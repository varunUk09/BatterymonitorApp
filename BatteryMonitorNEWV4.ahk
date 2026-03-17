#Requires AutoHotkey v2.0
#SingleInstance Force

; ╔══════════════════════════════════════════════════╗
; ║         BATTERY MONITOR by Varun                 ║
; ║  - Continuous beep until condition resolves      ║
; ║  - No annoying popups                            ║
; ║  - Custom sound support (.wav)                   ║
; ║  - Settings saved in Windows Registry            ║
; ╚══════════════════════════════════════════════════╝

REG_KEY := "HKCU\Software\BatteryMonitor"

ReadReg(name, default) {
    try return RegRead(REG_KEY, name)
    return default
}

global LOW_THRESH  := Integer(ReadReg("LowThresh",  30))
global HIGH_THRESH := Integer(ReadReg("HighThresh", 90))
global CHECK_SECS  := Integer(ReadReg("CheckSecs",  30))
global BEEP_SECS   := Integer(ReadReg("BeepSecs",   5))
global SOUND_LOW   := ReadReg("SoundLow",   "")
global SOUND_HIGH  := ReadReg("SoundHigh",  "")
global USE_CUSTOM  := Integer(ReadReg("UseCustom",  0))
global STARTUP     := Integer(ReadReg("Startup",    0))

global low_active      := false
global high_active     := false
global g_health_pct    := 0      ; cached battery health % , fetched once on startup

; ── Shared WMI connection (fast, reused) ─────────────
global g_wmi_svc := ""
InitWMI() {
    global g_wmi_svc
    try {
        wmi := ComObject("WbemScripting.SWbemLocator")
        g_wmi_svc := wmi.ConnectServer()
    }
}
InitWMI()

; ── Sleep/Wake handler ───────────────────────────────
OnMessage(0x218, WM_POWERBROADCAST)
WM_POWERBROADCAST(wParam, lParam, msg, hwnd) {
    global CHECK_SECS, BEEP_SECS
    if (wParam = 7) {
        Sleep(2000)
        InitWMI()
        CheckBattery()
        SetTimer(CheckBattery, CHECK_SECS * 1000)
        SetTimer(BeepLoop,     BEEP_SECS  * 1000)
        SetTimer(FastPlugCheck, 2000)
    }
    if (wParam = 4) {
        StopAlert()
    }
}

; ── Tray ─────────────────────────────────────────────
A_TrayMenu.Delete()
A_TrayMenu.Add("Battery Monitor", (*) => mainGui.Show())
A_TrayMenu.Add("Settings",        (*) => OpenSettings())
A_TrayMenu.Add("Exit",            (*) => ExitApp())
A_TrayMenu.Default := "Battery Monitor"
TraySetIcon("shell32.dll", 174)

; ════════════════════════════════════════════════════
; MAIN WINDOW
; ════════════════════════════════════════════════════
mainGui := Gui("+MinimizeBox -MaximizeBox", "Battery Monitor")
mainGui.BackColor := "0D0D1A"
mainGui.OnEvent("Close", (*) => mainGui.Hide())

mainGui.SetFont("s32 bold cFFD740", "Segoe UI")
global pctLbl := mainGui.Add("Text", "x24 y18 w220", "-- %")

mainGui.SetFont("s10 cAAAAAA", "Segoe UI")
global plugLbl := mainGui.Add("Text", "x248 y36 w80", "")

mainGui.SetFont("s8 cBBBBDD", "Segoe UI")
global cycleLbl := mainGui.Add("Text", "x170 y58 w158 Right", "Cycle: loading...")

mainGui.SetFont("s9 c888888", "Segoe UI")
global statusLbl := mainGui.Add("Text", "x24 y78 w290", "Starting...")

global batBar := mainGui.Add("Progress", "x24 y100 w290 h22 Background1A1A2E cFFD740 Range0-100", 0)

mainGui.SetFont("s8 c8888AA", "Segoe UI")
global timeLbl := mainGui.Add("Text", "x24 y128 w150", "")

; Actual % label (right side of timeLbl)
mainGui.SetFont("s8 cFFD740", "Segoe UI")
global actualPctLbl := mainGui.Add("Text", "x180 y128 w138 Right", "")

mainGui.Add("Text", "x24 y146 w290 h1 Background333355", "")

; Battery health info row
mainGui.SetFont("s8 c9999BB", "Segoe UI")
global designLbl := mainGui.Add("Text", "x24 y153 w290", "Health: loading...")

mainGui.Add("Text", "x24 y168 w290 h1 Background222233", "")

mainGui.SetFont("s8 c8888AA", "Segoe UI")
global threshLbl := mainGui.Add("Text", "x24 y175 w290", "")

mainGui.SetFont("s8 c333366", "Segoe UI")
global alertLbl := mainGui.Add("Text", "x24 y192 w290", "")

mainGui.SetFont("s9 bold cAAAAFF", "Segoe UI")
settBtn := mainGui.Add("Button", "x24 y212 w150 h30", "Settings")
settBtn.OnEvent("Click", (*) => OpenSettings())

mainGui.SetFont("s8 c7777AA", "Segoe UI")
mainGui.Add("Text", "x184 y220 w150", "Minimize = tray")

mainGui.Show("w338 h258")

; ── Start timers ─────────────────────────────────────
SetTimer(CheckBattery, CHECK_SECS * 1000)
SetTimer(BeepLoop,     BEEP_SECS  * 1000)
SetTimer(FastPlugCheck, 2000)   ; check plug status every 2 sec
; Fetch battery health once on startup (runs async so UI loads fast)
SetTimer(FetchBatteryHealth, -500)
CheckBattery()

; ════════════════════════════════════════════════════
; BATTERY CHECK
; ════════════════════════════════════════════════════
FetchBatteryHealth() {
    global g_health_pct
    health := GetBatteryHealth()
    if (health["design"] > 0) {
        designMah       := Round(health["design"] / 1000, 1)
        fullMah         := Round(health["full"]   / 1000, 1)
        g_health_pct    := health["health"]
        healthColor     := (g_health_pct < 50) ? "FF6666" : (g_health_pct < 70) ? "FFD740" : "00E676"
        designLbl.Text  := "Health: " g_health_pct "% — Design: " designMah "Wh  Max: " fullMah "Wh"
        designLbl.SetFont("c" healthColor)
    } else {
        g_health_pct   := 0
        designLbl.Text := "Health: N/A"
        designLbl.SetFont("c9999BB")
    }
    ; Also fetch cycle count once here
    cycles := GetCycleCount()
    if (cycles != "N/A" && cycles != "") {
        cycleNum   := Integer(cycles)
        cycleColor := (cycleNum > 500) ? "FF8888" : (cycleNum > 300) ? "FFD740" : "00E676"
        cycleLbl.Text := "Cycle: " cycles " cycles"
        cycleLbl.SetFont("c" cycleColor)
    } else {
        cycleLbl.Text := "Cycle: N/A"
        cycleLbl.SetFont("c9999BB")
    }
}

CheckBattery(*) {
    global low_active, high_active, LOW_THRESH, HIGH_THRESH, g_health_pct

    pct  := GetBatteryPct()
    plug := IsPluggedIn()

    ; Calculate displayPct FIRST before anything else
    if (g_health_pct > 0 && g_health_pct < 99)
        displayPct := Round(pct * g_health_pct / 100, 1)
    else
        displayPct := pct

    chkPct := displayPct

    if (chkPct < LOW_THRESH && !plug) {
        if (!low_active) {
            low_active := true
            alertLbl.Text := "!! LOW BATTERY — Beeping until plugged in..."
            alertLbl.SetFont("cFF4D4D")
            PlayAlert()
        }
    } else {
        if (low_active) {
            low_active := false
            alertLbl.Text := ""
        }
    }

    if (chkPct > HIGH_THRESH && plug) {
        if (!high_active) {
            high_active := true
            alertLbl.Text := "!! BATTERY FULL — Beeping until unplugged..."
            alertLbl.SetFont("c00E676")
            PlayAlert()
        }
    } else {
        if (high_active) {
            high_active := false
            alertLbl.Text := ""
        }
    }

    if (!low_active && !high_active)
        alertLbl.Text := ""

    ; Update Windows % label
    if (g_health_pct > 0 && g_health_pct < 99) {
        actualPctLbl.Text := "Windows: " pct "%"
        actualPctLbl.SetFont("c666688")
    } else {
        actualPctLbl.Text := ""
    }

    ; Color based on ACTUAL % and plug status
    if (plug)
        color := "00AAFF"
    else if (displayPct <= LOW_THRESH)
        color := "FF4D4D"
    else
        color := "FFD740"

    pctLbl.Text := displayPct " %"
    pctLbl.SetFont("c" color)
    plugLbl.Text := plug ? "[Charging]" : "[Battery]"
    plugLbl.SetFont("c" color)
    statusLbl.Text := plug ? "Charging" : "On Battery"
    statusLbl.SetFont("c" color)
    batBar.Value := displayPct
    batBar.Opt("c" color)
    timeLbl.Text  := GetTimeStr()
    threshLbl.Text := "Low alert < " LOW_THRESH "%     High alert > " HIGH_THRESH "%"

    ; Cycle count fetched once at startup via FetchBatteryHealth
}

; ════════════════════════════════════════════════════
; FAST PLUG CHECK — runs every 2s for instant response
; ════════════════════════════════════════════════════
FastPlugCheck(*) {
    global low_active, high_active, LOW_THRESH, HIGH_THRESH

    ; Only check plug status here (fast) — pct updated by CheckBattery
    plug := IsPluggedIn()
    pct  := GetBatteryPct()

    ; If low alert was active but now plugged in — stop immediately
    if (low_active && plug) {
        low_active := false
        StopAlert()             ; kill sound instantly
        alertLbl.Text := ""
        pctLbl.SetFont("c00AAFF")
        plugLbl.Text := "[Charging]"
        plugLbl.SetFont("c00AAFF")
        statusLbl.Text := "Charging"
        statusLbl.SetFont("c00AAFF")
        batBar.Opt("c00AAFF")
    }

    ; If high alert was active but now unplugged — stop immediately
    if (high_active && !plug) {
        high_active := false
        StopAlert()             ; kill sound instantly
        alertLbl.Text := ""
        pctLbl.SetFont("cFFD740")
        plugLbl.Text := "[Battery]"
        plugLbl.SetFont("cFFD740")
        statusLbl.Text := "On Battery"
        statusLbl.SetFont("cFFD740")
        batBar.Opt("cFFD740")
    }

    ; Also update plug label live every 2s
    if (!low_active && !high_active) {
        global g_health_pct
        fastDisplayPct := (g_health_pct > 0 && g_health_pct < 99) ? Round(pct * g_health_pct / 100, 1) : pct
        if (plug) {
            col := "00AAFF"
            plugLbl.Text := "[Charging]"
        } else {
            col := (fastDisplayPct <= LOW_THRESH) ? "FF4D4D" : "FFD740"
            plugLbl.Text := "[Battery]"
        }
        plugLbl.SetFont("c" col)
        statusLbl.Text := plug ? "Charging" : "On Battery"
        statusLbl.SetFont("c" col)
        batBar.Opt("c" col)
        pctLbl.SetFont("c" col)
    }
}

; ════════════════════════════════════════════════════
; BEEP LOOP — repeats every BEEP_SECS while alert active
; ════════════════════════════════════════════════════
BeepLoop(*) {
    global low_active, high_active
    if (low_active || high_active)
        PlayAlert()
}

; ════════════════════════════════════════════════════
; PLAY ALERT SOUND
; ════════════════════════════════════════════════════
PlayAlert() {
    global USE_CUSTOM, SOUND_LOW, SOUND_HIGH, low_active
    if (low_active) {
        if (USE_CUSTOM && SOUND_LOW != "") {
            if (FileExist(SOUND_LOW))
                SoundPlay(SOUND_LOW, 1)
            else {
                SetTimer(BeepLowAsync, -1)
                ShowToast("Low sound file missing! Using default beep.")
            }
        } else {
            SetTimer(BeepLowAsync, -1)
        }
    } else {
        if (USE_CUSTOM && SOUND_HIGH != "") {
            if (FileExist(SOUND_HIGH))
                SoundPlay(SOUND_HIGH, 1)
            else {
                SetTimer(BeepHighAsync, -1)
                ShowToast("High sound file missing! Using default beep.")
            }
        } else {
            SetTimer(BeepHighAsync, -1)
        }
    }
}

StopAlert() {
    ; Stop any playing sound immediately
    try SoundPlay("")          ; empty string stops current wav
    try DllCall("PlaySound", "Ptr", 0, "Ptr", 0, "UInt", 0)  ; stops all
}

BeepLowAsync() {
    global low_active
    if (low_active) {
        SoundBeep(440, 250)
        Sleep(80)
        SoundBeep(440, 250)
        Sleep(80)
        SoundBeep(440, 250)
    }
}

BeepHighAsync() {
    global high_active
    if (high_active) {
        SoundBeep(880, 250)
        Sleep(80)
        SoundBeep(880, 250)
    }
}

; ════════════════════════════════════════════════════
; SETTINGS WINDOW
; ════════════════════════════════════════════════════
OpenSettings(*) {
    global LOW_THRESH, HIGH_THRESH, CHECK_SECS, BEEP_SECS, USE_CUSTOM, SOUND_LOW, SOUND_HIGH, STARTUP, REG_KEY

    sGui := Gui("+Owner +AlwaysOnTop -MaximizeBox -MinimizeBox", "Settings")
    sGui.BackColor := "0D0D1A"

    sGui.SetFont("s11 bold cFFFFFF", "Segoe UI")
    sGui.Add("Text", "x20 y18 w300", "Battery Monitor - Settings")
    sGui.Add("Text", "x20 y40 w300 h1 Background333355", "")

    ; LOW
    sGui.SetFont("s9 cFF8888", "Segoe UI")
    sGui.Add("Text", "x20 y55 w210", "Low Battery Alert (%):")
    sGui.SetFont("s10 cFFFFFF", "Segoe UI")
    lowEdit := sGui.Add("Edit", "x230 y52 w60 h26 Center Background1A1A2E cFFFFFF", LOW_THRESH)
    sGui.Add("UpDown", "Range1-99", LOW_THRESH)
    sGui.SetFont("s8 c9999BB", "Segoe UI")
    sGui.Add("Text", "x20 y82 w280", "Beep starts when battery BELOW this % and unplugged")

    ; HIGH
    sGui.SetFont("s9 c88FF99", "Segoe UI")
    sGui.Add("Text", "x20 y106 w210", "High Battery Alert (%):")
    sGui.SetFont("s10 cFFFFFF", "Segoe UI")
    highEdit := sGui.Add("Edit", "x230 y103 w60 h26 Center Background1A1A2E cFFFFFF", HIGH_THRESH)
    sGui.Add("UpDown", "Range2-100", HIGH_THRESH)
    sGui.SetFont("s8 c9999BB", "Segoe UI")
    sGui.Add("Text", "x20 y133 w280", "Beep starts when battery ABOVE this % and charging")

    ; Check interval
    sGui.SetFont("s9 cAAAAFF", "Segoe UI")
    sGui.Add("Text", "x20 y157 w210", "Check Interval (seconds):")
    sGui.SetFont("s10 cFFFFFF", "Segoe UI")
    intEdit := sGui.Add("Edit", "x230 y154 w60 h26 Center Background1A1A2E cFFFFFF", CHECK_SECS)
    sGui.Add("UpDown", "Range5-300", CHECK_SECS)

    ; Beep repeat
    sGui.SetFont("s9 cFFD740", "Segoe UI")
    sGui.Add("Text", "x20 y184 w210", "Beep Repeat Every (seconds):")
    sGui.SetFont("s10 cFFFFFF", "Segoe UI")
    beepIntEdit := sGui.Add("Edit", "x230 y181 w60 h26 Center Background1A1A2E cFFFFFF", BEEP_SECS)
    sGui.Add("UpDown", "Range1-120", BEEP_SECS)
    sGui.SetFont("s8 c9999BB", "Segoe UI")
    sGui.Add("Text", "x20 y211 w280", "How often beep repeats while alert is active")

    sGui.Add("Text", "x20 y230 w300 h1 Background333355", "")

    ; Custom sounds
    sGui.Add("Text", "x20 y230 w300 h1 Background333355", "")

    sGui.SetFont("s9 cCCCCCC", "Segoe UI")
    customChk := sGui.Add("CheckBox", "x20 y240 w280 Background0D0D1A cCCCCCC", "Use custom sounds (else use default beep)")
    customChk.Value := USE_CUSTOM

    ; Low battery sound
    sGui.SetFont("s9 cFF8888", "Segoe UI")
    sGui.Add("Text", "x20 y268 w280", "Low Battery Sound (.wav):")
    lowSoundEdit := sGui.Add("Edit", "x20 y286 w218 h24 Background1A1A2E cFFFFFF ReadOnly", SOUND_LOW)
    sGui.SetFont("s8 bold cAAAAFF", "Segoe UI")
    browseLowBtn := sGui.Add("Button", "x246 y284 w54 h26", "Browse")
    browseLowBtn.OnEvent("Click", (*) => BrowseWav(lowSoundEdit))

    ; High battery sound
    sGui.SetFont("s9 c88FF99", "Segoe UI")
    sGui.Add("Text", "x20 y318 w280", "High Battery Sound (.wav):")
    highSoundEdit := sGui.Add("Edit", "x20 y336 w218 h24 Background1A1A2E cFFFFFF ReadOnly", SOUND_HIGH)
    sGui.SetFont("s8 bold cAAAAFF", "Segoe UI")
    browseHighBtn := sGui.Add("Button", "x246 y334 w54 h26", "Browse")
    browseHighBtn.OnEvent("Click", (*) => BrowseWav(highSoundEdit))

    sGui.SetFont("s8 c9999BB", "Segoe UI")
    sGui.Add("Text", "x20 y366 w280", "Leave blank = built-in beep tones")

    sGui.Add("Text", "x20 y382 w300 h1 Background333355", "")

    sGui.SetFont("s9 cCCCCCC", "Segoe UI")
    startChk := sGui.Add("CheckBox", "x20 y392 w280 Background0D0D1A cCCCCCC", "Run automatically at Windows Startup")
    startChk.Value := STARTUP

    sGui.Add("Text", "x20 y418 w300 h1 Background333355", "")

    sGui.SetFont("s9 bold cFFFFFF", "Segoe UI")
    saveBtn := sGui.Add("Button", "x20 y428 w130 h32", "Save Settings")
    sGui.SetFont("s9 bold cFF6666", "Segoe UI")
    cancelBtn := sGui.Add("Button", "x165 y428 w80 h32", "Cancel")
    cancelBtn.OnEvent("Click", (*) => sGui.Destroy())
    saveBtn.OnEvent("Click", SaveIt)

    sGui.Show("w320 h476")
    CenterWindow(sGui)

    SaveIt(*) {
        global LOW_THRESH, HIGH_THRESH, CHECK_SECS, BEEP_SECS, USE_CUSTOM, SOUND_LOW, SOUND_HIGH, STARTUP, REG_KEY

        newLow     := Integer(lowEdit.Value)
        newHigh    := Integer(highEdit.Value)
        newInt     := Integer(intEdit.Value)
        newBeepInt := Integer(beepIntEdit.Value)
        newCustom    := customChk.Value
        newSoundLow  := lowSoundEdit.Value
        newSoundHigh := highSoundEdit.Value
        newStartup   := startChk.Value

        if (newLow >= newHigh) {
            MsgBox("Low threshold must be less than High threshold!", "Invalid", "Icon! T3")
            return
        }
        if (newInt < 5) {
            MsgBox("Check interval must be at least 5 seconds!", "Invalid", "Icon! T3")
            return
        }
        if (newCustom && newSoundLow = "" && newSoundHigh = "") {
            MsgBox("Please select at least one .wav file or uncheck custom sound!", "Invalid", "Icon! T3")
            return
        }

        LOW_THRESH  := newLow
        HIGH_THRESH := newHigh
        CHECK_SECS  := newInt
        BEEP_SECS   := newBeepInt
        USE_CUSTOM  := newCustom
        SOUND_LOW   := newSoundLow
        SOUND_HIGH  := newSoundHigh
        STARTUP     := newStartup

        RegWrite(LOW_THRESH,  "REG_SZ", REG_KEY, "LowThresh")
        RegWrite(HIGH_THRESH, "REG_SZ", REG_KEY, "HighThresh")
        RegWrite(CHECK_SECS,  "REG_SZ", REG_KEY, "CheckSecs")
        RegWrite(BEEP_SECS,   "REG_SZ", REG_KEY, "BeepSecs")
        RegWrite(USE_CUSTOM,  "REG_SZ", REG_KEY, "UseCustom")
        RegWrite(SOUND_LOW,   "REG_SZ", REG_KEY, "SoundLow")
        RegWrite(SOUND_HIGH,  "REG_SZ", REG_KEY, "SoundHigh")
        RegWrite(STARTUP,     "REG_SZ", REG_KEY, "Startup")

        ; Startup registry
        startupKey := "HKCU\Software\Microsoft\Windows\CurrentVersion\Run"
        if (STARTUP)
            RegWrite('"' A_ScriptFullPath '"', "REG_SZ", startupKey, "BatteryMonitor")
        else
            try RegDelete(startupKey, "BatteryMonitor")

        SetTimer(CheckBattery, CHECK_SECS * 1000)
        SetTimer(BeepLoop,     BEEP_SECS  * 1000)
        CheckBattery()
        sGui.Destroy()
        ShowToast("Settings saved!")
    }
}

BrowseWav(editCtrl) {
    file := FileSelect(1, , "Select Alert Sound", "WAV Audio (*.wav)")
    if (file != "")
        editCtrl.Value := file
}

; ════════════════════════════════════════════════════
ShowToast(msg) {
    t := Gui("-Caption +AlwaysOnTop +ToolWindow")
    t.BackColor := "0A1A0A"
    t.SetFont("s9 bold c00E676", "Segoe UI")
    t.Add("Text", "x14 y10 w220", msg)
    t.Show("x" (A_ScreenWidth - 260) " y" (A_ScreenHeight - 80) " w245 h40 NoActivate")
    SetTimer(() => SafeClose(t), -2500)
}

GetBatteryPct() {
    ; Use Windows native API — same source as taskbar battery %
    try {
        static SYSTEM_POWER_STATUS := Buffer(12, 0)
        DllCall("GetSystemPowerStatus", "Ptr", SYSTEM_POWER_STATUS)
        pct := NumGet(SYSTEM_POWER_STATUS, 2, "UChar")
        if (pct <= 100)
            return pct
    }
    ; Fallback to WMI if API fails
    global g_wmi_svc
    try {
        for bat in g_wmi_svc.ExecQuery("SELECT * FROM Win32_Battery")
            return Integer(bat.EstimatedChargeRemaining)
    }
    return 50
}

IsPluggedIn() {
    try {
        static SYSTEM_POWER_STATUS := Buffer(12, 0)
        DllCall("GetSystemPowerStatus", "Ptr", SYSTEM_POWER_STATUS)
        acStatus := NumGet(SYSTEM_POWER_STATUS, 0, "UChar")
        return (acStatus = 1)   ; 1 = AC power (plugged in)
    }
    return false
}

GetBatteryHealth() {
    ; Returns Map with designMWh, fullMWh, healthPct
    result := Map()
    result["design"] := 0
    result["full"]   := 0
    result["health"] := 0
    try {
        psFile    := A_Temp "\bathealthy.ps1"
        resFile   := A_Temp "\bathealthy.txt"
        psScript  := '$b = Get-WmiObject -Class BatteryStaticData -Namespace "root\wmi" | Select-Object -First 1; '
                   . '$f = Get-WmiObject -Class BatteryFullChargedCapacity -Namespace "root\wmi" | Select-Object -First 1; '
                   . '"DESIGN=" + $b.DesignedCapacity + "|FULL=" + $f.FullChargedCapacity | Out-File "' resFile '" -Encoding ascii'
        fh := FileOpen(psFile, "w")
        fh.Write(psScript)
        fh.Close()
        RunWait('powershell.exe -WindowStyle Hidden -ExecutionPolicy Bypass -File "' psFile '"',, "Hide")
        FileDelete(psFile)
        if (FileExist(resFile)) {
            raw := Trim(FileRead(resFile))
            raw := StrReplace(raw, "`r", "")
            raw := StrReplace(raw, "`n", "")
            FileDelete(resFile)
            if RegExMatch(raw, "DESIGN=(\d+)", &m1)
                result["design"] := Integer(m1[1])
            if RegExMatch(raw, "FULL=(\d+)", &m2)
                result["full"] := Integer(m2[1])
            if (result["design"] > 0 && result["full"] > 0)
                result["health"] := Round(result["full"] / result["design"] * 100, 1)
        }
    } catch {
    }
    return result
}

GetCycleCount() {
    ; Method 1: WMI Win32_Battery CycleCount
    try {
        wmi := ComObject("WbemScripting.SWbemLocator")
        svc := wmi.ConnectServer()
        for bat in svc.ExecQuery("SELECT * FROM Win32_Battery") {
            if (bat.CycleCount > 0)
                return bat.CycleCount
        }
    } catch {
    }

    ; Method 2: PowerShell BatteryStatus WMI namespace
    try {
        resultFile := A_Temp "\batcyc.txt"
        psFile     := A_Temp "\batcyc.ps1"
        psScript   := 'try { $b = Get-WmiObject -Class BatteryStatus -Namespace "root\wmi"; $b.CycleCount | Out-File "' resultFile '" -Encoding ascii } catch { "N/A" | Out-File "' resultFile '" -Encoding ascii }'
        fh := FileOpen(psFile, "w")
        fh.Write(psScript)
        fh.Close()
        RunWait('powershell.exe -WindowStyle Hidden -ExecutionPolicy Bypass -File "' psFile '"',, "Hide")
        FileDelete(psFile)
        if (FileExist(resultFile)) {
            val := Trim(FileRead(resultFile))
            val := StrReplace(val, "`r", "")
            val := StrReplace(val, "`n", "")
            FileDelete(resultFile)
            if (val != "" && val != "N/A" && IsInteger(val) && Integer(val) > 0)
                return val
        }
    } catch {
    }

    ; Method 3: powercfg HTML report — parse cycle count
    try {
        repPath := A_Temp "\batreport.html"
        RunWait('powercfg.exe /batteryreport /output "' repPath '"',, "Hide")
        if (FileExist(repPath)) {
            html := FileRead(repPath)
            FileDelete(repPath)
            if RegExMatch(html, "(?i)CYCLE COUNT[^0-9]*(\d+)", &m)
                return m[1]
            if RegExMatch(html, ">(\d+)<[^>]*>.*(?i)cycle", &m)
                return m[1]
        }
    } catch {
    }

    return "N/A"
}
GetTimeStr() {
    global g_wmi_svc
    try {
        for bat in g_wmi_svc.ExecQuery("SELECT * FROM Win32_Battery") {
            mins := bat.EstimatedRunTime
            if (mins > 9990)
                return "Fully Charged / Calculating..."
            h := mins // 60
            m := Mod(mins, 60)
            return "~" h "h " m "m remaining"
        }
    }
    return ""
}

SafeClose(g) {
    try g.Destroy()
}

CenterWindow(g) {
    g.GetPos(, , &gw, &gh)
    g.Move((A_ScreenWidth - gw) // 2, (A_ScreenHeight - gh) // 2)
}
