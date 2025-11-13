/*
===============================================================================
AutoHelper - Vermintide 2 Automation Script
===============================================================================
Author: ChadMasodin
Version: 2.0
Type game UI: New
Language: English
Date: 10/11/2025
Description: Automated script for routine processes in Vermintide 2
===============================================================================
*/

; ============================================================================
; INITIALIZATION AND SETTINGS
; ============================================================================

#NoEnv
#SingleInstance Force
SetBatchLines -1
ListLines Off
SetMouseDelay, -1
;CoordMode, Mouse, Client
SetTitleMatchMode, 2
SetControlDelay, -1
DetectHiddenWindows, On
SetStoreCapsLockMode, Off

; ============================================================================
; CONSTANTS
; ============================================================================

global VERSION := "2.0"
global DELAY_SHORT := 100
global DELAY_MEDIUM := 300
global DELAY_LONG := 500
global DELAY_ANIMATION := 1500
global DELAY_OCR_COOLDOWN := 300
global COLOR_GREEN_DUST := 0x58A113
global COLOR_TOLERANCE := 15

; ============================================================================
; GLOBAL VARIABLES - COORDINATES
; ============================================================================

global X1 := "", Y1 := ""  ; White button
global X2 := "", Y2 := ""  ; Green button
global X3 := "", Y3 := ""  ; Blue button
global X4 := "", Y4 := ""  ; Yellow button
global X5 := "", Y5 := ""  ; Atanor button

global OCR_X := "", OCR_Y := "", OCR_W := "", OCR_H := ""
global GreenDust_X1 := "", GreenDust_Y1 := ""
global GreenDust_X2 := "", GreenDust_Y2 := ""

; ============================================================================
; GLOBAL VARIABLES - SCRIPT STATE
; ============================================================================

global Stop := 0
global Target_Window := ""
global LoopCount := ""
global SelectedProperties := ""
global savedPreset := ""
global lastOCRText := ""
global lastOCRTime := 0

; ============================================================================
; GLOBAL VARIABLES - HOTKEYS
; ============================================================================

global Stop_key := "F1"
global Pause_key := "F2"
global Reset_key := "F3"
global OpenChests_key := "F5"
global Salvage_items_key := "F7"
global Salvage_red_items_key := "F9"
global Reroll_Properties_key := "F11"
global Atanor_key := "F12"
global Close_script_key := "^ESC"

; ============================================================================
; APPLICATION PATHS
; ============================================================================

RegExMatch(A_ScriptName, "^(.*?)\.", match)
global basename1 := match
global WINTITLE := basename1 . " v" . VERSION
global presetsDir := A_AppData "\" basename1

if !FileExist(presetsDir)
    FileCreateDir, %presetsDir%

global logFile := presetsDir "\debug.log"
global presetFile := presetsDir "\presets.ini"

; ============================================================================
; LIBRARY INCLUDES
; ============================================================================

#include <Vis2>

; ============================================================================
; INITIALIZATION
; ============================================================================

ManageLogFile() {
    global logFile

    if (!FileExist(logFile))
        return

    FileGetSize, fileSize, %logFile%

    if (fileSize > 512000) {
        FileDelete, %logFile%
        FileAppend, [Log file was cleared due to size limit]`n, %logFile%
        LogAction("Log file cleared (exceeded 500 KB size)")
    }
}

ManageLogFile()

GoSub, CreateGUI
return

; ============================================================================
; HOTKEY MANAGEMENT
; ============================================================================


UpdateAllKeys()
{
    global
    Gui, Submit, NoHide
    Hotkey, %Stop_key%, Stope, On
    Hotkey, %Pause_key%, Pauza, On
    Hotkey, %Reset_key%, Reset, On
    Hotkey, %OpenChests_key%, OpenChests, On
    Hotkey, %Salvage_items_key%, SalvageItems, On
    Hotkey, %Salvage_red_items_key%, SalvageRedItems, On
    Hotkey, %Reroll_Properties_key%, RerollProperties, On
    Hotkey, %Atanor_key%, UpgradeAtanor, On
    Hotkey, %Close_script_key%, CloseScript, On
}

UpdateKey(KeyVariable, Label) {
    Global
    GuiControlGet, New_Key, , %KeyVariable%
    Current_Key := %KeyVariable%

    if (New_Key != Current_Key) {
        try Hotkey, %Current_Key%, %Label%, Off
        %KeyVariable% := New_Key
        Hotkey, %New_Key%, %Label%, On

        if (savedPreset != "")
            IniWrite, %New_Key%, %presetFile%, %savedPreset%, %KeyVariable%
    }
}

; ============================================================================
; GUI CREATION
; ============================================================================

CreateGUI:
    Gui, +AlwaysOnTop -DPIScale +OwnDialogs

    ; Preset Panel
    Gui, Font, s10
    Gui, Add, Text, yp+11 xm section, Preset list:
    Gui, Add, ComboBox, x+8 vsavedPreset gPresetChange
    Gui, Add, Button, x+10 h21 w70 gSavePreset, Save
    Gui, Add, Button, x+10 h21 w70 gDeletePreset vDELETEBUTTON, Delete
    Gui, Add, StatusBar

    ; Configuration Section
    Gui, Font, c535353 s13 Bold Q3, Segoe UI Black
    Gui, Add, GroupBox, x10 y40 w380 h300 Center Section, Script Settings
    Gui, Font
    Gui, Font, s11, Segoe UI

    ; Atanor Position
    Gui, Add, Text, xs+15 yp+30, #1 Athanor upgrade button position:
    Gui, Font, s9
    Gui, Add, Button, xs+15 yp+30 w80 h25 gSetPosAtanor, Set
    Gui, Add, Edit, x+15 yp+3 w80 h20 vPos_5 ReadOnly,
    Gui, Font, s11, Segoe UI

    ; Salvage Buttons
    Gui, Add, Text, xs+15 yp+30, #2 Item salvage button positions:
    Gui, Font, s9
    Gui, Add, Button, xs+15 yp+30 w70 h20 gSetPosWhite, White
    Gui, Add, Edit, x+15 w80 h20 vPos_1 ReadOnly,
    Gui, Add, Button, x+15 w70 h20 gSetPosBlue, Blue
    Gui, Add, Edit, x+15 w80 h20 vPos_3 ReadOnly ,
    Gui, Add, Button, xs+15 yp+25 w70 h20 gSetPosGreen, Green
    Gui, Add, Edit, x+15 w80 h20 vPos_2 ReadOnly,
    Gui, Add, Button, x+15 w70 h20 gSetPosYellow, Yellow
    Gui, Add, Edit, x+15 w80 h20 vPos_4 ReadOnly,

    ; OCR Area
    Gui, Font, s11, Segoe UI
    Gui, Add, Text, xs+15 yp+30, #3 Item properties area:
    Gui, Font, s9
    Gui, Add, Button, xs+15 yp+25 w140 h25 gSetOCRArea, Select area
    Gui, Add, Edit, x+15 yp+3 w140 h20 vOCR_Pos ReadOnly,

    ; Green Dust Area
    Gui, Font, s11, Segoe UI
    Gui, Add, Text, xs+15 yp+30, #4 Green dust icon area:
    Gui, Font, s9
    Gui, Add, Button, xs+15 yp+25 w140 h25 gSetGreenDustArea, Select area
    Gui, Add, Edit, x+15 yp+3 w140 h20 vGreenDust_Pos ReadOnly,

    ; Additional Settings
    Gui, Font, s11, Segoe UI
    Gui, Add, Text, x65 y580, General Loop Limiter (optional):
    Gui, Add, Edit, x+10 w75 h20 vLoopCount Number

    Gui, Font, s9 Bold, Segoe UI
    Gui, Add, Button, x25 y610 w130 h30 gSetLocation, Set Target Window:
    Gui, Add, Edit, x+15 yp+3 w210 h25 vTarget_Window ReadOnly, %Target_Window%
    Gui, Font

    ; Action Buttons Section
    Gui, Font, c535353 s13 Bold Q3, Segoe UI Black
    Gui, Add, GroupBox, x10 y350 w380 h210 Center Section, Buttons
    Gui, Font, s9 Bold, Segoe UI

    Gui, Add, Button, xs+15 yp+30 w110 h35 gOpenChests, OPEN CHESTS
    Gui, Add, Button, x+10 w110 h35 gSalvageItems, SALVAGE ITEMS
    Gui, Add, Button, x+10 w110 h35 gUpgradeAtanor, UPGRADE ATHANOR
    Gui, Add, Button, xs+15 yp+45 w170 h35 gSalvageRedItems, SALVAGE RED ITEMS
    Gui, Add, Button, x+10 w170 h35 gRerollProperties, REROLL PROPERTIES
    Gui, Add, Button, xs+40 yp+45 w300 h25 gShowInfo, <><><> HOTKEYS AND INFO  <><><>
    Gui, Add, Button, xs+15 yp+35 w110 h35 gStope, STOP
    Gui, Add, Button, x+10 w110 h35 gPauza, PAUSE
    Gui, Add, Button, x+10 w110 h35 gReset, RESTART
    Gui, Font

    Gui, Show, w400 h670, %WINTITLE%
    GoSub, UpdatePresetList
    #include <menu>
return

; ============================================================================
; INFO WINDOW
; ============================================================================

ShowInfo:
    Gui, InfoWindow:New, , Help
    Gui, +AlwaysOnTop -DPIScale +Owner1
    Gui, Add, Picture, x10 y33, icon.png

    Gui, Font, c666666 s11 Bold Q3, Segoe UI Black
    Gui, Add, GroupBox, x10 y10 w280 h150 Center Section, INFO
    Gui, Font,

    Gui, Add, Link, xs+145 yp+25, Author: <a href="https://steamcommunity.com/id/ChadMasodin">ChadMasodin</a>
    Gui, Add, Text, xs+145 yp+20, Release date:  27/12/24
    Gui, Add, Text, xs+145 yp+15, Update date:  10/11/25
    Gui, Add, Link, xs+10 yp+25, AutoHelper - is a script that automates a number`nof routine processes in Vermintide 2.`n`nMore information about it can be found in this <a href="https://steamcommunity.com/sharedfiles/filedetails/?id=3385048068\">guide.</a>

    ; Hotkeys Section
    Gui, Font, c666666 s11 Bold Q3, Segoe UI Black
    Gui, Add, GroupBox, x10 y170 w280 h300 Center Section, HOTKEYS
    Gui, Font, s11

; Define common coordinates for alignment
labelX := "xs+10"     ; X coordinate for all text labels
hotkeyX := "x+1"      ; X coordinate for all Hotkey fields (after label)
hotkeyWidth := "w90"  ; Width of all Hotkey fields
labelWidth := "w155"  ; Text label width for alignment

Gui, Add, Text, cBlue %labelX% yp+30 %labelWidth% Left, Stop:
Gui, Add, Hotkey, %hotkeyX% yp %hotkeyWidth% h20 vStop_key, %Stop_key%

Gui, Add, Text, cBlue %labelX% yp+30 %labelWidth% Left, Pause:
Gui, Add, Hotkey, %hotkeyX% yp %hotkeyWidth% h22 vPause_key, %Pause_key%

Gui, Add, Text, cBlue %labelX% yp+30 %labelWidth% Left, Restart:
Gui, Add, Hotkey, %hotkeyX% yp %hotkeyWidth% h22 vReset_key, %Reset_key%

Gui, Add, Text, cRed %labelX% yp+30 %labelWidth% Left, Open Chests:
Gui, Add, Hotkey, %hotkeyX% yp %hotkeyWidth% h20 vOpenChests_key, %OpenChests_key%

Gui, Add, Text, cRed %labelX% yp+30 %labelWidth% Left, Salvage Items:
Gui, Add, Hotkey, %hotkeyX% yp %hotkeyWidth% h20 vSalvage_items_key, %Salvage_items_key%

Gui, Add, Text, cRed %labelX% yp+30 %labelWidth% Left, Salvage Red Items:
Gui, Add, Hotkey, %hotkeyX% yp %hotkeyWidth% h20 vSalvage_red_items_key, %Salvage_red_items_key%

Gui, Add, Text, cRed %labelX% yp+30 %labelWidth% Left, Reroll Properties:
Gui, Add, Hotkey, %hotkeyX% yp %hotkeyWidth% h20 vReroll_Properties_key, %Reroll_Properties_key%

Gui, Add, Text, cRed %labelX% yp+30 %labelWidth% Left, Upgrade Athanor:
Gui, Add, Hotkey, %hotkeyX% yp %hotkeyWidth% h20 vAtanor_key, %Atanor_key%

Gui, Add, Text, cMaroon %labelX% yp+30 %labelWidth% Left, Close Script:
Gui, Add, Hotkey, %hotkeyX% yp %hotkeyWidth% h20 vClose_script_key, %Close_script_key%

    Gui, Font, s11 Bold Q3, Segoe UI
    Gui, Add, Button, x10 y480 w280 h45 gSaveHotkeySettings, OK
    Gui, Font

    Gui, Show, w300 h540
return

SaveHotkeySettings:
    Gui, Submit
    UpdateAllKeys()

    if (savedPreset != "") {
        try {
            Loop, Parse, % "Stop_key|Pause_key|Reset_key|OpenChests_key|Salvage_items_key|Salvage_red_items_key|Reroll_Properties_key|Atanor_key|Close_script_key", |
                IniWrite, % %A_LoopField%, %presetFile%, %savedPreset%, %A_LoopField%

            SB_SetText("Hotkeys saved to preset '" . savedPreset . "'")
        } catch e {
            LogAction("Error saving hotkeys: " . e.Message)
        }
    }

    Gui, Destroy
return

; ============================================================================
; VALIDATION FUNCTIONS
; ============================================================================

ValidateScript(requireCoords := false, coordsList := "") {
    global LoopCount, Target_Window

    if (Target_Window == "") {
        MsgBox, 48, Error, Target window not set. Set the window before running the script.
        return false
    }

    if (requireCoords) {
        Loop, Parse, coordsList, `,
        {
            coord := Trim(A_LoopField)
            if (%coord% == "") {
                MsgBox, 16, Error, Coordinate %coord% not set. Set required coordinates before running.
                return false
            }
        }
    }

    return true
}

CheckActiveWindow() {
    global Target_Window
    WinGetActiveTitle, Current_Window

    if (Current_Window != Target_Window) {
        ToolTip, Script paused. Activate window: %Target_Window%
        Sleep, DELAY_LONG
        return false
    }
    return true
}

; ============================================================================
; UTILITY FUNCTIONS
; ============================================================================

LogAction(message) {
    global logFile
    timestamp := A_Now
    FormatTime, formattedTime, %timestamp%, yyyy-MM-dd HH:mm:ss

    try {
        FileAppend, [%formattedTime%] %message%`n, %logFile%
    }
}

UpdateTooltip(action, count) {
    global LoopCount
    if (LoopCount != "" && LoopCount > 0)
        ToolTip, %action%... Cycle: %count%/%LoopCount%
    else
        ToolTip, %action%... Cycle: %count% (infinite)
}

SendKey(key, delay := 300) {
    Send, {%key% down}
    Send, {%key% up}
    Sleep, delay
}

Min(a, b) {
    return a < b ? a : b
}

; ============================================================================
; COORDINATE SETTING FUNCTIONS
; ============================================================================

SetPosWhite:
    Stop := 1
    GetClickPos(X1, Y1)
    GuiControl,, Pos_1, %X1% %Y1%
return

SetPosGreen:
    Stop := 1
    GetClickPos(X2, Y2)
    GuiControl,, Pos_2, %X2% %Y2%
return

SetPosBlue:
    Stop := 1
    GetClickPos(X3, Y3)
    GuiControl,, Pos_3, %X3% %Y3%
return

SetPosYellow:
    Stop := 1
    GetClickPos(X4, Y4)
    GuiControl,, Pos_4, %X4% %Y4%
return

SetPosAtanor:
    Stop := 1
    GetClickPos(X5, Y5)
    GuiControl,, Pos_5, %X5% %Y5%
return

SetOCRArea:
    Stop := 1
    GetOCRArea(OCR_X, OCR_Y, OCR_W, OCR_H)
    GuiControl,, OCR_Pos, %OCR_X% %OCR_Y% %OCR_W% %OCR_H%
return

SetGreenDustArea:
    Stop := 1
    GetGreenDustArea(GreenDust_X1, GreenDust_Y1, GreenDust_X2, GreenDust_Y2)
    GuiControl,, GreenDust_Pos, %GreenDust_X1% %GreenDust_Y1% %GreenDust_X2% %GreenDust_Y2%
return

SetLocation:
    Target_Window := SetWindow(Target_Window)
    GuiControl,, Target_Window, %Target_Window%
return

GetClickPos(ByRef X, ByRef Y) {
    isPressed := 0
    X := "", Y := ""

    Loop {
        if (GetKeyState("Esc", "P")) {
            ToolTip
            X := "", Y := ""
            return
        }

        Left_Mouse := GetKeyState("LButton")
        MouseGetPos, currentX, currentY
        ToolTip, Click LMB to select`nPress ESC to cancel`n`nCurrent coordinates:`nX=%currentX% Y=%currentY%

        if (Left_Mouse == False && isPressed == 0) {
            isPressed := 1
        } else if (Left_Mouse == True && isPressed == 1) {
            X := currentX
            Y := currentY
            ToolTip
            break
        }
    }
}

SetWindow(Target_Window) {
    isPressed := 0
    i := 0
    Target_Window := ""

    Loop {
        if (GetKeyState("Esc", "P")) {
            ToolTip
            return ""
        }

        Left_Mouse := GetKeyState("LButton")
        WinGetTitle, Temp_Window, A
        ToolTip, Double-click LMB on window`nPress ESC to cancel`n`nCurrent Window: %Temp_Window%

        if (Left_Mouse == False && isPressed == 0) {
            isPressed := 1
        } else if (Left_Mouse == True && isPressed == 1) {
            i++, isPressed := 0
            if (i >= 2) {
                Target_Window := Temp_Window
                ToolTip
                break
            }
        }
    }

    return Target_Window
}

GetOCRArea(ByRef X, ByRef Y, ByRef W, ByRef H) {

    ; First loop - wait for LMB press
    Loop {
        ; Check ESC press
        if (GetKeyState("Esc", "P")) {
            ToolTip
            X := Y := W := H := ""
            return
        }

        ; Check LMB press to exit first loop
        if (GetKeyState("LButton", "P")) {
            break
        }
        MouseGetPos, currentX, currentY
        ToolTip, Hold LMB and select item properties area`nPress ESC to cancel`n`nCurrent coordinates:`nX1=%currentX% Y1=%currentY%

        Sleep, 10
    }

    ; Get start coordinates after LMB press
    MouseGetPos, startX, startY

    ; Second loop - track area selection
    Loop {
        ; Check ESC press
        if (GetKeyState("Esc", "P")) {
            ToolTip
            X := Y := W := H := ""
            return
        }

        ; Check if LMB is still held
        if (!GetKeyState("LButton", "P")) {
            break
        }

        ; Get current coordinates
        MouseGetPos, currentX, currentY

        ; Calculate area dimensions
        width := Abs(currentX - startX)
        height := Abs(currentY - startY)
        minX := Min(startX, currentX)
        minY := Min(startY, currentY)

        ; Show selection info with moving ToolTip
        ToolTip, Selection: %minX% %minY% [%width% x %height%]`n`nRelease LMB to finish`nPress ESC to cancel

        Sleep, 10
    }

    ; Get end coordinates
    MouseGetPos, endX, endY

    ; Calculate final area
    X := Min(startX, endX)
    Y := Min(startY, endY)
    W := Abs(startX - endX)
    H := Abs(startY - endY)

    ToolTip, Item properties area set!`nX: %X% Y: %Y%`nW: %W% H: %H%
    Sleep, DELAY_ANIMATION
    ToolTip
}


GetGreenDustArea(ByRef X1, ByRef Y1, ByRef X2, ByRef Y2) {
    ; First loop - wait for LMB press
    Loop {
        ; Check ESC press
        if (GetKeyState("Esc", "P")) {
            ToolTip
            X1 := Y1 := X2 := Y2 := ""
            return
        }

        ; Check LMB press to exit first loop
        if (GetKeyState("LButton", "P")) {
            break
        }
                ; Get current coordinates
        MouseGetPos, currentX, currentY
        ToolTip, Hold LMB and select green dust area`nPress ESC to cancel`n`nCurrent coordinates:`nX1=%currentX% Y1=%currentY%

        Sleep, 10
    }

    ; Get start coordinates after LMB press
    MouseGetPos, startX, startY

    ; Second loop - track area selection
    Loop {
        ; Check ESC press
        if (GetKeyState("Esc", "P")) {
            ToolTip
            X1 := Y1 := X2 := Y2 := ""
            return
        }

        ; Check if LMB is still held
        if (!GetKeyState("LButton", "P")) {
            break
        }

        ; Get current coordinates
        MouseGetPos, currentX, currentY

        ; Calculate current rectangle coordinates
        currentX1 := Min(startX, currentX)
        currentY1 := Min(startY, currentY)
        currentX2 := Max(startX, currentX)
        currentY2 := Max(startY, currentY)

        ; Show selection info with moving ToolTip
        ToolTip, Selection: %currentX1% %currentY1% - %currentX2% %currentY2%`n`nRelease LMB to finish`nPress ESC to cancel

        Sleep, 10
    }

    ; Get end coordinates
    MouseGetPos, endX, endY

    ; Calculate final area as X1,Y1,X2,Y2
    X1 := Min(startX, endX)
    Y1 := Min(startY, endY)
    X2 := Max(startX, endX)
    Y2 := Max(startY, endY)

    ToolTip, Green dust area set!`nX1: %X1% Y1: %Y1%`nX2: %X2% Y2: %Y2%
    Sleep, DELAY_ANIMATION
    ToolTip
}
; ============================================================================
; CONTROL FUNCTIONS
; ============================================================================

Stope:
    Stop := 1
    ToolTip
    LogAction("Script stopped by user")
return

Pauza:
    if (A_IsPaused) {
        Pause, , 1
        ToolTip
        LogAction("Script resumed")
    } else {
        Pause, , 1
        ToolTip, PAUSE
        LogAction("Script paused")
        Sleep, DELAY_SHORT
    }
return

Reset:
    LogAction("Script restart")
    Reload
return

CloseScript:
    LogAction("Script closed by user")
    ExitApp
return

; ============================================================================
; MAIN AUTOMATION FUNCTIONS
; ============================================================================

OpenChests:
    Stop := 0
    CycleCount := 0
    Gui, Submit, NoHide

    if (!ValidateScript()) {
        LogAction("Open chests: validation error")
        return
    }

    LogAction("Starting chest opening. Loop limit: " . (LoopCount != "" ? LoopCount : "infinite"))

    loop {
        if (!CheckActiveWindow())
            continue

        if (Stop == 1 || (LoopCount != "" && CycleCount >= LoopCount))
            break

        SendInput, {Space}
        Sleep, DELAY_LONG
        SendInput, {Space}

        CycleCount++
        UpdateTooltip("Opening chests", CycleCount)
        Sleep, DELAY_LONG
    }

    ; Added sound when cycles complete
    if (LoopCount != "" && CycleCount >= LoopCount) {
        SoundPlay, *64
    }

    ToolTip
    LogAction("Chest opening completed. Total cycles: " . CycleCount)
return

SalvageItems:
    Stop := 0
    CycleCount := 0
    Gui, Submit, NoHide

    if (!ValidateScript(true, "X1,Y1,X2,Y2,X3,Y3,X4,Y4")) {
        LogAction("Item salvage: validation error")
        return
    }

    LogAction("Starting item salvage. Loop limit: " . (LoopCount != "" ? LoopCount : "infinite"))

    loop {
        if (!CheckActiveWindow())
            continue

        if (Stop == 1 || (LoopCount != "" && CycleCount >= LoopCount))
            break

        Sleep, DELAY_LONG

        MouseMove, %X1%, %Y1%
        Click
        Sleep, DELAY_MEDIUM

        MouseMove, %X2%, %Y2%
        Click
        Sleep, DELAY_MEDIUM

        MouseMove, %X3%, %Y3%
        Click
        Sleep, DELAY_MEDIUM

        MouseMove, %X4%, %Y4%
        Click
        Sleep, DELAY_MEDIUM

        Send, {Space down}
        Sleep, 400
        Send, {Space up}
        Sleep, 800

        CycleCount++
        UpdateTooltip("Salvaging items", CycleCount)
        Sleep, DELAY_MEDIUM
    }

    ; Added sound when cycles complete
    if (LoopCount != "" && CycleCount >= LoopCount) {
        SoundPlay, *64
    }

    ToolTip
    LogAction("Item salvage completed. Total cycles: " . CycleCount)
return

SalvageRedItems:
    Stop := 0
    CycleCount := 0
    Gui, Submit, NoHide

    if (!ValidateScript(true, "X1,Y1,X2,Y2,X3,Y3,X4,Y4")) {
        LogAction("Red item salvage: validation error")
        return
    }

    LogAction("Starting red item salvage. Loop limit: " . (LoopCount != "" ? LoopCount : "infinite"))

    loop {
        if (!CheckActiveWindow())
            continue

        if (Stop == 1 || (LoopCount != "" && CycleCount >= LoopCount))
            break

        SendKey("Space")
        Sleep, DELAY_SHORT

        Loop, 4 {
            SendKey("Right")
            SendKey("Space")
        }

        Sleep, DELAY_SHORT
        Send, {Space down}
        Sleep, DELAY_LONG
        Send, {Space up}
        Sleep, 2000

        CycleCount++
        UpdateTooltip("Salvaging red items", CycleCount)
        Sleep, DELAY_SHORT
    }

    ; Added sound when cycles complete
    if (LoopCount != "" && CycleCount >= LoopCount) {
        SoundPlay, *64
    }

    ToolTip
    LogAction("Red item salvage completed. Total cycles: " . CycleCount)
return

UpgradeAtanor:
    Stop := 0
    CycleCount := 0
    Gui, Submit, NoHide

    if (!ValidateScript(true, "X5,Y5")) {
        LogAction("Athanor upgrade: validation error")
        return
    }

    LogAction("Starting athanor upgrade. Loop limit: " . (LoopCount != "" ? LoopCount : "infinite"))

    loop {
        if (!CheckActiveWindow())
            continue

        if (Stop == 1 || (LoopCount != "" && CycleCount >= LoopCount))
            break

        Sleep, 150
        MouseMove, %X5%, %Y5%
        Click

        CycleCount++
        UpdateTooltip("Upgrading athanor", CycleCount)
        Sleep, DELAY_ANIMATION
    }

    ; Added sound when cycles complete
    if (LoopCount != "" && CycleCount >= LoopCount) {
        SoundPlay, *64
    }

    ToolTip
    LogAction("Athanor upgrade completed. Total cycles: " . CycleCount)
return

; ============================================================================
; REROLL PROPERTIES FUNCTION
; ============================================================================

RerollProperties:
    Gui, Submit, NoHide

    if (Target_Window == "") {
        MsgBox, 48, Error, Target window not set. Set the window before running the script.
        LogAction("Reroll properties: window not set")
        return
    }

    if (OCR_X = "" || OCR_Y = "" || OCR_W = "" || OCR_H = "") {
        MsgBox, 16, Error, Item properties area not set!`nSelect area before running.
        LogAction("Reroll properties: item properties area not set")
        return
    }

    if (GreenDust_X1 = "" || GreenDust_Y1 = "" || GreenDust_X2 = "" || GreenDust_Y2 = "") {
        MsgBox, 16, Error, Green dust area not set!`nSelect area before running.
        LogAction("Reroll properties: green dust area not set")
        return
    }

    Menu, RerollScript, Show
return
/*
menuHandler(itemName) {
    global Stop, Target_Window, LoopCount, OCR_X, OCR_Y, OCR_W, OCR_H
    global GreenDust_X1, GreenDust_Y1, GreenDust_X2, GreenDust_Y2
    global SelectedProperties := itemName
    global lastOCRText, lastOCRTime
    global COLOR_GREEN_DUST, COLOR_TOLERANCE, DELAY_OCR_COOLDOWN

    ; Split property names from menu selection
    Properties := StrSplit(itemName, " - ")
    Property1 := Properties[1]
    Property2 := Properties[2]

    LogAction("Starting property reroll: " . itemName)

    Gui, Submit, NoHide
    WinActivate, %Target_Window%

    CycleCount := 0
    Stop := 0

    ; Setup coordinates for OCR and pixel search
    ocrArea := [OCR_X, OCR_Y, OCR_W, OCR_H]
    dustX1 := GreenDust_X1
    dustY1 := GreenDust_Y1
    dustX2 := GreenDust_X2
    dustY2 := GreenDust_Y2

    ToolTip, Starting property reroll...`nLooking for: %Property1% - %Property2%
    Sleep, 1000

    ; Flag to track stop reason
    dustExhausted := false

    loop {
        ; Check if correct window is active
        if (!CheckActiveWindow())
            continue

        ; Check stop conditions
        if (Stop == 1 || (LoopCount != "" && CycleCount >= LoopCount)) {
            ToolTip
            LogAction("Property reroll stopped. Cycles: " . CycleCount)
            break
        }

        ; Check for green dust availability (crafting resource)
        PixelSearch, Px, Py, %dustX1%, %dustY1%, %dustX2%, %dustY2%, %COLOR_GREEN_DUST%, %COLOR_TOLERANCE%, RGB Fast

        ; If dust not found, set flag and exit
        if (ErrorLevel) {
            dustExhausted := true
            ToolTip, Green dust not found!
            Sleep, 3000
            ToolTip
            break
        }

        ; Continue while green dust is available
        While !ErrorLevel {
            ; Throttle OCR calls to prevent performance issues
            if (A_TickCount - lastOCRTime < DELAY_OCR_COOLDOWN) {
                Sleep, 50
                continue
            }

            ; Perform OCR to read item properties
            try {
                text := OCR(ocrArea)
                lastOCRTime := A_TickCount
                lastOCRText := text
            } catch e {
                LogAction("OCR error: " . e.Message)
                text := ""
            }

            ; Display OCR results
            ToolTip, Recognized text:`n%text%`n---`nLooking for: %Property1% - %Property2%`nCycle: %CycleCount%/%LoopCount%
            Sleep, 1000

            ; Check if target properties found
            If (InStr(text, Property1) && InStr(text, Property2)) {
                SoundPlay, *64
                ToolTip, Target properties found!`n%Property1% - %Property2%
                LogAction("Properties found after " . CycleCount . " cycles")
                Sleep, 5000
                ToolTip
                return

            } else {
                ; Reroll properties by pressing Space
                Send, {Space down}
                Sleep, 500
                Send, {Space up}
                Sleep, 1500

                CycleCount++
                UpdateTooltip("Rerolling properties", CycleCount)
                Sleep, DELAY_SHORT
                ; Recheck green dust availability
                ;PixelSearch, Px, Py, %dustX1%, %dustY1%, %dustX2%, %dustY2%, %COLOR_GREEN_DUST%, %COLOR_TOLERANCE%, RGB Fast
            }
        }
    }

    ; If exited loop due to dust exhaustion, show message
    if (dustExhausted && !Stop && (LoopCount = "" || CycleCount < LoopCount)) {
        ToolTip, Script stopped.
        LogAction("Property reroll completed. Cycles: " . CycleCount . " (dust not found)")
        Sleep, 3000
    }

    ToolTip
}
*/

menuHandler(itemName) {
    global Stop, Target_Window, LoopCount, OCR_X, OCR_Y, OCR_W, OCR_H
    global GreenDust_X1, GreenDust_Y1, GreenDust_X2, GreenDust_Y2
    global SelectedProperties := itemName
    global lastOCRText, lastOCRTime
    global COLOR_GREEN_DUST, COLOR_TOLERANCE, DELAY_OCR_COOLDOWN

    ; Split property names from menu selection
    Properties := StrSplit(itemName, " - ")
    Property1 := Properties[1]
    Property2 := Properties[2]

    LogAction("Starting property reroll: " . itemName)

    Gui, Submit, NoHide
    WinActivate, %Target_Window%

    CycleCount := 0
    Stop := 0

    ; Setup coordinates for OCR and pixel search
    ocrArea := [OCR_X, OCR_Y, OCR_W, OCR_H]
    dustX1 := GreenDust_X1
    dustY1 := GreenDust_Y1
    dustX2 := GreenDust_X2
    dustY2 := GreenDust_Y2

    ToolTip, Starting property reroll...`nLooking for: %Property1% - %Property2%
    Sleep, 1000


    loop {

        ; Check stop conditions
        if (Stop == 1 || (LoopCount != "" && CycleCount >= LoopCount)) {
            ToolTip
            LogAction("Property reroll stopped. Cycles: " . CycleCount)
            break
        }

        ; Check if correct window is active
        if (!CheckActiveWindow())
            continue

        ; Check for green dust availability (crafting resource)
        PixelSearch, Px, Py, %dustX1%, %dustY1%, %dustX2%, %dustY2%, %COLOR_GREEN_DUST%, %COLOR_TOLERANCE%, RGB Fast

        ; If dust not found, set flag and exit
        if (ErrorLevel) {
            ToolTip, Green dust not found! Script stopped
            Sleep, 3000
            ToolTip
            break
        }

        ; OCR чтение свойств предмета
        if (A_TickCount - lastOCRTime >= DELAY_OCR_COOLDOWN) {
            try {
                text := OCR(ocrArea)
                lastOCRTime := A_TickCount
                lastOCRText := text
            } catch e {
                LogAction("OCR error: " . e.Message)
                text := ""
            }

            ; Display OCR results
            ToolTip, Recognized text:`n%text%`n---`nLooking for: %Property1% - %Property2%`nCycle: %CycleCount%/%LoopCount%
            Sleep, 1000

            ; Check if target properties found
            If (InStr(text, Property1) && InStr(text, Property2)) {
                SoundPlay, *64
                ToolTip, Target properties found!`n%Property1% - %Property2%
                LogAction("Properties found after " . CycleCount . " cycles")
                Sleep, 5000
                ToolTip
                return
            }
        }

        ; Reroll properties by pressing Space
        Send, {Space down}
        Sleep, 700
        Send, {Space up}
        Sleep, 2000

        CycleCount++
        UpdateTooltip("Rerolling properties", CycleCount)
        Sleep, DELAY_SHORT

    }

    ToolTip
}
; ============================================================================
; PRESET MANAGEMENT FUNCTIONS
; ============================================================================

PresetChange:
    Gui, Submit, NoHide
    if (savedPreset = "")
        return

    ;LogAction("Loading preset: " . savedPreset)

    ; Removed try-catch for better error diagnostics
    Loop, 5 {
        IniRead, xVal, %presetFile%, %savedPreset%, Pos_%A_Index%_X
        IniRead, yVal, %presetFile%, %savedPreset%, Pos_%A_Index%_Y
        if (xVal != "ERROR" && yVal != "ERROR") {
            GuiControl,, Pos_%A_Index%, %xVal% %yVal%
            X%A_Index% := xVal
            Y%A_Index% := yVal
        }
    }

    IniRead, ocrX, %presetFile%, %savedPreset%, OCR_X
    IniRead, ocrY, %presetFile%, %savedPreset%, OCR_Y
    IniRead, ocrW, %presetFile%, %savedPreset%, OCR_W
    IniRead, ocrH, %presetFile%, %savedPreset%, OCR_H
    if (ocrX != "ERROR" && ocrY != "ERROR" && ocrW != "ERROR" && ocrH != "ERROR") {
        OCR_X := ocrX
        OCR_Y := ocrY
        OCR_W := ocrW
        OCR_H := ocrH
        GuiControl,, OCR_Pos, %ocrX% %ocrY% %ocrW% %ocrH%
    }

    IniRead, gdX1, %presetFile%, %savedPreset%, GreenDust_X1
    IniRead, gdY1, %presetFile%, %savedPreset%, GreenDust_Y1
    IniRead, gdX2, %presetFile%, %savedPreset%, GreenDust_X2
    IniRead, gdY2, %presetFile%, %savedPreset%, GreenDust_Y2
    if (gdX1 != "ERROR" && gdY1 != "ERROR" && gdX2 != "ERROR" && gdY2 != "ERROR") {
        GreenDust_X1 := gdX1
        GreenDust_Y1 := gdY1
        GreenDust_X2 := gdX2
        GreenDust_Y2 := gdY2
        GuiControl,, GreenDust_Pos, %gdX1% %gdY1% %gdX2% %gdY2%
    }

    IniRead, twVal, %presetFile%, %savedPreset%, Target_Window
    if (twVal != "ERROR") {
        GuiControl,, Target_Window, %twVal%
        Target_Window := twVal
    }

    IniRead, lcVal, %presetFile%, %savedPreset%, LoopCount
    if (lcVal != "ERROR") {
        GuiControl,, LoopCount, %lcVal%
        LoopCount := lcVal
    }

    Gui, Submit, NoHide
    Loop, Parse, % "Stop_key|Pause_key|Reset_key|OpenChests_key|Salvage_items_key|Salvage_red_items_key|Reroll_Properties_key|Atanor_key|Close_script_key", |
    {
        IniRead, val, %presetFile%, %savedPreset%, %A_LoopField%
        if (val != "ERROR") {
            %A_LoopField% := val
            GuiControl,, %A_LoopField%, %val%
        }
    }

    UpdateAllKeys()
    ;SB_SetText("Preset '" . savedPreset . "' loaded")

return

SavePreset:
    Gui, Submit, NoHide

    if (savedPreset = "") {
        SB_SetText("Enter preset name!")
        return
    }

    LogAction("Saving preset: " . savedPreset)

    Loop, 5 {
        GuiControlGet, pos,, Pos_%A_Index%
        if (pos != "") {
            StringSplit, coord, pos, %A_Space%
            IniWrite, %coord1%, %presetFile%, %savedPreset%, Pos_%A_Index%_X
            IniWrite, %coord2%, %presetFile%, %savedPreset%, Pos_%A_Index%_Y
        }
    }

    if (OCR_X != "" && OCR_Y != "" && OCR_W != "" && OCR_H != "") {
        IniWrite, %OCR_X%, %presetFile%, %savedPreset%, OCR_X
        IniWrite, %OCR_Y%, %presetFile%, %savedPreset%, OCR_Y
        IniWrite, %OCR_W%, %presetFile%, %savedPreset%, OCR_W
        IniWrite, %OCR_H%, %presetFile%, %savedPreset%, OCR_H
    }

    if (GreenDust_X1 != "" && GreenDust_Y1 != "" && GreenDust_X2 != "" && GreenDust_Y2 != "") {
        IniWrite, %GreenDust_X1%, %presetFile%, %savedPreset%, GreenDust_X1
        IniWrite, %GreenDust_Y1%, %presetFile%, %savedPreset%, GreenDust_Y1
        IniWrite, %GreenDust_X2%, %presetFile%, %savedPreset%, GreenDust_X2
        IniWrite, %GreenDust_Y2%, %presetFile%, %savedPreset%, GreenDust_Y2
    }

    if (Target_Window != "")
        IniWrite, %Target_Window%, %presetFile%, %savedPreset%, Target_Window

    if (LoopCount != "")
        IniWrite, %LoopCount%, %presetFile%, %savedPreset%, LoopCount

    ; Save hotkeys
    IniWrite, %Stop_key%, %presetFile%, %savedPreset%, Stop_key
    IniWrite, %Pause_key%, %presetFile%, %savedPreset%, Pause_key
    IniWrite, %Reset_key%, %presetFile%, %savedPreset%, Reset_key
    IniWrite, %OpenChests_key%, %presetFile%, %savedPreset%, OpenChests_key
    IniWrite, %Salvage_items_key%, %presetFile%, %savedPreset%, Salvage_items_key
    IniWrite, %Salvage_red_items_key%, %presetFile%, %savedPreset%, Salvage_red_items_key
    IniWrite, %Reroll_Properties_key%, %presetFile%, %savedPreset%, Reroll_Properties_key
    IniWrite, %Atanor_key%, %presetFile%, %savedPreset%, Atanor_key
    IniWrite, %Close_script_key%, %presetFile%, %savedPreset%, Close_script_key

    SB_SetText("Preset '" . savedPreset . "' saved!")
    GoSub, UpdatePresetList

return

DeletePreset:
    Gui, Submit, NoHide

    if (savedPreset = "") {
        SB_SetText("Select preset to delete")
        return
    }

    MsgBox, 4, Confirmation, Are you sure you want to delete preset "%savedPreset%"?
    IfMsgBox, Yes
    {
        IniDelete, %presetFile%, %savedPreset%
        SB_SetText("Preset '" . savedPreset . "' deleted")
        LogAction("Preset deleted: " . savedPreset)
        GoSub, UpdatePresetList
    }
return

UpdatePresetList:
    IniRead, sectionNames, %presetFile%

    if (sectionNames = "" || sectionNames = "ERROR") {
        GuiControl,, savedPreset, |
        return
    }

    sectionNames := RegExReplace(sectionNames, "(\R|$)", "|")
    GuiControl,, savedPreset, |%sectionNames%
return

; ============================================================================
; GUI EVENT HANDLERS
; ============================================================================

GuiClose:
    Gui, Submit, NoHide
    LogAction("Application closed")
    ExitApp
return

InfoWindowGuiClose:
    Gui, Destroy
return