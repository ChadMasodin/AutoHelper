﻿VERSION := 1.0
RegExMatch(A_ScriptName, "^(.*?)\.", basename)
WINTITLE := basename1 " " VERSION

presetsDir := A_AppData "\" basename1
if !FileExist(presetsDir)
	FileCreateDir, %presetsDir%

#SingleInstance force
#NoENV
SetBatchLines -1 ; Run script at maximum speed
ListLines Off ; Disable debug logging
SetMouseDelay, -1 ; Remove mouse delays
CoordMode, Mouse, Client
SetTitleMatchMode, 2
SetControlDelay, -1
DetectHiddenWindows, On
SetStoreCapsLockMode, Off

;---------------------- GLOBAL VARIABLES ----------------------
global X1, Y1, X2, Y2, X3, Y3, X4, Y4, X5, Y5, X6, Y6, X7, Y7
global LoopCount
Target_Window := ""
global Stop:=0
LoopCount := ""
Stop_key := "F1", Pause_key := "F2", Reset_key := "F3", OpenChests_key := "F5", Salvage_items_key := "F7", Salvage_red_items_key := "F9", Dust_convert_key := "F11", Atanor_key := "F12", Close_script_key := "^ESC"

;====================== UPDATE ALL HOTKEYS ======================
UpdateAllKeys()
{
    global
    Hotkey, %Stop_key%, Stope, On
    Hotkey, %Pause_key%, Pauza, On
    Hotkey, %Reset_key%, Reset, On
    Hotkey, %OpenChests_key%, OpenChests, On
    Hotkey, %Salvage_items_key%, Salvage_items, On
    Hotkey, %Salvage_red_items_key%, Salvage_red_items, On
    Hotkey, %Dust_convert_key%, Dust_convert, On
    Hotkey, %Atanor_key%, Upgrade_atanor, On
    Hotkey, %Close_script_key%, Close_script, On
}
;---------------------- CREATE GUI ----------------------
Gui,+AlwaysOnTop -DPIScale +OwnDialogs

;---------------------- Preset Panel ----------------------
Gui, Add, Text, xm section, Preset List:
Gui, Add, ComboBox, x+5 vfrmSAVEDPRESET gPresetChange
Gui, Add, Button, x+5 h21 w70 gSavePreset, Save
Gui, Add, Button, x+5 h21 w70 gDeletePreset vDELETEBUTTON, Delete
Gui, Add, StatusBar

;------------------- Coordinate Setup ----------------------
Gui, Font, c535353 s13 Bold Q3, Segoe UI Black
Gui, Add, GroupBox, x10 y40 w380 h260 Center Section, Script Settings
Gui, Font
Gui, Font, s11, Segoe UI
Gui, Add, Text, xs+10 yp+25, #1 Button position for opening chests:
Gui, Font, s9,
Gui, Add, Button, xs+10 yp+30 w80 h25 gSet_Pos_Chests, Set
Gui, Add, Edit, x+10 yp+1 w80 h20 vPos_1,
Gui, Font, s11, Segoe UI
Gui, Add, Text, xs+10 yp+30, #2 Button position for upgrade Athanor:
Gui, Font, s9
Gui, Add, Button, xs+10 yp+30 w80 h25 gSet_Pos_Atanor, Set
Gui, Add, Edit, x+10 yp+1 w80 h20 vPos_7,
Gui, Font, s11, Segoe UI
Gui,Add,Text,xs+10 yp+30,#3 Button positions for salvage items:

Gui, Font, s9
Gui,Add,Button,xs+10 yp+25 w140 h25 gSet_Pos_Button, Salvage Button
Gui,Add,Edit,x+10 yp+3 w140 h20 vPos_6,
Gui,Add,Button,xs+10 yp+30 w70 h20 gSet_Pos_White, White
Gui,Add,Edit,x+10 w80 h20 vPos_2,
Gui,Add,Button,x+10 w70 h20 gSet_Pos_Green, Green
Gui,Add,Edit,x+10 w80 h20 vPos_3,
Gui,Add,Button,xs+10 yp+25 w70 h20 gSet_Pos_Blue, Blue
Gui,Add,Edit,x+10 w80 h20 vPos_4,
Gui,Add,Button,x+10 w70 h20 gSet_Pos_Yellow, Yellow
Gui,Add,Edit,x+10 w80 h20 vPos_5,

;--------------------- Additional Settings ---------------
Gui, Add, Text, x70 y690, General Loop Limiter (optional):
Gui, Add, Edit, x+10 w80 h20 vLoopCount

Gui, Add, Button, x25 y716 w120 h30 gSet_Location, Set Target Window:
Gui, Add, Edit, x+15 yp+3 w210 h25 vTarget_Window, %Target_Window%

;------------------------ Action Buttons ---------------------
Gui, Font, c535353 s13 Bold Q3, Segoe UI Black
Gui, Add, GroupBox, x10 y295 w380 h200 Center Section, Buttons

Gui, Font, s9 Bold, Segoe UI
Gui, Add, Button, xs+15 yp+30 w110 h35 gOpenChests, OPEN CHESTS
Gui, Add, Button, x+10 w110 h35 gSalvage_items, SALVAGE ITEMS
Gui, Add, Button, x+10 w110 h35 gUpgrade_atanor, UPGRADE ATHANOR
Gui, Add, Button, xs+15 yp+45 w170 h35 gSalvage_red_items, SALVAGE RED ITEMS
Gui, Add, Button, x+10 w170 h35 gDust_convert, CRAFTING MECHANICS
Gui, Add, Button, xs+40 yp+45 w300 h25 gTarget_Info, <><><> HOTKEYS AND INFO <><><>
Gui, Add, Button, xs+15 yp+35 w110 h35 gStope, STOP
Gui, Add, Button, x+10 w110 h35 gPauza, PAUSE
Gui, Add, Button, x+10 w110 h35 gReset, RESTART
Gui, Font

Gui, Font, c666666 s12 Bold Q3, Segoe UI Black
Gui, Add, GroupBox, x10 y500 w380 h180 Center Section, Default Hotkeys
Gui, Font
Gui, Font, s12
Gui, Add, Text, cBlue xs+30 yp+30, Stop = F1
Gui, Add, Text, cBlue x+15, Pause = F2
Gui, Add, Text, cBlue x+15, Restart / Reset = F3
Gui, Add, Text, cRed xs+50 yp+30, Open Chests = F5
Gui, Add, Text, cRed x+10, Salvage Items = F7
Gui, Add, Text, cRed xs+100 yp+30, Salvage Red Items = F9
Gui, Add, Text, cRed xs+100 yp+30, Crafting Mechanics = F11
Gui, Add, Text, cRed xs+30 yp+30, Upgrade Atanor = F12
Gui, Add, Text, cMaroon x+10, Close Script = Ctrl+Esc
Gui, Font
Gui, Show, w400 h780, %WINTITLE%
GoSub, UpdatePresetList
return
;────────────────────────────────────────────────────────────
; Help Window (Target_Info)
;────────────────────────────────────────────────────────────
Target_Info:
Gui, NewWindow:New , , Help
Gui, +AlwaysOnTop -DPIScale +Owner1
Gui, Add, Picture, x10 y33, icon.png
Gui, Font, c666666 s11 Bold Q3, Segoe UI Black
Gui, Add, GroupBox, x10 y10 w280 h150 Center Section, INFO
Gui, Font
Gui, Add, Text, xs+145 yp+25, Author: !ChadMasodin
Gui, Add, Text, xs+145 yp+15, Release: 27/12/24
Gui, Add, Text, xs+145 yp+15, Update: 21/02/25
Gui, Add, Link, xs+10 yp+35, AutoHelper - is a script for Vermintide 2, that automates`na number of routine processes in the game.`nMore information about it can be found in this <a href="https://steamcommunity.com/sharedfiles/filedetails/?id=3385048068\">guide.</a>
Gui, Font, s11 Bold Q3, Segoe UI
Gui, Add, Button, x10 y480 w280 h45 gSaveTargetInfo, OK
Gui, Font
;────────────────────────────────────────────────────────────
; Hotkey Binding GUI
;────────────────────────────────────────────────────────────
Gui, Font, c666666 s11 Bold Q3, Segoe UI Black
Gui, Add, GroupBox, x10 y170 w280 h300 Center Section, HOTKEYS
Gui, Font

Gui, Add, Text, cBlue xs+10 yp+30, Stop:
Gui, Add, Hotkey, x+10 w90 h20 vStop_key, %Stop_key%
Gui, Add, Text, cBlue xs+10 yp+30, Pause:
Gui, Add, Hotkey, x+10 w90 h22 vPause_key, %Pause_key%
Gui, Add, Text, cBlue xs+10 yp+30, Restart/Reset:
Gui, Add, Hotkey, x+10 w90 h22 vReset_key, %Reset_key%
Gui, Add, Text, cRed xs+10 yp+30, Open Chests:
Gui, Add, Hotkey, x+10 w90 h20 vOpenChests_key, %OpenChests_key%
Gui, Add, Text, cRed xs+10 yp+30, Salvage Items:
Gui, Add, Hotkey, x+10 w90 h20 vSalvage_items_key, %Salvage_items_key%
Gui, Add, Text, cRed xs+10 yp+30, Salvage Red Items:
Gui, Add, Hotkey, x+10 w80 h20 vSalvage_red_items_key, %Salvage_red_items_key%
Gui, Add, Text, cRed xs+10 yp+30, Dust Convert:
Gui, Add, Hotkey, x+10 w90 h20 vDust_convert_key, %Dust_convert_key%
Gui, Add, Text, cRed xs+10 yp+30, Upgrade Atanor:
Gui, Add, Hotkey, x+10 w90 h20 vAtanor_Key, %Atanor_key%
Gui, Add, Text, cMaroon xs+10 yp+30, Close Script:
Gui, Add, Hotkey, x+10 w90 h20 vClose_script_key, %Close_script_key%
Gui, Show, w300 h540,
return

;────────────────────────────────────────────────────────────
; Event handlers for window setup and GUI updates
;────────────────────────────────────────────────────────────
Set_Location:
Target_Window := Set_Window(Target_Window)
GuiControl,, Target_Window, %Target_Window%
return

UpdateKey(KeyVariable, Label)
{
Global
GuiControlGet, New_Key, , %KeyVariable%
Current_Key := %KeyVariable%
if (New_Key != Current_Key)
{
Hotkey, %Current_Key%, %Label%, Off
%KeyVariable% := New_Key
Hotkey, %New_Key%, %Label%, On
if (frmSAVEDPRESET != "")
IniWrite, %New_Key%, %A_AppData%\%basename1%\presets.ini, %frmSAVEDPRESET%, %KeyVariable%
}
}

SendKey(key, delay := 300) {
Send, {%key% down}
Send, {%key% up}
Sleep, delay
}

;────────────────────────────────────────────────────────────
; Functions to set click coordinates
;────────────────────────────────────────────────────────────
Set_Pos_Chests:
Stop := 1
Get_Click_Pos(X1, Y1)
GuiControl,, Pos_1, %X1% %Y1%
return

Set_Pos_White:
Stop := 1
Get_Click_Pos(X2, Y2)
GuiControl,, Pos_2, %X2% %Y2%
return

Set_Pos_Green:
Stop := 1
Get_Click_Pos(X3, Y3)
GuiControl,, Pos_3, %X3% %Y3%
return

Set_Pos_Blue:
Stop := 1
Get_Click_Pos(X4, Y4)
GuiControl,, Pos_4, %X4% %Y4%
return

Set_Pos_Yellow:
Stop := 1
Get_Click_Pos(X5, Y5)
GuiControl,, Pos_5, %X5% %Y5%
return

Set_Pos_Button:
Stop := 1
Get_Click_Pos(X6, Y6)
GuiControl,, Pos_6, %X6% %Y6%
return

Set_Pos_Atanor:
Stop := 1
Get_Click_Pos(X7, Y7)
GuiControl,, Pos_7, %X7% %Y7%
return

;────────────────────────────────────────────────────────────
; Script control functions
;────────────────────────────────────────────────────────────
Stope:
Stop := 1
return

Pauza:
if (A_IsPaused)
{
Pause, , 1
Tooltip
}
else
{
Pause, , 1
Tooltip, PAUSED
Sleep, 100
}
return

Reset:
Reload
return

;────────────────────────────────────────────────────────────
; Open Chests automation
;────────────────────────────────────────────────────────────
OpenChests:
Stop := 0
CycleCount := 0
Gui, Submit, NoHide

if (LoopCount != "" && LoopCount < 0) {
    MsgBox, 16, Error, Please enter a positive number or leave empty for infinite loops!
    return
}
if (Target_Window == "") {
    MsgBox, 48, Error, Target window not set!
    return
}
if (!X1 || !Y1) {
    MsgBox, 16, Error, Coordinates not set! Set X and Y first.
    return
}
loop {
    WinGetActiveTitle, Current_Window
    if (Current_Window != Target_Window) {
        ToolTip, Script paused. Activate window: %Target_Window%
        Sleep, 500
        continue
    }
    if (Stop == 1 || (LoopCount != "" && CycleCount >= LoopCount))
        break
    Sleep, 50
    MouseMove, %X1%, %Y1%
    Click
    Sleep, 500
    Send, {Space}
    CycleCount++
    ToolTip, Running... 555: %CycleCount%/%LoopCount%
}
ToolTip

return

;────────────────────────────────────────────────────────────
; Salvage items
;────────────────────────────────────────────────────────────
Salvage_items:
Stop := 0
CycleCount := 0
Gui, Submit, NoHide
if (LoopCount != "" && LoopCount < 0) {
MsgBox, 16, Error, Please enter a positive number or leave empty for infinite loops!
return
}
if (Target_Window == "") {
MsgBox, 48, Error, Target window not set!
return
}
if ((!X2 || !Y2) || (!X3 || !Y3) || (!X4 || !Y4) || (!X5 || !Y5) || (!X6 || !Y6)) {
MsgBox, 16, Error, Coordinates not set! Set X and Y first.
return
}
loop {
WinGetActiveTitle, Current_Window
if (Current_Window != Target_Window) {
ToolTip, Script paused. Activate window: %Target_Window%
Sleep, 500
continue
}
if (Stop == 1 || (LoopCount != "" && CycleCount >= LoopCount))
break
Sleep, 500
MouseMove, %X2%, %Y2% ; White
Click
Sleep, 300
MouseMove, %X3%, %Y3% ; Green
Click
Sleep, 300
MouseMove, %X4%, %Y4% ; Blue
Click
Sleep, 300
MouseMove, %X5%, %Y5% ; Yellow
Click
Sleep, 300
MouseMove, %X6%, %Y6% ; Salvage button
Click, down
Sleep, 500
Click, up
Sleep, 1000
CycleCount++
ToolTip, Running... Loop: %CycleCount%/%LoopCount%
Sleep, 200
}
ToolTip
return

;────────────────────────────────────────────────────────────
; Salvage red items
;────────────────────────────────────────────────────────────
Salvage_red_items:
{
Stop := 0
CycleCount := 0
Gui, Submit, NoHide
if (LoopCount != "" && LoopCount < 0) {
MsgBox, 16, Error, Please enter a positive number or leave empty for infinite loops!
return
}
if (Target_Window == "") {
MsgBox, 48, Error, Target window not set!
return
}
loop {
WinGetActiveTitle, Current_Window
if (Current_Window != Target_Window) {
ToolTip, Script paused. Activate window: %Target_Window%
Sleep, 500
continue
}
if (Stop == 1 || (LoopCount != "" && CycleCount >= LoopCount))
break
SendKey("Space")
Sleep, 100
Loop, 4 {
SendKey("Right")
SendKey("Space")
}
Sleep, 100
Send, {Space down}
Sleep, 500
Send, {Space up}
Sleep, 2000
CycleCount++
ToolTip, Running... Loop: %CycleCount%/%LoopCount%
Sleep, 100
}
}
ToolTip
return
;────────────────────────────────────────────────────────────
; Upgrade Atanor
;────────────────────────────────────────────────────────────
Upgrade_atanor:
Stop := 0
CycleCount := 0
Gui, Submit, NoHide
if (LoopCount != "" && LoopCount < 0) {
MsgBox, 16, Error, Please enter a positive number or leave empty for infinite loops!
return
}
if (Target_Window == "") {
MsgBox, 48, Error, Target window not set!
return
}
if (!X7 || !Y7) {
MsgBox, 16, Error, Coordinates not set! Set X and Y first.
return
}
loop {
WinGetActiveTitle, Current_Window
if (Current_Window != Target_Window) {
ToolTip, Script paused. Activate window: %Target_Window%
Sleep, 500
continue
}
if (Stop == 1 || (LoopCount != "" && CycleCount >= LoopCount))
break
Sleep, 150
MouseMove, %X7%, %Y7%
Click
CycleCount++
ToolTip, Running... Loop: %CycleCount%/%LoopCount%
Sleep, 1500
}
ToolTip
return

;────────────────────────────────────────────────────────────
; Dust convers
;────────────────────────────────────────────────────────────
Dust_convert:
Stop := 0
CycleCount := 0
Gui, Submit, NoHide
if (LoopCount != "" && LoopCount < 0) {
MsgBox, 16, Error, Please enter a positive number or leave empty for infinite loops!
return
}
if (Target_Window == "") {
MsgBox, 48, Error, Target window not set!
return
}
Send, {Space}
loop {
WinGetActiveTitle, Current_Window
if (Current_Window != Target_Window) {
ToolTip, Script paused. Activate window: %Target_Window%
Sleep, 500
continue
}
if (Stop == 1 || (LoopCount != "" && CycleCount >= LoopCount))
break
Sleep, 200
Send, {Space down}
CycleCount++
ToolTip, Running... Loop: %CycleCount%/%LoopCount%
Sleep, 1750
}
ToolTip
return

;────────────────────────────────────────────────────────────
; Close script
;────────────────────────────────────────────────────────────
Close_script:
ExitApp
return

UpdateAllKeys()
Gui, Submit, NoHide
return

;────────────────────────────────────────────────────────────
; Get click coordinates
;────────────────────────────────────────────────────────────
Get_Click_Pos(ByRef X, ByRef Y)
{
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
    ToolTip, Click LMB to Set Pos`nPress ESC to cancel`n`nCurrent coordinates: X=%currentX% Y=%currentY%

    if (Left_Mouse == False && isPressed == 0) {
        isPressed := 1
    }
    else if (Left_Mouse == True && isPressed == 1) {
        X := currentX
        Y := currentY
        ToolTip
        break
    }
}

}

;────────────────────────────────────────────────────────────
; Set target window
;────────────────────────────────────────────────────────────
Set_Window(Target_Window)
{
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
    ToolTip, Double-click LMB on target window`nPress ESC to cancel`n`nCurrent Window: %Temp_Window%

    if (Left_Mouse == False && isPressed == 0) {
        isPressed := 1
    }
    else if (Left_Mouse == True && isPressed == 1) {
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

;============================================================
; Preset management
;============================================================
PresetChange:
    Gui, Submit, NoHide
    if (frmSAVEDPRESET = "")
    return

Loop, 7 {
    IniRead, xVal, %A_AppData%\%basename1%\presets.ini, %frmSAVEDPRESET%, Pos_%A_Index%_X
    IniRead, yVal, %A_AppData%\%basename1%\presets.ini, %frmSAVEDPRESET%, Pos_%A_Index%_Y
    if (xVal != "ERROR" && yVal != "ERROR") {
        GuiControl,, Pos_%A_Index%, %xVal% %yVal%
        X%A_Index% := xVal
        Y%A_Index% := yVal
    }
}

IniRead, twVal, %A_AppData%\%basename1%\presets.ini, %frmSAVEDPRESET%, Target_Window
if (twVal != "ERROR") {
    GuiControl,, Target_Window, %twVal%
    Target_Window := twVal
}

IniRead, lcVal, %A_AppData%\%basename1%\presets.ini, %frmSAVEDPRESET%, LoopCount
if (lcVal != "ERROR") {
    GuiControl,, LoopCount, %lcVal%
    LoopCount := lcVal
}

Loop, Parse, % "Stop_key|Pause_key|Reset_key|OpenChests_key|Salvage_items_key|Salvage_red_items_key|Dust_convert_key|Atanor_key|Close_script_key", |
{
IniRead, val, %A_AppData%\%basename1%\presets.ini, %frmSAVEDPRESET%, %A_LoopField%
    if (val != "ERROR") {
        %A_LoopField% := val
        GuiControl,, %A_LoopField%, %val%
    }
}
UpdateAllKeys()
return
;============================================================
; Save preset management
;============================================================
SavePreset:
    Gui, Submit, NoHide
    if (frmSAVEDPRESET = "") {
        SB_SetText("Enter preset name!")
        return
    }
        Loop, 7 {
        GuiControlGet, pos,, Pos_%A_Index%
        if (pos != "") {
        StringSplit, coord, pos, %A_Space%
        IniWrite, %coord1%, %A_AppData%\%basename1%\presets.ini, %frmSAVEDPRESET%, Pos_%A_Index%_X
        IniWrite, %coord2%, %A_AppData%\%basename1%\presets.ini, %frmSAVEDPRESET%, Pos_%A_Index%_Y
    }
}

if (Target_Window != "")
    IniWrite, %Target_Window%, %A_AppData%\%basename1%\presets.ini, %frmSAVEDPRESET%, Target_Window

if (LoopCount != "")
    IniWrite, %LoopCount%, %A_AppData%\%basename1%\presets.ini, %frmSAVEDPRESET%, LoopCount

IniWrite, %Stop_key%, %A_AppData%\%basename1%\presets.ini, %frmSAVEDPRESET%, Stop_key
IniWrite, %Pause_key%, %A_AppData%\%basename1%\presets.ini, %frmSAVEDPRESET%, Pause_key
IniWrite, %Reset_key%, %A_AppData%\%basename1%\presets.ini, %frmSAVEDPRESET%, Reset_key
IniWrite, %OpenChests_key%, %A_AppData%\%basename1%\presets.ini, %frmSAVEDPRESET%, OpenChests_key
IniWrite, %Salvage_items_key%, %A_AppData%\%basename1%\presets.ini, %frmSAVEDPRESET%, Salvage_items_key
IniWrite, %Salvage_red_items_key%, %A_AppData%\%basename1%\presets.ini, %frmSAVEDPRESET%, Salvage_red_items_key
IniWrite, %Dust_convert_key%, %A_AppData%\%basename1%\presets.ini, %frmSAVEDPRESET%, Dust_convert_key
IniWrite, %Atanor_key%, %A_AppData%\%basename1%\presets.ini, %frmSAVEDPRESET%, Atanor_key
IniWrite, %Close_script_key%, %A_AppData%\%basename1%\presets.ini, %frmSAVEDPRESET%, Close_script_key

SB_SetText("Preset '" frmSAVEDPRESET "' saved!")
GoSub, UpdatePresetList
return
;============================================================
; Delete preset management
;============================================================
DeletePreset:
gui, submit, nohide
RegExMatch(A_ScriptName, "^(.*?)\.", basename)
if (frmSAVEDPRESET = "") {
SB_SetText("Preset name required")
return
}
IniDelete, %A_AppData%\%basename1%\presets.ini, %frmSAVEDPRESET%
SB_SetText("Preset '" frmSAVEDPRESET "' deleted!")
GoSub, UpdatePresetList
Return

UpdatePresetList:
    RegExMatch(A_ScriptName, "^(.*?)\.", basename)
    IniRead, sectionNames, %A_AppData%\%basename1%\presets.ini

    if (sectionNames = "") {
    GuiControl,, frmSAVEDPRESET, |
    return
    }

sectionNames := RegExReplace(sectionNames, "(\R|$)", "|")
GuiControl,, frmSAVEDPRESET, |%sectionNames%
return

GuiClose:
    Gui, Submit, NoHide
ExitApp

;============================================================
; Save/Restore GUI settings
;============================================================
GuiSave(inifile,section,begin="",end="")
{
SplitPath, inifile, file, path, ext, base, drive

if (path = "") {
    RegExMatch(A_ScriptName, "^(.*?)\.", basename)
    inifile := A_AppData "\" basename1 "\" inifile
}

WinGet, List_controls, ControlList, A

if (begin = "")
    flag := 0
else
    flag := 1

Loop, Parse, List_controls, `n
{
    GuiControlGet, textvalue,,%A_Loopfield%,Text
    GuiControlGet, vname, Name, %A_Loopfield%

    If (vname = "")
        continue

    if (begin = vname) {
        flag := 0
        continue
    }

    if (flag)
        continue

    if (end = vname)
        break

    GuiControlGet, value ,, %A_Loopfield%
    value := RegExReplace(value, "`n", "|")
    IniWrite, % value, %inifile%, %section%, %vname%
}

return
}

GuiRestore(inifile,section)
{
SplitPath, inifile, file, path, ext, base, drive

if (path = "") {
    RegExMatch(A_ScriptName, "^(.*?)\.", basename)
    inifile := A_AppData "\" basename1 "\" inifile
}

WinGet, List_controls, ControlList, A

Loop, Parse, List_controls, `n
{
    GuiControlGet, vname, Name, %A_Loopfield%
    GuiControlGet, value ,, %A_Loopfield%

    If (vname = "")
        continue

    IniRead, value, %inifile%, %section%, %vname%, ERROR

    if (value != "ERROR") {
        value := RegExReplace(value, "\|", "`n")
        RegExMatch( A_Loopfield, "(.*?)\d+", name)
        if (name1 = "ComboBox") {
            GuiControl, ChooseString, %A_Loopfield%, %value%
        } else {
            GuiControl,  ,%A_Loopfield%, %value%
        }
    }
}
return

}

SaveTargetInfo:
Gui, Submit
UpdateAllKeys()
if (frmSAVEDPRESET != "")
{
Loop, Parse, % "Stop_key|Pause_key|Reset_key|OpenChests_key|Salvage_items_key|Salvage_red_items_key|Dust_convert_key|Atanor_key|Close_script_key", |
IniWrite, % %A_LoopField%, %A_AppData%\%basename1%\presets.ini, %frmSAVEDPRESET%, %A_LoopField%
}
Gui, Destroy
return