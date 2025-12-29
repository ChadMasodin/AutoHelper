/*
===============================================================================
AutoHelper - Vermintide 2 Automation Script
===============================================================================
Author: ChadMasodin
Version: 2.3
Language: Multilingual (Default: English)
Date: 29/12/2025
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
CoordMode, Mouse, Screen
SetTitleMatchMode, 2
SetControlDelay, -1
DetectHiddenWindows, On
SetStoreCapsLockMode, Off

; ============================================================================
; CONSTANTS
; ============================================================================

global VERSION := "2.3"
global DELAY_SHORT := 100
global DELAY_MEDIUM := 300
global DELAY_LONG := 500
global DELAY_ANIMATION := 1500

; ============================================================================
; GLOBAL VARIABLES - COORDINATES
; ============================================================================

global X1 := "", Y1 := ""  ; White button
global X2 := "", Y2 := ""  ; Green button
global X3 := "", Y3 := ""  ; Blue button
global X4 := "", Y4 := ""  ; Yellow button
global X5 := "", Y5 := ""  ; Atanor button
global X6 := "", Y6 := ""  ; Button
global X7 := "", Y7 := ""  ; AfterReroll

global OCR_X := "", OCR_Y := "", OCR_W := "", OCR_H := ""

; ============================================================================
; GLOBAL VARIABLES - SCRIPT STATE
; ============================================================================

global Stop := 0
global Target_Window := ""
global LoopCount := ""
global SelectedProperties := ""
global savedPreset := ""

; ============================================================================
; GLOBAL VARIABLES - LOCALIZATION
; ============================================================================

global CurrentLanguage := "English"
global OCR_Language := "eng"
global SupportedLanguages := "English|French|Italian|German|Spanish|Polish|Portuguese-Brazil|Russian|Chinese-Simplified"

; Language Codes Map for Tesseract
global LangCodes := {}
LangCodes["English"] := "eng"
LangCodes["French"] := "fra"
LangCodes["Italian"] := "ita"
LangCodes["German"] := "deu_latf"
LangCodes["Spanish"] := "spa"
LangCodes["Polish"] := "pol"
LangCodes["Portuguese-Brazil"] := "por"
LangCodes["Russian"] := "rus"
LangCodes["Chinese-Simplified"] := "chi_sim"

; UI Strings Variables (Global declaration)
global TXT_LangLabel, TXT_PresetList, TXT_Save, TXT_Delete, TXT_SettingsTitle
global TXT_PosAtanor, TXT_Set, TXT_PosSalvage, TXT_SalvageBtn
global TXT_White, TXT_Blue, TXT_Green, TXT_Yellow
global TXT_AreaProps, TXT_SelectArea, TXT_PosRerollOld, TXT_LoopLimit
global TXT_SetTargetWin, TXT_BtnGroup
global TXT_BtnOpenChests, TXT_BtnSalvage, TXT_BtnAtanor, TXT_BtnSalvageRed
global TXT_BtnReroll, TXT_BtnInfo, TXT_BtnStop, TXT_BtnRestart

; Messages & Tooltips Variables
global TXT_Msg_TargetWinNotSet, TXT_Msg_CoordNotSet, TXT_Msg_AreaNotSet
global TXT_Msg_BtnNotSet, TXT_Msg_DeleteConfirm, TXT_Msg_PresetDeleted
global TXT_SB_EnterName, TXT_SB_Saved, TXT_SB_Deleted, TXT_SB_SelectDelete, TXT_SB_HotkeysSaved
global TXT_Tip_Paused, TXT_Tip_Opening, TXT_Tip_Salvaging, TXT_Tip_SalvagingRed
global TXT_Tip_Upgrading, TXT_Tip_ClickSelect, TXT_Tip_DoubleClick, TXT_Tip_HoldLMB
global TXT_Tip_Selection, TXT_Tip_Selection2, TXT_Tip_AreaSmall, TXT_Tip_AreaSet, TXT_Tip_RerollStart
global TXT_Tip_Recognized, TXT_Tip_Found, TXT_Tip_Rerolling

; Info Window Variables
global TXT_Info_Author, TXT_Info_Release, TXT_Info_Update, TXT_Info_Desc
global TXT_Info_Link, TXT_Info_Hotkeys
global TXT_HK_Stop, TXT_HK_Pause, TXT_HK_Restart, TXT_HK_OpenChests
global TXT_HK_Salvage, TXT_HK_SalvageRed, TXT_HK_Reroll, TXT_HK_Atanor, TXT_HK_Close

; ============================================================================
; GLOBAL VARIABLES - HOTKEYS
; ============================================================================

global Stop_key := "Numpad1"
global Pause_key := "Numpad2"
global Reset_key := "Numpad3"
global OpenChests_key := "Numpad4"
global Salvage_items_key := "Numpad5"
global Salvage_red_items_key := "Numpad6"
global Reroll_Properties_key := "Numpad7"
global Atanor_key := "Numpad8"
global Close_script_key := "Numpad9"

; ============================================================================
; APPLICATION PATHS
; ============================================================================

RegExMatch(A_ScriptName, "^(.*?)\.", match)
global basename1 := match
global WINTITLE := basename1 . " v" . VERSION

global presetsDir := A_ScriptDir
global presetFile := presetsDir "\presets.ini"
global LangDir := A_ScriptDir "\Languages"

; ============================================================================
; LIBRARY INCLUDES
; ============================================================================

#include <Vis2>

; ============================================================================
; INITIALIZATION FLOW
; ============================================================================

LoadSettings()
LoadStrings()
BuildMenu()
GoSub, CreateGUI
return

; ============================================================================
; LOCALIZATION FUNCTIONS
; ============================================================================

LoadSettings() {
    global presetFile, CurrentLanguage, OCR_Language, LangCodes

    IniRead, lang, %presetFile%, Settings, Language, English
    CurrentLanguage := lang

    ; Validate Language Folder
    if (CurrentLanguage != "English") {
        targetDir := LangDir "\" CurrentLanguage
        if (!FileExist(targetDir)) {
            Gui, +OwnDialogs
            MsgBox, 262192, Error, Language folder for '%CurrentLanguage%' not found!`nPath: %targetDir%`n`nReverting to English.
            CurrentLanguage := "English"
            IniWrite, English, %presetFile%, Settings, Language
        }
    }

    if (LangCodes.HasKey(CurrentLanguage))
        OCR_Language := LangCodes[CurrentLanguage]
    else
        OCR_Language := "eng"
}

LoadStrings() {
    global

    ; --- ENGLISH DEFAULTS (Fallback) ---
    TXT_LangLabel := "Language:"
    TXT_PresetList := "Preset:"
    TXT_Save := "Save"
    TXT_Delete := "Delete"
    TXT_SettingsTitle := "Settings"
    TXT_PosAtanor := "#1 Athanor upgrade button position:"
    TXT_Set := "Set"
    TXT_PosSalvage := "#2 Item salvage button positions:"
    TXT_SalvageBtn := "Salvage Button"
    TXT_White := "White"
    TXT_Blue := "Blue"
    TXT_Green := "Green"
    TXT_Yellow := "Yellow"
    TXT_AreaProps := "#3 Item properties area:"
    TXT_SelectArea := "Select area"
    TXT_PosRerollOld := "#4 Item slot position after reroll (FOR OLD UI):"
    TXT_LoopLimit := "Loop Limiter:"
    TXT_SetTargetWin := "Set Window:"
    TXT_BtnGroup := "Buttons"
    TXT_BtnOpenChests := "OPEN CHESTS"
    TXT_BtnSalvage := "SALVAGE ITEMS"
    TXT_BtnAtanor := "UPGRADE ATHANOR"
    TXT_BtnSalvageRed := "SALVAGE RED ITEMS"
    TXT_BtnReroll := "REROLL PROPERTIES"
    TXT_BtnInfo := "<><><> HOTKEYS AND INFO <><><>"
    TXT_BtnStop := "STOP"
    TXT_BtnRestart := "RESTART"

    ; Messages & Tooltips
    TXT_Msg_TargetWinNotSet := "Target window not set. Set the window before running the script."
    TXT_Msg_CoordNotSet := "Coordinate %coord% not set. Set required coordinates before running."
    TXT_Msg_AreaNotSet := "Item properties area not set!`nSelect area before running."
    TXT_Msg_BtnNotSet := "Re-roll/Salvage button position not set!`nSet coordinates before running."
    TXT_Msg_DeleteConfirm := "Are you sure you want to delete preset"
    TXT_Msg_PresetDeleted := "Preset deleted:"

    TXT_SB_EnterName := "Enter preset name!"
    TXT_SB_Saved := "saved!"
    TXT_SB_Deleted := "deleted"
    TXT_SB_SelectDelete := "Select preset to delete"
    TXT_SB_HotkeysSaved := "Hotkeys saved to preset"

    TXT_Tip_Paused := "PAUSE"
    TXT_Tip_Opening := "Opening chests"
    TXT_Tip_Salvaging := "Salvaging items"
    TXT_Tip_SalvagingRed := "Salvaging red items"
    TXT_Tip_Upgrading := "Upgrading athanor"

    TXT_Tip_ClickSelect := "Click LMB to select`nPress ESC to cancel`n`nCurrent coordinates:"
    TXT_Tip_DoubleClick := "Double-click LMB on window`nPress ESC to cancel`n`nCurrent Window:"
    TXT_Tip_HoldLMB := "Hold LMB and select item properties area`nPress ESC to cancel`n`nCurrent coordinates:"
    TXT_Tip_Selection := "Selection:"
    TXT_Tip_Selection2 := "Release LMB to finish`nPress ESC to cancel"
    TXT_Tip_AreaSmall := "The area is too small! Please set a larger area"
    TXT_Tip_AreaSet := "Item properties area set!"

    TXT_Tip_RerollStart := "Starting property reroll...`nLooking for:"
    TXT_Tip_Recognized := "Recognized text:"
    TXT_Tip_Found := "Target properties found!"
    TXT_Tip_Rerolling := "Rerolling properties"

    ; Info Window
    TXT_Info_Author := "Author:"
    TXT_Info_Release := "Release date:"
    TXT_Info_Update := "Update date:"
    TXT_Info_Desc := "AutoHelper - is a script that automates a number`nof routine processes in Vermintide 2."
    TXT_Info_Link := "More information about it can be found in this"
    TXT_Info_Hotkeys := "HOTKEYS"

    TXT_HK_Stop := "Stop:"
    TXT_HK_Pause := "Pause:"
    TXT_HK_Restart := "Restart:"
    TXT_HK_OpenChests := "Open Chests:"
    TXT_HK_Salvage := "Salvage Items:"
    TXT_HK_SalvageRed := "Salvage Red Items:"
    TXT_HK_Reroll := "Reroll Properties:"
    TXT_HK_Atanor := "Upgrade Athanor:"
    TXT_HK_Close := "Close Script:"

    ; --- LOAD EXTERNAL LANGUAGE ---
    if (CurrentLanguage != "English") {
        ahkPath := LangDir "\" CurrentLanguage "\text.ahk"
        if (FileExist(ahkPath)) {
            Loop, Read, %ahkPath%
            {
                line := Trim(A_LoopReadLine)
                ; Simple parser: Variable := "Value"
                if (RegExMatch(line, "^(\w+)\s*:=\s*""(.*)""", match)) {
                    ; Need to handle `n in the string (escaped in file as `n or real newline? usually AHK file has real `n logic if included, but we are parsing)
                    ; We will manually replace literal `n string with newline char for parsing simple text
                    val := StrReplace(match2, "``n", "`n")
                    %match1% := val
                }
            }
        } else {
            ; ERROR: File missing
            Gui, +OwnDialogs
            MsgBox, 262192, Error, Localization file 'text.ahk' not found for language '%CurrentLanguage%'!`nPath: %ahkPath%`n`nReverting to English defaults.
            CurrentLanguage := "English"
            ; Update INI to reflect fallback
            IniWrite, English, %presetFile%, Settings, Language
        }
    }
}

BuildMenu() {
    global CurrentLanguage, LangDir

    try {
        Menu, RerollScript, DeleteAll
        Menu, Submenu1, DeleteAll
        Menu, Submenu2, DeleteAll
        Menu, Submenu3, DeleteAll
        Menu, Submenu4, DeleteAll
        Menu, Submenu5, DeleteAll
    }

    if (CurrentLanguage == "English") {
        ; Submenu 1: Melee
        Menu, Submenu1, Add, Attack Speed - Crit Chance, menuHandler
        Menu, Submenu1, Add, Attack Speed - Block Cost Reduction, menuHandler
        Menu, Submenu1, Add, Attack Speed - Crit Power, menuHandler
        Menu, Submenu1, Add, Attack Speed - Stamina , menuHandler
        Menu, Submenu1, Add, Attack Speed - vs Chaos, menuHandler
        Menu, Submenu1, Add, Attack Speed - vs Skaven, menuHandler
        Menu, Submenu1, Add, Block Cost Reduction - Crit Chance, menuHandler
        Menu, Submenu1, Add, Block Cost Reduction - Crit Power, menuHandler
        Menu, Submenu1, Add, Block Cost Reduction - Push/Block Angle, menuHandler
        Menu, Submenu1, Add, Block Cost Reduction - vs Chaos, menuHandler
        Menu, Submenu1, Add, Block Cost Reduction - vs Skaven, menuHandler
        Menu, Submenu1, Add, Block Cost Reduction - vs Attack Speed, menuHandler
        Menu, Submenu1, Add, Crit Chance - Crit Power, menuHandler
        Menu, Submenu1, Add, Crit Chance - vs Chaos, menuHandler
        Menu, Submenu1, Add, Crit Chance - vs Skaven, menuHandler
        Menu, Submenu1, Add, Crit Power - vs Chaos, menuHandler
        Menu, Submenu1, Add, Crit Power - vs Skaven, menuHandler
        Menu, Submenu1, Add, Stamina - Block Cost Reduction, menuHandler
        Menu, Submenu1, Add, Stamina - Crit Chance, menuHandler
        Menu, Submenu1, Add, Stamina - Crit Power, menuHandler
        Menu, Submenu1, Add, Stamina - Push/Block Angle, menuHandler
        Menu, Submenu1, Add, Stamina - vs Chaos, menuHandler
        Menu, Submenu1, Add, Stamina - vs Skaven, menuHandler

        ; Submenu 2: Ranged
        Menu, Submenu2, Add, Crit chance - Crit Power, menuHandler
        Menu, Submenu2, Add, Crit chance - vs Armoured, menuHandler
        Menu, Submenu2, Add, Crit chance - vs Berserkers, menuHandler
        Menu, Submenu2, Add, Crit chance - vs Chaos, menuHandler
        Menu, Submenu2, Add, Crit chance - vs Infantry, menuHandler
        Menu, Submenu2, Add, Crit chance - vs Monsters, menuHandler
        Menu, Submenu2, Add, Crit chance - vs Skaven, menuHandler
        Menu, Submenu2, Add, Crit Power - vs Armoured, menuHandler
        Menu, Submenu2, Add, Crit Power - vs Berserkers, menuHandler
        Menu, Submenu2, Add, Crit Power - vs Chaos, menuHandler
        Menu, Submenu2, Add, Crit Power - vs Infantry, menuHandler
        Menu, Submenu2, Add, Crit Power - vs Monsters, menuHandler
        Menu, Submenu2, Add, Crit Power - vs Skaven, menuHandler
        Menu, Submenu2, Add, vs Armoured - vs Berserkers, menuHandler
        Menu, Submenu2, Add, vs Armoured - vs Infantry, menuHandler
        Menu, Submenu2, Add, vs Armoured - vs Monsters, menuHandler
        Menu, Submenu2, Add, vs Chaos - vs Armoured, menuHandler
        Menu, Submenu2, Add, vs Chaos - vs Berserkers, menuHandler
        Menu, Submenu2, Add, vs Chaos - vs Infantry, menuHandler
        Menu, Submenu2, Add, vs Chaos - vs Monsters, menuHandler
        Menu, Submenu2, Add, vs Infantry - vs Berserkers, menuHandler
        Menu, Submenu2, Add, vs Monsters - vs Infantry, menuHandler
        Menu, Submenu2, Add, vs Monsters - vs Infantry, menuHandler
        Menu, Submenu2, Add, vs Skaven - vs Armoured, menuHandler
        Menu, Submenu2, Add, vs Skaven - vs Berserkers, menuHandler
        Menu, Submenu2, Add, vs Skaven - vs Infantry, menuHandler
        Menu, Submenu2, Add, vs Skaven - vs Monsters, menuHandler

        ; Submenu 3: Necklace
        Menu, Submenu3, Add, Block Cost Reduction - Health, menuHandler
        Menu, Submenu3, Add, Block Cost Reduction - Push/Block Angle, menuHandler
        Menu, Submenu3, Add, Stamina - Block Cost Reduction, menuHandler
        Menu, Submenu3, Add, Stamina - Health, menuHandler
        Menu, Submenu3, Add, Stamina - Push/Block Angle Angle, menuHandler

        ; Submenu 4: Charm
        Menu, Submenu4, Add, Attack Speed - Crit Power, menuHandler
        Menu, Submenu4, Add, Attack Speed - vs Armoured, menuHandler
        Menu, Submenu4, Add, Attack Speed - vs Berserkers, menuHandler
        Menu, Submenu4, Add, Attack Speed - vs Chaos, menuHandler
        Menu, Submenu4, Add, Attack Speed - vs Infantry, menuHandler
        Menu, Submenu4, Add, Attack Speed - vs Monsters, menuHandler
        Menu, Submenu4, Add, Attack Speed - vs Skaven, menuHandler
        Menu, Submenu4, Add, Crit Power - vs Armoured, menuHandler
        Menu, Submenu4, Add, Crit Power - vs Berserkers, menuHandler
        Menu, Submenu4, Add, Crit Power - vs Chaos, menuHandler
        Menu, Submenu4, Add, Crit Power - vs Infantry, menuHandler
        Menu, Submenu4, Add, Crit Power - vs Monsters, menuHandler
        Menu, Submenu4, Add, Crit Power - vs Skaven, menuHandler
        Menu, Submenu4, Add, vs Armoured - vs Berserkers, menuHandler
        Menu, Submenu4, Add, vs Armoured - vs Infantry, menuHandler
        Menu, Submenu4, Add, vs Armoured - vs Monsters, menuHandler
        Menu, Submenu4, Add, vs Chaos - vs Armoured, menuHandler
        Menu, Submenu4, Add, vs Chaos - vs Berserkers, menuHandler
        Menu, Submenu4, Add, vs Chaos - vs Infantry, menuHandler
        Menu, Submenu4, Add, vs Chaos - vs Monsters, menuHandler
        Menu, Submenu4, Add, vs Chaos - vs Skaven, menuHandler
        Menu, Submenu4, Add, vs Infantry - vs Berserkers, menuHandler
        Menu, Submenu4, Add, vs Monsters - vs Infantry, menuHandler
        Menu, Submenu4, Add, vs Skaven - vs Armoured, menuHandler
        Menu, Submenu4, Add, vs Skaven - vs Berserkers, menuHandler
        Menu, Submenu4, Add, vs Skaven - vs Infantry, menuHandler
        Menu, Submenu4, Add, vs Skaven - vs Monsters, menuHandler

        ; Submenu 5: Trinket
        Menu, Submenu5, Add, Cooldown - Crit Chance, menuHandler
        Menu, Submenu5, Add, Cooldown - Curse Resist, menuHandler
        Menu, Submenu5, Add, Cooldown - Movement speed, menuHandler
        Menu, Submenu5, Add, Cooldown - Revive Speed, menuHandler
        Menu, Submenu5, Add, Cooldown - Stamina recovery, menuHandler
        Menu, Submenu5, Add, Crit Chance - Curse Resist, menuHandler
        Menu, Submenu5, Add, Crit Chance - Movement Speed, menuHandler
        Menu, Submenu5, Add, Crit Chance - Revive Speed, menuHandler
        Menu, Submenu5, Add, Crit Chance - Stamina recovery, menuHandler
        Menu, Submenu5, Add, Curse Resist - Movement Speed, menuHandler
        Menu, Submenu5, Add, Curse Resist - Stamina recovery, menuHandler
        Menu, Submenu5, Add, Movement Speed - Revive Speed, menuHandler
        Menu, Submenu5, Add, Movement Speed - Stamina recovery, menuHandler
        Menu, Submenu5, Add, Revive Speed - Stamina recovery, menuHandler

        Menu, RerollScript, Add, Melee, :Submenu1
        Menu, RerollScript, Add, Ranged, :Submenu2
        Menu, RerollScript, Add ; Separator
        Menu, RerollScript, Add, Necklace, :Submenu3
        Menu, RerollScript, Add, Charm, :Submenu4
        Menu, RerollScript, Add, Trinkets, :Submenu5

    } else {
        menuPath := LangDir "\" CurrentLanguage "\menu.ahk"
        if (FileExist(menuPath)) {
            Loop, Read, %menuPath%
            {
                tline := Trim(A_LoopReadLine)
                if (RegExMatch(tline, "^Menu,\s*([^,]+),\s*Add\s*$", match)) {
                    Menu, %match1%, Add
                }
                else if (RegExMatch(tline, "^Menu,\s*([^,]+),\s*Add,\s*([^,]+),\s*([^,]+)", match)) {
                    Menu, %match1%, Add, %match2%, %match3%
                }
            }
        } else {
            Menu, RerollScript, Add, Error: menu.ahk missing, menuHandler
        }
    }
}

ChangeLanguage:
    Gui, Submit, NoHide
    if (NewLanguage != CurrentLanguage) {
        IniWrite, %NewLanguage%, %presetFile%, Settings, Language
        Reload
    }
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

    ; Language Selector
    Gui, Font, s9
    Gui, Add, Text, x10 y10, %TXT_LangLabel%
    Gui, Add, DropDownList, x+10 yp-3 w100 vNewLanguage gChangeLanguage Choose%CurrentLanguage%, %SupportedLanguages%
    Gui, Font

    Gui, Font, s9, Segoe UI
    Gui, Add, Text, x180 y10, %TXT_LoopLimit%
    Gui, Add, Edit, x+10 w65 h20 vLoopCount Number

    ; Preset Panel
    Gui, Font, s9
    Gui, Add, Text, x10 y40 section, %TXT_PresetList%
    Gui, Add, ComboBox, x+8 vsavedPreset gPresetChange
    Gui, Add, Button, x+10 h21 w70 gSavePreset, %TXT_Save%
    Gui, Add, Button, x+10 h21 w70 gDeletePreset vDELETEBUTTON, %TXT_Delete%
    Gui, Add, StatusBar

    ; Configuration Section
    Gui, Font, c535353 s13 Bold Q3, Segoe UI Black
    Gui, Add, GroupBox, x10 y70 w380 h330 Center Section, %TXT_SettingsTitle%
    Gui, Font
    Gui, Font, s11, Segoe UI

    ; Atanor Position
    Gui, Add, Text, xs+15 yp+30, %TXT_PosAtanor%
    Gui, Font, s9
    Gui, Add, Button, xs+15 yp+30 w80 h25 gSetPosAtanor, %TXT_Set%
    Gui, Add, Edit, x+15 yp+3 w80 h20 vPos_5 ReadOnly,
    Gui, Font, s11, Segoe UI

    ; Salvage Buttons
    Gui, Add, Text, xs+15 yp+30, %TXT_PosSalvage%
    Gui, Font, s9
    Gui,Add,Button,xs+15 yp+25 w140 h25 gSet_Pos_Button, %TXT_SalvageBtn%
    Gui,Add,Edit,x+10 yp+3 w195 h20 vPos_6 ReadOnly,
    Gui, Add, Button, xs+15 yp+30 w70 h20 gSetPosWhite, %TXT_White%
    Gui, Add, Edit, x+15 w80 h20 vPos_1 ReadOnly,
    Gui, Add, Button, x+15 w70 h20 gSetPosBlue, %TXT_Blue%
    Gui, Add, Edit, x+15 w80 h20 vPos_3 ReadOnly ,
    Gui, Add, Button, xs+15 yp+25 w70 h20 gSetPosGreen, %TXT_Green%
    Gui, Add, Edit, x+15 w80 h20 vPos_2 ReadOnly,
    Gui, Add, Button, x+15 w70 h20 gSetPosYellow, %TXT_Yellow%
    Gui, Add, Edit, x+15 w80 h20 vPos_4 ReadOnly,

    ; OCR Area
    Gui, Font, s11, Segoe UI
    Gui, Add, Text, xs+15 yp+30, %TXT_AreaProps%
    Gui, Font, s9
    Gui, Add, Button, xs+15 yp+25 w140 h25 gSetOCRArea, %TXT_SelectArea%
    Gui, Add, Edit, x+15 yp+3 w140 h20 vOCR_Pos ReadOnly,

    ; AfterReroll
    Gui, Font, s10, Segoe UI
    Gui, Add, Text, xs+15 yp+30, %TXT_PosRerollOld%
    Gui, Font, s9
    Gui, Add, Button, xs+15 yp+25 w80 h25 gSet_Pos_AfterReroll, %TXT_Set%
    Gui, Add, Edit, x+15 yp+3 w80 h20 vPos_7 ReadOnly,

    ; Additional Settings
    ;Gui, Font, s10, Segoe UI
    ;Gui, Add, Text, x20 y630, %TXT_LoopLimit%
    ;Gui, Add, Edit, x+10 w75 h20 vLoopCount Number

    Gui, Font, s9 Bold, Segoe UI
    Gui, Add, Button, x20 y645 w130 h30 gSetLocation, %TXT_SetTargetWin%
    Gui, Add, Edit, x+15 yp+3 w210 h25 vTarget_Window ReadOnly, %Target_Window%
    Gui, Font

    ; Action Buttons Section
    Gui, Font, c535353 s13 Bold Q3, Segoe UI Black
    Gui, Add, GroupBox, x10 y410 w380 h210 Center Section, %TXT_BtnGroup%
    Gui, Font, s9 Bold, Segoe UI

    Gui, Add, Button, xs+15 yp+30 w110 h35 gOpenChests, %TXT_BtnOpenChests%
    Gui, Add, Button, x+10 w110 h35 gSalvageItems, %TXT_BtnSalvage%
    Gui, Add, Button, x+10 w110 h35 gUpgradeAtanor, %TXT_BtnAtanor%
    Gui, Add, Button, xs+15 yp+45 w170 h35 gSalvageRedItems, %TXT_BtnSalvageRed%
    Gui, Add, Button, x+10 w170 h35 gRerollProperties, %TXT_BtnReroll%
    Gui, Add, Button, xs+40 yp+45 w300 h25 gShowInfo, %TXT_BtnInfo%
    Gui, Add, Button, xs+15 yp+35 w170 h35 gStope, %TXT_BtnStop%
    Gui, Add, Button, x+10 w170 h35 gReset, %TXT_BtnRestart%
    Gui, Font

    Gui, Show, w400 h720, %WINTITLE%
    GoSub, UpdatePresetList
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

    Gui, Add, Text, xs+145 yp+25, %TXT_Info_Author%
    Gui, Add, Link, x+5, <a href="https://steamcommunity.com/id/ChadMasodin">ChadMasodin</a>

    Gui, Add, Text, xs+145 yp+20, %TXT_Info_Release% 27/12/24
    Gui, Add, Text, xs+145 yp+15, %TXT_Info_Update% 27/12/25

    Gui, Add, Text, xs+10 yp+20 w260, %TXT_Info_Desc%
    Gui, Add, Link, xs+10 yp+30, %TXT_Info_Link% <a href="https://steamcommunity.com/sharedfiles/filedetails/?id=3385048068\">guide.</a>

    ; Hotkeys Section
    Gui, Font, c666666 s11 Bold Q3, Segoe UI Black
    Gui, Add, GroupBox, x10 y170 w280 h300 Center Section, %TXT_Info_Hotkeys%
    Gui, Font, s9

    labelX := "xs+10"
    hotkeyX := "x+1"
    hotkeyWidth := "w90"
    labelWidth := "w155"

    Gui, Add, Text, cBlue %labelX% yp+30 %labelWidth% Left, %TXT_HK_Stop%
    Gui, Add, Hotkey, %hotkeyX% yp %hotkeyWidth% h20 vStop_key, %Stop_key%

    Gui, Add, Text, cBlue %labelX% yp+30 %labelWidth% Left, %TXT_HK_Pause%
    Gui, Add, Hotkey, %hotkeyX% yp %hotkeyWidth% h22 vPause_key, %Pause_key%

    Gui, Add, Text, cBlue %labelX% yp+30 %labelWidth% Left, %TXT_HK_Restart%
    Gui, Add, Hotkey, %hotkeyX% yp %hotkeyWidth% h22 vReset_key, %Reset_key%

    Gui, Add, Text, cRed %labelX% yp+30 %labelWidth% Left, %TXT_HK_OpenChests%
    Gui, Add, Hotkey, %hotkeyX% yp %hotkeyWidth% h20 vOpenChests_key, %OpenChests_key%

    Gui, Add, Text, cRed %labelX% yp+30 %labelWidth% Left, %TXT_HK_Salvage%
    Gui, Add, Hotkey, %hotkeyX% yp %hotkeyWidth% h20 vSalvage_items_key, %Salvage_items_key%

    Gui, Add, Text, cRed %labelX% yp+30 %labelWidth% Left, %TXT_HK_SalvageRed%
    Gui, Add, Hotkey, %hotkeyX% yp %hotkeyWidth% h20 vSalvage_red_items_key, %Salvage_red_items_key%

    Gui, Add, Text, cRed %labelX% yp+30 %labelWidth% Left, %TXT_HK_Reroll%
    Gui, Add, Hotkey, %hotkeyX% yp %hotkeyWidth% h20 vReroll_Properties_key, %Reroll_Properties_key%

    Gui, Add, Text, cRed %labelX% yp+30 %labelWidth% Left, %TXT_HK_Atanor%
    Gui, Add, Hotkey, %hotkeyX% yp %hotkeyWidth% h20 vAtanor_key, %Atanor_key%

    Gui, Add, Text, cMaroon %labelX% yp+30 %labelWidth% Left, %TXT_HK_Close%
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

            SB_SetText(TXT_SB_HotkeysSaved . " '" . savedPreset . "'")
        }
    }

    Gui, Destroy
return

; ============================================================================
; VALIDATION FUNCTIONS
; ============================================================================

ValidateScript(requireCoords := false, coordsList := "") {
    global LoopCount, Target_Window, TXT_Msg_TargetWinNotSet, TXT_Msg_CoordNotSet
    Gui, +OwnDialogs

    if (Target_Window == "") {
        MsgBox, 262192, Error, %TXT_Msg_TargetWinNotSet%
        return false
    }

    if (requireCoords) {
        Loop, Parse, coordsList, `,
        {
            coord := Trim(A_LoopField)
            if (%coord% == "") {
                errText := StrReplace(TXT_Msg_CoordNotSet, "%coord%", coord)
                MsgBox, 262160, Error, %errText%
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

Set_Pos_Button:
    Stop := 1
    GetClickPos(X6, Y6)
    GuiControl,, Pos_6, %X6% %Y6%
return

Set_Pos_AfterReroll:
    Stop := 1
    GetClickPos(X7, Y7)
    GuiControl,, Pos_7, %X7% %Y7%
return

SetOCRArea:
    Stop := 1
    GetOCRArea(OCR_X, OCR_Y, OCR_W, OCR_H)
    GuiControl,, OCR_Pos, %OCR_X% %OCR_Y% %OCR_W% %OCR_H%
return

SetLocation:
    Target_Window := SetWindow(Target_Window)
    GuiControl,, Target_Window, %Target_Window%
return

GetClickPos(ByRef X, ByRef Y) {
    global TXT_Tip_ClickSelect
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
        ToolTip, %TXT_Tip_ClickSelect%`nX=%currentX% Y=%currentY%

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
    global TXT_Tip_DoubleClick
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
        ToolTip, %TXT_Tip_DoubleClick% %Temp_Window%

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
    global TXT_Tip_HoldLMB, TXT_Tip_Selection, TXT_Tip_AreaSmall, TXT_Tip_AreaSet, TXT_Tip_Selection2

    ; First loop - wait for LMB press
    Loop {
        ; Check ESC press
        if (GetKeyState("Esc", "P")) {
            ToolTip
            DestroySelectionGui()
            X := Y := W := H := ""
            return
        }

        ; Check LMB press to exit first loop
        if (GetKeyState("LButton", "P")) {
            break
        }
        MouseGetPos, currentX, currentY
        ToolTip, %TXT_Tip_HoldLMB%`nX1=%currentX% Y1=%currentY%
    }

    ; Get start coordinates after LMB press
    MouseGetPos, startX, startY

    ; Second loop - track area selection
    Loop {
        ; Check ESC press
        if (GetKeyState("Esc", "P")) {
            DestroySelectionGui()
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

        ; Update selection rectangle visualization
        DrawSelectionRect(minX, minY, width, height)

        ; Show selection info with moving ToolTip
        ToolTip, %TXT_Tip_Selection% %minX% %minY% [%width% x %height%]`n`n%TXT_Tip_Selection2%

        Sleep, 10
    }

    ; Get end coordinates
    MouseGetPos, endX, endY

    ; Calculate final area
    X := Min(startX, endX)
    Y := Min(startY, endY)
    W := Abs(startX - endX)
    H := Abs(startY - endY)

        ; Destroy visualization
    DestroySelectionGui()

    if (W < 15 || H < 15) {
        ToolTip, %TXT_Tip_AreaSmall%
        Sleep, 3000
        ToolTip
        X := Y := W := H := ""
        return
    }

    ToolTip, %TXT_Tip_AreaSet%`nX: %X% Y: %Y%`nW: %W% H: %H%
    Sleep, DELAY_ANIMATION
    ToolTip
}

DrawSelectionRect(X, Y, W, H) {
    global hSelectionGui

    ; Create or update selection window
    if (hSelectionGui = "") {
        Gui, SelectionOverlay: -Caption +AlwaysOnTop +ToolWindow +E0x20
        Gui, SelectionOverlay: Color, 0066FF
        Gui, SelectionOverlay: +LastFound
        SelectionOverlayGui := WinExist()
        WinSet, Transparent, 135, ahk_id %SelectionOverlayGui%
        hSelectionGui := 1
    }

    ; Show the rectangle with transparency
    Gui, SelectionOverlay: Show, x%X% y%Y% w%W% h%H% NoActivate, SelectionOverlay
}

DestroySelectionGui() {
    global hSelectionGui
    Gui, SelectionOverlay: Destroy
    hSelectionGui := ""
}

; ============================================================================
; CONTROL FUNCTIONS
; ============================================================================

Stope:
    Stop := 1
    ToolTip
return

Pauza:
    global TXT_Tip_Paused
    if (A_IsPaused) {
        Pause, , 1
        ToolTip
    } else {
        Pause, , 1
        ToolTip, %TXT_Tip_Paused%
        Sleep, DELAY_SHORT
    }
return

Reset:
    Reload
return

CloseScript:
    ExitApp
return

; ============================================================================
; MAIN AUTOMATION FUNCTIONS
; ============================================================================

OpenChests:
    global TXT_Tip_Opening
    Stop := 0
    CycleCount := 0
    Gui, Submit, NoHide

    if (!ValidateScript()) {
        return
    }

    loop {
        if (!CheckActiveWindow())
            continue

        if (Stop == 1 || (LoopCount != "" && CycleCount >= LoopCount))
            break

        SendInput, {Space}
        Sleep, DELAY_LONG
        SendInput, {Space}

        CycleCount++
        UpdateTooltip(TXT_Tip_Opening, CycleCount)
        Sleep, DELAY_LONG
    }

    if (LoopCount != "" && CycleCount >= LoopCount) {
        SoundPlay, *64
    }

    ToolTip
return

SalvageItems:
    global TXT_Tip_Salvaging
    Stop := 0
    CycleCount := 0
    Gui, Submit, NoHide

    if (!ValidateScript(true, "X1,Y1,X2,Y2,X3,Y3,X4,Y4,X6,Y6")) {
        return
    }

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

        MouseMove, %X6%, %Y6%
        Click, down
        Sleep, DELAY_LONG
        Click, up
        Sleep, DELAY_ANIMATION

        CycleCount++
        UpdateTooltip(TXT_Tip_Salvaging, CycleCount)
        Sleep, DELAY_MEDIUM
    }

    if (LoopCount != "" && CycleCount >= LoopCount) {
        SoundPlay, *64
    }

    ToolTip
return

SalvageRedItems:
    global TXT_Tip_SalvagingRed
    Stop := 0
    CycleCount := 0
    Gui, Submit, NoHide

    if (!ValidateScript(true, "X1,Y1,X2,Y2,X3,Y3,X4,Y4,X6,Y6")) {
        return
    }

    loop {
        if (!CheckActiveWindow())
            continue

        if (Stop == 1 || (LoopCount != "" && CycleCount >= LoopCount))
            break

        Sleep, DELAY_LONG

        MouseMove, %X1%, %Y1%
        Click right
        Sleep, DELAY_MEDIUM

        MouseMove, %X2%, %Y2%
        Click right
        Sleep, DELAY_MEDIUM

        MouseMove, %X3%, %Y3%
        Click right
        Sleep, DELAY_MEDIUM

        MouseMove, %X4%, %Y4%
        Click right
        Sleep, DELAY_MEDIUM

        MouseMove, %X6%, %Y6%
        Click, down
        Sleep, DELAY_LONG
        Click, up
        Sleep, DELAY_ANIMATION

        CycleCount++
        UpdateTooltip(TXT_Tip_SalvagingRed, CycleCount)
        Sleep, DELAY_SHORT
    }

    if (LoopCount != "" && CycleCount >= LoopCount) {
        SoundPlay, *64
    }

    ToolTip
return

UpgradeAtanor:
    global TXT_Tip_Upgrading
    Stop := 0
    CycleCount := 0
    Gui, Submit, NoHide

    if (!ValidateScript(true, "X5,Y5")) {
        return
    }

    loop {
        if (!CheckActiveWindow())
            continue

        if (Stop == 1 || (LoopCount != "" && CycleCount >= LoopCount))
            break

        Sleep, 150
        MouseMove, %X5%, %Y5%
        Click

        CycleCount++
        UpdateTooltip(TXT_Tip_Upgrading, CycleCount)
        Sleep, DELAY_ANIMATION
    }

    if (LoopCount != "" && CycleCount >= LoopCount) {
        SoundPlay, *64
    }

    ToolTip
return

; ============================================================================
; REROLL PROPERTIES FUNCTION
; ============================================================================

RerollProperties:
    Gui, Submit, NoHide
    Gui, +OwnDialogs
    global TXT_Msg_TargetWinNotSet, TXT_Msg_AreaNotSet, TXT_Msg_BtnNotSet

    if (Target_Window == "") {
        MsgBox, 262192, Error, %TXT_Msg_TargetWinNotSet%
        return
    }

    if (OCR_X = "" || OCR_Y = "" || OCR_W = "" || OCR_H = "") {
        MsgBox, 262160, Error, %TXT_Msg_AreaNotSet%
        return
    }

    if ((!X6 || !Y6)) {
        MsgBox, 262160, Error, %TXT_Msg_BtnNotSet%
        return
    }
    Menu, RerollScript, Show
return

menuHandler(itemName) {
    global Stop, Target_Window, LoopCount, OCR_X, OCR_Y, OCR_W, OCR_H, OCR_Language
    global SelectedProperties := itemName
    global X7, Y7
    global TXT_Tip_RerollStart, TXT_Tip_Recognized, TXT_Tip_Found, TXT_Tip_Rerolling

    ; Split property names from menu selection
    Properties := StrSplit(itemName, " - ")
    Property1 := Properties[1]
    Property2 := Properties[2]

    Gui, Submit, NoHide
    WinActivate, %Target_Window%

    CycleCount := 0
    Stop := 0

    ; Setup coordinates for OCR
    ocrArea := [OCR_X, OCR_Y, OCR_W, OCR_H]

    ToolTip, %TXT_Tip_RerollStart% %Property1% - %Property2%
    Sleep, 1000

    loop {
        ; Check stop conditions
        if (Stop == 1 || (LoopCount != "" && CycleCount >= LoopCount)) {
            ToolTip
            break
        }

        ; Check if correct window is active
        if (!CheckActiveWindow())
            continue

        try {
            ; OCR with specified Language Code
            text := OCR(ocrArea, OCR_Language)
        }

        ; Display OCR results
        ToolTip, %TXT_Tip_Recognized%`n%text%`n---`nLooking for: %Property1% - %Property2%`nCycle: %CycleCount%/%LoopCount%
        Sleep, 1000

        ; Check if target properties found
        If (InStr(text, Property1) && InStr(text, Property2)) {
            SoundPlay, *64
            ToolTip, %TXT_Tip_Found%`n%Property1% - %Property2%
            Sleep, 5000
            ToolTip
            return
        }

        MouseMove, %X6%, %Y6%
        Click, down
        Sleep, DELAY_LONG
        Click, up
        Sleep, 2500

        ; MOVE MOUSE TO SPECIFIED COORDINATE AFTER REROLL
        if (X7 != "" && Y7 != "") {
            MouseMove, %X7%, %Y7%
            Sleep, 1000
        }

        CycleCount++
        UpdateTooltip(TXT_Tip_Rerolling, CycleCount)
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

    Loop, 7 {
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

return

SavePreset:
    Gui, Submit, NoHide
    Gui, +OwnDialogs

    if (savedPreset = "") {
        SB_SetText(TXT_SB_EnterName)
        return
    }

    Loop, 7 {
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

    SB_SetText(TXT_SB_HotkeysSaved . " '" . savedPreset . "'")
    GoSub, UpdatePresetList

return

DeletePreset:
    Gui, Submit, NoHide
    Gui, +OwnDialogs

    if (savedPreset = "") {
        SB_SetText(TXT_SB_SelectDelete)
        return
    }

    MsgBox, 262148, Confirmation, %TXT_Msg_DeleteConfirm% "%savedPreset%"?
    IfMsgBox, Yes
    {
        IniDelete, %presetFile%, %savedPreset%
        SB_SetText(TXT_Msg_PresetDeleted . " '" . savedPreset . "'")
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
    ExitApp
return

InfoWindowGuiClose:
    Gui, Destroy
return