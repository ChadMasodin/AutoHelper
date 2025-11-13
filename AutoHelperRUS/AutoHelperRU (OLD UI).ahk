/*
===============================================================================
AutoHelper - Vermintide 2 Automation Script
===============================================================================
Author: ChadMasodin
Version: 2.0
Type game UI: Old
Language: Russian
Date: 10/11/2025
Description: Automated script for routine processes in Vermintide 2
===============================================================================
*/

; ============================================================================
; INITIALIZATION AND SETTINGS
; ============================================================================

#NoEnv
#SingleInstance Force
SetBatchLines -1 ; have the script run at maximum speed and never sleep
ListLines Off ; a debugging option
SetMouseDelay, -1 ; Убирает задержки между действиями мыши
;CoordMode, Mouse, Client ; Определяет координаты мыши относительно клиентской области окна
SetTitleMatchMode, 2  ; Позволяет частичное совпадение заголовков окон при поиске
SetControlDelay, -1 ;Убирает задержки при работе с элементами GUI
DetectHiddenWindows, On  ; Позволяет обнаруживать скрытые окна
SetStoreCapsLockMode, Off  ; Отключает запоминание состояния клавиши CapsLock

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
global COLOR_TOLERANCE := 20

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

; Очистка лог-файла при превышении размера (например, 1 МБ)
ManageLogFile() {
    global logFile

    ; Проверяем существование файла
    if (!FileExist(logFile))
        return

    ; Получаем размер файла в байтах
    FileGetSize, fileSize, %logFile%

    ; Если размер превышает 500 КБ (512000 байт), очищаем файл
    if (fileSize > 512000) {
        FileDelete, %logFile%
        FileAppend, [Лог-файл был очищен из-за превышения размера]`n, %logFile%
        LogAction("Лог-файл очищен (превышен размер 500 КБ)")
    }
}

; Вызов функции управления лог-файлом при запуске
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
    Gui, Add, Text, yp+11 xm section, Список пресетов:
    Gui, Add, ComboBox, x+10 vsavedPreset gPresetChange
    Gui, Add, Button, x+10 h21 w70 gSavePreset, Сохранить
    Gui, Add, Button, x+10 h21 w70 gDeletePreset vDELETEBUTTON, Удалить
    Gui, Add, StatusBar

    ; Configuration Section
    Gui, Font, c535353 s13 Bold Q3, Segoe UI Black
    Gui, Add, GroupBox, x10 y40 w380 h390 Center Section, Настройка скриптов
    Gui, Font
    Gui, Font, s11, Segoe UI

    ; Atanor Position
    Gui, Add, Text, xs+15 yp+30, #1 Позиция кнопки для улучшения Атанора:
    Gui, Font, s9
    Gui, Add, Button, xs+15 yp+30 w80 h25 gSetPosAtanor, Установить
    Gui, Add, Edit, x+15 yp+3 w80 h20 vPos_5 ReadOnly,
    Gui, Font, s11, Segoe UI

    ; Salvage Buttons
    Gui, Add, Text, xs+15 yp+30, #2 Позиция кнопок для утилизации предметов:
    Gui, Font, s9
    Gui,Add,Button,xs+15 yp+25 w140 h25 gSet_Pos_Button, Кнопка утилизации
    Gui,Add,Edit,x+10 yp+3 w195 h20 vPos_6 ReadOnly,
    Gui, Add, Button, xs+15 yp+30 w70 h20 gSetPosWhite, Белая
    Gui, Add, Edit, x+15 w80 h20 vPos_1 ReadOnly,
    Gui, Add, Button, x+15 w70 h20 gSetPosBlue, Синяя
    Gui, Add, Edit, x+15 w80 h20 vPos_3 ReadOnly ,
    Gui, Add, Button, xs+15 yp+25 w70 h20 gSetPosGreen, Зеленая
    Gui, Add, Edit, x+15 w80 h20 vPos_2 ReadOnly,
    Gui, Add, Button, x+15 w70 h20 gSetPosYellow, Желтая
    Gui, Add, Edit, x+15 w80 h20 vPos_4 ReadOnly,

    ; OCR Area
    Gui, Font, s11, Segoe UI
    Gui, Add, Text, xs+15 yp+30, #3 Область свойств предмета:
    Gui, Font, s9
    Gui, Add, Button, xs+15 yp+25 w140 h25 gSetOCRArea, Выделить область
    Gui, Add, Edit, x+15 yp+3 w140 h20 vOCR_Pos ReadOnly,

    ; Green Dust Area
    Gui, Font, s11, Segoe UI
    Gui, Add, Text, xs+15 yp+30, #4 Область иконки зеленой пыли:
    Gui, Font, s9
    Gui, Add, Button, xs+15 yp+25 w140 h25 gSetGreenDustArea, Выделить область
    Gui, Add, Edit, x+15 yp+3 w140 h20 vGreenDust_Pos ReadOnly,

    ; AfterReroll
    Gui, Font, s11, Segoe UI
    Gui, Add, Text, xs+15 yp+30, #5 Позиция предмета в слоте переброса:
    Gui, Font, s9
    Gui, Add, Button, xs+15 yp+25 w80 h25 gSet_Pos_AfterReroll, Установить
    Gui, Add, Edit, x+15 yp+3 w80 h20 vPos_7 ReadOnly,

    ; Additional Settings
    Gui, Font, s10, Segoe UI
    Gui, Add, Text, x25 y650, Общий ограничитель циклов (опционально):
    Gui, Add, Edit, x+10 w75 h20 vLoopCount Number

    Gui, Font, s9 Bold, Segoe UI
    Gui, Add, Button, x25 y680 w130 h30 gSetLocation, Установите окно:
    Gui, Add, Edit, x+15 yp+3 w210 h25 vTarget_Window ReadOnly, %Target_Window%
    Gui, Font

    ; Action Buttons Section
    Gui, Font, c535353 s13 Bold Q3, Segoe UI Black
    Gui, Add, GroupBox, x10 y430 w380 h210 Center Section, Кнопки
    Gui, Font, s9 Bold, Segoe UI

    Gui, Add, Button, xs+15 yp+30 w110 h35 gOpenChests, ОТКРЫТИЕ`nСУНДУКОВ
    Gui, Add, Button, x+10 w110 h35 gSalvageItems, УТИЛИЗАЦИЯ`nПРЕДМЕТОВ
    Gui, Add, Button, x+10 w110 h35 gUpgradeAtanor, УЛУЧШЕНИЕ`nАТАНОРА
    Gui, Add, Button, xs+15 yp+45 w170 h35 gSalvageRedItems, УТИЛИЗАЦИЯ`nКРАСНЫХ ПРЕДМЕТОВ
    Gui, Add, Button, x+10 w170 h35 gRerollProperties, ПЕРЕБРОСИТЬ`nСВОЙСТВА
    Gui, Add, Button, xs+40 yp+45 w300 h25 gShowInfo, <><><> ИНФА И ПРИВЯЗКА КЛАВИШ <><><>
    Gui, Add, Button, xs+15 yp+35 w110 h35 gStope, СТОП
    Gui, Add, Button, x+10 w110 h35 gPauza, ПАУЗА
    Gui, Add, Button, x+10 w110 h35 gReset, ПЕРЕЗАПУСК
    Gui, Font

    Gui, Show, w400 h750, %WINTITLE%
    GoSub, UpdatePresetList
    #include <menu>
return

; ============================================================================
; INFO WINDOW
; ============================================================================

ShowInfo:
    Gui, InfoWindow:New, , Помощь
    Gui, +AlwaysOnTop -DPIScale +Owner1
    Gui, Add, Picture, x10 y33, icon.png

    Gui, Font, c666666 s11 Bold Q3, Segoe UI Black
    Gui, Add, GroupBox, x10 y10 w280 h150 Center Section, ИНФА
    Gui, Font

    Gui, Add, Link, xs+145 yp+25, Автор: <a href="https://steamcommunity.com/id/ChadMasodin">ChadMasodin</a>
    Gui, Add, Text, xs+145 yp+20, Дата выхода:  27/12/24
    Gui, Add, Text, xs+145 yp+15, Дата обновы:  10/11/25
    Gui, Add, Link, xs+10 yp+25, AutoHelper — это скрипт автоматизирующий ряд `nрутинных процессов в Vermintide 2.`nБолее подробную информацию о нем можно`nнайти в этом <a href="https://steamcommunity.com/sharedfiles/filedetails/?id=3385048068\">руководстве.</a>

    ; Hotkeys Section
    Gui, Font, c666666 s11 Bold Q3, Segoe UI Black
    Gui, Add, GroupBox, x10 y170 w280 h300 Center Section, ПРИВЯЗКА КЛАВИШ
    Gui, Font, s10

; Определяем общие координаты для выравнивания
labelX := "xs+10"     ; X координата для всех текстовых меток
hotkeyX := "x+1"      ; X координата для всех полей Hotkey (после метки)
hotkeyWidth := "w70"  ; Ширина всех полей Hotkey
labelWidth := "w180"  ; Ширина текстовых меток для выравнивания

Gui, Add, Text, cBlue %labelX% yp+30 %labelWidth% Left, Стоп:
Gui, Add, Hotkey, %hotkeyX% yp %hotkeyWidth% h20 vStop_key, %Stop_key%

Gui, Add, Text, cBlue %labelX% yp+30 %labelWidth% Left, Пауза:
Gui, Add, Hotkey, %hotkeyX% yp %hotkeyWidth% h22 vPause_key, %Pause_key%

Gui, Add, Text, cBlue %labelX% yp+30 %labelWidth% Left, Перезапуск:
Gui, Add, Hotkey, %hotkeyX% yp %hotkeyWidth% h22 vReset_key, %Reset_key%

Gui, Add, Text, cRed %labelX% yp+30 %labelWidth% Left, Открытие сундуков:
Gui, Add, Hotkey, %hotkeyX% yp %hotkeyWidth% h20 vOpenChests_key, %OpenChests_key%

Gui, Add, Text, cRed %labelX% yp+30 %labelWidth% Left, Утиль предметов:
Gui, Add, Hotkey, %hotkeyX% yp %hotkeyWidth% h20 vSalvage_items_key, %Salvage_items_key%

Gui, Add, Text, cRed %labelX% yp+30 %labelWidth% Left, Утиль крас предметов:
Gui, Add, Hotkey, %hotkeyX% yp %hotkeyWidth% h20 vSalvage_red_items_key, %Salvage_red_items_key%

Gui, Add, Text, cRed %labelX% yp+30 %labelWidth% Left, Перебросить свойства:
Gui, Add, Hotkey, %hotkeyX% yp %hotkeyWidth% h20 vReroll_Properties_key, %Reroll_Properties_key%

Gui, Add, Text, cRed %labelX% yp+30 %labelWidth% Left, Улучшение атанора:
Gui, Add, Hotkey, %hotkeyX% yp %hotkeyWidth% h20 vAtanor_key, %Atanor_key%

Gui, Add, Text, cMaroon %labelX% yp+30 %labelWidth% Left, Закрыть скрипт:
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

            SB_SetText("Горячие клавиши сохранены в пресет '" . savedPreset . "'")
        } catch e {
            LogAction("Ошибка сохранения горячих клавиш: " . e.Message)
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
        MsgBox, 48, Ошибка, Целевое окно не установлено. Установите окно перед запуском скрипта.
        return false
    }

    if (requireCoords) {
        Loop, Parse, coordsList, `,
        {
            coord := Trim(A_LoopField)
            if (%coord% == "") {
                MsgBox, 16, Ошибка, Координата %coord% не установлена. Установите необходимые координаты перед запуском.
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
        ToolTip, Скрипт приостановлен. Активируйте окно: %Target_Window%
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
        ToolTip, %action%... Цикл: %count%/%LoopCount%
    else
        ToolTip, %action%... Цикл: %count% (бесконечно)
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
    ;if (X1 != "" && Y1 != "")
        GuiControl,, Pos_1, %X1% %Y1%
return

SetPosGreen:
    Stop := 1
    GetClickPos(X2, Y2)
    ;if (X2 != "" && Y2 != "")
        GuiControl,, Pos_2, %X2% %Y2%
return

SetPosBlue:
    Stop := 1
    GetClickPos(X3, Y3)
    ;if (X3 != "" && Y3 != "")
        GuiControl,, Pos_3, %X3% %Y3%
return

SetPosYellow:
    Stop := 1
    GetClickPos(X4, Y4)
    ;if (X4 != "" && Y4 != "")
        GuiControl,, Pos_4, %X4% %Y4%
return

SetPosAtanor:
    Stop := 1
    GetClickPos(X5, Y5)
    ;if (X5 != "" && Y5 != "")
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
    ;if (OCR_X != "" && OCR_Y != "" && OCR_W != "" && OCR_H != "")
        GuiControl,, OCR_Pos, %OCR_X% %OCR_Y% %OCR_W% %OCR_H%
return

SetGreenDustArea:
    Stop := 1
    GetGreenDustArea(GreenDust_X1, GreenDust_Y1, GreenDust_X2, GreenDust_Y2)
    ;if (GreenDust_X1 != "" && GreenDust_Y1 != "" && GreenDust_X2 != "" && GreenDust_Y2 != "")
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
        ToolTip, Щелкните ЛКМ для выбора`nНажмите ESC для отмены`n`nТекущие координаты:`nX=%currentX% Y=%currentY%

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
        ToolTip, Дважды щелкните ЛКМ по окну`nНажмите ESC для отмены`n`nТекущее окно: %Temp_Window%

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

    ; Первый цикл - ожидание нажатия ЛКМ
    Loop {
        ; Проверка нажатия ESC
        if (GetKeyState("Esc", "P")) {
            ToolTip
            X := Y := W := H := ""
            return
        }

        ; Проверка нажатия ЛКМ для выхода из первого цикла
        if (GetKeyState("LButton", "P")) {
            break
        }
        MouseGetPos, currentX, currentY
        ToolTip, Зажмите ЛКМ и выделите область свойств предмета`nНажмите ESC для отмены`n`nТекущие координаты:`nX1=%currentX% Y1=%currentY%

        Sleep, 10
    }

    ; Получаем начальные координаты после нажатия ЛКМ
    MouseGetPos, startX, startY

    ; Второй цикл - отслеживание выделения области
    Loop {
        ; Проверка нажатия ESC
        if (GetKeyState("Esc", "P")) {
            ToolTip
            X := Y := W := H := ""
            return
        }

        ; Проверяем, удерживается ли еще ЛКМ
        if (!GetKeyState("LButton", "P")) {
            break
        }

        ; Получаем текущие координаты
        MouseGetPos, currentX, currentY

        ; Вычисляем размеры области
        width := Abs(currentX - startX)
        height := Abs(currentY - startY)
        minX := Min(startX, currentX)
        minY := Min(startY, currentY)

        ; Показываем информацию о выделении с движущимся ToolTip
        ToolTip, Выделение: %minX% %minY% [%width% x %height%]`n`nОтпустите ЛКМ для завершения`nНажмите ESC для отмены

        Sleep, 10
    }

    ; Получаем конечные координаты
    MouseGetPos, endX, endY

    ; Вычисляем итоговую область
    X := Min(startX, endX)
    Y := Min(startY, endY)
    W := Abs(startX - endX)
    H := Abs(startY - endY)

    ToolTip, Область свойств предмета установлена!`nX: %X% Y: %Y%`nW: %W% H: %H%
    Sleep, DELAY_ANIMATION
    ToolTip
}


GetGreenDustArea(ByRef X1, ByRef Y1, ByRef X2, ByRef Y2) {
    ; Первый цикл - ожидание нажатия ЛКМ
    Loop {
        ; Проверка нажатия ESC
        if (GetKeyState("Esc", "P")) {
            ToolTip
            X1 := Y1 := X2 := Y2 := ""
            return
        }

        ; Проверка нажатия ЛКМ для выхода из первого цикла
        if (GetKeyState("LButton", "P")) {
            break
        }
                ; Получаем текущие координаты
        MouseGetPos, currentX, currentY
        ToolTip, Зажмите ЛКМ и выделите область зеленой пыли`nНажмите ESC для отмены`n`nТекущие координаты:`nX1=%currentX% Y1=%currentY%

        Sleep, 10
    }

    ; Получаем начальные координаты после нажатия ЛКМ
    MouseGetPos, startX, startY

    ; Второй цикл - отслеживание выделения области
    Loop {
        ; Проверка нажатия ESC
        if (GetKeyState("Esc", "P")) {
            ToolTip
            X1 := Y1 := X2 := Y2 := ""
            return
        }

        ; Проверяем, удерживается ли еще ЛКМ
        if (!GetKeyState("LButton", "P")) {
            break
        }

        ; Получаем текущие координаты
        MouseGetPos, currentX, currentY

        ; Вычисляем текущие координаты прямоугольника
        currentX1 := Min(startX, currentX)
        currentY1 := Min(startY, currentY)
        currentX2 := Max(startX, currentX)
        currentY2 := Max(startY, currentY)

        ; Показываем информацию о выделении с движущимся ToolTip
        ToolTip, Выделение: %currentX1% %currentY1% - %currentX2% %currentY2%`n`nОтпустите ЛКМ для завершения`nНажмите ESC для отмены

        Sleep, 10
    }

    ; Получаем конечные координаты
    MouseGetPos, endX, endY

    ; Вычисляем итоговую область как X1,Y1,X2,Y2
    X1 := Min(startX, endX)
    Y1 := Min(startY, endY)
    X2 := Max(startX, endX)
    Y2 := Max(startY, endY)

    ToolTip, Область зеленой пыли установлена!`nX1: %X1% Y1: %Y1%`nX2: %X2% Y2: %Y2%
    Sleep, DELAY_ANIMATION
    ToolTip
}
; ============================================================================
; CONTROL FUNCTIONS
; ============================================================================

Stope:
    Stop := 1
    ToolTip
    LogAction("Скрипт остановлен пользователем")
return

Pauza:
    if (A_IsPaused) {
        Pause, , 1
        ToolTip
        LogAction("Скрипт возобновлен")
    } else {
        Pause, , 1
        ToolTip, ПАУЗА
        LogAction("Скрипт приостановлен")
        Sleep, DELAY_SHORT
    }
return

Reset:
    LogAction("Перезапуск скрипта")
    Reload
return

CloseScript:
    LogAction("Скрипт закрыт пользователем")
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
        LogAction("Открытие сундуков: ошибка валидации")
        return
    }

    LogAction("Начало открытия сундуков. Лимит циклов: " . (LoopCount != "" ? LoopCount : "бесконечно"))

    loop {
        if (!CheckActiveWindow())
            continue

        if (Stop == 1 || (LoopCount != "" && CycleCount >= LoopCount))
            break

        SendInput, {Space}
        Sleep, DELAY_LONG
        SendInput, {Space}

        CycleCount++
        UpdateTooltip("Открытие сундуков", CycleCount)
        Sleep, DELAY_LONG
    }

    ; Добавлен звук при завершении циклов
    if (LoopCount != "" && CycleCount >= LoopCount) {
        SoundPlay, *64
    }

    ToolTip
    LogAction("Открытие сундуков завершено. Всего циклов: " . CycleCount)
return

SalvageItems:
    Stop := 0
    CycleCount := 0
    Gui, Submit, NoHide

    if (!ValidateScript(true, "X1,Y1,X2,Y2,X3,Y3,X4,Y4,X6,Y6")) {
        LogAction("Утилизация предметов: ошибка валидации")
        return
    }

    LogAction("Начало утилизации предметов. Лимит циклов: " . (LoopCount != "" ? LoopCount : "бесконечно"))

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
        UpdateTooltip("Утилизация предметов", CycleCount)
        Sleep, DELAY_MEDIUM
    }

    ; Добавлен звук при завершении циклов
    if (LoopCount != "" && CycleCount >= LoopCount) {
        SoundPlay, *64
    }

    ToolTip
    LogAction("Утилизация предметов завершена. Всего циклов: " . CycleCount)
return

SalvageRedItems:
    Stop := 0
    CycleCount := 0
    Gui, Submit, NoHide

    if (!ValidateScript(true, "X1,Y1,X2,Y2,X3,Y3,X4,Y4,X6,Y6")) {
        LogAction("Утилизация предметов: ошибка валидации")
        return
    }

    LogAction("Начало утилизации красных предметов. Лимит циклов: " . (LoopCount != "" ? LoopCount : "бесконечно"))

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
        UpdateTooltip("Утилизация красных предметов", CycleCount)
        Sleep, DELAY_SHORT
    }

    ; Добавлен звук при завершении циклов
    if (LoopCount != "" && CycleCount >= LoopCount) {
        SoundPlay, *64
    }

    ToolTip
    LogAction("Утилизация красных предметов завершена. Всего циклов: " . CycleCount)
return

UpgradeAtanor:
    Stop := 0
    CycleCount := 0
    Gui, Submit, NoHide

    if (!ValidateScript(true, "X5,Y5")) {
        LogAction("Улучшение атанора: ошибка валидации")
        return
    }

    LogAction("Начало улучшения атанора. Лимит циклов: " . (LoopCount != "" ? LoopCount : "бесконечно"))

    loop {
        if (!CheckActiveWindow())
            continue

        if (Stop == 1 || (LoopCount != "" && CycleCount >= LoopCount))
            break

        Sleep, 150
        MouseMove, %X5%, %Y5%
        Click

        CycleCount++
        UpdateTooltip("Улучшение атанора", CycleCount)
        Sleep, DELAY_ANIMATION
    }

    ; Добавлен звук при завершении циклов
    if (LoopCount != "" && CycleCount >= LoopCount) {
        SoundPlay, *64
    }

    ToolTip
    LogAction("Улучшение атанора завершено. Всего циклов: " . CycleCount)
return

; ============================================================================
; REROLL PROPERTIES FUNCTION
; ============================================================================

RerollProperties:
    Gui, Submit, NoHide

    if (Target_Window == "") {
        MsgBox, 48, Ошибка, Целевое окно не установлено. Установите окно перед запуском скрипта.
        LogAction("Переброс свойств: окно не установлено")
        return
    }

    if (OCR_X = "" || OCR_Y = "" || OCR_W = "" || OCR_H = "") {
        MsgBox, 16, Ошибка, Область свойств предмета не установлена!`nВыделите область перед запуском.
        LogAction("Переброс свойств: область свойств предмета не установлена")
        return
    }

    if (GreenDust_X1 = "" || GreenDust_Y1 = "" || GreenDust_X2 = "" || GreenDust_Y2 = "") {
        MsgBox, 16, Ошибка, Область зеленой пыли не установлено!`nВыделите область перед запуском.
        LogAction("Переброс свойств: область зеленой пыли не установлена")
        return
    }

    if ((!X6 || !Y6)) {
        MsgBox, 16, Ошибка,  Местоположение кнопки для изменения свойства/утилизации не задано!`nУстановите координаты перед запуском.
        return
    }

    if (!X7 || !Y7) {
        MsgBox, 16, Ошибка, Местоположение предмета в слоте не задано!`nУстановите координаты перед запуском.
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
    global X7, Y7

    ; Split property names from menu selection
    Properties := StrSplit(itemName, " - ")
    Property1 := Properties[1]
    Property2 := Properties[2]

    LogAction("Начало переброса свойств: " . itemName)

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

    ToolTip, Начинаем переброс свойств...`nИщем: %Property1% - %Property2%
    Sleep, 1000

    ; Флаг для отслеживания причины остановки
    dustExhausted := false

    loop {
        ; Check if correct window is active
        if (!CheckActiveWindow())
            continue

        ; Check stop conditions
        if (Stop == 1 || (LoopCount != "" && CycleCount >= LoopCount)) {
            ToolTip
            LogAction("Переброс свойств остановлен. Циклов: " . CycleCount)
            break
        }

        ; Check for green dust availability (crafting resource)
        PixelSearch, Px, Py, %dustX1%, %dustY1%, %dustX2%, %dustY2%, %COLOR_GREEN_DUST%, %COLOR_TOLERANCE%, RGB Fast

        ; Если пыль не найдена, устанавливаем флаг и выходим
        if (ErrorLevel) {
            dustExhausted := true
            ToolTip, Зеленая пыль не найдена!
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
                text := OCR(ocrArea, "rus")
                lastOCRTime := A_TickCount
                lastOCRText := text
            } catch e {
                LogAction("Ошибка OCR: " . e.Message)
                text := ""
            }

            ; Display OCR results
            ToolTip, Распознанный текст:`n%text%`n---`nИщем: %Property1% - %Property2%`nЦикл: %CycleCount%/%LoopCount%
            Sleep, 1000

            ; Check if target properties found
            If (InStr(text, Property1) && InStr(text, Property2)) {
                SoundPlay, *64
                ToolTip, Нужные свойства найдены!`n%Property1% - %Property2%
                LogAction("Свойства найдены после " . CycleCount . " циклов")
                Sleep, 5000
                ToolTip
                return
            } else {

            MouseMove, %X6%, %Y6%
            Click, down
            Sleep, DELAY_LONG
            Click, up
            Sleep, DELAY_ANIMATION

                ; ПЕРЕМЕЩАЕМ МЫШЬ В ЗАДАННУЮ КООРДИНАТУ ПОСЛЕ ПЕРЕБРОСА
                if (X7 != "" && Y7 != "") {
                    MouseMove, %X7%, %Y7%
                    Sleep, 1000
                }

                CycleCount++
                UpdateTooltip("Переброс свойств", CycleCount)
                Sleep, DELAY_SHORT
                ; Recheck green dust availability
                PixelSearch, Px, Py, %dustX1%, %dustY1%, %dustX2%, %dustY2%, %COLOR_GREEN_DUST%, %COLOR_TOLERANCE%, RGB Fast
            }
        }
    }

    ; Если вышли из цикла из-за окончания пыли, показываем сообщение
    if (dustExhausted && !Stop && (LoopCount = "" || CycleCount < LoopCount)) {
        ToolTip, Скрипт остановлен.
        LogAction("Переброс свойств завершен. Циклов: " . CycleCount . " (пыль не найдена)")
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
    global X7, Y7

    ; Split property names from menu selection
    Properties := StrSplit(itemName, " - ")
    Property1 := Properties[1]
    Property2 := Properties[2]

    LogAction("Начало переброса свойств: " . itemName)

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

    ToolTip, Начинаем переброс свойств...`nИщем: %Property1% - %Property2%
    Sleep, 1000


    loop {

        ; Check stop conditions - ДОБАВЛЕНО В НАЧАЛО ЦИКЛА
        if (Stop == 1 || (LoopCount != "" && CycleCount >= LoopCount)) {
            ToolTip
            LogAction("Переброс свойств остановлен пользователем. Циклов: " . CycleCount)
            break
        }

        ; Check if correct window is active
        if (!CheckActiveWindow())
            continue

        ; Check for green dust availability (crafting resource)
        PixelSearch, Px, Py, %dustX1%, %dustY1%, %dustX2%, %dustY2%, %COLOR_GREEN_DUST%, %COLOR_TOLERANCE%, RGB Fast

        ; Если пыль не найдена, выходим
        if (ErrorLevel) {
            ToolTip, Зеленая пыль не найдена! Скрипт остановлен.
            LogAction("Переброс свойств завершен. Циклов: " . CycleCount . " (пыль не найдена)")
            Sleep, 3000
            ToolTip
            break
        }
        ; OCR чтение свойств предмета
        if (A_TickCount - lastOCRTime >= DELAY_OCR_COOLDOWN) {
            try {
                text := OCR(ocrArea, "rus")
                lastOCRTime := A_TickCount
                lastOCRText := text
            } catch e {
                LogAction("Ошибка OCR: " . e.Message)
                text := ""
            }

            ; Display OCR results
            ToolTip, Распознанный текст:`n%text%`n---`nИщем: %Property1% - %Property2%`nЦикл: %CycleCount%/%LoopCount%
            Sleep, 1000

            ; Check if target properties found
            If (InStr(text, Property1) && InStr(text, Property2)) {
                SoundPlay, *64
                ToolTip, Нужные свойства найдены!`n%Property1% - %Property2%
                LogAction("Свойства найдены после " . CycleCount . " циклов")
                Sleep, 5000
                ToolTip
                return
            }
        }

            MouseMove, %X6%, %Y6%
            Click, down
            Sleep, DELAY_LONG
            Click, up
            Sleep, DELAY_ANIMATION

                ; MOVE MOUSE TO SPECIFIED COORDINATE AFTER REROLL
                if (X7 != "" && Y7 != "") {
                    MouseMove, %X7%, %Y7%
                    Sleep, 1000
                }

        CycleCount++
        UpdateTooltip("Переброс свойств", CycleCount)
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

    ;LogAction("Загрузка пресета: " . savedPreset)

    ; Убрана конструкция try-catch для лучшей диагностики ошибок
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
    ;SB_SetText("Пресет '" . savedPreset . "' загружен")

return

SavePreset:
    Gui, Submit, NoHide

    if (savedPreset = "") {
        SB_SetText("Введите имя пресета!")
        return
    }

    LogAction("Сохранение пресета: " . savedPreset)

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

    ; Сохранение горячих клавиш
    IniWrite, %Stop_key%, %presetFile%, %savedPreset%, Stop_key
    IniWrite, %Pause_key%, %presetFile%, %savedPreset%, Pause_key
    IniWrite, %Reset_key%, %presetFile%, %savedPreset%, Reset_key
    IniWrite, %OpenChests_key%, %presetFile%, %savedPreset%, OpenChests_key
    IniWrite, %Salvage_items_key%, %presetFile%, %savedPreset%, Salvage_items_key
    IniWrite, %Salvage_red_items_key%, %presetFile%, %savedPreset%, Salvage_red_items_key
    IniWrite, %Reroll_Properties_key%, %presetFile%, %savedPreset%, Reroll_Properties_key
    IniWrite, %Atanor_key%, %presetFile%, %savedPreset%, Atanor_key
    IniWrite, %Close_script_key%, %presetFile%, %savedPreset%, Close_script_key

    SB_SetText("Пресет '" . savedPreset . "' сохранен!")
    GoSub, UpdatePresetList

return

DeletePreset:
    Gui, Submit, NoHide

    if (savedPreset = "") {
        SB_SetText("Выберите пресет для удаления")
        return
    }

    MsgBox, 4, Подтверждение, Вы уверены, что хотите удалить пресет "%savedPreset%"?
    IfMsgBox, Yes
    {
        IniDelete, %presetFile%, %savedPreset%
        SB_SetText("Пресет '" . savedPreset . "' удален")
        LogAction("Пресет удален: " . savedPreset)
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
    LogAction("Приложение закрыто")
    ExitApp
return

InfoWindowGuiClose:
    Gui, Destroy
return