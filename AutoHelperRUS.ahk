VERSION := 1.0
RegExMatch(A_ScriptName, "^(.*?)\.", basename)
WINTITLE := basename1 " " VERSION

presetsDir := A_AppData "\" basename1
if !FileExist(presetsDir)
    FileCreateDir, %presetsDir%

#SingleInstance force
#NoENV
SetBatchLines -1    ; have the script run at maximum speed and never sleep
ListLines Off       ; a debugging option
SetMouseDelay, -1  ; Убирает задержки между действиями мыши
CoordMode, Mouse, Client  ; Определяет координаты мыши относительно клиентской области окна
SetTitleMatchMode, 2  ; Позволяет частичное совпадение заголовков окон при поиске
SetControlDelay, -1  ; Убирает задержки при работе с элементами GUI
DetectHiddenWindows, On  ; Позволяет обнаруживать скрытые окна
SetStoreCapsLockMode, Off  ; Отключает запоминание состояния клавиши CapsLock

;---------------------- ГЛОБАЛЬНЫЕ ПЕРЕМЕННЫЕ ----------------------
global X1, Y1, X2, Y2, X3, Y3, X4, Y4, X5, Y5, X6, Y6, X7, Y7
global LoopCount
Target_Window := ""
global Stop:=0
LoopCount := "" ; Количество циклов (по умолчанию пустое)
Stop_key := "F1", Pause_key := "F2", Reset_key := "F3", OpenChests_key := "F5", Salvage_items_key := "F7", Salvage_red_items_key := "F9", Dust_convert_key := "F11", Atanor_key := "F12", Close_script_key := "^ESC"

;====================== ОБНОВЛЕНИЕ ВСЕХ КЛАВИШ ======================
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
;---------------------- СОЗДАНИЕ GUI ----------------------
Gui,+AlwaysOnTop -DPIScale +OwnDialogs

;----------------------- Панель пресетов ----------------------
Gui, Add, Text, xm section, Список пресетов:
Gui, Add, ComboBox, x+5 vfrmSAVEDPRESET gPresetChange
Gui, Add, Button, x+5 h21 w70 gSavePreset, Сохранить
Gui, Add, Button, x+5 h21 w70 gDeletePreset vDELETEBUTTON, Удалить
Gui, Add, StatusBar

;------------------- Настройка координат ----------------------
Gui, Font, c535353 s13 Bold Q3, Segoe UI Black
Gui, Add, GroupBox, x10 y40 w380 h260 Center Section, Настройка скриптов
Gui, Font
Gui, Font, s11, Segoe UI
Gui, Add, Text,  xs+10 yp+25, #1 Позиция кнопки для открытия сундуков:
Gui, Font, s9,
Gui, Add, Button, xs+10 yp+30 w80 h25 gSet_Pos_Chests, Установить
Gui, Add, Edit, x+10 yp+1 w80 h20 vPos_1,
Gui, Font, s11, Segoe UI
Gui, Add, Text, xs+10 yp+30, #2 Позиция кнопки для улучшения Атанора:
Gui, Font, s9
Gui, Add, Button, xs+10 yp+30 w80 h25 gSet_Pos_Atanor, Установить
Gui, Add, Edit, x+10 yp+1 w80 h20 vPos_7,
Gui, Font, s11, Segoe UI
Gui,Add,Text,xs+10 yp+30,#3 Позиция кнопок для утилизации предметов:

Gui, Font, s9
Gui,Add,Button,xs+10 yp+25 w140 h25 gSet_Pos_Button, Кнопка утилизации
Gui,Add,Edit,x+10 yp+3 w140 h20 vPos_6,
Gui,Add,Button,xs+10 yp+30 w70 h20 gSet_Pos_White, Белая
Gui,Add,Edit,x+10 w80 h20 vPos_2,
Gui,Add,Button,x+10 w70 h20 gSet_Pos_Green, Зеленая
Gui,Add,Edit,x+10 w80 h20 vPos_3,
Gui,Add,Button,xs+10 yp+25 w70 h20 gSet_Pos_Blue, Синяя
Gui,Add,Edit,x+10 w80 h20 vPos_4,
Gui,Add,Button,x+10 w70 h20 gSet_Pos_Yellow, Желтая
Gui,Add,Edit,x+10 w80 h20 vPos_5,

;--------------------- Дополнительные настройки ---------------
Gui, Add, Text, x25 y690, Общий ограничитель циклов (опционально):
Gui, Add, Edit, x+10 w80 h20 vLoopCount

Gui, Add, Button, x25 y720 w120 h30 gSet_Location, Установите окно:
Gui, Add, Edit, x+15 yp+3 w210 h25 vTarget_Window, %Target_Window%

;------------------------ Кнопки действий ---------------------
Gui, Font, c535353 s13 Bold Q3, Segoe UI Black
Gui, Add, GroupBox, x10 y295 w380 h200 Center Section, Кнопки

Gui, Font, s9 Bold, Segoe UI
Gui, Add, Button, xs+15 yp+30 w110 h35 gOpenChests, ОТКРЫТИЕ СУНДУКОВ
Gui, Add, Button, x+10 w110 h35 gSalvage_items, УТИЛИЗАЦИЯ ПРЕДМЕТОВ
Gui, Add, Button, x+10 w110 h35 gUpgrade_atanor, УЛУЧШЕНИЕ АТАНОРА
Gui, Add, Button, xs+15 yp+45 w170 h35 gSalvage_red_items, УТИЛИЗАЦИЯ`nКРАСНЫХ ПРЕДМЕТОВ
Gui, Add, Button, x+10 w170 h35 gDust_convert, МЕХАНИКИ КРАФТА
Gui, Add, Button, xs+40 yp+45 w300 h25 gTarget_Info, <><><> ИНФА И ПРИВЯЗКА КЛАВИШ <><><>
Gui, Add, Button, xs+15 yp+35 w110 h35 gStope, СТОП
Gui, Add, Button, x+10 w110 h35 gPauza, ПАУЗА
Gui, Add, Button, x+10 w110 h35 gReset, ПЕРЕЗАПУСК
Gui, Font

Gui, Font, c666666 s12 Bold Q3, Segoe UI Black
Gui, Add, GroupBox, x10 y500 w380 h180 Center Section, Горячие клавиши по умолчанию
Gui, Font
Gui, Font, s11
Gui, Add, Text, cBlue xs+30 yp+35, Стоп = F1
Gui, Add, Text, cBlue x+10, Пауза = F2
Gui, Add, Text, cBlue x+10, Перезапуск / сброс = F3
Gui, Add, Text, cRed xs+30 yp+30, Открытие Сундуков = F5
Gui, Add, Text, cRed x+10, Утил. Предметов = F7
Gui, Add, Text, cRed xs+90 yp+30, Утил. Красных Предметов = F9
Gui, Add, Text, cRed xs+30 yp+30, Механики Крафта = F11
Gui, Add, Text, cRed x+10, Улуч. Атанора = F12
Gui, Add, Text, cMaroon xs+90 yp+30, Закрыть Скрипт = Ctrl+Esc
Gui, Font
Gui, Show, w400 h780, %WINTITLE%
GoSub, UpdatePresetList
return
;────────────────────────────────────────────────────────────
; Окно справки (Target_Info)
;────────────────────────────────────────────────────────────
Target_Info:
Gui, NewWindow:New, , Помощь
Gui, +AlwaysOnTop -DPIScale +Owner1
Gui, Add, Picture, x10 y33, icon.png
Gui, Font, c666666 s11 Bold Q3, Segoe UI Black
Gui, Add, GroupBox, x10 y10 w280 h150 Center Section, ИНФА
Gui, Font
Gui, Add, Text, xs+145 yp+25, Автор:   !ChadMasodin
Gui, Add, Text, xs+145 yp+15, Дата выхода:  27/12/24
Gui, Add, Text, xs+145 yp+15, Дата обновы:  21/02/25
Gui, Add, Link, xs+10 yp+30, AutoHelper — это скрипт для Vermintide 2,`nавтоматизирующий ряд рутинных процессов в игре.`nБолее подробную информацию о нем можно`nнайти в этом <a href="https://steamcommunity.com/sharedfiles/filedetails/?id=3385048068\">руководстве.</a>
Gui, Font, s11 Bold Q3, Segoe UI
Gui, Add, Button, x10 y480 w280 h45 gSaveTargetInfo, OK
Gui, Font
;────────────────────────────────────────────────────────────
; Привязка клавиш (GUI для горячих клавиш)
;────────────────────────────────────────────────────────────
Gui, Font, c666666 s11 Bold Q3, Segoe UI Black
Gui, Add, GroupBox, x10 y170 w280 h300 Center Section, ПРИВЯЗКА КЛАВИШ
Gui, Font

Gui, Add, Text, cBlue xs+10 yp+30, Стоп:
Gui, Add, Hotkey, x+10 w90 h20 vStop_key, %Stop_key%
Gui, Add, Text, cBlue xs+10 yp+30, Пауза:
Gui, Add, Hotkey, x+10 w90 h22 vPause_key, %Pause_key%
Gui, Add, Text, cBlue xs+10 yp+30, Перезапуск / сброс:
Gui, Add, Hotkey, x+10 w90 h22 vReset_key, %Reset_key%
Gui, Add, Text, cRed xs+10 yp+30, Открытие сундуков:
Gui, Add, Hotkey, x+10 w90 h20 vOpenChests_key, %OpenChests_key%
Gui, Add, Text, cRed xs+10 yp+30, Утил. предметов:
Gui, Add, Hotkey, x+10 w90 h20 vSalvage_items_key, %Salvage_items_key%
Gui, Add, Text, cRed xs+10 yp+30, Утил. красных предметов:
Gui, Add, Hotkey, x+10 w80 h20 vSalvage_red_items_key, %Salvage_red_items_key%
Gui, Add, Text, cRed xs+10 yp+30, Преобразование пыли:
Gui, Add, Hotkey, x+10 w90 h20 vDust_convert_key, %Dust_convert_key%
Gui, Add, Text, cRed xs+10 yp+30, Улучшение атанора:
Gui, Add, Hotkey, x+10 w90 h20 vAtanor_Key, %Atanor_key%
Gui, Add, Text, cMaroon xs+10 yp+30, Закрыть скрипт:
Gui, Add, Hotkey, x+10 w90 h20 vClose_script_key, %Close_script_key%
Gui, Show, w300 h540,
return
;────────────────────────────────────────────────────────────
; Обработчики событий для установки окна и обновления GUI
;────────────────────────────────────────────────────────────
Set_Location:
    Target_Window := Set_Window(Target_Window)
    GuiControl,, Target_Window, %Target_Window%
return
; Универсальная функция для обновления горячей клавиши
UpdateKey(KeyVariable, Label)
{
    Global
    GuiControlGet, New_Key, , %KeyVariable%
    Current_Key := %KeyVariable%
    if (New_Key != Current_Key)
    {
        Hotkey, %Current_Key%, %Label%, Off  ; Убираем старую привязку
        %KeyVariable% := New_Key  ; Обновляем переменную
        Hotkey, %New_Key%, %Label%, On  ; Устанавливаем новую привязку
    ; Автосохранение в активный пресет
        if (frmSAVEDPRESET != "")
            IniWrite, %New_Key%, %A_AppData%\%basename1%\presets.ini, %frmSAVEDPRESET%, %KeyVariable%
    }
}
; Функция SendKey, определённая на верхнем уровне
SendKey(key, delay := 300) {
    Send, {%key% down}
    Send, {%key% up}
    Sleep, delay
}
;────────────────────────────────────────────────────────────
; Функции для установки координат клика
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
; Функции для управления выполнением скрипта
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
        Tooltip, ПАУЗА
        Sleep, 100
    }
return

Reset:
    Reload
return
;────────────────────────────────────────────────────────────
; Функция автоматизации открытия сундуков
;────────────────────────────────────────────────────────────
OpenChests:
    Stop := 0
    CycleCount := 0  ; Счетчик итераций
    Gui, Submit, NoHide  ; Обновляем переменные из GUI

    if (LoopCount != "" && LoopCount < 0) {
        MsgBox, 16, Ошибка, Пожалуйста, введите положительное число циклов или оставьте поле пустым для бесконечного выполнения!
        return
    }
    if (Target_Window == "") {
        MsgBox, 48, Ошибка, Целевое окно не установлено!
        return
    }
    if (!X1 || !Y1) {
        MsgBox, 16, Ошибка, Местоположение X и Y не задано! Установите координаты перед запуском.
        return
    }
    loop {
        WinGetActiveTitle, Current_Window
        if (Current_Window != Target_Window) {
            ToolTip, Скрипт приостановлен. Активируйте окно: %Target_Window%
            Sleep, 500
            continue
        }
        if (Stop == 1 || (LoopCount != "" && CycleCount >= LoopCount))
            break
        ;Sleep, 250
        Sleep, 50
        MouseMove, %X1%, %Y1%
        Click
        ;Sleep, 600
        Sleep, 500
        Send, {Space}
        ;Sleep, 200
        CycleCount++  ; Увеличиваем счетчик итераций
        ToolTip, Работает... Цикл: %CycleCount%/%LoopCount%
        ;Sleep, 2200
    }
    ToolTip
return

;────────────────────────────────────────────────────────────
; Функция для утилизации предметов
;────────────────────────────────────────────────────────────
Salvage_items:
    Stop := 0
    CycleCount := 0 ; Счетчик итераций
    Gui, Submit, NoHide
    if (LoopCount != "" && LoopCount < 0) {
        MsgBox, 16, Ошибка, Пожалуйста, введите положительное число циклов или оставьте поле пустым для бесконечного выполнения!
        return
    }
    if (Target_Window == "") {
        MsgBox, 48, Ошибка, Целевое окно не установлено!
        return
    }
    if ((!X2 || !Y2) || (!X3 || !Y3) || (!X4 || !Y4) || (!X5 || !Y5) || (!X6 || !Y6)) {
        MsgBox, 16, Ошибка, Местоположение X и Y не задано! Установите координаты перед запуском.
        return
    }
    loop {
        WinGetActiveTitle, Current_Window
        if (Current_Window != Target_Window) {
            ToolTip, Скрипт приостановлен. Активируйте окно: %Target_Window%
            Sleep, 500
            continue
        }
        if (Stop == 1 || (LoopCount != "" && CycleCount >= LoopCount))
            break
        Sleep, 500
        MouseMove, %X2%, %Y2%  ; Белая кнопка
        Click
        Sleep, 300
        MouseMove, %X3%, %Y3%  ; Зеленая кнопка
        Click
        Sleep, 300
        MouseMove, %X4%, %Y4%  ; Синяя кнопка
        Click
        Sleep, 300
        MouseMove, %X5%, %Y5%  ; Желтая кнопка
        Click
        Sleep, 300
        MouseMove, %X6%, %Y6%  ; Кнопка утилизации
        Click, down
        Sleep, 500
        Click, up
        Sleep, 1000
        CycleCount++
        ToolTip, Работает... Цикл: %CycleCount%/%LoopCount%
        Sleep, 200
    }
    ToolTip
return

;────────────────────────────────────────────────────────────
; Функция для утилизации красных предметов
;────────────────────────────────────────────────────────────
Salvage_red_items:
{
    Stop := 0
    CycleCount := 0  ; Счетчик итераций
    Gui, Submit, NoHide
    if (LoopCount != "" && LoopCount < 0) {
        MsgBox, 16, Ошибка, Пожалуйста, введите положительное число циклов или оставьте поле пустым для бесконечного выполнения!
        return
    }
    if (Target_Window == "") {
        MsgBox, 48, Ошибка, Целевое окно не установлено!
        return
    }
    loop {
        WinGetActiveTitle, Current_Window
        if (Current_Window != Target_Window) {
            ToolTip, Скрипт приостановлен. Активируйте окно: %Target_Window%
            Sleep, 500
            continue
        }
        if (Stop == 1 || (LoopCount != "" && CycleCount >= LoopCount))
            break
        SendKey("Space")  ; Нажатие пробела
        Sleep, 100
        Loop, 4 {
            SendKey("Right")
            SendKey("Space")
        }
        Sleep, 100
        Send, {Space down}  ; Завершающее нажатие для утилизации
        Sleep, 500
        Send, {Space up}
        Sleep, 2000
        CycleCount++
        ToolTip, Работает... Цикл: %CycleCount%/%LoopCount%
        Sleep, 100
    }
}
ToolTip
return
;────────────────────────────────────────────────────────────
; Функция для улучшения Атанора
;────────────────────────────────────────────────────────────
Upgrade_atanor:
    Stop := 0
    CycleCount := 0  ; Счетчик итераций
    Gui, Submit, NoHide
    if (LoopCount != "" && LoopCount < 0) {
        MsgBox, 16, Ошибка, Пожалуйста, введите положительное число циклов или оставьте поле пустым для бесконечного выполнения!
        return
    }
    if (Target_Window == "") {
        MsgBox, 48, Ошибка, Целевое окно не установлено!
        return
    }
    if (!X7 || !Y7) {
        MsgBox, 16, Ошибка, Местоположение X и Y не задано! Установите координаты перед запуском.
        return
    }
    loop {
        WinGetActiveTitle, Current_Window
        if (Current_Window != Target_Window) {
            ToolTip, Скрипт приостановлен. Активируйте окно: %Target_Window%
            Sleep, 500
            continue
        }
        if (Stop == 1 || (LoopCount != "" && CycleCount >= LoopCount))
            break
        Sleep, 150
        MouseMove, %X7%, %Y7%
        Click
        CycleCount++
        ToolTip, Работает... Цикл: %CycleCount%/%LoopCount%
        Sleep, 1500
    }
    ToolTip
return
;────────────────────────────────────────────────────────────
; Функция преобразования пыли (удержание пробела)
;────────────────────────────────────────────────────────────
Dust_convert:
    Stop := 0
    CycleCount := 0  ; Счетчик итераций
    Gui, Submit, NoHide
    if (LoopCount != "" && LoopCount < 0) {
        MsgBox, 16, Ошибка, Пожалуйста, введите положительное число циклов или оставьте поле пустым для бесконечного выполнения!
        return
    }
    if (Target_Window == "") {
        MsgBox, 48, Ошибка, Целевое окно не установлено!
        return
    }
    Send, {Space}
    loop {
        WinGetActiveTitle, Current_Window
        if (Current_Window != Target_Window) {
            ToolTip, Скрипт приостановлен. Активируйте окно: %Target_Window%
            Sleep, 500
            continue
        }
        if (Stop == 1 || (LoopCount != "" && CycleCount >= LoopCount))
            break
        Sleep, 200
        Send, {Space down}
        CycleCount++
        ToolTip, Работает... Цикл: %CycleCount%/%LoopCount%
        Sleep, 1750
    }
    ToolTip
return
;────────────────────────────────────────────────────────────
; Функция завершения работы скрипта
;────────────────────────────────────────────────────────────
Close_script:
    ExitApp
return

UpdateAllKeys()
    Gui, Submit, NoHide
return

;────────────────────────────────────────────────────────────
; Функция получения позиции курсора по клику
;────────────────────────────────────────────────────────────
Get_Click_Pos(ByRef X, ByRef Y)
{
    isPressed := 0
    X := "", Y := "" ; Сбрасываем координаты при старте
    Loop {
        ; Проверка нажатия ESC
        if (GetKeyState("Esc", "P")) {
            ToolTip
            X := "", Y := "" ; Явный сброс координат
            return ; Прерываем функцию
        }

        Left_Mouse := GetKeyState("LButton")
        MouseGetPos, currentX, currentY
        ToolTip, Щелкните ЛКМ для выбора`nНажмите ESC для отмены`n`nТекущие координаты: `nX=%currentX% Y=%currentY%

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
; Функция для установки целевого окна
;────────────────────────────────────────────────────────────
Set_Window(Target_Window)
{
    isPressed := 0
    i := 0
    Target_Window := "" ; Сбрасываем значение при старте

    Loop {
        ; Проверка нажатия ESC
        if (GetKeyState("Esc", "P")) {
            ToolTip
            return "" ; Возвращаем пустую строку при отмене
        }

        Left_Mouse := GetKeyState("LButton")
        WinGetTitle, Temp_Window, A
        ToolTip, Дважды щелкните ЛКМ по окну`nНажмите ESC для отмены`n`nТекущее окно: %Temp_Window%

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
; do a guirestore for newly selected preset
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
      ; Загрузка Target_Window
    IniRead, twVal, %A_AppData%\%basename1%\presets.ini, %frmSAVEDPRESET%, Target_Window
    if (twVal != "ERROR") {
        GuiControl,, Target_Window, %twVal%
        Target_Window := twVal  ; Явное обновление переменной
    }

    ; Загрузка LoopCount
    IniRead, lcVal, %A_AppData%\%basename1%\presets.ini, %frmSAVEDPRESET%, LoopCount
    if (lcVal != "ERROR") {
        GuiControl,, LoopCount, %lcVal%
        LoopCount := lcVal  ; Явное обновление переменной
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

SavePreset:
    Gui, Submit, NoHide
    if (frmSAVEDPRESET = "") {
        SB_SetText("Введите имя пресета!")
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
    ; Сохранение только если значение не пустое
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

    SB_SetText("Пресет '" frmSAVEDPRESET "' сохранен!")
    GoSub, UpdatePresetList
return
;============================================================
; delete selected preset section from presets.ini
;============================================================

DeletePreset:
    gui, submit, nohide
    RegExMatch(A_ScriptName, "^(.*?)\.", basename)
    ; if drop down text is blank then error message and return
    if (frmSAVEDPRESET = "") {
        SB_SetText("Preset name required")
        return
    }
    ; delete entire section from ini file
    IniDelete, %A_AppData%\%basename1%\presets.ini, %frmSAVEDPRESET%

    SB_SetText("Пресет '" frmSAVEDPRESET "' удален")
    GoSub, UpdatePresetList  ; update drop down to show all preset section names in ini file
Return

;============================================================
; update drop down to show all preset section names in ini file, except section1
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
;============================================================
; when you click x or close button
;============================================================

GuiClose:
    Gui, Submit, NoHide      ; update control variables
ExitApp
;============================================================
; save all gui control values for active gui to ini file
;============================================================

GuiSave(inifile,section,begin="",end="")
{
    SplitPath, inifile, file, path, ext, base, drive     ; splitpath expects paths with \

    if (path = "") {   ; if no path given then use default path
        RegExMatch(A_ScriptName, "^(.*?)\.", basename)    ; dont use splitpath to get basename because it cant handle DeltaRush.1.3.exe
        inifile := A_AppData "\" basename1 "\" inifile
    }

    WinGet, List_controls, ControlList, A    ; get list of all controls active gui

    if (begin = "")
        flag := 0
    else
        flag := 1

    Loop, Parse, List_controls, `n
    {
        ;ControlGet, cid, hWnd,, %A_LoopField%         ; get the id of current control
        GuiControlGet, textvalue,,%A_Loopfield%,Text  ; get associated text
        GuiControlGet, vname, Name, %A_Loopfield%     ; get controls vname

        If (vname = "")   ; only save controls which have a vname
            continue

        if (begin = vname) {
            flag := 0
            continue
        }

        if (flag)
            continue

        if (end = vname)
            break

        GuiControlGet, value ,, %A_Loopfield%         ; get controls value
        value := RegExReplace(value, "`n", "|")       ; convert newlines to pipes (for multiline edit fields, because newlines are not valid for ini file)

        ; todo: truncate edit values to not exceed ini fieldsize limit (1024?)  OR blank (all or nothing)

        IniWrite, % value, %inifile%, %section%, %vname%

    }

   return
}

;============================================================
; Update gui controls with values from ini file.
;============================================================

GuiRestore(inifile,section)
{

    SplitPath, inifile, file, path, ext, base, drive     ; splitpath expects paths with \

    if (path = "") {   ; if no path given then use default path
        RegExMatch(A_ScriptName, "^(.*?)\.", basename)    ; dont use splitpath to get basename because it cant handle DeltaRush.1.3.exe
        inifile := A_AppData "\" basename1 "\" inifile
    }

    ;============================================================
    ; update gui controls with values from ini file
    ;============================================================

    WinGet, List_controls, ControlList, A   ; get list of all controls for active gui

    Loop, Parse, List_controls, `n
    {

        ;ControlGet, cid, hWnd,, %A_LoopField%         ; get the id of current control
        ;GuiControlGet, textvalue,,%A_Loopfield%,Text  ; get controls associated text
        GuiControlGet, vname, Name, %A_Loopfield%     ; get controls vname
        GuiControlGet, value ,, %A_Loopfield%         ; get controls value

        If (vname = "")   ; only process controls which have a vname
            continue

        IniRead, value, %inifile%, %section%, %vname%, ERROR

        if (value != "ERROR") {

            value := RegExReplace(value, "\|", "`n")       ; convert pipes to newlines (for multiline edit fields, because newlines are not valid for ini file)

            RegExMatch( A_Loopfield, "(.*?)\d+", name)   ; extract the control name without numbers
            if (name1 = "ComboBox") {
                GuiControl, ChooseString, %A_Loopfield%, %value%   ; select item in dropdownlist
            } else {
                GuiControl,  ,%A_Loopfield%, %value%    ; update the control
            }
        }

    }

    return

}
SaveTargetInfo:
    Gui, Submit
    UpdateAllKeys()
    ; Сохранение в активный пресет
    if (frmSAVEDPRESET != "")
    {
        Loop, Parse, % "Stop_key|Pause_key|Reset_key|OpenChests_key|Salvage_items_key|Salvage_red_items_key|Dust_convert_key|Atanor_key|Close_script_key", |
            IniWrite, % %A_LoopField%, %A_AppData%\%basename1%\presets.ini, %frmSAVEDPRESET%, %A_LoopField%
    }

    Gui, Destroy
return