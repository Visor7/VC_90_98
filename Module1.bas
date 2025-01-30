Attribute VB_Name = "Module1"
'*************************************************************************************************
'************Журнал ВЦ-98 v2.5 старт 17.12.2023***останя*зміна****02.08.2024**********************
'*************************************************************************************************

Public ws As Worksheet '
Public wsData As Worksheet 'Вказуємо аркуш ThisWorkbook.Worksheets("Data")
Public ws_Name As String ' Глобальна змінна Ім'я аркуша
Public actShIndex As Integer   'глобальна змінна Номер аркуша
Public room As String ' Глобальна змінна Кімната в журналі
Public check As Integer ' Глобальну змінна Кількість перевірок
Public press As Integer ' Глобальну змінна Тиск
Public ws_Visible As String ' Глобальна змінну Використання аркушу
Public ws_passw As String ' Глобальна змінна пароль аркушу
Public control_passw As String ' Глобальна змінна пароль Form control
Public ws_Number As Integer ' Глобальна змінна номер обраного OptionButton або Label в формі GeneralForm.
'Відповідае номеру аркушу та номеру розумної таблиці на аркуші Дані та номерам розумних таблиць на кожному аркуші 90 або 98


'*************************************************************************************************
'*************************************************************************************************
Public Sub FindParam()
    Set ws = ActiveSheet ' Отримуємо активний аркуш
    ws_Name = ws.name
    'MsgBox "Ім'я активного аркуша:" & ws_Name
    actShIndex = ActiveSheet.Index
    'MsgBox "№ активного аркуша:" & actShIndex
    ws_Visible = ThisWorkbook.Worksheets("Data").Cells(3, 1 + actShIndex).Value
    'MsgBox "Видимість аркуша:" & ms_Visible
    room = ThisWorkbook.Worksheets("Data").Cells(6, 1 + actShIndex).Value
    'MsgBox "Кімната в журналі: " & room
    check = ThisWorkbook.Worksheets("Data").Cells(9, 2).Value
    'MsgBox "Кількість перевірок:" & check
    press = ThisWorkbook.Worksheets("Data").Cells(10, 2).Value
    'MsgBox "Тиск: " & press
    If Left(ws.name, 2) = "Da" Then
        ws_passw = "lab"
    Else
        ws_passw = "lab123"
    End If
    control_passw = "lab"
End Sub
Public Sub ShowControlPanel()
    Application.ExecuteExcel4Macro "SHOW.TOOLBAR(""Ribbon"", True)"
    Application.DisplayFormulaBar = True
End Sub
Public Sub HideControlPanel()
    'заблокувати, приховати
    'Application.ExecuteExcel4Macro "SHOW.TOOLBAR(""Ribbon"", False)" 'ексель 2016,2019,2021,2023 приховати
    'Application.ExecuteExcel4Macro "SHOW.TOOLBAR(""Ribbon"", True)" 'ексель 2016,2019,2021,2023 показати
    'Application.ExecuteExcel4Macro "RibbonToggle()" 'єксель365 приховати - показати
    'Application.CommandBars("Ribbon").Enabled = False 'єксель2013 приховати
    'Application.CommandBars("Ribbon").Enabled = True 'єксель2013 показати
    Dim excelVersion As String
    excelVersion = Application.Version
    'ThisWorkbook.Worksheets(ws_Name).Activate
        ' Перевіряємо версію Excel та робимо відповідні дії
    If excelVersion = "15.0" Then ' Для Excel 2013
        'Application.CommandBars("Worksheet Menu Bar").Enabled = False
        'Application.CommandBars("Ribbon").Enabled = False 'єксель2013 приховати
        Application.ExecuteExcel4Macro "SHOW.TOOLBAR(""Ribbon"", False)"
    ElseIf excelVersion = "16.0" Then ' Для Excel 2016 та пізніших версій
        'Application.CommandBars("Ribbon").Enabled = True
        Application.ExecuteExcel4Macro "SHOW.TOOLBAR(""Ribbon"", False)"
    End If
    Application.DisplayFormulaBar = False
    'MsgBox excelVersion
End Sub
Public Sub Protect_ws() 'Для подвійного кліку на C1
    Call FindParam
    Call HideControlPanel
    ThisWorkbook.Worksheets(ws_Name).Protect (ws_passw), DrawingObjects:=True, Contents:=True, Scenarios:=True
End Sub
Public Sub Protect_ws_from_form() 'Зняття паролю з форми
    'Call FindParam
    Call HideControlPanel
    ThisWorkbook.Worksheets(ws_Number).Protect (ws_passw), DrawingObjects:=True, Contents:=True, Scenarios:=True
End Sub
Public Sub Protect_ws_all()
    Dim wb As Workbook
    Dim ws As Worksheet
    Set wb = ThisWorkbook    ' Вказуємо робочу книгу
    For Each ws In wb.Sheets ' Проходимося по всіх аркушах 90 у книзі
    'MsgBox ws.name
            If Left(ws.name, 2) = "98" Or Left(ws.name, 2) = "90" Or Left(ws.name, 4) = "Zvit" Then ' Якщо аркуш починається з 98,90,Zvit
                ThisWorkbook.Worksheets(ws.name).Protect ("lab123"), DrawingObjects:=True, Contents:=True, Scenarios:=True
                'MsgBox "lab123"
            Else
                If Left(ws.name, 4) = "Data" Then ThisWorkbook.Worksheets(ws.name).Protect ("lab"), DrawingObjects:=True, Contents:=True, Scenarios:=True
                'MsgBox "lab"
            End If
    Next ws
End Sub
Public Sub Unprotect_ws() 'Для подвійного кліку на А1
    Call FindParam
   'MsgBox "Ім'я активного аркуша: " & ws_Name
    'ThisWorkbook.Worksheets(ws_Name).Activate
    'MsgBox "Пароль: " & ws_passw
    ThisWorkbook.Worksheets(ws_Name).Unprotect (ws_passw)
End Sub
Public Sub Unprotect_ws_from_form() 'Зняття паролю з форми
    Call FindParam
   'MsgBox "Ім'я активного аркуша: " & ws_Name
    'ThisWorkbook.Worksheets(ws_Name).Activate
    'MsgBox "Пароль: " & ws_passw
    ThisWorkbook.Worksheets(ws_Number).Unprotect (ws_passw)
End Sub
Public Sub Unprotect_ws_all()
    Dim wb As Workbook
    Dim ws As Worksheet
    Set wb = ThisWorkbook    ' Вказуємо робочу книгу
    For Each ws In wb.Sheets ' Проходимося по всіх аркушах 90 у книзі
    Application.ScreenUpdating = False ' Вимикаємо оновлення екрану для швидкості
    'MsgBox ws.name
            If Left(ws.name, 2) = "98" Or Left(ws.name, 2) = "90" Or Left(ws.name, 4) = "Zvit" Then ' Якщо аркуш починається з 98,90,Zvit
                ThisWorkbook.Worksheets(ws.name).Unprotect ("lab123")
                'MsgBox "lab123"
            Else
                If Left(ws.name, 4) = "Data" Then ThisWorkbook.Worksheets(ws.name).Unprotect ("lab")
                'MsgBox "lab"
            End If
    Next ws
    Call ShowControlPanel
    Application.ScreenUpdating = True ' Вмикаємо оновлення екрану для швидкості
End Sub

Sub OpenControlForm() 'Форма FormGeneral98 та ControlForm

    Dim ListObj As ListObject
    Dim Count As Integer
    Call FindParam
    Set ws = ThisWorkbook.Worksheets(ws_Name)
    Set ListObj = ws.ListObjects(1)
    Count = ListObj.ListRows.Count
    'якщо рядки в ListObj відсутні то виникає помилка, тому перевірємо Count
    If ListObj.ListRows.Count = 0 Then
        If Left(ws.name, 2) = "98" Then
                ws.Cells(11, 8).Select
        ElseIf Left(ws.name, 2) = "90" Then
                ws.Cells(11, 7).Select
        End If
        MsgBox "Неможливо додати запис про перевірку" & Chr(10) & "Будь ласка, додайте спочатку запис до таблиці", vbOKOnly + vbExclamation, "Перевірка введення"
        Exit Sub
    End If
    FormControl.Show
End Sub
Sub Add_row_98() 'Форма FormGeneral98 та ControlForm
    Dim ListObj As ListObject
    Dim Count As Integer
    Call FindParam
    Set ws = ThisWorkbook.Worksheets(ws_Name)
    If Left(ws.name, 2) = "98" Then
        Form98.Show
    Else
        Form90.Show
    End If
End Sub
Sub Add_row_90_sh_3to26() 'Форма FormGeneral98 та ControlForm actShIndex
    Dim ListObj As ListObject
'    Dim s As Integer
    Call FindParam
    'Set ws = ThisWorkbook.Worksheets(ws_Name)
'    If Left(ws.name, 2) = "98" Then
'        FormGeneral98.Show
'    Else
'        FormGeneral90.Show
'    End If
   Select Case actShIndex
        Case 2
            Form90Ar.Show
        Case 3 To 26
            Form90.Show
    End Select
End Sub
Sub OpenGeneralForm() 'форма FormGeneral
    FormGeneral.Show vbModeless
End Sub
Sub ClearTablesVC9_OnAllSheets()
    ' Зняття захисту з усіх аркушів
    ' Unprotect all worksheets
    Unprotect_ws_all
    
    Dim ws As Worksheet
    Dim lo As ListObject
    
    ' Перебір аркушів
    ' Loop through each worksheet
    For Each ws In ThisWorkbook.Worksheets
        ' Перебір всіх таблиць на аркуші
        ' Loop through all tables on the worksheet
        For Each lo In ws.ListObjects
            ' Перевірка, чи назва таблиці починається з "VC9"
            ' Check if the table name starts with "VC9"
            If Left(lo.name, 3) = "VC9" Then
                ' Перевірка, чи таблиця має більше одного рядка та чи містить дані
                ' Check if the table has more than one row and contains data
                If Not lo.DataBodyRange Is Nothing Then
                    If Application.WorksheetFunction.CountA(lo.DataBodyRange) > 0 Then
                        lo.DataBodyRange.Rows.Delete
                    End If
                End If
            End If
        Next lo
    Next ws
    
    ' Встановлення захисту на всі аркуші
    ' Protect all worksheets
    Protect_ws_all
End Sub
