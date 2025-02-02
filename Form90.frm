VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Form90 
   Caption         =   "         ВЦ-90"
   ClientHeight    =   11220
   ClientLeft      =   8250
   ClientTop       =   2895
   ClientWidth     =   15480
   OleObjectBlob   =   "Form90.frx":0000
End
Attribute VB_Name = "Form90"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Write_close_Form90_Click()
    
    '---------блокування вводу якщо перевірка не виконувалась-----++//------------------------------------------------------------------
    '---------Block input if the check was not performed--------------------------------------------------------------------------
    'Dim ws As Worksheet
    Dim lastRow As Long
    Dim m As Long ' лічильник / counter
    Dim foundRow As Long
    Dim lastCheckedDate As Date
    Dim currentDate As Date
    
    '-------- Визначення аркушу та останнього рядка в стовпчику 8
    '-------- Define the worksheet and the last row in column 8
    Set ws = ThisWorkbook.Worksheets(ws_Number)
    '---------код для вц-98-----------------------------------------------
    '---------Code for VC-98-----------------------------------------------
    'If Left(ws.name, 2) = "98" Then
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    '---- Пошук останнього значення "Перевірено" в стовпці 8
    '---- Find the last "Checked" value in column 8
    If lastRow > 0 Then
        For m = lastRow To 11 Step -1
            If InStr(1, ws.Cells(m, 7).Value, "еревір") <> 0 Then
                lastCheckedDate = ws.Cells(m, 1).Value
                Exit For
            End If
        Next m
        If lastCheckedDate = 0 Then lastCheckedDate = ws.Cells(11, 1).Value
        
        '---- Порівняння дат та відображення повідомлення в залежності від різниці днів
        '---- Compare dates and display a message based on the difference in days
        If lastCheckedDate <> 0 Then
            currentDate = Date
            Dim dayDifference As Integer
            dayDifference = DateDiff("d", lastCheckedDate, currentDate)
        'MsgBox "ws_Number -  " & ws_Number & Chr(10) & "ws_Name -  " & ws_Name
        'MsgBox "tableNameVC90 - " & tableNameVC90
        'MsgBox "lastCheckedDate -  " & lastCheckedDate & Chr(10) & "dayDifference -  " & dayDifference

            Select Case dayDifference
                Case Is < 27
                    GoTo End_Select
                Case 27
                    MsgBox "Будь ласка, нагадайте керівнику," & Chr(10) & "остання перевірка виконувалась 27 діб тому.", vbOKOnly + vbInformation, "Перевірка введення"
                Case 28
                    MsgBox "Будь ласка, нагадайте керівнику," & Chr(10) & "остання перевірка виконувалась 28 діб тому.", vbOKOnly + vbInformation, "Перевірка введення"
                Case 29
                    MsgBox "Будь ласка, нагадайте керівнику," & Chr(10) & "остання перевірка виконувалась 29 діб тому.", vbOKOnly + vbInformation, "Перевірка введення"
                Case 30
                    MsgBox "Будь ласка, нагадайте керівнику," & Chr(10) & "остання перевірка виконувалась 30 діб тому." & Chr(10) & "Введення даних через 3 доби буде заблоковано.", vbOKOnly + vbExclamation, "Перевірка введення"
                Case 31
                    MsgBox "Будь ласка, нагадайте керівнику," & Chr(10) & "остання перевірка виконувалась 31 добу тому." & Chr(10) & "Введення даних через 2 доби буде заблоковано.", vbOKOnly + vbExclamation, "Перевірка введення"
                Case 32
                    MsgBox "Будь ласка, нагадайте керівнику," & Chr(10) & "остання перевірка виконувалась 32 доби тому." & Chr(10) & "Введення даних через 1 добу буде заблоковано.", vbOKOnly + vbExclamation, "Перевірка введення"
                Case Else
                    '-------- Відображення повідомлення, якщо різниця днів не відповідає жодному з умов
                    '-------- Display a message if the day difference does not match any condition
                    MsgBox "Кількість діб без перевірки: " & dayDifference, vbOKOnly + vbCritical, "Перевірка введення"
                    MsgBox "Будь ласка, передайте керівнику привіт :)" & Chr(10) & "Введення даних заблоковано." & Chr(10) & "Відновлення роботи можливе лише після перевірки.", vbOKOnly + vbCritical, "Перевірка введення"
                    FormGeneral.Controls("Label_choice_" & ws_Number).Caption = "Заблоковано"
                    FormGeneral.Controls("Label_choice_" & ws_Number).ForeColor = RGB(255, 0, 0) ' RGB(темноЧЕРВОНИЙ, зелений, синій)
                    Unload Form90
                    Exit Sub
End_Select: End Select
        End If
    End If
    
    'Перевірка введення прізвища 00000000000000000000000000000000000000000000000000000000000000000000
    'Surname input validation 00000000000000000000000000000000000000000000000000000000000000000000
    Dim inputText As String
    Dim i As Integer
    Dim hasDigits As Boolean
    
    inputText = ComboBox_name.Text
    hasDigits = False
    
    ' ----------Перевіряємо кожен символ у введеному тексті
    ' ----------Check each character in the input text
    For i = 1 To Len(inputText)
        If IsNumeric(Mid(inputText, i, 1)) Then
            hasDigits = True
            Exit For
        End If
    Next i
    
    ' ---------Виводимо повідомлення, якщо цифри знайдені
    ' ---------Display a message if digits are found
    If hasDigits Then
        MsgBox "Будь ласка, видаліть цифри з прізвища!", vbOKOnly + vbExclamation, "Перевірка введення"
        GoTo name
    End If
    
    If ComboBox_name.Value = Empty Then
        MsgBox "Будь ласка, введіть прізвище виконавця", vbOKOnly + vbExclamation, "Перевірка введення"
        GoTo name
    End If
    GoTo status

name:
    ComboBox_name.SetFocus
    ComboBox_name.SelStart = 0
    ComboBox_name.SelLength = Len(inputText)
    Exit Sub

status:
    'Перевірка введення робочого стану 00000000000000000000000000000000000000000000000000000000000000000000
    'Work status input validation 00000000000000000000000000000000000000000000000000000000000000000000
    
    If ComboBox_status.Value = Empty Then
        GoTo next1
    ElseIf ComboBox_status.Value = "Робочий" Or ComboBox_status.Value = "Неробочий" Then
        GoTo next2
    End If

next1:
    MsgBox "Будь ласка, введіть робочий стан", vbOKOnly + vbExclamation, "Перевірка введення"
    ComboBox_status.SetFocus
    ComboBox_status.SelStart = 0
    ComboBox_status.SelLength = Len(inputText)
    Exit Sub

next2:
    ' Визначити ім'я таблиці на вказаному аркуші00000000000000000000000000000000000000000000000000000000000000000000
    ' Define the table name on the specified worksheet00000000000000000000000000000000000000000000000000000000000000000000
    Dim Таблиця As String 'ім'я Таблиці / Table name
    Таблиця = "VC90_tab_" & CStr(ws_Number)
    
    ' Отримати обрані CheckBox
    ' Get selected CheckBox
    Dim обраніЧекБокси As String
    Dim n As Integer ' лічильник / counter
    For n = 1 To 20
        ' Перевірити, чи обраний CheckBox
        ' Check if CheckBox is selected
        If Controls("CheckBox" & n).Value = True Then
            ' Якщо обрано, додати текст Label до рядка
            ' If selected, add Label text to the string
            обраніЧекБокси = обраніЧекБокси & Controls("Label_Item_" & n).Caption & "; "
        End If
    Next n
    
    ' Видалити останню кому та пробіл якщо CheckBox обрані, якщо ні то "-"
    ' Remove the last comma and space if CheckBoxes are selected, otherwise "-"
    If Len(обраніЧекБокси) <> 0 Then
        обраніЧекБокси = Left(обраніЧекБокси, Len(обраніЧекБокси) - 1)
    Else
        обраніЧекБокси = "-"
    End If
    
    ' Перевірка
    ' Validation
    If обраніЧекБокси = "-" And TextBox_solution = "-" And TextBox_fault = "-" Then
        MsgBox "Будь ласка, оберіть пункти ТО," & Chr(10) & "або введіть Несправність," & Chr(10) & "або введіть Спосіб усунення.", vbOKOnly + vbExclamation, "Перевірка введення"
        Exit Sub
    End If
    
    ' Перевірка Несправність
    ' Fault input validation
    If TextBox_fault = Empty Then
        MsgBox "Будь ласка, введіть Несправність." & Chr(10) & "Несправність це текст, або "" - """, vbOKOnly + vbExclamation, "Перевірка введення"
        Exit Sub
    End If
    
    ' Перевірка Спосіб усунення
    ' Solution input validation
    If TextBox_solution = Empty Then
        MsgBox "Будь ласка, введіть Спосіб усунення." & Chr(10) & "Спосіб усунення це текст, або "" - """, vbOKOnly + vbExclamation, "Перевірка введення"
        Exit Sub
    End If
    
    ' Додати новий рядок та записати дані
    ' Add a new row and write data
    Call Unprotect_ws_from_form
    Dim ListObj As ListObject
    Dim ListRow As ListRow
    Set ws = ThisWorkbook.Worksheets(ws_Number)
    Set ListObj = ws.ListObjects(Таблиця)
    Set ListRow = ListObj.ListRows.Add
    
    ListRow.Range(1).NumberFormat = "dd.mm.yyyy;@"
    ListRow.Range(1) = Date
    ListRow.Range(2).NumberFormat = "@" ' Встановлюємо формат як текст / Set format as text
    ListRow.Range(2).HorizontalAlignment = xlLeft ' Вирівнюємо текст по лівому краю / Align text to the left
    ListRow.Range(2).WrapText = True ' Встановлюємо властивість WrapText / Set WrapText property
    ListRow.Range(2) = CStr(обраніЧекБокси) ' Записуємо значення / Record value
    ListRow.Range(3) = CStr(Form90.TextBox_fault.Value)
    ListRow.Range(4) = CStr(Form90.TextBox_solution.Value)
    ListRow.Range(5) = CStr(Form90.ComboBox_status.Value)
    ListRow.Range(6) = CStr(Form90.ComboBox_name.Value)
    ListRow.Range(7) = CStr("-")
    Call Protect_ws_from_form
    
    ' В FormGeneral в Label_choice_x відмітити що Form90 запис додала
    ' In FormGeneral, mark Label_choice_x that Form90 added the record
    FormGeneral.Controls("Label_choice_" & ws_Number).Caption = "Запис додано"
    FormGeneral.Controls("Label_choice_" & ws_Number).ForeColor = RGB(0, 128, 0) ' RGB(червоний, темноЗЕЛЕНИЙ, синій / red, darkGREEN, blue)
    Unload Form90
    
    ws.Activate
    Dim Count As Integer
    Count = ListObj.ListRows.Count
    ListObj.DataBodyRange(Count, 1).Select
End Sub


Private Sub UserForm_Initialize()
    
    Dim i As Integer
    Dim LabelItem As Object
    Dim LabelOperation As Object
    Dim LabelPeriodicity As Object
    Dim LabelResponsible As Object
    Dim CheckBox As Object
    Dim LabelValidity As Object
    Dim LabelDayDiffTO As Object
    Set wsData = ThisWorkbook.Worksheets("Data") ' Вказуємо аркуш "Data".
    
    ' Встановлюємо назву форми
    Form90.Caption = "         ВЦ-90     " & wsData.Range("A2").Value
    
    ' Перевірка значення вибраного OptionButton або Label'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Select Case ws_Number
        Case 1 To 25
            ' Вибір відповідної розумної таблиці
            Dim tableName As String 'Змінна в якій ім'я таблиці ТО_х на аркуші Дані
            tableName = "TO_" & ws_Number
            
        '----при виборі обладнання активувати комірку першого стовпчика станнього рядка
            Dim ListObj As ListObject
            Dim tableNameVC90 As String 'Змінна в якій ім'я таблиці VC90_tab_ на аркуші обранного обладнання
            tableNameVC90 = "VC90_tab_" & ws_Number
            'MsgBox "ws_Number - " & ws_Number
            'MsgBox "tableNameVC90 - " & tableNameVC90
            Set ws = ThisWorkbook.Worksheets(ws_Number)
            Set ListObj = ws.ListObjects(tableNameVC90)
            Dim Count As Integer 'Змінна в якій кількість рядків
            Count = ws.ListObjects(tableNameVC90).ListRows.Count
            Set ws = ThisWorkbook.Worksheets(ws_Number)
            'MsgBox "Count - " & Count
            ws.Activate
            ' Якщо є дані в таблиці, виділяємо комірку першого стовпця останнього рядка
            If Count > 0 Then
                ListObj.DataBodyRange(Count, 1).Select
            Else
                ' Якщо таблиця порожня, виділяємо комірку A10
                ws.Range("A10").Select
                'MsgBox "Таблиця порожня, вибрано комірку A10"
            End If
        '----
            ' Перевірка наявності розумної таблиці
'            If TableExists(wsData, tableName) Then
                ' Визначення кількості рядків в розумній таблиці
                Dim rowCount As Integer
                rowCount = wsData.ListObjects(tableName).ListRows.Count
                ' Якщо розумна таблиця порожня, замінити значення для всіх Label_Item
                'MsgBox "Кількість рядків таблиці ТО на аркуші Дані- " & rowCount
                If rowCount = 0 Then
                    For i = 1 To 20
                        Set LabelItem = Me.Controls("Label_Item_" & i)
                        LabelItem.Caption = "Таблиця " & tableName & " порожня"
                        LabelItem.Visible = True
                    Next i
                Else
                    ' Цикл для ітерації по Label_Item
                    For i = 1 To 20
                        ' Знаходження відповідного Label
                        Set LabelItem = Me.Controls("Label_Item_" & i)
                        LabelItem.Visible = True                        ' Зробити видимим Label
                        Set LabelOperation = Me.Controls("Label_Operation_" & i)
                        LabelOperation.Visible = True                    ' Зробити видимим Label
                        Set LabelPeriodicity = Me.Controls("Label_Periodicity_" & i)
                        LabelPeriodicity.Visible = True                   ' Зробити видимим Label
                        Set LabelResponsible = Me.Controls("Label_Responsible_" & i)
                        LabelResponsible.Visible = True                   ' Зробити видимим Label
                        Set LabelDayDiffTO = Me.Controls("Label_DayDiffTO_" & i)
                        LabelDayDiffTO.Visible = True                   ' Зробити видимим Label
                    ' Перевірка, чи існує рядок в розумній таблиці'''''''''''''''''''''''''''''''''''''''''''''''''''''
                        If i <= rowCount Then
                            ' Отримання значення з розумної таблиці
                            LabelItem.Caption = wsData.ListObjects(tableName).DataBodyRange.Cells(i, 1).Text
                            LabelOperation.Caption = wsData.ListObjects(tableName).DataBodyRange.Cells(i, 2).Text
                            LabelPeriodicity.Caption = wsData.ListObjects(tableName).DataBodyRange.Cells(i, 3).Text
                            LabelResponsible.Caption = wsData.ListObjects(tableName).DataBodyRange.Cells(i, 4).Text
                            LabelDayDiffTO.Caption = wsData.ListObjects(tableName).DataBodyRange.Cells(i, 8).Value
                        End If
                    Next i
                    ' Приховати Label_Item, номер якого більше ніж rowCount''''''''''''''''''''''''''''''''''''''''''''''
                    For i = rowCount + 1 To 20
                        Set LabelItem = Me.Controls("Label_Item_" & i)
                        LabelItem.Visible = False                        ' Зробити не видимим Label
                        Set LabelOperation = Me.Controls("Label_Operation_" & i)
                        LabelOperation.Visible = False                     ' Зробити не видимим Label
                        Set LabelPeriodicity = Me.Controls("Label_Periodicity_" & i)
                        LabelPeriodicity.Visible = False                    ' Зробити невидимим Label
                        Set LabelResponsible = Me.Controls("Label_Responsible_" & i)
                        LabelResponsible.Visible = False                    ' Зробити невидимим Label
                        Set CheckBox = Me.Controls("CheckBox" & i)
                        CheckBox.Visible = False                    ' Зробити невидимим OptionButton
                        Set LabelValidity = Me.Controls("Label_validity_" & i)
                        LabelValidity.Visible = False                    ' Зробити невидимим Label
                        Set LabelDayDiffTO = Me.Controls("Label_DayDiffTO_" & i)
                        LabelDayDiffTO.Visible = False                    ' Зробити невидимим Label
                    Next i
                End If
    End Select
    ' Внесення дати останнього виконання пункту ТО в Label_Validity_х '''''''''''''''''''''''''''''''''''''''''''''''''''''
    Dim searchValue As String 'Змінна в якій пункт ТО при його пошуку в таблицях VC90_tab_х
    Dim foundDate As Date 'Змінна в якій знайдена дата останнього виконання пункту ТО
    Dim tableRange As Range 'Змінна в якій  діапазон VC90_tab_х
    Dim labelIndex As Integer 'Змінна для номера Label_Item_x для вибору потрыбного пункту ТО
    Dim rowIndex As Integer 'Змінна для перебору рядків в таблиці при пошуку останньої дати виконання пунктуТО
    Dim found As Boolean 'Змінна яка визначає чи знайдена дата останнього виконная ТО    ' Задайте значення глобальних змінних tableName та ws_Number
    ' Отримати посилання на робочий аркуш
    Set ws = ThisWorkbook.Worksheets(ws_Number)
    ' Отримати діапазон таблиці VC90_tab_X
    Set tableRange = ws.ListObjects("VC90_tab_" & ws_Number).Range
    ' Перебір усіх LabelItem
    For labelIndex = 1 To 20
        ' Отримати посилання на LabelItem_X і LabelValidity_X
        Set LabelItem = Me.Controls("Label_Item_" & labelIndex)
        Set LabelValidity = Me.Controls("Label_Validity_" & labelIndex)
        ' Отримати значення з LabelItem_X
        searchValue = LabelItem.Caption
        ' Скинути прапорець знаходження
        found = False
        ' Перебір рядків у зворотньому порядку
        For rowIndex = tableRange.Rows.Count To 1 Step -1
            ' Перевірити, чи значення з таблиці містить шукане значення
            If InStr(1, tableRange.Cells(rowIndex, 2).Value, searchValue) > 0 Then
                ' Знайдено відповідне значення, отримати дату з стовпчика 1
                foundDate = tableRange.Cells(rowIndex, 1).Value

                ' Записати отриману дату в LabelValidity_X
                LabelValidity.Caption = Format(foundDate, "dd.mm.yyyy")
                'MsgBox LabelValidity.Caption
                Dim DayDiff_TO As Integer ' Змінна допустима кількість днів між датою та датою остннього виконання пункту ТО
                DayDiff_TO = DateDiff("d", LabelValidity, Date)
                
                'DayDiff_TO = Date - LabelValidity 'wsData.ListObjects(tableName).DataBodyRange.Cells(i, 8).Value
                If DayDiff_TO > wsData.ListObjects(tableName).DataBodyRange.Cells(labelIndex, 8).Value Then
                    LabelValidity.ForeColor = RGB(128, 0, 0) ' Встановлення колір для LabelValidity у випадку відсутності ТО
                    LabelValidity.Font.Bold = True
                End If
                ' Помітити, що знайдено значення
                found = True 'Дата останнього виконная ТО знайдена
                Exit For
            End If
        Next rowIndex
        ' Якщо значення пункт ТО не знайдено в таблиці, записати "Немає даних"
        If Not found Then
            foundDate = tableRange.Cells(2, 1).Value
            If foundDate = 0 Then
                    LabelValidity.Caption = "Немає даних"
                    'LabelValidity.ForeColor = RGB(0, 0, 128) ' RGB(червоний, зелений, темноСИНІЙ)
                    LabelValidity.Font.Bold = True
            Else
                    DayDiff_TO = DateDiff("d", foundDate, Date)
                    If DayDiff_TO > wsData.ListObjects(tableName).DataBodyRange.Cells(labelIndex, 8).Value Then
                        LabelValidity.Caption = "Немає даних"
                        LabelValidity.ForeColor = RGB(128, 0, 0) ' RGB(темноЧЕРВОНИЙ, зелений, темносиній)
                        LabelValidity.Font.Bold = True
                    Else
                        LabelValidity.Caption = "Немає даних"
                        'LabelValidity.ForeColor = RGB(0, 0, 128) ' RGB(червоний, зелений, темноСИНІЙ)
                        LabelValidity.Font.Bold = True
                    End If
            End If
        End If
'''''''''''''''''''''''''

    Next labelIndex
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Label_date_90.Caption = Date
    ' Зміна розмір шрифта, жирний та назви форми в залежності обраного обладнання ws_Number
    Label_equipment.Caption = wsData.ListObjects(tableName).DataBodyRange.Cells(1, 6).Text
    Label_equipment.Font.Size = 9
    Label_equipment.Font.Bold = True
    TextBox_solution.Value = "-"
    TextBox_fault.Value = "-"
    ' Зміна розміру форми в залежності від rowCount
    Me.Height = 615 - (20 - rowCount) * 25 ' Змініть 20 на бажаний висоту для кожного нового рядка
    
End Sub

' Перевірка існування таблиці в заданому аркуші
Function TableExists(ws As Worksheet, tableName As String) As Boolean
    On Error Resume Next
    TableExists = Not ws.ListObjects(tableName) Is Nothing
    On Error GoTo 0
End Function
