VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ЭтаКнига"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Workbook_Open()
    'Dim ws As Worksheet
    Dim ListObj As ListObject
    Dim ListRow As ListRow
    Dim Count As Integer

    Set ws = ThisWorkbook.Sheets(1)
    Set ListObj = ws.ListObjects(1)
    ws.Activate
    Count = ListObj.ListRows.Count
    If Count = 0 Then
            ws.Cells(11, 1).Select
        Else
            ws.Cells(Count, 1).Select
            ListObj.DataBodyRange(Count, 1).Select
    End If
    Call Protect_ws_all
End Sub

Private Sub Workbook_SheetDeactivate(ByVal sh As Object)
    Dim ws As Worksheet
    Dim checkRange As Range
    Dim n As Integer
    Dim emptyCellAddress As String ' Оголошення змінної для адреси порожньої комірки
    Set ws = ThisWorkbook.Sheets("Data")
    '**************Перевіряємо, чи заповнені комірки Використання аркушу на аркуш "Data"
    If sh.name = ws.name Then
        For n = 2 To 17
            If IsEmpty(ws.Cells(3, n).Value) Then
                emptyCellAddress = ws.Cells(3, n).Address ' Отримання адреси порожньої комірки
                MsgBox "Будь ласка," & Chr(10) & "Введіть Використання аркушу у комірку " & emptyCellAddress & Chr(10) & "на аркуші 'Дані'.", vbExclamation, "Перевірка введення"
                ws.Activate
                ws.Cells(3, n).Select ' Вибір порожньої комірки
                Exit Sub
            End If
        Next n
        '**************Перевіряємо, чи заповнені комірки Підрозділ на аркуші "Data"
        Set checkRange = ws.Range("B4:Z4")
        If checkRange.MergeCells Then
            If checkRange.Cells(1, 1).Value = "" Then
                MsgBox "Будь ласка," & Chr(10) & "Введіть Підрозділ у комірку B4:Z4 на аркуші 'Дані'.", vbExclamation, "Перевірка введення"
                ' встановлюємо фокус на комірку B4:Z4
                ws.Activate
                ws.Range("B4").Select
                Exit Sub
            End If
        End If

        If Not checkRange.MergeCells Then
            MsgBox "Будь ласка," & Chr(10) & "Введіть Підрозділ у комірку B4:Z4 на аркуші 'Дані'." & Chr(10) & "Комірка B4:Z4 має бути об'єднаною", vbExclamation, "Перевірка введення"
            ' встановлюємо фокус на комірку B4
            ws.Activate
            ws.Range("B4").Select
            Exit Sub
        End If
        '**************Перевіряємо, чи заповнені В назві аркушу  на аркуші "Data"
        For n = 2 To 17
            If ws.Cells(3, n).Value = "Yes" Then
                If IsEmpty(ws.Cells(5, n).Value) Then
                    emptyCellAddress = ws.Cells(5, n).Address ' Отримання адреси порожньої комірки
                    MsgBox "Будь ласка," & Chr(10) & "Введіть В назві аркушу у комірку " & emptyCellAddress & Chr(10) & "на аркуші 'Дані'.", vbExclamation, "Перевірка введення"
                    ws.Activate
                    ws.Cells(5, n).Select ' Вибір порожньої комірки
                    Exit Sub
                End If
            End If
        Next n
        '**************Перевіряємо, чи заповнені Назва в журналі  на аркуші "Data"
        For n = 2 To 17
            If ws.Cells(3, n).Value = "Yes" Then
                If IsEmpty(ws.Cells(6, n).Value) Then
                    emptyCellAddress = ws.Cells(6, n).Address ' Отримання адреси порожньої комірки
                    MsgBox "Будь ласка," & Chr(10) & "Введіть Назва в журналі у комірку " & emptyCellAddress & Chr(10) & "на аркуші 'Дані'.", vbExclamation, "Перевірка введення"
                    ws.Activate
                    ws.Cells(6, n).Select ' Вибір порожньої комірки
                    Exit Sub
                End If
            End If
        Next n
        '**************Перевіряємо, чи заповнені комірки Початок журналу на аркуш "Data"
        Set checkRange = ws.Range("B7:Z7")
        If checkRange.MergeCells Then
            If checkRange.Cells(1, 1).Value = "" Then
                MsgBox "Будь ласка," & Chr(10) & "Введіть Початок журналу у комірку B7:Z7 на аркуші 'Дані'.", vbExclamation, "Перевірка введення"
                ' встановлюємо фокус на комірку B7:Z7
                ws.Activate
                ws.Range("B7").Select
                Exit Sub
            End If
        End If

        If Not checkRange.MergeCells Then
            MsgBox "Будь ласка," & Chr(10) & "Введіть Початок журналу у комірку B7:Z7 на аркуші 'Дані'." & Chr(10) & "Комірка B7:Z7 має бути об'єднаною", vbExclamation, "Перевірка введення"
            ' встановлюємо фокус на комірку B7
            ws.Activate
            ws.Range("B7").Select
            Exit Sub
        End If
        '**************Перевіряємо, чи заповнена Кількість перевірок на добу
        Set checkRange = ws.Range("B9")
            ' Перевірка, чи комірка порожня
            If ws.Cells(3, 2).Value = "Yes" Then
                If IsEmpty(checkRange.Value) Then
                    MsgBox "Будь ласка," & Chr(10) & "Введіть Кількість перевірок на добу у комірку $B$9" & Chr(10) & "на аркуші 'Дані'.", vbExclamation, "Перевірка введення"
                    ' встановлюємо фокус на комірку B9
                    ws.Activate
                    checkRange.Select
                    Exit Sub
                End If
            End If
        '**************Перевіряємо, чи заповнений Тиск
        Set checkRange = ws.Range("B10")
        ' Перевірка, чи комірка порожня
            If ws.Cells(3, 2).Value = "Yes" Then
                If IsEmpty(checkRange.Value) Then
                    MsgBox "Будь ласка," & Chr(10) & "Введіть Кількість перевірок на добу у комірку $B$10" & Chr(10) & "на аркуші 'Дані'.", vbExclamation, "Перевірка введення"
                    ' встановлюємо фокус на комірку B9
                    ws.Activate
                    checkRange.Select
                    Exit Sub
                End If
            End If
    End If
    Call Add_text_to_name_sheet_90_98
End Sub
Sub Add_text_to_name_sheet_90_98()
    Dim wb As Workbook
    Dim wsData As Worksheet
    Dim ws As Worksheet
    Dim i As Integer
    Dim workbookPassword  As String
    workbookPassword = "lab123"    ' Задаємо пароль книги
    Set wb = ThisWorkbook    ' Вказуємо робочу книгу
    On Error Resume Next    ' Знімаємо захист паролем
    wb.Unprotect workbookPassword
    Set wsData = wb.Sheets("Data")    ' Вказуємо аркуш "Data"
    Application.ScreenUpdating = False ' Вимикаємо оновлення екрану для швидкості
    i = 1
    For Each ws In wb.Sheets ' Проходимося по всіх аркушах 90 у книзі
            i = i + 1
            If Left(ws.name, 2) = "90" Then ' Якщо аркуш починається з "90"
                If i <= wsData.Range("B5:Z5").Columns.Count + 1 Then ' Перевірка, чи існують дані в комірках B5:Z5 для цього аркуша
                    Dim newSheetName As String
                    newSheetName = "90_" & wsData.Cells(5, i).Value
                    ' Перевірка на існування аркуша з новою назвою
                    On Error Resume Next
                    If Not WorksheetExists(newSheetName) Then
                        ws.name = newSheetName
                    End If
                    On Error GoTo 0
                        ' якщо використання аркуша 'Так' аркуш видимий, інакше приховуємо
                        If wsData.Cells(3, i).Value = "Yes" Then
                            ThisWorkbook.Sheets(newSheetName).Visible = True
                            'MsgBox newSheetName & " .Visible = True"
                        Else
                            ThisWorkbook.Sheets(newSheetName).Visible = False
                            'MsgBox newSheetName & " .Visible = False"
                        End If
                 End If
            End If
            If Left(ws.name, 2) = "98" Then ' Якщо аркуш починається з "98"
                If i <= wsData.Range("B5:Z5").Columns.Count + 1 Then ' Перевірка, чи існують дані в комірках B5:Z5 для цього аркуша
                    'Dim newSheetName As String
                    newSheetName = "98_" & wsData.Cells(5, i).Value
                    ' Перевірка на існування аркуша з новою назвою
                    On Error Resume Next
                    If Not WorksheetExists(newSheetName) Then
                        ws.name = newSheetName
                    End If
                    On Error GoTo 0
                        ' якщо використання аркуша 'Так' аркуш видимий, інакше приховуємо
                        If wsData.Cells(3, i).Value = "Yes" Then
                            ThisWorkbook.Sheets(newSheetName).Visible = True
                            'MsgBox newSheetName & " .Visible = True"
                        Else
                            ThisWorkbook.Sheets(newSheetName).Visible = False
                            'MsgBox newSheetName & " .Visible = False"
                        End If
                 End If
            End If
    Next ws
    ' Встановлюємо захист паролем
    'MsgBox "временно отключен пароль на книгу"
    wb.Protect password:=workbookPassword
    Application.ScreenUpdating = True ' Увімкнути оновлення екрану назад
End Sub

Function WorksheetExists(shtName As String) As Boolean
    On Error Resume Next
    WorksheetExists = Not Worksheets(shtName) Is Nothing
    On Error GoTo 0
End Function
