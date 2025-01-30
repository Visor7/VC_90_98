VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Лист10"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Worksheet_BeforeDoubleClick(ByVal Target As Range, Cancel As Boolean)
    'Pозблокувати, показати Аркуш.
    Set ws = ActiveSheet ' Отримуємо активний аркуш
    ws_Name = ws.name
    'MsgBox "Ім'я активного аркуша: " & sheet_Name
    'Перевіряємо, чи це подія на комірці A1
    If Target.Address = "$E$1" Then
        ' Отримання введеного значення від користувача через InputBox
        Dim pasw As String
        pasw = InputBox("Password:")
        ' Перевірка введеного пароля та виведення відповідного повідомлення
        If pasw = "lab123" Then
            Call Unprotect_ws
            Call ShowControlPanel
            ThisWorkbook.Worksheets(ws_Name).Range("A7").Select
        ElseIf pasw = "" Then
            ThisWorkbook.Worksheets(ws_Name).Range("A7").Select
            Exit Sub
        Else
            MsgBox "Please enter the correct password" & Chr(10) & "and try again", vbOKOnly + vbInformation
            ThisWorkbook.Worksheets(ws_Name).Range("A7").Select
            Exit Sub
        End If
    ElseIf Target.Address = "$G$1" Then 'Захищаємо аркуш
        Call Protect_ws
        ThisWorkbook.Worksheets(ws_Name).Range("C7").Select
    ElseIf Target.Address = "$A$1" Then 'Захищаємо аркуш
        Call FormGeneral.Show
    End If
End Sub

