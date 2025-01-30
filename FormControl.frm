VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FormControl 
   Caption         =   "Перевірка"
   ClientHeight    =   4080
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5790
   OleObjectBlob   =   "FormControl.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FormControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Cbx_name_contr_Change()
    Txb_passw_contr.SetFocus
End Sub

Private Sub Passw_exit_contr_Click()
    Unload FormControl
End Sub

Private Sub Txb_passw_contr_Change()
    Dim ListObj As ListObject
    Dim ListColumn As ListColumn
    Dim Count As Integer
    'Call FindParam
    Set ws = ThisWorkbook.Worksheets(ws_Name)
    Set ListObj = ws.ListObjects(1)
    Dim inputText As String
    Dim i As Integer
    Dim hasDigits As Boolean
    
    If Txb_passw_contr.Value = "lab" Then
        ' -------------------------перевірка комбо прізвище-------------------------
        inputText = Cbx_name_contr.Text
        hasDigits = False
        ' Перевіряємо кожен символ у введеному тексті, цифра чи ні
        For i = 1 To Len(inputText)
            If IsNumeric(Mid(inputText, i, 1)) Then
                hasDigits = True
                Exit For
            End If
        Next i
    
        ' Виводимо повідомлення, якщо цифри знайдені
        If hasDigits Then
            MsgBox "Будь ласка, видаліть цифри з прізвища!", vbOKOnly + vbExclamation, "Перевірка введення"
            GoTo name
        End If
         ' Перевіряємо хоть щось ведено
        If Cbx_name_contr.Value = Empty Then
            MsgBox "Будь ласка, введіть  прізвище," & Chr(10) & "та спробуйте ще раз", vbOKOnly + vbExclamation, "Перевірка введення"
            Txb_passw_contr.Value = ""
            GoTo name
        End If
        Call Unprotect_ws

        'пишемо перевірку в таблицю
        Dim ListRow As ListRow
        Set ListRow = ListObj.ListRows.Add
        
        ListRow.Range(1).NumberFormat = "dd.mm.yyyy;@"
        ListRow.Range(1) = Date
        ListRow.Range(2).NumberFormat = "@" ' Встановлюємо формат як текст / Set format as text
        ListRow.Range(2).HorizontalAlignment = xlLeft ' Вирівнюємо текст по лівому краю / Align text to the left
        ListRow.Range(2).WrapText = True ' Встановлюємо властивість WrapText / Set WrapText property
        ListRow.Range(2) = CStr("-") ' Записуємо значення / Record value
        ListRow.Range(3) = CStr("-")
        ListRow.Range(4) = CStr("-")
        ListRow.Range(5) = CStr("-")
        ListRow.Range(6) = CStr("-")
        
        If Left(ws.name, 2) = "98" Then
                ListRow.Range(7) = CStr("-")
                ListRow.Range(8) = Cbx_name_contr.Value & " Перевірено " & Label_date_contr & " " & Format(Now, "Short Time")
        ElseIf Left(ws.name, 2) = "90" Then
            ListRow.Range(7) = Cbx_name_contr.Value & " Перевірено " & Label_date_contr & " " & Format(Now, "Short Time")
        End If
        
        Unload FormControl
        Call Protect_ws
        
        If Left(ws.name, 2) = "98" Then
            ListObj.DataBodyRange(Count, 8).Select
        ElseIf Left(ws.name, 2) = "90" Then
            ListObj.DataBodyRange(Count, 7).Select
        End If
        
        ThisWorkbook.Save
        Exit Sub
    End If
        
    If (Len(Txb_passw_contr.Value)) < 3 Then
        Exit Sub
    End If
    MsgBox "Будь ласка, введіть правильний пароль" & Chr(10) & "та спробуйте ще раз", vbOKOnly + vbExclamation, "Перевірка введення"
    GoTo passw
    Exit Sub
name:   Cbx_name_contr.SetFocus
        Cbx_name_contr.SelStart = 0
        Cbx_name_contr.SelLength = Len(inputText)
        Txb_passw_contr.Value = ""
    Exit Sub
passw:  Txb_passw_contr.Value = ""
        Txb_passw_contr.SetFocus
    Exit Sub
End Sub

Private Sub UserForm_Initialize()
    Call FindParam
    Label_date_contr.Caption = Date
    Cbx_name_contr.Value = ""
    Txb_passw_contr.Value = ""
End Sub


