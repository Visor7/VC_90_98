VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Form98 
   Caption         =   "         ¬÷-98"
   ClientHeight    =   3975
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8025
   OleObjectBlob   =   "Form98.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Form98"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Close_General_Click() '--------- нопка вийти----------------------
    Unload Form98
End Sub

Private Sub Close_Save_General_Click() '---- нопка зберегти та вийти----------

    '---------блокуванн¤ вводу ¤кщо перев≥рка не виконувалась-----------------------------------------------------------------------
    'Dim ws As Worksheet
    Dim lastRow As Long
    Dim n As Long
    Dim foundRow As Long
    Dim lastCheckedDate As Date
    Dim currentDate As Date
'******************
    Set ws = ActiveSheet ' ќтримуЇмо активний аркуш
    ' ќтримуЇмо ≥м'¤ активного аркуша
    ''Dim ws_Name As String
    'ws_Name = ws.name
    'MsgBox "≤м'¤ активного аркуша:" & ws_Name

    '-------- ¬изначенн¤ аркушу та останнього р¤дка в стовпчику 8
    Set ws = ThisWorkbook.Worksheets(ws_Name)
    
    ' ¬становлюЇмо назву форми
    Form98.Caption = "         ¬÷-98     " & ThisWorkbook.Worksheets("Data").Range("A2").Value
    
    '---------код дл¤ вц-98-----------------------------------------------
    If Left(ws.name, 2) = "98" Then
        lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
            '---- ѕошук останнього значенн¤ "ѕерев≥рено" в стовпц≥ 8
        If lastRow > 0 Then
    
            For n = lastRow To 11 Step -1
                If InStr(1, ws.Cells(n, 8).Value, "ерев≥р") <> 0 Then
                    lastCheckedDate = ws.Cells(n, 1).Value
                    Exit For
                End If
            Next n
            If lastCheckedDate = 0 Then lastCheckedDate = ws.Cells(11, 1).Value
            '---- ѕор≥вн¤нн¤ дат та в≥дображенн¤ пов≥домленн¤ в залежност≥ в≥д р≥зниц≥ дн≥в
            If lastCheckedDate <> 0 Then
                currentDate = Date
                Dim dayDifference As Integer
                dayDifference = DateDiff("d", lastCheckedDate, currentDate)
    
                Select Case dayDifference
                    Case Is < 27
                        GoTo End_Select
                    Case 27
                        MsgBox "Ѕудь ласка, нагадайте кер≥внику," & Chr(10) & "останн¤ перев≥рка виконувалась 27 д≥б тому.", vbOKOnly + vbInformation, "ѕерев≥рка введенн¤"
                    Case 28
                        MsgBox "Ѕудь ласка, нагадайте кер≥внику," & Chr(10) & "останн¤ перев≥рка виконувалась 28 д≥б тому.", vbOKOnly + vbInformation, "ѕерев≥рка введенн¤"
                    Case 29
                        MsgBox "Ѕудь ласка, нагадайте кер≥внику," & Chr(10) & "останн¤ перев≥рка виконувалась 29 д≥б тому.", vbOKOnly + vbInformation, "ѕерев≥рка введенн¤"
                    Case 30
                        MsgBox "Ѕудь ласка, нагадайте кер≥внику," & Chr(10) & "останн¤ перев≥рка виконувалась 30 д≥б тому." & Chr(10) & "¬веденн¤ даних через 3 доби буде заблоковано.", vbOKOnly + vbExclamation, "ѕерев≥рка введенн¤"
                    Case 31
                        MsgBox "Ѕудь ласка, нагадайте кер≥внику," & Chr(10) & "останн¤ перев≥рка виконувалась 31 добу тому." & Chr(10) & "¬веденн¤ даних через 2 доби буде заблоковано.", vbOKOnly + vbExclamation, "ѕерев≥рка введенн¤"
                    Case 32
                        MsgBox "Ѕудь ласка, нагадайте кер≥внику," & Chr(10) & "останн¤ перев≥рка виконувалась 32 доби тому." & Chr(10) & "¬веденн¤ даних через 1 добу буде заблоковано.", vbOKOnly + vbExclamation, "ѕерев≥рка введенн¤"
                    Case Else
            '-------- ¬≥дображенн¤ пов≥домленн¤, ¤кщо р≥зниц¤ дн≥в не в≥дпов≥даЇ жодному з умов
                        MsgBox " ≥льк≥сть д≥б без перев≥рки: " & dayDifference, vbOKOnly + vbCritical, "ѕерев≥рка введенн¤"
                        MsgBox "Ѕудь ласка, передайте кер≥внику прив≥т :)" & Chr(10) & "¬веденн¤ даних заблоковано." & Chr(10) & "¬≥дновленн¤ роботи можливе лише п≥сл¤ перев≥рки.", vbOKOnly + vbCritical, "ѕерев≥рка введенн¤"
                        Me.Cbx_name.Value = "„ас"
                        Me.Txb_temperature.Value = "пити"
                        Me.Txb_humidity.Value = "каву!"
                        Me.Txb_pressure.Value = ChrW(&H263A) ' —майлик (O)
                    Exit Sub
End_Select:     End Select
            End If
        End If
    
        'Dim ShGeneral As Worksheet
        Dim ListObj As ListObject
        Dim ListRow As ListRow
        Dim Count As Integer
        'Set ShGeneral = ThisWorkbook.Worksheets("¬÷-98")
    
        Set ListObj = ws.ListObjects(1)
        '------------ѕерев≥рка введенн¤ пр≥звища-----------------------------
    
        Dim inputText As String
        Dim i As Integer
        Dim hasDigits As Boolean
        
        inputText = Cbx_name.Text
        hasDigits = False
        
        ' ----------ѕерев≥р¤Їмо кожен символ у введеному текст≥
        For i = 1 To Len(inputText)
            If IsNumeric(Mid(inputText, i, 1)) Then
                hasDigits = True
                Exit For
            End If
        Next i
        
        ' ---------¬иводимо пов≥домленн¤, ¤кщо цифри знайден≥
        If hasDigits Then
            MsgBox "Ѕудь ласка, видал≥ть цифри з пр≥звища!", vbOKOnly + vbExclamation, "ѕерев≥рка введенн¤"
            GoTo name
        End If
        If Cbx_name.Value = Empty Then
            MsgBox "Ѕудь ласка, введ≥ть пр≥звище", vbOKOnly + vbExclamation, "ѕерев≥рка введенн¤"
            GoTo name
        End If
        GoTo Chk_temperatura
name:       Cbx_name.SetFocus
            Cbx_name.SelStart = 0
            Cbx_name.SelLength = Len(inputText)
    Exit Sub
        
        '----------ѕерев≥рка введенн¤ температури---------------------------
Chk_temperatura:
        If Txb_temperature.Value = Empty Then
            MsgBox "Ѕудь ласка, введ≥ть температуру", vbOKOnly + vbExclamation, "ѕерев≥рка введенн¤"
            Cancel = True
            GoTo temperature
        End If
        Temp = IsNumeric(Txb_temperature)
        If (Temp = "False") Or (Txb_temperature.Value > 40) Or (Txb_temperature.Value < 0) Then
            MsgBox "“емпература - це число в≥д 0 до 40", vbOKOnly + vbExclamation, "ѕерев≥рка введенн¤"
            GoTo temperature
        End If
        GoTo Chk_humidity
temperature:        Txb_temperature.SetFocus
                    Txb_temperature.SelStart = 0
                    Txb_temperature.SelLength = 5
    Exit Sub
        '----------ѕерев≥рка введенн¤ вологост≥---------------------------
Chk_humidity:
        If Txb_humidity.Value = Empty Then
            MsgBox "Ѕудь ласка, введ≥ть волог≥сть", vbOKOnly + vbExclamation, "ѕерев≥рка введенн¤"
            GoTo humidity
        End If
            Temp = IsNumeric(Txb_humidity)
        If (Temp = "False") Or (Txb_humidity.Value > 90) Or (Txb_humidity.Value < 20) Then
            MsgBox "¬олог≥сть - це число в≥д 20 до 90", vbOKOnly + vbExclamation, "ѕерев≥рка введенн¤"
            GoTo humidity
        End If
        GoTo Chk_pressure
humidity:     Txb_humidity.SetFocus
            Txb_humidity.SelStart = 0
            Txb_humidity.SelLength = 5
        Exit Sub
        '----------ѕерев≥рка введенн¤ тиску--------------------ThisWorkbook.Worksheets("¬÷-98")-----------
Chk_pressure:
        If press = 3 Then
            If Txb_pressure.Value = Empty Then
                    MsgBox "Ѕудь ласка, введ≥ть тиск в мм.рт.ст.", vbOKOnly + vbExclamation, "ѕерев≥рка введенн¤"
                    GoTo pressure
            End If
            Temp = IsNumeric(Txb_pressure)
            If (Temp = "False") Or (Txb_pressure.Value > 798) Or (Txb_pressure.Value < 650) Then
                    MsgBox "“иск мм.рт.ст - це число в≥д 650 до 798", vbOKOnly + vbExclamation, "ѕерев≥рка введенн¤"
                    GoTo pressure
            End If
        End If
    
        If press = 2 Then
            If Txb_pressure.Value = Empty Then
                MsgBox "Ѕудь ласка, введ≥ть  тиск в кѕа", vbOKOnly + vbExclamation, "ѕерев≥рка введенн¤"
                GoTo pressure
            End If
            Temp = IsNumeric(Txb_pressure)
            If (Temp = "False") Or (Txb_pressure.Value > 106.5) Or (Txb_pressure.Value < 86.6) Then
                MsgBox "“иск кѕа - це число в≥д 86,6 до 106,5", vbOKOnly + vbExclamation, "ѕерев≥рка введенн¤"
                GoTo pressure
            End If
        End If
        GoTo write_row
pressure:       Txb_pressure.SetFocus
                Txb_pressure.SelStart = 0
                Txb_pressure.SelLength = 5
        Exit Sub
write_row:
        Call Unprotect_ws
        '----------ƒодаванн¤ р¤дка------------------------------------------
        Set ListRow = ListObj.ListRows.Add
        ListRow.Range(1) = Date
        ListRow.Range(1).NumberFormat = "dd.mm.yyyy;@"
        ListRow.Range(2) = Time
        ListRow.Range(2).NumberFormat = "hh:mm;@"
        'ListRow.Range(3) = CStr(ThisWorkbook.Worksheets("Data").Range("B6"))
        ListRow.Range(3) = room
        ListRow.Range(4) = CStr(Form98.Cbx_name.Value)
        ListRow.Range(5) = CDbl(Form98.Txb_temperature.Value)
        ListRow.Range(6) = CByte(Form98.Txb_humidity.Value)
        If press = 1 Then
            ListRow.Range(7) = CStr("-")
            ThisWorkbook.Worksheets(ws_Name).Range("G9").Value = "“иск," & Chr(10) & "мм.рт.ст." & Chr(10) & "кѕа"
        End If
        If press = 2 Or press = 3 Then
            ListRow.Range(7) = CDbl(Form98.Txb_pressure.Value)
        End If
        
        If press = 2 Then
            ThisWorkbook.Worksheets(ws_Name).Range("G9").Value = "“иск," & Chr(10) & "кѕа"
        End If
            
        If press = 3 Then
            ThisWorkbook.Worksheets(ws_Name).Range("G9").Value = "“иск," & Chr(10) & "мм.рт.ст."
        End If
       
        ListRow.Range(8) = CStr("-")
        Call Protect_ws
        'Call ProtectData
        Unload Form98
        ws.Activate
        Count = ListObj.ListRows.Count
        ListObj.DataBodyRange(Count, 1).Select
        'MsgBox "   «апис додано   ", vbOKOnly + vbInformation, "  ¬≥таю!  "
        FormGeneral.Label_choice_1.Caption = "«апис додано"
        FormGeneral.Controls("Label_choice_" & ws_Number).ForeColor = RGB(0, 128, 0) ' RGB(червоний, темно«≈Ћ≈Ќ»…, син≥й / red, darkGREEN, blue)
        ThisWorkbook.Save
        Exit Sub
    End If
    '****код дл¤ вц-90*****************************************
    'MsgBox "дл¤ вц-90 код не поки що зроблено"
End Sub

Private Sub UserForm_Initialize()
    Call FindParam
    
        '----при вибор≥ обладнанн¤ активувати ком≥рку першого стовпчика станнього р¤дка
            Dim ListObj As ListObject
            'Dim tableNameVC90 As String '«м≥нна в ¤к≥й ≥м'¤ таблиц≥ VC90_tab_ на аркуш≥ обранного обладнанн¤
            'tableNameVC90 = "VC90_tab_" & ws_Number
            'MsgBox "ws_Number - " & ws_Number
            'MsgBox "tableNameVC90 - " & tableNameVC90
            Set ws = ThisWorkbook.Worksheets(ws_Number)
            Set ListObj = ws.ListObjects("VC98_tab")
            Dim Count As Integer '«м≥нна в ¤к≥й к≥льк≥сть р¤дк≥в
            Count = ws.ListObjects("VC98_tab").ListRows.Count
            Set ws = ThisWorkbook.Worksheets(ws_Number)
            
            ' ¬становлюЇмо назву форми
            Form98.Caption = "         ¬÷-98     " & ThisWorkbook.Worksheets("Data").Range("A2").Value
            
            'MsgBox "Count - " & Count
            ws.Activate
            ' якщо Ї дан≥ в таблиц≥, вид≥л¤Їмо ком≥рку першого стовпц¤ останнього р¤дка
            If Count > 0 Then
                ListObj.DataBodyRange(Count, 1).Select
            Else
                ' якщо таблиц¤ порожн¤, вид≥л¤Їмо ком≥рку A10
                ws.Range("A10").Select
                'MsgBox "“аблиц¤ порожн¤, вибрано ком≥рку A10"
            End If
        '----
    
    Me.Label_date.Caption = Date
    Me.Label_place.Caption = ThisWorkbook.Worksheets("Data").Range("B6").Value & " " & ThisWorkbook.Worksheets("Data").Range("A2").Value
    'Me.Label_place.Caption = room
    Me.Cbx_name.Value = ""
    Me.Txb_humidity.Value = ""
    Me.Txb_pressure.Value = ""
    FlagGeneral = 1
    If press = 1 Then
        Me.Label_pressure.Visible = False
        Me.Txb_pressure.Visible = False
        Me.Label7.Visible = False
        Exit Sub
    End If
        Me.Label_pressure.Visible = True
        Me.Txb_pressure.Visible = True
        Me.Label7.Visible = True
    If press = 2 Then
        Me.Label_pressure.Caption = "кѕа"
        Exit Sub
    End If
    If press = 3 Then
        Me.Label_pressure.Caption = "мм.рт.ст"
    End If

End Sub


