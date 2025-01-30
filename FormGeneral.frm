VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FormGeneral 
   Caption         =   "Вибір обладнання"
   ClientHeight    =   11670
   ClientLeft      =   11850
   ClientTop       =   630
   ClientWidth     =   12090
   OleObjectBlob   =   "FormGeneral.frx":0000
End
Attribute VB_Name = "FormGeneral"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public EnableEvents As Boolean

Private Sub UserForm_Initialize()
    EnableEvents = True
    Dim i As Integer
    Dim OptionButton As Object
    Dim LabelControl As Object
    Dim LabelChoice As Object
    Dim LabelSheet As Object
    Dim wsData As Worksheet
    Set wsData = ThisWorkbook.Worksheets("Data")    ' Вказуємо аркуш "Data"

    ' Встановлюємо назву форми
    FormGeneral.Caption = "         Вибір обладнання     " & wsData.Range("A2").Value

    For i = 2 To 26
        If wsData.Cells(3, i).Value = "Yes" Then
            ' Знайти та зробити видимими відповідні OptionButton та Label
            Set OptionButton = Me.Controls("OptionButton" & i - 1)
            Set LabelControl = Me.Controls("Label" & i - 1)
            Set LabelChoice = Me.Controls("Label_choice_" & i - 1)
            Set LabelSheet = Me.Controls("Label_sheet_" & i - 1)
            OptionButton.Visible = True
            LabelControl.Visible = True
            LabelChoice.Visible = True
            LabelChoice.Caption = "" ' Встановлюємо порожній рядок для LabelChoice
            LabelControl.Caption = wsData.ListObjects("TO_" & i - 1).ListColumns(5).DataBodyRange.Cells(1).Value & "  " & wsData.ListObjects("TO_" & i - 1).ListColumns(6).DataBodyRange.Cells(1).Value & " " & wsData.ListObjects("TO_" & i - 1).ListColumns(7).DataBodyRange.Cells(1).Value
            LabelSheet.Visible = True
            LabelSheet.Caption = wsData.Cells(5, i).Value
        Else
            ' Якщо значення дорівнює "Ні", зробити невидимими OptionButton та Label
            Set OptionButton = Me.Controls("OptionButton" & i - 1)
            Set LabelControl = Me.Controls("Label" & i - 1)
            Set LabelChoice = Me.Controls("Label_choice_" & i - 1)
            Set LabelSheet = Me.Controls("Label_sheet_" & i - 1)
            OptionButton.Visible = False
            LabelControl.Visible = False
            'LabelControl.Caption = "" ' Встановлюємо порожній рядок для LabelControl
            LabelChoice.Visible = False
            'LabelChoice.Caption = "" ' Встановлюємо порожній рядок для LabelChoice
            LabelSheet.Visible = False
            'LabelSheet.Caption = ""
        
        End If
    Next i
    
    ' Розмір форми в залежності від останнього видимого аркуша
    For i = 25 To 2 Step -1
        If wsData.Cells(3, i).Value = "Yes" Then
            Me.Height = 625 - 25 * (25 - i)
            Exit For
        End If
    Next i
End Sub

'Private Sub OptionButton1_Click()
'    If EnableEvents Then
'        ws_Number = 1
'        Form98.Show
'        'MsgBox "Був натиснутий OptionButton" & ws_Number
'    End If
'End Sub
'
'
'Private Sub OptionButton2_Click()
'    If EnableEvents Then
'        ws_Number = 2
'        Form90Ar.Show
'    End If
'End Sub
'Private Sub OptionButton3_Click()
'    If EnableEvents Then
'        ws_Number = 3
'        Form90.Show
'    End If
'End Sub
'Private Sub OptionButton4_Click()
'    If EnableEvents Then
'        ws_Number = 4
'        Form90.Show
'    End If
'End Sub
'Private Sub OptionButton5_Click()
'    If EnableEvents Then
'        ws_Number = 5
'        Form90.Show
'    End If
'End Sub
'Private Sub OptionButton6_Click()
'    If EnableEvents Then
'        ws_Number = 6
'        Form90.Show
'    End If
'End Sub
'Private Sub OptionButton7_Click()
'    If EnableEvents Then
'        ws_Number = 7
'        Form90.Show
'    End If
'End Sub
'Private Sub OptionButton8_Click()
'    If EnableEvents Then
'        ws_Number = 8
'        Form90.Show
'    End If
'End Sub
'Private Sub OptionButton9_Click()
'    If EnableEvents Then
'        ws_Number = 9
'        Form90.Show
'    End If
'End Sub
'Private Sub OptionButton10_Click()
'    If EnableEvents Then
'        ws_Number = 10
'        Form90.Show
'    End If
'End Sub
'Private Sub OptionButton11_Click()
'    If EnableEvents Then
'        ws_Number = 11
'        Form90.Show
'    End If
'End Sub
'Private Sub OptionButton12_Click()
'    If EnableEvents Then
'        ws_Number = 12
'        Form90.Show
'    End If
'End Sub
'Private Sub OptionButton13_Click()
'    If EnableEvents Then
'        ws_Number = 13
'        Form90.Show
'    End If
'End Sub
'Private Sub OptionButton14_Click()
'    If EnableEvents Then
'        ws_Number = 14
'        Form90.Show
'    End If
'End Sub
'Private Sub OptionButton15_Click()
'    If EnableEvents Then
'        ws_Number = 15
'        Form90.Show
'    End If
'End Sub
'Private Sub OptionButton16_Click()
'    If EnableEvents Then
'        ws_Number = 16
'        Form90.Show
'    End If
'End Sub
'Private Sub OptionButton17_Click()
'    If EnableEvents Then
'        ws_Number = 17
'        Form90.Show
'    End If
'End Sub
'Private Sub OptionButton18_Click()
'    If EnableEvents Then
'        ws_Number = 18
'        Form90.Show
'    End If
'End Sub
'Private Sub OptionButton19_Click()
'    If EnableEvents Then
'        ws_Number = 19
'        Form90.Show
'    End If
'End Sub
'Private Sub OptionButton20_Click()
'    If EnableEvents Then
'        ws_Number = 20
'        Form90.Show
'    End If
'End Sub
'Private Sub OptionButton21_Click()
'    If EnableEvents Then
'        ws_Number = 21
'        Form90.Show
'    End If
'End Sub
'Private Sub OptionButton22_Click()
'    If EnableEvents Then
'        ws_Number = 22
'        Form90.Show
'    End If
'End Sub
'Private Sub OptionButton23_Click()
'    If EnableEvents Then
'        ws_Number = 23
'        Form90.Show
'    End If
'End Sub
'Private Sub OptionButton24_Click()
'    If EnableEvents Then
'        ws_Number = 24
'        Form90.Show
'    End If
'End Sub
'Private Sub OptionButton25_Click()
'    If EnableEvents Then
'        ws_Number = 25
'        Form90.Show
'    End If
'End Sub
'Private Sub Label1_Click()
'    ws_Number = 1
'    OptionButton1.Value = True
'    Form98.Show
'End Sub
Private Sub Label1_Click()
    If EnableEvents Then
        ws_Number = 1
        OptionButton1.Value = True
        Form98.Show
    End If
End Sub
'Private Sub Label2_Click()
'    ws_Number = 2
'    OptionButton1.Value = True
'    Form90Ar.Show
'End Sub
Private Sub Label2_Click()
    If EnableEvents Then
        ws_Number = 2
        OptionButton2.Value = True
        Form90Ar.Show
    End If
End Sub
Private Sub Label3_Click()
    If EnableEvents Then
        ws_Number = 3
        OptionButton3.Value = True
        Form90.Show
    End If
End Sub
Private Sub Label4_Click()
    If EnableEvents Then
        ws_Number = 4
        OptionButton4.Value = True
        Form90.Show
    End If
End Sub
Private Sub Label5_Click()
    If EnableEvents Then
        ws_Number = 5
        OptionButton5.Value = True
        Form90.Show
    End If
End Sub
Private Sub Label6_Click()
    If EnableEvents Then
        ws_Number = 6
        OptionButton6.Value = True
        Form90.Show
    End If
End Sub
Private Sub Label7_Click()
    If EnableEvents Then
        ws_Number = 7
        OptionButton7.Value = True
        Form90.Show
    End If
End Sub
Private Sub Label8_Click()
    If EnableEvents Then
        ws_Number = 8
        OptionButton8.Value = True
        Form90.Show
    End If
End Sub
Private Sub Label9_Click()
    If EnableEvents Then
        ws_Number = 9
        OptionButton9.Value = True
        Form90.Show
    End If
End Sub
Private Sub Label10_Click()
    If EnableEvents Then
        ws_Number = 10
        OptionButton10.Value = True
        Form90.Show
    End If
End Sub
Private Sub Label11_Click()
    If EnableEvents Then
        ws_Number = 11
        OptionButton11.Value = True
        Form90.Show
    End If
End Sub
Private Sub Label12_Click()
    If EnableEvents Then
        ws_Number = 12
        OptionButton12.Value = True
        Form90.Show
    End If
End Sub
Private Sub Label13_Click()
    If EnableEvents Then
        ws_Number = 13
        OptionButton13.Value = True
        Form90.Show
    End If
End Sub
Private Sub Label14_Click()
    If EnableEvents Then
        ws_Number = 14
        OptionButton14.Value = True
        Form90.Show
    End If
End Sub
Private Sub Label15_Click()
    If EnableEvents Then
        ws_Number = 15
        OptionButton15.Value = True
        Form90.Show
    End If
End Sub
Private Sub Label16_Click()
    If EnableEvents Then
        ws_Number = 16
        OptionButton16.Value = True
        Form90.Show
    End If
End Sub
Private Sub Label17_Click()
    If EnableEvents Then
        ws_Number = 17
        OptionButton17.Value = True
        Form90.Show
    End If
End Sub
Private Sub Label18_Click()
    If EnableEvents Then
        ws_Number = 18
        OptionButton18.Value = True
        Form90.Show
    End If
End Sub
Private Sub Label19_Click()
    If EnableEvents Then
        ws_Number = 19
        OptionButton19.Value = True
        Form90.Show
    End If
End Sub
Private Sub Label20_Click()
    If EnableEvents Then
        ws_Number = 20
        OptionButton20.Value = True
        Form90.Show
    End If
End Sub
Private Sub Label21_Click()
    If EnableEvents Then
        ws_Number = 21
        OptionButton21.Value = True
        Form90.Show
    End If
End Sub
Private Sub Label22_Click()
    If EnableEvents Then
        ws_Number = 22
        OptionButton22.Value = True
        Form90.Show
    End If
End Sub
Private Sub Label23_Click()
    If EnableEvents Then
        ws_Number = 23
        OptionButton23.Value = True
        Form90.Show
    End If
End Sub
Private Sub Label24_Click()
    If EnableEvents Then
        ws_Number = 24
        OptionButton24.Value = True
        Form90.Show
    End If
End Sub
Private Sub Label25_Click()
    If EnableEvents Then
        ws_Number = 25
        OptionButton25.Value = True
        Form90.Show
    End If
End Sub
