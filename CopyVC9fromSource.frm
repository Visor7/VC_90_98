VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} CopyVC9fromSource 
   Caption         =   "CopyVC9fromSource"
   ClientHeight    =   3150
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10065
   OleObjectBlob   =   "CopyVC9fromSource.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "CopyVC9fromSource"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()
    ClearTablesVC9_OnAllSheets ' Очистити всі таблиці
    Dim sourcePath As String
    Dim sourceWorkbook As Workbook
    Dim sourceSheet As Worksheet
    Dim sourceTable As ListObject
    Dim destTable As ListObject
    Dim destSheet As Worksheet
    Dim destWorkbook As Workbook
    Dim tblName As String
    Dim password As String
    Dim i As Long, j As Long
    Dim sourceCell As Range
    Dim destCell As Range
    Dim hasVC9Table As Boolean
    ClearTablesVC9_OnAllSheets
    ' Розблокувати всі аркуші destWorkbook / Unprotect all sheets in the destination workbook
    Unprotect_ws_all
    ' Встановити пароль / Set the password
    password = "lab123"
    
    ' Отримати шлях до файлу з TextBox / Get the file path from the TextBox
    sourcePath = TextBox1.Text
    
    ' Відкрити файл джерела / Open the source workbook
    Set sourceWorkbook = Workbooks.Open(sourcePath)
    
    ' Відкрити файл призначення (поточний файл) / Open the destination workbook (current workbook)
    Set destWorkbook = ThisWorkbook
    
    ' Перебрати всі листи у файлі джерела / Loop through all sheets in the source workbook
    For Each sourceSheet In sourceWorkbook.Sheets
        hasVC9Table = False
        
        ' Перевірити, чи на аркуші є таблиця, яка починається з "VC9" / Check if the sheet contains a table that starts with "VC9"
        For Each sourceTable In sourceSheet.ListObjects
            If Left(sourceTable.name, 3) = "VC9" Then
                hasVC9Table = True
                Exit For
            End If
        Next sourceTable
        
        ' Якщо на аркуші є таблиця з "VC9", знімаємо захист / If the sheet contains a "VC9" table, unprotect the sheet
        If hasVC9Table Then
            sourceSheet.Unprotect password:=password
        
            ' Перебрати всі таблиці на листі / Loop through all tables on the sheet
            For Each sourceTable In sourceSheet.ListObjects
                If Left(sourceTable.name, 3) = "VC9" Then
                    tblName = sourceTable.name
                    
                    ' Знайти відповідну таблицю у файлі призначення / Find the corresponding table in the destination workbook
                    Set destSheet = destWorkbook.Sheets(sourceSheet.name)
                    
                    ' Знімаємо захист з аркуша призначення / Unprotect the destination sheet
                    destSheet.Unprotect password:=password
                    
                    ' Перевірити, чи існує таблиця у файлі призначення / Check if the table exists in the destination workbook
                    On Error Resume Next
                    Set destTable = destSheet.ListObjects(tblName)
                    On Error GoTo 0
                    
                    If Not destTable Is Nothing Then
                        ' Перевірити наявність даних у таблиці джерела / Check if the source table contains data
                        If Not sourceTable.DataBodyRange Is Nothing Then
                            If Application.WorksheetFunction.CountA(sourceTable.DataBodyRange) > 0 Then
                                ' Перевірити та додати рядки, якщо потрібно / Check and add rows if needed
                                Dim rowDiff As Long
                                rowDiff = sourceTable.ListRows.Count - destTable.ListRows.Count
                                If rowDiff > 0 Then
                                    destTable.ListRows.Add AlwaysInsert:=True
                                End If
                                
                                MsgBox ("Джерело: " & sourceSheet.name & " : " & tblName & " row - " & sourceTable.DataBodyRange.Rows.Count & " col - " & sourceTable.DataBodyRange.Columns.Count)
                                'MsgBox ("Source: " & sourceSheet.name & " : " & tblName & " row - " & sourceTable.DataBodyRange.Rows.Count & " col - " & sourceTable.DataBodyRange.Columns.Count)
                                
                                ' Копіювати дані по одній комірці / Copy data cell by cell
                                For i = 1 To sourceTable.DataBodyRange.Rows.Count
                                    For j = 1 To sourceTable.DataBodyRange.Columns.Count
                                        Set sourceCell = sourceTable.DataBodyRange.Cells(i, j)
                                        Set destCell = destTable.DataBodyRange.Cells(i, j)
                                        destCell.Value = sourceCell.Value
                                    Next j
                                Next i
                            Else
                                MsgBox "Таблиця " & tblName & " у файлі джерела порожня."
                                'MsgBox "Table " & tblName & " in the source file is empty."
                            End If
                        Else
                            MsgBox "Таблиця " & tblName & " у файлі джерела не містить даних."
                            'MsgBox "Table " & tblName & " in the source file contains no data."
                        End If
                    Else
                        MsgBox "Таблиця " & tblName & " не знайдена у файлі призначення."
                        'MsgBox "Table " & tblName & " not found in the destination file."
                    End If
                    
                    ' Заблокувати аркуш призначення / Protect the destination sheet
                    destSheet.Protect password:=password
                End If
            Next sourceTable
        
            ' Заблокувати аркуш джерела після обробки / Protect the source sheet after processing
            sourceSheet.Protect password:=password
        End If
    Next sourceSheet
    
    ' Закрити файл джерела без збереження / Close the source workbook without saving
    sourceWorkbook.Close SaveChanges:=False
    
    ' Заблокувати всі аркуші destWorkbook / Protect all sheets in the destination workbook
    Protect_ws_all
    ' Повідомлення про завершення / Completion message
    MsgBox "Копіювання завершено!"
    'MsgBox "Copying completed!"
    Unload CopyVC9fromSource
End Sub

