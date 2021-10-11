Attribute VB_Name = "Word"
Option Explicit
Sub Word_Tables_To_Excel()
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''                                This Macro was written by                                '''
'''                                      Asher Danner                                       '''
'''                                       03/06/2012                                        '''
'''The purpose is to export all tables from selected MS Word documents into an Excel sheet. '''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Dim table_no As Integer 'table number in Word
Dim word_row As Long 'row index in Word
Dim word_col As Long 'column index in Word
Dim excel_col As Integer 'column index in Excel
Dim excel_row As Integer 'row index in Excel
Dim wd_doc As Object
Dim wd_file_name As Variant
Dim i As Long

wd_file_name = Application.GetOpenFilename("Word files (*.do*),*.do*", , _
"Browse for file containing table to be imported", MultiSelect:=True)

Application.ScreenUpdating = False

If IsArray(wd_file_name) Then
    For i = LBound(wd_file_name) To UBound(wd_file_name)
        Set wd_doc = GetObject(wd_file_name(i))
        With wd_doc
            If wd_doc.tables.Count <> 0 Then
                For table_no = 1 To wd_doc.tables.Count
                    excel_row = excel_row + 1
                    excel_col = 1
                    With .tables(table_no)
                        'copy cell contents from Word table cells to Excel cells
                        For word_row = 1 To .Rows.Count
                            For word_col = 1 To .Columns.Count
                                excel_col = excel_col + 1
                                On Error Resume Next
                                ActiveSheet.Cells(excel_row, excel_col) = WorksheetFunction.Clean(Replace(.cell(word_row, word_col).Range.Text, Chr(13), " "))
                                On Error GoTo 0
                            Next word_col
                        Next word_row
                    End With
                Next table_no
            End If
        End With
        Set wd_doc = Nothing
    Next i
Else
    If wd_file_name = False Then Exit Sub      '(user cancelled import file browser)
End If

Application.ScreenUpdating = True
End Sub
Sub Word_Tables_To_Excel_Row()
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''                                This Macro was written by                                '''
'''                                      Asher Danner                                       '''
'''                                       07/13/2015                                        '''
'''The purpose is to export all tables from selected MS Word documents into an Excel sheet. '''
'''Each table will be in a single Excel row.                                                '''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Dim table_no As Integer 'table number in Word
Dim rCell As Variant
Dim excel_row As Long ' row index in Excel
Dim excel_col As Integer ' column index in Excel
Dim wd_doc As Object
Dim wd_file_name As Variant
Dim i As Long

wd_file_name = Application.GetOpenFilename("Word files (*.do*),*.do*", , _
"Browse for file containing table to be imported", MultiSelect:=True)

Application.ScreenUpdating = False

If IsArray(wd_file_name) Then
    For i = LBound(wd_file_name) To UBound(wd_file_name)
        Set wd_doc = GetObject(wd_file_name(i))
            If wd_doc.tables.Count <> 0 Then
                For table_no = 1 To wd_doc.tables.Count
                    excel_row = excel_row + 1
                    excel_col = 1
                    With wd_doc.tables(table_no)
                        'copy cell contents from Word table cells to Excel cells
                        For Each rCell In .Range.Cells
                                excel_col = excel_col + 1
                                On Error Resume Next
                                ActiveSheet.Cells(excel_row, excel_col) = Trim(WorksheetFunction.Clean(Replace(rCell.Range.Text, Chr(13), " ")))
                                On Error GoTo 0
                        Next rCell
                    End With
                Next table_no
            End If
        Set wd_doc = Nothing
    Next i
Else
    If wd_file_name = False Then Exit Sub      '(user cancelled import file browser)
End If

Application.ScreenUpdating = True
End Sub

