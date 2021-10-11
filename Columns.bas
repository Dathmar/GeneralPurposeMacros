Attribute VB_Name = "Columns"
Option Explicit
Sub Fill_Below_in_Selected_Columns()
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''                                This Macro was written by                                '''
'''                                      Asher Danner                                       '''
'''                                       03/01/2012                                        '''
'''The purpose is copy cell information to cells below in selected columns.                 '''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Dim sht As Worksheet
Dim row_count As Long
Dim rCol As Range
Dim n As Long
Dim this_name As String

Application.ScreenUpdating = False

Set sht = ActiveSheet
Call delete_extraneous_blank_rows_and_columns(sht)

row_count = sht.UsedRange.Rows.Count
For Each rCol In Selection.Columns
    this_name = sht.Cells(2, rCol.Column)
    For n = 1 To row_count
        If sht.Cells(n, rCol.Column) = "" Then
            sht.Cells(n, rCol.Column) = this_name
        Else
            this_name = sht.Cells(n, rCol.Column)
        End If
    Next n
Next rCol

Application.ScreenUpdating = True
End Sub
Sub Delete_Unselected_Columns()
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''                                This Macro was written by                                '''
'''                                      Asher Danner                                       '''
'''                                       03/01/2012                                        '''
'''The purpose is to delete all columns that do not have selected cells in them, or if the  '''
'''whole column is selected.                                                                '''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Application.ScreenUpdating = False

Dim col_count As Long
Dim this_col As Long

col_count = ActiveSheet.UsedRange.Columns.Count
For this_col = col_count To 1 Step -1
    If is_column_or_row_selected("columns", this_col) = False Then
        ActiveSheet.Columns(this_col).EntireColumn.Delete
    End If
Next this_col

Application.ScreenUpdating = True
End Sub
Sub Hide_Unselected_Columns()
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''                                This Macro was written by                                '''
'''                                      Asher Danner                                       '''
'''                                       03/01/2012                                        '''
'''The purpose is to hide all columns that do not have selected cells in them, or if the    '''
'''whole column is selected.                                                                '''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Application.ScreenUpdating = False

Dim col_count As Long
Dim this_col As Long

col_count = ActiveSheet.UsedRange.Columns.Count
For this_col = col_count To 1 Step -1
    If is_column_or_row_selected("columns", this_col) = False Then
        ActiveSheet.Columns(this_col).EntireColumn.Hidden = True
    End If
Next this_col

Application.ScreenUpdating = True
End Sub
Sub Unhide_All_Columns()
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''                                This Macro was written by                                '''
'''                                      Asher Danner                                       '''
'''                                       03/01/2012                                        '''
'''The purpose is to unhide all columns in the usedrange of a sheet                         '''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Application.ScreenUpdating = False
Call func_Unhide_All_Columns(ActiveSheet)
Application.ScreenUpdating = True
End Sub
Function func_Unhide_All_Columns(ByRef this_sht As Worksheet)
this_sht.UsedRange.EntireColumn.Hidden = False
End Function
Sub Delete_All_Hidden_Columns()
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''                                This Macro was written by                                '''
'''                                      Asher Danner                                       '''
'''                                       03/01/2012                                        '''
'''The purpose is to delete all columns in a worksheet that are hidden in the usedrange of  '''
'''a sheet.                                                                                 '''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Application.ScreenUpdating = False
Dim col_count As Long
Dim this_col As Long

col_count = ActiveSheet.UsedRange.Columns.Count
For this_col = col_count To 1 Step -1
    If ActiveSheet.Columns(this_col).EntireColumn.Hidden = True Then
        ActiveSheet.Columns(this_col).EntireColumn.Delete
    End If
Next this_col
Application.ScreenUpdating = True
End Sub
Sub Batch_Delete_All_Hidden_Columns()
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''                                This Macro was written by                                '''
'''                                      Asher Danner                                       '''
'''                                       10/14/2012                                        '''
'''The purpose is to delete all columns in a worksheet that are hidden in the usedrange of  '''
'''a sheet.                                                                                 '''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Dim xl_file_name As Variant
Dim this_book As Workbook
Dim this_workbook As Long

xl_file_name = Application.GetOpenFilename("Excel files (*.xl*),*.xl*", , _
    "Browse for file to be merged", MultiSelect:=True)

If IsArray(xl_file_name) Then
    For this_workbook = LBound(xl_file_name) To UBound(xl_file_name) ' iterate through each book
        Application.Workbooks.Open filename:=xl_file_name(this_workbook), ReadOnly:=False ' open the books
        Set this_book = ActiveWorkbook
        Call Delete_All_Hidden_Columns
        this_book.Save
        this_book.Close
    Next this_workbook
End If
End Sub
Sub Add_Hyperlinks_to_Selected_Column()
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''                                This Macro was written by                                '''
'''                                      Asher Danner                                       '''
'''                                       10/01/2012                                        '''
'''The purpose is to adds hyperlinks to filepath text in the selected column.               '''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Dim n As Long
Dim i As Long
Application.ScreenUpdating = False
If is_only_one_column_or_row_selected("Column", Selection.Address(ReferenceStyle:=xlR1C1)) = False Then Exit Sub
i = Selection.Column
For n = 2 To ActiveSheet.UsedRange.Rows.Count
    ActiveSheet.Hyperlinks.Add Anchor:=Cells(n, i), Address:= _
            Cells(n, i).value, _
            TextToDisplay:=Cells(n, i).value
Next n
Application.ScreenUpdating = True
End Sub
