Attribute VB_Name = "Rows_and_Columns"
Option Explicit
Sub Delete_Unselected_Rows_and_Columns()
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''                                This Macro was written by                                '''
'''                                      Asher Danner                                       '''
'''                                       03/01/2012                                        '''
'''The purpose is to delete all rows and columns that do not have selected cells in them,   '''
'''or if the whole row or column selected.                                                  '''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Application.ScreenUpdating = False

Dim row_count As Long
Dim this_row As Long
Dim col_count As Long
Dim this_col As Long

col_count = ActiveSheet.UsedRange.Columns.Count
row_count = ActiveSheet.UsedRange.Rows.Count
For this_row = row_count To 1 Step -1
    If is_column_or_row_selected("rows", this_row) = False Then
        ActiveSheet.Rows(this_row).EntireRow.Delete
    End If
Next this_row
For this_col = col_count To 1 Step -1
    If is_column_or_row_selected("columns", this_col) = False Then
        ActiveSheet.Columns(this_col).EntireColumn.Delete
    End If
Next this_col

Application.ScreenUpdating = True
End Sub
Sub Hide_Unselected_Rows_and_Columns()
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''                                This Macro was written by                                '''
'''                                      Asher Danner                                       '''
'''                                       03/01/2012                                        '''
'''The purpose is to hide all rows and columns that do not have selected cells in them, or  '''
'''if the whole row or column selected.                                                     '''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Application.ScreenUpdating = False

Dim row_count As Long
Dim this_row As Long
Dim col_count As Long
Dim this_col As Long

col_count = ActiveSheet.UsedRange.Columns.Count
row_count = ActiveSheet.UsedRange.Rows.Count
For this_row = row_count To 1 Step -1
    If is_column_or_row_selected("rows", this_row) = False Then
        ActiveSheet.Rows(this_row).EntireRow.Hidden = True
    End If
Next this_row
For this_col = col_count To 1 Step -1
    If is_column_or_row_selected("columns", this_col) = False Then
        ActiveSheet.Columns(this_col).EntireColumn.Hidden = True
    End If
Next this_col

Application.ScreenUpdating = True
End Sub
Function is_column_or_row_selected(row_or_col As String, row_or_col_location As Long) As Boolean
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''                                This Macro was written by                                '''
'''                                      Asher Danner                                       '''
'''                                       03/01/2012                                        '''
'''The purpose is to return true or false based on if the row or column is selected.        '''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Dim selection_address As String
Dim selection_address_split As Variant
Dim selection_range As Variant
Dim first_in_range As Long
Dim last_in_range As Long
Dim i As Long

row_or_col = LCase(row_or_col) ' format row_or_col string
selection_address = Selection.Address(ReferenceStyle:=xlR1C1)

selection_address_split = split(selection_address, ",")
' check to see if all rows or all columns are selected
' if a row is selected and we are looking for columns then all columns are selected
' if a column is selected and we are looking for rows then all rows are selected
If InStr(selection_address, "R") <> 0 And InStr(selection_address, "C") = 0 Then ' only rows selected
    If row_or_col = "columns" Then        ' if we are looking for columns then
        is_column_or_row_selected = True  ' all columns are selected
        Exit Function
    End If
ElseIf InStr(selection_address, "C") <> 0 And InStr(selection_address, "R") = 0 Then ' only columns selected
    If row_or_col = "rows" Then           ' if we are looking for rows then
        is_column_or_row_selected = True  ' all rows are selected
        Exit Function
    End If
End If

' depending on if we are looking for rows or columns then
' we must only look at row or column information
For i = LBound(selection_address_split) To UBound(selection_address_split)
    Set selection_range = convert_R1C1_to_range(selection_address_split(i))
    If row_or_col = "columns" Then
        first_in_range = selection_range.Column
        last_in_range = _
            selection_range.Cells(selection_range.Cells.Count).Column
        If row_or_col_location >= first_in_range And row_or_col_location <= last_in_range Then
            is_column_or_row_selected = True
            Exit Function
        Else
            is_column_or_row_selected = False
        End If
    ElseIf row_or_col = "rows" Then
        first_in_range = selection_range.Row
        last_in_range = _
            selection_range.Cells(selection_range.Cells.Count).Row
        If row_or_col_location >= first_in_range And row_or_col_location <= last_in_range Then
            is_column_or_row_selected = True
            Exit Function
        Else
            is_column_or_row_selected = False
        End If
    End If
Next i
End Function
Function convert_R1C1_to_range(r1c1_string As Variant) As Range
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''                                This Macro was written by                                '''
'''                                      Asher Danner                                       '''
'''                                       03/01/2012                                        '''
'''The purpose is to return a range based on cells(R1,C1) given a string range or single    '''
'''cell as a string in R1C1 format                                                          '''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Dim cell_array As Variant

If InStr(r1c1_string, ":") <> 0 Then ' this is a range
    cell_array = split(r1c1_string, ":")
    Set convert_R1C1_to_range = _
        Range(convert_R1C1_to_cell(cell_array(0)), convert_R1C1_to_cell(cell_array(1)))
Else ' not a range
    Set convert_R1C1_to_range = convert_R1C1_to_cell(r1c1_string)
End If
End Function
Function convert_R1C1_to_cell(r1c1_string As Variant)
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''                                This Macro was written by                                '''
'''                                      Asher Danner                                       '''
'''                                       03/01/2012                                        '''
'''The purpose is to return a cell based on cells(R1,C1) given a single cell as a string in '''
'''R1C1 format                                                                              '''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Dim range_array As Variant
range_array = split(Replace(r1c1_string, "R", ""), "C")

If InStr(r1c1_string, "R") = 0 Then
    Set convert_R1C1_to_cell = ActiveSheet.Columns(CLng(range_array(1)))
ElseIf InStr(r1c1_string, "C") = 0 Then
    Set convert_R1C1_to_cell = ActiveSheet.Rows(CLng(range_array(0)))
Else
    range_array = split(Replace(r1c1_string, "R", ""), "C")
    Set convert_R1C1_to_cell = Cells(CLng(range_array(0)), CLng(range_array(1)))
End If
End Function
Function is_only_one_column_or_row_selected(row_or_col As String, r1c1_address As Variant) As Boolean
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''                                This Macro was written by                                '''
'''                                      Asher Danner                                       '''
'''                                       03/09/2012                                        '''
'''This checks to make sure there is not a range of columns or rows selected.               '''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
If InStr(r1c1_address, ",") <> 0 Or InStr(r1c1_address, ":") <> 0 Then
    is_only_one_column_or_row_selected = False
    Exit Function
End If
If LCase(row_or_col) = "row" Or LCase(row_or_col) = "r" Then
    If InStr(r1c1_address, "C") <> 0 Then
        is_only_one_column_or_row_selected = False
        Exit Function
    End If
ElseIf LCase(row_or_col) = "column" Or LCase(row_or_col) = "c" Or LCase(row_or_col) = "col" Then
    If InStr(r1c1_address, "R") <> 0 Then
        is_only_one_column_or_row_selected = False
        Exit Function
    End If
End If
is_only_one_column_or_row_selected = True
End Function
Sub Border_All_Cells_With_Data()
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''                                This Macro was written by                                '''
'''                                      Asher Danner                                       '''
'''                                       03/12/2012                                        '''
'''The purpose is to put all borders around cells that contain data.                        '''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Dim rCell       As Range:   Dim rngEnd     As Range

For Each rCell In ActiveSheet.UsedRange.Cells
    If rCell.value <> "" Then
        rCell.Borders(xlEdgeLeft).LineStyle = xlContinuous
        rCell.Borders(xlEdgeTop).LineStyle = xlContinuous
        rCell.Borders(xlEdgeRight).LineStyle = xlContinuous
        rCell.Borders(xlEdgeBottom).LineStyle = xlContinuous
        
        If rCell.MergeCells Then
            Set rngEnd = rCell.MergeArea
            rngEnd.Borders(xlEdgeLeft).LineStyle = xlContinuous
            rngEnd.Borders(xlEdgeTop).LineStyle = xlContinuous
            rngEnd.Borders(xlEdgeRight).LineStyle = xlContinuous
            rngEnd.Borders(xlEdgeBottom).LineStyle = xlContinuous
        End If
    End If
Next rCell
End Sub
Sub test_del()
Call delete_extraneous_blank_rows_and_columns(ActiveSheet)
End Sub
Sub delete_extraneous_blank_rows_and_columns(work_sheet As Worksheet)
Dim row_count As Long
Dim col_count As Long
Dim del_rng As Range

With work_sheet
    row_count = .UsedRange.Rows.Count
    col_count = .UsedRange.Columns.Count
    If row_count = 0 Then row_count = 1
    If col_count = 0 Then col_count = 1
    
    If row_count = 1 And col_count = 1 And .Cells(1, 1) = "" Then Exit Sub
    Do Until Application.WorksheetFunction.CountA(.Columns(col_count)) <> 0 Or col_count = 1
        If Not del_rng Is Nothing Then
            Set del_rng = Union(del_rng, .Columns(col_count))
        Else
            Set del_rng = .Columns(col_count)
        End If
        col_count = col_count - 1
    Loop
    If Not del_rng Is Nothing Then del_rng.EntireColumn.Delete
    Set del_rng = Nothing
    Do Until Application.WorksheetFunction.CountA(.Rows(row_count)) <> 0 Or row_count = 1
        If Not del_rng Is Nothing Then
            Set del_rng = Union(del_rng, .Rows(row_count))
        Else
            Set del_rng = .Rows(row_count)
        End If
        row_count = row_count - 1
    Loop
    If Not del_rng Is Nothing Then del_rng.EntireRow.Delete
    
End With
End Sub
Function is_only_one_cell_selected(this_address As Range) As Boolean
If this_address.Cells.Count > 1 Then
    is_only_one_cell_selected = False
Else
    is_only_one_cell_selected = True
End If
End Function
