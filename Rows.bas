Attribute VB_Name = "Rows"
Option Explicit
Sub Delete_Unselected_Rows()
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''                                This Macro was written by                                '''
'''                                      Asher Danner                                       '''
'''                                       03/01/2012                                        '''
'''The purpose is to delete all rows that do not have selected cells in them, or if the     '''
'''whole row is selected.                                                                   '''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Application.ScreenUpdating = False

Dim row_count As Long
Dim this_row As Long

row_count = ActiveSheet.UsedRange.Rows.Count
For this_row = row_count To 1 Step -1
    If is_column_or_row_selected("rows", this_row) = False Then
        ActiveSheet.Rows(this_row).EntireRow.Delete
    End If
Next this_row

Application.ScreenUpdating = True
End Sub
Sub Hide_Unselected_Rows()
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''                                This Macro was written by                                '''
'''                                      Asher Danner                                       '''
'''                                       03/01/2012                                        '''
'''The purpose is to hide all rows that do not have selected cells in them, or if the       '''
'''whole row is selected.                                                                   '''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Application.ScreenUpdating = False

Dim row_count As Long
Dim this_row As Long

row_count = ActiveSheet.UsedRange.Rows.Count
For this_row = row_count To 1 Step -1
    If is_column_or_row_selected("rows", this_row) = False Then
        ActiveSheet.Rows(this_row).EntireRow.Hidden = True
    End If
Next this_row

Application.ScreenUpdating = True
End Sub
Sub Unhide_All_Rows()
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''                                This Macro was written by                                '''
'''                                      Asher Danner                                       '''
'''                                       03/01/2012                                        '''
'''The purpose is to unhide all rows in the usedrange of a sheet                            '''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Application.ScreenUpdating = False
Call func_Unhide_All_Rows(ActiveSheet)
Application.ScreenUpdating = True
End Sub
Function func_Unhide_All_Rows(ByRef this_sht As Worksheet)
this_sht.UsedRange.EntireRow.Hidden = False
End Function
Sub Delete_All_Hidden_Rows()
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''                                This Macro was written by                                '''
'''                                      Asher Danner                                       '''
'''                                       03/01/2012                                        '''
'''The purpose is to delete all rows in a worksheet that are hidden in the usedrange of a   '''
'''sheet.                                                                                   '''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Application.ScreenUpdating = False
Dim row_count As Long
Dim this_row As Long

row_count = ActiveSheet.UsedRange.Rows.Count
For this_row = row_count To 1 Step -1
    If ActiveSheet.Rows(this_row).EntireRow.Hidden = True Then
        ActiveSheet.Rows(this_row).EntireRow.Delete
    End If
Next this_row
Application.ScreenUpdating = True
End Sub
Sub Unique_Value_Spacing_for_Selected_Column()
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''                                This Macro was written by                                '''
'''                                      Asher Danner                                       '''
'''                                       03/09/2012                                        '''
'''The purpose is to add line spacing between each unique value in the selected column.     '''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Dim row_count As Long
Dim this_row As Long
Dim this_str As String
Dim last_str As String
Dim selected_column As Long
Dim space_number As String
Dim n As Long
Dim selection_address As String

space_number = InputBox(Prompt:="How many spaces do you want added?", _
    Title:="Spacing", Default:="1")
If IsNumeric(space_number) = False Then
    MsgBox "Please enter a number. Run the macro again."
    Exit Sub
End If

Application.ScreenUpdating = False
row_count = ActiveSheet.UsedRange.Rows.Count

selection_address = Selection.Address(ReferenceStyle:=xlR1C1)

If is_only_one_column_or_row_selected("column", selection_address) = False Then
    MsgBox "Please only select one column at a time"
    Exit Sub
End If
selected_column = CLng(Right(selection_address, Len(selection_address) - InStr(selection_address, "C")))
For this_row = row_count To 1 Step -1
    this_str = Cells(this_row, selected_column)
    If this_str <> last_str Then
        For n = 1 To space_number
            ActiveSheet.Rows(this_row + 1).EntireRow.Insert
        Next n
        last_str = this_str
    End If
Next this_row
Application.ScreenUpdating = True
End Sub

