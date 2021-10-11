Attribute VB_Name = "Testing"
Option Explicit
Sub Columns_Times_Rows_Test()
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''                                This Macro was written by                                '''
'''                                      Asher Danner                                       '''
'''                                       03/02/2012                                        '''
'''This macro really has no purpose other than to fill in a range of cells with the column  '''
'''number times the row number.  This is used for testing other macros.                     '''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Dim to_address As String
Dim cell_range As Range
Dim range_array As Variant
Dim cell_in_range As Range
Dim r1c1_string As String

to_address = InputBox(Prompt:="Give cell address in R1C1 format.", _
          Title:="Cell address", Default:="R50C29")
If to_address = vbNullString Then
    Exit Sub
Else
    Set cell_range = convert_R1C1_to_range("R1C1:" & to_address)
    For Each cell_in_range In cell_range
        r1c1_string = cell_in_range.Address(ReferenceStyle:=xlR1C1)
        range_array = split(Replace(r1c1_string, "R", ""), "C")
        cell_in_range.value = CLng(range_array(0)) * CLng(range_array(1))
    Next cell_in_range
End If
End Sub
Sub Spiral_Test()
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''                                This Macro was written by                                '''
'''                                      Asher Danner                                       '''
'''                                       03/02/2012                                        '''
'''This macro really has no purpose other than to fill in a range of cells with a spiral    '''
'''pattern starting with the selected cells.                                                '''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Dim from_col As Long
Dim from_row As Long
Dim go_left As Boolean
Dim go_down As Boolean
Dim left_col As Long
Dim right_col As Long
Dim top_row As Long
Dim bot_row As Long
Dim n As Long

If is_only_one_cell_selected(Selection) = False Then
    MsgBox "Please only select one cell"
Else
    from_col = Selection.Column
    from_row = Selection.Row
End If
left_col = from_col
right_col = from_col
bot_row = from_row
top_row = from_row
Do Until from_col = 0 And from_row = 0

    If from_col > 0 And from_row > 0 Then
        Cells(from_row, from_col) = n
        n = n + 1
    End If
    If from_col = left_col - 1 Then
        left_col = from_col
        go_left = False
    ElseIf from_col = right_col + 1 Then
        right_col = from_col
        go_left = True
    End If
    If from_row = top_row - 1 Then
        top_row = from_row
        go_down = True
    ElseIf from_row = bot_row + 1 Then
        bot_row = from_row
        go_down = False
    End If
    
    If go_left = False And go_down = False Then
        from_col = from_col + 1
    ElseIf go_left = False And go_down = True Then
        from_row = from_row + 1
    ElseIf go_left = True And go_down = False Then
        from_row = from_row - 1
    Else
        from_col = from_col - 1
    End If
Loop

End Sub
