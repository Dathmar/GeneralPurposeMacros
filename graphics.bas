Attribute VB_Name = "graphics"
Sub graphics_merge()

Dim wb As Workbook
Dim ws1 As Worksheet
Dim ws2 As Worksheet
Dim ws3 As Worksheet

Set wb = ActiveWorkbook

If wb.Sheets.Count < 2 Then
    MsgBox ("Workbook does not have enough sheets")
    Exit Sub
End If

Set ws1 = wb.Sheets(1)
Set ws2 = wb.Sheets(2)

If wb.Sheets.Count < 3 Then wb.Sheets.Add after:=wb.Sheets(wb.Sheets.Count)

Set ws3 = wb.Sheets(3)

ws3.Cells(1, 1) = "Art Code/Name/Accnum"
ws3.Cells(1, 2) = "Project/Program"
ws3.Cells(1, 3) = "Priority"
ws3.Cells(1, 4) = "Subject"
ws3.Cells(1, 5) = "Grade Level"
ws3.Cells(1, 6) = "Workflow"
ws3.Cells(1, 7) = "Modified"
ws3.Cells(1, 8) = "Modified By"


If ws1.Cells(1, 1) = "Art Code/Name/Accnum" Then
    Call get_IBIS(ws2, ws3)
    Call get_SP(ws1, ws3)
Else
    Call get_IBIS(ws1, ws3)
    Call get_SP(ws2, ws3)
End If

End Sub
Function get_IBIS(from_sht As Worksheet, to_sht As Worksheet)
Dim col As Long
Dim to_row As Long

to_row = to_sht.UsedRange.Rows.Count + 1

col = get_column_by_header(from_sht, "Name")
If col <> 0 Then
    Call copy_col_to_location(from_sht, col, to_sht, 1, to_row)
    col = 0
Else
    col = get_column_by_header(from_sht, "External Client ID")
    If col <> 0 Then
        Call copy_col_to_location(from_sht, col, to_sht, 1, to_row)
        col = 0
    Else
        col = get_column_by_header(from_sht, "Item Accnum")
        If col <> 0 Then
            Call copy_col_to_location(from_sht, col, to_sht, 1, to_row)
            col = 0
        End If
    End If
End If

col = get_column_by_header(from_sht, "Program")
If col <> 0 Then
    Call copy_col_to_location(from_sht, col, to_sht, 2, to_row)
    col = 0
End If

col = get_column_by_header(from_sht, "Ad hoc Step")
If col <> 0 Then
    Call copy_col_to_location(from_sht, col, to_sht, 3, to_row)
    col = 0
End If

col = get_column_by_header(from_sht, "Test")
If col <> 0 Then
    Call copy_col_to_location(from_sht, col, to_sht, 4, to_row)
    Call copy_col_to_location(from_sht, col, to_sht, 5, to_row)
    col = 0
End If


col = get_column_by_header(from_sht, "Workflow Step")
If col <> 0 Then
    Call copy_col_to_location(from_sht, col, to_sht, 6, to_row)
    col = 0
End If

col = get_column_by_header(from_sht, "Last Updated Date")
If col <> 0 Then
    Call copy_col_to_location(from_sht, col, to_sht, 7, to_row)
    col = 0
End If

col = get_column_by_header(from_sht, "Last Updated By")
If col <> 0 Then
    Call copy_col_to_location(from_sht, col, to_sht, 8, to_row)
    col = 0
End If

End Function
Function get_SP(from_sht As Worksheet, to_sht As Worksheet)
Dim col As Long
Dim to_row As Long

to_row = to_sht.UsedRange.Rows.Count + 1

col = get_column_by_header(from_sht, "Art Code/Name/Accnum")
If col <> 0 Then
    Call copy_col_to_location(from_sht, col, to_sht, 1, to_row)
    col = 0
End If

col = get_column_by_header(from_sht, "Project/Program")
If col <> 0 Then
    Call copy_col_to_location(from_sht, col, to_sht, 2, to_row)
    col = 0
End If

col = get_column_by_header(from_sht, "Priority")
If col <> 0 Then
    Call copy_col_to_location(from_sht, col, to_sht, 3, to_row)
    col = 0
End If

col = get_column_by_header(from_sht, "Subject")
If col <> 0 Then
    Call copy_col_to_location(from_sht, col, to_sht, 4, to_row)
    col = 0
End If

col = get_column_by_header(from_sht, "Grade Level")
If col <> 0 Then
    Call copy_col_to_location(from_sht, col, to_sht, 5, to_row)
    col = 0
End If

col = get_column_by_header(from_sht, "Workflow")
If col <> 0 Then
    Call copy_col_to_location(from_sht, col, to_sht, 6, to_row)
    col = 0
End If

col = get_column_by_header(from_sht, "Modified")
If col <> 0 Then
    Call copy_col_to_location(from_sht, col, to_sht, 7, to_row)
    col = 0
End If

col = get_column_by_header(from_sht, "Modified By")
If col <> 0 Then
    Call copy_col_to_location(from_sht, col, to_sht, 8, to_row)
    col = 0
End If

End Function
Function copy_col_to_location(from_sht As Worksheet, from_col As Long, to_sht As Worksheet, to_col As Long, to_row As Long)
Dim lr As Long

With from_sht
    lr = .UsedRange.Rows.Count + 1
    .Range(.Cells(2, from_col), .Cells(lr, from_col)).Copy Destination:=to_sht.Cells(to_row, to_col)
End With

End Function
Function get_column_by_header(ws As Worksheet, header As String, Optional check_case = True) As Long
Dim lc As Long

lc = ws.UsedRange.Columns.Count

For n = 1 To lc
    If ws.Cells(1, n) = header And check_case Then
        get_column_by_header = n
        Exit Function
    ElseIf LCase(ws.Cells(1, n)) = LCase(header) And Not check_case Then
        get_column_by_header = n
        Exit Function
    End If
Next n
get_column_by_header = 0

End Function
