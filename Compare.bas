Attribute VB_Name = "Compare"
Sub Compare_Sheets_1and2()
Dim sht_1 As Worksheet
Dim sht_2 As Worksheet
Dim rCell As Range
Set sht_1 = ActiveWorkbook.Sheets(1)
Set sht_2 = ActiveWorkbook.Sheets(2)

sht_1.UsedRange.Cells.Interior.Color = xlNone
sht_2.UsedRange.Cells.Interior.Color = xlNone

If sht_1.UsedRange.Cells.Count < sht_2.UsedRange.Cells.Count Then
Set check_range = sht_2.UsedRange.Cells
Else
Set check_range = sht_1.UsedRange.Cells
End If

For Each rCell In check_range
If IsError(sht_1.Range(rCell.Address).value) Or IsError(sht_2.Range(rCell.Address)) Then
    If IsError(sht_1.Range(rCell.Address).value) <> IsError(sht_2.Range(rCell.Address)) Then
        sht_1.Range(rCell.Address).Interior.Color = RGB(255, 0, 0)
        sht_2.Range(rCell.Address).Interior.Color = RGB(255, 0, 0)
    End If
ElseIf sht_1.Range(rCell.Address) <> sht_2.Range(rCell.Address) Then
    sht_1.Range(rCell.Address).Interior.Color = RGB(255, 0, 0)
    sht_2.Range(rCell.Address).Interior.Color = RGB(255, 0, 0)
End If
Next rCell
End Sub
