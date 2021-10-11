Attribute VB_Name = "formatting"
Sub Trim_Cells()
Dim rCell As Range

Application.ScreenUpdating = False
Call delete_extraneous_blank_rows_and_columns(ActiveSheet)

For Each rCell In ActiveSheet.UsedRange.Cells
    If rCell.Value2 <> "" Then
        rCell.Value2 = Replace(rCell.Value2, Chr(160), " ")
        rCell.Value2 = Trim(rCell.Value2)
    End If
Next rCell

Application.ScreenUpdating = True

End Sub
