Attribute VB_Name = "adv_pst_spc"
Option Explicit
Sub Advanced_Paste_Special()
Adv_Paste_Special.Show
End Sub
Function paste_spec_cur_bk(only_acive As Boolean, paste_all As Boolean, paste_formats As Boolean, paste_formulas As Boolean, _
                      paste_values As Boolean, paste_widths As Boolean)
Dim wb As Workbook
Dim sht As Worksheet
Dim from_sht As Long
Dim to_sht As Long
Dim n As Long
Dim rng As Range

Set wb = ActiveWorkbook
Set sht = wb.ActiveSheet

If only_acive Then
    from_sht = sht.Index
    to_sht = sht.Index
Else
    from_sht = 1
    to_sht = wb.Sheets.Count
End If

For n = from_sht To to_sht
    Set sht = wb.Sheets(n)
    If paste_all Then
        sht.UsedRange.Copy
        sht.UsedRange.PasteSpecial xlPasteAll
    End If
    If paste_formats Then
        sht.UsedRange.Copy
        sht.UsedRange.PasteSpecial xlPasteFormats
    End If
    If paste_formulas Then
        sht.UsedRange.Copy
        sht.UsedRange.PasteSpecial xlPasteFormulas
    End If
    If paste_values Then
        sht.UsedRange.Copy
        sht.UsedRange.PasteSpecial xlPasteValues
    End If
    If paste_widths Then
        sht.UsedRange.Copy
        sht.UsedRange.PasteSpecial xlPasteColumnWidths
    End If
Next n
Application.CutCopyMode = False
End Function
Function paste_spec_split(folder_path As String, only_acive As Boolean, paste_all As Boolean, paste_formats As Boolean, paste_formulas As Boolean, _
                     paste_values As Boolean, paste_widths As Boolean, whole_book As String)

Dim wb As Workbook
Dim sht As Worksheet
Dim from_sht As Long
Dim to_sht As Long
Dim n As Long
Dim to_bk As Workbook
Dim this_sht As Worksheet
Dim created As Boolean
Dim save_name As String

Set wb = ActiveWorkbook
Set sht = wb.ActiveSheet

If only_acive Then
    from_sht = sht.Index
    to_sht = sht.Index
Else
    from_sht = 1
    to_sht = wb.Sheets.Count
End If

created = False
For n = from_sht To to_sht
    Set sht = wb.Sheets(n)
    If Not whole_book Or Not created Then
        Application.Workbooks.Add
        Set to_bk = ActiveWorkbook
        Call delete_unneeded_sheets(to_bk)
        Set this_sht = to_bk.ActiveSheet
        created = True
    ElseIf whole_book Then
        to_bk.Sheets.Add after:=to_bk.Sheets(to_bk.Sheets.Count)
        Set this_sht = to_bk.Sheets(to_bk.Sheets.Count)
    End If
    If paste_all Then
        sht.UsedRange.Copy
        this_sht.Cells(1, 1).PasteSpecial xlPasteAll
    End If
    If paste_formats Then
        sht.UsedRange.Copy
        this_sht.Cells(1, 1).PasteSpecial xlPasteFormats
    End If
    If paste_formulas Then
        sht.UsedRange.Copy
        this_sht.Cells(1, 1).PasteSpecial xlPasteFormulas
    End If
    If paste_values Then
        sht.UsedRange.Copy
        this_sht.Cells(1, 1).PasteSpecial xlPasteValues
    End If
    If paste_widths Then
        sht.UsedRange.Copy
        this_sht.Cells(1, 1).PasteSpecial xlPasteColumnWidths
    End If
    If Not whole_book Then
        to_bk.SaveAs filename:=folder_path & "\" & get_unique_filename(this_sht.Name & ".xlsx"), FileFormat:=xlOpenXMLWorkbook
    End If
    
Next n

If whole_book Then
    save_name = wb.Name
    to_bk.SaveAs filename:=folder_path & "\" & get_unique_filename(save_name), FileFormat:=wb.FileFormat
End If
Application.CutCopyMode = False
End Function

