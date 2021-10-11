Attribute VB_Name = "Spliting"
Option Explicit
Sub Split_Sheets_to_Workbooks()
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''                                This Macro was written by                                '''
'''                                      Asher Danner                                       '''
'''                                       03/14/2012                                        '''
'''The purpose is split all sheets in a workbook into individual workbooks.                 '''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Dim n As Long
Dim this_book As Workbook
Dim new_book As Workbook
Dim file_path As String
Dim file_name As String

file_path = browse_for_folder()
If file_path = "False" Then Exit Sub
Application.ScreenUpdating = False
Set this_book = ActiveWorkbook
For n = 1 To this_book.Sheets.Count
    Application.Workbooks.Add
    Set new_book = ActiveWorkbook
    Call delete_unneeded_sheets(new_book)
    this_book.Sheets(n).Rows(1).Copy
    new_book.Sheets("Sheet1").Cells(1, 1).PasteSpecial 8
    this_book.Sheets(n).UsedRange.Copy Destination:=new_book.Sheets("Sheet1").Cells(1, 1)
    With new_book.ActiveSheet.PageSetup
        .LeftHeader = this_book.ActiveSheet.PageSetup.LeftHeader
        .CenterHeader = this_book.ActiveSheet.PageSetup.CenterHeader
        .RightHeader = this_book.ActiveSheet.PageSetup.RightHeader
        .LeftFooter = this_book.ActiveSheet.PageSetup.LeftFooter
        .CenterFooter = this_book.ActiveSheet.PageSetup.CenterFooter
        .RightFooter = this_book.ActiveSheet.PageSetup.RightFooter
        .LeftMargin = this_book.ActiveSheet.PageSetup.LeftMargin
        .RightMargin = this_book.ActiveSheet.PageSetup.RightMargin
        .TopMargin = this_book.ActiveSheet.PageSetup.TopMargin
        .BottomMargin = this_book.ActiveSheet.PageSetup.BottomMargin
        .HeaderMargin = this_book.ActiveSheet.PageSetup.HeaderMargin
        .FooterMargin = this_book.ActiveSheet.PageSetup.FooterMargin
        .CenterHorizontally = this_book.ActiveSheet.PageSetup.CenterHorizontally
        .CenterVertically = this_book.ActiveSheet.PageSetup.CenterVertically
        .Orientation = this_book.ActiveSheet.PageSetup.Orientation
        .PaperSize = this_book.ActiveSheet.PageSetup.PaperSize
        .Zoom = this_book.ActiveSheet.PageSetup.Zoom
    End With
    
    file_name = get_unique_filename(file_path & "\" & this_book.Sheets(n).Name & ".xlsx")
    
    new_book.SaveAs filename:=file_name
    new_book.Close
Next n
Application.ScreenUpdating = True
End Sub
Sub Split_Unique_Values_to_Sheets()
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''                                This Macro was written by                                '''
'''                                      Asher Danner                                       '''
'''                                       03/14/2012                                        '''
'''The purpose is split all unique value sets in a selected column to new worksheets which  '''
'''are then named after the unique values.                                                  '''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Dim this_sht, new_sht As Worksheet
Dim this_array As Variant
Dim elmt As Variant
Dim file_path As String
Dim wb As Workbook
Dim this_col As Long


Set wb = ActiveWorkbook
Set this_sht = ActiveSheet

If is_only_one_column_or_row_selected("col", Selection.Address(ReferenceStyle:=xlR1C1)) = False Then Exit Sub
Application.ScreenUpdating = False
this_col = CLng(Replace(Selection.Address(ReferenceStyle:=xlR1C1), "C", ""))
this_array = get_unique_values(this_sht, this_col)
For Each elmt In this_array
    wb.Sheets.Add after:=wb.Sheets(wb.Sheets.Count)
    Set new_sht = ActiveSheet
    this_sht.UsedRange.AutoFilter Field:=this_col, Criteria1:=elmt
    this_sht.Rows(1).Copy
    new_sht.Cells(1, 1).PasteSpecial 8
    this_sht.UsedRange.SpecialCells(xlCellTypeVisible).Copy Destination:=new_sht.Cells(1, 1)
    If elmt = "" Then elmt = "Blanks"
    new_sht.Name = format_sheet_name(CStr(elmt), wb)
Next elmt
this_sht.ShowAllData
Application.ScreenUpdating = True
End Sub
Sub Split_Unique_Values_to_Books()
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''                                This Macro was written by                                '''
'''                                      Asher Danner                                       '''
'''                                       03/14/2012                                        '''
'''The purpose is split all unique value sets in a selected column to new workbooks which   '''
'''are then named after the unique values.                                                  '''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Dim this_sht As Worksheet
Dim this_array As Variant
Dim elmt As Variant
Dim file_path As String
Dim new_book As Workbook
Dim this_col As Long
Dim save_path As String

Set this_sht = ActiveSheet
If is_only_one_column_or_row_selected("col", Selection.Address(ReferenceStyle:=xlR1C1)) = False Then Exit Sub
file_path = browse_for_folder()
If file_path = "False" Then Exit Sub
Application.ScreenUpdating = False
this_col = CLng(Replace(Selection.Address(ReferenceStyle:=xlR1C1), "C", ""))
this_array = get_unique_values(this_sht, this_col)
For Each elmt In this_array
    Application.Workbooks.Add
    Set new_book = ActiveWorkbook
    Call delete_unneeded_sheets(new_book)
    this_sht.UsedRange.AutoFilter Field:=this_col, Criteria1:=elmt
    this_sht.Rows(1).Copy
    new_book.ActiveSheet.Cells(1, 1).PasteSpecial 8
    this_sht.UsedRange.SpecialCells(xlCellTypeVisible).Copy Destination:=new_book.Sheets(1).Cells(1, 1)
    If elmt = "" Then elmt = "Blanks"
    elmt = Replace(elmt, "?", "")
    elmt = Replace(elmt, "\", "")
    elmt = Replace(elmt, "/", "")
    elmt = Replace(elmt, ":", "")
    elmt = Replace(elmt, "<", "")
    elmt = Replace(elmt, ">", "")
    elmt = Replace(elmt, "|", "")
    
    save_path = file_path & "\" & elmt & ".xlsx"
    
    save_path = get_unique_filename(save_path)
    new_book.SaveAs filename:=save_path, FileFormat:=xlOpenXMLWorkbook
    new_book.Close
Next elmt
Application.ScreenUpdating = True
End Sub
Sub Split_Equal_Row_Count_to_Books()
Split_by_equal_rows.Show
End Sub
Function browse_for_folder(Optional OpenAt As Variant) As Variant
     'Function purpose:  To Browser for a user selected folder.
     'If the "OpenAt" path is provided, open the browser at that directory
     'NOTE:  If invalid, it will open at the Desktop level
     'Written by Ken Puls from VBAexpress.com
     
    Dim ShellApp As Object
     
     'Create a file browser window at the default folder
    Set ShellApp = CreateObject("Shell.Application"). _
    BrowseForFolder(0, "Please choose a folder", 0, OpenAt)
     
     'Set the folder to that selected.  (On error in case cancelled)
    On Error Resume Next
    browse_for_folder = ShellApp.self.path
    On Error GoTo 0
     
     'Destroy the Shell Application
    Set ShellApp = Nothing
     
     'Check for invalid or non-entries and send to the Invalid error
     'handler if found
     'Valid selections can begin L: (where L is a letter) or
     '\\ (as in \\servername\sharename.  All others are invalid
    Select Case Mid(browse_for_folder, 2, 1)
    Case Is = ":"
        If Left(browse_for_folder, 1) = ":" Then GoTo Invalid
    Case Is = "\"
        If Not Left(browse_for_folder, 1) = "\" Then GoTo Invalid
    Case Else
        GoTo Invalid
    End Select
     
    Exit Function
     
Invalid:
     'If it was determined that the selection was invalid, set to False
    browse_for_folder = False
     
End Function

