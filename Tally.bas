Attribute VB_Name = "Tally"
Option Explicit
Sub IBIS_Tally_Results()
Call IBIS_Tally.pre_fill("G", "H", "I", "K")
IBIS_Tally.Show
End Sub
Sub Tally_Results()
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''                                This Macro was written by                                '''
'''                                      Asher Danner                                       '''
'''                                       05/23/2012                                        '''
'''This macro tallies results in the selected area of each sheet of the workbookon sheet 1. '''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Application.ScreenUpdating = False
Dim sheet_count As Long
Dim n As Long
Dim selection_address As String
Dim rCell As Range
selection_address = Selection.Address
Selection.Validation.Delete
Sheets(1).Range(selection_address) = 0
For n = 2 To ActiveWorkbook.Sheets.Count
    For Each rCell In Sheets(n).Range(selection_address)
        If Trim(rCell) <> "" Then
            Sheets(1).Cells(rCell.Row, rCell.Column) = Sheets(1).Cells(rCell.Row, rCell.Column) + 1
        End If
    Next rCell
Next n
Application.ScreenUpdating = True
End Sub
Sub Tally_Comments()
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''                                This Macro was written by                                '''
'''                                      Asher Danner                                       '''
'''                                       05/23/2012                                        '''
'''This macro prints the text in the selected area of each sheet of the workbook on sheet 1.'''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Dim selection_address As String
Dim rCell As Range
Dim n As Byte

Application.ScreenUpdating = False
Selection.Clear
selection_address = Selection.Address
For n = 2 To ActiveWorkbook.Sheets.Count
    For Each rCell In Sheets(n).Range(selection_address)
        If rCell <> "" Then
            If Sheets(n).Cells(rCell.Row, rCell.Column) <> "" Then
                If Sheets(1).Cells(rCell.Row, rCell.Column) = "" Then
                    Sheets(1).Cells(rCell.Row, rCell.Column) = "• " & Sheets(n).Cells(rCell.Row, rCell.Column)
                Else
                    Sheets(1).Cells(rCell.Row, rCell.Column) = _
                        Sheets(1).Cells(rCell.Row, rCell.Column) & Chr(10) & _
                        "• " & Sheets(n).Cells(rCell.Row, rCell.Column)
                End If
            End If
        End If
    Next rCell
Next n
Application.ScreenUpdating = True
End Sub
Sub Add_X_Number_of_Sheets()
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''                                This Macro was written by                                '''
'''                                      Asher Danner                                       '''
'''                                       07/20/2012                                        '''
'''This macro creates X copies of the current sheet.                                        '''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Dim user_input As String
Dim sht_name As String
Dim cur_sht As Worksheet
Dim n As Byte
Set cur_sht = ActiveSheet
user_input = InputBox(Prompt:="How many sheets do you want to add?", _
          Title:="Sheet count", Default:="1")
If CStr(user_input) = "" Or IsNumeric(user_input) = False Then Exit Sub

For n = 1 To user_input
    Sheets.Add after:=Sheets(ActiveWorkbook.Sheets.Count)
    ActiveSheet.Name = format_sheet_name(cur_sht.Name, ActiveWorkbook)
    cur_sht.UsedRange.Copy Destination:=ActiveSheet.Cells(1, 1)
    cur_sht.UsedRange.Copy
    ActiveSheet.Cells(1, 1).PasteSpecial 8
    With ActiveSheet.PageSetup
        .LeftHeader = cur_sht.PageSetup.LeftHeader
        .CenterHeader = cur_sht.PageSetup.CenterHeader
        .RightHeader = cur_sht.PageSetup.RightHeader
        .LeftFooter = cur_sht.PageSetup.LeftFooter
        .CenterFooter = cur_sht.PageSetup.CenterFooter
        .RightFooter = cur_sht.PageSetup.RightFooter
        .LeftMargin = cur_sht.PageSetup.LeftMargin
        .RightMargin = cur_sht.PageSetup.RightMargin
        .TopMargin = cur_sht.PageSetup.TopMargin
        .BottomMargin = cur_sht.PageSetup.BottomMargin
        .HeaderMargin = cur_sht.PageSetup.HeaderMargin
        .FooterMargin = cur_sht.PageSetup.FooterMargin
        .CenterHorizontally = cur_sht.PageSetup.CenterHorizontally
        .CenterVertically = cur_sht.PageSetup.CenterVertically
        .Orientation = cur_sht.PageSetup.Orientation
        .PaperSize = cur_sht.PageSetup.PaperSize
        .Zoom = cur_sht.PageSetup.Zoom
    End With
    Application.CutCopyMode = False
Next n
End Sub
Sub Make_Sheet_X_Active()
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''                                This Macro was written by                                '''
'''                                      Asher Danner                                       '''
'''                                       10/01/2012                                        '''
'''This macro makes sheet number X active on all selected workbooks.                        '''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Dim this_book As Workbook
Dim xl_file_name As Variant
Dim this_workbook As Long
Dim user_input As String

xl_file_name = Application.GetOpenFilename("Excel files (*.xl*),*.xl*", , _
    "Browse for file to be merged", MultiSelect:=True)
If Not IsArray(xl_file_name) Then Exit Sub
user_input = InputBox(Prompt:="Which sheet number do you want to select?", _
          Title:="Sheet Number", Default:="1")
Application.DisplayAlerts = False
Application.ScreenUpdating = False

If IsArray(xl_file_name) Then
    For this_workbook = LBound(xl_file_name) To UBound(xl_file_name) ' iterate through each book
        Application.Workbooks.Open filename:=xl_file_name(this_workbook), UpdateLinks:=False
        Set this_book = ActiveWorkbook
        If ActiveWorkbook.Sheets.Count < user_input Then
            MsgBox this_book.Name & " does not have " & user_input & " sheets."
        Else
            this_book.Sheets(CInt(user_input)).Select
            this_book.Save
        End If
        this_book.Close
    Next this_workbook
End If
Application.DisplayAlerts = True
Application.ScreenUpdating = True
End Sub
Sub Get_Combinations()
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''                                This Macro was written by                                '''
'''                                      Asher Danner                                       '''
'''                                       09/04/2013                                        '''
'''This macro give the combinations of a string of characters that are seperated by commas. '''
'''The max correct tells how many of the characters can be in a set.                        '''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Combinations.Show
End Sub
