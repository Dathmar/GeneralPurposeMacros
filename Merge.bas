Attribute VB_Name = "Merge"
Option Explicit
Sub Countif_Merge()
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''                                This Macro was written by                                '''
'''                                      Asher Danner                                       '''
'''                                       03/14/2012                                        '''
'''The purpose is to merge two workbook with two columns that have the same data.  The macro'''
'''checks each row of both books and if they match the data is merged into a new book.      '''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Dim to_book As Workbook
Dim to_sht As Worksheet
Dim from1_book As Workbook
Dim from1_sht As Worksheet
Dim from2_book As Workbook
Dim from2_sht As Worksheet
Dim user_input As String
Dim X As Integer
Dim Y As Integer
Dim n As Long
Dim xl_file_name As Variant
Dim col_count As Integer
Dim last_row As Long


' open the first book and ask for the column that contains the unique field match on
xl_file_name = Application.GetOpenFilename("Excel files (*.xl*),*.xl*", , _
    "Browse for first file to be merged", MultiSelect:=False)
If xl_file_name = False Then Exit Sub
Workbooks.Open filename:=xl_file_name, ReadOnly:=True, Editable:=True
Set from1_book = ActiveWorkbook
Set from1_sht = from1_book.ActiveSheet
user_input = InputBox(Prompt:="In the book showing what column do you want to compare by?", _
          Title:="Column letter or number", Default:="1")
If user_input = "" Then
    from1_book.Close SaveChanges:=False
    Exit Sub
End If
If IsNumeric(user_input) = False Then
    X = get_column_number(user_input)
Else: X = CLng(user_input): End If

' open the second book and ask for the column that contains the unique field match on
xl_file_name = Application.GetOpenFilename("Excel files (*.xl*),*.xl*", , _
    "Browse for second file to be merged", MultiSelect:=False)
If xl_file_name = False Then Exit Sub
Workbooks.Open filename:=xl_file_name, ReadOnly:=True, Editable:=True
Set from2_book = ActiveWorkbook
Set from2_sht = from2_book.ActiveSheet
user_input = InputBox(Prompt:="In the book showing what column do you want to compare by?", _
          Title:="Column letter or number", Default:="1")
If user_input = "" Then Exit Sub
If IsNumeric(user_input) = False Then
    Y = get_column_number(user_input)
Else: Y = CLng(user_input): End If

Application.ScreenUpdating = False
Application.DisplayAlerts = False

Call delete_extraneous_blank_rows_and_columns(from1_sht)
Call delete_extraneous_blank_rows_and_columns(from2_sht)

col_count = from1_sht.UsedRange.Columns.Count

Application.Workbooks.Add
Set to_book = ActiveWorkbook
Set to_sht = to_book.ActiveSheet

'merge each sheet with validation columns between them
from1_sht.UsedRange.Copy Destination:=to_sht.Cells(1, 1)
from2_sht.UsedRange.Copy Destination:=to_sht.Cells(1, col_count + 4)

to_sht.Cells(1, col_count + 1) = "<-- count"
to_sht.Cells(1, col_count + 2) = "Validation"
to_sht.Cells(1, col_count + 3) = "count -->"

last_row = to_sht.UsedRange.Rows.Count

' add count if information
For n = 2 To last_row
    to_sht.Cells(n, col_count + 1) = Application.WorksheetFunction.CountIf(to_sht.Columns(col_count + 3 + Y), to_sht.Cells(n, X))
    to_sht.Cells(n, col_count + 3) = Application.WorksheetFunction.CountIf(to_sht.Columns(X), to_sht.Cells(n, col_count + 3 + Y))
Next n


With to_sht
    ' sort by countif values and the sort data values in book 1
    .Sort.SortFields.Clear
    .Sort.SortFields.Add key:=.Columns(col_count + 1), _
        SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
    .Sort.SortFields.Add key:=.Columns(X), _
        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    .Sort.SetRange .Range(.Cells(1, 1), .Cells(last_row, col_count + 1))
    .Sort.header = xlYes
    .Sort.MatchCase = False
    .Sort.Orientation = xlTopToBottom
    .Sort.SortMethod = xlPinYin
    .Sort.Apply
    
    ' sort by countif values and the sort data values in book 2
    .Sort.SortFields.Clear
    .Sort.SortFields.Add key:=.Columns(col_count + 3), _
        SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
    .Sort.SortFields.Add key:=.Columns(col_count + 3 + Y), _
        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    .Sort.SetRange .Range(.Cells(1, col_count + 3), .Cells(last_row, .UsedRange.Columns.Count))
    .Sort.header = xlYes
    .Sort.MatchCase = False
    .Sort.Orientation = xlTopToBottom
    .Sort.SortMethod = xlPinYin
    .Sort.Apply

    For n = 2 To last_row
        If .Cells(n, X) = .Cells(n, col_count + 3 + Y) Then
            .Cells(n, col_count + 2) = "True"
        Else
            .Cells(n, col_count + 2) = "False"
        End If
    Next n
End With
from1_book.Close SaveChanges:=False
from2_book.Close SaveChanges:=False

Application.DisplayAlerts = True
Application.ScreenUpdating = True
End Sub


Sub Merge_Files_To_End_Of_Sheet()
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''                                This Macro was written by                                '''
'''                                      Asher Danner                                       '''
'''                                       03/14/2012                                        '''
'''                                   Updated 07/10/14                                    '''
'''The purpose is to merge all files that are selected to the end of the current open sheet.'''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Files_to_EOS.Show
End Sub
Sub Merge_Matching_Headers_to_End_of_Sheet()
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''                                This Macro was written by                                '''
'''                                      Asher Danner                                       '''
'''                                       09/11/2012                                        '''
'''Merge many files to the end of a sheet but only merge matching headers.                  '''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Headers_to_EOS.Show
End Sub
Sub Merge_Files_to_Sheets()
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''                                This Macro was written by                                '''
'''                                      Asher Danner                                       '''
'''                                       03/14/2012                                        '''
'''The purpose is to merge all files that are selected to a new sheet.                      '''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Dim from_book As Workbook
Dim to_book As Workbook
Dim xl_file_name As Variant
Dim this_workbook As Long
Dim sht_name As String
Dim n As Long
Dim check_passwords As Variant

xl_file_name = Application.GetOpenFilename("Excel files (*.xl*),*.xl*", , _
    "Browse for file to be merged", MultiSelect:=True)
Application.DisplayAlerts = False
Application.ScreenUpdating = False

If IsArray(xl_file_name) Then
    Workbooks.Add
    Set to_book = ActiveWorkbook
    Call delete_unneeded_sheets(to_book)
    For this_workbook = LBound(xl_file_name) To UBound(xl_file_name) ' iterate through each book
        Application.Workbooks.Open filename:=xl_file_name(this_workbook), ReadOnly:=True, Editable:=True ' open the books
        Set from_book = ActiveWorkbook
        sht_name = format_sheet_name(from_book.name, to_book)
        
        to_book.Worksheets.Add(after:=to_book.Worksheets(to_book.Worksheets.Count)).name = sht_name
        
        from_book.ActiveSheet.UsedRange.Copy Destination:=to_book.Sheets(to_book.Sheets.Count).Cells(1, 1)
        from_book.ActiveSheet.Rows(1).Copy
        to_book.ActiveSheet.Cells(1, 1).PasteSpecial 8
        
        With to_book.ActiveSheet.PageSetup
            .LeftHeader = from_book.ActiveSheet.PageSetup.LeftHeader
            .CenterHeader = from_book.ActiveSheet.PageSetup.CenterHeader
            .RightHeader = from_book.ActiveSheet.PageSetup.RightHeader
            .LeftFooter = from_book.ActiveSheet.PageSetup.LeftFooter
            .CenterFooter = from_book.ActiveSheet.PageSetup.CenterFooter
            .RightFooter = from_book.ActiveSheet.PageSetup.RightFooter
            .LeftMargin = from_book.ActiveSheet.PageSetup.LeftMargin
            .RightMargin = from_book.ActiveSheet.PageSetup.RightMargin
            .TopMargin = from_book.ActiveSheet.PageSetup.TopMargin
            .BottomMargin = from_book.ActiveSheet.PageSetup.BottomMargin
            .HeaderMargin = from_book.ActiveSheet.PageSetup.HeaderMargin
            .FooterMargin = from_book.ActiveSheet.PageSetup.FooterMargin
            .CenterHorizontally = from_book.ActiveSheet.PageSetup.CenterHorizontally
            .CenterVertically = from_book.ActiveSheet.PageSetup.CenterVertically
            .Orientation = from_book.ActiveSheet.PageSetup.Orientation
            .PaperSize = from_book.ActiveSheet.PageSetup.PaperSize
            .Zoom = from_book.ActiveSheet.PageSetup.Zoom
        End With
        from_book.Close
    Next this_workbook
    to_book.Sheets(1).Delete
End If
Application.DisplayAlerts = True
Application.ScreenUpdating = True
End Sub
Sub testdelete_unneeded_sheets()
Call delete_unneeded_sheets(ActiveWorkbook)
End Sub
Function delete_unneeded_sheets(this_book As Workbook)
Dim n As Long
Dim i As Long
If this_book.Sheets.Count > 1 Then
    For n = this_book.Sheets.Count To 2 Step -1
    If Application.WorksheetFunction.CountA(this_book.Sheets(n).Cells) = 0 Then
        this_book.Sheets(n).Delete
    End If
    Next n
End If
End Function
Function format_sheet_name(sht_name As String, ByVal to_book As Workbook) As String
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''                                This Macro was written by                                '''
'''                                      Asher Danner                                       '''
'''                                       03/14/2012                                        '''
'''Returns a string that is formatted to be usable as a sheet name.                         '''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Dim n As Integer
Dim counter As Long
Dim not_same As Boolean
Dim tst_nam As String

counter = 1
not_same = False

sht_name = Replace(sht_name, " ", "_")
sht_name = Replace(sht_name, "?", "")
sht_name = Replace(sht_name, "\", "")
sht_name = Replace(sht_name, "/", "")
sht_name = Replace(sht_name, ":", "")
sht_name = Replace(sht_name, "[", "")
sht_name = Replace(sht_name, "]", "")
If sht_name <> ".xlsx" And sht_name <> ".xls" Then
    sht_name = Replace(sht_name, ".xlsx", "")
    sht_name = Replace(sht_name, ".xls", "")
End If
If Len(sht_name) > 31 Then
    sht_name = Left(sht_name, 29)
End If
Do While Right(sht_name, 1) = "_"
    sht_name = Left(sht_name, Len(sht_name) - 1)
Loop
If sht_name = "" Then
    sht_name = "Blank"
End If
Do While not_same = False
    For n = 1 To to_book.Sheets.Count
        If to_book.Sheets(n).name = sht_name Then
            If Right(sht_name, Len(CStr(counter))) = counter Then
                sht_name = Left(sht_name, Len(sht_name) - Len(CStr(counter)))
            End If
            
            If Len(sht_name & counter) > 31 Then
                sht_name = Left(sht_name, Len(sht_name) - Len(CStr(counter)))
            End If
            sht_name = sht_name & counter
            
            Do While to_book.Sheets(n).name = sht_name
                sht_name = Left(sht_name, Len(sht_name) - Len(CStr(counter)))
                counter = counter + 1
                sht_name = sht_name & counter
            Loop
        Else
            not_same = True
        End If
    Next n
    For n = 1 To to_book.Sheets.Count
        If to_book.Sheets(n).name = sht_name Then
            not_same = False
        End If
    Next n
Loop

format_sheet_name = sht_name
End Function
Sub Merge_Books_With_Same_Data_In_Columns_X_and_Y()
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''                                This Macro was written by                                '''
'''                                      Asher Danner                                       '''
'''                                       03/14/2012                                        '''
'''The purpose is to merge two workbook with two columns that have the same data.  The macro'''
'''checks each row of both books and if they match the data is merged into a new book.      '''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Dim to_book As Workbook
Dim from_book As Workbook
Dim user_input As String
Dim X As Long
Dim Y As Long
Dim xl_file_name As Variant
Dim already_open As Boolean

Set to_book = ActiveWorkbook
Call delete_extraneous_blank_rows_and_columns(to_book.ActiveSheet)
xl_file_name = Application.GetOpenFilename("Excel files (*.xl*),*.xl*", , _
    "Browse for file to be merged", MultiSelect:=False)
If xl_file_name = False Then Exit Sub
user_input = InputBox(Prompt:="In the book showing what column X should match column Y.", _
          Title:="Column X letter or number", Default:="1")
If user_input = "" Then Exit Sub
If IsNumeric(user_input) = False Then
    X = get_column_number(user_input)
Else: X = CLng(user_input): End If

If is_workbook_open(get_filename(xl_file_name)) Then
    Set from_book = Workbooks(get_filename(CStr(xl_file_name)))
    already_open = True
Else
    Application.Workbooks.Open filename:=xl_file_name, ReadOnly:=True
    Set from_book = ActiveWorkbook
End If

Call delete_extraneous_blank_rows_and_columns(from_book.ActiveSheet)
from_book.ActiveSheet.Activate
user_input = InputBox(Prompt:="In the book showing what column Y should match column X.", _
          Title:="Column Y letter or number", Default:="1")
If user_input = "" Then Exit Sub
If IsNumeric(user_input) = False Then
    Y = get_column_number(user_input)
Else: Y = CLng(user_input): End If

Application.ScreenUpdating = False
Application.DisplayAlerts = False
Call merge_books_to_end_column(to_book.ActiveSheet, from_book.ActiveSheet, X, Y)

If Not already_open Then from_book.Close
Application.DisplayAlerts = True
Application.ScreenUpdating = True
End Sub
Function copy_row_to_end(from_this_book As Workbook, to_this_book As Workbook, copy_row As Long, last_row As Long)
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''                                This Macro was written by                                '''
'''                                      Asher Danner                                       '''
'''                                       03/14/2012                                        '''
'''The purpose is to copy a row from one Excel workbook to another excel Workbook.          '''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Dim col_count As Long
Dim i As Long
col_count = from_this_book.ActiveSheet.UsedRange.Columns.Count
For i = 1 To col_count
    to_this_book.ActiveSheet.Cells(last_row, i) = from_this_book.ActiveSheet.Cells(copy_row, i)
Next i
End Function
Function get_column_number(column_string As String) As Long
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''                                This Macro was written by                                '''
'''                                      Asher Danner                                       '''
'''                                       03/14/2012                                        '''
'''Returns the column number when passed a column string. A -> 1, AA -> 27                  '''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
get_column_number = ActiveWorkbook.Sheets(1).Columns(column_string).Column
End Function
Function merge_books_to_end_column(to_sheet As Worksheet, from_sheet As Worksheet, to_match_col As Long, from_match_col As Long, Optional rep_dups As Boolean)
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''                                This Macro was written by                                '''
'''                                      Asher Danner                                       '''
'''                                       03/14/2012                                        '''
'''The purpose is to merge two workbook with two columns that have the same data.  The macro'''
'''uses filters to get unique values and loops through these to match books.                '''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Dim this_array As Variant
Dim status_array() As Variant
Dim n As Long
Dim i As Long
Dim areas_row_cnt As Long
Dim this_str As String
Dim this_sht As String
Dim sht1_count As Long
Dim sht2_count As Long
Dim col_count As Long
Dim report As Variant
Dim this_row As Long
Dim rng1 As Range
Dim rng2 As Range
Dim rArea As Range
Dim found_duplicates As Boolean

If IsMissing(rep_dups) Then rep_dups = True

'sheet 1 is to_sheet
'sheet 2 is from_sheet
Call func_Unhide_All_Rows(to_sheet)
Call func_Unhide_All_Rows(from_sheet)
Call func_Unhide_All_Columns(to_sheet)
Call func_Unhide_All_Columns(from_sheet)
'setup sheets to avoid errors
from_sheet.Cells.UnMerge
to_sheet.Cells.UnMerge
from_sheet.AutoFilterMode = False
to_sheet.AutoFilterMode = False

Application.ScreenUpdating = False
col_count = to_sheet.UsedRange.Columns.Count
' check that the resulting data isn
If col_count + from_sheet.UsedRange.Columns.Count > to_sheet.Columns.Count Then
    MsgBox "There are too many columns to merge all the data as a XLS file." & vbNewLine _
        & "Please save your current book as a XLSX file and rerun this macro."
    Exit Function
End If
this_array = get_unique_values(to_sheet, to_match_col)

from_sheet.UsedRange.Sort key1:=from_sheet.Cells(1, from_match_col), header:=xlYes

For n = LBound(this_array) To UBound(this_array)

    
    status_array() = status_bar(status_array(), n, UBound(this_array)) ' setup status bar
    to_sheet.UsedRange.AutoFilter Field:=to_match_col, Criteria1:=CStr(this_array(n))
    Set rng1 = to_sheet.UsedRange.SpecialCells(xlCellTypeVisible)
    sht1_count = count_rows_in_range(rng1) - 1
    If sht1_count > 1 Then
        found_duplicates = True
    Else
        from_sheet.AutoFilterMode = False
        from_sheet.UsedRange.AutoFilter Field:=from_match_col, Criteria1:=CStr(this_array(n))
        Set rng2 = from_sheet.UsedRange.SpecialCells(xlCellTypeVisible)
        sht2_count = count_rows_in_range(rng2) - 1
        this_row = to_sheet.Cells(to_sheet.Rows.Count, to_match_col).End(xlUp).Row
        
        to_sheet.AutoFilterMode = False
        Call copy_row_n_times(to_sheet, sht2_count - 1, this_row) ' can fix this function to enter the rows all at once.
        If rng2.Areas.Count <> 1 Then
            For i = 2 To rng2.Areas.Count
                rng2.Areas(i).Copy (to_sheet.Cells(this_row + areas_row_cnt, col_count + 1))
                areas_row_cnt = rng2.Areas(i).Rows.Count
            Next i
            areas_row_cnt = 0
        ElseIf rng2.Areas.Count = 1 And rng2.Rows.Count > 1 Then
            rng2.Offset(1).Resize(rng2.Rows.Count - 1).Copy (to_sheet.Cells(this_row, col_count + 1))
        End If
    End If
Next n
from_sheet.Range(from_sheet.Cells(1, 1), from_sheet.Cells(1, from_sheet.UsedRange.Columns.Count)).Copy (to_sheet.Cells(1, col_count + 1))
If found_duplicates And (Not rep_dups) Then
    MsgBox "Found Duplicates these have been skipped"
End If
Application.StatusBar = False
to_sheet.Range("A1").AutoFilter
from_sheet.Range("A1").AutoFilter
Application.ScreenUpdating = True
End Function
Function get_unique_values(ByVal this_sht As Worksheet, this_col As Long) As Variant
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''                                This Macro was written by                                '''
'''                                      Asher Danner                                       '''
'''                                       03/14/2012                                        '''
'''The purpose is to return all unique values in a column as an array.                      '''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Dim row_count As Long
Dim next_col As Long
With this_sht
    next_col = .Cells(1, .Columns.Count).End(xlToLeft).Column + 1
    
    .Columns(this_col).AdvancedFilter Action:=xlFilterCopy, CopyToRange:=.Cells(1, next_col), unique:=True
    row_count = .Cells(.Rows.Count, next_col).End(xlUp).Row
    .Sort.SortFields.Clear
    .Sort.SortFields.Add key:=.Range(.Cells(1, next_col), .Cells(row_count, next_col)), _
        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortNormal
    With .Sort
        .SetRange this_sht.Range(this_sht.Cells(1, next_col), this_sht.Cells(row_count, next_col))
        .header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    get_unique_values = WorksheetFunction.Transpose(.Range(.Cells(2, next_col), .Cells(row_count, next_col)))
    .Columns(next_col).EntireColumn.Delete
    
End With
End Function
Function copy_row_n_times(this_sht As Worksheet, n As Long, copy_row As Long)
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''                                This Macro was written by                                '''
'''                                      Asher Danner                                       '''
'''                                       03/14/2012                                        '''
'''Copies a row n times in the row below current location.                                  '''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Dim insert_at As Long
Dim insert_to As Long
Dim total_rows As Long

If n < 1 Then Exit Function

total_rows = n + 1

insert_at = copy_row + 1
insert_to = copy_row + total_rows

With this_sht
    .Range(.Rows(insert_at), .Rows(insert_to)).Insert
    .Rows(copy_row).Copy Destination:=.Range(.Rows(insert_at), .Rows(insert_to - 1))
End With
End Function
Function count_rows_in_range(rng As Range) As Long
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''                                This Macro was written by                                '''
'''                                      Asher Danner                                       '''
'''                                       03/14/2012                                        '''
'''Count the rows in a continuous or non-continuous range.                                  '''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Dim i As Long
Dim sub_rng As Range

Set sub_rng = rng.End(xlUp)

If Application.WorksheetFunction.CountA(sub_rng) = Application.WorksheetFunction.CountA(rng) Then
    Set rng = sub_rng
End If

For i = 1 To rng.Areas.Count
count_rows_in_range = count_rows_in_range + rng.Areas(i).Rows.Count
Next i
End Function
Function status_bar(status_array() As Variant, current_val As Long, max_val As Long) As Variant
'status_array(0) last percent
'status_array(1) running average time
'status_array(2) last running time
'status_array(3) Times run
ReDim Preserve status_array(0 To 3)
If status_array(0) = "" Then status_array(0) = 0

If status_array(0) = Int((current_val / max_val) * 100) Then
    status_bar = status_array()
    Exit Function
Else
    If status_array(1) = "" Then
        status_array(1) = 0
        status_array(3) = 0
    ElseIf status_array(1) = 0 Then
        status_array(1) = ((status_array(1)) + Time - status_array(2))
        status_array(3) = status_array(3) + 1
    Else
        status_array(1) = ((status_array(1) * status_array(3)) + Time - status_array(2)) / (status_array(3) + 1)
        status_array(3) = status_array(3) + 1
    End If
    status_array(2) = Time
    status_array(0) = Int((current_val / max_val) * 100)
    Application.StatusBar = "Processing..." & status_array(0) & "%" & _
        "   Average % time - " & Round(status_array(1) * 3600 * 24, 2) & " Seconds    Finish time - " & _
        (100 - status_array(0)) * status_array(1) + Time
    status_bar = status_array()
    
End If
End Function



