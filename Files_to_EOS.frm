VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Files_to_EOS 
   Caption         =   "Merge Options"
   ClientHeight    =   5685
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   10335
   OleObjectBlob   =   "Files_to_EOS.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Files_to_EOS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub active_sheet_Click()
sht_no.Enabled = False
End Sub
Private Sub all_sheets_Click()
sht_no.Enabled = False
End Sub
Private Sub Cancel_Click()
Unload Me
End Sub
Private Sub OK_Click()
Dim from_book As Workbook
Dim to_book As Workbook
Dim to_sht As Worksheet
Dim xl_file_name As Variant
Dim this_row As Long
Dim this_workbook As Long
Dim last_row As Long
Dim include_fn As Integer
Dim start_column As Integer
Dim sheet_count As Integer
Dim this_sheet As Integer
Dim n As Long
Dim start_sheet As Integer
Dim already_open As Boolean
Dim skip_merge As Boolean
Dim rpt_sht As Worksheet
Dim merge_cnt As Long
Dim from_sht As Worksheet
Dim merge_header As Boolean
Dim cpy_rng As Range

Me.Hide

xl_file_name = file_manager_guis.list_to_array(list_Files)
If Not IsArray(xl_file_name) Then
    If xl_file_name = False Then Exit Sub
End If
Application.DisplayAlerts = False
Application.ScreenUpdating = False
Application.AskToUpdateLinks = False
If ActiveWorkbook Is Nothing Then Application.Workbooks.Add

start_column = 1
If filename.value Then
    start_column = start_column + 1
End If
If sheetname.value Then
    start_column = start_column + 1
End If
Set to_book = ActiveWorkbook
Set to_sht = to_book.ActiveSheet
Call delete_extraneous_blank_rows_and_columns(to_sht)

to_book.Worksheets.Add after:=to_book.Sheets(to_book.Sheets.Count)
Set rpt_sht = to_book.Sheets(to_book.Sheets.Count)
rpt_sht.Activate
rpt_sht.name = format_sheet_name("Summary", to_book)

' delete leading blank rows Need to create function to do this.
Do While Application.WorksheetFunction.CountA(to_sht.Rows(1)) = 0 And to_sht.UsedRange.Rows.Count <> 0 And to_sht.UsedRange.Rows.Count <> 1
    to_sht.Rows(1).EntireRow.Delete
Loop
Do While Application.WorksheetFunction.CountA(to_sht.Columns(1)) = 0 And to_sht.UsedRange.Columns.Count <> 0 And to_sht.UsedRange.Columns.Count <> 1
    to_sht.Columns(1).EntireColumn.Delete
Loop

last_row = to_sht.UsedRange.Rows.Count + 1

If to_sht.UsedRange.Columns.Count = 1 And to_sht.UsedRange.Rows.Count = 1 And to_sht.Cells(1, 1) = "" Then last_row = 1
merge_cnt = 2
If IsArray(xl_file_name) Then

    For this_workbook = LBound(xl_file_name) To UBound(xl_file_name)
        rpt_sht.Cells(1, 1) = "Workbook"
        rpt_sht.Cells(1, 2) = "Worksheet"
        rpt_sht.Cells(1, 3) = "Merge Status"
        rpt_sht.Cells(1, 4) = "In rows"
        rpt_sht.Cells(merge_cnt, 1) = get_filename(CStr(xl_file_name(this_workbook))) ' workbook
        rpt_sht.Cells(merge_cnt, 2) = "Not Started" ' worksheet
        rpt_sht.Cells(merge_cnt, 3) = "Not Started" ' merge status
        rpt_sht.Cells(merge_cnt, 4) = "Unknown" ' in rows
    Next this_workbook

    For this_workbook = LBound(xl_file_name) To UBound(xl_file_name) ' iterate through each book
        
        ' handle merging headers.
        If Me.merge_no_header Then
            merge_header = False
        ElseIf Me.merge_all_headers Then
            merge_header = True
        ElseIf Me.merge_first_header And this_workbook = 1 Then
            merge_header = True
        Else
            merge_header = False
        End If
            
        rpt_sht.Cells(merge_cnt, 2) = "Working"
        ' check to see if the workbook is already open
        ' if it isn't open then open the workbook
        If is_workbook_open(get_filename(xl_file_name(this_workbook))) Then
            Set from_book = Workbooks(get_filename(CStr(xl_file_name(this_workbook))))
            already_open = True
        Else
            Application.Workbooks.Open filename:=xl_file_name(this_workbook), ReadOnly:=True
            Set from_book = ActiveWorkbook
        End If
        
        If all_sheets Then
            sheet_count = from_book.Sheets.Count
            start_sheet = 1
        ElseIf spec_sheet Then
            On Error Resume Next
            Set from_sht = from_book.Sheets(CStr(sht_no))
            On Error GoTo 0
            If IsNumeric(sht_no) And Not from_sht Is Nothing Then
                sheet_count = sht_no
                start_sheet = sht_no
            ElseIf IsNumeric(sht_no) Then
                If sht_no <= from_book.Sheets.Count Then
                    sheet_count = from_book.Sheets(sht_no).Index
                    start_sheet = from_book.Sheets(sht_no).Index
                ElseIf LenB(Dir("C:\users\dbaum\")) Then
                    skip_merge = True
                    rpt_sht.Cells(merge_cnt, 1) = get_filename(CStr(xl_file_name(this_workbook))) ' worksheet
                    rpt_sht.Cells(merge_cnt, 2) = sht_no
                    rpt_sht.Cells(merge_cnt, 3) = "Did you check sheets vis-" & Chr(224) & "-vis acceptible sheets in workbook." ' merge status
                    rpt_sht.Cells(merge_cnt, 4) = "NA" ' in rows
                Else
                    skip_merge = True
                    rpt_sht.Cells(merge_cnt, 1) = get_filename(CStr(xl_file_name(this_workbook)))
                    rpt_sht.Cells(merge_cnt, 2) = sht_no ' worksheet
                    rpt_sht.Cells(merge_cnt, 3) = "Sheet not found" ' merge status
                    rpt_sht.Cells(merge_cnt, 4) = "NA" ' in rows
                End If
            Else
                sheet_count = from_book.ActiveSheet.Index
                start_sheet = from_book.ActiveSheet.Index
            End If
        Else
            sheet_count = from_book.ActiveSheet.Index
            start_sheet = from_book.ActiveSheet.Index
        End If
        
        For this_sheet = start_sheet To sheet_count
            'If from_book.Sheets(this_sheet) = rpt_sht Then
            '    skip_merge = True
            'End If
            
            ' Do not merge the sheet you are merging into.
            If this_sheet = to_sht.Index And to_book.name = from_book.name And skip_merge = False Then
                skip_merge = True
            End If
            
            If Not skip_merge Then
                With from_book.Sheets(this_sheet)
                If Application.WorksheetFunction.CountA(.Cells) <> 0 Then ' only do this if the sheet is not empty
                    If .FilterMode Then .ShowAllData
                    ' delete leading blank rows
                    Do While Application.WorksheetFunction.CountA(.Rows(1)) = 0 And .UsedRange.Rows.Count <> 0 And .UsedRange.Rows.Count <> 1
                        .Rows(1).EntireRow.Delete
                    Loop
                    Do While Application.WorksheetFunction.CountA(.Columns(1)) = 0 And .UsedRange.Columns.Count <> 0 And .UsedRange.Columns.Count <> 1
                        .Columns(1).EntireColumn.Delete
                    Loop
                    Call delete_extraneous_blank_rows_and_columns(from_book.Sheets(this_sheet))
    
                    If to_sht.Rows.Count < last_row + .UsedRange.Rows.Count Then
                        rpt_sht.Cells(merge_cnt, 1) = get_filename(CStr(xl_file_name(this_workbook))) ' workbook
                        rpt_sht.Cells(merge_cnt, 2) = .name ' worksheet
                        rpt_sht.Cells(merge_cnt, 3) = "Ran out of rows"
                        rpt_sht.Cells(merge_cnt, 4) = "NA"
                        If Not already_open Then
                            from_book.Close
                        End If
                        Exit Sub
                        Application.AskToUpdateLinks = True
                        Application.DisplayAlerts = True
                        Application.ScreenUpdating = True
                    End If
                    
                    ' set copy range based on merge header information
                    If merge_header Then
                        ' if you are merging the header copy everything.
                        Set cpy_rng = .UsedRange
                        
                        ' after the first header is copied then don't merge any more if merge_first_header is set.
                        If Me.merge_first_header Then
                            merge_header = False
                        End If
                    Else
                        ' If there is only header then copy the first cell without the header
                        Debug.Print .UsedRange.Rows.Count
                        
                        If .UsedRange.Rows.Count <= CInt(Me.num_header_rows) Then
                            Set cpy_rng = .Cells(Me.num_header_rows + 1, 1)
                        Else
                            Set cpy_rng = .Range(.Cells(Me.num_header_rows + 1, 1), .Cells(.UsedRange.Rows.Count, .UsedRange.Columns.Count))
                        End If
                    End If
                    
                    If Me.paste_values Then
                        cpy_rng.Copy
                        to_sht.Cells(last_row, start_column).PasteSpecial xlPasteValues
                    Else
                        cpy_rng.Copy Destination:=to_sht.Cells(last_row, start_column)
                    End If
                    rpt_sht.Cells(merge_cnt, 1) = get_filename(CStr(xl_file_name(this_workbook))) ' workbook
                    rpt_sht.Cells(merge_cnt, 2) = .name ' worksheet
                    rpt_sht.Cells(merge_cnt, 3) = "Merged"
                    rpt_sht.Cells(merge_cnt, 4).NumberFormat = "@"
                    rpt_sht.Cells(merge_cnt, 4) = last_row & "-" & to_sht.UsedRange.Rows.Count
                    
                    If filename.value Then to_sht.Range(to_sht.Cells(last_row, 1), to_sht.Cells(to_sht.UsedRange.Rows.Count, 1)) = from_book.name
                    If sheetname.value Then to_sht.Range(to_sht.Cells(last_row, start_column - 1), to_sht.Cells(to_sht.UsedRange.Rows.Count, start_column - 1)) = .name
                Else
                    rpt_sht.Cells(merge_cnt, 1) = get_filename(CStr(xl_file_name(this_workbook)))
                    rpt_sht.Cells(merge_cnt, 2) = .name
                    rpt_sht.Cells(merge_cnt, 3) = "Sheet Blank"
                    rpt_sht.Cells(merge_cnt, 4) = "NA"
                End If
                End With
            Else
                rpt_sht.Cells(merge_cnt, 1) = get_filename(CStr(xl_file_name(this_workbook))) ' workbook
            End If
            last_row = to_sht.UsedRange.Rows.Count + 1
            skip_merge = False
            merge_cnt = merge_cnt + 1
        Next this_sheet
        
        If Not already_open Then
            from_book.Close
        End If
        Set from_book = Nothing
        rpt_sht.Cells(merge_cnt, 1) = "Complete"
    Next this_workbook
End If
to_sht.Activate
Unload Me
Application.AskToUpdateLinks = True
Application.DisplayAlerts = True
Application.ScreenUpdating = True
End Sub
Private Sub spec_sheet_Click()
sht_no.Enabled = True
End Sub
Private Sub add_files_Click()
Call file_manager_guis.add_files_to_list(list_Files)
End Sub
Private Sub clear_list_Click()
list_Files.Clear
End Sub
Private Sub deselect_button_Click()
Call file_manager_guis.clear_list(list_Files)
End Sub
Private Sub remove_button_Click()
Call file_manager_guis.remove_selected(list_Files)
End Sub
Private Sub top_button_Click()
Call file_manager_guis.move_selected_to_top(list_Files)
End Sub
Private Sub bottom_button_Click()
Call file_manager_guis.move_selected_to_bottom(list_Files)
End Sub
Private Sub up_button_Click()
Call file_manager_guis.move_selected_up(list_Files)
End Sub
Private Sub down_button_Click()
Call file_manager_guis.move_selected_down(list_Files)
End Sub
