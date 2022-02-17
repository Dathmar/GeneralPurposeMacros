VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Split_by_equal_rows 
   Caption         =   "Split By Equal Number of Rows"
   ClientHeight    =   3645
   ClientLeft      =   150
   ClientTop       =   585
   ClientWidth     =   3900
   OleObjectBlob   =   "Split_by_equal_rows.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Split_by_equal_rows"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim wb As Workbook
Private Sub btn_folder_select_Click()
file_path_input.value = browse_for_folder()
End Sub
Private Sub btn_ok_Click()
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''                                This Macro was written by                                '''
'''                                      Asher Danner                                       '''
'''                                       04/22/2021                                        '''
'''The purpose is split a worksheet into many books based on number of user provided row    '''
'''count                                                                                    '''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Dim this_sht As Worksheet
Dim error_list As String
Dim start_row As Integer

Set wb = ActiveWorkbook
Set this_sht = wb.ActiveSheet

Application.ScreenUpdating = False

error_list = ""

If Not IsNumeric(start_row_input) Then
    error_list = append_text(error_list, "Please provide a start row.", Chr(13))
End If
If Not IsNumeric(number_of_rows_input) Then
    error_list = append_text(error_list, "Please provide a number of rows.", Chr(13))
End If

If Len(Dir(file_path_input, vbDirectory)) = 0 Then
    error_list = append_text(error_list, "Please provide a valid output folder that exists.", Chr(13))
End If

If error_list <> "" Then
    MsgBox (error_list)
    Exit Sub
End If

Me.Hide
start_row = start_row_input.value
If all_shts_cb.value Then
    For n = find_sht_index(start_sht_input.value) To wb.Sheets.Count
        Set this_sht = wb.Sheets(n)
        Run split_rows(this_sht, start_row)
        start_row = 2
    Next n
Else
    Set this_sht = ActiveSheet
    Run split_rows(this_sht, start_row)
End If

Application.ScreenUpdating = True
End Sub
Private Function split_rows(this_sht As Worksheet, start_row As Integer)
Dim end_row As Integer
Dim new_sht As Worksheet
Dim lr As Integer
Dim number_of_rows As String
Dim file_path As String
Dim new_book As Workbook
Dim save_path As String
Dim xls As Object

number_of_rows = number_of_rows_input.value

file_path = file_path_input.value

lr = this_sht.UsedRange.Rows.Count
While start_row < lr
    end_row = start_row + number_of_rows - 1
    If end_row > lr Then end_row = lr
    
    Application.Workbooks.Add
    Set new_book = ActiveWorkbook
    Set new_sht = new_book.Sheets(1)
    this_sht.Rows(1).Copy Destination:=new_sht.Rows(1)
    this_sht.Range(this_sht.Rows(start_row).EntireRow, this_sht.Rows(end_row).EntireRow).Copy Destination:=new_sht.Rows(2)
    
    save_path = file_path & "/" & get_name_without_extension(wb.name) & "_" & this_sht.name & "_Rows_" & start_row & "-" & end_row & ".xlsx"
    
    new_book.SaveAs filename:=save_path, FileFormat:=xlOpenXMLWorkbook
    new_book.Close SaveChanges:=False
    
    While IsWorkBookOpen(save_path)
        Set xls = GetObject(save_path)
        xls.Close True
    Wend
    
    start_row = end_row + 1
Wend
End Function
Private Function find_sht_index(sht_indentifier As String) As Long
Dim n As Integer

For n = 1 To wb.Sheets.Count
    If wb.Sheets(n).name = sht_indentifier Then
        find_sht_index = n
        Exit Function
    End If
Next n

For n = 1 To wb.Sheets.Count
    If n = Int(sht_indentifier) Then
        find_sht_index = n
        Exit Function
    End If
Next n

find_sht_index = -1
End Function
Private Function IsWorkBookOpen(filename As String)
    Dim ff As Long, ErrNo As Long

    On Error Resume Next
    ff = FreeFile()
    Open filename For Input Lock Read As #ff
    Close ff
    ErrNo = Err
    On Error GoTo 0

    Select Case ErrNo
    Case 0:    IsWorkBookOpen = False
    Case 70:   IsWorkBookOpen = True
    Case Else: Error ErrNo
    End Select
End Function
