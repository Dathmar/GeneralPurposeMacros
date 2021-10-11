VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Headers_to_EOS 
   Caption         =   "Merge Matching Headers to End of Sheet"
   ClientHeight    =   4575
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8310
   OleObjectBlob   =   "Headers_to_EOS.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Headers_to_EOS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub OK_Click()
Dim file_list() As Variant
Dim n As Integer
Dim this_book As Workbook
Dim merge_sht As Worksheet
Dim headers() As Variant
Dim i As Long
Dim add_headers As Variant
Dim this_headers() As Variant

ReDim file_list(0 To list_Files.ListCount) As Variant
file_list = list_to_array

Application.Workbooks.Add
Set merge_sht = ActiveSheet

For n = LBound(file_list) To UBound(file_list)
    Application.Workbooks.Open (file_list(n))
    Set this_book = ActiveWorkbook
    
    'get headers and add them to the headers variable
    If n = LBound(file_list) Then
        ReDim headers(1 To this_book.ActiveSheet.UsedRange.Columns.Count)
        headers() = get_headers(this_book.ActiveSheet)
        
        this_book.Sheets(1).UsedRange.Copy Destination:=merge_sht.Cells(1, 1)
    Else
        this_headers = get_headers(this_book.ActiveSheet)
        add_headers = new_headers(headers(), this_headers())
        ReDim Preserve headers(1 To UBound(headers) + UBound(add_headers))
        For i = 1 To UBound(add_headers)
            headers(UBound(headers) - UBound(add_headers) + i) = add_headers(i)
            this_book.Sheets(1).Cells(1, this_book.Sheets(1).UsedRange.Columns.Count + 1) = add_headers(i)
        Next i
        
        For i = 1 To mrg_sheet.UsedRange.Columns.Count
            If matching(headers(), this_book.Sheets(1).Cells(1, i)) Then
                
            End If
        Next i
    End If
    
Next n

End Sub
Function new_headers(headers() As Variant, mrg_headers() As Variant) As Variant
Dim n As Integer
Dim col_count As Long
Dim i As Long
Dim merge_headers() As Variant
i = 1
For n = LBound(mrg_headers) To UBound(mrg_headers)
    If mrg_headers(n) <> "" Then
        If Not matching(headers(), CStr(mrg_headers(n))) Then
            ReDim Preserve merge_headers(LBound(headers) To i)
            merge_headers(i) = mrg_headers(n)
            i = i + 1
        End If
    End If
Next n
new_headers = headers
End Function
Function matching(headers() As Variant, header As String) As Boolean
Dim n As Long
For n = LBound(headers) To UBound(headers)
    If headers(n) = header Then
        matching = True
        Exit Function
    End If
Next n
matching = False
End Function
Function get_headers(this_sht As Worksheet) As Variant
Dim n As Integer
Dim col_count As Long
Dim headers() As Variant

col_count = this_sht.UsedRange.Columns.Count
ReDim headers(1 To col_count)

For n = 1 To col_count
    headers(n) = this_sht.Cells(1, n)
Next n

get_headers = headers
End Function
Private Sub Cancel_Click()
Unload Me
End Sub
Private Sub add_files_Click()
Dim n As Long
Dim i As Long
Dim xl_files As Variant

xl_files = Application.GetOpenFilename("Excel files (*.xl*),*.xl*", , _
    "Browse for file to be merged", MultiSelect:=True)
If Not IsArray(xl_files) Then Exit Sub
Call array_to_list(xl_files)
End Sub
Function array_to_list(aArray As Variant)
Dim n As Integer
Dim blank_list As Boolean
Dim l_bound As Integer

l_bound = LBound(aArray)

If list_Files.ListCount = 0 Then blank_list = True
For n = l_bound To UBound(aArray)
    ' skip checking element match if the list is blank
    ' this speeds up operation of the program
    If Not blank_list Then
        If list_elmt_match(CStr(aArray(n)), list_Files.List) Then GoTo continueloop
    End If
    xl_path = Left(aArray(n), InStrRev(aArray(n), "\")) 'path
    xl_name = Right(aArray(n), Len(aArray(n)) - (InStrRev(aArray(n), "\")))
    xl_name = Left(xl_name, InStrRev(xl_name, ".") - 1)
    xl_type = Right(aArray(n), Len(aArray(n)) - (InStrRev(aArray(n), "."))) ' type
    
    list_Files.AddItem
    list_Files.List(n - l_bound, 0) = xl_path
    list_Files.List(n - l_bound, 1) = xl_name
    list_Files.List(n - l_bound, 2) = xl_type
continueloop:
Next n
End Function
Function list_to_array() As Variant
Dim aArray As Variant
If list_Files.ListCount = 0 Then Exit Function
ReDim aArray(LBound(list_Files.List, 1) To UBound(list_Files.List, 1)) As Variant
For n = LBound(list_Files.List, 1) To UBound(list_Files.List, 1)
    aArray(n) = CStr(list_Files.List(n, 0) & list_Files.List(n, 1) & "." & list_Files.List(n, 2))
Next n
list_to_array = aArray
End Function
Function list_elmt_match(elmt As String, match_array As Variant) As Boolean
Dim n As Long

For n = LBound(match_array, 1) To UBound(match_array, 1)
If elmt = CStr(match_array(n, 0) & match_array(n, 1) & "." & match_array(n, 2)) Then
    elmt_match = True
    Exit Function
End If
Next n
elmt_match = False
End Function
Private Sub clear_list_Click()
list_Files.Clear
End Sub
Private Sub deselect_button_Click()
Dim n As Long
For n = 0 To list_Files.ListCount - 1
    list_Files.Selected(n) = False
Next n
End Sub
Private Sub top_button_Click()
Dim aArray As Variant
Dim slc_items As Variant
Dim i As Integer

If list_Files.ListCount = 0 Then Exit Sub
slc_items = selected_items
If Not IsArray(slc_items) Then Exit Sub

aArray = list_to_array()
For n = UBound(slc_items) To LBound(slc_items) Step -1
    slc_items(n) = slc_items(n) + UBound(slc_items) - n
    If slc_items(n) <> LBound(list_Files.List, 1) Then
        For i = slc_items(n) To 1 Step -1
            aArray = swap_elmt(i, i - 1, aArray)
        Next i
    End If
Next n
list_Files.Clear
Call array_to_list(aArray)
For n = 0 To UBound(slc_items) - 1
    list_Files.Selected(n) = True
Next n
End Sub
Private Sub bottom_button_Click()
Dim aArray As Variant
Dim slc_items As Variant
Dim i As Integer

If list_Files.ListCount = 0 Then Exit Sub
slc_items = selected_items
If Not IsArray(slc_items) Then Exit Sub
aArray = list_to_array()
For n = UBound(slc_items) To LBound(slc_items) Step -1
    If slc_items(n) <> UBound(list_Files.List, 1) Then
        For i = slc_items(n) To UBound(list_Files.List, 1) - 1 - (UBound(slc_items) - n)
            aArray = swap_elmt(i, i + 1, aArray)
        Next i
    End If
Next n
list_Files.Clear
Call array_to_list(aArray)
For n = 0 To UBound(slc_items) - 1
    list_Files.Selected(list_Files.ListCount - n - 1) = True
Next n
End Sub
Private Sub up_button_Click()
Dim aArray As Variant
Dim slc_items As Variant
If list_Files.ListCount = 0 Then Exit Sub
slc_items = selected_items
If Not IsArray(slc_items) Then Exit Sub
aArray = list_to_array()
For n = LBound(slc_items) To UBound(slc_items)
    If slc_items(n) <> LBound(list_Files.List, 1) Then aArray = swap_elmt(CInt(slc_items(n)), CInt(slc_items(n) - 1), aArray)
Next n
list_Files.Clear
Call array_to_list(aArray)
For n = LBound(slc_items) To UBound(slc_items)
    If slc_items(n) <> LBound(list_Files.List, 1) Then
        list_Files.Selected(slc_items(n) - 1) = True
    Else
        list_Files.Selected(slc_items(n)) = True
    End If
Next n
End Sub
Private Sub down_button_Click()
Dim aArray As Variant
Dim slc_items As Variant
If list_Files.ListCount = 0 Then Exit Sub
slc_items = selected_items
If Not IsArray(slc_items) Then Exit Sub
aArray = list_to_array()
For n = LBound(slc_items) To UBound(slc_items)
    If slc_items(n) <> UBound(list_Files.List, 1) Then aArray = swap_elmt(CInt(slc_items(n)), CInt(slc_items(n) + 1), aArray)
Next n
list_Files.Clear
Call array_to_list(aArray)
For n = LBound(slc_items) To UBound(slc_items)
    If slc_items(n) <> UBound(list_Files.List, 1) Then
        list_Files.Selected(slc_items(n) + 1) = True
    Else
        list_Files.Selected(slc_items(n)) = True
    End If
Next n
End Sub
Function swap_elmt(frm_indx As Integer, to_indx As Integer, aArray As Variant) As Variant
Dim fmr_emlt As String
Dim to_elmt As String

'store store elmts as a string
frm_elmt = aArray(frm_indx)
to_elmt = aArray(to_indx)

'write over changed elmts
aArray(frm_indx) = to_elmt
aArray(to_indx) = frm_elmt

swap_elmt = aArray
End Function
Function selected_items() As Variant
Dim aArray() As Integer
Dim n As Long
Dim i As Long
i = 1
For n = LBound(list_Files.List, 1) To UBound(list_Files.List, 1)
    If list_Files.Selected(n) Then
        ReDim Preserve aArray(1 To i) As Integer
        aArray(i) = CInt(n)
        i = i + 1
    End If
Next n
If i = 1 Then
  selected_items = Empty
Else
 selected_items = aArray
End If
End Function




















