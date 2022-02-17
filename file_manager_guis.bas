Attribute VB_Name = "file_manager_guis"
Option Explicit
Function array_to_list(aArray As Variant, list_to_add As Variant)
Dim n As Integer
Dim blank_list As Boolean
Dim l_bound As Integer
Dim xl_path As String
Dim xl_name As String
Dim xl_type As String

l_bound = LBound(aArray)

If list_to_add.ListCount = 0 Then blank_list = True
For n = l_bound To UBound(aArray)
    ' skip checking element match if the list is blank
    ' this speeds up operation of the program
    If Not blank_list Then
        If list_elmt_match(CStr(aArray(n)), list_to_add.List) Then GoTo continueloop
    End If
    xl_path = Left(aArray(n), InStrRev(aArray(n), "\")) 'path
    xl_name = Right(aArray(n), Len(aArray(n)) - (InStrRev(aArray(n), "\"))) ' file name
    xl_name = Left(xl_name, InStrRev(xl_name, ".") - 1)
    xl_type = Right(aArray(n), Len(aArray(n)) - (InStrRev(aArray(n), "."))) ' type
    
    list_to_add.AddItem
    list_to_add.List(n - l_bound, 0) = xl_path
    list_to_add.List(n - l_bound, 1) = xl_name
    list_to_add.List(n - l_bound, 2) = xl_type
continueloop:
Next n
End Function
Function list_elmt_match(elmt As String, match_array As Variant) As Boolean
Dim n As Long

For n = LBound(match_array, 1) To UBound(match_array, 1)
If elmt = CStr(match_array(n, 0) & match_array(n, 1) & "." & match_array(n, 2)) Then
    list_elmt_match = True
    Exit Function
End If
Next n
list_elmt_match = False
End Function
Function list_to_array(list_Files As Variant) As Variant
Dim aArray As Variant
Dim n As Long

If list_Files.ListCount = 0 Then Exit Function
ReDim aArray(LBound(list_Files.List, 1) To UBound(list_Files.List, 1)) As Variant
For n = LBound(list_Files.List, 1) To UBound(list_Files.List, 1)
    aArray(n) = CStr(list_Files.List(n, 0) & list_Files.List(n, 1) & "." & list_Files.List(n, 2))
Next n
list_to_array = aArray
End Function
Function unselected_items(list_box As Variant) As Variant
Dim aArray() As Integer
Dim n As Long
Dim i As Long
i = 1
For n = LBound(list_box.List, 1) To UBound(list_box.List, 1)
    If list_box.Selected(n) = False Then
        ReDim Preserve aArray(1 To i) As Integer
        aArray(i) = CInt(n)
        i = i + 1
    End If
Next n
If i = 0 Then
  unselected_items = Empty
Else
 unselected_items = aArray
End If
End Function
Function selected_items(list_Files As Variant) As Variant
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
Function remove_selected(list_Files As Variant) As Variant
Dim aArray As Variant
Dim unslc_items As Variant
Dim list_array As Variant
Dim n As Long
Dim i As Long

If list_Files.ListCount = 0 Then Exit Function
unslc_items = unselected_items(list_Files)
If Not IsArray(unslc_items) Then Exit Function

aArray = list_to_array(list_Files)
ReDim list_array(LBound(unslc_items) - 1 To UBound(unslc_items) - 1)
For n = LBound(unslc_items) To UBound(unslc_items)
    list_array(n - 1) = aArray(unslc_items(n))
Next n

list_Files.Clear
Call array_to_list(list_array, list_Files)
End Function
Function swap_elmt(frm_indx As Integer, to_indx As Integer, aArray As Variant) As Variant
Dim frm_elmt As String
Dim to_elmt As String

'store store elmts as a string
frm_elmt = aArray(frm_indx)
to_elmt = aArray(to_indx)

'write over changed elmts
aArray(frm_indx) = to_elmt
aArray(to_indx) = frm_elmt

swap_elmt = aArray
End Function
Function move_selected_to_top(list_Files As Variant)
Dim aArray As Variant
Dim slc_items As Variant
Dim i As Integer
Dim n As Long

If list_Files.ListCount = 0 Then Exit Function
slc_items = selected_items(list_Files)
If Not IsArray(slc_items) Then Exit Function

aArray = list_to_array(list_Files)
For n = UBound(slc_items) To LBound(slc_items) Step -1
    slc_items(n) = slc_items(n) + UBound(slc_items) - n
    If slc_items(n) <> LBound(list_Files.List, 1) Then
        For i = slc_items(n) To 1 Step -1
            aArray = swap_elmt(i, i - 1, aArray)
        Next i
    End If
Next n
list_Files.Clear
Call array_to_list(aArray, list_Files)
For n = 0 To UBound(slc_items) - 1
    list_Files.Selected(n) = True
Next n
End Function
Function move_selected_up(list_Files As Variant)
Dim aArray As Variant
Dim slc_items As Variant
Dim n As Long

If list_Files.ListCount = 0 Then Exit Function
slc_items = selected_items(list_Files)
If Not IsArray(slc_items) Then Exit Function
aArray = list_to_array(list_Files)
For n = LBound(slc_items) To UBound(slc_items)
    If slc_items(n) <> LBound(list_Files.List, 1) Then aArray = swap_elmt(CInt(slc_items(n)), CInt(slc_items(n) - 1), aArray)
Next n
list_Files.Clear
Call array_to_list(aArray, list_Files)
For n = LBound(slc_items) To UBound(slc_items)
    If slc_items(n) <> LBound(list_Files.List, 1) Then
        list_Files.Selected(slc_items(n) - 1) = True
    Else
        list_Files.Selected(slc_items(n)) = True
    End If
Next n
End Function
Function move_selected_down(list_Files As Variant)
Dim aArray As Variant
Dim slc_items As Variant
Dim n As Long

If list_Files.ListCount = 0 Then Exit Function
slc_items = selected_items(list_Files)
If Not IsArray(slc_items) Then Exit Function
aArray = list_to_array(list_Files)
For n = LBound(slc_items) To UBound(slc_items)
    If slc_items(n) <> UBound(list_Files.List, 1) Then aArray = swap_elmt(CInt(slc_items(n)), CInt(slc_items(n) + 1), aArray)
Next n
list_Files.Clear
Call array_to_list(aArray, list_Files)
For n = LBound(slc_items) To UBound(slc_items)
    If slc_items(n) <> UBound(list_Files.List, 1) Then
        list_Files.Selected(slc_items(n) + 1) = True
    Else
        list_Files.Selected(slc_items(n)) = True
    End If
Next n
End Function
Function move_selected_to_bottom(list_Files As Variant)
Dim aArray As Variant
Dim slc_items As Variant
Dim i As Integer
Dim n As Long

If list_Files.ListCount = 0 Then Exit Function
slc_items = selected_items(list_Files)
If Not IsArray(slc_items) Then Exit Function
aArray = list_to_array(list_Files)
For n = UBound(slc_items) To LBound(slc_items) Step -1
    If slc_items(n) <> UBound(list_Files.List, 1) Then
        For i = slc_items(n) To UBound(list_Files.List, 1) - 1 - (UBound(slc_items) - n)
            aArray = swap_elmt(i, i + 1, aArray)
        Next i
    End If
Next n
list_Files.Clear
Call array_to_list(aArray, list_Files)
For n = 0 To UBound(slc_items) - 1
    list_Files.Selected(list_Files.ListCount - n - 1) = True
Next n
End Function
Function add_files_to_list(list_Files As Variant)
Dim n As Long
Dim i As Long
Dim aArray As Variant
Dim xl_files As Variant
Dim xl_file As Long
Dim cnt As Long


xl_files = Application.GetOpenFilename("Excel files (*.xl*),*.xl*", , _
    "Browse for file to be merged", MultiSelect:=True)
If Not IsArray(xl_files) Then Exit Function
aArray = list_to_array(list_Files)
If Not IsArray(aArray) Then
    ReDim aArray(0 To 0)
    cnt = 0
Else
    cnt = UBound(aArray) + 1
End If
For xl_file = LBound(xl_files) To UBound(xl_files)

    ReDim Preserve aArray(0 To cnt)
    aArray(cnt) = xl_files(xl_file)
    cnt = cnt + 1

Next xl_file

Call array_to_list(aArray, list_Files)
End Function
Function clear_list(list_Files As Variant)
Dim n As Long
For n = 0 To list_Files.ListCount - 1
    list_Files.Selected(n) = False
Next n
End Function
