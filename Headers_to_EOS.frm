VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Headers_to_EOS 
   Caption         =   "Merge Matching Headers to End of Sheet"
   ClientHeight    =   5385
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

file_list = file_manager_guis.list_to_array(list_Files)

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























