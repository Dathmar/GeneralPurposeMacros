VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} IBIS_Tally 
   Caption         =   "IBIS Tally Format"
   ClientHeight    =   4215
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4140
   OleObjectBlob   =   "IBIS_Tally.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "IBIS_Tally"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Function pre_fill(accept_col As String, accept_with_edit_col As String, reject_col As String, results_col As String, Optional my_title As String)

If my_title = "" Then
    my_title = "IBIS Tally Results"
End If

Me.accept.value = accept_col
Me.accept_with_edit.value = accept_with_edit_col
Me.reject.value = reject_col
Me.results.value = results_col

End Function
Private Sub go_Click()
Dim col_count As Long
Dim accept_col As Long, accept_with_edit_col As Long, reject_col As Long
Dim results_col As Long

Dim n As Long

Dim accept_val As Long
Dim awe_val As Long
Dim reject_val As Long
Application.ScreenUpdating = False
col_count = ActiveWorkbook.ActiveSheet.UsedRange.Rows.Count

accept_col = get_numeric_value(Me.accept.value)
accept_with_edit_col = get_numeric_value(Me.accept_with_edit.value)
reject_col = get_numeric_value(Me.reject.value)
results_col = get_numeric_value(Me.results.value)

For n = 1 To col_count
    accept_val = get_result(n, accept_col)
    awe_val = get_result(n, accept_with_edit_col)
    reject_val = get_result(n, reject_col)
    
    If accept_val <> 0 Or awe_val <> 0 Or reject_val <> 0 Then ' all have values
        If accept_val = awe_val Or accept_val = reject_val Or awe_val = reject_val Then
             ActiveWorkbook.ActiveSheet.Cells(n, results_col) = "Split Tally"
             ActiveWorkbook.ActiveSheet.Cells(n, results_col).Interior.Color = RGB(255, 255, 255)
        End If
        If accept_val > awe_val And accept_val > reject_val Then
            ActiveWorkbook.ActiveSheet.Cells(n, results_col) = "Accept as is"
        ElseIf awe_val > accept_val And awe_val > reject_val Then
            ActiveWorkbook.ActiveSheet.Cells(n, results_col) = "Accept with edit"
        ElseIf reject_val > accept_val And reject_val > awe_val Then
            ActiveWorkbook.ActiveSheet.Cells(n, results_col) = "Reject"
        End If
    End If
    
    If ActiveWorkbook.ActiveSheet.Cells(n, results_col) = "Split Tally" Then
        With ActiveWorkbook.ActiveSheet.Cells(n, results_col).Interior
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .Color = 255
            .TintAndShade = 0
            .PatternTintAndShade = 0
        End With
    End If
Next n

Application.ScreenUpdating = True
Unload Me
Exit Sub
Err_handl:
MsgBox "Must supply proper column."
Application.ScreenUpdating = True
Unload Me
End Sub
Private Function get_result(n As Long, col As Long) As Long
Dim result As String

If col = 0 Then
    get_result = 0
    Exit Function
End If

result = ActiveWorkbook.ActiveSheet.Cells(n, col)

If IsNumeric(result) Then
    get_result = result
Else
    get_result = 0
End If

End Function
Private Function get_numeric_value(value As String) As Long
On Error GoTo Err_handl
If IsNumeric(value) Then
    get_numeric_value = CLng(value)
ElseIf value = "" Then
    get_numeric_value = 0
Else
    get_numeric_value = ActiveWorkbook.Sheets(1).Columns(value).Column
End If
Exit Function
Err_handl:
get_numeric_value = 0
End Function
