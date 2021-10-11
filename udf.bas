Attribute VB_Name = "udf"
Option Explicit
Function concatgpm(rng As Range, Optional delim As String = ",", Optional include_blanks As Boolean = False, Optional unique As Boolean = False) As String
Dim rCell As Range
Dim ret_val As String

For Each rCell In rng.Cells
    If rCell = "" And include_blanks Then
        If ret_val = "" Then
            ret_val = rCell
        Else
            If Not unique Or (unique And InStr(ret_val, rCell.value) = 0) Or rCell.value = "" Then
                ret_val = ret_val & delim & rCell
            End If
        End If
    ElseIf rCell.Row = rng.Parent.UsedRange.Rows.Count + 1 Then ' Stop execution at last row with data.
        concatgpm = ret_val
        Exit Function
    ElseIf rCell <> "" Then
        If ret_val = "" Then
            ret_val = rCell
        Else
            If Not unique Or (unique And InStr(ret_val, rCell.value) = 0) Then
                ret_val = ret_val & delim & rCell
            End If
        End If
    End If
Next rCell
concatgpm = ret_val
End Function
Function concatif(if_rng As Range, if_con As String, concat_rng As Range, Optional delim As String = ",", Optional include_blanks As Boolean = False, Optional unique As Boolean = False)
Dim rCell As Range
Dim ret_val As String
Dim sht As Worksheet

Set sht = if_rng.Parent


For Each rCell In if_rng.Cells
    If CStr(rCell) = if_con Then
        If sht.Cells(rCell.Row, concat_rng.Column) <> "" And include_blanks Then
            If ret_val = "" Then
                ret_val = sht.Cells(rCell.Row, concat_rng.Column)
            Else
                If Not unique Or (unique And InStr(ret_val, sht.Cells(rCell.Row, concat_rng.Column).value) = 0) Then
                    ret_val = ret_val & delim & sht.Cells(rCell.Row, concat_rng.Column)
                End If
            End If
        ElseIf sht.Cells(rCell.Row, concat_rng.Column).Row = if_rng.Parent.UsedRange.Rows.Count + 1 Then
            concatif = ret_val
            Exit Function
        ElseIf sht.Cells(rCell.Row, concat_rng.Column) <> "" Then
            If ret_val = "" Then
                ret_val = sht.Cells(rCell.Row, concat_rng.Column)
            Else
                If Not unique Or (unique And InStr(ret_val, sht.Cells(rCell.Row, concat_rng.Column).value) = 0) Then
                    ret_val = ret_val & delim & sht.Cells(rCell.Row, concat_rng.Column)
                End If
            End If
        End If
    End If
Next rCell
concatif = ret_val


End Function
Function left_before(rng As Variant, before As String, Optional trim_str As Boolean = False) As String
Dim ret As String
Dim txt As String

txt = CStr(rng)

If InStr(txt, before) = 0 Then
    left_before = txt
    Exit Function
End If

ret = Left(txt, InStr(txt, before) - 1)

If trim_str Then
    ret = Trim(ret)
End If
left_before = ret
End Function

Function right_after(rng As Variant, after As String, Optional trim_str As Boolean = False) As String
Dim ret As String
Dim txt As String

txt = CStr(rng)

If InStr(txt, after) = 0 Then
    right_after = txt
    Exit Function
End If

ret = Right(txt, Len(txt) - InStr(txt, after) - Len(after) + 1)

If trim_str Then
    ret = Trim(ret)
End If
right_after = ret
End Function
Function mid_between(rng As Variant, first As String, second As String, Optional trim_str As Boolean = False) As String
Dim ret As String
Dim txt As String

txt = CStr(rng)

ret = Mid(txt, InStr(txt, first) + Len(first), InStr(txt, second) - InStr(txt, first) - Len(first))

If trim_str Then
    ret = Trim(ret)
End If
mid_between = ret

End Function
Function remove_extra_spaces(rng As Range) As String

Dim txt As String

txt = rng.Value2

txt = Replace(txt, " ", " ")

If Not txt <> "" Then
    Do While InStr(txt, "  ") <> 0
        txt = Replace(txt, "  ", " ")
    Loop
End If

remove_extra_spaces = Trim(txt)

End Function
Function append_text(append_to As Variant, append_val As String, Optional del As String = ",") As String
Dim ret_val As String

append_to = CStr(append_to)

If append_to = "" Then
    ret_val = append_val
Else
    ret_val = append_to & del & append_val
End If
append_text = ret_val
End Function





