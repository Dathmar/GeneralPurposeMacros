Attribute VB_Name = "Search_Form_Functions"
Option Explicit
Sub Text_Coloring()
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''                                This Macro was written by                                '''
'''                                      Asher Danner                                       '''
'''                                       05/09/2012                                        '''
'''The purpose is to launch the Font_Properties UserForm which is used for a veriety of     '''
'''tasks.                                                                                   '''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
If Not ActiveWorkbook Is Nothing Then
Font_Properties.Show
End If
End Sub
Function search_initialize(font_name As String, font_size As Long, _
                           font_color As Long, font_bold As Boolean, _
                           font_ital As Boolean, font_under As Long, _
                           font_strike As Boolean, font_super As Boolean, _
                           font_sub As Boolean, search_pref As String, _
                           search_1 As String, search_2 As String, _
                           include_1 As Boolean, include_2 As Boolean, _
                           search_area As String)
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''                                This Macro was written by                                '''
'''                                      Asher Danner                                       '''
'''                                       05/09/2012                                        '''
'''The purpose is to use the User's submissions in the Font_Properties UserForm to format   '''
'''text in the correct ways.                                                                '''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Dim rng As Range
Dim n As Long
Dim cStart As Integer
Dim cLen As Integer
Dim rCell As Range
Dim do_format As Boolean
Application.ScreenUpdating = False
n = 1
If search_area = "selection" Then
    Set rng = Selection
ElseIf search_area = "sheet" Then
    Set rng = ActiveSheet.UsedRange
Else
    n = ActiveWorkbook.Sheets.Count
End If

Do While n <> 0
If search_area = "workbook" Then
    Set rng = ActiveWorkbook.Sheets(n).UsedRange
    Sheets(n).Select
End If

For Each rCell In ActiveSheet.UsedRange
    cStart = InStr(rCell.value, search_1)
    do_format = False
    If cStart <> 0 Then ' only process relavent cells
        If search_pref = "before" Then
            If include_1 = True Then
                cLen = cStart + Len(search_1) - 1
            Else
                cLen = cStart - 1
            End If
            cStart = 1
            do_format = True
        ElseIf search_pref = "after" Then
            If include_1 = True Then
                cLen = Len(rCell.value)
            Else
                cStart = cStart + Len(search_1)
                cLen = Len(rCell.value)
            End If
            do_format = True
        ElseIf search_pref = "only" Then
            cLen = Len(search_1)
            do_format = True
        ElseIf search_pref = "between" And InStr(rCell.value, search_2) > cStart Then
            If include_1 = True And include_2 = True Then
                cLen = InStr(rCell.value, search_2) + Len(search_2) - cStart
            ElseIf include_1 = False And include_2 = True Then
                cStart = cStart + Len(search_1)
                cLen = InStr(rCell.value, search_2) + Len(search_2) - cStart
            ElseIf include_1 = False And include_2 = False Then
                cLen = InStr(rCell.value, search_2) - cStart - 1
                cStart = cStart + Len(search_1)
            Else ' include_1 = true and include_2 = false
                cLen = InStr(rCell.value, search_2) - cStart
            End If
            do_format = True
        ElseIf search_pref = "before and after" And InStr(rCell.value, search_2) > cStart Then
            If include_1 = True Then
                cLen = cStart + Len(search_1) - 1
            Else
                cLen = cStart - 1
            End If
            cStart = 1
            With rCell.Characters(cStart, cLen).Font
                .Name = font_name
                .Size = font_size
                .Color = font_color
                .Bold = font_bold
                .Italic = font_ital
                .Underline = font_under
                .Strikethrough = font_strike
                .Superscript = font_super
                .Subscript = font_sub
            End With
            cStart = InStr(rCell.value, search_2)
            If include_2 = True Then
                cLen = Len(rCell.value)
            Else
                cStart = cStart + Len(search_2)
                cLen = Len(rCell.value)
            End If
            do_format = True
        End If
        If do_format = True Then
            With rCell.Characters(cStart, cLen).Font
                .Name = font_name
                .Size = font_size
                .Color = font_color
                .Bold = font_bold
                .Italic = font_ital
                .Underline = font_under
                .Strikethrough = font_strike
                .Superscript = font_super
                .Subscript = font_sub
            End With
        End If
    End If
Next rCell
n = n - 1
Loop
Application.ScreenUpdating = True
End Function
Function installed_fonts() As Variant
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''                                This Macro was written by                                '''
'''                                      Asher Danner                                       '''
'''                                       05/09/2012                                        '''
'''The purpose is to return an array of all installed system fonts avalible to Excel.       '''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Dim FontList As CommandBarComboBox
Dim i As Long
Dim tempbar As CommandBar
Dim arr As Variant

On Error Resume Next
Set FontList = Application.CommandBars("Formatting").FindControl(ID:=1728)

If FontList Is Nothing Then
    Set tempbar = Application.CommandBars.Add
    Set FontList = tempbar.Controls.Add(ID:=1728)
End If

ReDim arr(1 To FontList.ListCount)
On Error GoTo 0
For i = 1 To FontList.ListCount - 1
arr(i) = FontList.List(i)
Next i
installed_fonts = arr
' Delete temp CommandBar if it exists
On Error Resume Next
tempbar.Delete
End Function
