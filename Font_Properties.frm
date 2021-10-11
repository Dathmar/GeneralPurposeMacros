VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Font_Properties 
   Caption         =   "Search"
   ClientHeight    =   5835
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7485
   OleObjectBlob   =   "Font_Properties.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Font_Properties"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cencel_button_Click()
Unload Me
End Sub
Private Sub color_black_Click()
sample_label.ForeColor = &H0&
End Sub
Private Sub color_blue_Click()
sample_label.ForeColor = &HFF0000
End Sub
Private Sub color_green_Click()
sample_label.ForeColor = 49152
End Sub
Private Sub color_green_Enter()
sample_label.ForeColor = 49152
End Sub
Private Sub color_orange_Click()
sample_label.ForeColor = &H80FF&
End Sub
Private Sub color_purple_Click()
sample_label.ForeColor = &HC000C0
End Sub
Private Sub color_red_Click()
sample_label.ForeColor = &HFF&
End Sub
Private Sub color_yellow_Click()
sample_label.ForeColor = &HFFFF&
End Sub
Private Sub color_black_Enter()
sample_label.ForeColor = &H0&
End Sub
Private Sub color_blue_Enter()
sample_label.ForeColor = &HFF0000
End Sub
Private Sub color_orange_Enter()
sample_label.ForeColor = &H80FF&
End Sub
Private Sub color_purple_Enter()
sample_label.ForeColor = &HC000C0
End Sub
Private Sub color_red_Enter()
sample_label.ForeColor = &HFF&
End Sub
Private Sub color_yellow_Enter()
sample_label.ForeColor = &HFFFF&
End Sub
Private Sub font_bold_Change()
sample_label.Font.Bold = font_bold.value
End Sub
Private Sub font_ital_Change()
sample_label.Font.Italic = font_ital.value
End Sub
Private Sub font_list_box_Change()
If font_list_box.value <> "" Then
    font_list_box.Font = font_list_box.value
    sample_label.Font = font_list_box.value
End If
End Sub
Private Sub font_list_box_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
If font_list_box.value <> "" Then
    font_list_box.Font = font_list_box.value
    sample_label.Font = font_list_box.value
End If
End Sub
Private Sub font_size_Change()
If font_size.value >= 1 And font_size.value <= 409 And IsNumeric(font_size.value) = True Then
    sample_label.Font.Size = font_size.value
End If
End Sub
Private Sub only_search_Change()
include_1.Visible = False
include_2.Visible = False
search_2.Visible = False
search_2.value = ""
include_1.value = False
include_2.value = False
End Sub
Private Sub before_search_Change()
include_1.Visible = True
include_2.Visible = False
search_2.Visible = False
search_2.value = ""
include_2.value = False
End Sub
Private Sub after_search_Change()
include_1.Visible = True
include_2.Visible = False
search_2.Visible = False
search_2.value = ""
include_2.value = False
End Sub
Private Sub before_and_after_search_Change()
include_1.Visible = True
include_2.Visible = True
search_2.Visible = True
End Sub
Private Sub between_search_Change()
include_1.Visible = True
include_2.Visible = True
search_2.Visible = True
End Sub
Private Sub submit_button_Click()
If font_list_box.value = "" Then
    MsgBox "Please select a font."
    Exit Sub
End If
If font_size.value = "" Or font_size.value < 1 Or font_size.value > 409 Then
    MsgBox "Please use a valid font size."
    Exit Sub
End If
If search_2.Visible = True And search_2.value = "" Then
    MsgBox "Please include a search term."
    Exit Sub
End If
If search_1.Visible = True And search_1.value = "" Then
    MsgBox "Please include a search term."
    Exit Sub
End If

Call search_initialize(font_list_box.value, font_size.value, _
                       sample_label.ForeColor, font_bold.value, _
                       font_ital.value, underline_check, _
                       font_strike.value, font_super.value, _
                       font_sub.value, search_preference, _
                       search_1.value, search_2.value, _
                       include_1.value, include_2.value, _
                       search_area)
Unload Me
End Sub
Private Sub UserForm_Initialize()
Dim font_list As Variant

'fill font list combobox
font_list = installed_fonts()
font_size.Font.Size = 12
font_size.value = 12
font_list_box.Font.Size = 12
For n = LBound(font_list) To UBound(font_list)
    If font_list(n) <> "" Then
        font_list_box.AddItem font_list(n)
    End If
Next n
search_1.Font.Size = 12
search_2.Font.Size = 12
font_list_box.value = "Arial"
und_none.value = True
before_search.value = True
sheet_radio = True
End Sub
Private Function underline_check() As Long
If und_none = True Then
    underline_check = xlUnderlineStyleNone
ElseIf und_single = True Then
    underline_check = xlUnderlineStyleSingle
ElseIf und_double = True Then
    underline_check = xlUnderlineStyleDouble
Else
    underline_check = xlUnderlineStyleNone
End If
End Function
Private Function search_preference() As String
If before_search.value = True Then
    search_preference = "before"
ElseIf after_search.value = True Then
    search_preference = "after"
ElseIf only_search.value = True Then
    search_preference = "only"
ElseIf between_search.value = True Then
    search_preference = "between"
ElseIf before_and_after_search.value = True Then
    search_preference = "before and after"
Else
    search_preference = "only"
End If
End Function
Private Function search_area() As String
If selection_radio.value = True Then
    search_area = "selection"
ElseIf workbook_radio.value = True Then
    search_area = "workbook"
ElseIf sheet_radio.value = True Then
    search_area = "sheet"
Else
    search_area = "sheet"
End If
End Function
