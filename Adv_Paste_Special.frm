VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Adv_Paste_Special 
   Caption         =   "Advanced Paste"
   ClientHeight    =   4695
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4695
   OleObjectBlob   =   "Adv_Paste_Special.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Adv_Paste_Special"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cb_run_button_Click()
Call paste_spec_cur_bk(curBook_curSheet, paste_all.value, paste_formats.value, paste_formulas.value, paste_values.value, paste_widths.value)
Unload Me
End Sub
Private Sub nb_cancel_button_Click()
Unload Me
End Sub
Private Sub cb_cancel_button_Click()
Unload Me
End Sub
Private Sub folder_path_Change()
If Len(Dir(folder_path.value, vbDirectory)) = 0 Then
    folder_error.Caption = "Invalid Folder"
Else
    folder_error.Caption = ""
End If
End Sub
Private Sub nb_run_button_Click()

If folder_path.value = "" Then
    folder_error = "Please select a folder to save to."
    Exit Sub
End If

If folder_error.Caption = "Invalid Folder" Then
    folder_error = "Please select a valid folder to save to."
    Exit Sub
End If

Call paste_spec_split(folder_path.value, newBook_curSheet, paste_all.value, paste_formats.value, paste_formulas.value, paste_values.value, paste_widths.value, whole_book.value)
Unload Me
End Sub
Private Sub newBook_allSheets_Change()
If newBook_allSheets.value Then
    folder_path.Enabled = True
    select_folder.Enabled = True
Else
    folder_path.Enabled = False
    folder_path.value = ""
    
    select_folder.Enabled = False
End If
End Sub
Private Sub paste_all_Change()
If paste_all.value Then
    paste_formats.value = False
    paste_formats.Enabled = False
    
    paste_formulas.value = False
    paste_formulas.Enabled = False
    
    paste_values.value = False
    paste_values.Enabled = False
    
    paste_widths.value = False
    paste_widths.Enabled = False
Else
    paste_formats.Enabled = True
    paste_formulas.Enabled = True
    paste_values.Enabled = True
    paste_widths.Enabled = True
End If
End Sub
Private Sub paste_formats_Change()
If paste_formats Then
    paste_all.Enabled = False
    paste_all.value = False
ElseIf all_false() Then
    paste_all.Enabled = True
End If
End Sub
Private Sub paste_formulas_Change()
If paste_formulas.value Then
    paste_values.value = False
    paste_values.Enabled = False
    
    paste_all.value = False
    paste_all.Enabled = False
ElseIf all_false() Then
    paste_values.Enabled = True
    paste_all.Enabled = True
End If
End Sub
Private Sub paste_values_Change()
If paste_values.value Then
    paste_formulas.value = False
    paste_formulas.Enabled = False
    
    paste_all.value = False
    paste_all.Enabled = False
ElseIf all_false() Then
    paste_formulas.Enabled = True
    paste_all.Enabled = True
End If
End Sub
Private Sub paste_widths_Change()
If paste_widths.value Then
    paste_all.value = False
    paste_all.Enabled = False
ElseIf all_false() Then
    paste_all.Enabled = True
End If
End Sub
Private Sub select_folder_Click()
folder_path.Text = BrowseFolder("Select Save Location")
End Sub
Private Function all_false() As Boolean
all_false = False
If Not paste_formats.value And Not paste_formulas.value And Not paste_values.value And Not paste_widths.value Then
    all_false = True
End If
End Function
