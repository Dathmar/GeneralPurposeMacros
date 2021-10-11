Attribute VB_Name = "Menus"
Option Explicit
Sub Add_Menus()
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''                                This Macro was written by                                '''
'''                                      Asher Danner                                       '''
'''                                       03/05/2012                                        '''
'''The purpose is to add a drop down menu with all macros in macro_list to Excel.           '''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Dim cMenu1 As CommandBarControl
Dim cb_main_menu_bar As CommandBar
Dim i_help_menu As Integer
Dim cbc_custom_menu As CommandBarControl
Dim delete_menu As CommandBarControl
Dim hide_menu As CommandBarControl
Dim rows_menu As CommandBarControl
Dim columns_menu As CommandBarControl
Dim tables_menu As CommandBarControl
Dim word_menu As CommandBarControl
Dim other_menu As CommandBarControl
Dim merge_menu As CommandBarControl
Dim split_menu As CommandBarControl
Dim tally_menu As CommandBarControl
Dim files_menu As CommandBarControl

Dim i As Long
Dim macro_list As Variant
Dim tip_list As Variant
Dim face_id As Long
Dim in_menu As Boolean

macro_list = Array("Delete_Unselected_Columns", "Delete_All_Hidden_Columns", "Batch_Delete_All_Hidden_Columns", "Delete_Unselected_Rows", "Delete_All_Hidden_Rows", "Delete_Unselected_Rows_and_Columns", "Hide_Unselected_Columns", _
                   "Hide_Unselected_Rows", "Hide_Unselected_Rows_and_Columns", "Unhide_All_Columns", "Unhide_All_Rows", "Columns_Times_Rows_Test", "Word_Tables_To_Excel", "Unique_Value_Spacing_for_Selected_Column", "Border_All_Cells_With_Data", "Spiral_Test", "Merge_Files_To_End_Of_Sheet", _
                   "Merge_Books_With_Same_Data_In_Columns_X_and_Y", _
                   "Merge_Files_to_Sheets", _
                   "Split_Sheets_to_Workbooks", _
                   "Split_Unique_Values_to_Books", _
                   "Split_Unique_Values_to_Sheets", _
                   "Text_Coloring", _
                   "Tally_Results", _
                   "Tally_Comments", _
                   "Add_X_Number_of_Sheets", _
                   "List_Files", _
                   "File_Summary", _
                   "Move_Files_to_Type_Folders", _
                   "Add_Hyperlinks_to_Selected_Column", _
                   "Make_Sheet_X_Active", _
                   "Remove_Passwords_From_Sheets", _
                   "Get_Combinations", _
                   "Compare_Sheets_1and2", _
                   "IBIS_Tally_Results", _
                   "Countif_Merge", "Fill_Below_in_Selected_Columns", "Trim_Cells", "copy_files_in_C1_to_C2", "Advanced_Paste_Special", "Split_Equal_Row_Count_to_Books", _
                   "About_GPMs")
                  
tip_list = Array("Delete_Unselected_Columns - Delete all columns that do not have selected cells in them, or if the whole column is selected.", "Delete_All_Hidden_Columns - Delete all columns in a worksheet that are hidden in the usedrange of a sheet.", "Batch_Delete_All_Hidden_Columns - Delete all columns in the activesheet of all selected workbooks that are hidden in the usedrange of a sheet.", _
                 "Delete_Unselected_Rows - Delete all rows that do not have selected cells in them, or if the whole row is selected.", "Delete_All_Hidden_Rows - Delete all rows in a worksheet that are hidden in the usedrange of a sheet.", _
                 "Delete_Unselected_Rows_and_Columns - Delete all rows and columns that do not have selected cells in them, or if the whole row or column selected.", "Hide_Unselected_Columns - Hide all columns that do not have selected cells in them, or if the whole column is selected.", "Hide_Unselected_Rows - Hide all rows that do not have selected cells in them, or if the whole row is selected.", "Hide_Unselected_Rows_and_Columns - Hide all rows and columns that do not have selected cells in them, or if the whole row or column selected.", "Unhide_All_Columns - Unhide all columns in the usedrange of a sheet.", "Unhide_All_Rows - Unhide all rows in the usedrange of a sheet.", _
                 "Columns_Times_Rows_Test - This macro really has no purpose other than to fill in a range of cells with the column number times the row number.  This is used for testing other macros.", _
                 "Word_Tables_To_Excel - Export all tables from selected MS Word documents into an Excel sheet", "Unique_Value_Spacing_for_Selected_Column - Add line spacing between each unique value in the selected column.", "Border_All_Cells_With_Data - Put all borders around cells that contain data.", "Spiral_Test - This macro really has no purpose other than to fill in a range of cells with a spiral pattern starting with the selected cells.", "Merge_Files_To_End_Of_Sheet - Merge all files that are selected to the end of the current open sheet.", _
                 "Merge_Books_With_Same_Data_In_Columns_X_and_Y - Merge two workbook with two columns that have the same data.  The macro checks each row of both books and if they match the data is merged into a new book.", _
                 "Merge_Files_to_Sheets - Merge selected files to individual files.", _
                 "Split_Sheets_to_Workbooks - Split all sheets in a workbook into individual workbooks.", "Split_Unique_Values_to_Books - Split all unique value sets in a selected column to new workbooks which are then named after the unique values.", _
                 "Split_Unique_Values_to_Sheets - Split all unique value sets in a selected column to new worksheets which are then named after the unique values.", "Text_Coloring - Search specified text and format based on user submissions.", _
                 "Tally_Results - Tally results in the selected area of each sheet of the workbookon sheet 1.", "Tally_Comments - Print the text in the selected area of each sheet of the workbook on sheet 1.", _
                 "Add_X_Number_of_Sheets - Create X copies of the current sheet.", _
                 "List_Files - List all files in a folder with the option to include subfolders.", _
                 "File_Summary - summarize files by unique base name and type (highlight a column of URIs).", _
                 "Move_Files_to_Type_Folders - Moves many files into a new filepath by type.", _
                 "Add_Hyperlinks_to_Selected_Column - Adds hyperlinks to filepath text in the selected column.", _
                 "Make_Sheet_X_Active - Makes sheet number X active on all selected workbooks.", _
                 "Remove_Passwords_From_Sheets - Removes passwords from sheets of selected books.", _
                 "Get_Combinations - Returns the combinations of CAs given a max correct number.", _
                 "Compare_Sheets_1and2 - Compares sheets 1 and 2 of the currently open workbook", _
                 "IBIS_Tally_Results - Returns Accept as is, Accept with edit, or Reject on a committee comment form.", _
                 "Countif_Merge - Creates a workbook with merged data validated with countifs and match.", "Fill_Below_in_Selected_Columns - Fills the cells below each value in a column.", "Trim_Cells - Removes blank spaces befor and after all cells in this worksheet.", "copy_files_in_C1_to_C2 - Copies the files from the path in column 1 to the path in column 2.", "Advanced_Paste_Special - Paste workbook in different ways", _
                 "Split_Equal_Row_Count_to_Books - The purpose is split a worksheet into many books based on number of user provided row count", _
                 "About_GPMs - Shows version number of GPMs")

'Delete any existing menu. We must use On Error Resume next _
in case it does not exist.
On Error Resume Next
Run delete_menu_function()
On Error GoTo 0
'Set a CommandBar variable to Worksheet menu bar
Set cb_main_menu_bar = Application.CommandBars("Worksheet Menu Bar")

'Return the Index number of the Help menu. We can then use this to place a custom menu before.
i_help_menu = cb_main_menu_bar.Controls("Help").Index

'Add a Control to the "Worksheet Menu Bar" before Help.
'Set a CommandBarControl variable to it
Set cbc_custom_menu = cb_main_menu_bar.Controls.Add(Type:=msoControlPopup, before:=i_help_menu)
                  
'Give the control a caption
cbc_custom_menu.Caption = "&General Macros"

' set the delete submenu
Set delete_menu = cbc_custom_menu.Controls.Add(Type:=msoControlPopup)
   ' Give the control a caption
    delete_menu.Caption = "&Delete"

' set the hide submenu
Set hide_menu = cbc_custom_menu.Controls.Add(Type:=msoControlPopup)
   ' Give the control a caption
    hide_menu.Caption = "&Hide"

Set rows_menu = cbc_custom_menu.Controls.Add(Type:=msoControlPopup)
   ' Give the control a caption
    rows_menu.Caption = "&Rows"
    
Set columns_menu = cbc_custom_menu.Controls.Add(Type:=msoControlPopup)
   ' Give the control a caption
    columns_menu.Caption = "&Columns"
    
Set merge_menu = cbc_custom_menu.Controls.Add(Type:=msoControlPopup)
   ' Give the control a caption
    merge_menu.Caption = "&Merge"
    
Set split_menu = cbc_custom_menu.Controls.Add(Type:=msoControlPopup)
   ' Give the control a caption
    split_menu.Caption = "&Split"
    
Set tables_menu = cbc_custom_menu.Controls.Add(Type:=msoControlPopup)
   ' Give the control a caption
    tables_menu.Caption = "&Tables"
    
Set word_menu = cbc_custom_menu.Controls.Add(Type:=msoControlPopup)
   ' Give the control a caption
    word_menu.Caption = "MS &Word"

Set tally_menu = cbc_custom_menu.Controls.Add(Type:=msoControlPopup)
   ' Give the control a caption
    tally_menu.Caption = "Comment &Forms"

Set files_menu = cbc_custom_menu.Controls.Add(Type:=msoControlPopup)
   ' Give the control a caption
    files_menu.Caption = "File Functions"

Set other_menu = cbc_custom_menu.Controls.Add(Type:=msoControlPopup)
   ' Give the control a caption
    other_menu.Caption = "&Other"
    
'Working with our new Control, add a sub control and _
'give it a Caption and tell it which macro to run (OnAction).
'With cbcCutomMenu.Controls.Add(Type:=msoControlButton)
'               .Caption = "Menu 1"
'               .OnAction = "MyMacro1"
'End With
If UBound(macro_list) <> UBound(tip_list) Then
    MsgBox "macro list does not match tip list"
    Exit Sub
End If

For i = LBound(macro_list) To UBound(macro_list)
'Add a contol to the sub menu, just created above
    in_menu = False
    If macro_list(i) <> Left(tip_list(i), InStr(tip_list(i), " - ") - 1) Then
        MsgBox "Macro " & macro_list(i) & " does not match it's tip"
        Exit Sub
    End If
    face_id = 0
    face_id = set_face_id(LCase(macro_list(i)))
    If InStr(LCase(macro_list(i)), "delete") <> 0 Then
        With delete_menu.Controls.Add(Type:=msoControlButton)
                       .Caption = Replace(macro_list(i), "_", " ")
                       .FaceId = face_id
                       .OnAction = macro_list(i)
                       .TooltipText = Right(tip_list(i), Len(tip_list(i)) - Len(macro_list(i)) - 3)
        End With
        in_menu = True
    End If
    If InStr(LCase(macro_list(i)), "hide") <> 0 Or InStr(LCase(macro_list(i)), "Hidden") Then
        With hide_menu.Controls.Add(Type:=msoControlButton)
                       .Caption = Replace(macro_list(i), "_", " ")
                       .FaceId = face_id
                       .OnAction = macro_list(i)
                       .TooltipText = Right(tip_list(i), Len(tip_list(i)) - Len(macro_list(i)) - 3)
        End With
        in_menu = True
    End If
    If InStr(LCase(macro_list(i)), "column") <> 0 Then
        With columns_menu.Controls.Add(Type:=msoControlButton)
                       .Caption = Replace(macro_list(i), "_", " ")
                       .FaceId = face_id
                       .OnAction = macro_list(i)
                       .TooltipText = Right(tip_list(i), Len(tip_list(i)) - Len(macro_list(i)) - 3)
        End With
        in_menu = True
    End If
    If InStr(LCase(macro_list(i)), "row") <> 0 Then
        With rows_menu.Controls.Add(Type:=msoControlButton)
                       .Caption = Replace(macro_list(i), "_", " ")
                       .FaceId = face_id
                       .OnAction = macro_list(i)
                       .TooltipText = Right(tip_list(i), Len(tip_list(i)) - Len(macro_list(i)) - 3)
        End With
        in_menu = True
    End If
    If InStr(LCase(macro_list(i)), "merge") <> 0 Then
        With merge_menu.Controls.Add(Type:=msoControlButton)
                       .Caption = Replace(macro_list(i), "_", " ")
                       .FaceId = face_id
                       .OnAction = macro_list(i)
                       .TooltipText = Right(tip_list(i), Len(tip_list(i)) - Len(macro_list(i)) - 3)
        End With
        in_menu = True
    End If
    If InStr(LCase(macro_list(i)), "table") <> 0 Then
        With tables_menu.Controls.Add(Type:=msoControlButton)
                       .Caption = Replace(macro_list(i), "_", " ")
                       .FaceId = face_id
                       .OnAction = macro_list(i)
                       .TooltipText = Right(tip_list(i), Len(tip_list(i)) - Len(macro_list(i)) - 3)
        End With
        in_menu = True
    End If
    If InStr(LCase(macro_list(i)), "word") <> 0 And InStr(LCase(macro_list(i)), "password") = 0 Then
        With word_menu.Controls.Add(Type:=msoControlButton)
                       .Caption = Replace(macro_list(i), "_", " ")
                       .FaceId = face_id
                       .OnAction = macro_list(i)
                       .TooltipText = Right(tip_list(i), Len(tip_list(i)) - Len(macro_list(i)) - 3)
        End With
        in_menu = True
    End If
    If InStr(LCase(macro_list(i)), "split") <> 0 Then
        With split_menu.Controls.Add(Type:=msoControlButton)
                       .Caption = Replace(macro_list(i), "_", " ")
                       .FaceId = face_id
                       .OnAction = macro_list(i)
                       .TooltipText = Right(tip_list(i), Len(tip_list(i)) - Len(macro_list(i)) - 3)
        End With
        in_menu = True
    End If
    If InStr(LCase(macro_list(i)), "tally") <> 0 Then
        With tally_menu.Controls.Add(Type:=msoControlButton)
                       .Caption = Replace(macro_list(i), "_", " ")
                       .FaceId = face_id
                       .OnAction = macro_list(i)
                       .TooltipText = Right(tip_list(i), Len(tip_list(i)) - Len(macro_list(i)) - 3)
        End With
        in_menu = True
    End If
    If InStr(LCase(macro_list(i)), "file") <> 0 Then
        With files_menu.Controls.Add(Type:=msoControlButton)
                       .Caption = Replace(macro_list(i), "_", " ")
                       .FaceId = face_id
                       .OnAction = macro_list(i)
                       .TooltipText = Right(tip_list(i), Len(tip_list(i)) - Len(macro_list(i)) - 3)
        End With
        in_menu = True
    End If
    If in_menu = False Then
        With other_menu.Controls.Add(Type:=msoControlButton)
                       .Caption = Replace(macro_list(i), "_", " ")
                       .FaceId = face_id
                       .OnAction = macro_list(i)
                       .TooltipText = Right(tip_list(i), Len(tip_list(i)) - Len(macro_list(i)) - 3)
        End With
    End If
    
Next i
End Sub
Function delete_menu_function()
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''                                This Macro was written by                                '''
'''                                      Asher Danner                                       '''
'''                                       03/05/2012                                        '''
'''The purpose is to delete the General Macros menue.                                       '''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
On Error Resume Next
Application.CommandBars("Worksheet Menu Bar").Controls("&General Macros").Delete
On Error GoTo 0
End Function
Function set_face_id(macro_name As Variant) As Long
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''                                This Macro was written by                                '''
'''                                      Asher Danner                                       '''
'''                                       03/05/2012                                        '''
'''The purpose is to set the face id of menu items based on conditions.                     '''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
If InStr(macro_name, "delete") <> 0 And InStr(macro_name, "column") <> 0 And InStr(macro_name, "row") <> 0 Then
    set_face_id = 0
    Exit Function
ElseIf InStr(macro_name, "delete") <> 0 And InStr(macro_name, "column") <> 0 Then
    set_face_id = 294
    Exit Function
ElseIf InStr(macro_name, "delete") <> 0 And InStr(macro_name, "row") <> 0 Then
    set_face_id = 293
    Exit Function
End If
If InStr(macro_name, "merge") <> 0 And InStr(macro_name, "sheets") <> 0 Then
    set_face_id = 658
    Exit Function
ElseIf InStr(macro_name, "merge") <> 0 And InStr(macro_name, "files") <> 0 Then
    set_face_id = 3683
    Exit Function
ElseIf InStr(macro_name, "merge") <> 0 And InStr(macro_name, "column") <> 0 Then
    set_face_id = 3688
    Exit Function
End If
If InStr(macro_name, "hide") <> 0 And InStr(macro_name, "unhide") = 0 Then
    set_face_id = 286
    Exit Function
End If
If InStr(macro_name, "tables") Then
    set_face_id = 2153
    Exit Function
End If
If InStr(macro_name, "spacing") Then
    set_face_id = 296
    Exit Function
End If
set_face_id = 0

End Function
Function About_GPMs()
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''                                This Macro was written by                                '''
'''                                      Asher Danner                                       '''
'''                                       08/21/2013                                        '''
'''The purpose is to set the face id of menu items based on conditions.                     '''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
AboutGPM.Show
End Function




















