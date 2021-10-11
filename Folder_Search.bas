Attribute VB_Name = "Folder_Search"
'Force the explicit delcaration of variables
Option Explicit
Function IsUserFormLoaded(ByVal UFName As String) As Boolean
    Dim UForm As Object
     
    IsUserFormLoaded = False
    For Each UForm In VBA.UserForms
        If UForm.Name = UFName Then
            IsUserFormLoaded = True
            Exit For
        End If
    Next
End Function 'IsUserFormLoaded
Sub Move_Files_to_Type_Folders()
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''                                This Macro was written by                                '''
'''                                      Asher Danner                                       '''
'''                                       03/10/2015                                        '''
'''The purpose is to move many files into a new filepath by type.                           '''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Dim list_col As Long
Dim wb As Workbook
Dim list_sht As Worksheet
Dim inc_thumbs As Boolean
Dim inc_sys As Boolean
Dim inc_sub As Boolean
Dim sum_sht As Worksheet
Dim from_path As String
Dim to_path As String
Dim to_folder As String
Dim lr As Long
Dim n As Long
Dim file_name As String
Dim move_status As Boolean
Dim move_file As Boolean
Dim file_ext As String

' check that only one column is selected
If is_only_one_column_or_row_selected("col", Selection.Address(ReferenceStyle:=xlR1C1)) = False Then
    MsgBox "Select the column with file paths."
    Exit Sub
End If

Move_fileByType_Form.Show

to_folder = BrowseFolder("Choose Base Folder to Move To")
If to_folder = "" Then Exit Sub

If IsUserFormLoaded("Move_fileByType_Form") = False Then Exit Sub
inc_thumbs = Move_fileByType_Form.thumbs.value
inc_sys = Move_fileByType_Form.sys_files
Unload Move_fileByType_Form
' get number of selected column
list_col = CLng(Replace(Selection.Address(ReferenceStyle:=xlR1C1), "C", ""))
Set wb = ActiveWorkbook
Set list_sht = wb.ActiveSheet

' add a summary sheet
wb.Sheets.Add after:=list_sht
Set sum_sht = wb.Sheets(list_sht.Index + 1)
sum_sht.Name = format_sheet_name("Summary", wb)
sum_sht.Cells(1, 1) = "From"
sum_sht.Cells(1, 2) = "To"
sum_sht.Cells(1, 3) = "Transfer Status"
lr = list_sht.Cells(list_sht.Rows.Count, list_col).End(xlUp).Row
' loop over all files
For n = 2 To lr
    move_file = True
    from_path = list_sht.Cells(n, list_col)
    file_name = get_filename(from_path)
    file_ext = get_extension(file_name)
    to_path = to_folder & "\" & file_ext & "\" & file_name
    
    If Not create_folder(to_folder & "\" & file_ext & "\") Then
        move_file = False
    End If
    If inc_thumbs = False And file_name = "Thumbs.db" Then
        move_file = False
    ElseIf inc_sys = True And Left(file_name, 1) = "." Then
        move_file = False
    End If
    move_status = False
    sum_sht.Cells(n, 1) = from_path
    sum_sht.Cells(n, 2) = to_path
    sum_sht.Cells(n, 3) = "Moving"
    If move_file Then
        move_status = Copy_File_to_Location(from_path, to_path)
    End If
    sum_sht.Cells(n, 3) = move_status

    
Next n
End Sub

Sub File_Summary()
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''                                This Macro was written by                                '''
'''                                      Asher Danner                                       '''
'''                                       03/09/2015                                        '''
'''The purpose is to summarize files by unique base name and type (highlight a column of    '''
'''URIs).                                                                                   '''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Dim wb As Workbook
Dim list_sht As Worksheet
Dim sum_sht As Worksheet
Dim list_col As Long
Dim lr As Long
Dim file_name As String
Dim n As Long
Dim i As Long

Application.ScreenUpdating = False

' check that only one column is selected
If is_only_one_column_or_row_selected("col", Selection.Address(ReferenceStyle:=xlR1C1)) = False Then
    MsgBox "Select the column with file paths."
    Exit Sub
End If
' get number of selected column
list_col = CLng(Replace(Selection.Address(ReferenceStyle:=xlR1C1), "C", ""))
Set wb = ActiveWorkbook
Set list_sht = wb.ActiveSheet

' add a summary sheet
wb.Sheets.Add after:=list_sht
Set sum_sht = wb.Sheets(list_sht.Index + 1)
sum_sht.Name = format_sheet_name("Summary", wb)

lr = list_sht.Cells(list_sht.Rows.Count, list_col).End(xlUp).Row

With sum_sht
' copy the hyperlinks to the summary sheet
list_sht.Range(list_sht.Cells(1, list_col), list_sht.Cells(lr, list_col)).Copy Destination:=sum_sht.Cells(1, 1)

' split out the file names
For n = 2 To lr
    file_name = get_filename(.Cells(n, 1))
    .Cells(n, 3) = Trim(get_extension(file_name))
    .Cells(n, 2) = Trim(Replace(file_name, "." & .Cells(n, 3), ""))
Next n

' col 1 = hyperlink
' col 2 = name
' col 3 = type

' build file type headers
.Columns("B:C").Copy Destination:=.Cells(1, 4)
.Columns(4).RemoveDuplicates Columns:=1, header:=xlNo
.Columns(5).RemoveDuplicates Columns:=1, header:=xlNo
lr = .Cells(.Rows.Count, 5).End(xlUp).Row

' save file types in an array
Dim file_types() As String
ReDim file_types(0 To lr - 1)
.Cells(1, 4) = "File Name"
For n = 2 To lr
    file_types(n - 2) = .Cells(n, 5)
    .Cells(1, 3 + n) = .Cells(n, 5)
    .Cells(n, 5) = ""
Next n

' loop through all files and hyperlink each avalible type
Dim r_row As Range
Dim cur_type As String
lr = .Cells(.Rows.Count, 4).End(xlUp).Row
.Range(.Cells(2, 5), .Cells(lr, 4 + UBound(file_types))) = "X"
For n = 2 To lr
    .UsedRange.AutoFilter Field:=2, Criteria1:=.Cells(n, 4)
    
    For Each r_row In .UsedRange.SpecialCells(xlCellTypeVisible)
        If r_row.Row <> 1 Then
            cur_type = .Cells(r_row.Row, 3)
            i = type_match(cur_type, file_types)
            .Cells(n, 5 + i) = .Cells(r_row.Row, 1)
        End If
    Next r_row
Next n

.Range(.Cells(1, 1), Cells(1, 3)).EntireColumn.Delete

End With
Application.ScreenUpdating = True
End Sub
Function type_match(cur_type As String, file_types() As String) As Long
Dim n As Long
For n = LBound(file_types) To UBound(file_types)
    If cur_type = file_types(n) Then
        type_match = n
        Exit Function
    End If
Next n
End Function
Sub list_Files()
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''                                This Macro was written by                                '''
'''                                      Asher Danner                                       '''
'''                                       03/09/2015                                        '''
'''The purpose is to list all files in a given file path.  Can include search of all        '''
'''sub folders and the ability to run a md5 checksum.                                       '''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'Declare the variable
Dim objFso As FileSystemObject
Dim FolderPath As String
Dim searchSub As Boolean
Dim md5 As Boolean

list_files_form.Show

FolderPath = BrowseFolder("Choose Folder For Import")
If FolderPath = "" Then Exit Sub

If IsUserFormLoaded("List_Files_Form") = False Then Exit Sub

md5 = list_files_form.checksum.value
searchSub = list_files_form.subfolder.value

If ActiveWorkbook Is Nothing Then Application.Workbooks.Add

'Insert the headers for Columns A through F

Cells(1, 1).value = "File Path"
Cells(1, 2).value = "File Name"
Cells(1, 3).value = "File Extension"
Cells(1, 4).value = "File Size"
Cells(1, 5).value = "File Type"
Cells(1, 6).value = "Date Created"
Cells(1, 7).value = "Date Last Accessed"
Cells(1, 8).value = "Date Last Modified"
If md5 Then
    Cells(1, 9).value = "MD5 Checksum"
End If

'Create an instance of the FileSystemObject
Set objFso = CreateObject("Scripting.FileSystemObject")
Application.ScreenUpdating = False
'Call the RecursiveFolder routine

Call RecursiveFolder(objFso, FolderPath, searchSub, md5)
Application.StatusBar = "Processing complete - Found " & ActiveSheet.UsedRange.Rows.Count - 1 & " Total Files"
Application.ScreenUpdating = True
'Change the width of the columns to achieve the best fit

End Sub
Function BrowseFolder(Title As String, _
    Optional InitialFolder As String = vbNullString, _
    Optional InitialView As Office.MsoFileDialogView = _
        msoFileDialogViewList) As String
Dim V As Variant
Dim InitFolder As String
With Application.FileDialog(msoFileDialogFolderPicker)
    .Title = Title
    .InitialView = InitialView
    If Len(InitialFolder) > 0 Then
        If Dir(InitialFolder, vbDirectory) <> vbNullString Then
            InitFolder = InitialFolder
            If Right(InitFolder, 1) <> "\" Then
                InitFolder = InitFolder & "\"
            End If
            .InitialFileName = InitFolder
        End If
    End If
    .Show
    On Error Resume Next
    Err.Clear
    V = .SelectedItems(1)
    If Err.Number <> 0 Then
        V = vbNullString
    End If
End With
BrowseFolder = CStr(V)
End Function
Function RecursiveFolder( _
FSO As FileSystemObject, _
MyPath As String, _
IncludeSubFolders As Boolean, md5 As Boolean)

'Declare the variables
Dim File As File
Dim Folder As Folder
Dim subfolder As Folder
Dim NextRow As Long
Dim myStatus As String
Dim lastFolder() As String
Dim n As Long

'Find the next available row
NextRow = ActiveSheet.Cells(ActiveSheet.Rows.Count, "A").End(xlUp).Row + 1

'Get the folder
If FSO.FolderExists(MyPath) Then
    Set Folder = FSO.GetFolder(MyPath)
Else
    Exit Function
End If
myStatus = "Found - " & NextRow - 1 & " Files Total - Checking... " & Folder
If Len(myStatus) > 255 Then
    myStatus = "Found - " & NextRow - 1 & " Files Total - File path too long to display"
End If
Application.StatusBar = myStatus
'Loop through each file in the folder
For Each File In Folder.files
    ActiveSheet.Cells(NextRow, 1).value = get_UNC(Folder.path) & "\" & File.Name
    ActiveSheet.Cells(NextRow, 2).value = File.Name
    ActiveSheet.Cells(NextRow, 3).value = get_extension(File.Name)
    ActiveSheet.Cells(NextRow, 4).value = File.Size
    ActiveSheet.Cells(NextRow, 5).value = File.Type
    ActiveSheet.Cells(NextRow, 6).value = File.DateCreated
    ActiveSheet.Cells(NextRow, 7).value = File.DateLastAccessed
    ActiveSheet.Cells(NextRow, 8).value = File.DateLastModified
    If md5 And File.Name <> "Thumbs.db" Then
        ActiveSheet.Cells(NextRow, 9).value = FileToMD5Hex(File.path)
    End If
    If NextRow Mod 100 = 0 Then
        myStatus = "Found - " & NextRow - 1 & " Files Total - Checking... " & Folder
        If Len(myStatus) > 255 Then
            myStatus = "Found - " & NextRow - 1 & " Files Total - File path too long to display"
        End If
        Application.StatusBar = myStatus
    End If
    NextRow = NextRow + 1
Next File
'Loop through files in the subfolders
If IncludeSubFolders Then
    For Each subfolder In Folder.SubFolders
        Call RecursiveFolder(FSO, subfolder.path, True, md5)
    Next subfolder
End If
End Function

