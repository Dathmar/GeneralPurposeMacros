Attribute VB_Name = "File_Functions"
Sub copy_files_in_C1_to_C2()
Dim path1 As String
Dim path2 As String
Dim sht As Worksheet
Dim wb
Dim lr As Long

Set wb = ActiveWorkbook

If wb Is Nothing Then
    Exit Sub
End If

Set sht = wb.ActiveSheet

lr = sht.Cells(sht.Rows.Count, 1).End(xlUp).Row

For n = 2 To lr
    path1 = sht.Cells(n, 1)
    path2 = sht.Cells(n, 2)
    sht.Cells(n, 3) = Copy_File_to_Location(path1, path2)
Next n

End Sub
Function pick_files(multi As Boolean) As Variant
Dim file_list As Variant
Dim file_arr(0 To 0) As Variant
file_list = Application.GetOpenFilename("Excel files (*.xls; *.xlsx; *.csv), *.xls; *.xlsx; *.csv", , _
    "Browse for file to be merged", MultiSelect:=multi)

If IsArray(file_list) Then
    pick_files = file_list
Else
    file_arr(0) = file_list
    pick_files = file_arr
End If
End Function
Function Copy_File_to_Location(from_file As String, to_file As String) As Boolean
Dim FSO As Object
Dim path As String
Set FSO = CreateObject("scripting.filesystemobject")

path = Replace(to_file, get_filename(to_file), "")

Call recursive_mkdir(path) ' should make a make path recursive function

On Error GoTo Copy_File_to_Location_Er
'FSO.MoveFile Source:=from_file, Destination:=to_file
FSO.CopyFile Source:=from_file, Destination:=to_file
On Error GoTo 0
Set FSO = Nothing
Copy_File_to_Location = True
Exit Function
Copy_File_to_Location_Er:
Copy_File_to_Location = False
Set FSO = Nothing
End Function
Function recursive_mkdir(path As String)
Dim parent_path As String
Dim folders_split As Variant

If Len(Dir(parent_path, vbDirectory)) = 1 Or Len(parent_path) = 0 Then Exit Function
'remove trailing \
If Right(path, 1) = "\" Then
    path = Left(path, Len(path) - 1)
End If

folders_split = split(path, "\")

parent_path = Left(path, Len(path) - 1 - Len(folders_split(UBound(folders_split))))

If Len(Dir(parent_path, vbDirectory)) = 0 Then
    Call recursive_mkdir(parent_path)
End If

Call recursive_mkdir(path)

End Function
Function get_UNC(strMappedDrive As String) As String
    Dim objFso As FileSystemObject
    Set objFso = New FileSystemObject
    Dim strDrive As String
    Dim strShare As String
    
    If Left(strMappedDrive, 2) = "\\" Then
        get_UNC = strMappedDrive
        Exit Function
    End If
    
    'Separated the mapped letter from
    'any following sub-folders
    strDrive = objFso.GetDriveName(strMappedDrive)
    'find the UNC share name from the mapped letter
    strShare = objFso.Drives(strDrive).ShareName
    'The Replace function allows for sub-folders
    'of the mapped drive
    get_UNC = Replace(strMappedDrive, strDrive, strShare)
    Set objFso = Nothing 'Destroy the object
End Function
Public Function get_name_without_extension(filename As String) As String
Dim extension As Variant
extension = get_extension(filename)

get_name_without_extension = Replace(filename, "." & extension, "")

End Function
Public Function get_extension(filename As String) As String
Dim name_split As Variant
name_split = split(filename, ".")
get_extension = CStr(name_split(UBound(name_split)))
End Function
Public Function get_filename(ByVal filepath As String) As String
Dim name_split As Variant
name_split = split(filepath, "\")
get_filename = CStr(name_split(UBound(name_split)))
End Function
Public Function FileToMD5Hex(sFileName As String) As String
    Dim enc
    Dim bytes
    Dim outstr As String
    Dim pos As Integer
    Set enc = CreateObject("System.Security.Cryptography.MD5CryptoServiceProvider")
    'Convert the string to a byte array and hash it
    bytes = GetFileBytes(sFileName)
    
    If UBound(bytes) = 0 And bytes(0) = CByte(0) Then
        FileToMD5Hex = ""
        Exit Function
    End If
    bytes = enc.ComputeHash_2((bytes))
    'Convert the byte array to a hex string
    For pos = 1 To LenB(bytes)
        outstr = outstr & LCase(Right("0" & Hex(AscB(MidB(bytes, pos, 1))), 2))
    Next
    FileToMD5Hex = outstr
    Set enc = Nothing
End Function
Private Function GetFileBytes(ByVal path As String) As Byte()
    Dim lngFileNum As Long
    Dim bytRtnVal() As Byte
    Dim has_err As Boolean
    Dim n As Long
    
    n = 0
    has_err = False
    lngFileNum = FreeFile
    If LenB(Dir(path)) Then ''// Does file exist?
        
        Do
            On Error Resume Next
            Open path For Binary Access Read As lngFileNum
            
            If Err.Number <> 0 Then
                has_err = True
                Err.Number = 0
                n = n + 1
            Else
                has_err = False
                On Error GoTo 0
            End If
            
            If n > 6 Then
                Application.Wait (Now + TimeValue("0:00:01"))
            End If
        Loop Until n <> 10 Or has_err = False
        
        If LOF(lngFileNum) = 0 Or has_err Then
            ReDim bytRtnVal(0 To 0) As Byte
            bytRtnVal(0) = CByte(0)
            GetFileBytes = bytRtnVal
            Close lngFileNum
            Exit Function
        End If
        ReDim bytRtnVal(LOF(lngFileNum) - 1&) As Byte
        
        Get lngFileNum, , bytRtnVal
        Close lngFileNum
    Else
        ReDim bytRtnVal(0 To 0) As Byte
        bytRtnVal(0) = CByte(0)
        GetFileBytes = bytRtnVal
        Close lngFileNum
        Exit Function
    End If
    GetFileBytes = bytRtnVal
    Erase bytRtnVal
End Function
Function create_folder(folder_path As String) As Boolean
Dim FSO As Object
On Error GoTo Err_2
Set FSO = CreateObject("scripting.filesystemobject")


If Right(folder_path, 1) <> "\" Then
    folder_path = folder_path & "\"
End If

If FSO.FolderExists(folder_path) Then
    create_folder = True
Else
    MkDir folder_path
    create_folder = True
End If

Set FSO = Nothing
On Error GoTo 0
Exit Function
Err_2:
On Error GoTo 0
create_folder = False
End Function
Function file_exists(filename As String) As Boolean
file_exists = False
If LenB(Dir(filename)) Then file_exists = True
End Function
Function get_unique_filename(filename As String) As String
Dim filepath As String
Dim justname As String
Dim extension As String
Dim counter As String

counter = 1

Do While file_exists(filename)
    justname = get_filename(filename)
    filepath = Replace(filename, justname, "")
    extension = get_extension(filename)
    justname = Replace(justname, "." & extension, "")
    
    If counter <> 1 Then
        justname = Left(justname, Len(justname) - Len(counter))
    End If
    
    justname = justname & counter
    counter = counter + 1
    filename = filepath & justname & "." & extension
    
Loop
get_unique_filename = filename
End Function


