Attribute VB_Name = "AutoUpdateInstall"
Option Explicit
Const default_install_dir = "%APPDATA%\Microsoft\AddIns"
Public Const ADDINNAME = "General_Purpose_Macros"
Const ADDINFILENAME = "General_Purpose_Macros.xlam"
Const ADDINDOWNLOADURL = "https://github.com/Dathmar/GeneralPurposeMacros/raw/main/General_Purpose_Macros.xlam"
Const LASTCOMMITURL = "https://api.github.com/repos/Dathmar/GeneralPurposeMacros/commits?path=General_Purpose_Macros.xlam&page=1&per_page=1"
Dim mcUpdate As AutoUpdates
Dim moWB As Workbook

Public Declare Function InternetGetConnectedState _
                         Lib "wininet.dll" (lpdwFlags As Long, _
                                            ByVal dwReserved As Long) As Boolean
Function get_app_data_path() As String
get_app_data_path = Environ("AppData")
End Function
Function IsConnected() As Boolean
    Dim Stat As Long
    IsConnected = (InternetGetConnectedState(Stat, 0&) <> 0)
End Function
Sub AutoUpdate()
    If IsConnected Then
        CheckAndUpdate False
    End If
End Sub
Sub ManualUpdate()
    On Error Resume Next
    If IsConnected Then
        Application.OnTime Now, "CheckAndUpdate"
    Else
        MsgBox "Connect to the internet to update."
    End If
End Sub
Sub CheckAndUpdate(Optional bManual = True)
    Set mcUpdate = New AutoUpdates
    With mcUpdate
        'Set intial values of class
        'Name of this app, probably a global variable, such as GSAPPNAME
        .AppName = ADDINNAME
        .Manual = bManual
        'Get rid of possible old backup copy
        .RemoveOldCopy
        .DownloadName = ADDINDOWNLOADURL
        'URL which contains build # of new version
        .CheckURL = LASTCOMMITURL
        .DoUpdateIfThereIsAnUpdate
    End With
End Sub
Public Function IsInstalled() As Boolean
    Dim oAddIn As addin
    On Error Resume Next
    If ThisWorkbook.IsAddin Then
        For Each oAddIn In Application.AddIns
            If LCase(oAddIn.FullName) = LCase(ThisWorkbook.FullName) Then
                If oAddIn.Installed Then
                    IsInstalled = True
                    Exit Function
                End If
            End If
        Next
    Else
        IsInstalled = True
    End If
End Function
Sub remove_existing_addin(keep_path As String)
Dim this_addin As addin
For Each this_addin In Application.AddIns
    If this_addin.name = ADDINFILENAME And this_addin.FullName <> keep_path Then
        this_addin.Installed = False
        Kill this_addin.FullName
    End If
Next
End Sub
Public Sub FinishInstall()
Call DeleteInstallSettings
End Sub
Public Sub DeleteInstallSettings()
If GetSetting(ADDINNAME, "Settings", "InstallLocation", "") <> "" Then
    DeleteSetting AppName:=ADDINNAME, Section:="Settings"
End If
End Sub
Public Sub CheckInstall()
    Dim oAddIn As addin
    Dim addin_name As String
    Dim install_dir As String
    Dim added_books As Boolean
    Dim resp As Integer
    Dim addin_path As String
    Dim old_path As String
    Dim addin_filename As String
    
    addin_name = Replace(ADDINNAME, "_", " ")
    If GetSetting(ADDINNAME, "Settings", "PromptToInstall", "") = "" Then
        If Not IsInstalled Then
            If InStr(LCase(ThisWorkbook.path), ".zip") > 0 Then
                MsgBox "It appears you have opened the add-in from a compressed folder" & vbNewLine & _
                       "(zip file). Please uncompress the file and open again." & vbNewLine & vbNewLine & _
                       "The add-in will automatically close now.", vbExclamation + vbOKOnly, addin_name
                ThisWorkbook.Close False
            End If
            
            install_dir = default_install_dir
            If InStr(LCase(install_dir), "%appdata%") <> 0 Then
                install_dir = get_app_data_path & right_after(LCase(install_dir), "%appdata%")
            End If
            If MsgBox("Do you wish to install " & addin_name & " as an addin?", vbQuestion + vbYesNo, addin_name) = vbYes Then
                resp = MsgBox("Do you want to install in the default location " & install_dir & "?", vbQuestion + vbYesNoCancel, addin_name)
                If resp = vbNo Then
                    install_dir = BrowseFolder("Choose Folder For Import")
                End If
                If install_dir <> "" And resp <> vbCancel And Dir(install_dir, vbDirectory) <> "" Then
                    SaveSetting ADDINNAME, "Settings", "InstallStatus", "Installing"
                    If ActiveWorkbook Is Nothing Then AddEmptyBook
                        addin_filename = ThisWorkbook.name
                        old_path = ThisWorkbook.FullName
                        addin_path = PathJoin(install_dir, addin_filename)
                        
                        If file_exists(addin_path) Then Kill addin_path
                        
                        ThisWorkbook.SaveAs filename:=addin_path, FileFormat:=xlOpenXMLAddIn
                        Call remove_existing_addin(addin_path)
                        DoEvents
                        
                        Set oAddIn = Application.AddIns.Add(addin_path, False)
                        oAddIn.Installed = True
                        
                        SaveSetting ADDINNAME, "Settings", "InstallLocation", addin_path
                        SaveSetting ADDINNAME, "Settings", "InstallFromLocation", old_path
                    RemoveEmptyBooks
                End If
            ElseIf MsgBox("Do you want me to stop asking this question?", vbQuestion + vbYesNo, addin_name) = vbYes Then
                SaveSetting ADDINNAME, "Settings", "PromptToInstall", "No"
            End If
        End If
    End If
End Sub
'---------------------------------------------------------------------------------------
' Procedure : PathJoin
' Author    : Mike Wolfe (mike@nolongerset.com)
' Source    : https://nolongerset.com/joining-paths-in-vba/
' Date      : 10/21/2015
' Purpose   : Intelligently joins path components automatically dealing with backslashes.
' Notes     - To add a trailing backslash, pass a single backslash as the final parameter.
'           - If there is no single backslash passed at the end, there will be no trailing
'               backslash (even if the final parameter contains a trailing backslash).
'           - A leading backslash in the first parameter will be left in place.
'           - Empty path components are ignored.
'---------------------------------------------------------------------------------------
'>>> PathJoin("C:\", "Users", "Public", "\")
'C:\Users\Public\
'>>> PathJoin("C:", "Users", "Public", "Settings.ini")
'C:\Users\Public\Settings.ini
'>>> PathJoin("\\localpc\C$", "\Users\", "\Public\")
'\\localpc\C$\Users\Public
'>>> PathJoin("\\localpc\C$", "\Users\", "\Public\", "")
'\\localpc\C$\Users\Public
'>>> PathJoin("\\localpc\C$", "\Users\", "\Public\", "\")
'\\localpc\C$\Users\Public\
'>>> PathJoin("Users", "Public")
'Users\Public
'>>> PathJoin("\Users", "Public\Documents", "New Text Document.txt")
'\Users\Public\Documents\New Text Document.txt
'>>> PathJoin("C:\Users\", "", "Public", "\", "Documents", "\")
'C:\Users\Public\Documents\
'>>> PathJoin("C:\Users\Public\")
'C:\Users\Public
'>>> PathJoin("C:\Users\Public\", "\")
'C:\Users\Public\
'>>> PathJoin("C:\Users\Public", "\")
'C:\Users\Public\
Public Function PathJoin(ParamArray PathComponents() As Variant) As String
    Dim LowerBound As Integer
    LowerBound = LBound(PathComponents)

    Dim UpperBound As Integer
    UpperBound = UBound(PathComponents)

    Dim i As Integer
    For i = LowerBound To UpperBound
        Dim Component As String
        Component = CStr(PathComponents(i))

        If Component = "\" And i = UpperBound Then
            'Add a trailing slash
            PathJoin = PathJoin & "\"
        Else
            'Strip trailing slash if necessary
            If Right(Component, 1) = "\" Then Component = Left(Component, Len(Component) - 1)

            'Strip leading slash if necessary
            If i > LowerBound And Left(Component, 1) = "\" Then Component = Mid(Component, 2)

            If Len(Component) = 0 Then
                'do nothing
            Else
                PathJoin = append_text(PathJoin, Component, "\")
            End If
        End If
    Next i
End Function
Sub AddEmptyBook()
Dim addin_name As String
addin_name = Replace(ADDINNAME, "_", " ")
'Adds an empty workbook if needed.
    If ActiveWorkbook Is Nothing Then
        Workbooks.Add
        Set moWB = ActiveWorkbook
        moWB.CustomDocumentProperties.Add "MyEmptyWorkbook", False, msoPropertyTypeString, "This is a temporary workbook added by " & addin_name
        moWB.Saved = True
    End If
End Sub
Sub RemoveEmptyBooks()
    Dim oWb As Workbook
    For Each oWb In Workbooks
        If IsIn(oWb.CustomDocumentProperties, "MyEmptyWorkbook") Then
            oWb.Close False
        End If
    Next
End Sub
Function IsIn(col As Variant, name As String) As Boolean
    Dim obj As Object
    On Error Resume Next
    Set obj = col(name)
    IsIn = (Err.Number = 0)
End Function
