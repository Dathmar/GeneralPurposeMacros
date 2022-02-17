VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AutoUpdates"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public WithEvents sht As Worksheet
Attribute sht.VB_VarHelpID = -1

'-------------------------------------------------------------------------
' Module : clsUpdate
' Company   : JKP Application Development Services (c)
' Author : Jan Karel Pieterse
' Created   : 19-2-2007
' Purpose   : Class to check for program updates
'-------------------------------------------------------------------------

#If VBA7 Then
Private Declare PtrSafe Function URLDownloadToFile Lib "urlmon" _
        Alias "URLDownloadToFileA" (ByVal pCaller As Long, _
                                    ByVal szURL As String, ByVal szFileName As String, _
                                    ByVal dwReserved As Long, ByVal lpfnCB As Long) As Long
#Else
Private Declare Function URLDownloadToFile Lib "urlmon" _
        Alias "URLDownloadToFileA" (ByVal pCaller As Long, _
                                    ByVal szURL As String, ByVal szFileName As String, _
                                    ByVal dwReserved As Long, ByVal lpfnCB As Long) As Long
#End If


Private mdtLastUpdate As Date
Private mdtRemoteUpdateDate As Date
Private msAppName As String
Private msBuild As String
Private msCheckURL As String
Private msCurrentAddinName As String
Private msDownloadName As String
Private msTempAddInName As String
Private mbManual As Boolean
Private Sub Class_Terminate()
    Set sht = Nothing
End Sub
Private Sub DownloadFile(strWebFilename As String, strSaveFileName As String)
' Download the file.
    URLDownloadToFile 0, strWebFilename, strSaveFileName, 0, 0
End Sub
Private Function json_date_to_dt(json_date As String) As String
Dim d As Date

d = DateValue(Mid$(json_date, 1, 10)) + TimeValue(Mid(json_date, 12, 8))

json_date_to_dt = d
End Function
Public Function GetGitHubDate() As Date
    Dim httpObject As Object
    Dim sGetResult As String
    Dim sURL As String
    Dim sRequest As String
    Dim sItem As Variant
    Dim blahdate As Date
    
    Set httpObject = CreateObject("MSXML2.XMLHTTP")
    
    sURL = CheckURL

    sRequest = sURL
    httpObject.Open "GET", sRequest, False
    httpObject.send
    sGetResult = httpObject.responseText
    sGetResult = Mid(sGetResult, 2, Len(sGetResult) - 2)
    
    Dim oJSON As Object
    Set oJSON = JSON_Tools.ParseJSON(sGetResult)
    
    Dim commit_date As String
    commit_date = oJSON("obj.commit.author.date")
    GetGitHubDate = json_date_to_dt(commit_date)
End Function
Public Function IsThereAnUpdate() As Boolean
    RemoteUpdateDate = GetGitHubDate
    IsThereAnUpdate = False
    If RemoteUpdateDate <> LastUpdate Then
        IsThereAnUpdate = True
    End If
End Function
Public Sub DoUpdateIfThereIsAnUpdate()
    If IsThereAnUpdate() Then
        Debug.Print "There was an update"
        Call DoUpdate
    Else
        Debug.Print "There was no update"
    End If
End Sub
Public Sub DoUpdate()
    Dim addin_name As String
    If GetUpdate Then
        Application.Cursor = xlDefault
        addin_name = Replace(AppName, "_", " ")
        MsgBox "Successfully updated the " & addin_name & " , please restart Excel to start using the new version!", vbOKOnly + vbInformation
    Else
        Application.Cursor = xlDefault
        If Manual Then
            MsgBox "Updating has failed.", vbInformation + vbOKOnly
        End If
    End If
TidyUp:
    On Error GoTo 0
    Exit Sub
End Sub
Public Property Get Build() As String
    Build = msBuild
End Property
Public Property Let Build(ByVal sBuild As String)
    msBuild = GetSetting(AppName:=AppName, Section:="Updates", key:="LastUpdate", Default:=0)
End Property
Public Sub RemoveOldCopy()
    CurrentAddinName = ThisWorkbook.FullName
    TempAddInName = CurrentAddinName & "(OldVersion)"
    On Error Resume Next
    Kill TempAddInName
End Sub
Public Function GetUpdate() As Boolean
    On Error Resume Next
    'If workbook has been saved readonly, we can safely delete the file!
    If ThisWorkbook.ReadOnly Then
        Err.Clear
        Kill CurrentAddinName
    End If
    LastUpdate = RemoteUpdateDate
    ThisWorkbook.SaveAs TempAddInName
    DoEvents
    Kill CurrentAddinName
    On Error GoTo 0
    DownloadFile DownloadName, CurrentAddinName
    If Err = 0 Then GetUpdate = True
End Function
Private Property Get CurrentAddinName() As String
    CurrentAddinName = msCurrentAddinName
End Property
Private Property Let CurrentAddinName(ByVal sCurrentAddinName As String)
    msCurrentAddinName = sCurrentAddinName
End Property
Private Property Get TempAddInName() As String
    TempAddInName = msTempAddInName
End Property
Private Property Let TempAddInName(ByVal sTempAddInName As String)
    msTempAddInName = sTempAddInName
End Property
Public Property Get DownloadName() As String
    DownloadName = msDownloadName
End Property
Public Property Let DownloadName(ByVal sDownloadName As String)
    msDownloadName = sDownloadName
End Property
Public Property Get CheckURL() As String
    CheckURL = msCheckURL
End Property
Public Property Let CheckURL(ByVal sCheckURL As String)
    msCheckURL = sCheckURL
End Property
Public Property Get LastUpdate() As Date
    Dim dtNow As Date
    dtNow = Now
    mdtLastUpdate = CDate(GetSetting(AppName, "Updates", "LastUpdate", "0"))
    If mdtLastUpdate = 0 Then
        'Never checked for an update, save today!
        SaveSetting AppName, "Updates", "LastUpdate", CStr(dtNow)
        mdtLastUpdate = dtNow
    End If
    LastUpdate = mdtLastUpdate
End Property
Public Property Let LastUpdate(ByVal dtLastUpdate As Date)
    mdtLastUpdate = dtLastUpdate
    SaveSetting AppName, "Updates", "LastUpdate", CStr(mdtLastUpdate)
End Property
Public Property Let RemoteUpdateDate(ByVal dtUpdateDate As Date)
    mdtRemoteUpdateDate = dtUpdateDate
End Property
Public Property Get RemoteUpdateDate() As Date
    RemoteUpdateDate = mdtRemoteUpdateDate
End Property
Public Property Get AppName() As String
    AppName = msAppName
End Property
Public Property Let AppName(ByVal sAppName As String)
    msAppName = sAppName
End Property
Public Property Get Manual() As Boolean
    Manual = mbManual
End Property
Public Property Let Manual(ByVal bManual As Boolean)
    mbManual = bManual
End Property

