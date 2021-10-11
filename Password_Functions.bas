Attribute VB_Name = "Password_Functions"
Option Explicit
Function password_usr_form() As Variant
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''                              This function was written by                               '''
'''                                      Asher Danner                                       '''
'''                                       04/30/2013                                        '''
'''The purpose is to prompt a user for passwords.                                           '''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Dim this_pass As Variant
Dim n As Long
Dim pass_worked As Boolean
Pass_Form.Show
this_pass = Pass_Form.passwords.value
Unload Pass_Form
If this_pass = "***User has canceled the form***" Then
    password_usr_form = False
    Exit Function
End If

If InStr(this_pass, ",") <> 0 Then
    password_usr_form = split(this_pass, ",")
Else
    password_usr_form = this_pass
End If
End Function
Function try_passwords(this_sht As Worksheet, these_pass As Variant) As Boolean
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''                              This function was written by                               '''
'''                                      Asher Danner                                       '''
'''                                       04/30/2013                                        '''
'''The purpose is to try a number of passwords on a protected sheet.                        '''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Dim n As Long
Application.DisplayAlerts = False
try_passwords = True
If check_protected(this_sht) = False Then Exit Function
' try the rest
If IsArray(these_pass) Then
    For n = LBound(these_pass) To UBound(these_pass)
        On Error Resume Next
        this_sht.Unprotect Password:=these_pass(n)
        On Error GoTo 0
        If check_protected(this_sht) = False Then Exit Function
    Next n
Else
    On Error Resume Next
    this_sht.Unprotect Password:=CStr(these_pass)
    On Error GoTo 0
    If check_protected(this_sht) = False Then Exit Function
End If
' try null string last as this displayes a box
On Error Resume Next
this_sht.Unprotect Password:=vbNullString
On Error GoTo 0
If check_protected(this_sht) = True Then
    try_passwords = False
End If
Application.DisplayAlerts = True
End Function
Function check_protected(this_sht As Worksheet) As Boolean
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''                              This function was written by                               '''
'''                                      Asher Danner                                       '''
'''                                       04/30/2013                                        '''
'''The purpose is to check if a sheet is protected by a password.                           '''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' returns false if not protected
check_protected = True
If this_sht.ProtectContents = False Then check_protected = False
End Function
Sub Remove_Passwords_From_Sheets()
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''                                This Macro was written by                                '''
'''                                      Asher Danner                                       '''
'''                                       04/30/2013                                        '''
'''The purpose is to remove passwords from all selected documents.  The macro will ask for  '''
'''a password.  Multiple passwords maybe entered but seperated with a comma.                '''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Dim xl_file_name As Variant
Dim this_workbook As Long
Dim passwords As Variant
Dim report_book As Workbook
Dim this_book As Workbook
Dim sht_count As Long
Dim n As Long

xl_file_name = Application.GetOpenFilename("Excel files (*.xl*),*.xl*", , _
    "Browse for file to be merged", MultiSelect:=True)
Application.DisplayAlerts = False
Application.ScreenUpdating = False

If IsArray(xl_file_name) Then
    passwords = password_usr_form()
    Application.Workbooks.Add
    Set report_book = ActiveWorkbook
    For this_workbook = LBound(xl_file_name) To UBound(xl_file_name) ' iterate through each book
        Application.Workbooks.Open filename:=xl_file_name(this_workbook), ReadOnly:=False ' open the books
        Set this_book = ActiveWorkbook
        sht_count = this_book.Sheets.Count
        For n = 1 To sht_count
            report_book.Sheets(1).Cells(this_workbook, 1) = this_book.Name
            report_book.Sheets(1).Cells(this_workbook, n + 1) = try_passwords(this_book.Sheets(n), passwords)
        Next n
        this_book.Save
        this_book.Close
        Set this_book = Nothing
    Next this_workbook
End If
Application.DisplayAlerts = True
Application.ScreenUpdating = True
End Sub
