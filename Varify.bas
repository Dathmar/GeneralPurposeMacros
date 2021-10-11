Attribute VB_Name = "Varify"
Option Explicit
Function is_workbook_open(ByVal wrk_name As String) As Boolean
Dim n As Long
is_workbook_open = False
For n = 1 To Application.Workbooks.Count
    If Workbooks(n).Name = wrk_name Then
        is_workbook_open = True
        Exit Function
    End If
Next n

End Function
Function is_not_saved(ByRef OWB As Workbook) As Boolean
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''                                This Macro was written by                                '''
'''                                      Asher Danner                                       '''
'''                                       03/01/2012                                        '''
'''The purpose is to check if the opened file has been saved recently.                      '''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
If OWB.Saved = False Then
    is_not_saved = True
End If
End Function
Function is_never_saved(ByRef OWB As Workbook) As Boolean
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''                                This Macro was written by                                '''
'''                                      Asher Danner                                       '''
'''                                       03/01/2012                                        '''
'''The purpose is to check if the opened file has ever been saved.                          '''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
If OWB.path = "" Then
is_never_saved = True
End If
End Function
Function want_save()
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''                                This Macro was written by                                '''
'''                                      Asher Danner                                       '''
'''                                       03/01/2012                                        '''
'''Warns the user that the workbook has not been saved recently or ever then asks the user  '''
'''if they would like to save before continuing.                                            '''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Dim i_ret As Integer
Dim file_name As String

i_ret = MsgBox("This workbook has not been saved recently." & vbNewLine & _
               "Running this macro will make irreversible changes to your workbook." & vbNewLine & _
               "Would you like to save before continuing?", vbYesNo)
               
Application.DisplayAlerts = False
If i_ret = vbYes Then
    If is_never_saved(ActiveWorkbook) = False Then
        file_name = ActiveWorkbook.FullName
        ActiveWorkbook.SaveAs filename:=file_name
    Else
        file_name = Application.GetSaveAsFilename( _
            fileFilter:="Excel Files (*.xlsx), *.xlsx")
        If file_name <> "False" Then
            ActiveWorkbook.SaveAs filename:=file_name
        End If
    End If
End If
Application.DisplayAlerts = True
End Function
Function check_save()
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''                                This Macro was written by                                '''
'''                                      Asher Danner                                       '''
'''                                       03/01/2012                                        '''
'''Checks if the files has been saved. This function should be called before any macro that '''
'''will make irreversible changes to the file.                                              '''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
If is_not_saved(ActiveWorkbook) = True Or is_never_saved(ActiveWorkbook) = True Then
    Run want_save
End If
End Function
