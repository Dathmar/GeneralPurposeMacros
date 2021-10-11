VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Pass_Form 
   Caption         =   "Password Protected"
   ClientHeight    =   2010
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5130
   OleObjectBlob   =   "Pass_Form.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Pass_Form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cancel_button_Click()
Pass_Form.Hide
passwords.value = "***User has canceled the form***"
End Sub
Private Sub ok_button_Click()
Pass_Form.Hide
End Sub

