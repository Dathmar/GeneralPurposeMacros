VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} AboutGPM 
   Caption         =   "About General Purpose Macros"
   ClientHeight    =   2070
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4470
   OleObjectBlob   =   "AboutGPM.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "AboutGPM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub ok_button_Click()
Unload Me
End Sub
Private Sub UserForm_Initialize()
about_label.Caption = "General Purpose Macros" & Chr(13) & _
              "Created by" & Chr(13) & _
              "Asher Danner" & Chr(13) & _
              "Version 1.37" & Chr(13) & _
              "Updated 08/27/2019"
End Sub
