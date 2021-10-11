VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Combinations 
   Caption         =   "Combinations"
   ClientHeight    =   2700
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4710
   OleObjectBlob   =   "Combinations.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Combinations"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Cancel_Click()
Unload Me
End Sub
Private Sub max_len_box_Change()
Dim op_len As Long
Dim max_len As Long

op_len = Len(CAs_box.value) - Len(Replace(CAs_box.value, ",", "")) + 1
If max_len_box.value <> "" Then max_len = max_len_box.value
If op_len < 2 Then
    Label4.Caption = "Total Combinations: 0"
Else
    Label4.Caption = "Total Combinations: " & Application.WorksheetFunction.Combin(op_len, max_len)
End If
End Sub
Private Sub OK_Click()
Dim n As Long
Dim options() As String
Dim i As Integer
Dim pool() As String
'create an array with a string
options = split(CAs_box.value, ",")
i = 1
n = Abs(UBound(options) - LBound(options) + 1)
ReDim pool(1 To n)
For n = LBound(options) To UBound(options)
    pool(i) = Trim(options(n))
    i = i + 1
Next n
Erase options

printCombinations pool(), max_len_box.value
Unload Me
End Sub
Private Function printCombinations(ByRef pool() As String, ByVal r As Integer)
 Dim n As Integer
 Dim i As Integer
 Dim j As Integer
 Dim c As Integer
 Dim combo As String
 n = UBound(pool) - LBound(pool) + 1

' Please do add error handling for when r>n

 Dim idx() As Integer
 ReDim idx(1 To r)
 For i = 1 To r
     idx(i) = i
 Next i
 c = 1
 Do
     'Write current combination
     For j = 1 To r
         If j = r Then
             combo = Trim(combo & pool(idx(j)))
         Else
             combo = Trim(combo & pool(idx(j)) & ",")
         End If
     Next j
     ' Locate last non-max index
     i = r
     While (idx(i) = n - r + i)
         i = i - 1
         If i = 0 Then
             'All indexes have reached their max, so we're done
             Cells(c, 1) = combo
             Exit Function
         End If
     Wend
     Cells(c, 1) = combo
     combo = ""
     c = c + 1
     'Increase it and populate the following indexes accordingly
     idx(i) = idx(i) + 1
     For j = i + 1 To r
         idx(j) = idx(i) + j - i
     Next j
 Loop
End Function
Private Function factorial(n As Long) As Long
If n <= 1 Then factorial = 1: Exit Function
factorial = n * factorial(n - 1)
End Function
