VERSION 5.00
Object = "{B0F710F8-474B-4DE4-BE20-2C604C6801EE}#9.1#0"; "RichUtil1.ocx"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9285
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   9285
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Height          =   855
      Left            =   5520
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   1680
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   615
      Left            =   5880
      TabIndex        =   1
      Top             =   720
      Width           =   855
   End
   Begin RichUtil1.EListBox EListBox1 
      Height          =   2895
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4335
      _extentx        =   7646
      _extenty        =   5106
      backcolor       =   -2147483643
      forecolor       =   -2147483640
      font            =   "Test.frx":0000
      checkboxes      =   -1  'True
      rowdivisions    =   -1  'True
      columndivisions =   -1  'True
      columns         =   4
      column          =   "Test.frx":0024
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit



Private Declare Function DrawFocusRect Lib "user32" (ByVal hdc As Long, lpRect As RECT) As Long
Private Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type





Private Sub EListBox1_DblClick()
    EListBox1.EditBoxOn EListBox1.ListIndex, Int(Rnd * 4) + 1
End Sub


Private Sub EListBox1_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then
        EListBox1.RemoveItem EListBox1.ListIndex
    End If

End Sub



Private Sub EListBox1_ValidateEdit(ByVal Row As Long, ByVal Col As Long, ByVal EditText As String, Cancel As Boolean)
    If Len(EditText) = 0 Then Cancel = True
End Sub


Private Sub Form_DblClick()
    EListBox1.ListCount = 50
    EListBox1.List(50) = "The Last Item is HERE"
End Sub




Private Sub Form_Load()
    
    EListBox1.AddItem "Hi"
    EListBox1.AddItem "Bye"
    EListBox1.AddItem2 "C1", 0, "C2", "Column 3"
    
    Dim intI As Integer
    Dim lngJ As Long
    
    For lngJ = 1 To 100
        For intI = Asc("A") To Asc("Z")
            EListBox1.AddItem2 String(Int(Rnd * 20) + 1, Chr(intI)), 0, intI, Asc(LCase(Chr(intI)))
        Next intI
    Next lngJ
    
    'lst2.AddItem "Hi"
End Sub


Private Sub Form_Paint()
    Dim tRect As RECT
    
    tRect.Right = 10
    tRect.Bottom = 10
    DrawFocusRect hdc, tRect
End Sub


Private Sub Form_Resize()
    EListBox1.Top = 0
    EListBox1.Height = ScaleHeight
End Sub

