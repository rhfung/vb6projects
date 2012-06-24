VERSION 5.00
Begin VB.Form frmAbout 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About Strands"
   ClientHeight    =   5085
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7650
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   FontTransparent =   0   'False
   ForeColor       =   &H00000000&
   Icon            =   "Strands_About.frx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Strands_About.frx":000C
   ScaleHeight     =   5085
   ScaleWidth      =   7650
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer timRotate 
      Enabled         =   0   'False
      Interval        =   750
      Left            =   1800
      Top             =   1920
   End
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Default         =   -1  'True
      Height          =   375
      Left            =   3360
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   4560
      Width           =   975
   End
   Begin VB.Image imgEarth 
      Height          =   4980
      Left            =   -3720
      Picture         =   "Strands_About.frx":C216
      Top             =   4320
      Visible         =   0   'False
      Width           =   4920
   End
   Begin VB.Image imgShk 
      Height          =   2835
      Left            =   1680
      Picture         =   "Strands_About.frx":14C51
      Top             =   4320
      Visible         =   0   'False
      Width           =   2205
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'April 9, 2000 by Richard Fung

Private mPlace As BaseThoughts
Private mintPos As Integer
Private Sub DrawText()
    Dim intN As Integer
    
    Cls
    For intN = 1 To 5 'for mPlace items
        Select Case intN
            Case 1, 5
                Font.Size = 8
            Case 2, 4
                Font.Size = 12
            Case 3
                Font.Size = 16
        End Select
        mPlace(intN).Refresh Me
        CurrentX = mPlace(intN).Left
        CurrentY = mPlace(intN).Top
        Print mPlace(intN)
    Next intN
    Refresh
End Sub

Private Sub ShiftText(Text As String)
    Dim intN As Integer
    
    For intN = 4 To 1 Step -1           'start at one less than total numbers
        mPlace(intN + 1) = mPlace(intN) 'because last one deleted
    Next intN
    mPlace(1) = Text 'add in new text
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Dim intPlace As Integer
    Dim sngDivide As Single
    
    Set mPlace = New BaseThoughts
    'window
    WindowState = vbNormal
    Scale (0, 0)-(conWidth, conHeight)
    'move button
    cmdClose.Left = ScaleWidth / 2 - cmdClose.Width / 2
    'split top region into 7,
    sngDivide = cmdClose.Top / 7
    'Debug.Print "CenterX", "CenterY", "ScaleWidth", "ScaleHeight"
    'fill in between 5 with text, from bottom up
    For intPlace = 5 To 1 Step -1
        mPlace.Add "", "", "", "", ScaleWidth / 2, sngDivide * intPlace
        'Debug.Print mPlace(mPlace.Count).CenterX, mPlace(mPlace.Count).CenterY, ScaleWidth, ScaleHeight
    Next intPlace
    timRotate.Enabled = True
End Sub


Private Sub Form_Unload(Cancel As Integer)
    Set mPlace = Nothing
End Sub



Private Sub timRotate_Timer()
    mintPos = mintPos + 1
    Select Case mintPos
        Case 1, 15
            ShiftText "Strands"
        Case 2, 16
            ShiftText "By Richard Fung"
        Case 4, 18
            ShiftText "Copyright © Richard Fung, 2000."
        Case 5, 19
            ShiftText "Build date April 18, 2000."
        Case 7, 21
            ShiftText """We all live in the yellow submarine..."""
        Case 8, 22
            ShiftText "-Ringo Starr"
        Case 10, 24
            ShiftText "Created in"
        Case 11, 25
            ShiftText "Microsoft® Visual Basic 5 (SP3)"
        Case 28
            ShiftText "Galileo Galilei?"
        Case 30
            ShiftText "Haa haa"
        Case 34
            ShiftText ""
            mintPos = 0
        Case Else
            ShiftText ""
    End Select
    DrawText
End Sub


