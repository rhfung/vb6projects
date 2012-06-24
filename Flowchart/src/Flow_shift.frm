VERSION 5.00
Begin VB.Form frmShift 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Shift"
   ClientHeight    =   1170
   ClientLeft      =   6345
   ClientTop       =   5580
   ClientWidth     =   1170
   Icon            =   "Flow_shift.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   2  'Custom
   ScaleHeight     =   1170
   ScaleWidth      =   1170
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdLeft 
      Caption         =   "<"
      Height          =   375
      Left            =   15
      TabIndex        =   3
      Top             =   390
      Width           =   375
   End
   Begin VB.CommandButton cmdDown 
      Caption         =   "\/"
      Height          =   375
      Left            =   375
      TabIndex        =   2
      Top             =   750
      Width           =   375
   End
   Begin VB.CommandButton cmdRight 
      Caption         =   ">"
      Height          =   375
      Left            =   735
      TabIndex        =   1
      Top             =   390
      Width           =   375
   End
   Begin VB.CommandButton cmdUp 
      Caption         =   "/\"
      Height          =   375
      Left            =   375
      TabIndex        =   0
      Top             =   30
      Width           =   375
   End
End
Attribute VB_Name = "frmShift"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Update()
    cmdDown.Enabled = Not frmMain.IsInMode
    cmdUp.Enabled = cmdDown.Enabled
    cmdRight.Enabled = cmdDown.Enabled
    cmdLeft.Enabled = cmdDown.Enabled
End Sub

Private Sub cmdDown_Click()
    Update
    If cmdDown.Enabled Then 'all enabled
        frmMain.Shift Down:=True
    End If
End Sub

Private Sub cmdLeft_Click()
    Update
    If cmdLeft.Enabled Then 'all enabled
        frmMain.Shift Left:=True
    End If
End Sub

Private Sub cmdRight_Click()
    Update
    If cmdRight.Enabled Then 'all enabled
        frmMain.Shift Right:=True
    End If
End Sub

Private Sub cmdUp_Click()
    Update
    If cmdUp.Enabled Then 'all enabled
        frmMain.Shift Up:=True
    End If
End Sub


Private Sub Form_Activate()
    Update
End Sub

'July 3, 2001

Private Sub Form_Load()
    If WindowState <> vbMinimized Then
        Move Screen.Width * 0.8 - Width, Screen.Height * 0.95 - Height
    End If
End Sub


