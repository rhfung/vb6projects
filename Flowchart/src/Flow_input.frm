VERSION 5.00
Begin VB.Form frmInput 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Flow Chart"
   ClientHeight    =   1815
   ClientLeft      =   975
   ClientTop       =   1305
   ClientWidth     =   5265
   Icon            =   "Flow_input.frx":0000
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   2  'Custom
   ScaleHeight     =   1815
   ScaleWidth      =   5265
   ShowInTaskbar   =   0   'False
   Begin VB.ComboBox cboList 
      Height          =   315
      Left            =   120
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   1320
      Visible         =   0   'False
      Width           =   5055
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   345
      Left            =   4200
      TabIndex        =   4
      Top             =   600
      Width           =   855
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   345
      Left            =   4200
      TabIndex        =   3
      Top             =   180
      Width           =   855
   End
   Begin VB.TextBox txtEdit 
      Height          =   315
      Left            =   120
      TabIndex        =   1
      Top             =   1320
      Width           =   5055
   End
   Begin VB.Label lblPrompt 
      Height          =   1215
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3855
   End
End
Attribute VB_Name = "frmInput"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Richard Fung.  Tues August 15, 2000.
 
Public mblnOK As Boolean

 


Private Sub cmdCancel_Click()
    mblnOK = False
    Hide
End Sub

Private Sub cmdOK_Click()
    mblnOK = True
    Hide
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode <> vbFormCode Then
        Cancel = 1
        mblnOK = False
        Hide
    End If
End Sub


Private Sub txtEdit_GotFocus()
    SendKeys "{Home}+{End}"
End Sub


