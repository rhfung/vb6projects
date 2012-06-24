VERSION 5.00
Object = "{B0F710F8-474B-4DE4-BE20-2C604C6801EE}#9.0#0"; "RichUtil1.ocx"
Begin VB.Form frmDoc 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Create Documentation"
   ClientHeight    =   3195
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   Icon            =   "frmDoc.frx":0000
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   2280
      TabIndex        =   5
      Top             =   2640
      Width           =   855
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   960
      TabIndex        =   4
      Top             =   2640
      Width           =   855
   End
   Begin VB.CheckBox chkWriteBlank 
      Caption         =   "Write documentation for members with no descriptions?"
      Height          =   855
      Left            =   2280
      TabIndex        =   2
      Top             =   360
      Width           =   2055
   End
   Begin RichUtil1.EListBox lstMembers 
      Height          =   1815
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   3201
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CheckBoxes      =   -1  'True
      BeginProperty Column {276D682E-2140-4AB5-8D2C-F105DB576FFF} 
         Count           =   1
      EndProperty
   End
   Begin VB.Label lblName 
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   0
      Width           =   3375
   End
   Begin VB.Label Label2 
      Caption         =   "The HTML documentation will be written in the same directory as the code."
      Height          =   735
      Left            =   2280
      TabIndex        =   3
      Top             =   1680
      Width           =   1935
   End
   Begin VB.Label Label1 
      Caption         =   "Choose members to expose: "
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   2175
   End
End
Attribute VB_Name = "frmDoc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public m_blnOK As Boolean

Private Sub cmdCancel_Click()
    m_blnOK = False
    Hide
End Sub

Private Sub cmdOK_Click()
    m_blnOK = True
    Hide
End Sub


Private Sub Form_Load()
    m_blnOK = False

    lstMembers.AddItem "Public"
    lstMembers.AddItem "Friend"
    lstMembers.AddItem "Private"
    lstMembers.AddItem "Dim"
    
    lstMembers.Selected(1) = True
    lstMembers.Selected(2) = True
    lstMembers.Selected(3) = True
    lstMembers.Selected(4) = True
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode = vbFormControlMenu Then
        Cancel = 1
        m_blnOK = False
        Hide
    End If
End Sub


