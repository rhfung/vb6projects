VERSION 5.00
Object = "{B0F710F8-474B-4DE4-BE20-2C604C6801EE}#9.2#0"; "RICHUTIL1.OCX"
Begin VB.Form frmUpdate 
   Caption         =   "Updated Comments"
   ClientHeight    =   3060
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3150
   LockControls    =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3060
   ScaleWidth      =   3150
   StartUpPosition =   2  'CenterScreen
   Begin RichUtil1.EListBox lstUpd 
      Height          =   3015
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   5318
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
      RowDivisions    =   -1  'True
      ColumnDivisions =   -1  'True
      Columns         =   2
      BeginProperty Column {276D682E-2140-4AB5-8D2C-F105DB576FFF} 
         Count           =   2
      EndProperty
   End
End
Attribute VB_Name = "frmUpdate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Sub AddItem(PropertyName As String, NewComment As String)
    lstUpd.AddItem2 PropertyName, 0, NewComment
End Sub


Private Sub Form_Load()
    lstUpd.AddItem2 "Member Name", 1, "Action"
End Sub

Private Sub Form_Resize()
    If WindowState <> vbMinimized Then
        On Error Resume Next
        lstUpd.Move 0, 0, ScaleWidth, ScaleHeight
    End If
End Sub


Private Sub lstUpd_DblClick()
    MsgBox "You cannot select items from here.  Close this window.", vbInformation
End Sub


