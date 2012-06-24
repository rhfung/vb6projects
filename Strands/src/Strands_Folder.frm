VERSION 5.00
Begin VB.Form frmFolder 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Select Folder"
   ClientHeight    =   3870
   ClientLeft      =   1530
   ClientTop       =   2160
   ClientWidth     =   4470
   Icon            =   "Strands_Folder.frx":0000
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3870
   ScaleWidth      =   4470
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtExt 
      Height          =   315
      Left            =   120
      TabIndex        =   1
      Text            =   "*.*"
      Top             =   480
      Width           =   3015
   End
   Begin VB.DirListBox dirFolder 
      Height          =   2565
      Left            =   120
      TabIndex        =   3
      Top             =   1200
      Width           =   3015
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3360
      TabIndex        =   5
      Top             =   840
      Width           =   975
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   3360
      TabIndex        =   4
      Top             =   240
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "&Filter:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "&Folders:"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   900
      Width           =   1215
   End
End
Attribute VB_Name = "frmFolder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'April 8, 2000

Private mblnOK As Boolean

Public Function GetFolder(ByVal Ext As Boolean, OwnerForm As Variant, ByRef lpstrFolder As String, ByRef lpstrExt As String) As Boolean
    'dirFolder.Path = Default
    Label1.Enabled = Ext
    txtExt.Enabled = Ext
    If Not Ext Then
        txtExt.BackColor = BackColor
        txtExt.ForeColor = ForeColor
    End If
    
    Show 1, OwnerForm
    If mblnOK Then
        GetFolder = mblnOK
        ChDrive dirFolder.Path
        ChDir dirFolder.Path
        lpstrFolder = dirFolder.Path
        If Left(txtExt, 1) = "." Then
            lpstrExt = "*" & txtExt
        ElseIf InStr(1, txtExt, "*") = 0 Then
            lpstrExt = "*" & txtExt & "*"
        Else
            lpstrExt = txtExt
        End If
    End If
    Unload Me
End Function

Private Sub cmdCancel_Click()
    mblnOK = False
    Hide
End Sub


Private Sub cmdOK_Click()
    mblnOK = True
    Hide
End Sub



Private Sub dirFolder_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeySpace Or KeyCode = vbKeyReturn Then
        dirFolder.Path = dirFolder.List(dirFolder.ListIndex)
    End If
End Sub





