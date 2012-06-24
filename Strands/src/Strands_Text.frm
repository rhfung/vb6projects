VERSION 5.00
Begin VB.Form frmText 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Edit Text"
   ClientHeight    =   1560
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   2400
   Icon            =   "Strands_Text.frx":0000
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1560
   ScaleWidth      =   2400
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtEdit 
      Height          =   1095
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   240
      Width           =   1935
   End
End
Attribute VB_Name = "frmText"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Richard Fung.  March 31, 2000.

Private mobjAttached As Thought
Private mblnEdit     As Boolean

Public Property Get Attached() As Thought
    Set Attached = mobjAttached
End Property

Public Property Set Attached(ByVal objNew As Thought)
    mblnEdit = True
    Set mobjAttached = objNew
    If Not objNew Is Nothing Then
        txtEdit = objNew.Text
        txtEdit.Visible = True
    Else
        txtEdit.Visible = False
    End If
    mblnEdit = False
End Property

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode = vbFormControlMenu Then
        Cancel = 1
        Hide
        Set Attached = Nothing
    End If
End Sub

Private Sub Form_Resize()
    If WindowState <> vbMinimized Then
        txtEdit.Move 0, 0, ScaleWidth, ScaleHeight
    End If
End Sub


Private Sub txtEdit_Change()
    If Not mblnEdit And Not Attached Is Nothing Then
        Attached.Text = txtEdit
        If Not frmMain.Selected Is Nothing Then
            With frmMain
                If Len(.Mind.Header.Comment) Then
                    .lblText = .Mind.Header.Comment & vbNewLine & .Selected.Text
                Else
                    .lblText = .Selected.Text
                End If
            End With
        End If
    End If
End Sub


