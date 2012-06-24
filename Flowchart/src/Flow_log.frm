VERSION 5.00
Begin VB.Form frmLog 
   Caption         =   "Log"
   ClientHeight    =   1350
   ClientLeft      =   525
   ClientTop       =   810
   ClientWidth     =   8805
   Icon            =   "Flow_log.frx":0000
   LockControls    =   -1  'True
   MinButton       =   0   'False
   PaletteMode     =   2  'Custom
   ScaleHeight     =   1350
   ScaleWidth      =   8805
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtLog 
      Height          =   1335
      Left            =   0
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   0
      Width           =   8775
   End
End
Attribute VB_Name = "frmLog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'February 11, 2001. 1.1.0

Public Sub AddLine(Msg As String)
    txtLog = txtLog & Msg & vbNewLine
    If Not Visible Then
        #If Prev Then
        Show vbModeless, frmForm
        #Else
        If frmMain Is Nothing Then
            Show vbModeless, frmForm
        Else
            Show vbModeless, frmMain
        End If
        #End If
    End If
End Sub

#If Prev Then
Private Sub Form_Load()
    Caption = "Log - Viewer Version"
End Sub
#End If

Private Sub Form_Resize()
    On Error Resume Next
    If WindowState <> vbMinimized Then
        txtLog.Move 0, 0, ScaleWidth, ScaleHeight
    End If
End Sub



Private Sub txtLog_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        Unload Me
    End If
End Sub


