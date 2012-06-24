VERSION 5.00
Begin VB.Form frmVal 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Properties"
   ClientHeight    =   2715
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5415
   Icon            =   "flow_val.frx":0000
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   2  'Custom
   ScaleHeight     =   2715
   ScaleWidth      =   5415
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   2220
      TabIndex        =   5
      Top             =   2280
      Width           =   975
   End
   Begin VB.Frame fraProp 
      Height          =   2055
      Left            =   600
      TabIndex        =   0
      Top             =   120
      Width           =   4695
      Begin VB.Label lblValue 
         Height          =   255
         Index           =   0
         Left            =   2040
         TabIndex        =   4
         Top             =   240
         Width           =   2610
      End
      Begin VB.Label lblName 
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   3
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Value"
         Height          =   195
         Left            =   2040
         TabIndex        =   2
         Top             =   0
         Width           =   405
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Name"
         Height          =   195
         Left            =   240
         TabIndex        =   1
         Top             =   0
         Width           =   420
      End
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   0
      Picture         =   "flow_val.frx":000C
      Top             =   120
      Width           =   480
   End
End
Attribute VB_Name = "frmVal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'July 6, 2001

Private mlngCount As Long

Public Sub AddItem(Name As String, ByVal Value As String)
    mlngCount = mlngCount + 1
    If mlngCount = 0 Then
        lblName(0) = Name
        lblValue(0) = Value
    Else
        'load new row
        Load lblName(mlngCount)
        Load lblValue(mlngCount)
        'position new row
        lblName(mlngCount).Top = lblName(mlngCount - 1).Top + lblName(mlngCount - 1).Height
        lblName(mlngCount).Visible = True
        lblValue(mlngCount).Top = lblName(mlngCount).Top
        lblValue(mlngCount).Visible = True
        'set data
        lblName(mlngCount) = Name
        lblValue(mlngCount) = Value
    End If
    If TextWidth(Value) > lblValue(mlngCount).Width Then
        lblValue(mlngCount).ToolTipText = Value
    End If
End Sub


Public Sub FitToSize()
    Dim sngHeight As Single
    
    With GetLastItem()
        fraProp.Height = .Top + .Height + 120 'of last label added
    End With
    cmdOK.Top = fraProp.Top + fraProp.Height + 120
    'the code prevents window from becoming too small
    sngHeight = frmVal.Height - frmVal.ScaleHeight + cmdOK.Top + cmdOK.Height + 120
    If sngHeight > frmVal.Height Then 'make window larger
        frmVal.Height = frmVal.Height - frmVal.ScaleHeight + cmdOK.Top + cmdOK.Height + 120
    End If
End Sub

Public Function GetLastItem(Optional Value As Boolean = False) As Label
    If Value Then
        Set GetLastItem = lblValue(lblValue.UBound)
    Else
        Set GetLastItem = lblName(lblName.UBound)
    End If
End Function


Private Sub cmdOK_Click()
    Unload Me
End Sub


Private Sub Form_Load()
    mlngCount = -1
End Sub


Private Sub Form_Unload(Cancel As Integer)
    mlngCount = -1
End Sub


