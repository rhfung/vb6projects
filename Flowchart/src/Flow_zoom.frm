VERSION 5.00
Begin VB.Form frmZoom 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Zoom"
   ClientHeight    =   2535
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3090
   Icon            =   "Flow_zoom.frx":0000
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   2  'Custom
   ScaleHeight     =   2535
   ScaleWidth      =   3090
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.OptionButton opZoom 
      Caption         =   "25%"
      Height          =   255
      Index           =   25
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   1778
      TabIndex        =   9
      Top             =   2040
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   240
      TabIndex        =   8
      Top             =   2040
      Width           =   1095
   End
   Begin VB.OptionButton opOther 
      Caption         =   "Other"
      Height          =   255
      Left            =   1560
      TabIndex        =   7
      Top             =   1320
      Width           =   1095
   End
   Begin VB.OptionButton opZoom 
      Caption         =   "200%"
      Height          =   255
      Index           =   200
      Left            =   1560
      TabIndex        =   6
      Top             =   960
      Width           =   1095
   End
   Begin VB.OptionButton opZoom 
      Caption         =   "150%"
      Height          =   255
      Index           =   150
      Left            =   1560
      TabIndex        =   5
      Top             =   600
      Width           =   1095
   End
   Begin VB.OptionButton opZoom 
      Caption         =   "125%"
      Height          =   255
      Index           =   125
      Left            =   1560
      TabIndex        =   4
      Top             =   240
      Width           =   1095
   End
   Begin VB.OptionButton opZoom 
      Caption         =   "100%"
      Height          =   255
      Index           =   100
      Left            =   240
      TabIndex        =   3
      Top             =   1320
      Width           =   1095
   End
   Begin VB.OptionButton opZoom 
      Caption         =   "75%"
      Height          =   255
      Index           =   75
      Left            =   240
      TabIndex        =   2
      Top             =   960
      Width           =   1095
   End
   Begin VB.OptionButton opZoom 
      Caption         =   "50%"
      Height          =   255
      Index           =   50
      Left            =   240
      TabIndex        =   1
      Top             =   600
      Width           =   1095
   End
   Begin VB.Label lblOtherZoom 
      Height          =   255
      Left            =   1680
      MousePointer    =   3  'I-Beam
      TabIndex        =   10
      Top             =   1560
      Width           =   1215
   End
End
Attribute VB_Name = "frmZoom"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'August 21, 2000.

Public Changed As Boolean

Private mlngZoom As Long


Private Sub UpdateZoom()
    If opOther Then
        lblOtherZoom = "( " & mlngZoom & "% )"
    Else
        lblOtherZoom = ""
    End If
End Sub


Public Property Let Zoom(ByVal pZoom As Long)
    mlngZoom = pZoom
    opOther.Tag = pZoom
    
    Select Case pZoom
        Case 25, 50, 75, 100, 125, 150, 200
            opZoom(pZoom) = True
        Case Else
            opOther = True
    End Select
    Changed = False
End Property

Public Property Get Zoom() As Long
    Zoom = mlngZoom
End Property


Private Sub cmdCancel_Click()
    Changed = False
    Hide
End Sub

Private Sub cmdOK_Click()
    Hide
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode <> vbFormCode Then
        Cancel = 1
        Changed = False
        Hide
    End If
End Sub


Private Sub lblOtherZoom_Click()
    opOther_Click
End Sub

Private Sub opOther_Click()
    mlngZoom = opOther.Tag
    Changed = True
    
    If Visible Then
        Dim strNewZoom As String
        Dim lngNewZoom As Long
        
        strNewZoom = InputBox$("Enter in a percent value for zoom from 10 to 500 percent.", Caption, mlngZoom)
        If Len(strNewZoom) And IsNumeric(strNewZoom) Then
            lngNewZoom = strNewZoom
            If 10 <= lngNewZoom And lngNewZoom <= 500 Then
                mlngZoom = lngNewZoom 'change zoom value
                opOther.Tag = mlngZoom 'update tag
            End If
        End If
    End If
    UpdateZoom
End Sub

Private Sub opZoom_Click(Index As Integer)
    mlngZoom = Index
    Changed = True
    UpdateZoom
End Sub


