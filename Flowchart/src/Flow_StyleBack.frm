VERSION 5.00
Begin VB.Form frmStyleBack 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Background Style"
   ClientHeight    =   510
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   3855
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   34
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   257
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox PicSample 
      BackColor       =   &H80000005&
      Height          =   495
      Index           =   7
      Left            =   3360
      ScaleHeight     =   435
      ScaleWidth      =   420
      TabIndex        =   7
      Top             =   0
      Width           =   480
      Begin VB.Shape ShpSample 
         DrawMode        =   2  'Blackness
         FillStyle       =   7  'Diagonal Cross
         Height          =   255
         Index           =   7
         Left            =   120
         Top             =   120
         Width           =   255
      End
   End
   Begin VB.PictureBox PicSample 
      BackColor       =   &H80000005&
      Height          =   495
      Index           =   6
      Left            =   2880
      ScaleHeight     =   435
      ScaleWidth      =   420
      TabIndex        =   6
      Top             =   0
      Width           =   480
      Begin VB.Shape ShpSample 
         DrawMode        =   2  'Blackness
         FillStyle       =   6  'Cross
         Height          =   255
         Index           =   6
         Left            =   120
         Top             =   120
         Width           =   255
      End
   End
   Begin VB.PictureBox PicSample 
      BackColor       =   &H80000005&
      Height          =   495
      Index           =   5
      Left            =   2400
      ScaleHeight     =   435
      ScaleWidth      =   420
      TabIndex        =   5
      Top             =   0
      Width           =   480
      Begin VB.Shape ShpSample 
         DrawMode        =   2  'Blackness
         FillStyle       =   5  'Downward Diagonal
         Height          =   255
         Index           =   5
         Left            =   120
         Top             =   120
         Width           =   255
      End
   End
   Begin VB.PictureBox PicSample 
      BackColor       =   &H80000005&
      Height          =   495
      Index           =   4
      Left            =   1920
      ScaleHeight     =   435
      ScaleWidth      =   420
      TabIndex        =   4
      Top             =   0
      Width           =   480
      Begin VB.Shape ShpSample 
         DrawMode        =   2  'Blackness
         FillStyle       =   4  'Upward Diagonal
         Height          =   255
         Index           =   4
         Left            =   120
         Top             =   120
         Width           =   255
      End
   End
   Begin VB.PictureBox PicSample 
      BackColor       =   &H80000005&
      Height          =   495
      Index           =   3
      Left            =   1440
      ScaleHeight     =   435
      ScaleWidth      =   420
      TabIndex        =   3
      Top             =   0
      Width           =   480
      Begin VB.Shape ShpSample 
         DrawMode        =   2  'Blackness
         FillStyle       =   3  'Vertical Line
         Height          =   255
         Index           =   3
         Left            =   120
         Top             =   120
         Width           =   255
      End
   End
   Begin VB.PictureBox PicSample 
      BackColor       =   &H80000005&
      Height          =   495
      Index           =   2
      Left            =   960
      ScaleHeight     =   435
      ScaleWidth      =   420
      TabIndex        =   2
      Top             =   0
      Width           =   480
      Begin VB.Shape ShpSample 
         DrawMode        =   2  'Blackness
         FillStyle       =   2  'Horizontal Line
         Height          =   255
         Index           =   2
         Left            =   120
         Top             =   120
         Width           =   255
      End
   End
   Begin VB.PictureBox PicSample 
      BackColor       =   &H80000005&
      Height          =   495
      Index           =   1
      Left            =   480
      ScaleHeight     =   435
      ScaleWidth      =   420
      TabIndex        =   1
      Top             =   0
      Width           =   480
      Begin VB.Shape ShpSample 
         DrawMode        =   2  'Blackness
         Height          =   255
         Index           =   1
         Left            =   120
         Top             =   120
         Width           =   255
      End
   End
   Begin VB.PictureBox PicSample 
      BackColor       =   &H80000005&
      Height          =   495
      Index           =   0
      Left            =   0
      ScaleHeight     =   435
      ScaleWidth      =   420
      TabIndex        =   0
      Top             =   0
      Width           =   480
      Begin VB.Shape ShpSample 
         DrawMode        =   2  'Blackness
         FillStyle       =   0  'Solid
         Height          =   255
         Index           =   0
         Left            =   120
         Top             =   120
         Width           =   255
      End
   End
End
Attribute VB_Name = "frmStyleBack"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Implements IToolbar
'date: December 28, 2002
'For flowchart

Private mItem       As FlowItem
Private mDefaults   As FlowChart
Private mintLast    As Integer

Private Sub ChangeHighlight(NewOne As Integer)
    If (mintLast > -1) Then
        picSample(mintLast).BackColor = vbWindowBackground
    End If
    picSample(NewOne).BackColor = vbHighlight
    mintLast = NewOne
End Sub

Private Sub ItemChanged()
     If Not mItem Is Nothing Then
        frmMain.ToolBoxChanged conUndoObjectFormatFillStyle  'undone
        mItem.P.FillStyle = mintLast
        frmMain.ToolBoxDone True, False, False, False, False, False, False, False
    End If
End Sub

Public Sub UpdateToForm(Item As FlowItem, Defaults As FlowChart)
    Set mItem = Item
    Set mDefaults = Defaults
    
    If Not mItem Is Nothing Then
        ChangeHighlight mItem.P.FillStyle
        SetEnabled Me, True
    Else
        SetEnabled Me, False
    End If
End Sub


Private Sub Form_Initialize()
    mintLast = -1 'nothing selected yet
End Sub



Private Sub Form_Load()
    gblnWindowBack = True
    UpdateToForm Nothing, Nothing
End Sub


Private Sub Form_Unload(Cancel As Integer)
    gblnWindowBack = False
End Sub


Private Property Get IToolbar_Height() As Single
    IToolbar_Height = Height
End Property

Private Sub IToolbar_Move(Left As Single, Top As Single)
    Move Left, Top
End Sub


Private Sub IToolbar_Show(Mode As FormShowConstants, Form As Form)
    Show Mode, Form
End Sub


Private Sub IToolbar_UpdateToForm(Item As FlowItem, Defaults As FlowChart)
    UpdateToForm Item, Defaults
End Sub


Private Sub picSample_Click(Index As Integer)
    ChangeHighlight Index
    ItemChanged
End Sub
