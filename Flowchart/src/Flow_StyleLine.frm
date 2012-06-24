VERSION 5.00
Begin VB.Form frmStyleLine 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Line Style"
   ClientHeight    =   2160
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   1335
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   144
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   89
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picSample 
      BackColor       =   &H80000005&
      Height          =   360
      Index           =   1
      Left            =   0
      ScaleHeight     =   300
      ScaleWidth      =   1275
      TabIndex        =   1
      Top             =   360
      Width           =   1335
      Begin VB.Line linSample 
         BorderStyle     =   2  'Dash
         DrawMode        =   6  'Mask Pen Not
         Index           =   1
         X1              =   120
         X2              =   1080
         Y1              =   120
         Y2              =   120
      End
   End
   Begin VB.PictureBox picSample 
      BackColor       =   &H80000005&
      Height          =   360
      Index           =   5
      Left            =   0
      ScaleHeight     =   300
      ScaleWidth      =   1275
      TabIndex        =   5
      Top             =   1800
      Width           =   1335
      Begin VB.Line linSample 
         BorderStyle     =   0  'Transparent
         DrawMode        =   6  'Mask Pen Not
         Index           =   5
         X1              =   120
         X2              =   1080
         Y1              =   120
         Y2              =   120
      End
   End
   Begin VB.PictureBox picSample 
      BackColor       =   &H80000005&
      Height          =   360
      Index           =   4
      Left            =   0
      ScaleHeight     =   300
      ScaleWidth      =   1275
      TabIndex        =   4
      Top             =   1440
      Width           =   1335
      Begin VB.Line linSample 
         BorderStyle     =   5  'Dash-Dot-Dot
         DrawMode        =   6  'Mask Pen Not
         Index           =   4
         X1              =   120
         X2              =   1080
         Y1              =   120
         Y2              =   120
      End
   End
   Begin VB.PictureBox picSample 
      BackColor       =   &H80000005&
      Height          =   360
      Index           =   3
      Left            =   0
      ScaleHeight     =   300
      ScaleWidth      =   1275
      TabIndex        =   3
      Top             =   1080
      Width           =   1335
      Begin VB.Line linSample 
         BorderStyle     =   4  'Dash-Dot
         DrawMode        =   6  'Mask Pen Not
         Index           =   3
         X1              =   120
         X2              =   1080
         Y1              =   120
         Y2              =   120
      End
   End
   Begin VB.PictureBox picSample 
      BackColor       =   &H80000005&
      Height          =   360
      Index           =   2
      Left            =   0
      ScaleHeight     =   300
      ScaleWidth      =   1275
      TabIndex        =   2
      Top             =   720
      Width           =   1335
      Begin VB.Line linSample 
         BorderStyle     =   3  'Dot
         DrawMode        =   6  'Mask Pen Not
         Index           =   2
         X1              =   120
         X2              =   1080
         Y1              =   120
         Y2              =   120
      End
   End
   Begin VB.PictureBox picSample 
      BackColor       =   &H80000005&
      Height          =   360
      Index           =   0
      Left            =   0
      ScaleHeight     =   300
      ScaleWidth      =   1275
      TabIndex        =   0
      Top             =   0
      Width           =   1335
      Begin VB.Line linSample 
         DrawMode        =   6  'Mask Pen Not
         Index           =   0
         X1              =   120
         X2              =   1080
         Y1              =   120
         Y2              =   120
      End
   End
End
Attribute VB_Name = "frmStyleLine"
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
        frmMain.ToolBoxChanged conUndoObjectFormatLineStyle  'undone
        mItem.P.LineStyle = mintLast
        frmMain.ToolBoxDone False, True, False, False, False, False, False, False
    End If
End Sub

Public Sub UpdateToForm(Item As FlowItem, Defaults As FlowChart)
    Set mItem = Item
    Set mDefaults = Defaults
    
    If Not mItem Is Nothing Then
        ChangeHighlight mItem.P.LineStyle
        SetEnabled Me, True
    Else
        SetEnabled Me, False
    End If
End Sub

Private Sub Form_Initialize()
    mintLast = -1 'nothing selected yet
End Sub



Private Sub Form_Load()
    gblnWindowLine = True
    UpdateToForm Nothing, Nothing
End Sub


Private Sub Form_Unload(Cancel As Integer)
    gblnWindowLine = False
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


