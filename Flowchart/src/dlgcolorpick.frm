VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmCPick 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Colour"
   ClientHeight    =   750
   ClientLeft      =   2760
   ClientTop       =   3705
   ClientWidth     =   2895
   Icon            =   "dlgcolorpick.frx":0000
   KeyPreview      =   -1  'True
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   750
   ScaleWidth      =   2895
   ShowInTaskbar   =   0   'False
   Begin MSComDlg.CommonDialog dlg 
      Left            =   1200
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox Pic 
      Height          =   375
      Index           =   15
      Left            =   2520
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   15
      ToolTipText     =   "Press R to Reset colours"
      Top             =   360
      Width           =   375
   End
   Begin VB.PictureBox Pic 
      Height          =   375
      Index           =   14
      Left            =   2160
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   14
      Top             =   360
      Width           =   375
   End
   Begin VB.PictureBox Pic 
      Height          =   375
      Index           =   13
      Left            =   1800
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   13
      Top             =   360
      Width           =   375
   End
   Begin VB.PictureBox Pic 
      Height          =   375
      Index           =   12
      Left            =   1440
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   12
      Top             =   360
      Width           =   375
   End
   Begin VB.PictureBox Pic 
      Height          =   375
      Index           =   11
      Left            =   2520
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   11
      Top             =   0
      Width           =   375
   End
   Begin VB.PictureBox Pic 
      Height          =   375
      Index           =   10
      Left            =   2160
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   10
      Top             =   0
      Width           =   375
   End
   Begin VB.PictureBox Pic 
      Height          =   375
      Index           =   9
      Left            =   1800
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   9
      Top             =   0
      Width           =   375
   End
   Begin VB.PictureBox Pic 
      Height          =   375
      Index           =   8
      Left            =   1440
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   8
      Top             =   0
      Width           =   375
   End
   Begin VB.PictureBox Pic 
      Height          =   375
      Index           =   7
      Left            =   1080
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   7
      Top             =   360
      Width           =   375
   End
   Begin VB.PictureBox Pic 
      Height          =   375
      Index           =   6
      Left            =   720
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   6
      Top             =   360
      Width           =   375
   End
   Begin VB.PictureBox Pic 
      Height          =   375
      Index           =   5
      Left            =   360
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   5
      Top             =   360
      Width           =   375
   End
   Begin VB.PictureBox Pic 
      Height          =   375
      Index           =   4
      Left            =   0
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   4
      Top             =   360
      Width           =   375
   End
   Begin VB.PictureBox Pic 
      Height          =   375
      Index           =   3
      Left            =   1080
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   3
      Top             =   0
      Width           =   375
   End
   Begin VB.PictureBox Pic 
      Height          =   375
      Index           =   2
      Left            =   720
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   2
      Top             =   0
      Width           =   375
   End
   Begin VB.PictureBox Pic 
      Height          =   375
      Index           =   1
      Left            =   360
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   1
      Top             =   0
      Width           =   375
   End
   Begin VB.PictureBox Pic 
      Height          =   375
      Index           =   0
      Left            =   0
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   0
      Top             =   0
      Width           =   375
   End
End
Attribute VB_Name = "frmCPick"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Implements IToolbar

Option Explicit

Public Enum FColourPart
    conPartFore
    conPartBack
    conPartText
End Enum

Private m_lngColours(0 To 16) As Long
Private mItem               As FlowItem
Public m_ColourPart        As FColourPart

Private Sub ItemChanged(ByVal Index As Integer)
    If Not mItem Is Nothing Then
        
        Select Case m_ColourPart
        Case conPartBack
            If mItem.P.FillStyle = vbFSTransparent Then
                frmMain.ToolBoxChanged conUndoObjectFormatFillStyle
                mItem.P.FillStyle = vbFSSolid
                frmMain.Prompt "Colour box changed fill style to solid."
            End If
            
            frmMain.ToolBoxChanged conUndoBackColourChange  'undone
            mItem.P.BackColour = Pic(Index).BackColor
            
            frmMain.ToolBoxDone False, False, False, False, False, True, False, False
        Case conPartFore
            frmMain.ToolBoxChanged conUndoForeColourChange 'undone
            mItem.P.ForeColour = Pic(Index).BackColor
            frmMain.ToolBoxDone False, False, False, False, False, False, True, False
        Case conPartText
            frmMain.ToolBoxChanged conUndoTextColourChange  'undone
            mItem.P.TextColour = Pic(Index).BackColor
            frmMain.ToolBoxDone False, False, False, False, False, False, False, True
        End Select
    End If
End Sub

Public Sub SetColourPart(ByVal ColourPart As FColourPart)
    m_ColourPart = ColourPart
    Select Case ColourPart
    Case conPartBack: Caption = "Back Colour"
    Case conPartFore: Caption = "Line Colour"
    Case conPartText: Caption = "Text Colour"
    End Select
End Sub


Private Sub Form_Initialize()
    m_lngColours(0) = QBColor(1)
    m_lngColours(1) = QBColor(9)
    
    m_lngColours(2) = QBColor(2)
    m_lngColours(3) = QBColor(10)
    
    m_lngColours(4) = QBColor(3)
    m_lngColours(5) = QBColor(11)
    
    m_lngColours(6) = QBColor(4)
    m_lngColours(7) = QBColor(12)
    
    m_lngColours(8) = QBColor(5)
    m_lngColours(9) = QBColor(13)
    
    m_lngColours(10) = QBColor(6)
    m_lngColours(11) = QBColor(14)
    
    m_lngColours(12) = RGB(36, 149, 217)
    m_lngColours(13) = QBColor(7)
    
    m_lngColours(14) = QBColor(15)
    m_lngColours(15) = vbBlack
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyR Then
        Form_Initialize
        Form_Load
    End If
End Sub


Private Sub Form_Load()
    Dim a As Integer
    
    For a = 0 To 15
        Pic(a).BackColor = m_lngColours(a)
    Next a
    
    gblnWindowPick = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim a As Integer
    
    For a = 0 To 15
        m_lngColours(a) = Pic(a).BackColor
    Next a
    
    gblnWindowPick = False
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

Public Sub UpdateToForm(Item As FlowItem, Defaults As FlowChart)
        Set mItem = Item
    
    If Not mItem Is Nothing Then
        'UNDONE: write code to show selected colour
        SetEnabled Me, True
    Else
        SetEnabled Me, False
    End If
End Sub

Private Sub Pic_Click(Index As Integer)
    ItemChanged Index
    
End Sub


Private Sub Pic_DblClick(Index As Integer)
    dlg.Color = Pic(Index).BackColor
    dlg.ShowColor
    Pic(Index).BackColor = dlg.Color
End Sub
