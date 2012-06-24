VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmColour 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Colour Box"
   ClientHeight    =   1905
   ClientLeft      =   7635
   ClientTop       =   4845
   ClientWidth     =   1275
   FontTransparent =   0   'False
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   1905
   ScaleWidth      =   1275
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox picText 
      Height          =   255
      Left            =   720
      ScaleHeight     =   195
      ScaleWidth      =   435
      TabIndex        =   7
      Top             =   1560
      Width           =   495
   End
   Begin MSComDlg.CommonDialog dlgColour 
      Left            =   1320
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox picBackColour 
      Height          =   255
      Left            =   720
      ScaleHeight     =   195
      ScaleWidth      =   435
      TabIndex        =   5
      Top             =   1200
      Width           =   495
   End
   Begin VB.PictureBox picBackFill 
      DrawStyle       =   5  'Transparent
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   0
      ScaleHeight     =   195
      ScaleWidth      =   555
      TabIndex        =   4
      Top             =   1200
      Width           =   615
   End
   Begin VB.PictureBox picForeColour 
      Height          =   255
      Left            =   720
      ScaleHeight     =   195
      ScaleWidth      =   435
      TabIndex        =   2
      Top             =   240
      Width           =   495
   End
   Begin VB.PictureBox picLineStyle 
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   0
      ScaleHeight     =   255
      ScaleWidth      =   615
      TabIndex        =   1
      ToolTipText     =   "Right-click for line width."
      Top             =   240
      Width           =   615
      Begin VB.Line linLineStyle 
         X1              =   0
         X2              =   600
         Y1              =   120
         Y2              =   120
      End
   End
   Begin VB.Label lblFillStyle 
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   0
      TabIndex        =   8
      Top             =   840
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "Text"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   0
      TabIndex        =   6
      Top             =   1560
      Width           =   855
   End
   Begin VB.Label Label2 
      Caption         =   "Background"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   0
      TabIndex        =   3
      Top             =   600
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Line"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   615
   End
   Begin VB.Menu mnuFill 
      Caption         =   ""
      Visible         =   0   'False
      Begin VB.Menu mnuFillItem 
         Caption         =   "Solid"
         Index           =   0
      End
      Begin VB.Menu mnuFillItem 
         Caption         =   "Transparent"
         Index           =   1
      End
      Begin VB.Menu mnuFillItem 
         Caption         =   "Horizontal Line"
         Index           =   2
      End
      Begin VB.Menu mnuFillItem 
         Caption         =   "Vertical Line"
         Index           =   3
      End
      Begin VB.Menu mnuFillItem 
         Caption         =   "Upward Diagonal"
         Index           =   4
      End
      Begin VB.Menu mnuFillItem 
         Caption         =   "Downward Diagonal"
         Index           =   5
      End
      Begin VB.Menu mnuFillItem 
         Caption         =   "Cross"
         Index           =   6
      End
      Begin VB.Menu mnuFillItem 
         Caption         =   "Diagonal Cross"
         Index           =   7
      End
   End
   Begin VB.Menu mnuLine 
      Caption         =   ""
      Visible         =   0   'False
      Begin VB.Menu mnuLineItem 
         Caption         =   "Solid"
         Index           =   0
      End
      Begin VB.Menu mnuLineItem 
         Caption         =   "Dash"
         Index           =   1
      End
      Begin VB.Menu mnuLineItem 
         Caption         =   "Dot"
         Index           =   2
      End
      Begin VB.Menu mnuLineItem 
         Caption         =   "Dash-Dot"
         Index           =   3
      End
      Begin VB.Menu mnuLineItem 
         Caption         =   "Dash-Dot-Dot"
         Index           =   4
      End
      Begin VB.Menu mnuLineItem 
         Caption         =   "Invisible"
         Index           =   5
      End
      Begin VB.Menu mnuLineItem 
         Caption         =   "Inside Solid"
         Index           =   6
      End
   End
   Begin VB.Menu mnuWidth 
      Caption         =   ""
      Visible         =   0   'False
      Begin VB.Menu mnuWidthItem 
         Caption         =   "Thin Line"
         Index           =   0
      End
   End
End
Attribute VB_Name = "frmColour"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Colour box for Flow Chart


Private Sub ItemChanged(ByVal UndoAction As PUndoAction)
    frmMain.ToolBoxChanged UndoAction
End Sub

Private Property Get mSelected() As FlowItem
    Set mSelected = frmMain.mSelected
End Property

Private Sub SetLineStyle(ByVal Style As Integer)
    Select Case Style
        Case vbSolid: linLineStyle.BorderStyle = vbBSSolid
        Case vbDash: linLineStyle.BorderStyle = vbBSDash
        Case vbDot: linLineStyle.BorderStyle = vbBSDot
        Case vbDashDot: linLineStyle.BorderStyle = vbBSDashDot
        Case vbDashDotDot: linLineStyle.BorderStyle = vbBSDashDotDot
        Case vbInvisible: linLineStyle.BorderStyle = vbBSNone
        Case vbInsideSolid: linLineStyle.BorderStyle = vbBSInsideSolid
    End Select
End Sub

Public Sub Update()
    If frmMain.IsSelected Then
        With mSelected
'        If TypeOf .mSelected Is FPicture Then
'        ElseIf TypeOf .mSelected Is FText Then
'        ElseIf TypeOf .mSelected Is FCircle Then
'        Else
            SetLineStyle .P.LineStyle
            picBackFill.FillStyle = .P.FillStyle
            lblFillStyle = mnuFillItem(picBackFill.FillStyle).Caption
            picForeColour.BackColor = .P.ForeColour
            picBackColour.BackColor = .P.BackColour
            picText.BackColor = .P.TextColour
'        End If
        End With
        picLineStyle.Visible = True
        picForeColour.Visible = True
        lblFillStyle.Visible = True
        picBackFill.Visible = True
        picBackColour.Visible = True
        picText.Visible = True
    Else
        picLineStyle.Visible = False
        picForeColour.Visible = False
        lblFillStyle.Visible = False
        picBackFill.Visible = False
        picBackColour.Visible = False
        picText.Visible = False
    End If
End Sub

Private Sub Form_Activate()
    Update
End Sub

Private Sub Form_Load()
    Dim intI As Integer
    
    gblnWindowColour = True
    For intI = 1 To 6
        Load mnuWidthItem(intI)
        mnuWidthItem(intI).Caption = CStr(intI)
    Next intI
    If WindowState <> vbMinimized Then
        Move Screen.Width * 0.8, Screen.Height * 0.95 - Height
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    gblnWindowColour = False
End Sub






Private Sub lblFillStyle_Click()
    PopupMenu mnuFill
End Sub

Private Sub mnuFillItem_Click(Index As Integer)
    picBackFill.FillStyle = Index
    If frmMain.IsSelected Then
        ItemChanged conUndoObjectFormatFillStyle
        mSelected.P.FillStyle = Index
        frmMain.ToolBoxDone True, False, False, False, False, False, False, False
    End If
    picBackFill_Paint
End Sub



Private Sub mnuLineItem_Click(Index As Integer)
    SetLineStyle Index
    
    If frmMain.IsSelected Then
        ItemChanged conUndoObjectFormatLineStyle
        mSelected.P.LineStyle = Index
        frmMain.ToolBoxDone False, True, False, False, False, False, False, False
    End If
End Sub




Private Sub mnuWidth_Click()
    Dim intItem As Integer
    For intItem = mnuWidthItem.LBound To mnuWidthItem.UBound
        mnuWidthItem(intItem).Checked = False
    Next intItem
    mnuWidthItem(mSelected.P.LineWidth).Checked = True
End Sub


Private Sub mnuWidthItem_Click(Index As Integer)
    frmMain.ToolBoxChanged conUndoObjectFormatLineWidth
    mSelected.P.LineWidth = Index
    frmMain.ToolBoxDone False, False, True, False, False, False, False, False
End Sub


Private Sub picBackColour_Click()
    On Error Resume Next
    With dlgColour
        .CancelError = True
        .Color = mSelected.P.BackColour
        .Flags = cdlCCFullOpen Or cdlCCRGBInit
        .ShowColor
        If Err = 0 Then
            picBackColour.BackColor = .Color
            If frmMain.IsSelected Then
                ItemChanged conUndoBackColourChange
                'change fill style automatically
                If mSelected.P.FillStyle = vbFSTransparent Then
                    mSelected.P.FillStyle = vbFSSolid
                    picBackFill.FillStyle = mSelected.P.FillStyle
                    picBackFill_Paint
                
                    mSelected.P.BackColour = .Color
                    frmMain.ToolBoxDone True, False, False, False, False, True, False, False
                Else
                    mSelected.P.BackColour = .Color
                    frmMain.ToolBoxDone False, False, False, False, False, True, False, False
                End If
                
            End If
        End If
    End With
End Sub

Private Sub picBackFill_Click()
    PopupMenu mnuFill
End Sub

Private Sub picBackFill_Paint()
    picBackFill.Cls
    picBackFill.Line (-15, -15)-(picBackFill.ScaleWidth, picBackFill.ScaleHeight), , B
    lblFillStyle = mnuFillItem(picBackFill.FillStyle).Caption
End Sub


Private Sub picForeColour_Click()
    On Error Resume Next
    With dlgColour
        .CancelError = True
        .Color = mSelected.P.ForeColour
        .Flags = cdlCCFullOpen Or cdlCCRGBInit
        .ShowColor
        If Err = 0 Then
            picForeColour.BackColor = .Color
            If frmMain.IsSelected Then
                ItemChanged conUndoForeColourChange
                mSelected.P.ForeColour = .Color
                frmMain.ToolBoxDone False, False, False, False, False, False, True, False
            End If
        End If
    End With
End Sub


Private Sub picLineStyle_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If (Button And vbRightButton) Then
        PopupMenu mnuWidth, vbPopupMenuLeftButton Or vbPopupMenuRightButton
    Else
        PopupMenu mnuLine, vbPopupMenuLeftButton Or vbPopupMenuRightButton
    End If
End Sub


Private Sub picText_Click()
        On Error Resume Next
        With dlgColour
            .CancelError = True
            .Color = mSelected.P.TextColour
            .Flags = cdlCCFullOpen Or cdlCCRGBInit
            .ShowColor
            If Err = 0 Then
                picText.BackColor = .Color
                If frmMain.IsSelected Then
                    ItemChanged conUndoTextColourChange
                    mSelected.P.TextColour = .Color
                    frmMain.ToolBoxDone False, False, False, False, False, False, False, True
                End If
            End If
        End With

End Sub


