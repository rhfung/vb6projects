VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmFont 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Font Box"
   ClientHeight    =   765
   ClientLeft      =   5565
   ClientTop       =   645
   ClientWidth     =   3660
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   2  'Custom
   ScaleHeight     =   765
   ScaleWidth      =   3660
   ShowInTaskbar   =   0   'False
   Begin VB.ComboBox cboSize 
      Enabled         =   0   'False
      Height          =   315
      Left            =   2760
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   45
      Width           =   735
   End
   Begin MSComctlLib.Toolbar tlbFont 
      Height          =   330
      Left            =   90
      TabIndex        =   3
      Top             =   405
      Width           =   2280
      _ExtentX        =   4022
      _ExtentY        =   582
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Wrappable       =   0   'False
      Appearance      =   1
      Style           =   1
      ImageList       =   "ilsPic"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   7
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Bold"
            Description     =   "Bold"
            Object.ToolTipText     =   "Bold"
            ImageKey        =   "bold"
            Style           =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Italic"
            Description     =   "Italic"
            Object.ToolTipText     =   "Italic"
            ImageKey        =   "italic"
            Style           =   1
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Underline"
            Description     =   "Underline"
            Object.ToolTipText     =   "Underline"
            ImageKey        =   "underline"
            Style           =   1
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Left"
            Description     =   "Left"
            Object.ToolTipText     =   "Left"
            ImageKey        =   "left"
            Style           =   2
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Centre"
            Description     =   "Centre"
            Object.ToolTipText     =   "Centre"
            ImageKey        =   "centre"
            Style           =   2
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Right"
            Description     =   "Right"
            Object.ToolTipText     =   "Right"
            ImageKey        =   "right"
            Style           =   2
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ilsPic 
      Left            =   3240
      Top             =   720
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Flow_font.frx":0000
            Key             =   "bold"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Flow_font.frx":0114
            Key             =   "italic"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Flow_font.frx":0228
            Key             =   "underline"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Flow_font.frx":033C
            Key             =   "left"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Flow_font.frx":0450
            Key             =   "centre"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Flow_font.frx":0564
            Key             =   "right"
         EndProperty
      EndProperty
   End
   Begin VB.ComboBox cboScreenFonts 
      Enabled         =   0   'False
      Height          =   315
      ItemData        =   "Flow_font.frx":0678
      Left            =   300
      List            =   "Flow_font.frx":067A
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   45
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.CheckBox chkFontFace 
      Enabled         =   0   'False
      Height          =   255
      Left            =   60
      TabIndex        =   0
      Top             =   75
      Width           =   375
   End
End
Attribute VB_Name = "frmFont"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mblnCancel  As Boolean
Private mblnUpdating As Boolean
Private mblnBuilding As Boolean
 
Private Sub ItemChanged()
    On Error Resume Next
    If Not mblnUpdating And frmMain.IsSelected Then
        frmMain.ToolBoxChanged conUndoFontChanged
        If Not mblnBuilding Then 'building font list; can't change this
            If chkFontFace Then 'change font name
                frmMain.mSelected.P.FontFace = cboScreenFonts
            Else
                frmMain.mSelected.P.FontFace = ""
            End If
        End If
'        If chkSize Then 'change font size
            frmMain.mSelected.P.TextSize = CLng(cboSize)
'        Else
'            frmMain.mSelected.P.TextSize = frmMain.mFlowChart.FontSize 'all font sizes absolute
'        End If
        frmMain.mSelected.P.TextBold = tlbFont.Buttons!Bold.Value 'chkBold
        frmMain.mSelected.P.TextItalic = tlbFont.Buttons!Italic.Value 'chkItalic
        frmMain.mSelected.P.TextUnderline = tlbFont.Buttons!Underline.Value 'chkUnderline
        'SetTextFlags frmMain.mSelected, IIf(tlbFont.Buttons!Left.Value, DT_LEFT, 0) Or IIf(tlbFont.Buttons!Centre.Value, DT_CENTER, 0) Or IIf(tlbFont.Buttons!Right.Value, DT_RIGHT, 0)
        frmMain.mSelected.P.TextAlign = IIf(tlbFont.Buttons!Left.Value, DT_LEFT, 0) Or IIf(tlbFont.Buttons!Centre.Value, DT_CENTER, 0) Or IIf(tlbFont.Buttons!Right.Value, DT_RIGHT, 0)
        If Err Then 'beep on error
            Beep
        End If
        frmMain.ToolBoxDone
    ElseIf Not mblnUpdating Then
        Beep
    End If
End Sub

Public Sub Update()
    Dim lngFlags As Long
    
    mblnUpdating = True
    If frmMain.IsSelected And Not mblnBuilding Then
        'enable everything
        chkFontFace.Enabled = Not mblnBuilding
        cboScreenFonts.Enabled = Not mblnBuilding
        'chkSize.Enabled = Not mblnBuilding
        cboSize.Enabled = Not mblnBuilding
        tlbFontEnabled True
        'chkBold.Enabled = True
        'chkItalic.Enabled = True
        'chkUnderline.Enabled = True
        
        With frmMain.mSelected.P
            If Len(.FontFace) And cboScreenFonts.ListCount > 1 Then
                On Error Resume Next
                cboScreenFonts = .FontFace
                If Err Then
                    cboScreenFonts.ListIndex = -1 'don't select anything
                    Err.Clear
                    chkFontFace = vbUnchecked
                Else
                    chkFontFace = vbChecked
                End If
            Else
                chkFontFace = vbUnchecked
            End If

            On Error Resume Next
            cboSize = .TextSize
            If Err Then
                Err.Clear
                If .TextSize > conFontMin And .TextSize < conFontMax Then
                    cboSize.AddItem .TextSize
                    cboSize = .TextSize 'add and select proper size
                Else
                    cboSize.ListIndex = -1 'do not select anything
                End If
            End If
            
            tlbFont.Buttons!Bold.Value = IIf(.TextBold, tbrPressed, tbrUnpressed)
            tlbFont.Buttons!Italic.Value = IIf(.TextItalic, tbrPressed, tbrUnpressed)
            tlbFont.Buttons!Underline.Value = IIf(.TextUnderline, tbrPressed, tbrUnpressed)
            lngFlags = .TextAlign  'GetTextFlags(frmMain.mSelected)
            tlbFont.Buttons!Left.Value = IIf(lngFlags = DT_LEFT, tbrPressed, tbrUnpressed)
            tlbFont.Buttons!Centre.Value = IIf(lngFlags = DT_CENTER, tbrPressed, tbrUnpressed)
            tlbFont.Buttons!Right.Value = IIf(lngFlags = DT_RIGHT, tbrPressed, tbrUnpressed)
        End With
    Else
        On Error Resume Next
        'disable everything
        chkFontFace.Enabled = False
        cboScreenFonts.Enabled = False
        'chkSize.Enabled = False
        cboSize.Enabled = False
        tlbFontEnabled False
'        chkBold.Enabled = False
'        chkItalic.Enabled = False
'        chkUnderline.Enabled = False
    End If
    mblnUpdating = False
End Sub

Private Sub tlbFontEnabled(Enable As Boolean)
    Dim objButton As Button
    
    tlbFont.Enabled = Enable
    For Each objButton In tlbFont.Buttons
        objButton.Enabled = Enable
    Next objButton
End Sub


Private Sub cboScreenFonts_Change()
    chkFontFace = vbChecked
    ItemChanged
End Sub


Private Sub cboScreenFonts_Click()
    chkFontFace = vbChecked
    ItemChanged
End Sub


Private Sub cboSize_Change()
    'chkSize = vbChecked
    ItemChanged
End Sub


Private Sub cboSize_Click()
    'chkSize = vbChecked
    ItemChanged
End Sub



'Private Sub chkBold_Click()
'    ItemChanged
'End Sub
'
Private Sub chkFontFace_Click()
    ItemChanged
End Sub

'Private Sub chkItalic_Click()
'    ItemChanged
'End Sub
'
'Private Sub chkSize_Click()
'    ItemChanged
'End Sub

'Private Sub chkUnderline_Click()
'    ItemChanged
'End Sub
'
Private Sub Form_Activate()
    Update
End Sub

Private Sub Form_Load()
    Dim lngFontNo As Long
    
    
    If WindowState <> vbMinimized Then
        Move Screen.Width * 0.9 - Width, Screen.Height * 0.05
    End If
    
    gblnWindowFont = True
    Show 0, frmMain
    mblnCancel = False
    mblnBuilding = True
    
    'add font sizes
    For lngFontNo = 8 To 12
        cboSize.AddItem lngFontNo
    Next lngFontNo
    
    For lngFontNo = 14 To 20 Step 2
        cboSize.AddItem lngFontNo
    Next lngFontNo
    
    For lngFontNo = 24 To 40 Step 4
        cboSize.AddItem lngFontNo
    Next lngFontNo
    
    'add font face names
    For lngFontNo = 0 To Screen.FontCount - 1
        cboScreenFonts.AddItem Screen.Fonts(lngFontNo)
        DoEvents
        If mblnCancel Then Exit Sub
    Next lngFontNo
    
    cboScreenFonts.Visible = True
    mblnBuilding = False
    Update
    
'    For lngFontNo = 0 To Printer.FontCount - 1
'        cboPrinterFonts.AddItem Printer.Fonts(lngFontNo)
'    Next lngFontNo
End Sub




Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    mblnCancel = True
End Sub


Private Sub Form_Unload(Cancel As Integer)
    gblnWindowFont = False
End Sub


Private Sub tlbFont_ButtonClick(ByVal Button As MSComctlLib.Button)
    ItemChanged
End Sub

