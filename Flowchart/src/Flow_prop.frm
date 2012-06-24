VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmProp 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Flow Chart Options"
   ClientHeight    =   4575
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7290
   Icon            =   "Flow_prop.frx":0000
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   2  'Custom
   ScaleHeight     =   4575
   ScaleWidth      =   7290
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraArea 
      Caption         =   "Drawing Area"
      Height          =   975
      Left            =   4680
      TabIndex        =   18
      Top             =   2880
      Width           =   2535
      Begin VB.Label lblArea 
         Caption         =   "(1) Add an object (2) Go to the menu Shapes and select Area."
         Height          =   615
         Left            =   120
         TabIndex        =   19
         Top             =   240
         Width           =   2295
      End
   End
   Begin VB.Frame fraUnits 
      Caption         =   "Units"
      Height          =   1815
      Left            =   4680
      TabIndex        =   15
      Top             =   960
      Width           =   2535
      Begin VB.ComboBox cboUnits 
         Height          =   315
         ItemData        =   "Flow_prop.frx":000C
         Left            =   120
         List            =   "Flow_prop.frx":001F
         Style           =   2  'Dropdown List
         TabIndex        =   16
         Top             =   240
         Width           =   2295
      End
      Begin VB.Label Label1 
         Caption         =   "Not saved in the flow chart until version 8."
         Height          =   855
         Left            =   120
         TabIndex        =   17
         Top             =   720
         Width           =   2295
      End
   End
   Begin VB.Frame fraDefaults 
      Caption         =   "Defaults (for all flow charts)"
      Height          =   975
      Left            =   120
      TabIndex        =   10
      Top             =   2880
      Width           =   4455
      Begin VB.CheckBox chkAutoEditText 
         Alignment       =   1  'Right Justify
         Caption         =   "Edit text on add:"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   600
         Width           =   1695
      End
      Begin VB.CommandButton cmdExt 
         Caption         =   "Set Extensions"
         Height          =   375
         Left            =   3000
         TabIndex        =   13
         Top             =   240
         Width           =   1335
      End
      Begin VB.ComboBox cboMargin 
         Height          =   315
         ItemData        =   "Flow_prop.frx":0059
         Left            =   1560
         List            =   "Flow_prop.frx":006F
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label lblM 
         Caption         =   "&Margins (inches):"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3998
      TabIndex        =   1
      Top             =   4080
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   2198
      TabIndex        =   0
      Top             =   4080
      Width           =   1095
   End
   Begin VB.Frame fraObj 
      Caption         =   "Objects"
      Height          =   735
      Left            =   4680
      TabIndex        =   8
      Top             =   120
      Width           =   2535
      Begin VB.Label lblObj 
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   2295
      End
   End
   Begin VB.Frame fraVer 
      Caption         =   "&Version"
      Height          =   1815
      Left            =   120
      TabIndex        =   5
      Top             =   960
      Width           =   4455
      Begin VB.ListBox lstVer 
         Height          =   1425
         ItemData        =   "Flow_prop.frx":0095
         Left            =   240
         List            =   "Flow_prop.frx":00B4
         Sorted          =   -1  'True
         TabIndex        =   6
         Top             =   240
         Width           =   735
      End
      Begin VB.Label lblVerInfo 
         Height          =   1455
         Left            =   1080
         TabIndex        =   7
         Top             =   240
         Width           =   3255
      End
   End
   Begin VB.Frame fraFont 
      Caption         =   "Default Font"
      Height          =   735
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   4455
      Begin VB.CommandButton cmdFontChange 
         Caption         =   "&Change"
         Height          =   375
         Left            =   3480
         TabIndex        =   4
         Top             =   240
         Width           =   855
      End
      Begin VB.Label lblFont 
         Height          =   375
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   3255
      End
   End
   Begin MSComDlg.CommonDialog dlgFile 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      FilterIndex     =   1
   End
End
Attribute VB_Name = "frmProp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'July 3, 2001

Private mFile       As FlowChart
Private mRegistry   As PRegistry
Private mstrFontName As String
Private mcurFontSize As Currency
Private msngVersion As Single

Public Sub LoadProperties(ByVal File As FlowChart, ByVal Registry As PRegistry)
    Dim intItem As Integer
    Dim sngMargin As Single
    
    Set mFile = File
    Set mRegistry = Registry
    mstrFontName = mFile.FontName
    mcurFontSize = mFile.FontSize
    msngVersion = mFile.Version
    sngMargin = mRegistry.Margin
    
    For intItem = 0 To cboMargin.ListCount - 1
        If CSng(cboMargin.List(intItem)) = sngMargin Then
            cboMargin.ListIndex = intItem
            Exit For
        End If
    Next intItem
    chkAutoEditText = IIf(mRegistry.AutoEditText, vbChecked, vbUnchecked)
    
    Dim intL As Integer
    
    For intL = 0 To cboUnits.ListCount - 1
        If cboUnits.ItemData(intL) = File.UnitScale Then
            cboUnits.ListIndex = intL
            Exit For
        End If
    Next intL
    
    Dim objI As FlowItem
    
    For Each objI In File
        If objI.Number = conAddExtra1 Then 'area
            lblArea = "To change or remove area, click on the area whitespace and delete area object."
            Exit For
        End If
    Next objI
    
    Update
End Sub



'data from var to window
Private Sub Update()
    Const conVerFormat = "#0.0#"
    
    lblFont = mstrFontName & ", size " & mcurFontSize
    lblObj = mFile.Count & IIf(mFile.Count = 1, " object", " objects")
    
    On Error Resume Next
    lstVer = Format(msngVersion, conVerFormat)
    If Err Or lstVer.ListIndex = -1 Then 'version not there
        Err.Clear
        lstVer.AddItem Format(msngVersion, conVerFormat)
        lstVer = Format(msngVersion, conVerFormat)
    End If
End Sub


Private Sub cmdCancel_Click()
    Unload Me
End Sub



Private Sub cmdExt_Click()
    RegisterFlowFile
End Sub

Private Sub cmdFontChange_Click()
    On Error Resume Next
    With dlgFile
        .Flags = cdlCFBoth Or cdlCFLimitSize Or cdlCFWYSIWYG Or cdlCFScalableOnly Or cdlCFForceFontExist
        .Min = conFontMin
        .Max = conFontMax
        'only FontName and FontSize
        'can be changed globally for the
        'document, and only FontName
        'is the same in any zoom percentage.
        .FontName = mstrFontName
        .FontSize = mcurFontSize
        .FontBold = False
        .FontItalic = False
        .FontUnderline = False
            .ShowFont
        If Err = cdlCancel Then 'cancelled
            Exit Sub
        ElseIf Err <> 0 Then
            MsgBox "Error showing Font dialog box.", vbExclamation, "Font"
        Else
            mstrFontName = .FontName
            mcurFontSize = .FontSize
            Update
        End If
    End With
End Sub

Private Sub cmdOK_Click()
    If msngVersion < mFile.Version Then
        If MsgBox("Some formatting may be lost.  Do you want to change the version?", vbYesNo Or vbQuestion) = vbYes Then
            mFile.Version = msngVersion
        End If
    Else 'no problem moving to higher version
        mFile.Version = msngVersion
    End If
    mFile.FontName = mstrFontName
    mFile.FontSize = mcurFontSize
    mFile.UnitScale = cboUnits.ItemData(cboUnits.ListIndex)
    If cboMargin.ListIndex <> -1 Then
        mRegistry.Margin = CSng(cboMargin)
    End If
    mRegistry.AutoEditText = CBool(chkAutoEditText)
    
    frmMain.SetView FontName:=mstrFontName, FontSize:=mcurFontSize
    Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set mFile = Nothing
    Set mRegistry = Nothing
End Sub


Private Sub lstVer_Click()
    Dim intAdd As Integer
    
    If lstVer.ListIndex > -1 Then 'item selected
        msngVersion = CSng(lstVer)
        On Error Resume Next
        intAdd = msngVersion '# version
        If msngVersion >= 5.1 And msngVersion < 6 Then intAdd = 15 'mid version
        lblVerInfo = LoadResString(100 + intAdd)
        If Err Then 'no version info found
            Err.Clear
            lblVerInfo = "No information is available for the selected version."
        End If
    End If
End Sub


