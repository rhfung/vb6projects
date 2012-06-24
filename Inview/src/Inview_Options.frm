VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmOptions 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Options"
   ClientHeight    =   3390
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   Icon            =   "Inview_Options.frx":0000
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3390
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame fraGeneral 
      BorderStyle     =   0  'None
      Height          =   2175
      Left            =   240
      TabIndex        =   16
      Top             =   480
      Visible         =   0   'False
      Width           =   4215
      Begin VB.Frame fraGen1 
         Caption         =   "Temporary Files"
         Height          =   1335
         Left            =   0
         TabIndex        =   19
         Top             =   600
         Width           =   2775
         Begin VB.OptionButton opTempApp 
            Caption         =   "Location of this program"
            Height          =   195
            Left            =   240
            TabIndex        =   22
            Top             =   960
            Width           =   2175
         End
         Begin VB.OptionButton opTempWin 
            Caption         =   "Windows\Temp"
            Height          =   255
            Left            =   240
            TabIndex        =   21
            Top             =   600
            Width           =   1935
         End
         Begin VB.OptionButton opTempDefault 
            Caption         =   "Best choice"
            Height          =   255
            Left            =   240
            TabIndex        =   20
            Top             =   240
            Width           =   1935
         End
      End
      Begin VB.CommandButton cmdClear 
         Caption         =   "Clear Settings"
         Height          =   375
         Left            =   2400
         TabIndex        =   18
         Top             =   120
         Width           =   1575
      End
      Begin VB.CheckBox chkNoSaveSetting 
         Caption         =   "Do not save settings"
         Height          =   375
         Left            =   120
         TabIndex        =   17
         Top             =   120
         Width           =   2055
      End
   End
   Begin VB.Frame fraText 
      BorderStyle     =   0  'None
      Height          =   2175
      Left            =   240
      TabIndex        =   6
      Top             =   480
      Visible         =   0   'False
      Width           =   4215
      Begin VB.CommandButton cmdTextFont 
         Caption         =   "..."
         Height          =   255
         Left            =   3720
         TabIndex        =   9
         Top             =   240
         Width           =   375
      End
      Begin VB.HScrollBar hsbTextSize 
         Height          =   255
         LargeChange     =   2
         Left            =   1200
         Max             =   20
         Min             =   10
         TabIndex        =   11
         Top             =   600
         Value           =   10
         Width           =   1455
      End
      Begin VB.CheckBox chkTextBold 
         Caption         =   "B&old"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   960
         Width           =   855
      End
      Begin VB.Label lblTextFont 
         Height          =   255
         Left            =   1200
         TabIndex        =   8
         Top             =   240
         Width           =   2415
      End
      Begin VB.Label Label2 
         Caption         =   "Fon&t:"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label3 
         Caption         =   "Font &size:"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   600
         Width           =   855
      End
      Begin VB.Label lblTextSize 
         Height          =   255
         Left            =   2760
         TabIndex        =   12
         Top             =   600
         Width           =   975
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   2513
      TabIndex        =   15
      Top             =   2880
      Width           =   975
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   1193
      TabIndex        =   14
      Top             =   2880
      Width           =   975
   End
   Begin VB.Frame fraBinary 
      BorderStyle     =   0  'None
      Height          =   2175
      Left            =   240
      TabIndex        =   1
      Top             =   480
      Visible         =   0   'False
      Width           =   4215
      Begin VB.CheckBox chkBinBold 
         Caption         =   "B&old"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   600
         Width           =   855
      End
      Begin VB.HScrollBar hsbBinSize 
         Height          =   255
         LargeChange     =   2
         Left            =   1200
         Max             =   20
         Min             =   10
         TabIndex        =   3
         Top             =   240
         Value           =   10
         Width           =   1455
      End
      Begin VB.Label lblBinSize 
         Height          =   255
         Left            =   2760
         TabIndex        =   4
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Font &size:"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   855
      End
   End
   Begin MSComctlLib.TabStrip tab1 
      Height          =   2655
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   4683
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   3
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "&Binary"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "T&ext"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Curren&t Session"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog dlgFont 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      FontSize        =   10
      Max             =   20
      Min             =   10
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mlngTemp    As Long


Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdClear_Click()
    gData.ClearData
End Sub

Private Sub cmdOK_Click()
    Dim r As Long
    
    gData.BinBold = chkBinBold
    gData.BinFontSize = hsbBinSize
    
    gData.TextBold = chkTextBold
    gData.TextFontName = lblTextFont
    gData.TextFontSize = hsbTextSize
    
    gData.NoSaveSetting = chkNoSaveSetting
    gData.SaveData 'save here because values did change
    
    Select Case True
    Case opTempDefault
        gstrPath = "" 'will rebuild itself
        mlngTemp = 1
    Case opTempWin 'see GetTempFile()
        gstrPath = Space(1024) 'put temp files in Temp directory if possible
        r = GetTempPath(1024, gstrPath)
        If r > 1024 Or r = 0 Then 'error - place temp file in App.Path
            gstrPath = App.Path
        Else 'good
            gstrPath = Left$(gstrPath, InStr(1, gstrPath, vbNullChar) - 1)
        End If
        mlngTemp = 2
    Case opTempApp
        gstrPath = App.Path
        mlngTemp = 3
    End Select
    
    Unload Me
End Sub

Private Sub cmdTextFont_Click()
    On Error Resume Next
    With dlgFont
        .Flags = cdlCFForceFontExist Or cdlCFLimitSize Or cdlCFANSIOnly Or cdlCFBoth Or cdlCFWYSIWYG Or cdlCFScalableOnly
        .FontName = lblTextFont
        .FontSize = hsbTextSize
        .FontBold = chkTextBold
        .ShowFont
        If Err = 0 Then 'set properties
            lblTextFont = .FontName
            hsbTextSize = .FontSize
            lblTextSize = .FontSize
            chkTextBold = IIf(.FontBold, vbChecked, vbUnchecked)
        Else
            Err.Clear
        End If
    End With
End Sub


Private Sub Form_Load()
    hsbBinSize = gData.BinFontSize
    lblBinSize = gData.BinFontSize
    chkBinBold = IIf(gData.BinBold, vbChecked, vbUnchecked)
    
    lblTextFont = gData.TextFontName
    hsbTextSize = gData.TextFontSize
    lblTextSize = gData.TextFontSize
    chkTextBold = IIf(gData.TextBold, vbChecked, vbUnchecked)
    
    chkNoSaveSetting = IIf(gData.NoSaveSetting, vbChecked, vbUnchecked)
    
    Select Case mlngTemp
    Case 0, 1
        opTempDefault = True
        opTempDefault.ToolTipText = gstrPath
    Case 2
        opTempWin = True
        opTempWin.ToolTipText = gstrPath
    Case 3
        opTempApp = True
        opTempApp.ToolTipText = gstrPath
    End Select
    
    tab1_Click
End Sub


Private Sub hsbBinSize_Change()
    lblBinSize = hsbBinSize
End Sub


Private Sub hsbBinSize_Scroll()
    lblBinSize = hsbBinSize
End Sub


Private Sub hsbTextSize_Change()
    lblTextSize = hsbTextSize
End Sub


Private Sub hsbTextSize_Scroll()
    lblTextSize = hsbTextSize
End Sub


Private Sub tab1_Click()
    fraText.Visible = False
    fraBinary.Visible = False
    fraGeneral.Visible = False

    Select Case tab1.SelectedItem.Index
    Case 1 'binary
        fraBinary.Visible = True
    Case 2 'text
        fraText.Visible = True
    Case 3
        fraGeneral.Visible = True
    End Select
End Sub


