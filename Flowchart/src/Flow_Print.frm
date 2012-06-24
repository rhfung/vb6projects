VERSION 5.00
Begin VB.Form frmPrint 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Print"
   ClientHeight    =   3855
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5430
   ClipControls    =   0   'False
   Icon            =   "Flow_Print.frx":0000
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3855
   ScaleWidth      =   5430
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraRange 
      Caption         =   "Print Range"
      Height          =   1335
      Left            =   120
      TabIndex        =   8
      Top             =   1800
      Width           =   3255
      Begin VB.OptionButton opSelection 
         Caption         =   "&Selection"
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   960
         Width           =   975
      End
      Begin VB.TextBox txtTo 
         Height          =   285
         Left            =   2520
         TabIndex        =   14
         Top             =   600
         Width           =   495
      End
      Begin VB.TextBox txtFrom 
         Height          =   285
         Left            =   1560
         TabIndex        =   12
         Top             =   600
         Width           =   495
      End
      Begin VB.OptionButton opRange 
         Caption         =   "Ran&ge"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   600
         Width           =   975
      End
      Begin VB.OptionButton opAll 
         Caption         =   "&All"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   975
      End
      Begin VB.CheckBox chkSame 
         Caption         =   "&On same page"
         Height          =   255
         Left            =   1440
         TabIndex        =   24
         ToolTipText     =   "If option not selected, will print each layer on different pages"
         Top             =   960
         Width           =   1575
      End
      Begin VB.Label lblTo 
         Alignment       =   1  'Right Justify
         Caption         =   "&to:"
         Height          =   255
         Left            =   2160
         TabIndex        =   13
         Top             =   600
         Width           =   255
      End
      Begin VB.Label lblFrom 
         Alignment       =   1  'Right Justify
         Caption         =   "&from:"
         Height          =   255
         Left            =   960
         TabIndex        =   11
         Top             =   600
         Width           =   495
      End
   End
   Begin VB.Frame fraOrientation 
      Caption         =   "Orientation"
      Height          =   1335
      Left            =   120
      TabIndex        =   16
      Top             =   1800
      Visible         =   0   'False
      Width           =   3255
      Begin VB.OptionButton opOLandscape 
         Caption         =   "&Landscape"
         Height          =   255
         Left            =   360
         TabIndex        =   18
         Top             =   840
         Width           =   2055
      End
      Begin VB.OptionButton opOPortrait 
         Caption         =   "P&ortrait"
         Height          =   255
         Left            =   360
         TabIndex        =   17
         Top             =   360
         Width           =   2055
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4200
      TabIndex        =   23
      Top             =   3360
      Width           =   975
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   3000
      TabIndex        =   22
      Top             =   3360
      Width           =   975
   End
   Begin VB.Frame fraCopies 
      Caption         =   "Copies"
      Height          =   1335
      Left            =   3600
      TabIndex        =   19
      Top             =   1800
      Width           =   1695
      Begin VB.TextBox txtCopies 
         Height          =   285
         Left            =   240
         TabIndex        =   21
         Top             =   480
         Width           =   735
      End
      Begin VB.Label lblC 
         Caption         =   "Number of &copies:"
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.Frame fraPrinter 
      Caption         =   "Printer"
      Height          =   1575
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5175
      Begin VB.CommandButton cmdProp 
         Caption         =   "&Properties"
         Height          =   375
         Left            =   3720
         TabIndex        =   3
         Top             =   210
         Width           =   1335
      End
      Begin VB.ComboBox cboName 
         Height          =   315
         Left            =   960
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   240
         Width           =   2655
      End
      Begin VB.Label lblWhere 
         Height          =   255
         Left            =   960
         TabIndex        =   7
         Top             =   1080
         Width           =   3855
      End
      Begin VB.Label lblW 
         Caption         =   "Where:"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   1080
         Width           =   735
      End
      Begin VB.Label lblType 
         Height          =   255
         Left            =   960
         TabIndex        =   5
         Top             =   720
         Width           =   3855
      End
      Begin VB.Label lblT 
         Caption         =   "Driver:"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   720
         Width           =   735
      End
      Begin VB.Label lblN 
         Caption         =   "&Name:"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   735
      End
   End
End
Attribute VB_Name = "frmPrint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'frmPrint or PrintDlg


'2 May 2003
'15-16 February 2002

'By Richard Fung
'ItemChanged returns True if failed.
'Update         changes window
'ItemChanged    changes variables/objects

'1. Load
'2. Initialize or InitializeSetup 'returns True if Good

Private Declare Function PrinterProperties Lib "winspool.drv" (ByVal hwnd As Long, ByVal hPrinter As Long) As Long
Private Declare Function OpenPrinter Lib "winspool.drv" Alias "OpenPrinterA" (ByVal pPrinterName As String, phPrinter As Long, pDefault As Long) As Long
Private Declare Function ClosePrinter Lib "winspool.drv" (ByVal hPrinter As Long) As Long

Public Enum EnumPrintRange
    conPrintAll = 1
    conPrintRange = 2
    conPrintSelection = 4
End Enum

Public PrintRange   As EnumPrintRange 'in/out
Public FromPage     As Integer 'in/out
Public ToPage       As Integer 'in/out
Public Copies       As Integer 'in/out value saved on PrinterChanged
Public Orientation  As Integer 'in/out value saved on PrinterChanged

Private m_lngPrintOptions As EnumPrintRange 'in
Private m_intMin          As Integer 'in
Private m_intMax          As Integer 'in
Private m_blnPrintOnSamePage As Boolean 'out

Private mblnOK      As Boolean
Private mblnSetup   As Boolean
Private Const conPrintSetup = "Print Setup"
Private Const conMinCopies = 1
Private Const conMaxCopies = 10

Private Function fraCopiesChanged() As Boolean
    On Error Resume Next
    Copies = txtCopies
    If Err Then fraCopiesChanged = True: Err.Clear
    If Copies < conMinCopies Or Copies > conMaxCopies Then fraCopiesChanged = True
End Function

Private Sub fraCopiesUpdate()
    txtCopies = Copies
End Sub

Private Sub fraOrientationUpdate()
    Select Case Orientation
    Case vbPRORLandscape
        opOLandscape = True
    Case vbPRORPortrait
        opOPortrait = True
    End Select
End Sub

Private Sub fraPrinterUpdate()
    If ListIndex = -1 Then Exit Sub
    lblType = GetPrinter(ListIndex).DriverName
    lblWhere = GetPrinter(ListIndex).Port
End Sub

Private Function fraRangeChanged() As Boolean
    If opAll Then PrintRange = conPrintAll
    If opRange Then PrintRange = conPrintRange
    If opSelection Then PrintRange = conPrintSelection
    If opRange Or txtFrom.Enabled Then
        On Error Resume Next
        FromPage = txtFrom
        ToPage = txtTo
        If opRange Then 'only generate errors on Range selected
            If Err Then fraRangeChanged = True: Err.Clear
            If Not GetInRange(FromPage) Then fraRangeChanged = True
            If Not GetInRange(ToPage) Then fraRangeChanged = True
            If FromPage > ToPage Then fraRangeChanged = True
        End If
    End If
    lblFrom.Enabled = opRange
    txtFrom.Enabled = opRange
    txtFrom = FromPage
    lblTo.Enabled = opRange
    txtTo.Enabled = opRange
    txtTo = ToPage
End Function

Private Sub fraRangeUpdate()
    opAll = (PrintRange And conPrintAll) = conPrintAll
    opRange = (PrintRange And conPrintRange) = conPrintRange
    opSelection = (PrintRange And conPrintSelection) = conPrintSelection
    
    lblFrom.Enabled = opRange
    txtFrom.Enabled = opRange
    txtFrom = FromPage
    lblTo.Enabled = opRange
    txtTo.Enabled = opRange
    txtTo = ToPage
End Sub


Private Function GetInRange(ByVal Number As Integer) As Boolean
    GetInRange = CBool(Min <= Number) And CBool(Number <= Max)
End Function

Private Function GetPrinter(ByVal Index As Integer) As Printer
    Set GetPrinter = Printers(Index)
End Function

Private Function ListIndex() As Integer
    ListIndex = cboName.ListIndex
End Function

Public Function Initialize(ByVal Options As EnumPrintRange, ByVal Range As EnumPrintRange, Optional ByVal FromPage As Integer = 1, Optional ByVal ToPage As Integer = 1, Optional ByVal Min As Integer = 1, Optional ByVal Max As Integer = 1, Optional ByVal Copies As Integer) As Boolean
    'Options - combine using Or
    'Range  - choose one only
    
    'returns True if Good
    cmdProp.Enabled = False
    
    PrintOptions = Options
    PrintRange = Range
    Me.Min = Min
    Me.Max = Max
    If GetInRange(FromPage) Then Me.FromPage = FromPage Else Me.FromPage = Min
    If GetInRange(ToPage) Then Me.ToPage = ToPage Else Me.ToPage = Max
    If conMinCopies <= Copies And Copies <= conMaxCopies Then Me.Copies = Copies
    
    fraPrinterUpdate
    fraRangeUpdate
    fraCopiesUpdate
    Initialize = (cboName.ListCount > 0)
End Function

Public Function InitializeSetup(ByVal Orientation As Integer) As Boolean
    'returns True if good
    cmdProp.Enabled = True
    
    Caption = conPrintSetup
    mblnSetup = True
    fraPrinterUpdate
    fraRange.Visible = False
    fraOrientation.Visible = True
    Me.Orientation = Orientation
    fraOrientationUpdate
    fraCopies.Visible = False
    InitializeSetup = (cboName.ListCount > 0)
End Function

Private Sub ItemChanged()
'Main change code to update Printer.
    Set Printer = GetPrinter(ListIndex)
    If mblnSetup Then
        If opOPortrait Then
            Printer.Orientation = vbPRORPortrait
        Else
            Printer.Orientation = vbPRORLandscape
        End If
        'WARNING: Bugs if not Visual Basic 6 SP5
        'and orientation is changed.
    Else
        Printer.Copies = Copies
    End If
End Sub

Public Property Let Max(ByVal pMax As Integer)
    m_intMax = pMax
End Property

Private Property Get Max() As Integer
    Max = m_intMax
End Property

Public Property Let Min(ByVal pMin As Integer)
    m_intMin = pMin
End Property

Private Property Get Min() As Integer
    Min = m_intMin
End Property

Private Property Let PrintOnSamePage(ByVal pPrn As Boolean)
     m_blnPrintOnSamePage = pPrn
End Property

Public Property Get PrintOnSamePage() As Boolean
    PrintOnSamePage = m_blnPrintOnSamePage
End Property

Public Property Let PrintOptions(ByVal pOptions As EnumPrintRange)
    m_lngPrintOptions = pOptions

    opAll.Enabled = (pOptions And conPrintAll) = conPrintAll
    opRange.Enabled = (pOptions And conPrintRange) = conPrintRange
    opSelection.Enabled = (pOptions And conPrintSelection) = conPrintSelection

End Property

Private Property Get PrintOptions() As EnumPrintRange
    PrintOptions = m_lngPrintOptions
End Property

Public Function ShowModal(ByVal OwnerForm As Form) As Boolean
'returns True if Ok
    Show vbModal, OwnerForm
    ShowModal = mblnOK
'form is unloaded afterwards
End Function

Private Sub cboName_Click()
    fraPrinterUpdate
End Sub

Private Sub chkSame_Click()
    m_blnPrintOnSamePage = chkSame
End Sub

Private Sub cmdCancel_Click()
    mblnOK = False
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim blnFail As Boolean
    
    If mblnSetup Then
        ItemChanged
        mblnOK = True
        Unload Me
    Else
        If fraRangeChanged Then
            MsgBox "The page range is outside the pages from " & Min & " to " & Max & ", or are not integers.", vbExclamation, Caption
            blnFail = True
        Else
            If fraCopiesChanged Then
                MsgBox "The number of copies is outside of the range from 1 to 10 or is not an integer.", vbExclamation, Caption
                blnFail = True
            End If
            If Not blnFail Then
                ItemChanged
                mblnOK = True
                Unload Me
            End If
        End If
    End If
'    Debug.Print String(50, "-")
'    Debug.Print "PrintRange = ", PrintRange
'    Debug.Print "FromPage = ", FromPage
'    Debug.Print "ToPage = ", ToPage
'    Debug.Print "Copies = ", Copies
'    Debug.Print "Orientation = ", Orientation
'    Debug.Print "Setup = ", mblnSetup
'    Debug.Print "OK = ", mblnOk
End Sub



Private Sub cmdProp_Click()
    Dim r As Long
    Dim hwndPrinter As Long
    
    r = OpenPrinter(CStr(cboName), hwndPrinter, 0&)
    r = PrinterProperties(hwnd, hwndPrinter)
    If r = 0 Then
        MsgBox "Unable to show printer properties.", vbExclamation
    End If
    r = ClosePrinter(hwndPrinter)
End Sub

Private Sub Form_Load()
    Dim intPrn As Integer
    Dim strDeviceName As String
    
    Copies = 1
    mblnOK = False
    Initialize 0, 0, 1, 1, 1, 1, Copies
    
    'assumption: there is only one DeviceName
    For intPrn = 0 To Printers.Count - 1
        strDeviceName = GetPrinter(intPrn).DeviceName
        cboName.AddItem strDeviceName
        If Printer.DeviceName = strDeviceName Then cboName.ListIndex = intPrn
    Next intPrn
End Sub



Private Sub opAll_Click()
    fraRangeChanged
End Sub


Private Sub opRange_Click()
    fraRangeChanged
End Sub


Private Sub opSelection_Click()
    fraRangeChanged
End Sub


