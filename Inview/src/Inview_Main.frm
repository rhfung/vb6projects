VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   Caption         =   "In View"
   ClientHeight    =   5100
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   6990
   ClipControls    =   0   'False
   Icon            =   "Inview_Main.frx":0000
   LockControls    =   -1  'True
   OLEDropMode     =   1  'Manual
   ScaleHeight     =   5100
   ScaleWidth      =   6990
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdNewCancel 
      Caption         =   "Skip"
      Height          =   255
      Left            =   6360
      TabIndex        =   10
      Top             =   0
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.CommandButton cmdNewSet 
      Caption         =   "Set"
      Height          =   255
      Left            =   5520
      TabIndex        =   9
      Top             =   0
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox txtNewEdit 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   4440
      TabIndex        =   8
      Top             =   0
      Visible         =   0   'False
      Width           =   975
   End
   Begin MSComDlg.CommonDialog dlgFile 
      Left            =   2640
      Top             =   2040
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.HScrollBar hsbBinView 
      Height          =   255
      Left            =   1200
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   3960
      Width           =   4095
   End
   Begin VB.VScrollBar vsbBinView 
      Height          =   3135
      Left            =   5280
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   960
      Width           =   255
   End
   Begin VB.PictureBox picBinView 
      ClipControls    =   0   'False
      Height          =   3495
      Left            =   960
      ScaleHeight     =   3435
      ScaleWidth      =   4395
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   840
      Visible         =   0   'False
      Width           =   4455
      Begin VB.PictureBox picBin 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   3015
         Left            =   0
         ScaleHeight     =   3015
         ScaleWidth      =   3855
         TabIndex        =   3
         Top             =   0
         Width           =   3855
         Begin VB.TextBox txtBinEdit 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   480
            TabIndex        =   7
            Top             =   720
            Visible         =   0   'False
            Width           =   255
         End
      End
   End
   Begin VB.TextBox txt 
      Height          =   3735
      Left            =   240
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   1
      Top             =   600
      Visible         =   0   'False
      Width           =   5295
   End
   Begin MSComctlLib.TabStrip tabView 
      Height          =   4935
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   6975
      _ExtentX        =   12303
      _ExtentY        =   8705
      MultiRow        =   -1  'True
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   4
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Unicode"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "&Hexadecimal"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "&Decimal"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "&Character"
            ImageVarType    =   2
         EndProperty
      EndProperty
      OLEDropMode     =   1
   End
   Begin VB.Label lblMessage 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   0
      TabIndex        =   6
      Top             =   4800
      Width           =   3735
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileNew 
         Caption         =   "&New"
      End
      Begin VB.Menu mnuFileOpen 
         Caption         =   "&Open..."
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuFileSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileSave 
         Caption         =   "&Save"
         Index           =   0
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuFileSave 
         Caption         =   "Save &As..."
         Index           =   1
      End
      Begin VB.Menu mnuFileSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
      Begin VB.Menu mnuEditConvert 
         Caption         =   "&Convert..."
      End
      Begin VB.Menu mnuEditCopyB 
         Caption         =   "Copy &Bytes..."
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuEditSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditFindA 
         Caption         =   "Find &ASCII..."
         Index           =   0
      End
      Begin VB.Menu mnuEditFindA 
         Caption         =   "Find Next ASCII"
         Index           =   1
         Shortcut        =   {F2}
      End
      Begin VB.Menu mnuEditSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditFindW 
         Caption         =   "Find &Unicode..."
         Index           =   0
      End
      Begin VB.Menu mnuEditFindW 
         Caption         =   "Find Next Unicode"
         Index           =   1
         Shortcut        =   {F3}
      End
      Begin VB.Menu mnuEditSep3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditFind 
         Caption         =   "&Find..."
         Index           =   0
      End
      Begin VB.Menu mnuEditFind 
         Caption         =   "Find Next"
         Index           =   1
         Shortcut        =   {F4}
      End
      Begin VB.Menu mnuEditSep5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFindCaseInsensitive 
         Caption         =   "Find case insensitive..."
         Index           =   0
      End
      Begin VB.Menu mnuEditSep4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditFindByte 
         Caption         =   "Find byte..."
         Index           =   0
         Shortcut        =   {F6}
      End
      Begin VB.Menu mnuEditFindByte 
         Caption         =   "Find and replace byte..."
         Index           =   1
         Shortcut        =   +{F6}
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "&View"
      Begin VB.Menu mnuViewBlock 
         Caption         =   "Nothing"
         Enabled         =   0   'False
         Index           =   1
      End
   End
   Begin VB.Menu mnuTools 
      Caption         =   "&Tools"
      Begin VB.Menu mnuToolsProp 
         Caption         =   "&Properties..."
      End
      Begin VB.Menu mnuToolsOpt 
         Caption         =   "&Options..."
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "&About"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Richard Fung, Programming Applications 1

'variables to store implemented file objects
Private mlngFileType    As EnumViews
Private mTextFile       As TextFile
Private mBinFile        As BinFile

'drawing data for binary viewer
Private mlngDataWidth   As Long '# of characters
Private mlngDataWUsed   As Long '# of characters used
Private mlngSoftCols    As Long 'column # changes with the view
Private mlngLastPos     As Long 'last ch. position for drawing
'for picBay in the binary viewer
Private msngTextWidth   As Single 'full box width
Private msngTextChWidth As Single 'single character width
Private msngTextHeight  As Single 'text height
'file save
Private mblnLoading     As Boolean
Private mblnChanged     As Boolean

'this creates a 8 px border inside the tab container
Private Const conBorder = 120 '8 px

'The following views correspond to the tab index values
'on the form.
Private Enum EnumViews
    conNoFile = 0
    conText = 1
    conHex = 2
    conDec = 3
    conCharacter = 4
'    conReference = 5
End Enum

Private Type TRC
    Row As Single
    Col As Single
End Type

''The following two are for the binary view.
'Const conRow = 1
'Const conCol = 2

'These are the colours for the binary view.
Const conColourBlue = &H800000 '&H800000&
Const conColourGreen = &H8000& '&H8000&
Const conColourRed = &H80& '&H80&
Const conColourGray = &H808080 '&H808080&
Const conColourOrange = &H80FF& '&H80FF&
'Const conColourLightGray = &HC0C0C0  '&HC0C0C0&
'Const conColourPurple = &H800080  '&H800080&

Private Const conVsbAdjust As Long = 5

'This is the rectangle structure for the binary view.
Private Type TPRect
    Left    As Long 'twips
    Top     As Long
    Right   As Long
    Bottom  As Long
End Type


'Binary viewer: turns off the edit box
'   Hide=False is implemented for tab switching only
Private Sub BinEditOff(Optional ByVal Hide As Boolean = True)
    Dim bytItem As Byte
    
    On Error GoTo ResumeHere
    'if the edit box is visible and the last position
    'in the binary view is not zero then
    If txtBinEdit.Visible And mlngLastPos <> 0 Then
        'choose the view to determine the type
        'of data that has been set & convert
        Select Case View
            Case conHex
                'to convert to hexadecimal, Visual Basic
                'needs the &H in front of the hex value,
                'which the user should have returned
                'because of the tab name
                bytItem = CByte("&H" & txtBinEdit)
            Case conDec
                'convert the number the user typed in the
                'edit box to a byte value
                bytItem = CByte(txtBinEdit)
            Case conCharacter
                'take the first character in the text box
                'to an ASCII value, which correlates
                'to a byte value
                bytItem = Asc(txtBinEdit)
        End Select
ResumeHere:
        If Err Then
            Beep 'notifies user the change was unsuccessful
        Else
            'set the binary data to the file stream
            mBinFile.Item(mlngLastPos) = bytItem
            'set change flag
            mblnChanged = True
        End If
        If Hide Then 'Hide if passed to the function
            txtBinEdit.Visible = False
        End If
    Else
        If Hide Then 'Hide if passed to the function
            txtBinEdit.Visible = False
        End If
    End If
End Sub

'Binary Viewer: turns the edit box on
'     if a KeyAscii value is present, it will be set
'     to the edit box
Private Sub BinEditOn(Optional KeyAscii As Integer)
    Dim tRect As TPRect
    
    'if the position in the view is not 0 then
    If mlngLastPos <> 0 Then
        'get the selected item position
        tRect = GetChrRect(mlngLastPos)
        'move the box to the position with an indent
        'on the left for blank spaces -- items are
        'right aligned in decimal and hexadecimal
        txtBinEdit.Move tRect.Left + (mlngDataWidth - mlngDataWUsed) * msngTextChWidth, tRect.Top, tRect.Right - tRect.Left - msngTextChWidth * (mlngDataWidth - mlngDataWUsed), tRect.Bottom - tRect.Top
        txtBinEdit.MaxLength = mlngDataWUsed
        'choose the format the data will be presented in
        Select Case View
            Case conHex
                'hexadecimal view displays characters
                'using the Hex() function provided with VB
                txtBinEdit = Hex(mBinFile.Item(mlngLastPos))
            Case conDec
                'decimal view displays characters as it is
                'read from Item() in Byte data.  The Byte data
                'is automatically converted to Decimal format
                'by Visual Basic
                txtBinEdit = mBinFile.Item(mlngLastPos)
            Case conCharacter
                'character view displays the ASCII character corresponding
                'to the byte value 0 - 255, when possible, using VB's Chr()
                If mBinFile.Item(mlngLastPos) > 31 Then 'displayable characters are 32 and greater
                    txtBinEdit = Chr(mBinFile.Item(mlngLastPos))
                ElseIf KeyAscii = 0 Then 'no replacement character
                    'cannot edit non-visible characters because text box
                    'will not allow these characters
                    Beep
                    txtBinEdit.Visible = False
                    lblMessage = "Use the hexadecimal or decimal viewers to edit."
                    Exit Sub
                End If
        End Select
        If KeyAscii Then
            'place the typed character into the text box
            txtBinEdit = Chr(KeyAscii)
            'set the edit cursor to the start of the text
            txtBinEdit.SelStart = Len(txtBinEdit)
        Else
            'select all text that was retrieved in the
            'Select Case decision routine
            txtBinEdit.SelStart = 0
            txtBinEdit.SelLength = Len(txtBinEdit)
        End If
        'show the edit box and begin edit
        txtBinEdit.Visible = True
        txtBinEdit.SetFocus
        lblMessage = "Editing"
    End If
End Sub

Private Sub CharDraw(Pos As Long)
    Dim tRect As TPRect
    Dim bytNext As Byte
    Dim strDisp As String

    'See Also: picBin_Paint
    tRect = GetChrRect(Pos)
    bytNext = mBinFile(Pos)
    
    If bytNext = 13 Or bytNext = 10 Then 'CR and LF are green
        picBin.ForeColor = conColourGreen
    ElseIf bytNext < 32 Then 'not displayable characters (now blocks) are gray
        picBin.ForeColor = conColourGray
    Else 'display other characters in black
        picBin.ForeColor = vbBlack
    End If
    
    If bytNext = 13 Then 'if CR, display paragraph mark
        strDisp = "¶"
    ElseIf bytNext = 10 Then 'if LF (from CR-LF new line combination), put an unobtrusive dot
        strDisp = "·"
    ElseIf bytNext < 32 Then 'not space, but not displayable then show a block
        strDisp = "" 'that is - cannot be displayed properly
    Else 'displayable character
        strDisp = Chr$(bytNext)
    End If
    
    picBin.DrawMode = vbCopyPen
    picBin.Line (tRect.Left, tRect.Top)-(tRect.Right, tRect.Bottom), vbWindowBackground, BF
    picBin.CurrentX = tRect.Left: picBin.CurrentY = tRect.Top
    picBin.Print strDisp;
End Sub

Private Property Get CurrentRow() As Long
    If vsbBinView > 0 Then
        CurrentRow = vsbBinView * conVsbAdjust
    End If
End Property

Private Property Let CurrentRow(ByVal r As Long)
    If r < 0 Then 'behind the beginning of the file
        vsbBinView = 0
    ElseIf (r / conVsbAdjust) < vsbBinView.Max Then 'in limit
        vsbBinView = r / conVsbAdjust
    Else 'out of limit
        vsbBinView = vsbBinView.Max
        lblMessage = "Display: File is TOO LARGE!"
    End If
End Property

Public Sub DoFindReplace(ByVal FindPart As String, ByVal ReplacePart As String, ByVal Replace As Boolean)
'November 12, 2003
    Dim intByte As Integer
    Dim lngPos  As Long
    Dim lngElement As Long
    Dim lower   As Long
    Dim lngLenSequence As Long
    'Dim strInput As String
    Dim vntInput As Variant
    Dim bytInput() As Byte
    'Dim strOutput As String
    Dim vntOutput As Variant
    Dim bytOutput() As Byte
    Dim blnFlag   As Boolean
    
    On Error GoTo BadConversion
    ''get input
    'strInput = InputBox$("Please enter in a byte or sequence of bytes seperated by spaces to find.", "Find Byte")
    'If Len(strInput) = 0 Then Exit Sub
    
    'validate input
    vntInput = Split(FindPart, " ")
    lower = LBound(vntInput)
    lngLenSequence = UBound(vntInput) - LBound(vntInput) + 1
    
    ReDim bytInput(lower To UBound(vntInput))
    ReDim bytOutput(lower To UBound(vntInput))
    
    For lngElement = LBound(vntInput) To UBound(vntInput)
        If Not IsNumeric(vntInput(lngElement)) Then
            MsgBox "One of the entries is not a number.", vbInformation, "Find Byte"
            Exit Sub
        End If
        bytInput(lngElement) = CInt(vntInput(lngElement))
    Next lngElement
    
    If lngLenSequence < 1 Then
        MsgBox "Nothing to do.", vbInformation, "Find Byte"
        Exit Sub
    End If
    

    If Replace Then
        ''get output
        'strOutput = InputBox$("Please enter in " & lngLenSequence & " bytes to replace **without asking**.", "Replace Byte")
        'If Len(strOutput) = 0 Then Exit Sub
        
        
        'validate output
        vntOutput = Split(ReplacePart, " ")
        If UBound(vntOutput) <> UBound(vntInput) Or LBound(vntOutput) <> LBound(vntInput) Then
            MsgBox "Length of the replace byte sequence does not match that of input.", vbInformation, "Replace Byte"
            Exit Sub
        End If
        
        For lngElement = LBound(vntOutput) To UBound(vntOutput)
            If Not IsNumeric(vntOutput(lngElement)) Then
                MsgBox "One of the entries is not a number.", vbInformation, "Replace Byte"
                Exit Sub
            End If
            bytOutput(lngElement) = CInt(vntOutput(lngElement))
        Next lngElement
    End If
    
    On Error GoTo SeekReplace
    For lngPos = mlngLastPos + 1 To mBinFile.FileLength - lngLenSequence
        'does first byte match?
        If mBinFile.Item(lngPos) = bytInput(lower) Then
            blnFlag = True
            'check each element
            For lngElement = lower To UBound(bytInput)
                If mBinFile.Item(lngPos + lngElement - lower) <> bytInput(lngElement) Then
                    blnFlag = False
                End If
            Next lngElement
            
            If blnFlag Then 'we have a match
                If Replace Then
                    'overwrite data
                    For lngElement = lower To UBound(bytInput)
                        mBinFile.Item(lngPos + lngElement - lower) = bytOutput(lngElement)
                    Next lngElement
                    Debug.Print "Overwite occured at " & lngPos
                    lngPos = lngPos + lngLenSequence - 1
                    mblnChanged = True
                Else
                    'found a match
                    SetCursor lngPos
                    Exit Sub
                End If
            End If
        End If
    Next lngPos
    
    picBin.Refresh
    
    
    Exit Sub
BadConversion:
    MsgBox "Bad conversion from the numeric entry to an integer.", vbInformation, "Find Byte"
    Exit Sub
SeekReplace:
    MsgBox "Problem while looking for the byte sequence at position " & lngPos & ".  Please report " & Err.Number & " to Richard Fung.", vbInformation, "Find Byte"
    Exit Sub

End Sub

Private Property Get Filename() As String
    If Not mBinFile Is Nothing Then
        Filename = mBinFile.Filename
    ElseIf Not mTextFile Is Nothing Then
        Filename = mTextFile.Filename
    End If
End Property

Private Sub FormResize2()
    On Error Resume Next
    hsbBinView.Visible = False
    hsbBinView.Visible = False
    
    If picBin.Width > picBinView.ScaleWidth Then
        hsbBinView.Enabled = True
        hsbBinView.Max = picBin.Width - picBinView.ScaleWidth
        hsbBinView.LargeChange = picBinView.ScaleWidth
        hsbBinView.SmallChange = msngTextWidth
    Else
        hsbBinView.Enabled = False
    End If
    'hsbBinView.SmallChange = 1000 '/ ((picBin.Width - picBinView.ScaleWidth) / picBinView.ScaleWidth) / (TextWidth("O") * mlngDataWidth)
    picBin.Height = picBinView.ScaleHeight
    
    Err.Clear
    vsbBinView.Min = 0
    vsbBinView.Max = (mBinFile.FileLength \ mlngSoftCols - GetRowsVisible() + conVsbAdjust + 1) / conVsbAdjust
    vsbBinView.Enabled = (vsbBinView.Max > vsbBinView.Min)
    
    vsbBinView.LargeChange = Fix(GetRowsVisible() / conVsbAdjust)
    vsbBinView.SmallChange = 1
    
    hsbBinView.Visible = True
    hsbBinView.Visible = True
End Sub

Private Function IsASCII(ByVal B As Byte) As Boolean
    '0 to 9, A to Z, a to z
    'IsASCII = (48 <= B And B <= 57) Or (65 <= B And B <= 90) Or (97 <= B And B <= 122)
    IsASCII = (32 <= B And B <= 125)
End Function

Private Function mnuFileOpenClick2() As Boolean
    On Error Resume Next
    
    If QuerySave = False Then Exit Function
    
    With dlgFile
        .CancelError = True
        .Flags = cdlOFNFileMustExist Or cdlOFNHideReadOnly
        .Filter = "Binary Files (*.*)|*.*|Unicode Text Files (*.txt;*.log)|*.txt;*.log"
        .ShowOpen
        If Err = cdlCancel Then
            mnuFileOpenClick2 = False
            Exit Function
        ElseIf Err <> 0 Then
            MsgBox "Error showing open dialog box!", vbExclamation
            mnuFileOpenClick2 = False
        Else
            If .FilterIndex = 1 Then
                OpenBin .Filename
            Else
                OpenText .Filename
            End If
            mnuFileOpenClick2 = True 'did work
        End If
    End With
End Function

Private Sub mnuViewBuild(ByVal Blocks As Long)
    Dim lngI As Long
    
    For lngI = mnuViewBlock.UBound To 2 Step -1
        Unload mnuViewBlock(lngI)
    Next lngI
    
    If Blocks = 0 Then
        mnuViewBlock(1).Caption = "Nothing"
        mnuViewBlock(1).Enabled = False
        mnuViewBlock(1).Checked = False
    Else
        mnuViewBlock(1).Caption = "Block 1"
        mnuViewBlock(1).Enabled = True
        For lngI = 2 To Blocks
            Load mnuViewBlock(lngI)
            mnuViewBlock(lngI).Caption = "Block " & lngI
            mnuViewBlock(lngI).Enabled = True
        Next lngI
    End If
End Sub

'OpenFile opens a file and chooses the views
'that will be used.
Private Sub OpenBin(ByVal Filename As String)
    'create new file objects
    Set mBinFile = New BinFile
    Set mTextFile = Nothing
    
    'turn on loading flag (so data is not set)
    mblnLoading = True
    'the OpenFile() functions of the file objects
    'return error numbers
'    Select Case mTextFile.OpenFile(FileName)
'        Case 62 'input past end of file
'            'file could be too big
'            Set mTextFile = Nothing
'        Case Is <> 0 'not equals 0
'            'do not know problem
'            MsgBox "Error loading file.  " & Err.Description & " (" & Err.Number & ")", vbExclamation, "Text File Open"
'            Set mTextFile = Nothing
'    End Select
    
    If mBinFile.OpenFile(Filename) Then
        'file could not be opened
        MsgBox "Error loading file.  " & Err.Description & " (" & Err.Number & ")", vbExclamation, "Open File"
        Set mBinFile = Nothing
        mlngFileType = conNoFile
    Else
        mlngFileType = conHex
    End If
    
    mlngLastPos = 0 'reset last position; for binary viewer
    Caption = conTitle & " [" & Filename & "]"
    
    'select the default tab
    'if not a binary viewer
    If Not IsABinaryView Then
         'force the viewer, character hex is usually good
         Set tabView.SelectedItem = tabView.Tabs(conCharacter)
    Else
         UpdateView 'refresh the current binary view
    End If
    mnuViewBuild 0
     
    mblnChanged = False 'reset changed flag
    mblnLoading = False 'turn off loading flag
End Sub

Private Sub OpenText(ByVal Filename As String)
    Dim lngE As Long
    
    'create new file objects
    Set mBinFile = Nothing
    Set mTextFile = New TextFile
    
    'turn on loading flag (so data is not set)
    mblnLoading = True
    
    'focus on the proper tab
    Set tabView.SelectedItem = tabView.Tabs(conText)
    
    'the OpenFile() functions of the file objects
    lngE = mTextFile.OpenFile(Filename)
'    If lngE = -4 Then 'see TextFile code
'        MsgBox "This file is larger than 32 KB for ASCII files or 64 KB for Unicode files.  " & conTitle & " cannot open such large files.", vbInformation, "Open File"
'        mnuFileNew_Click
'        Exit Sub
    
    mlngLastPos = 0 'reset last position; for binary viewer
    
    If lngE <> 0 Then
        'file could not be opened
        MsgBox "Error loading file.  " & Err.Description & " (" & Err.Number & ")", vbExclamation, "Open File"
        mlngFileType = conNoFile
        Set mTextFile = Nothing
        UpdateView
    Else
        mlngFileType = conText
    
        'load tab name
        If mTextFile.IsUnicode Then
            tabView.Tabs(conText).Caption = "Unicode"
        Else
            tabView.Tabs(conText).Caption = "ASCII"
'            If mTextFile.FileLength <= 32000 Then
'                If MsgBox("Since you are editing an ASCII file, you should use Rich Pad." & vbNewLine & "Do you want to run Rich Pad or Notepad now?", vbQuestion Or vbYesNo Or vbDefaultButton2, "Open File") = vbYes Then
'                    On Error Resume Next
'                    Shell "Richpad.exe " & Filename, vbNormalFocus
'                    If Err Then
'                        Err.Clear
'                        Shell "Notepad.exe " & Filename, vbNormalFocus
'                        If Err Then
'                            Err.Clear
'                            MsgBox "Could not run Rich Pad or Notepad.", vbExclamation, "Open File"
'                        Else
'                            mnuFileNew_Click
'                            Exit Sub
'                        End If
'                    Else
'                        mnuFileNew_Click
'                        Exit Sub
'                    End If
'                End If
'            End If
        End If
        
        Caption = conTitle & " [" & Filename & "]"
        'load text
        txt.Text = mTextFile.Text
        If txt.Text <> mTextFile.Text Then 'did not load correctly
            txt.Locked = True
            MsgBox "The text file cannot be edited because this version does not" & vbNewLine & "support unicode characters.  Unknown characters are" & vbNewLine & "denoted by question marks (?)." & vbNewLine & "Press F1 to see the value of the character.", vbInformation, "Open File"
            Caption = Caption & " (Read-Only)"
            mTextFile.CanSave = False
        ElseIf mTextFile.Blocks > 0 Then
            'cannot edit partitioned file
            txt.Locked = True
            mTextFile.CanSave = False
            Caption = Caption & " (Read-Only)"
        Else
            txt.Locked = False
            mTextFile.CanSave = True
        End If
    
        'show the text viewer
        Set tabView.SelectedItem = tabView.Tabs(conText)
        mnuViewBuild mTextFile.Blocks
        UpdateView 'refresh the viewer at whatever it is at
    End If
    
    mblnChanged = False 'reset changed flag
    mblnLoading = False 'turn off loading flag
End Sub

'Formats a number with commas but no decimal points.
'Not to be confused with FormatNumber in VB6
Private Function FormatNum2(ByVal Number As Long) As String
    FormatNum2 = Format(Number, "#,##0")
End Function


'Gets the viewable area of the binary edit.
Private Function GetBinViewable() As TPRect 'viewable rectangle
    With GetBinViewable
        .Left = -picBin.Left
        .Top = -picBin.Top
        .Right = .Left + picBinView.ScaleWidth
        .Bottom = .Top + picBinView.ScaleHeight
    End With
End Function


'Gets the rectangular dimensions of a character's display area.
Private Function GetChrRect(ByVal Position As Long) As TPRect
    'Dim lngRowCol() As Long
    Dim tRowCol As TRC
    Dim lngRow As Long, lngCol As Long

    tRowCol = GetChrRowCol(Position)
    'these variables easier to work with
    lngRow = tRowCol.Row
    lngCol = tRowCol.Col
    
    With GetChrRect
        .Left = lngCol * msngTextWidth
        .Top = (lngRow - CurrentRow) * msngTextHeight
        .Right = .Left + msngTextWidth
        .Bottom = .Top + msngTextHeight
    End With
End Function

'Gets the position number of a character in the file
'from the X and Y values in the view.
Private Function GetChrPos(ByVal X As Long, ByVal Y As Long) As Long
    Dim lngRow As Long, lngCol As Long

    'figures out the nearest column and row
    'based on X and Y values using integer division
    lngCol = X \ msngTextWidth
    lngRow = Y \ msngTextHeight + CurrentRow
    
    GetChrPos = lngRow * mlngSoftCols + lngCol
    'if the position is beyond the file dimensions,
    'then clear tthe character position
    If Not (GetChrPos >= 1 And GetChrPos <= mBinFile.FileLength + 1) Then
        GetChrPos = 0 'no position
    ElseIf (X Mod msngTextWidth) \ msngTextChWidth < mlngDataWidth - mlngDataWUsed Then
        GetChrPos = 0 'no position
    End If
End Function


'Returns the row and column of a character position number.
'#If VB6 Then 'VB6 can return Long() arrays.
'Private Function GetChrRowCol(Position As Long) As Long()
'#Else 'VB5 can only return arrays through the Variant data type.
'Private Function GetChrRowCol(ByVal Position As Long) As Variant
'#End If
Private Function GetChrRowCol(ByVal Position As Long) As TRC
    'Dim lngReturn(conRow To conCol) As Long
    
    'using the position in a file, it determines
    'the position's row and column in the binary view
    'lngReturn(conRow) = Position \ mlngSoftCols
    'lngReturn(conCol) = Position Mod mlngSoftCols
    'an array always returns using left-array = right-source-array
    'GetChrRowCol = lngReturn

    'guard against divide by 0
    If mlngSoftCols > 0 Then
        GetChrRowCol.Row = Position \ mlngSoftCols
        GetChrRowCol.Col = Position Mod mlngSoftCols
    End If
End Function


'Gets the number of rows visible based on
'the view's dimensions.
Private Function GetRowsVisible() As Long
    GetRowsVisible = picBinView.ScaleHeight \ msngTextHeight
End Function

'Determines if a binary view is selected.
Private Function IsABinaryView() As Boolean
    IsABinaryView = (View = conDec) Or (View = conHex) Or (View = conCharacter)
End Function

Private Function QuerySave() As Boolean
    If mblnChanged = False Then
        QuerySave = True
    Else
        Select Case MsgBox("The current file has been changed.  Do you want to save this file?", vbQuestion Or vbYesNoCancel)
        Case vbYes
            mnuFileSave_Click 0
            QuerySave = Not mblnChanged
        Case vbNo
            QuerySave = True
        Case vbCancel
            QuerySave = False
        End Select
    End If
End Function

'This selects a binary item based on the provided position in the
'file.  The Scrolled value moves the selection rectangle to
'the newer position.
Private Sub SelectBinItem(ByVal Position As Long, Optional ByVal Scrolled As Boolean)
    Dim tRect As TPRect
    
    picBin.DrawMode = vbInvert
    
    If mlngLastPos <> 0 And (mlngLastPos <> Position Or Scrolled = False) Then
        BinEditOff
        tRect = GetChrRect(mlngLastPos)
        picBin.Line (tRect.Left + (mlngDataWidth - mlngDataWUsed) * msngTextChWidth, tRect.Top)-(tRect.Right, tRect.Bottom), , BF
    End If
    
    If Position > 0 Then 'And Position <= mBinFile.FileLength Then
        tRect = GetChrRect(Position)
        picBin.Line (tRect.Left + (mlngDataWidth - mlngDataWUsed) * msngTextChWidth, tRect.Top)-(tRect.Right, tRect.Bottom), , BF
        If Position <= mBinFile.FileLength Then
            lblMessage = "Position: " & FormatNum2(Position) & "    Byte Value: Dec=" & mBinFile.Item(Position) & "  Hex=" & Hex(mBinFile.Item(Position)) & "  Chr=" & Chr(mBinFile.Item(Position))
        Else
            lblMessage = "Position: after EOF"
        End If
    Else
        lblMessage = ""
    End If
    picBin.CurrentX = tRect.Right
    picBin.CurrentY = tRect.Top
    mlngLastPos = Position
End Sub

Public Sub SetFont()
     'sets up font sizes
     With gData
        picBin.Font.Name = .BinFontName
        picBin.Font.Size = .BinFontSize
        picBin.Font.Bold = .BinBold
        txtBinEdit.Font.Name = .BinFontName
        txtBinEdit.Font.Size = .BinFontSize
        txtBinEdit.Font.Bold = .BinBold
        txt.Font.Name = .TextFontName
        txt.Font.Size = .TextFontSize
        txt.Font.Bold = .TextBold
     End With
End Sub

Private Sub SetCursor(ByVal Start As Long)
    Dim lngRow As Long
        
        lngRow = GetChrRowCol(Start).Row
    If Not (CurrentRow <= lngRow And lngRow <= CurrentRow + vsbBinView.LargeChange * conVsbAdjust) Then 'not in view
        CurrentRow = lngRow
    End If
    SelectBinItem Start
End Sub

Private Sub txtPosition()
    Dim lngSelStart As Long
    
    If mlngFileType = conText Then
        If mTextFile.Blocks = 0 Then
            lngSelStart = txt.SelStart
        Else
            If mTextFile.IsUnicode Then
                lngSelStart = mTextFile.GetRealPos(txt.SelStart * 2) / 2
            Else
                lngSelStart = mTextFile.GetRealPos(txt.SelStart)
            End If
        End If
        If mTextFile.IsUnicode Then
            lblMessage = "Position: " & FormatNum2(lngSelStart) & "     Byte: " & FormatNum2(lngSelStart * 2)
        Else
            lblMessage = "Position: " & FormatNum2(lngSelStart)
        End If
    End If
End Sub

'Called by UpdateView() and was separated for clarity
Private Sub UpdateBinView()
    'calculate text widths and heights
    msngTextChWidth = picBin.TextWidth("0")
    msngTextWidth = msngTextChWidth * mlngDataWidth
    msngTextHeight = picBin.TextHeight("0")

    If Not mBinFile Is Nothing Then
        picBin.Width = msngTextWidth * mlngSoftCols 'fixed width
        'picBin.Height = (mBinFile.FileLength \ mlngSoftCols + 1) * msngTextHeight
    End If
    
    'reset the bin area position to 0
    picBin.Move 0, 0
    
    FormResize2
    
    'reset scroll bars
    hsbBinView = 0
    CurrentRow = GetChrRowCol(mlngLastPos).Row - conVsbAdjust
    
    'show it & it should redraw automatically
    picBinView.Visible = True
End Sub


'When a tab is changed, this is called.  It
'refreshes the data for the current view.
Private Sub UpdateView()
    'if some objects are missing, like when no file
    'is loaded, show nothing, do nothing
    If mlngFileType = conNoFile Then
        picBinView.Visible = False
        txt.Visible = False
        lblMessage = ""
    ElseIf mlngFileType = conText And Not mTextFile Is Nothing And tabView.SelectedItem.Index = conText Then
        picBinView.Visible = False
        txt.Visible = True
        txt_Click
    ElseIf Not mBinFile Is Nothing And tabView.SelectedItem.Index <> conText Then
        'clear out message
        lblMessage = ""
        
        'hide all viewers
        picBinView.Visible = False
        txt.Visible = False
        'set bits and pieces based on the current View
        Select Case View
        Case conHex
            mlngDataWidth = 3 'like FF_
            mlngDataWUsed = 2
            mlngSoftCols = gData.HexCols
            UpdateBinView
        Case conDec
            mlngDataWidth = 4 'like 255_
            mlngDataWUsed = 3
            mlngSoftCols = gData.DecCols
            UpdateBinView
        Case conCharacter
            mlngDataWidth = 1 'like A or C
            mlngDataWUsed = 1
            mlngSoftCols = gData.ChCols
            UpdateBinView
        End Select
    Else
        picBinView.Visible = False
        txt.Visible = False
        lblMessage = "This tab is unavailable."
    End If
   
    'show/hide scroll bars
    'following picBinView
    hsbBinView.Visible = picBinView.Visible
    vsbBinView.Visible = picBinView.Visible
End Sub

'Gets the view for the tabView index.
Private Property Get View() As EnumViews
    View = tabView.SelectedItem.Index
End Property



Private Sub Form_Load()
     Dim strLoad As String
     Dim blnB    As Boolean
     Dim blnU    As Boolean
     
     'loads registry data
     Set gData = New IVRegistry 'load data object
     SetFont
     'show default tab
     tabView_Click
     'updates the view
     UpdateView
     
     If Len(Command$) Then
        Show
        strLoad = Command$
        'check for binary flag
        If Left$(strLoad, 3) = "/b " Or Left$(strLoad, 3) = "/B " Then
            strLoad = Mid(strLoad, 4)
            blnB = True
        'check for unicode text flag
        ElseIf Left$(strLoad, 3) = "/u " Or Left$(strLoad, 3) = "/U " Then
            strLoad = Mid$(strLoad, 4)
            blnU = True
        End If
        'remove extra quotation marks
        If Left$(strLoad, 1) = """" And Right$(strLoad, 1) = """" Then
            strLoad = Mid$(strLoad, 2, Len(strLoad) - 2)
        End If
        'open up
        If Len(strLoad) Then
            If blnB Then
                OpenBin strLoad
            ElseIf blnU Then 'ascii and unicode text
                OpenText strLoad
            Else 'ask and auto-detect
                dlgFile.Filename = strLoad
                If Not mnuFileOpenClick2 Then
                    dlgFile.Filename = ""
                End If
            End If
        End If
     End If
End Sub

Private Sub Form_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Data.GetFormat(vbCFFiles) Then
        dlgFile.Filename = Data.Files(1)
        Effect = vbDropEffectCopy
        If Not mnuFileOpenClick2 Then
           dlgFile.Filename = "" 'clear this
        End If
    Else
        Effect = vbDropEffectNone
    End If
End Sub

Private Sub Form_OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single, State As Integer)
    If Data.GetFormat(vbCFFiles) Then
        Effect = vbDropEffectCopy
    Else
        Effect = vbDropEffectNone
    End If
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If QuerySave = False Then
        Cancel = 1
    End If
End Sub


'scroll bar code, position controls code
Private Sub Form_Resize()
    On Error Resume Next
    
    If WindowState <> vbMinimized Then
        'common
        lblMessage.Move 0, ScaleHeight - lblMessage.Height, ScaleWidth
        tabView.Move 0, 0, ScaleWidth, ScaleHeight - lblMessage.Height
        'text viewer
        txt.Move tabView.ClientLeft + conBorder, tabView.ClientTop + conBorder, tabView.ClientWidth - conBorder * 2, tabView.ClientHeight - conBorder * 2
        'binary viewer
        picBinView.Move tabView.ClientLeft + conBorder, tabView.ClientTop + conBorder, tabView.ClientWidth - conBorder * 2 - vsbBinView.Width, tabView.ClientHeight - conBorder * 2 - hsbBinView.Height
        hsbBinView.Move picBinView.Left, picBinView.Top + picBinView.Height, picBinView.Width
        vsbBinView.Move picBinView.Left + picBinView.Width, picBinView.Top, vsbBinView.Width, picBinView.Height
        
        If picBinView.Visible Then
            Err.Clear
            FormResize2
        End If
    End If
End Sub


Private Sub Form_Unload(Cancel As Integer)
    Set gData = Nothing 'registry data
    Set mBinFile = Nothing
    Set mTextFile = Nothing
End Sub


'since horizontal scroll bar stores screen twips,
'it is possible just to move the position of the
'view box and Windows will paint it
Private Sub hsbBinView_Change()
    picBin.Left = -(hsbBinView)
End Sub

Private Sub mnuEdit_Click()
    mnuEditConvert.Enabled = mlngFileType <> conText And mlngFileType <> conNoFile And mlngLastPos > 0
    mnuEditFindA(0).Enabled = (mlngFileType = conHex) And IsABinaryView
    mnuEditFindA(1).Enabled = mnuEditFindA(0).Enabled
    mnuEditFindW(0).Enabled = mnuEditFindA(0).Enabled
    mnuEditFindW(1).Enabled = mnuEditFindA(0).Enabled
    mnuEditFindByte(0).Enabled = mnuEditFindA(0).Enabled
    mnuEditFindByte(1).Enabled = mnuEditFindA(0).Enabled
    
    mnuEditFind(0).Enabled = (mlngFileType = conHex And IsABinaryView) Or (mlngFileType = conText And View = conText)
    mnuEditFind(1).Enabled = mnuEditFind(0).Enabled
    
    
    'mnuEditCopy.Enabled = (mlngFileType = conHex) And IsABinaryView And (mlngLastPos <> 0)
    mnuEditCopyB.Enabled = IsABinaryView And (mlngLastPos > 0) 'mnuEditCopy.Enabled
End Sub

Private Sub mnuEditConvert_Click()
    Load frmConvert
    frmConvert.Initialize mBinFile, mlngLastPos
    frmConvert.Show vbModal, Me
    
End Sub

'Private Sub mnuEditCopy_Click()
''copies the selected character
'    Dim strCopy As String
'
'    strCopy = Chr(mBinFile.Item(mlngLastPos))
'
'    Clipboard.Clear
'    Clipboard.SetText strCopy
'
'    If Clipboard.GetText(vbCFText) <> strCopy Then
'        MsgBox "The character was not copied successfully.  Windows Clipboard does not like the data.", vbExclamation
'    End If
'End Sub
'
Private Sub mnuEditCopyB_Click()
'copys a length of bytes starting at the selected position
    Static lngLen   As Long
    Dim strBuild    As String
    Dim lngItem     As Long
    Dim lngMax      As Long
    Dim strPath     As String
    Dim F           As Long
    
    On Error GoTo Handler
    If lngLen = 0 Then lngLen = 1
    lngLen = Val(InputBox("How many bytes do you want to copy?", "Copy", lngLen))
    
    If lngLen > 0 Then
        MousePointer = vbHourglass
        If mlngLastPos + lngLen - 1 > mBinFile.FileLength Then
            lngMax = mBinFile.FileLength
        Else
            lngMax = mlngLastPos + lngLen - 1
        End If
        For lngItem = mlngLastPos To lngMax
            strBuild = strBuild & Chr(mBinFile.Item(lngItem))
        Next lngItem
        Clipboard.Clear
        Clipboard.SetText strBuild
        
        If Clipboard.GetText(vbCFText) <> strBuild Then
            MsgBox "The byte string was not copied successfully.  Windows Clipboard does not like the data.", vbExclamation
            
            'write clipboard to HD
            If Len(gstrPath) Then
                strPath = gstrPath
            Else
                strPath = App.Path
            End If
            If Right$(strPath, 1) <> "\" Then strPath = strPath & "\"
            F = FreeFile
            Open strPath & "Clipboard_" & App.EXEName & ".tmp" For Output As #F
            Print #F, strBuild;
            Close #F
        End If
        MousePointer = vbDefault
    End If
    
    Exit Sub
Handler:
    Exit Sub
End Sub

Private Sub mnuEditFind_Click(Index As Integer)
    Static strFind    As String
    Dim lngCounter    As Long
    Dim bytSearchA()  As Byte
    Dim bytSearchW()  As Byte
    Dim bytFound      As Byte
    Dim blnFoundA     As Boolean
    Dim blnFoundW     As Boolean
    Dim lngLenW       As Long
    Dim lngPos        As Long
    Dim lngStart      As Long 'found start
    Dim lngBeginS     As Long 'begin search point
    Dim strAppend     As String
    
    On Error GoTo Handler
    If mlngFileType = conHex Then strAppend = "  This searches both ASCII and Unicode text."
    If Len(strFind) = 0 Then 'find next
        strFind = InputBox("Type in the text that you want to find." & strAppend, "Find", strFind)
    ElseIf Index = 0 Then
        strFind = InputBox("Type in the text that you want to find." & strAppend, "Find", strFind)
    End If
    
    MousePointer = vbHourglass
    
    
    If Len(strFind) And mlngFileType = conHex Then
        lblMessage = "Searching for ASCII/Unicode text..."
        lblMessage.Refresh
        
        'assumption: array start at 0
        bytSearchA = StrConv(strFind, vbFromUnicode)
        bytSearchW = strFind
        lngLenW = LenB(strFind)
        
        lngBeginS = 1
        'find next and there is a position, start after that match
        If mlngLastPos > 0 And Index <> 0 Then lngBeginS = mlngLastPos + Len(strFind)
        
        
        For lngPos = lngBeginS To mBinFile.FileLength
            bytFound = mBinFile.Item(lngPos)
            
            If bytFound = bytSearchA(0) Then
                'test ASCII
                blnFoundA = True
                For lngCounter = 1 To (lngLenW \ 2) - 1
                    If mBinFile.Item(lngPos + lngCounter) <> bytSearchA(lngCounter) Then
                        blnFoundA = False
                        Exit For
                    End If
                Next lngCounter
                
                If Not blnFoundA Then
                    'test unicode
                    blnFoundW = True
                    For lngCounter = 1 To lngLenW - 1
                        If mBinFile.Item(lngPos + lngCounter) <> bytSearchW(lngCounter) Then
                            blnFoundW = False
                            Exit For
                        End If
                    Next lngCounter
                End If
                If blnFoundA Or blnFoundW Then
                    lngStart = lngPos
                    Exit For
                End If
            End If
        Next lngPos
        
        If lngStart > 0 Then
                SetCursor lngStart
            If blnFoundA Then
                lblMessage = "Found ASCII match at " & FormatNum2(lngStart) & " for """ & strFind & """."
            ElseIf blnFoundW Then
                lblMessage = "Found Unicode match at " & FormatNum2(lngStart) & " for """ & strFind & """."
            End If
         Else
            MsgBox "Text not found.", vbInformation
            lblMessage = "Text not found."
        End If
    ElseIf Len(strFind) And mlngFileType = conText Then
        lngBeginS = 1
        If txt.SelStart + 1 > 0 And Index <> 0 Then 'find next, not at start
            lngBeginS = txt.SelStart + Len(strFind) + 1
        End If
        
        lngStart = InStr(lngBeginS, txt.Text, strFind, vbTextCompare)
        If lngStart > 0 Then
            txt.SelStart = lngStart - 1
            txt.SelLength = Len(strFind)
        ElseIf txt.Locked Then 'this txt was not changed
            lngStart = InStr(lngBeginS, mTextFile.Text, strFind, vbTextCompare)
            If lngStart > 0 And lngStart <= Len(txt.Text) Then 'within range of txt
                txt.SelStart = lngStart - 1
                txt.SelLength = Len(strFind)
            Else
                MsgBox "Text not found.", vbInformation, conTitle & "."
                lblMessage = "Text not found."
            End If
        Else
            MsgBox "Text not found.", vbInformation
            lblMessage = "Text not found."
        End If
    End If
    MousePointer = vbDefault
    Exit Sub
Handler:
    MousePointer = vbDefault
    Exit Sub
End Sub

Private Sub mnuEditFindA_Click(Index As Integer)
    Static lngL    As Long  'For binary files
    Dim lngCounter As Long
    Dim lngPos     As Long
    Dim lngStart   As Long 'found start
    Dim lngBeginS  As Long 'begin search point
    
    On Error GoTo Handler
    If lngL = 0 Then 'find next
        lngL = 10
        lngL = Abs(InputBox("How many ASCII characters should be in a line before you consider it ASCII?", "Find ASCII", lngL))
    ElseIf Index = 0 Then
        lngL = Abs(InputBox("How many ASCII characters should be in a line before you consider it ASCII?", "Find ASCII", lngL))
    End If
    
    MousePointer = vbHourglass
    If lngL <> 0 Then
        lblMessage = "Searching for ASCII..."
        lblMessage.Refresh
        lngBeginS = 1
        'find next and there is a position, start after that match
        If mlngLastPos > 0 And Index <> 0 Then lngBeginS = mlngLastPos + lngL
        
        For lngPos = lngBeginS To mBinFile.FileLength
            If IsASCII(mBinFile.Item(lngPos)) Then
                lngCounter = lngCounter + 1
                If lngCounter >= lngL Then
                    'pattern matched
                    lngStart = lngPos - lngCounter + 1
                    Exit For
                End If
            Else
                lngCounter = 0
            End If
        Next lngPos
        
        If lngStart > 0 Then
                SetCursor lngStart
            lblMessage = "An ASCII pattern was found at " & FormatNum2(lngStart) & "."
         Else
            MsgBox "No ASCII pattern match found.", vbInformation
            lblMessage = "No ASCII pattern match found."
        End If
    End If
    MousePointer = vbDefault
    Exit Sub
Handler:
    MousePointer = vbDefault
    Exit Sub
End Sub
                                  
Private Sub mnuEditFindByte_Click(Index As Integer)
    Load frmFindByte
    
    frmFindByte.chkReplace = IIf(Index, vbChecked, vbUnchecked)
    
    frmFindByte.Show vbModal, Me

''November 12, 2003
'    Dim intByte As Integer
'    Dim lngPos  As Long
'    Dim lngElement As Long
'    Dim lower   As Long
'    Dim lngLenSequence As Long
'    Dim strInput As String
'    Dim vntInput As Variant
'    Dim bytInput() As Byte
'    Dim strOutput As String
'    Dim vntOutput As Variant
'    Dim bytOutput() As Byte
'    Dim blnFlag   As Boolean
'
'    On Error GoTo BadConversion
'    'get input
'    strInput = InputBox$("Please enter in a byte or sequence of bytes seperated by spaces to find.", "Find Byte")
'    If Len(strInput) = 0 Then Exit Sub
'
'    'validate input
'    vntInput = Split(strInput, " ")
'    lower = LBound(vntInput)
'    lngLenSequence = UBound(vntInput) - LBound(vntInput) + 1
'
'    ReDim bytInput(lower To UBound(vntInput))
'    ReDim bytOutput(lower To UBound(vntInput))
'
'    For lngElement = LBound(vntInput) To UBound(vntInput)
'        If Not IsNumeric(vntInput(lngElement)) Then
'            MsgBox "One of the entries is not a number.", vbInformation, "Find Byte"
'            Exit Sub
'        End If
'        bytInput(lngElement) = CInt(vntInput(lngElement))
'    Next lngElement
'
'    If lngLenSequence < 1 Then
'        MsgBox "Nothing to do.", vbInformation, "Find Byte"
'        Exit Sub
'    End If
'
'
'    If Index Then
'        'get output
'        strOutput = InputBox$("Please enter in " & lngLenSequence & " bytes to replace **without asking**.", "Replace Byte")
'        If Len(strOutput) = 0 Then Exit Sub
'
'        'validate output
'        vntOutput = Split(strOutput, " ")
'        If UBound(vntOutput) <> UBound(vntInput) Or LBound(vntOutput) <> LBound(vntInput) Then
'            MsgBox "Length of the replace byte sequence does not match that of input.", vbInformation, "Replace Byte"
'            Exit Sub
'        End If
'
'        For lngElement = LBound(vntOutput) To UBound(vntOutput)
'            If Not IsNumeric(vntOutput(lngElement)) Then
'                MsgBox "One of the entries is not a number.", vbInformation, "Replace Byte"
'                Exit Sub
'            End If
'            bytOutput(lngElement) = CInt(vntOutput(lngElement))
'        Next lngElement
'    End If
'
'    On Error GoTo SeekReplace
'    For lngPos = mlngLastPos + 1 To mBinFile.FileLength - lngLenSequence
'        'does first byte match?
'        If mBinFile.Item(lngPos) = bytInput(lower) Then
'            blnFlag = True
'            'check each element
'            For lngElement = lower To UBound(bytInput)
'                If mBinFile.Item(lngPos + lngElement - lower) <> bytInput(lngElement) Then
'                    blnFlag = False
'                End If
'            Next lngElement
'
'            If blnFlag Then 'we have a match
'                If Index Then
'                    'overwrite data
'                    For lngElement = lower To UBound(bytInput)
'                        mBinFile.Item(lngPos + lngElement - lower) = bytOutput(lngElement)
'                    Next lngElement
'                    Debug.Print "Overwite occured at " & lngPos
'                    lngPos = lngPos + lngLenSequence - 1
'                    mblnChanged = True
'                Else
'                    'found a match
'                    SetCursor lngPos
'                    Exit Sub
'                End If
'            End If
'        End If
'    Next lngPos
'
'    picBin.Refresh
'
'
'    Exit Sub
'BadConversion:
'    MsgBox "Bad conversion from the numeric entry to an integer.", vbInformation, "Find Byte"
'    Exit Sub
'SeekReplace:
'    MsgBox "Problem while looking for the byte sequence at position " & lngPos & ".  Please report " & Err.Number & " to Richard Fung.", vbInformation, "Find Byte"
'    Exit Sub
End Sub

Private Sub mnuEditFindW_Click(Index As Integer)
    Static lngL As Long  'For binary files
    Dim lngCounter As Long
    Dim lngPos     As Long
    Dim lngStart   As Long 'found start
    Dim lngBeginS  As Long 'begin search point

    On Error GoTo Handler
    If lngL = 0 Then 'find next
        lngL = 10
        lngL = Abs(InputBox("How many ASCII characters should be in a line before you consider it ASCII?", "Find Unicode", lngL))
    ElseIf Index = 0 Then
        lngL = Abs(InputBox("How many ASCII characters should be in a line before you consider it ASCII?", "Find Unicode", lngL))
    End If
    
    MousePointer = vbHourglass
    If lngL <> 0 Then
        lblMessage = "Searching for Unicode text..."
        lblMessage.Refresh
        lngBeginS = 1
        'find next and there is a position, start after that match
        If mlngLastPos > 0 And Index <> 0 Then lngBeginS = mlngLastPos + lngL * 2
        
        For lngPos = lngBeginS To mBinFile.FileLength Step 2
            If IsASCII(mBinFile.Item(lngPos)) And mBinFile.Item(lngPos + 1) = 0 Then
                lngCounter = lngCounter + 1
                If lngCounter >= lngL Then
                    'pattern matched
                    lngStart = lngPos - lngCounter * 2 + 2
                    Exit For
                End If
            Else
                lngCounter = 0
            End If
        Next lngPos
        
        If lngStart > 0 Then
                SetCursor lngStart
            lblMessage = "A Unicode text pattern was found at " & FormatNum2(lngStart) & "."
         Else
            MsgBox "No Unicode text pattern match found.", vbInformation
            lblMessage = "No Unicode text pattern match found."
        End If
    End If
    MousePointer = vbDefault
    Exit Sub
Handler:
    MousePointer = vbDefault
    Exit Sub
End Sub


Private Sub mnuFile_Click()
    If mTextFile Is Nothing Then
        mnuFileSave(0).Enabled = Not (mBinFile Is Nothing)
    Else 'text file exists
        mnuFileSave(0).Enabled = mTextFile.CanSave
    End If
    mnuFileSave(1).Enabled = mnuFileSave(0).Enabled
End Sub

Private Sub mnuFileExit_Click()
    Unload Me
End Sub

Private Sub mnuFileNew_Click()
    Dim lngSize As Long
    Dim lngLoop As Long
    
    If txtBinEdit.Visible Then BinEditOff
    If QuerySave = False Then Exit Sub
    
    If View = conText Then
        Set mTextFile = New TextFile
        mlngFileType = conText
        tabView.Tabs(conText).Caption = "Unicode"
        
        mTextFile.OpenFile ""
        dlgFile.Filename = ""
        
        txt = ""
        txt.Locked = False
        mTextFile.CanSave = True
        Caption = conTitle & " [Untitled]"
        
        mnuViewBuild 0
        UpdateView
        mblnChanged = False 'reset changed flag to 'not changed'
    Else
        lngSize = Val(InputBox("How long do you want the binary file?", "New Binary File", "1"))
        If lngSize <> 0 Then
            Set mBinFile = New BinFile
            mlngFileType = conHex
            mBinFile.OpenFile ""
            dlgFile.Filename = ""
            
            If lngSize > 0 Then
                For lngLoop = 1 To lngSize
                    mBinFile.Item(lngLoop) = 0
                Next lngLoop
            ElseIf lngSize < 0 Then
                mBinFile.Item(-lngSize) = mBinFile.Item(-lngSize)
            End If
            Caption = conTitle & " [Untitled]"
            mnuViewBuild 0
            UpdateView
            mblnChanged = False 'reset changed flag to 'not changed'
        End If
    End If
End Sub

Private Sub mnuFileOpen_Click()
    mnuFileOpenClick2
End Sub


'Saves the file.
Private Sub mnuFileSave_Click(Index As Integer)
    On Error Resume Next
    If Index <> 0 Or Filename = "" Then 'a request to ask for a filename
        'set data
        dlgFile.Filter = "All Files (*.*)|*.*|All Files (*.*)|*.*"
        dlgFile.CancelError = True
        dlgFile.Flags = cdlOFNPathMustExist Or cdlOFNOverwritePrompt Or cdlOFNHideReadOnly
        'do operation
        dlgFile.ShowSave
        If Err = cdlCancel Then
            'user cancelled
            Exit Sub
        ElseIf Err <> 0 Then
            'some unknown error, but do not save
            MsgBox "Error showing save dialog box.  " & Err.Description & " " & Err.Number, vbExclamation
            Exit Sub
        End If
        'change the caption to reflect the filename changes
        Caption = conTitle & " [" & dlgFile.Filename & "]"
    Else
        dlgFile.Filename = Filename 'reinstate this
    End If
    
    'if each of the objects is loaded,
    'then save the file
    If mlngFileType = conText Then
        mTextFile.Text = txt.Text
        mTextFile.Filename = dlgFile.Filename
        mTextFile.SaveFile
    ElseIf mlngFileType <> conNoFile Then
        BinEditOff False 'just record the changed item
        mBinFile.Filename = dlgFile.Filename
        mBinFile.SaveFile
    End If
    If Err = 0 Then
        mblnChanged = False
    Else
        MsgBox "An error was encountered when saving the file.  " & Err.Description & " " & Err.Number, vbExclamation, "Save File"
    End If
End Sub

Private Function UcaseByte(ByVal B As Byte) As Byte
    If B >= 97 And B <= 122 Then 'lowercase ASCII characters
        UcaseByte = B - 32
    Else
        UcaseByte = B
    End If
End Function

Private Sub mnuFindCaseInsensitive_Click(Index As Integer)
    Static strFind    As String
    Dim lngCounter    As Long
    Dim bytSearchA()  As Byte
    Dim bytSearchW()  As Byte
    Dim bytFound      As Byte
    Dim blnFoundA     As Boolean
    Dim blnFoundW     As Boolean
    Dim lngLenW       As Long
    Dim lngPos        As Long
    Dim lngStart      As Long 'found start
    Dim lngBeginS     As Long 'begin search point
    Dim strAppend     As String
    
    On Error GoTo Handler
    If mlngFileType = conHex Then strAppend = "  This searches both ASCII and Unicode text."
    If Len(strFind) = 0 Then 'find next
        strFind = InputBox("Type in the text that you want to find." & strAppend, "Find", strFind)
    ElseIf Index = 0 Then
        strFind = InputBox("Type in the text that you want to find." & strAppend, "Find", strFind)
    End If
    
    MousePointer = vbHourglass
    
    
    If Len(strFind) And mlngFileType = conHex Then
        lblMessage = "Searching for ASCII/Unicode text..."
        lblMessage.Refresh
        
        'always check case insensitive
        strFind = UCase$(strFind)
        
        'assumption: array start at 0
        bytSearchA = StrConv(strFind, vbFromUnicode)
        bytSearchW = strFind
        lngLenW = LenB(strFind)
        
        lngBeginS = 1
        'find next and there is a position, start after that match
        If mlngLastPos > 0 And Index <> 0 Then lngBeginS = mlngLastPos + Len(strFind)
        
        
        For lngPos = lngBeginS To mBinFile.FileLength
            bytFound = mBinFile.Item(lngPos)
            
            If bytFound = bytSearchA(0) Then
                'test ASCII
                blnFoundA = True
                For lngCounter = 1 To (lngLenW \ 2) - 1
                    If UcaseByte(mBinFile.Item(lngPos + lngCounter)) <> bytSearchA(lngCounter) Then
                        blnFoundA = False
                        Exit For
                    End If
                Next lngCounter
                
                If Not blnFoundA Then
                    'test unicode
                    blnFoundW = True
                    For lngCounter = 1 To lngLenW - 1
                        If UcaseByte(mBinFile.Item(lngPos + lngCounter)) <> bytSearchW(lngCounter) Then
                            blnFoundW = False
                            Exit For
                        End If
                    Next lngCounter
                End If
                If blnFoundA Or blnFoundW Then
                    lngStart = lngPos
                    Exit For
                End If
            End If
        Next lngPos
        
        If lngStart > 0 Then
                SetCursor lngStart
            If blnFoundA Then
                lblMessage = "Found ASCII match at " & FormatNum2(lngStart) & " for """ & strFind & """."
            ElseIf blnFoundW Then
                lblMessage = "Found Unicode match at " & FormatNum2(lngStart) & " for """ & strFind & """."
            End If
         Else
            MsgBox "Text not found.", vbInformation
            lblMessage = "Text not found."
        End If
    ElseIf Len(strFind) And mlngFileType = conText Then
        lngBeginS = 1
        If txt.SelStart + 1 > 0 And Index <> 0 Then 'find next, not at start
            lngBeginS = txt.SelStart + Len(strFind) + 1
        End If
        
        lngStart = InStr(lngBeginS, txt.Text, strFind, vbTextCompare)
        If lngStart > 0 Then
            txt.SelStart = lngStart - 1
            txt.SelLength = Len(strFind)
        ElseIf txt.Locked Then 'this txt was not changed
            lngStart = InStr(lngBeginS, mTextFile.Text, strFind, vbTextCompare)
            If lngStart > 0 And lngStart <= Len(txt.Text) Then 'within range of txt
                txt.SelStart = lngStart - 1
                txt.SelLength = Len(strFind)
            Else
                MsgBox "Text not found.", vbInformation, conTitle & "."
                lblMessage = "Text not found."
            End If
        Else
            MsgBox "Text not found.", vbInformation
            lblMessage = "Text not found."
        End If
    End If
    MousePointer = vbDefault
    Exit Sub
Handler:
    MousePointer = vbDefault
    Exit Sub

End Sub

Private Sub mnuHelpAbout_Click()
    frmAbout.Show vbModal, Me
End Sub


Private Sub mnuTools_Click()
    mnuToolsProp.Enabled = (mlngFileType <> conNoFile)
End Sub

'Private Sub mnuToolsInfo_Click()
'    If mlngFileType = conHex Then
'        MsgBox "File Size: " & vbNewLine & FormatNum2(mBinFile.FileLength) & " bytes" & vbNewLine & FormatNum2(mBinFile.FileLength / 1024) & " KB" & vbNewLine & FormatNum2(mBinFile.FileLength / 1024 / 1024) & " MB", vbInformation
'    ElseIf mlngFileType = conText Then
'        If mTextFile.IsUnicode Then
'            MsgBox "Unicode File Size:" & vbNewLine & FormatNum2(Len(txt) * 2) & " bytes" & vbNewLine & FormatNum2(Len(txt) * 2 / 1024) & " KB", vbInformation
'        Else
'            MsgBox "ASCII File Size: " & vbNewLine & FormatNum2(Len(txt)) & " bytes" & vbNewLine & FormatNum2(Len(txt) / 1024) & " KB", vbInformation
'        End If
'    End If
'End Sub

Private Sub mnuToolsOpt_Click()
    frmOptions.Show vbModal, Me
    SetFont 'in case changes were made to gData
    If IsABinaryView Then
        UpdateBinView
    End If
End Sub

Private Sub mnuToolsProp_Click()
    If Not mBinFile Is Nothing Then
        Load frmProp
        With frmProp
            .chkSaved = IIf(mblnChanged, vbUnchecked, vbChecked)
            '.chkNotViewer is Modifiable
            .txtFormat = "Binary"
            .chkNotViewer = vbChecked
            .txtPath = GetPath(mBinFile.Filename)
            .txtFile = Mid$(mBinFile.Filename, Len(.txtPath) + 1)
            .txtSize = FormatNum2(mBinFile.FileLength) & " bytes (" & FormatNum2(mBinFile.FileLength / 1024) & " KB) (" & FormatNum2(mBinFile.FileLength / 1048576) & " MB)"
            .Show vbModal, Me
        End With
    ElseIf Not mTextFile Is Nothing Then
        Load frmProp
        With frmProp
            .chkSaved = IIf(mblnChanged, vbUnchecked, vbChecked)
            .cboFormat.Visible = True
            If mTextFile.IsUnicode Then
                .cboFormat.ListIndex = 0
            Else
                .cboFormat.ListIndex = 1
            End If
            '.txtFormat = tabView.Tabs(conText).Caption
            '.chkNotViewer is Modifiable
            .chkNotViewer = IIf(txt.Locked, vbUnchecked, vbChecked)
            .txtPath = GetPath(mTextFile.Filename)
            .txtFile = Mid$(mTextFile.Filename, Len(.txtPath) + 1)
            .txtSize = FormatNum2(mTextFile.FileLength) & " bytes (" & FormatNum2(mTextFile.FileLength / 1024) & " KB) (" & FormatNum2(mTextFile.FileLength / 1048576) & " MB)"
            .cmdClose.Caption = "Okay"
            .Show vbModal, Me
            If mTextFile.IsUnicode <> (.CboFormatIndex = 0) Then
                mTextFile.IsUnicode = (.CboFormatIndex = 0)
                tabView.Tabs(1).Caption = IIf(mTextFile.IsUnicode, "Unicode", "ASCII")
                MsgBox "Format has changed.", vbInformation
            End If
            
        End With
    End If
End Sub

Private Sub mnuView_Click()
    Dim lngI As Long
    Dim lngB As Long
    
    If Not mTextFile Is Nothing Then
        lngB = mTextFile.Block
        If mnuViewBlock(1).Enabled Then
            For lngI = 1 To mnuViewBlock.UBound
                mnuViewBlock(lngI).Checked = (lngB = lngI)
            Next lngI
        End If
    End If
End Sub

Private Sub mnuViewBlock_Click(Index As Integer)
    mTextFile.Block = Index
    txt.Text = mTextFile.Text
    mblnChanged = False 'not changed
End Sub

Private Sub picBin_DblClick()
    BinEditOn
End Sub

Private Sub picBin_KeyPress(KeyAscii As Integer)
    If View = conCharacter Then
        Select Case KeyAscii
        Case vbKeyBack
            'do nothing
        Case Else
            mblnChanged = True
            
            'change the item
            mBinFile.Item(mlngLastPos) = KeyAscii
            CharDraw mlngLastPos
            
            'move cursor
            mlngLastPos = mlngLastPos + 1
            SelectBinItem mlngLastPos, True
        End Select
    ElseIf KeyAscii >= 48 And KeyAscii <= 57 Then 'number can be accepted in all binary views
        BinEditOn KeyAscii 'turn on editor
    ElseIf View = conHex Then 'hexadecimal view accepts A to F
        Select Case KeyAscii
            Case 65 To 70 'A to F
                BinEditOn KeyAscii
            Case 97 To 102 'a to f
                BinEditOn KeyAscii
        End Select
'    ElseIf View = conCharacter Then 'character view accepts characters above 31
'        If KeyAscii > 31 Then
'            BinEditOn KeyAscii
'        End If
    End If
End Sub


'On the binary viewer...
Private Sub picBin_KeyDown(KeyCode As Integer, Shift As Integer)
    'if the text is not being edited
    If txtBinEdit.Visible = False Then
        Select Case KeyCode 'decide if the selected item should be moved
        Case vbKeyLeft, vbKeyBack  'move left 1 character
            If mlngLastPos - 1 > 0 Then 'check to see if move is within file dimensions
               SelectBinItem mlngLastPos - 1
            End If
        Case vbKeyRight, vbKeySpace 'move right 1 character
            If (View = conCharacter And KeyCode = vbKeyRight) Or View <> conCharacter Then
                If mlngLastPos + 1 <= mBinFile.FileLength + 1 Then 'check to see if move is within file dimensions
                    SelectBinItem mlngLastPos + 1
                End If
            End If
        Case vbKeyUp 'move up by subtracting # of characters in row
            If mlngLastPos - mlngSoftCols > 0 Then 'check to see if move is within file dimensions
                SelectBinItem mlngLastPos - mlngSoftCols
            End If
        Case vbKeyDown 'move down by adding # of characters in row
            If mlngLastPos + mlngSoftCols <= mBinFile.FileLength Then 'check to see if move is within file dimensions
                SelectBinItem mlngLastPos + mlngSoftCols
            End If
        End Select
    End If
End Sub



Private Sub picBin_KeyUp(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case vbKeyPageUp
        If vsbBinView - vsbBinView.LargeChange >= vsbBinView.Min Then
            vsbBinView = vsbBinView - vsbBinView.LargeChange
        End If
    Case vbKeyPageDown
        If vsbBinView + vsbBinView.LargeChange <= vsbBinView.Max Then
            vsbBinView = vsbBinView + vsbBinView.LargeChange
        End If
    End Select
End Sub

Private Sub picBin_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button <> 0 Then 'selects the item at (X, Y)
        SelectBinItem GetChrPos(X, Y)
    End If
End Sub

Private Sub picBin_Paint()
    Dim strDisp    As String 'the display string
    Dim bytNext    As Byte 'the next byte from the file
    Dim tViewChar  As TPRect 'view of character
    Dim tView      As TPRect 'view on the view picture box
    Dim lngCh      As Long 'position in a String
    Dim lngBegin    As Long 'place to begin reading in file
    Dim blnCursor  As Boolean
    Dim lngLeft    As Long
    Dim lngTop     As Long
    Static blnAntiLoop As Boolean 'prevents this from repeating infinitely
        
    If blnAntiLoop Then Exit Sub
    
    blnAntiLoop = True 'turn on drawing flag
    MousePointer = vbArrowHourglass
    
    lngLeft = picBin.ScaleWidth
    lngTop = -15
    
    'The below code has been divided three times
    'to make it display faster.  It displays faster
    'because less decisions in the loop have to be
    'made.  However, this sacrifices memory, but
    'it is not of concern.
    '********************* BinaryFile *********************
    If Not mBinFile Is Nothing And IsABinaryView Then  'text not displayed here
        'move around in the file to begin approx. at the
        'location where the view is visible
        lngBegin = IIf(CurrentRow * mlngSoftCols = 0, CurrentRow * mlngSoftCols + 1, CurrentRow * mlngSoftCols) - 1
        mBinFile.FilePos = lngBegin
        picBin.Cls
        tView = GetBinViewable()
        'add in some leeway by one character left-right, and one line up-down
        tView.Left = tView.Left - msngTextWidth
        tView.Top = tView.Top - msngTextHeight
        tView.Right = tView.Right + msngTextWidth
        tView.Bottom = tView.Bottom + msngTextHeight
        
        Select Case View
        Case conDec '********************* Decimal View
        Do While Not mBinFile.EOF
            strDisp = Space(mlngDataWidth) 'create a String with a pre-defined number of spaces
            'bytNext = mBinFile.Data(conFileNextByte) 'read the next data from the file object
            bytNext = mBinFile.Item(mBinFile.FilePos + 1)
        
            RSet strDisp = bytNext 'put the next byte into the String, right-aligned (RSet); VB automatically translates bytes into numerical data
            tViewChar = GetChrRect(mBinFile.FilePos) 'get the visible area of the item
            If tViewChar.Left >= tView.Left And tViewChar.Right <= tView.Right _
             And tViewChar.Top >= tView.Top And tViewChar.Bottom <= tView.Bottom Then 'make sure the item is in visible range
                If tViewChar.Left < lngLeft Then 'only adjust (X,Y) when needed
                    picBin.CurrentX = tViewChar.Left
                    picBin.CurrentY = tViewChar.Top
                End If
                lngLeft = tViewChar.Left
                lngTop = tViewChar.Top
            
                'decide colours to use
                If mBinFile.FilePos = mBinFile.FileLength Then 'last character is red
                    picBin.ForeColor = vbRed
                    picBin.Print strDisp;
                Else 'other characters
                    If bytNext = 13 Or bytNext = 10 Then 'CR and LF are green
                        picBin.ForeColor = conColourGreen
                    ElseIf bytNext < 32 Then 'not displayable ch. are gray
                        picBin.ForeColor = conColourGray
                    Else 'other numbers are black
                        picBin.ForeColor = vbBlack
                    End If
                    picBin.Print strDisp; 'display character
                End If
            End If
            If tViewChar.Bottom > tView.Bottom Then 'if character beyond bottom, stop drawing -- it won't be visible
                Exit Do
            End If
            If mlngLastPos = mBinFile.FilePos Then SelectBinItem mlngLastPos, True: blnCursor = True             'select the last selected item, if any
        Loop
                
        Case conHex '********************* Hexadecimal view
        'remember this code has been duplicated from above
        'for speed issues
        Do While Not mBinFile.EOF
            strDisp = Space(mlngDataWidth)  'create a String with a pre-defined number of spaces
            'bytNext = mBinFile.Data(conFileNextByte) 'read the next data from the file object
            bytNext = mBinFile.Item(mBinFile.FilePos + 1)
                
            RSet strDisp = Hex(bytNext) 'put in the String the next byte in hexadecimal form, right-aligned using RSet
            
            tViewChar = GetChrRect(mBinFile.FilePos) 'calculate the viewable area for item
            If tViewChar.Left >= tView.Left And tViewChar.Right <= tView.Right _
            And tViewChar.Top >= tView.Top And tViewChar.Bottom <= tView.Bottom Then 'make sure the item is in visible range
                If tViewChar.Left < lngLeft Then 'only adjust (X,Y) when needed
                    picBin.CurrentX = tViewChar.Left
                    picBin.CurrentY = tViewChar.Top
                End If
                lngLeft = tViewChar.Left
                lngTop = tViewChar.Top
                
                'decide colours to use
                If mBinFile.FilePos = mBinFile.FileLength Then 'last character is always going to be red (rouge)
                    picBin.ForeColor = vbRed
                    picBin.Print strDisp;
                Else
                    'go through each character in the display string, which hasn't been displayed yet
                    If bytNext = 13 Or bytNext = 10 Then 'CR and LF are green
                        picBin.ForeColor = conColourGreen
                    ElseIf bytNext < 32 Then 'not displayable characters (now blocks) are gray
                        picBin.ForeColor = conColourGray
                    Else 'display other characters in black
                        picBin.ForeColor = vbBlack
                    End If
                    picBin.Print strDisp;
'                   'this code sets Blue/Orange/Black/Gray/Green for each letter
'                    For lngCh = 1 To Len(strDisp)
'                        If bytNext = 13 Or bytNext = 10 Then 'CR and LF characters are green
'                            '13 and 10 are letters (D and A respectively) in hexadecimal
'                            picBin.ForeColor = conColourGreen
'                        ElseIf bytNext < 32 Then 'undisplayable byte value
'                            Select Case Mid$(strDisp, lngCh, 1)
'                            Case "A" To "F" 'letters of this are orange
'                                picBin.ForeColor = conColourOrange
'                            Case Else 'and the numbers are grey
'                                picBin.ForeColor = conColourGray
'                            End Select
'                        Else 'displayable character equivalent
'                            Select Case Mid$(strDisp, lngCh, 1)
'                            Case "A" To "F" 'letters are blue
'                                picBin.ForeColor = conColourBlue
'                            Case Else 'numbers are black
'                                picBin.ForeColor = vbBlack
'                            End Select
'                        End If
'                        picBin.Print Mid$(strDisp, lngCh, 1); 'display each character one by one
'                    Next lngCh
                End If
            End If
            If tViewChar.Bottom > tView.Bottom Then 'exit when the character has moved beyond visible range
                Exit Do
            End If
            If mlngLastPos = mBinFile.FilePos Then SelectBinItem mlngLastPos, True: blnCursor = True             'select the last selected item, if any
        Loop
            
        Case conCharacter '********************* Character view
        'code has been duplicated from above for speed issues yet again
        'See Also: CharDraw
        Do While Not mBinFile.EOF
            'strDisp = Space(mlngDataWidth) 'create a String with a pre-defined number of spaces
            'bytNext = mBinFile.Data(conFileNextByte) 'read the next data from the file object
            bytNext = mBinFile.Item(mBinFile.FilePos + 1)
                
            If bytNext = 13 Then 'if CR, display paragraph mark
                strDisp = "¶"
            ElseIf bytNext = 10 Then 'if LF (from CR-LF new line combination), put an unobtrusive dot
                strDisp = "·"
            ElseIf bytNext < 32 Then 'not space, but not displayable then show a block
                strDisp = "" 'that is - cannot be displayed properly
            Else 'displayable character
                strDisp = Chr$(bytNext)
            End If
        
            tViewChar = GetChrRect(mBinFile.FilePos) 'calculate visible area
            If tViewChar.Left >= tView.Left And tViewChar.Right <= tView.Right _
            And tViewChar.Top >= tView.Top And tViewChar.Bottom <= tView.Bottom Then 'make sure the character will be seen
                
                'If tViewChar.Left < lngLeft Then 'only adjust (X,Y) when needed
                    picBin.CurrentX = tViewChar.Left
                    picBin.CurrentY = tViewChar.Top
                'End If
                'lngLeft = tViewChar.Left
                'lngTop = tViewChar.Top
                
                If mBinFile.FilePos = mBinFile.FileLength Then 'last character red
                    picBin.ForeColor = vbRed
                    picBin.Print strDisp;
                Else
                    If bytNext = 13 Or bytNext = 10 Then 'CR and LF are green
                        picBin.ForeColor = conColourGreen
                    ElseIf bytNext < 32 Then 'not displayable characters (now blocks) are gray
                        picBin.ForeColor = conColourGray
                    Else 'display other characters in black
                        picBin.ForeColor = vbBlack
                    End If
                    picBin.Print strDisp;
                End If
            End If
            If tViewChar.Bottom > tView.Bottom Then 'get out once it is out of the viewing arena
                Exit Do
            End If
            If mlngLastPos = mBinFile.FilePos Then SelectBinItem mlngLastPos, True: blnCursor = True 'select the last selected item
        Loop
        End Select
        If mlngLastPos = mBinFile.FilePos + 1 Then SelectBinItem mlngLastPos, True: blnCursor = True 'select the last item, but this one is after EOF
        If Not blnCursor Then  'a cursor was not drawn
            mlngLastPos = lngBegin + 1 'move 1 so it the cursor will follow the scroll bar
            SelectBinItem mlngLastPos, True
        End If
    End If

    MousePointer = vbDefault
    blnAntiLoop = False 'turn off drawing flag
End Sub

'turns off the binary edit if editing is happening
Private Sub tabView_BeforeClick(Cancel As Integer)
    If txtBinEdit.Visible Then
        BinEditOff False
        picBinView.Visible = False
        txtBinEdit.Visible = False
    End If
End Sub

'updates the view when a tab is clicked
Private Sub tabView_Click()
    UpdateView
End Sub

Private Sub tabView_KeyUp(KeyCode As Integer, Shift As Integer)
    If (KeyCode = vbKeyReturn Or KeyCode = vbKeyDown) And IsABinaryView Then
        If picBin.Visible Then picBin.SetFocus
    ElseIf (KeyCode = vbKeyReturn Or KeyCode = vbKeyDown) And View = conText Then
        If txt.Visible Then txt.SetFocus
    End If
End Sub

Private Sub tabView_OLEDragDrop(Data As MSComctlLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Data.GetFormat(vbCFFiles) Then
        dlgFile.Filename = Data.Files(1)
        Effect = vbDropEffectCopy
        If Not mnuFileOpenClick2 Then
           dlgFile.Filename = "" 'clear this
        End If
    Else
        Effect = vbDropEffectNone
    End If
End Sub

Private Sub tabView_OLEDragOver(Data As MSComctlLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single, State As Integer)
    If Data.GetFormat(vbCFFiles) Then
        Effect = vbDropEffectCopy
    Else
        Effect = vbDropEffectNone
    End If
End Sub

Private Sub txt_KeyUp(KeyCode As Integer, Shift As Integer)
    Dim bytCh() As Byte
    Dim lngL    As Long
    
    txtPosition
    If KeyCode = vbKeyF1 And txt.Locked = True And txt.SelLength = 1 Then 'cannot be changed from the text file stream
        bytCh = Mid$(mTextFile.Text, txt.SelStart + 1, 1)
        lngL = LBound(bytCh)
        MsgBox "The selected character contains the bytes: " & bytCh(lngL) & " " & bytCh(lngL + 1) & vbNewLine & _
               "In ASCII characters: " & Chr(bytCh(lngL)) & " " & Chr(bytCh(lngL + 1)), vbInformation
    End If
End Sub

Private Sub txtBinEdit_KeyPress(KeyAscii As Integer)
    If KeyAscii > 31 And View = conCharacter Then
        If txtBinEdit.SelStart = mlngDataWUsed Then
            'move right
            BinEditOff
            picBin_KeyDown vbKeyRight, 0
            BinEditOn KeyAscii
            KeyAscii = 0
        End If
    End If
End Sub

Private Sub txtBinEdit_KeyUp(KeyCode As Integer, Shift As Integer)
    'friendly feature
    '- it selects the next item in the text box
    'by pressing left or right keys so the user
    'does not have to use the mouse
    '- all key presses are passed to the binary view box
    If View = conCharacter Then
        Select Case KeyCode
        Case vbKeyEscape
            txtBinEdit.Visible = False
        Case vbKeyRight
            If txtBinEdit.SelStart = Len(txtBinEdit) Then
                BinEditOff
                picBin.SetFocus
                picBin_KeyDown KeyCode, Shift
            End If
            KeyCode = 0
        Case vbKeyLeft
            If txtBinEdit.SelStart = 0 Then
                BinEditOff
                picBin.SetFocus
                picBin_KeyDown KeyCode, Shift
            End If
            KeyCode = 0
        Case vbKeyReturn
            BinEditOff
        End Select
    Else
        Select Case KeyCode
        Case vbKeyEscape
            txtBinEdit.Visible = False
        Case vbKeyRight
            If txtBinEdit.SelStart = Len(txtBinEdit) Then
                BinEditOff
                picBin.SetFocus
                picBin_KeyDown vbKeyRight, Shift
            End If
            KeyCode = 0
        Case vbKeySpace
            BinEditOff
            picBin.SetFocus
            picBin_KeyDown vbKeyRight, Shift
            KeyCode = 0
        Case vbKeyLeft
            If txtBinEdit.SelStart = 0 Then
                BinEditOff
                picBin.SetFocus
                picBin_KeyDown KeyCode, Shift
            End If
            KeyCode = 0
        Case vbKeyReturn
            BinEditOff
        End Select
    End If
End Sub

#If VB6 Then
'on validation (whatever that means)
'turn of the binary edit
Private Sub txtBinEdit_Validate(Cancel As Boolean)
    BinEditOff
End Sub
#End If

'shows the position in the text file view
'and sets up the file for changes
Private Sub txt_Change()
    txtPosition
    mblnChanged = True
End Sub

'shows the position in the text file view
Private Sub txt_Click()
    txtPosition
End Sub


'shows the position in the text file view
Private Sub txt_KeyDown(KeyCode As Integer, Shift As Integer)
    txtPosition
End Sub


'vertical scroll bar stores line numbers
'so only redrawing needs be done
Private Sub vsbBinView_Change()
    picBin_Paint
End Sub


