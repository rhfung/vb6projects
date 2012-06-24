VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   Caption         =   "Flow Chart Viewer"
   ClientHeight    =   5625
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   9000
   Icon            =   "Flowprev_main.frx":0000
   KeyPreview      =   -1  'True
   LockControls    =   -1  'True
   ScaleHeight     =   5625
   ScaleWidth      =   9000
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin MSComDlg.CommonDialog dlgFile 
      Left            =   2760
      Top             =   1680
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      Filter          =   "Flow Chart (*.flc)|*.flc"
      FilterIndex     =   1
   End
   Begin VB.VScrollBar vsbBar 
      Height          =   2895
      LargeChange     =   1500
      Left            =   5280
      SmallChange     =   100
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   720
      Width           =   255
   End
   Begin VB.HScrollBar hsbBar 
      Height          =   255
      LargeChange     =   1500
      Left            =   720
      SmallChange     =   100
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   3480
      Width           =   3615
   End
   Begin VB.PictureBox picView 
      ClipControls    =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2535
      Left            =   960
      ScaleHeight     =   2475
      ScaleWidth      =   3315
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   840
      Width           =   3375
      Begin VB.PictureBox picBay 
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ClipControls    =   0   'False
         FillColor       =   &H80000005&
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   1455
         Left            =   0
         MouseIcon       =   "Flowprev_main.frx":0442
         ScaleHeight     =   1455
         ScaleWidth      =   1335
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   0
         Width           =   1335
      End
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileNew 
         Caption         =   "&New"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuFileOpen 
         Caption         =   "&Open..."
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuFileSave 
         Caption         =   "&Save"
         Enabled         =   0   'False
         Index           =   0
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuFileSave 
         Caption         =   "Save &As..."
         Enabled         =   0   'False
         Index           =   1
      End
      Begin VB.Menu mnuFileSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFilePrintSetup 
         Caption         =   "Prin&t Setup..."
      End
      Begin VB.Menu mnuFilePrint 
         Caption         =   "&Print"
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuFileSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
      Begin VB.Menu mnuEditCut 
         Caption         =   "Cu&t"
         Enabled         =   0   'False
         Shortcut        =   ^X
      End
      Begin VB.Menu mnuEditCopy 
         Caption         =   "&Copy"
         Enabled         =   0   'False
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuEditPaste 
         Caption         =   "&Paste"
         Enabled         =   0   'False
         Shortcut        =   ^V
      End
      Begin VB.Menu mnuEditSep3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditEdit 
         Caption         =   "&Edit Text"
         Enabled         =   0   'False
         Shortcut        =   {F2}
      End
      Begin VB.Menu mnuEditCancel 
         Caption         =   "Cancel Edit"
         Enabled         =   0   'False
         Shortcut        =   {F3}
      End
      Begin VB.Menu mnuEditDel 
         Caption         =   "&Delete"
         Enabled         =   0   'False
         Shortcut        =   {DEL}
      End
      Begin VB.Menu mnuEditSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditToFront 
         Caption         =   "&Bring to front"
         Enabled         =   0   'False
         Shortcut        =   ^J
      End
      Begin VB.Menu mnuEditToBack 
         Caption         =   "&Send to back"
         Enabled         =   0   'False
         Shortcut        =   ^K
      End
   End
   Begin VB.Menu mnuText 
      Caption         =   "&Text"
      Begin VB.Menu mnuTextLeft 
         Caption         =   "&Left"
         Enabled         =   0   'False
         Shortcut        =   ^L
      End
      Begin VB.Menu mnuTextCentre 
         Caption         =   "&Centre"
         Enabled         =   0   'False
         Shortcut        =   ^E
      End
      Begin VB.Menu mnuTextRight 
         Caption         =   "&Right"
         Enabled         =   0   'False
         Shortcut        =   ^R
      End
      Begin VB.Menu mnuTextSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuTextBold 
         Caption         =   "&Bold"
         Enabled         =   0   'False
         Shortcut        =   ^B
      End
      Begin VB.Menu mnuTextItalic 
         Caption         =   "&Italic"
         Enabled         =   0   'False
         Shortcut        =   ^I
      End
      Begin VB.Menu mnuTextSize 
         Caption         =   "Si&ze"
         Begin VB.Menu mnuTextSizeExL 
            Caption         =   "Extra large"
            Enabled         =   0   'False
         End
         Begin VB.Menu mnuTextSizeL 
            Caption         =   "Large"
            Enabled         =   0   'False
         End
         Begin VB.Menu mnuTextSizeNormal 
            Caption         =   "Normal"
            Enabled         =   0   'False
         End
         Begin VB.Menu mnuTextSizeS 
            Caption         =   "Small"
            Enabled         =   0   'False
         End
         Begin VB.Menu mnuTextSizeExS 
            Caption         =   "Extra small"
            Enabled         =   0   'False
         End
      End
      Begin VB.Menu mnuTextSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuTextFont 
         Caption         =   "Font..."
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuTextSep3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuTextSpell 
         Caption         =   "&Spell Check..."
         Enabled         =   0   'False
      End
   End
   Begin VB.Menu mnuFormat 
      Caption         =   "For&mat"
      Begin VB.Menu mnuFormatFont 
         Caption         =   "&Font..."
         Shortcut        =   ^F
      End
      Begin VB.Menu mnuFormatSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFormatAlign 
         Caption         =   "&Align"
         Begin VB.Menu mnuFormatAlignOne 
            Caption         =   "To &Grid"
            Enabled         =   0   'False
            Shortcut        =   ^A
         End
         Begin VB.Menu mnuFormatAlignAll 
            Caption         =   "&All to Grid"
            Enabled         =   0   'False
            Shortcut        =   {F8}
         End
      End
      Begin VB.Menu mnuFormatGrid 
         Caption         =   "&Grid"
         Begin VB.Menu mnuFormatGridSnap 
            Caption         =   "&Snap to Grid"
            Enabled         =   0   'False
            Shortcut        =   ^G
         End
         Begin VB.Menu mnuFormatGridShow 
            Caption         =   "Sho&w Grid"
            Enabled         =   0   'False
            Shortcut        =   ^H
         End
      End
      Begin VB.Menu mnuFormatSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFormatVer 
         Caption         =   "&Version"
         Begin VB.Menu mnuFormatVerOld 
            Caption         =   "1.0"
            Enabled         =   0   'False
         End
      End
      Begin VB.Menu mnuFormatSep3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFormatZoom 
         Caption         =   "&Zoom..."
      End
   End
   Begin VB.Menu mnuTools 
      Caption         =   "&Tools"
      Begin VB.Menu mnuToolsAutoFitText 
         Caption         =   "&Auto-fit text"
         Enabled         =   0   'False
         Shortcut        =   {F6}
      End
      Begin VB.Menu mnuToolsDuplicate 
         Caption         =   "&Duplicate"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuToolsSplit 
         Caption         =   "S&plit Line"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuToolsSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuToolsExport 
         Caption         =   "&Export doc as picture..."
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuToolsSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuToolsReset 
         Caption         =   "&Reset Picture"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuToolsPicZoom 
         Caption         =   "&Picture Zoom..."
         Enabled         =   0   'False
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelpPrintInfo 
         Caption         =   "&Printer Statistics"
      End
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "&About Flow Chart Viewer"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Richard Fung.  August 12, 2000.
'Updated Feb. 2001

'Objects
Public mFlowChart As FlowChart
Private mRegistry  As PRegistry
'for interfaces
Private mMousePos  As Rect
Private mblnDown   As Boolean
'Scale values
Private msngScale   As Single
Private msngGrid    As Single


Private Const conFontName = "Times New Roman"
Private Const conFontSize = 12
'SetScrollBars()
Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Private Const SM_CXHSCROLL = 21
Private Const SM_CXVSCROLL = 2
Private Const SM_CYHSCROLL = 3
Private Const SM_CYVSCROLL = 20

Private Function CalcView() As Rect
    CalcView.x1 = -picBay.ScaleX(picBay.Left, picView.ScaleMode, picBay.ScaleMode)
    CalcView.y1 = -picBay.ScaleY(picBay.Top, picView.ScaleMode, picBay.ScaleMode)
    CalcView.x2 = CalcView.x1 + picBay.ScaleX(picView.ScaleWidth, picView.ScaleMode, picBay.ScaleMode)
    CalcView.y2 = CalcView.y1 + picBay.ScaleY(picView.ScaleHeight, picView.ScaleMode, picBay.ScaleMode)
End Function

Private Property Get conShiftValue() As Single
    conShiftValue = msngGrid * 2
End Property

Private Sub DefaultVal() 'for new, open
    'handled by FlowChart class
'    mcurFontSize = conFontSize
'    mstrFontName = conFontName
'    mlngPercentage = 100
    'mlngDefaultTextFlags = 0
    'mlngAddType = 0
    hsbBar.Value = 0
    vsbBar.Value = 0
End Sub

Public Sub OpenFile(FileName As String)
    Const conDiffers = "differs from the last time you edited."
    
    On Error Resume Next 'at top
    Set mFlowChart = New FlowChart
    Reset
    picBay.Visible = False
    DefaultVal
    If mFlowChart.Load(FileName) = conFail Then
        MsgBox "Problem opening the flow chart.  The flow chart may be corrupted or the file is too new for this version of the program.", vbExclamation
        mFlowChart.FileName = "" 'erase filename
    End If
    
    'parts of Reset()
    Recaption
    'mblnChanged = False
    
    frmLog.txtLog = "" 'log beginning
    
    'check version information
    If mFlowChart.Version > conCurrentVersion Then
        MsgBox "The file version is newer than the compiled version.  Some incompatibilities may arise.", vbInformation
        frmLog.AddLine "The file version is newer than the compiled version.  Some incompatibilities may arise."
    End If
    
    'set other data
    If mFlowChart.Version >= 4 Then
        With mFlowChart
            If .PrinterError Then  'for v4+
                frmLog.AddLine "PRINTER ERRORS! Many problems will arise with the program.  To correct, you must restart the program."
            Else
                If .Header1PDevName <> Printer.DeviceName And Len(.Header1PDevName) Then
                    frmLog.AddLine "The file was edited with the printer " & .Header1PDevName & ".  Now you are using " & Printer.DeviceName & "."
                End If
                If .Orientation <> Printer.Orientation Then
                    frmLog.AddLine "The orientation of page " & conDiffers
                    Printer.Orientation = .Orientation
                End If
                If .PaperSize <> Printer.PaperSize Then
                    frmLog.AddLine "The paper size " & conDiffers
                End If
                If .PScaleHeight > Printer.ScaleHeight Or _
                  .PScaleWidth > Printer.ScaleWidth Then
                    frmLog.AddLine "The drawing area " & conDiffers
                    frmLog.AddLine "Some of the drawing area may be clipped off."
                ElseIf .PScaleHeight <> Printer.ScaleHeight Or _
                  .PScaleWidth <> Printer.ScaleWidth Then
                    frmLog.AddLine "The drawing area " & conDiffers
                End If
            End If
            On Error Resume Next
            'check scroll X and Y values
            If .ScrollX < conScrollParts And .ScrollY < conScrollParts Then
                hsbBar.Value = .ScrollX / conScrollParts * hsbBar.Max
                vsbBar.Value = .ScrollY / conScrollParts * vsbBar.Max
            End If
        End With
    End If
    
    If mFlowChart.Version >= 2 Then 'V2
        SetView mFlowChart.ZoomPercent, mFlowChart.FontName, mFlowChart.FontSize
    Else 'V1
        SetView FontName:=conFontName, FontSize:=conFontSize
    End If
    
    OpenBoundaryCheck

    If Len(frmLog.txtLog) = 0 Then
        Unload frmLog
    End If
    picBay.Visible = True
End Sub

Public Sub MainShow()
    Form_Resize
    Show
    Refresh
End Sub


Private Sub picBayPaint(Optional All As Boolean)
    Dim objItem As FlowItem
    Dim sngMrg As Single
    Dim tView As Rect
    
    tView = CalcView()
        
    On Error GoTo Handler
'    If mnuToolsGridShow.Checked Then
'        Set objItem = New FlowItem
'        With objItem
'            .Left = -picBay.ScaleX(picBay.Left, picView.ScaleMode, picBay.ScaleMode)
'            .Top = -picBay.ScaleY(picBay.Top, picView.ScaleMode, picBay.ScaleMode)
'        End With
'        Align objItem, True
'        picBay.PaintPicture picGrid.Image, objItem.Left, objItem.Top
'        Set objItem = Nothing
'    End If
    
    'margins
    sngMrg = picBay.ScaleX(1, vbInches, picBay.ScaleMode) * mFlowChart.ZoomPercent / 100
    picBay.DrawStyle = vbDot 'to signal a viewer
    picBay.Line (sngMrg, sngMrg)-Step(picBay.ScaleWidth - sngMrg * 2, picBay.ScaleHeight - sngMrg * 2), vb3DFace, B
    picBay.DrawStyle = vbSolid
    
    'bottom
    For Each objItem In mFlowChart
        If objItem.DrawOrder = conBottom Then
            GoSub DrawItem
        End If
    Next objItem

    'middle
    For Each objItem In mFlowChart
        If objItem.DrawOrder = conMiddle Then
            GoSub DrawItem
        End If
    Next objItem

    'top
    For Each objItem In mFlowChart
        If objItem.DrawOrder = conTop Then
            GoSub DrawItem
        End If
    Next objItem
    
    SetDrawProps Nothing, picBay
    Exit Sub
DrawItem:
    If (objItem.Left >= tView.x1 And _
       objItem.Left + objItem.Width <= tView.x2) And _
       (objItem.Top >= tView.y1 And _
       objItem.Top + objItem.Height <= tView.y2) Or All Or _
      (objItem.Left + objItem.Width >= tView.x1 And _
       objItem.Left <= tView.x2) And _
       (objItem.Top + objItem.Height >= tView.y1 And _
       objItem.Top <= tView.y2) Then
            objItem.Draw picBay, Nothing
'            Debug.Print "{NEW ITEM} "; objItem.Text
    End If
    Return
    Exit Sub
Handler:
    Resume Next
End Sub


Private Sub Recaption()
    Dim strCaption As String
    Dim intTrim As Long
    
    If Len(mFlowChart.FileName) Then
        strCaption = String(255, vbNull)
        intTrim = GetFileTitle(mFlowChart.FileName & vbNullChar, strCaption, 255)
        strCaption = Left(strCaption, InStr(1, strCaption, vbNullChar, vbBinaryCompare) - 1)
        Caption = strCaption & " - Flow Chart Viewer"
    Else
        Caption = "Flow Chart Viewer"
    End If
End Sub



Private Sub SetScrollBars()
    Dim sngWidth As Single 'of scroll bar
    Dim sngHeight As Single
    'Const conScrollDif = 0
    
    'takes the image of the arrow
    sngWidth = GetSystemMetrics(SM_CXVSCROLL) '+ conScrollDif
    sngHeight = GetSystemMetrics(SM_CYHSCROLL) '+ conScrollDif
    sngWidth = ScaleX(sngWidth, vbPixels, ScaleMode)
    sngHeight = ScaleY(sngHeight, vbPixels, ScaleMode)
    vsbBar.Width = sngWidth
    hsbBar.Height = sngHeight
End Sub

Private Sub SetView(Optional ByVal Percentage As Long, Optional ByVal FontName As String, Optional ByVal FontSize As Currency)
'Necessary for both Viewer and main program.
'Sets up the necessary stuff for drawing to picBay
'and to the Printer.
    Dim blnGridChanged As Boolean
    Dim hVal As Single, vVal As Single
        
    'adjust values
    If Percentage = 0 Then Percentage = mFlowChart.ZoomPercent
    If Len(FontName) = 0 Then FontName = mFlowChart.FontName
    If FontSize = 0 Then FontSize = mFlowChart.FontSize
    
    blnGridChanged = (Percentage <> mFlowChart.ZoomPercent)
    
    'Adjust scale
    msngScale = conScale * (Percentage / 100)
    'msngGrid2 = 150 / msngScale
    
    On Error Resume Next
    'Match font
    picBay.Font.Name = FontName
    picBay.Font.Size = FontSize
    picBay.Font.Bold = False
    picBay.Font.Italic = False
    
    Printer.Font.Name = FontName
    Printer.Font.Size = FontSize
    Printer.Font.Bold = False
    Printer.Font.Italic = False
    
    If Err Then frmLog.AddLine "Screen font and printer font cannot be matched."
    
    'msngPercentAdjust = 1 'ScaleX(Screen.Width, vbTwips, vbPixels) / Screen.Width
    
    'Size and scale
    If Err Or mFlowChart.PrinterError Then
        Err.Clear
        picBay.Move 0, 0, mFlowChart.PScaleWidth * msngScale, mFlowChart.PScaleHeight * msngScale
        picBay.Scale (0, 0)-(mFlowChart.PScaleWidth, mFlowChart.PScaleHeight)
    Else
        picBay.Move 0, 0, Printer.ScaleWidth * msngScale, Printer.ScaleHeight * msngScale
        picBay.Scale (0, 0)-(Printer.ScaleWidth, Printer.ScaleHeight)
    End If
    
    'scale font size
    Err.Clear
    'picBay.Font.Size = Printer.TextWidth("Today") / picBay.ScaleX(picBay.TextWidth("Today"), picBay.ScaleMode, vbTwips) * FontSize * (Percentage / 100)
    picBay.Font.Size = Printer.Font.Size * (Percentage / 100)
    
    If Err Then
        picBay.Font.Size = FontSize * (Percentage / 100)
        frmLog.AddLine "The screen font cannot be scaled to the printer output."
        Err.Clear
    End If
    
    'tell user about any discrepancies
    If FontSize <> picBay.Font.Size Then
        'red colour to tell of font discrepancies
        Line (hsbBar.Left + hsbBar.Width, vsbBar.Top + vsbBar.Height)-(ScaleWidth, ScaleHeight), vbRed, BF
    Else
        Form_Paint 'show icon to tell its okay
    End If
    
    picBay.DrawWidth = IIf(Percentage >= 100, (Percentage / 100), 1)
    'note changed values
    mFlowChart.ZoomPercent = Percentage
    mFlowChart.FontName = FontName
    mFlowChart.FontSize = FontSize
    
    'remember the percentage position of scroll bar
    On Error Resume Next
    vVal = vsbBar.Value / vsbBar.Max
    hVal = hsbBar.Value / hsbBar.Max
    'change view position to prevent errors
    'when the scroll bar max values are changed
    vsbBar.Value = 0
    hsbBar.Value = 0
    
    'update scroll bars by calling ::Resize
    UpdateScrollBars
    
    'set the percentage position of scroll bars
    On Error Resume Next
    vsbBar.Value = vVal * vsbBar.Max
    hsbBar.Value = hVal * hsbBar.Max
    
    Redraw
End Sub

'Called from FileOpen.
Private Sub OpenBoundaryCheck()
    Dim objItem As FlowItem
    Dim strText As String
    Dim lngOutCount As Long 'count of off-boundary items
    
    For Each objItem In mFlowChart
        If (objItem.Left + objItem.Width > picBay.ScaleWidth) Or _
           (objItem.Top + objItem.Height > picBay.ScaleHeight) Then
            strText = GetText(objItem)
            frmLog.AddLine "The item with the text """ & IIf(Len(strText) > 20, Left(strText, 20) & "...", strText) & """ is off the screen."
            lngOutCount = lngOutCount + 1
        End If
    Next objItem
    
    If lngOutCount > 0 Then
        frmLog.AddLine "There are " & lngOutCount & " item(s) that are off the view of the page.  These items may not print properly."
        frmLog.AddLine "Close the program, change to the original page size, and re-open the program to fix this problem."
        MsgBox "There are " & lngOutCount & " item(s) that are off the view of the page.  These items may not print properly.  Close the program, change to the original page size, and re-open the program to fix this problem.", vbExclamation
    End If
End Sub


'Redraws everything.
Private Sub Redraw()
    Screen.MousePointer = vbArrowHourglass
    picBay.Cls
    picBayPaint
'    RedrawHandles mSelected
    Screen.MousePointer = vbDefault
End Sub

Private Sub UpdateScrollBars()
    If picBay.Width < picView.ScaleWidth Then
        hsbBar.Max = 0
        hsbBar.Enabled = False
    Else
        hsbBar.Max = ScaleX(picBay.Width - picView.ScaleWidth, vbTwips, conScrollScale)
        hsbBar.LargeChange = ScaleX(picView.ScaleWidth, vbTwips, conScrollScale)
        hsbBar.Enabled = True
    End If
    If picBay.Height < picView.ScaleHeight Then
        vsbBar.Max = 0
        vsbBar.Enabled = False
    Else
        vsbBar.Max = ScaleY(picBay.Height - picView.ScaleHeight, vbTwips, conScrollScale)
        vsbBar.LargeChange = ScaleY(picView.ScaleHeight, vbTwips, conScrollScale)
        vsbBar.Enabled = True
    End If
End Sub

Private Sub Form_Initialize()
    Set mFlowChart = New FlowChart
    Set mRegistry = New PRegistry
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'    If Not IsBusy Then
        Select Case KeyCode
            Case vbKeyUp
                vsbBar.Value = IIf(vsbBar.Value > vsbBar.SmallChange, vsbBar.Value - vsbBar.SmallChange, 0)
            Case vbKeyDown
                vsbBar.Value = IIf(vsbBar.Value < (vsbBar.Max - vsbBar.SmallChange), vsbBar.Value + vsbBar.SmallChange, vsbBar.Max)
            Case vbKeyLeft
                hsbBar.Value = IIf(hsbBar.Value > hsbBar.SmallChange, hsbBar.Value - hsbBar.SmallChange, 0)
            Case vbKeyRight
                hsbBar.Value = IIf(hsbBar.Value < (hsbBar.Max - hsbBar.SmallChange), hsbBar.Value + hsbBar.SmallChange, hsbBar.Max)
            Case vbKeyPageDown
                vsbBar.Value = IIf(vsbBar.Value < (vsbBar.Max - vsbBar.LargeChange), vsbBar.Value + vsbBar.LargeChange, vsbBar.Max)
            Case vbKeyPageUp
                vsbBar.Value = IIf(vsbBar.Value > vsbBar.LargeChange, vsbBar.Value - vsbBar.LargeChange, 0)
'            Case 192 'key with "`~" above tab key
'                SelectNextItem
            'more cases on KeyUp
        End Select
End Sub


Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
        Select Case KeyCode
            Case vbKeyHome
                hsbBar.Value = 0
            Case vbKeyEnd
                hsbBar.Value = hsbBar.Max
            Case vbKeyX
                mnuHelpAbout_Click
            Case vbKeyF5 'force redraw
                picBay.Cls
                picBayPaint True
        End Select
'    End If
End Sub

Private Sub Form_Load()
    Dim i As Integer
    
    On Error Resume Next
    'Match the scales
    Printer.ScaleMode = vbTwips
    If Err Then
        MsgBox "Printer error.  Cannot set proper scale mode.", vbExclamation
    End If
    
'    'Load handle boxes
'    For i = conHMin To conHMax
'        Load ctlBox(i)
'    Next i
'    'Fonts
'    SetEditFont
    
    'Registry settings
    If mRegistry.ReadRegistry <> 0 Then
        MsgBox "Error reading the registry.", vbExclamation
    End If
    
    'Check for any printer errors again
    If mFlowChart.PrinterError Then
        frmLog.AddLine "PRINTER ERRORS!  A printer must be installed on your computer."
        frmLog.AddLine "                 Problems may also arise if you are using a network printer."
    End If
    
    'View settings - must go at front
    SetView 100, conFontName, conFontSize
    SetScrollBars
    Form_Resize
    
    mnuFormatGridSnap.Checked = mRegistry.GridSnap
    mnuFormatGridShow.Checked = mRegistry.GridShow ': RedrawGrid 'refresh changes
End Sub


Private Sub Form_Paint()
    Line (hsbBar.Left + hsbBar.Width, vsbBar.Top + vsbBar.Height)-Step(16, 16), BackColor, BF
    CurrentX = hsbBar.Left + hsbBar.Width
    CurrentY = vsbBar.Top + vsbBar.Height
    Print "Vw"
    'PaintPicture Icon, hsbBar.Left + hsbBar.Width, vsbBar.Top + vsbBar.Height, 16, 16
End Sub


Private Sub Form_Resize()
    On Error GoTo Handler
    If WindowState <> vbMinimized Then
        picView.Move 0, 0, ScaleWidth - vsbBar.Width, ScaleHeight - hsbBar.Height
        vsbBar.Move picView.Width, picView.Top, vsbBar.Width, picView.Height
        hsbBar.Move picView.Left, ScaleHeight - hsbBar.Height, picView.Width
        UpdateScrollBars
    End If
Handler:
End Sub

Private Sub Form_Terminate()
    Set mFlowChart = Nothing
End Sub

Private Sub hsbBar_Change()
    If Not mblnDown Then picBay.Left = ScaleX(-hsbBar.Value, conScrollScale, vbTwips)
End Sub


Private Sub mnuFileExit_Click()
    Unload Me
End Sub

Private Sub mnuFileOpen_Click()
    On Error Resume Next
'    If QuerySave Then 'saved
        With dlgFile
            .Flags = cdlOFNFileMustExist Or cdlOFNHideReadOnly
            .ShowOpen
            If Err = cdlCancel Then 'cancelled
                Exit Sub
            ElseIf Err <> 0 Then 'error
                MsgBox "Unable to show Open dialog box.", vbExclamation
            Else 'good
                OpenFile .FileName
            End If
        End With
'    End If
End Sub

Private Sub mnuFilePrint_Click()
    Dim objItem As FlowItem
    
    On Error GoTo Handler
    
    If MsgBox("Do you want to print this flow chart?", vbQuestion Or vbYesNo) = vbNo Then Exit Sub
    
    MousePointer = vbHourglass
    mFlowChart.PrintFile Printer, picBay.Font
    
    MousePointer = vbDefault
    Exit Sub
Handler:
    On Error Resume Next
    MsgBox "There was a problem while printing.  The print job could not be completed.", vbExclamation, "Print"
    Printer.KillDoc 'clear the document from Windows spooling
    MousePointer = vbDefault
End Sub


Private Sub mnuFilePrintSetup_Click()
    On Error Resume Next
    Printer.KillDoc 'release printer object handle
    dlgFile.Flags = cdlPDPrintSetup
    dlgFile.Orientation = mFlowChart.Orientation
    dlgFile.Min = 1 'reset these values
    dlgFile.Max = 1
    dlgFile.ShowPrinter
    If Err = 0 Then
        Printer.Copies = dlgFile.Copies
        mFlowChart.Orientation = dlgFile.Orientation
        Printer.Orientation = mFlowChart.Orientation
        mFlowChart.PrintFile Printer, picBay.Font, True
        
        SetView 'resize the view area
    ElseIf Err = cdlCancel Then
        'ignore
    Else
        MsgBox "Problem showing the printer setup dialog box.", vbExclamation, "Print Setup"
    End If
    
End Sub

Private Sub mnuFormat_Click()
    mnuFormatVerOld.Caption = Format(mFlowChart.Version, "#0.0#")
    mnuFormatVerOld.Checked = True
    mnuFormatFont.Enabled = (mFlowChart.Version < 6) 'old version
End Sub


Private Sub mnuFormatFont_Click()
    On Error Resume Next
    With dlgFile
        .Flags = cdlCFBoth Or cdlCFLimitSize Or cdlCFWYSIWYG Or cdlCFScalableOnly
        .Min = 4
        .Max = 96
        .FontName = mFlowChart.FontName
        .FontSize = mFlowChart.FontSize
        MsgBox "Font changes cannot be saved.", vbInformation, "Change Font"
        .ShowFont
        If Err = cdlCancel Then 'cancelled
            Exit Sub
        ElseIf Err <> 0 Then
            MsgBox "Error showing Font dialog box.", vbExclamation, "Font"
        Else
            SetView FontName:=.FontName, FontSize:=.FontSize
        End If
    End With
End Sub


Private Sub mnuFormatZoom_Click()
    Load frmZoom
    frmZoom.Zoom = mFlowChart.ZoomPercent
    frmZoom.Show 1, Me
    If frmZoom.Changed Then
'        RedrawHandles Nothing
        SetView frmZoom.Zoom
        'version changed
'        UpgradeVersion 2
'        mblnChanged = True
    End If
    Unload frmZoom
End Sub


Private Sub mnuHelpAbout_Click()
    frmAbout.Show 1, Me
End Sub

Private Sub mnuHelpPrintInfo_Click()
    On Error GoTo Handler
    frmLog.txtLog = ""
    frmLog.AddLine "This is a viewer.  The full program includes full drawing features."
    frmLog.AddLine ""
    frmLog.AddLine "--- PRINTER INFO ---"
    With Printer
        frmLog.AddLine "# of Copies: " & .Copies
        frmLog.AddLine "Device Name: " & .DeviceName
        frmLog.AddLine "Driver Name: " & .DriverName
        frmLog.AddLine "Font Name: " & .Font.Name
        frmLog.AddLine "Font Size: " & .Font.Size
        frmLog.AddLine "Orientation: " & Choose(.Orientation, "Portrait", "Landscape") & " (" & .Orientation & ")"
        frmLog.AddLine "Paper Size: " & .PaperSize
        frmLog.AddLine "Port: " & .Port
        frmLog.AddLine "Print Quality: " & IIf(.PrintQuality < 0, Choose(-.PrintQuality, "Draft", "Low", "Medium", "High"), .PrintQuality & " dpi")
        frmLog.AddLine "Zoom: " & IIf(.Zoom = 0, "100%", .Zoom)
        frmLog.AddLine ""
    End With
    Exit Sub
Handler:
    frmLog.AddLine "Error accessing printer object.  " & Err.Description & " (" & Err.Number & ")."
End Sub

Private Sub picBay_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbLeftButton Then
        mMousePos.x1 = picBay.Left
        mMousePos.y1 = picBay.Top
        mMousePos.x2 = X
        mMousePos.y2 = Y
        picBay.MousePointer = vbCustom
        mblnDown = True
    End If
End Sub

Private Sub picBay_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If mblnDown Then
        mMousePos.x1 = picBay.Left + X - mMousePos.x2
        mMousePos.y1 = picBay.Top + Y - mMousePos.y2
        If Not (-mMousePos.x1 > 0) Then
            mMousePos.x1 = 0 'picBay.Left
        ElseIf Not (mMousePos.x1 + picBay.Width >= picView.Width) Then
            If picBay.Width > picView.Width Then
                mMousePos.x1 = picView.Width - picBay.Width 'picBay.Left
            Else
                mMousePos.x1 = 0
            End If
        End If
        If Not (-mMousePos.y1 > 0) Then
            mMousePos.y1 = 0 'picBay.Top
        ElseIf Not (mMousePos.y1 + picBay.Height >= picView.Height) Then
            If picBay.Height > picView.Height Then
                mMousePos.y1 = picView.Height - picBay.Height
            Else
                mMousePos.y1 = 0
            End If
        End If
        picBay.Move mMousePos.x1, mMousePos.y1
    End If
End Sub


Private Sub picBay_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If mblnDown Then
        picBay_MouseMove Button, Shift, X, Y
        On Error Resume Next
        hsbBar.Value = ScaleX(-mMousePos.x1, vbTwips, vbPixels)
        vsbBar.Value = ScaleY(-mMousePos.y1, vbTwips, vbPixels)
        picBay.MousePointer = vbDefault
        mblnDown = False
    End If
End Sub

Private Sub picBay_Paint()
    picBayPaint
End Sub

Private Sub vsbBar_Change()
    If Not mblnDown Then picBay.Top = ScaleY(-vsbBar.Value, conScrollScale, vbTwips)
End Sub


