Attribute VB_Name = "modBase"
Option Explicit
'Richard Fung.  Aug 12, 2000
'frmI... means that frm... is defined in code

'Constants
Public Const conCurrentVersion = 8!
Public Const conFail = -1
Public Const conIndent = 30 'twips of 2 pixels
Public Const conPi = 3.14159256358979
Public Const conR = conPi / 180
Public Const conSizeFactorMultiplier = 2@
Public Const conFontName = "Times New Roman"
Public Const conFontSize = 12
Public Const conScrollParts = 10000
Public Const conScrollScale = vbPoints
Public Const conFontMin = 4
Public Const conFontMax = 96
Public Const conDefaultFilter = "Flow Chart (*.flc)|*.flc"
Public Const conDefaultExt = "flc"

Public Const DT_CENTER = &H1
Public Const DT_LEFT = &H0
Public Const DT_RIGHT = &H2
Public Const DT_WORDBREAK = &H10
Public Const DT_CALCRECT = &H400
'Public Const DT_RASPRINTER = 2
'Public Const DT_RASDISPLAY = 1
Public Const DT_NOPREFIX = &H800
'Public Const DT_NOCLIP = &H100
Public Const TRANSPARENT = 1
Public Const OPAQUE = 2

'Type Structures

Public Type Rect 'For Flowchart
    X1 As Long 'Single
    Y1 As Long 'Single
    X2 As Long 'Single
    Y2 As Long 'Single
End Type

Public Type apiRECT 'For WinAPI
    Left    As Long
    Top     As Long
    Right   As Long
    Bottom  As Long
End Type

Public Type POINTAPI 'For WinAPI
    X As Long
    Y As Long
End Type

'Enumerations

Public Enum FAddType 'this order should match the toolbar order on the main form
    conAddRect = 1
    conAddInOut
    conAddDecision
    conAddTerminator
    conAddCircle
    conAddLine
    conAddMidArrowLine
    conAddEndArrowLine
    conAddText
    conAddPicture
    conAddEllipse 'not added
    conAddShapeA  'not used below
    conAddShapeB
    conAddShapeC
    conAddShapeD
    conAddShapeE
    conAddShapeF
    conAddButton    'used
    conAddLayerDefaults = 9999
    conAddExtra1 = 10000 'area
    conAddExtra2 = 10010
    conAddExtra3 = 10020
    conAddSecurity = 10030
End Enum

'Import DLL Functions
Public Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As apiRECT) As Long
Public Declare Function GetDesktopWindow Lib "user32" () As Long
'0-based array for Polygon, used by many shapes
Public Declare Function Polygon Lib "gdi32" (ByVal hdc As Long, lpPoint As POINTAPI, ByVal nCount As Long) As Long
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Public Declare Function apiDrawText Lib "user32" Alias "DrawTextA" (ByVal hdc As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As apiRECT, ByVal wFormat As Long) As Long
Public Declare Function GetFileTitle Lib "comdlg32.dll" Alias "GetFileTitleA" (ByVal lpszFile As String, ByVal lpszTitle As String, ByVal cbBuf As Integer) As Integer
Public Declare Function SetBkMode Lib "gdi32" (ByVal hdc As Long, ByVal nBkMode As Long) As Long
'SetBkMode is required because BUG: Setting FontTransparent
'Has No Effect on Windows 95/98/Me for Visual Basic (Q145726)
'- FontTransparent does not work

'Variable declarations

#If Prev = 0 Then
'Public gblnWindowColour As Boolean
'Public gblnWindowFont   As Boolean
Public gblnWindowMultis As Boolean
Public gblnWindowBack   As Boolean
Public gblnWindowLine   As Boolean
Public gblnWindowThick  As Boolean
Public gblnWindowLayer  As Boolean
Public gblnWindowPick   As Boolean
#Else
    #If active Then
    Public gblnMainStarted  As Boolean
    Public g_lngAutomation   As Long
    Public g_frmSplash      As frmForm
    Public g_Flowchart      As FlowChart
    #End If
#End If

'Flag for FlowChart class to tell only once that no Printer available
Public gblnPrnErrMsg As Boolean
Public gstrErrorMsg  As String 'used in this module and in FlowChart class

#If active Then
Public Sub ConditionSplash(ByVal objSplash As frmForm)
#Else
Private Sub ConditionSplash(ByVal objSplash As frmForm)
#End If
    objSplash.MenusVisible = False
    objSplash.Caption = "Loading..."
    objSplash.Move Screen.Width / 4, Screen.Height / 4, Screen.Width / 2, Screen.Height / 2
    objSplash.SetZoom 100
    objSplash.Show
    objSplash.picPreview.Refresh
End Sub

Public Sub CopyProperties(ByVal Org As Properties, ByVal Dest As Properties)
    Dest.EmulationNumber = Org.EmulationNumber
    Dest.ArrowEngg = Org.ArrowEngg
    Dest.ArrowSize = Org.ArrowSize
    Dest.CanEdit = Org.CanEdit
    Dest.BackColour = Org.BackColour
    Dest.DrawOrder = Org.DrawOrder
    Dest.FillStyle = Org.FillStyle 'or Arrow Solid property
    Dest.FontFace = Org.FontFace
    Dest.ForeColour = Org.ForeColour
    Dest.Height = Org.Height
    Dest.Left = Org.Left
    Dest.LineStyle = Org.LineStyle
    Dest.LineWidth = Org.LineWidth
    Dest.Name = Org.Name 'V8
    Dest.Tag1 = Org.Tag1
    Dest.Tag3 = Org.Tag3
    Dest.Text = Org.Text
    Dest.TextAlign = Org.TextAlign
    Dest.TextBold = Org.TextBold
    Dest.TextColour = Org.TextColour
    Dest.TextItalic = Org.TextItalic
    Dest.TextSize = Org.TextSize
    Dest.TextUnderline = Org.TextUnderline
    Dest.Top = Org.Top
    Dest.Width = Org.Width
    Dest.ZBase = Org.ZBase
    Dest.ZHeight = Org.ZHeight
    'Group and Layer numbers are not copied
End Sub





Public Sub DrawTextForLines(ByVal View As Object, ByVal FI As FlowItem, ByVal Parent As FlowChart)
    Dim r As Long
    
    If Len(FI.P.Text) = 0 Then Exit Sub
    SetFontProps View, FI, Parent
    With FI.P
        Select Case .Width * .Height
        Case Is < 0 'negative
            View.CurrentX = .CenterX + conIndent
            View.CurrentY = .CenterY + conIndent
        Case Is > 0 'positive
            View.CurrentX = .CenterX + conIndent
            View.CurrentY = .CenterY - conIndent - View.TextHeight(.Text)
        Case Else
            If .Width = 0 Then 'up-down
                View.CurrentX = .CenterX + conIndent
                View.CurrentY = .CenterY - View.TextHeight(.Text) / 2
            Else '.Height = 0 then 'left-right
                View.CurrentX = .CenterX - View.TextWidth(.Text) / 2
                View.CurrentY = .CenterY - View.TextHeight(.Text) - conIndent
            End If
        End Select
        r = SetBkMode(View.hdc, TRANSPARENT)
        If InStr(1, .Text, vbCr) > 0 Then 'remove CR-LF
            View.Print Left$(.Text, InStr(1, .Text, vbCr) - 1)
        Else
            View.Print .Text
        End If
    End With
End Sub

Public Function Duplicate(Optional ByVal Source As FlowItem, Optional ByVal AddNumber As FAddType) As FlowItem
    Dim objNew As FlowItem
        
    If Not Source Is Nothing And AddNumber = 0 Then
        AddNumber = GetFItemNo(Source)
    End If
    
    Select Case AddNumber
    Case conAddCircle:      Set objNew = New FCircle
    Case conAddDecision:    Set objNew = New FDecision
    Case conAddEndArrowLine: Set objNew = New FArrowLine
    Case conAddInOut:       Set objNew = New FInOut
    Case conAddLine:        Set objNew = New FLine
    Case conAddMidArrowLine: Set objNew = New FMidArrowLine
    Case conAddPicture:     Set objNew = New FPicture
    Case conAddRect:        Set objNew = New FRect
    Case conAddTerminator:  Set objNew = New FTerminator
    Case conAddText:        Set objNew = New FText
    Case conAddButton:      Set objNew = New FButton
    Case conAddEllipse:     Set objNew = New FEllipse
    Case conAddExtra1 To conAddExtra3, conAddShapeA To conAddShapeF
                            Set objNew = New FEmulation
    Case Else:              Set objNew = New FlowItem
    End Select
    
    objNew.P.EmulationNumber = AddNumber
    
    If Not Source Is Nothing Then
        CopyProperties Source.P, objNew.P
    End If
    Set Duplicate = objNew
    Set objNew = Nothing
End Function




Public Function GetFItemNo(ByVal Sample As FlowItem) As Long
    GetFItemNo = Sample.Number
'    Select Case True 'See Also: SetToEntry
'        Case TypeOf Sample Is FRect: GetFItemNo = conAddRect
'        Case TypeOf Sample Is FPicture: GetFItemNo = conAddPicture
'        Case TypeOf Sample Is FInOut: GetFItemNo = conAddInOut
'        Case TypeOf Sample Is FDecision: GetFItemNo = conAddDecision
'        Case TypeOf Sample Is FTerminator: GetFItemNo = conAddTerminator
'        Case TypeOf Sample Is FCircle: GetFItemNo = conAddCircle
'        Case TypeOf Sample Is FLine: GetFItemNo = conAddLine
'        Case TypeOf Sample Is FArrowLine: GetFItemNo = conAddEndArrowLine
'        Case TypeOf Sample Is FMidArrowLine: GetFItemNo = conAddMidArrowLine
'        Case TypeOf Sample Is FText: GetFItemNo = conAddText
'        Case Else: GetFItemNo = 0
'    End Select
End Function

Public Function GetFItemStr(ByVal Sample As FlowItem) As String
    Select Case True 'See Also: SetToEntry
        Case TypeOf Sample Is FRect: GetFItemStr = "Rect"
        Case TypeOf Sample Is FPicture: GetFItemStr = "Picture"
        Case TypeOf Sample Is FInOut: GetFItemStr = "InOut"
        Case TypeOf Sample Is FDecision: GetFItemStr = "Decision"
        Case TypeOf Sample Is FTerminator: GetFItemStr = "Terminator"
        Case TypeOf Sample Is FCircle: GetFItemStr = "Circle"
        Case TypeOf Sample Is FLine: GetFItemStr = "Line"
        Case TypeOf Sample Is FArrowLine: GetFItemStr = "ArrowLine"
        Case TypeOf Sample Is FMidArrowLine: GetFItemStr = "MidArrowLine"
        Case TypeOf Sample Is FText: GetFItemStr = "Text"
        Case TypeOf Sample Is FButton: GetFItemStr = "Button"
        Case Else: GetFItemStr = "?"
    End Select
End Function



Public Function IsBetween(Lower As Long, Value As Long, Upper As Long) As Boolean
    IsBetween = (Lower <= Value And Value <= Upper)
End Function

Public Function IsBetweenSng(Lower As Single, Value As Single, Upper As Single) As Boolean
    IsBetweenSng = (Lower <= Value And Value <= Upper)
End Function


#If Prev Then
Public Sub Main()
    CPlusPlus
End Sub
#End If

Public Sub SetEnabled(Form As Form, State As Boolean)
    Dim objItem As Object
    
    On Error Resume Next
    For Each objItem In Form.Controls
        If Not (TypeOf objItem Is Line Or TypeOf objItem Is Shape) Then
            objItem.Enabled = State
        End If
    Next objItem
    Err.Clear
End Sub

'GetText... was here before it was moved into FlowChart class
Public Function SetFontNmSz(ByVal Dest As Object, FontName As String, ByVal Size As Currency, ByVal Parent As FlowChart) As StdFont
    Const conAdjust = 0.15
    Dim curSize As Currency
    Dim destFont As StdFont
    
    Set destFont = Dest.Font
    destFont.Name = IIf(Len(FontName) > 0, FontName, Parent.FontName)
    curSize = IIf(Size > 0, Size, Parent.FontSize) * Parent.Percentage
    destFont.Size = curSize
    
    If destFont.Size > curSize Then 'if destination font too big reduce size
        curSize = curSize - conAdjust * Parent.Percentage
        destFont.Size = curSize
'        If DestFont.Size > curSize Then
'            curSize = curSize - 0.25
'            DestFont.Size = curSize
'        End If
    ElseIf destFont.Size < curSize And Not TypeOf Dest Is Printer Then 'if destination font too small increase size
        curSize = curSize + conAdjust * Parent.Percentage
        destFont.Size = curSize
    End If
    
    'the Printer font is usually larger
    'than the screen font.  Adjusting this
    'should help achieve WYSIWYG.
    Set SetFontNmSz = destFont
    Set destFont = Nothing
End Function

Public Sub SetTextSizeFactor(Obj As FlowItem, SizeFactor As Integer, DefaultSize As Integer)
'Called by frmMain when TextSize changed from Extra Small to Extra Large
'    Dim intSizeFactor   As Integer
'    Dim strBuild        As String
'    Dim lngLoop         As Long
'
'    If Len(Obj.Tag1) Then 'old style
'        'remove + and -
'        For lngLoop = 1 To Len(Obj.Tag1)
'            If Mid$(Obj.Tag1, lngLoop, 1) <> "+" And Mid$(Obj.Tag1, lngLoop, 1) <> "-" Then
'                strBuild = strBuild & Mid$(Obj.Tag1, lngLoop, 1)
'            End If
'        Next lngLoop
'        'rebuild + -
'        If intSizeFactor > 0 Then
'            Obj.Tag1 = Obj.Tag1 & String(intSizeFactor, "+")
'        ElseIf intSizeFactor < 0 Then
'            Obj.Tag1 = Obj.Tag1 & String(-intSizeFactor, "-")
'        End If
'    Else
    Obj.P.TextSize = DefaultSize + SizeFactor * conSizeFactorMultiplier
'    End If
End Sub


Public Sub SetFontProps(ByVal Dest As Object, ByVal FI As FlowItem, ByVal Parent As FlowChart)
    Dim objViewFont As StdFont
    
    On Error Resume Next
'    Set objViewFont = Dest.Font
'    'retain original (default) font settings
'    strFaceName = objViewFont.Name
'    curFontSize = objViewFont.Size
    
    'font formatting
    Dest.ForeColor = FI.P.TextColour
    SetFontNmSz Dest, FI.P.FontFace, FI.P.TextSize, Parent
    Set objViewFont = Dest.Font
'    objViewFont.Size = FI.P.TextSize
'    If Len(FI.P.FontFace) Then 'font name
'        objViewFont.Name = FI.P.FontFace
'    Else
'        objViewFont.Name = Parent.FontName
'    End If
'    If FI.P.TextSize > 0 Then 'fixed text size
'        If TypeOf Dest Is Printer Then
'            SetPrinterFontSize Dest, FI, Parent
'            'objViewFont.Size = Fi.P.TextSize
'        Else
'            objViewFont.Size = FI.P.TextSize * Parent.Percentage
'        End If
'    Else
'        objViewFont.Size = FI.P.TextSize * Parent.Percentage
'    End If
    
    'Tag1 formatting not accepted internally
    objViewFont.Bold = FI.P.TextBold Or Parent.ForceBold
    objViewFont.Italic = FI.P.TextItalic
    objViewFont.Underline = FI.P.TextUnderline
    
    Set objViewFont = Nothing 'release font object

    If Err Then
        gstrErrorMsg = gstrErrorMsg & Err.Description & vbNewLine
        Err.Clear
    End If
End Sub



#If Prev = 0 Then
Public Sub RegisterFlowFile()
'Registers flow chart files to the registry.
    Dim objRegistry As RegEntry 'real registry control
    Dim strAppFile  As String
    
    strAppFile = App.Path
    If Right$(strAppFile, 1) = "\" Then 'add in the \ in necessary locations
        strAppFile = strAppFile & App.EXEName & ".exe"
    Else
        strAppFile = strAppFile & "\" & App.EXEName & ".exe"
    End If
    
    Set objRegistry = New RegEntry
    
    objRegistry.KeyRoot = HKEY_CLASSES_ROOT
    objRegistry.SaveSetting ".flc", "", "flcfile"
    objRegistry.SaveSetting "flcfile", "", "Flow Chart File"
    objRegistry.SaveSetting "flcfile\DefaultIcon", "", strAppFile & ",1"
    objRegistry.SaveSetting "flcfile\Shell", "", ""
    objRegistry.SaveSetting "flcfile\Shell\Open\Command$", "", strAppFile & " ""%1"""
    objRegistry.SaveSetting "flcfile\Shell\Preview\Command$", "", strAppFile & " /P ""%1"""
    
    Set objRegistry = Nothing
End Sub
#End If

'FI - nothing, restores original settings
'FI - sets to FI properties
Public Sub SetDrawProps(ByVal Dest As Object, ByVal FI As FlowItem, ByVal Parent As FlowChart)
    Dim intPx As Integer
    
    If Not FI Is Nothing Then
        Dest.DrawStyle = FI.P.LineStyle
        If FI.P.LineStyle = vbSolid And FI.P.LineWidth > 0 Then
            intPx = FI.P.GetLineWidthPx(Dest)
            If Parent.ZoomPercent >= 100 Then
                Dest.DrawWidth = intPx * Parent.Percentage
            ElseIf CInt(intPx * Parent.Percentage) > 0 Then '1 or above
                Dest.DrawWidth = intPx * Parent.Percentage
            Else
                Dest.DrawWidth = 1
            End If
        Else
            Dest.DrawWidth = 1 'only for dashes, dots
        End If
'        If TypeOf Dest Is Printer Then
'            Dest.DrawWidth = IIf(FI.P.LineStyle = vbSolid, 2, 1) 'only thin lines can have style formatting on Printer
'        ElseIf Parent.ZoomPercent > 100 Then  'zoom
'            Dest.DrawWidth = IIf(FI.P.LineStyle = vbSolid, 2, 1) 'only thin lines can have style formatting when others thickened
'        Else
'            Dest.DrawWidth = 1 'standard
'        End If
        Dest.FillStyle = FI.P.FillStyle
        Dest.FillColor = FI.P.BackColour
        Dest.ForeColor = FI.P.ForeColour
    Else
        Dest.DrawWidth = 1
        Dest.DrawStyle = vbSolid
        Dest.FillStyle = vbFSTransparent
        Dest.FillColor = vbWhite
        Dest.ForeColor = vbBlack
    End If
End Sub

'SetText... was moved into FlowChart class

Function GetRect(FlowItem As FlowItem) As Rect
    With GetRect
        .X1 = FlowItem.P.Left
        .Y1 = FlowItem.P.Top
        .X2 = FlowItem.P.Left + FlowItem.P.Width
        .Y2 = FlowItem.P.Top + FlowItem.P.Height
    End With
End Function

#If Prev = 0 Then
Public Function InputBox(Prompt As String, Title As String, Optional ByVal Default As String) As String
'Text Box
    Load frmInput
    frmInput.lblPrompt = Prompt
    frmInput.Caption = Title
    frmInput.txtEdit = Default
    frmInput.Show vbModal
    If frmInput.mblnOK Then
        InputBox = frmInput.txtEdit
    Else
        InputBox = Default
    End If
    Unload frmInput
End Function
#End If

#If Prev = 0 Then
Public Function InputListBox(Prompt As String, Title As String, ByVal Default As String, ParamArray List() As Variant) As String
    Dim lngItem As Long
    
    Load frmInput
    frmInput.lblPrompt = Prompt
    frmInput.Caption = Title
    frmInput.txtEdit.Visible = False
    frmInput.cboList.Visible = True
    On Error Resume Next
    If Not IsEmpty(List) Then
        For lngItem = LBound(List) To UBound(List)
            frmInput.cboList.AddItem List(lngItem)
        Next lngItem
        frmInput.cboList = Default 'sometimes it won't select
        If Err Then
            Err.Clear
            frmInput.cboList.AddItem Default
            frmInput.cboList = Default
        End If
    Else
        frmInput.cboList.AddItem Default
        frmInput.cboList = Default
    End If
    frmInput.Show vbModal
    If frmInput.mblnOK Then
        If frmInput.cboList.ListIndex = -1 Then 'nothing selected
            InputListBox = Default
        Else
            InputListBox = frmInput.cboList
        End If
    Else
        InputListBox = Default
    End If
    Unload frmInput
End Function
#End If

#If Prev = 0 Then
Public Function InputListBox2(Prompt As String, Title As String, ByVal Default As Integer, ParamArray List() As Variant) As Integer
'Starts at 0.
    Dim lngItem As Long
    
    Load frmInput
    frmInput.lblPrompt = Prompt
    frmInput.Caption = Title
    frmInput.txtEdit.Visible = False
    frmInput.cboList.Visible = True
    On Error Resume Next
    If Not IsEmpty(List) Then
        For lngItem = LBound(List) To UBound(List)
            frmInput.cboList.AddItem List(lngItem)
        Next lngItem
        frmInput.cboList.ListIndex = Default 'sometimes it won't select
        If Err Then
            Err.Clear
        End If
    Else
        frmInput.cboList.AddItem Default
        frmInput.cboList.ListIndex = 0 'select only item
    End If
    frmInput.Show vbModal
    If frmInput.mblnOK Then
        If frmInput.cboList.ListIndex = -1 Then 'nothing selected
            InputListBox2 = Default
        Else
            InputListBox2 = frmInput.cboList.ListIndex
        End If
    Else
        InputListBox2 = Default
    End If
    Unload frmInput
End Function
#End If


Public Sub Main()
    #If Prev = 0 Then
    Dim objRegistry As RegEntry
    #Else
    Dim objFile     As FlowChart
    #End If
    Dim strFile     As String
    Dim blnPrev     As Boolean
    Dim objSplash   As frmForm
    
    On Error GoTo Handler
       
    Set objSplash = New frmForm
    Set objSplash.mFlowChart = New FlowChart
    
    #If active Then
    gblnMainStarted = True
    Set g_frmSplash = objSplash 'hold ref for future use
    Set g_Flowchart = objSplash.mFlowChart
    #End If
    
    
    CreateSplash objSplash.mFlowChart
    ConditionSplash objSplash
    
    'read the Command$ line
    If Len(Trim$(Command$)) Then
        #If Prev Then 'only choice
        blnPrev = True
        #End If
        
        Select Case UCase$(Left$(Command$, 3))
        Case "/O "
            objSplash.mFlowChart(2).P.Text = objSplash.mFlowChart(2).P.Text & vbNewLine & "Command: Open " & Mid$(Command$, 4)
            strFile = Mid$(Command$, 4)
        Case "/P ", "/P"
            objSplash.mFlowChart(2).P.Text = objSplash.mFlowChart(2).P.Text & vbNewLine & "Command: Print " & Mid$(Command$, 4)
            strFile = Mid$(Command$, 4)
            blnPrev = True
        Case "/H ", "/H", "/? ", "/?"
            objSplash.mFlowChart(2).P.Text = objSplash.mFlowChart(2).P.Text & vbNewLine & "Command: Help"
'            frmLog.Show vbModeless, frmMain
            
            Load frmVal
            frmVal.AddItem "Open:", "[/O] <filename>"
            frmVal.AddItem "Preview:", "/P <filename>"
            frmVal.AddItem "Help:", "/H or /?"
            frmVal.Show vbModal, objSplash
        Case Else
            If LCase$(Command$) = "-embedding" Then
                objSplash.mFlowChart(2).P.Text = objSplash.mFlowChart(2).P.Text & vbNewLine & "Command: Embedded"
            ElseIf Left$(Command$, 1) <> "/" Then
                strFile = Command$
                objSplash.mFlowChart(2).P.Text = objSplash.mFlowChart(2).P.Text & vbNewLine & "Command: Open " & Mid$(Command$, 4)
            Else
                objSplash.mFlowChart(2).P.Text = objSplash.mFlowChart(2).P.Text & vbNewLine & "Command: (invalid)"
                MsgBox "Command not understood: " & vbNewLine & Command$, vbInformation
            End If
        End Select
        
        objSplash.picPreview.Refresh
        
        If blnPrev Then
            Load frmForm
            With frmForm
                Set .mFlowChart = New FlowChart
                If Len(strFile) Then
                    If .mFlowChart.Load(strFile, .picPreview) <> 0 Then
                        MsgBox "Error loading flow chart.  " & strFile, vbExclamation
                        .SetZoom 50
                    Else
                        .SetZoom .mFlowChart.ZoomPercent
                    End If
                End If
                .mnuFileOpen.Visible = True
                .Show
            End With
        Else
            #If Prev = 0 Then
            Set frmMain = New frmIMain
            Load frmMain
            If Len(strFile) Then
                'remove quotation marks from file
                If Left$(strFile, 1) = """" And Right$(strFile, 1) = """" Then
                    strFile = Mid$(strFile, 2, Len(strFile) - 2)
                End If
                'load the file
                frmMain.OpenFile strFile
            End If
                        
            frmMain.MainShow
            #End If
        End If
        Screen.MousePointer = vbDefault
    Else
        #If Prev Then
            #If active Then
            If App.StartMode = 0 Then
            #End If
                'only show if not activeX startup
                Set frmForm.mFlowChart = New FlowChart
                frmForm.Show
            #If active Then
            End If
            #End If
        #Else
            Set frmMain = New frmIMain
            Load frmMain 'load main form
            
            objSplash.picPreview.Refresh
            Sleep 500
            
            frmMain.MainShow
        #End If
    End If
    
    #If Prev = 0 Then
    'see if file associations are good
    Set objRegistry = New RegEntry
    
    objRegistry.KeyRoot = HKEY_CLASSES_ROOT
    If StrComp(objRegistry.GetSetting(".flc", "", ""), "flcfile", vbTextCompare) <> 0 Then 'nothing
        If MsgBox("Flow Chart file extensions are not registered.  Do you want extensions to be registered?", vbQuestion Or vbYesNo) = vbYes Then
            'after asking for the user's permission, register flow chart files
            RegisterFlowFile
        End If
    End If
    Set objRegistry = Nothing
    #Else
        #If active Then
        'register with Univector
        Register.AddApp "FlowView", "Flowview.FlowChartApplication"
        #End If
    #End If
    
    Unload objSplash
    Set objSplash = Nothing
    
    Exit Sub
Handler:

    #If Prev Then
    If Len(Command$) > 0 Or Len(strFile) > 0 Then
        MsgBox "Print feature " & Command$

    Else
        MsgBox "Failure loading Flow Chart.  " & Err.Description & ".  (" & Err.Number & ")  If you want to print a file, you can drag and drop the file onto the executable file.  It should ask to print.", vbExclamation
    End If
    
    Unload frmForm
    Unload objSplash
    Set objSplash = Nothing
    
    #If active Then
        'release pointers
        Set g_frmSplash = Nothing
        Set g_Flowchart = Nothing
    #End If
    
    #Else
    
    Unload objSplash
    Set objSplash = Nothing
    
    Select Case MsgBox("Failure loading Flow Chart.  " & Err.Description & ".  (" & Err.Number & ")", vbExclamation Or vbAbortRetryIgnore)
    Case vbAbort
        Exit Sub
    Case vbRetry
        Resume
    Case vbIgnore
        Resume Next
    End Select
    #End If
End Sub

'used by Main() sub
Private Sub CreateSplash(ByVal File As FlowChart)
    Dim strTitle As String
    Dim objAdd   As FlowItem
    
    File.PScaleWidth = Screen.Width / 2 - 40 * Screen.TwipsPerPixelX
    File.PScaleHeight = Screen.Height / 2 - 40 * Screen.TwipsPerPixelY
    
    Set objAdd = File.AddParam(New FText, 0, File.PScaleHeight / 6, File.PScaleWidth, File.PScaleHeight / 4, "Flow Chart")
    objAdd.P.FontFace = "Arial"
    objAdd.P.TextAlign = DT_CENTER
    objAdd.P.TextSize = 18
    objAdd.P.TextBold = True
    
    strTitle = "Version " & App.Major & "." & Format(App.Minor, "#00") & "." & Format(App.Revision, "#00")
    
    Set objAdd = File.AddParam(New FText, 0, File.PScaleHeight / 2, File.PScaleWidth, File.PScaleHeight / 4, strTitle)
    objAdd.P.TextAlign = DT_CENTER
    objAdd.P.FontFace = "Courier New"
    objAdd.P.TextColour = QBColor(3) 'dark blue
    objAdd.P.TextBold = True
    
    Set objAdd = File.AddParam(New FText, 0, File.PScaleHeight * 2 / 3, File.PScaleWidth, File.PScaleHeight / 4, "Program developed by" & vbNewLine & "Richard H Fung")
    objAdd.P.TextAlign = DT_CENTER
    objAdd.P.TextColour = QBColor(1) 'dark blue
    objAdd.P.TextBold = True
    
    Set objAdd = File.AddParam(New FMidArrowLine, File.PScaleWidth / 15, File.PScaleHeight / 3, File.PScaleWidth / 5, File.PScaleHeight / 2, "")
    objAdd.P.LineWidth = 2
    objAdd.P.ForeColour = QBColor(8)
End Sub



Public Sub DrawText(ByVal View As Object, ByVal FI As FlowItem, ByVal Parent As FlowChart)  ', Optional WindowFont As StdFont)
'DrawText should be called after the rectangles and other
'shapes are drawn to prevent text from disappearing on
'some printing systems.
    Dim r           As Long
    Dim strString   As String
'    Dim strRect     As String
    Dim typRect     As apiRECT
'    Dim curFontSize As Currency
'    Dim blnBold     As Boolean
'    Dim blnItalic   As Boolean
'    Dim blnUnderline As Boolean
'    Dim strFaceName As String 'font face name

    If Len(FI.P.Text) = 0 Then Exit Sub
    
    SetFontProps View, FI, Parent
    
    'draw box
    strString = FI.P.Text & vbNullChar
    
    With typRect
        'Int() erases the decimal values following the decimal point
        .Left = Int(View.ScaleX(FI.P.Left + FI.TextLeftMrg + conIndent, View.ScaleMode, vbPixels))
        .Top = Int(View.ScaleY(FI.P.Top + FI.TextTopMrg + conIndent, View.ScaleMode, vbPixels))
        'CLng() rounds accordingly
        .Right = CLng(View.ScaleX(FI.P.Left + FI.P.Width - FI.TextRightMrg - conIndent, View.ScaleMode, vbPixels))
        .Bottom = CLng(View.ScaleY(FI.P.Top + FI.P.Height - FI.TextBottomMrg - conIndent, View.ScaleMode, vbPixels))
    End With
    
    'DT_NOCLIP turns off clipping.  The text can go out of the box if this flag is set.
    r = SetBkMode(View.hdc, TRANSPARENT)
    r = apiDrawText(View.hdc, strString, -1, typRect, FI.P.TextAlign Or DT_NOPREFIX Or DT_WORDBREAK)
    'error return is 0
'    If r = 0 Then
'        'GetLastError
'        'FormatMessage
'        Dim lpstrBuffer As String
'        lpstrBuffer = Space(100) & vbNull
'        FormatMessage &H1000 Or &H200, 0, GetLastError(), 0, lpstrBuffer, Len(lpstrBuffer), 0
'        MsgBox lpstrBuffer, , "Error"
'    End If
    
    'restore original settings
'    objViewFont.Name = strFaceName
'    objViewFont.Size = curFontSize
'    objViewFont.Bold = blnBold
'    objViewFont.Italic = blnItalic
'    objViewFont.Underline = blnUnderline
    If Err Then
        gstrErrorMsg = gstrErrorMsg & Err.Description & vbNewLine
        Err.Clear
    End If
    
End Sub
'    Dim tLeft As Single, tTop As Single
'    Dim tRight As Single, tBottom As Single
'    Dim lngPos As Long, ch As String
'    Dim lngP2 As Long, c2 As String * 1
'    #If VB6 Then
'        Dim vntText() As String
'    #Else
'        Dim vntText As Variant
'    #End If
'
'    tLeft = Fi.P.Left + FI.TextLeftMrg + conIndent
'    tTop = Fi.P.Top + FI.TextTopMrg + conIndent
'    tRight = Fi.P.Left + FI.P.Width - FI.TextRightMrg - conIndent
'    tBottom = Fi.P.Top + Fi.P.Height - FI.TextBottomMrg - conIndent
'
'    'cover up background
'    If View Is Printer Then 'must be white only
'        View.Line (tLeft, tTop)-(tRight, tBottom), vbWhite, BF
'    Else 'use the object's backcolor
'        View.Line (tLeft, tTop)-(tRight, tBottom), View.BackColor, BF
'    End If
'
'    View.CurrentX = tLeft
'    View.CurrentY = tTop
'
'    vntText = Split(Fi.P.Text)
'    For lngPos = LBound(vntText) To UBound(vntText)
'        ch = vntText(lngPos)
'        If View.TextWidth(ch) >= (tRight - tLeft) Then
'            For lngP2 = 1 To Len(ch)
'                c2 = Mid$(ch, lngP2, 1)
'                If View.CurrentX + View.TextWidth(c2) > tRight + conIndent Then 'indents are extra space
'                    View.Print
'                    View.CurrentX = tLeft
'                End If
'                If View.CurrentY > tBottom - View.TextHeight("Hello") Then
'                    View.Print "..." 'show user out of space
'                    Exit For
'                End If
'                View.Print c2;
'            Next lngP2
'            View.Print " ";
'        Else
'            If View.CurrentX + View.TextWidth(ch) > tRight Then
'                View.Print
'                View.CurrentX = tLeft
'            End If
'            If View.CurrentY > tBottom - View.TextHeight("Hello") Then
'                View.Print "..." 'show user out of space
'                Exit For
'            End If
'            View.Print ch; " ";
'        End If
'    Next lngPos


#If VB6 Then
Public Function DiverseReplace(Text As String, ByVal Compare As VbCompareMethod, ParamArray Patterns() As Variant) As String
    Dim strNewString As String
    Dim intUBound    As Integer
    Dim intLBound    As Integer
    Dim intCompare   As Integer
    
    If Text = "" Then
        Let DiverseReplace = Text
        Exit Function
    End If

    strNewString = Text
    'defing lower and upper bounds
    intLBound = LBound(Patterns)
    intUBound = UBound(Patterns)
    'search for strings and replace them
    For intCompare = intLBound To intUBound Step 2
        If intCompare + 1 <= intUBound Then
            'Function Replace(Expression As String, Find As String, Replace As String, [Start As Long = 1], [Count As Long = -1], [Compare As VbCompareMethod = vbBinaryCompare]) As String
            strNewString = Replace(strNewString, Patterns(intCompare), Patterns(intCompare + 1), , , Compare)
        End If
    Next intCompare

    Let DiverseReplace = strNewString
End Function
#Else
Public Function DiverseReplace(Text As String, ByVal Compare As VbCompareMethod, ParamArray Patterns() As Variant) As String
    Dim lngCount     As Long
    Dim strNewString As String
    Dim intCompare   As Integer
    Dim blnDone      As Boolean
    Dim intLen       As Integer
    Dim intUBound    As Integer
    Dim intLBound    As Integer
    
    If Text = "" Then
        Let DiverseReplace = Text
        Exit Function
    End If

    'defing lower and upper bounds
    intLBound = LBound(Patterns)
    intUBound = UBound(Patterns)
    'search for strings and replace them
    For lngCount = 1 To Len(Text)
        blnDone = False
        For intCompare = intLBound To intUBound Step 2
            intLen = Len(Patterns(intCompare))
            If StrComp(Mid$(Text, lngCount, intLen), Patterns(intCompare), Compare) = 0 And intCompare + 1 <= intUBound Then
                strNewString = strNewString & Patterns(intCompare + 1)
                lngCount = lngCount + intLen - 1
                blnDone = True
                Exit For
            End If
        Next intCompare
        If Not blnDone Then
            strNewString = strNewString & Mid$(Text, lngCount, 1)
        End If
    Next lngCount

    Let DiverseReplace = strNewString
End Function


'Split(expression[, delimiter[, count[, compare]]])
'  Zero-based array
Public Function Split(Expression As String, Optional Delimiter As String = " ", Optional ByVal Count As Long = -1, Optional ByVal Compare As VbCompareMethod = vbBinaryCompare) As Variant
    Dim strArray() As String
    Dim lngArray   As Long
    Dim lngChPos   As Long
    Dim lngNext    As Long
    Dim strPart    As String
    
    If Len(Expression) Then
        lngChPos = 1
        lngNext = InStr(lngChPos, Expression, Delimiter, Compare)
        If lngNext > 0 Then
            strPart = Mid$(Expression, lngChPos, lngNext - lngChPos)
        Else
            strPart = Mid$(Expression, lngChPos)
        End If
        Do
            ReDim Preserve strArray(0 To lngArray)
            strArray(lngArray) = strPart
            lngArray = lngArray + 1
            
            If lngNext = 0 Then
                Exit Do
            Else
                lngChPos = lngNext + Len(Delimiter)
            End If
            
            lngNext = InStr(lngChPos, Expression, Delimiter, Compare)
            If lngNext > 0 Then
                strPart = Mid$(Expression, lngChPos, lngNext - lngChPos)
            Else
                strPart = Mid$(Expression, lngChPos)
            End If
        Loop Until ((lngArray >= Count) And (Count <> -1))
    Else
        ReDim strArray(0)
    End If
    Split = strArray
End Function
#End If

