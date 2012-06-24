VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FlowChart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"FlowItem"
Attribute VB_Ext_KEY = "Member0" ,"FlowItem"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
'Richard Fung.  August 12, 2000.
'1-based collection

Public FileName     As String

'File V2
Public FontName     As String
Public FontSize     As Currency
Public ZoomPercent  As Integer
Public Version      As Single '#.# format

'File V4
Public ScrollX      As Integer 'V4 px V5 %
Public ScrollY      As Integer 'V4 px V5 %
Public PaperSize    As Integer
Public Orientation  As Integer
Public PScaleWidth  As Single 'px V5 twips
Public PScaleHeight As Single 'px V5 twips
Public Header1PDevName As String 'printer device name e.g. HP LaserJet 4L
'Public Header2      As String
'Public Header3      As String
'Public Header4      As String
'Public Header5      As String

'File V8
Public UnitScale    As ScaleModeConstants
Public Groups       As Groups

'Run-time variables are not saved
Public Changed      As Boolean
Public AskToSave    As Boolean 'before spell check? Default: Yes
Public PrinterError As Boolean 'on Initialize
Public ForceBold    As Boolean 'ignored by PrintFile
Public ShowSelected As Boolean 'ignored by PrintFile

Public Customized   As Boolean 'for conAddExtra1
Public mlngBackColour As Long
Public DefaultLayer  As Long 'changes using Layer tool window
Public Layers       As Layers 'Default Layer information saved at end of version 7 & 8 file

Private m_Convert   As FlowOldFile
Private Const conPx = 4.8 'see also FlowOldFile

'local variable to hold collection
Private mCol        As Collection


'It is a good idea to have Public/Private
'declared on all declaration items to
'avoid confusion.
Private Type FTHeaderData '--do not change order
    Version     As Single
    FontName    As String
    FontSize    As Currency
    Zoom        As Integer 'won't be too large
    ScrollX     As Integer 'like in V/H ScrollBar
    ScrollY     As Integer
    PaperSize   As Integer 'For Printer
    Orientation As Integer 'For Printer
    PScaleWidth As Single 'For Printer
    PScaleHeight As Single 'For Printer
    Header1PDevName As String
    Header2     As String
    Header3     As String
    Header4     As String
    Header5     As String
End Type

Private Type FTHeaderDataV8 '--do not change order
    Version     As Single
    FontName    As String
    FontSize    As Currency
    Zoom        As Integer 'won't be too large
    ScrollX     As Integer 'like in V/H ScrollBar
    ScrollY     As Integer
    PaperSize   As Integer 'For Printer
    Orientation As Integer 'For Printer
    PScaleWidth As Single 'For Printer
    PScaleHeight As Single 'For Printer
    UnitScale   As Long 'inch, cm, mm, etc.
    DeviceName  As String 'header1
    Header2     As String
    Header3     As String
    Header4     As String
    Header5     As String
    Header6     As String
    Header7     As String
    Header8     As String
    Header9     As String
    Tag1        As Long
    Tag2        As Long
    Tag3        As Long
    Tag4        As Long
    Tag5        As Long
    Tag6        As Long
    Tag7        As Long
    Tag8        As Long
    Tag9        As Long
    Bool1       As Boolean
    Bool2       As Boolean
    Bool3       As Boolean
    Bool4       As Boolean
    Bool5       As Boolean
    Bool6       As Boolean
    Bool7       As Boolean
    Bool8       As Boolean
    Bool9       As Boolean
End Type

Private Type FTEntryV45 'V4 -- do not change order
    FlowItemType    As String
    Left        As Single
    Top         As Single
    Width       As Single
    Height      As Single
    Text        As String 'remember it contains text alignment
    DrawOrder   As Integer '1 to 3
    Tag1        As String 'Text: text properties; Pic: picture properties
    Tag2        As String 'Text: custom font name
    Tag3        As String
End Type

Private Type FTEntryV6
    FlowItemNo  As Long 'FAddType
    Left        As Single
    Top         As Single
    Width       As Single
    Height      As Single
    Text        As String
    DrawOrder   As Integer '1 to 3
    'absolute text properties
    FontFace    As String
    TextSize    As Currency
    ForeColour  As Long
    BackColour  As Long
    TextBold    As Boolean
    TextItalic  As Boolean
    TextAlign   As Long
    TextUnderline As Boolean
    LineStyle   As Integer
    FillStyle   As Integer
    'tags for compatibility
    Tag1        As String 'Text: relative text properties; Pic: picture properties
    'Tag2 replaced by FontFace string
    Tag3        As String
    'NON-SAVING DATA below
    Name        As String 'tag4 new expansion strings
    Tag5        As String
    TextColour  As Long 'Tag6 this is saved
    LineWidth   As Long 'Tag7 saved
    Tag8        As Boolean
End Type

Private Type FTEntryV7
    FlowItemNo  As Long 'FAddType
    Left        As Single
    Top         As Single
    Width       As Single
    Height      As Single
    Text        As String
    DrawOrder   As Integer '1 to 3
    'absolute text properties
    FontFace    As String
    TextSize    As Currency
    ForeColour  As Long
    BackColour  As Long
    TextBold    As Boolean
    TextItalic  As Boolean
    TextAlign   As Long
    TextUnderline As Boolean
    LineStyle   As Integer
    FillStyle   As Integer
    'Tag1 - Text: relative text properties; Pic: picture properties
    Tag1        As String
    'Tag2 replaced by FontFace string
    Tag3        As String
    'NON-SAVING DATA below unless otherwise specified
    'new expansion strings
    Name        As String
    Tag5        As String
    TextColour  As Long 'Tag6 saved
    LineWidth   As Long 'Tag7 saved
    Tag8        As Boolean
    Tag9        As Boolean 'Tag9
    ArrowEngg   As Boolean 'Tag10
    ArrowSize   As Long 'Tag11
    GroupNo     As Long 'Tag12
    Tag13       As Long 'tag13
    Tag14       As Long 'tag14
    Layer       As Long 'tag15
    Tag16       As String
    Tag17       As String
    Tag18       As String
    ZBase       As Single '19
    ZHeight     As Single '20
    Tag21       As Boolean
    Tag22       As Boolean
    Tag23       As Boolean
End Type

Private Const conHeader = "## FlowChart File "
Private Const conHeaderL = 18& 'header length




Public Function AddParam(ByVal Sample As FlowItem, ByVal Left As Single, ByVal Top As Single, ByVal Width As Single, ByVal Height As Single, Text As String, Optional ByVal DrawOrder As FOrder = conMiddle, Optional sKey As String) As FlowItem
    'Version: 1 to 3
    'create a new object
    Dim objNew As FlowItem
    
    If Sample Is Nothing Then
        Set objNew = New FlowItem
    Else
        Set objNew = Sample 'New FlowItem
    End If

    'set the properties passed into the method
    With objNew.P
        .Left = Left
        .Top = Top
        .Width = Width
        .Height = Height
        .Text = Text
        .DrawOrder = DrawOrder
        'layer
        .Layer = DefaultLayer
    End With
    
    If Len(sKey) = 0 Then
        mCol.Add objNew
    Else
        mCol.Add objNew, sKey
    End If
    

    

    'return the object created
    Set AddParam = objNew
    Set objNew = Nothing
End Function

Public Function Add(ByVal Sample As FlowItem, Optional sKey As String) As FlowItem
    'Version 4 and up
    'create a new object
    Dim objNew As FlowItem
    
    If Sample Is Nothing Then
        Set objNew = New FlowItem
    Else
        Set objNew = Sample 'New FlowItem
    End If

    If Len(sKey) = 0 Then
        mCol.Add objNew
    Else
        mCol.Add objNew, sKey
    End If

    'add to group
    If Sample.P.GroupNo <> 0 Then Groups.Add Sample, Sample.P.GroupNo
    
    'layer
    Sample.P.Layer = DefaultLayer
    
    'return the object created
    Set Add = objNew
    Set objNew = Nothing
End Function

#If Prev = 0 Then
Public Function AddBefore(ByVal Item As FlowItem, ByVal Before As FlowItem) As FlowItem
    'create a new object
    'if not First, then it must be Last
    Dim objNew As FlowItem
    
    If Item Is Nothing Then
        Set objNew = New FlowItem
    Else
        Set objNew = Item
    End If

    If Before Is Nothing Then
        If mCol.Count > 0 Then
            mCol.Add objNew, , 1
        Else
            mCol.Add objNew
        End If
    Else
        mCol.Add objNew, , GetIndex(Before)
    End If

    'return the object created
    Set AddBefore = objNew
    Set objNew = Nothing
End Function
#End If

Private Sub Added6Obj(ByVal Item As FlowItem)
    'cannot do this unless Version 5 because
    'of Pixel scale mode in ver 1 to 4; 5 and up are in Twips
    If Version < 5 And m_Convert.mblnConverted = False Then Exit Sub
    If Item.Number = conAddExtra1 Then
        Customized = True
        PrinterError = True
        PScaleWidth = Item.P.Left + Item.P.Width
        PScaleHeight = Item.P.Top + Item.P.Height
        If Item.P.FillStyle <> vbFSTransparent Then
            mlngBackColour = Item.P.BackColour
        Else
            mlngBackColour = -1
        End If
    End If
    'add to group
    If Item.P.GroupNo <> 0 Then Groups.Add Item, Item.P.GroupNo
End Sub

Private Sub AddFromLine()
    'version 1,2,3
    Dim strFlowItem As String 'inputs
    Dim sngLeft     As Single  '  |
    Dim sngTop      As Single  '  V
    Dim sngWidth    As Single
    Dim sngHeight   As Single
    Dim strText     As String
    Dim lngOrder    As FOrder
    Dim objAdd      As FlowItem
        
    Input #1, strFlowItem, sngLeft, sngTop, sngWidth, sngHeight, strText, lngOrder
    strText = DiverseReplace(strText, vbBinaryCompare, "''", """", vbTab, "") ' '' becomes " AND tabs erased
    Select Case LCase$(strFlowItem)
        Case "rect"
            Set objAdd = New FRect
        Case "picture"
            Set objAdd = New FPicture
        Case "inout"
            Set objAdd = New FInOut
        Case "decision"
            Set objAdd = New FDecision
        Case "terminator"
            Set objAdd = New FTerminator
        Case "circle"
            Set objAdd = New FCircle
        Case "line"
            Set objAdd = New FLine
        Case "arrowline"
            Set objAdd = New FArrowLine
        Case "midarrowline"
            Set objAdd = New FMidArrowLine
        Case "text"
            Set objAdd = New FText
        Case "button"
            Set objAdd = New FButton
        Case "ellipse"
            Set objAdd = New FEllipse
        Case Else
            Set objAdd = New FlowItem
    End Select
    AddParam objAdd, sngLeft, sngTop, sngWidth, sngHeight, strText, lngOrder
    
    If InStr(1, strText, "¿", vbBinaryCompare) < 3 Then 'remove flag formats for FileVersion6
        objAdd.P.TextAlign = m_Convert.GetTextFlags(objAdd)   'order important!
        objAdd.P.Text = m_Convert.GetText(objAdd)
    End If
End Sub

Private Sub AddFromEntry45(Entry As FTEntryV45)  'V45, LOAD
    Dim objAdd As FlowItem
    
    Select Case LCase$(Entry.FlowItemType)
    Case "rect"
        Set objAdd = New FRect
    Case "button" 'new
        Set objAdd = New FButton
    Case "picture"
        Set objAdd = New FPicture
    Case "inout"
        Set objAdd = New FInOut
    Case "decision"
        Set objAdd = New FDecision
    Case "terminator"
        Set objAdd = New FTerminator
    Case "circle"
        Set objAdd = New FCircle
    Case "line"
        Set objAdd = New FLine
    Case "arrowline"
        Set objAdd = New FArrowLine
    Case "midarrowline"
        Set objAdd = New FMidArrowLine
    Case "text"
        Set objAdd = New FText
    Case "ellipse"
        Set objAdd = New FEllipse
    Case Else
        Set objAdd = New FlowItem
    End Select
    With objAdd.P
        .DrawOrder = Entry.DrawOrder
        .Height = Entry.Height
        .Left = Entry.Left
        .Tag1 = Entry.Tag1
        'GetText... must be put after tag 1 initialized
        .TextBold = m_Convert.GetTextFormatBold(objAdd)
        .TextItalic = m_Convert.GetTextFormatItalic(objAdd)
        .TextSize = FontSize + m_Convert.GetTextFormatSizeFactor(objAdd) * conSizeFactorMultiplier
        .Tag1 = "" 'Tag1 not supported in GUI
        'end get text
        .FontFace = Entry.Tag2 'FontFace replaces Tag2
        .Tag3 = Entry.Tag3
        .Text = Entry.Text
        'read text flags & text in format "flags¿text"
        If InStr(1, Entry.Text, "¿", vbBinaryCompare) < 3 Then 'remove flag formats for FileVersion6
            .TextAlign = m_Convert.GetTextFlags(objAdd)
            .Text = m_Convert.GetText(objAdd)
        End If
        .Top = Entry.Top
        .Width = Entry.Width
    End With
    mCol.Add objAdd
End Sub

Private Sub AddFromEntry6(Entry As FTEntryV6) 'V6, LOAD
    Dim objAdd As FlowItem
    
    'See Also: AddFromEntry7
    Select Case Entry.FlowItemNo
'    Case conAddEllipse
'        MsgBox "Ellipse not available for this version of Flow Chart.", vbExclamation
    Case conAddExtra2 To conAddExtra3
        MsgBox """Extra"" not available for this version of Flow Chart.", vbExclamation
    Case conAddShapeA To conAddShapeF
        MsgBox "New ""shape"" not available for this version of Flow Chart.", vbExclamation
    End Select
    
    Set objAdd = Duplicate(, Entry.FlowItemNo)
    With objAdd.P
        'FlowItemNo cannot be changed by variable
        .EmulationNumber = Entry.FlowItemNo
        .Height = Entry.Height
        .Left = Entry.Left
        .Tag1 = Entry.Tag1
        'removed objAdd.Tag2 = Entry.Tag2
        .Tag3 = Entry.Tag3
        .Text = Entry.Text
        .Top = Entry.Top
        .Width = Entry.Width
        .DrawOrder = Entry.DrawOrder
        
        .BackColour = Entry.BackColour
        .FillStyle = Entry.FillStyle
        If Entry.FlowItemNo = conAddEndArrowLine Then 'must be stated
            .FillStyle = vbFSSolid
        End If
        .FontFace = Entry.FontFace
        .ForeColour = Entry.ForeColour
        .LineStyle = Entry.LineStyle
        .LineWidth = Entry.LineWidth
        .Name = Entry.Name
        .TextAlign = Entry.TextAlign
        .TextBold = Entry.TextBold
        .TextItalic = Entry.TextItalic
        .TextSize = Entry.TextSize
        .TextUnderline = Entry.TextUnderline
        .TextColour = Entry.TextColour
        '20 items
    End With
    mCol.Add objAdd
    Added6Obj objAdd
End Sub

Private Sub AddFromEntry7(Entry As FTEntryV7) 'V6, LOAD
    Dim objAdd As FlowItem
    
    'See Also: AddFromEntry6
    Select Case Entry.FlowItemNo
    Case conAddExtra2 To conAddExtra3
        MsgBox """Extra"" not available for this version of Flow Chart.", vbExclamation
    Case conAddShapeA To conAddShapeF
        MsgBox "New ""shape"" not available for this version of Flow Chart.", vbExclamation
    End Select
    
    Set objAdd = Duplicate(, Entry.FlowItemNo)
    With objAdd.P
        'FlowItemNo cannot be changed by variable
        .EmulationNumber = Entry.FlowItemNo
        .Height = Entry.Height
        .Left = Entry.Left
        .Tag1 = Entry.Tag1
        .Tag3 = Entry.Tag3
        .Text = Entry.Text
        .Top = Entry.Top
        .Width = Entry.Width
        .DrawOrder = Entry.DrawOrder
        
        .ArrowEngg = Entry.ArrowEngg
        .ArrowSize = Entry.ArrowSize
        .BackColour = Entry.BackColour
        .FillStyle = Entry.FillStyle
        .FontFace = Entry.FontFace
        .ForeColour = Entry.ForeColour
        .GroupNo = Entry.GroupNo
        .LineStyle = Entry.LineStyle
        .LineWidth = Entry.LineWidth
        .Name = Entry.Name
        .TextAlign = Entry.TextAlign
        .TextBold = Entry.TextBold
        .TextItalic = Entry.TextItalic
        .TextSize = Entry.TextSize
        .TextUnderline = Entry.TextUnderline
        .TextColour = Entry.TextColour
        '22 items
        
        .ZBase = Entry.ZBase
        .ZHeight = Entry.ZHeight
        .Layer = Entry.Layer
    End With
    
    If objAdd.P.EmulationNumber = conAddLayerDefaults Then
        Set Layers.DefaultShape(Entry.Layer) = objAdd
    Else
        mCol.Add objAdd
        Added6Obj objAdd
    End If
End Sub


Public Function CustomBack() As Boolean
    CustomBack = CustomBack > -1 And Customized
End Function

Public Sub DrawSelected(ByVal Output As Object, ByVal IsSelected As Boolean)
    'Each object must have the Selected flag True
    Dim objItem As FlowItem
    Dim lngOrder As Long
    Const conBox = 120
    
    On Error GoTo Handler
    
    Output.DrawWidth = 2
    Output.DrawStyle = vbSolid
    If IsSelected Then
        Output.ForeColor = vbHighlight
    Else
        If CustomBack Then
            Output.ForeColor = mlngBackColour
        Else
            Output.ForeColor = Output.BackColor
        End If
    End If
    
    For lngOrder = conBottom To conTop
        For Each objItem In Me
            If objItem.P.DrawOrder = lngOrder And objItem.P.Enabled And objItem.P.Selected Then
                GoSub Draw
            End If
        Next objItem
    Next lngOrder
    
    If Not IsSelected Then 'redraw lines
        For lngOrder = conBottom To conTop
            For Each objItem In Me
                If objItem.P.DrawOrder = lngOrder And objItem.P.Enabled And objItem.P.Selected And IsObjLine(objItem) Then
                    objItem.Draw Output, Me
                End If
            Next objItem
        Next lngOrder
        
    End If
    
    
Draw: 'Selection lines

    Select Case objItem.Number
    Case conAddEndArrowLine, conAddMidArrowLine, conAddLine
        With objItem.P
            Output.Line (.Left, .Top)-Step(.Width, .Height)
        End With
    Case Else
        With objItem.P
            Output.Line (.Left - conBox, .Top)-(.Left, .Top)
            Output.Line -(.Left, .Top - conBox)
            Output.Line (.Left + .Width + conBox, .Top)-(.Left + .Width, .Top)
            Output.Line -(.Left + .Width, .Top - conBox)
            Output.Line (.Left - conBox, .Top + .Height)-(.Left, .Top + .Height)
            Output.Line -(.Left, .Top + .Height + conBox)
            Output.Line (.Left + .Width + conBox, .Top + .Height)-(.Left + .Width, .Top + .Height)
            Output.Line -(.Left + .Width, .Top + .Height + conBox)
        End With
    End Select
    Return

Handler:
    Exit Sub
End Sub

Public Property Get Font() As StdFont
    Set Font = New StdFont
    Font.Name = FontName
    Font.Size = FontSize
    Font.Bold = False
    Font.Italic = False
    Font.Underline = False
End Property

Public Function GetFilePart() As String
    Dim lngI As Long
    lngI = InStrRev(FileName, "\") + 1
    If lngI > 0 Then
        GetFilePart = Mid$(FileName, lngI)
    Else
        GetFilePart = FileName
    End If
End Function

Public Function GetIndex(ByVal Item As FlowItem) As Long
    Dim lngIndex As Long
    
    For lngIndex = 1 To Count
        If Item Is Me.Item(lngIndex) Then
            GetIndex = lngIndex
            Exit Function
        End If
    Next lngIndex
    lngIndex = 0 'no index
End Function

Public Function GetPath() As String
    Dim lngI As Long
    lngI = InStrRev(FileName, "\")
    If lngI > 0 Then
        GetPath = Left$(FileName, lngI)
    End If
End Function

Public Property Get Item(ByVal vntIndexKey As Variant) As FlowItem
Attribute Item.VB_UserMemId = 0
    'used when referencing an element in the collection
    'vntIndexKey contains either the Index or Key to the collection,
    'this is why it is declared as a Variant
    'Syntax: Set foo = x.Item(xyz) or Set foo = x.Item(5)
  Set Item = mCol(vntIndexKey)
End Property

Public Property Get Count() As Long
    Count = mCol.Count
End Property

Public Sub DrawFile(ByVal Output As Object, ByVal ForceBold As Boolean, ByVal All As Boolean, ByVal X1 As Single, ByVal Y1 As Single, ByVal X2 As Single, ByVal Y2 As Single)
'All = True Or tRect must be defined
    Dim objItem As FlowItem
    Dim blnAll  As Boolean
    Dim tView   As Rect
    Dim lngOrder As Long
    Const conBox = 120
   
    On Error Resume Next
    
    gstrErrorMsg = "" 'clear error msg buffer
    Me.ForceBold = ForceBold
    blnAll = All
    tView.X1 = X1
    tView.X2 = X2
    tView.Y1 = Y1
    tView.Y2 = Y2
    
    'adjustments
    'Do not set scalemode.
    Output.CurrentX = 0
    Output.CurrentY = 0
    If CustomBack Then
        Output.BackColor = mlngBackColour
    End If
    'FontTransparent does not work - use SetBkMode
    'Output.FontTransparent = True 'reinstate this
    
    On Error GoTo Handler
    For lngOrder = conBottom To conTop
        For Each objItem In Me
            If objItem.P.DrawOrder = lngOrder And objItem.P.Enabled Then
                GoSub DrawItem
            End If
        Next objItem
    Next lngOrder

    If ShowSelected Then DrawSelected Output, True

    If Len(gstrErrorMsg) Then 'display error message
        MsgBox gstrErrorMsg, vbExclamation, "Draw Log"
        gstrErrorMsg = "" 'clear message
    End If
    
    Me.ForceBold = False
    
    Exit Sub
Handler:
    gstrErrorMsg = gstrErrorMsg & Err.Description & " (" & Err.Number & ")" & vbNewLine
    Resume Next
DrawItem:
    'checks to see if the object is in the visible
    'area before displaying to picBay
    If blnAll Then
        objItem.Draw Output, Me
    Else
        If (objItem.P.Left >= tView.X1 And _
           objItem.P.Left + objItem.P.Width <= tView.X2) And _
           (objItem.P.Top >= tView.Y1 And _
           objItem.P.Top + objItem.P.Height <= tView.Y2) Or _
          (objItem.P.Left + objItem.P.Width >= tView.X1 And _
           objItem.P.Left <= tView.X2) And _
           (objItem.P.Top + objItem.P.Height >= tView.Y1 And _
           objItem.P.Top <= tView.Y2) Then
                objItem.Draw Output, Me
        End If
    End If
    Return
End Sub


Public Function Load(ByVal FileName As String, ByVal View As Object) As Long
    Dim F           As Long 'file stream
    'V1 to V3 headers
    Dim strHeader   As String
    Dim strFontName As String 'V2
    Dim curFontSize As Currency
    Dim lngZoomPercent As Long
    'V4 and above - binary data structures
    Dim typHeader   As FTHeaderData 'V4 to V7
    Dim typHeader8  As FTHeaderDataV8 'V8
    Dim typEntry    As FTEntryV45
    Dim typEntry6   As FTEntryV6
    Dim typEntry7   As FTEntryV7
    Dim lngByte     As Long
    Dim bytLastTwo(1 To 2) As Byte
    Dim objAdd      As FlowItem
    
    On Error GoTo Handler
    
    Set mCol = New Collection
    Me.FileName = FileName
    If InStr(1, FileName, ":") = 2 Then
        ChDrive Left$(FileName, 1)
        ChDir GetPath()
    End If
    
    F = FreeFile
    
    Open FileName For Input As #F
    Line Input #F, strHeader
    'V2 File - must process here
    
    If Left$(strHeader, conHeaderL) = conHeader Then
        Version = Val(Mid$(strHeader, conHeaderL + 1))
        If Version >= 2 And Version < 4 Then '2 to 3
            Input #F, strFontName, curFontSize, lngZoomPercent
            FontName = strFontName
            FontSize = curFontSize
            ZoomPercent = lngZoomPercent
        ElseIf Version >= 4 And Version < 8 Then 'V4 or greater, V6 or greater -- V6 header is same as V4
            Close #F
            Open FileName For Binary As #F
            
            lngByte = conHeaderL
            Do While Loc(F) < LOF(F)
                lngByte = lngByte + 1
                bytLastTwo(1) = bytLastTwo(2)
                Get #F, lngByte, bytLastTwo(2)
                If bytLastTwo(1) = 13 And bytLastTwo(2) = 10 Then
                    Seek #F, lngByte + 1
                    Exit Do
                End If
            Loop
            
            Get #F, , typHeader
            With typHeader
                FontName = .FontName
                FontSize = .FontSize
                Header1PDevName = .Header1PDevName
                Orientation = .Orientation
                PaperSize = .PaperSize
                ZoomPercent = .Zoom
                If .Version <> Version Then
                    If MsgBox("The version in the file header does not match the version in the file identifier.  Do you want to continue loading?", vbQuestion Or vbYesNo) = vbNo Then
                        Load = conFail
                        GoTo EndThis
                    End If
                    Version = .Version 'use the header version only
                End If
                If Version >= 5 Then 'V5
                    PScaleHeight = .PScaleHeight
                    PScaleWidth = .PScaleWidth
                Else
                    PScaleHeight = .PScaleHeight * conPx
                    PScaleWidth = .PScaleWidth * conPx
                End If
                'Note: V4 & V5 values differ
                '      Newer V5 is 00/10 000 (like %),
                '      older V5 and V4 are px
                ScrollX = .ScrollX
                ScrollY = .ScrollY
            End With
        ElseIf Version = 8 Then
            Close #F
            Open FileName For Binary As #F
            
            lngByte = conHeaderL
            Do While Loc(F) < LOF(F)
                lngByte = lngByte + 1
                bytLastTwo(1) = bytLastTwo(2)
                Get #F, lngByte, bytLastTwo(2)
                If bytLastTwo(1) = 13 And bytLastTwo(2) = 10 Then
                    Seek #F, lngByte + 1
                    Exit Do
                End If
            Loop
        
            Get #F, , typHeader8
            With typHeader8
                FontName = .FontName
                FontSize = .FontSize
                Header1PDevName = .DeviceName
'                Header2 = .Header2
'                Header3 = .Header3
'                Header4 = .Header4
'                Header5 = .Header5
                Orientation = .Orientation
                PaperSize = .PaperSize
                ZoomPercent = .Zoom
                If .Version <> Version Then
                    If MsgBox("The version in the file header does not match the version in the file identifier.  Do you want to continue loading?", vbQuestion Or vbYesNo) = vbNo Then
                        Load = conFail
                        GoTo EndThis
                    End If
                    Version = .Version 'use the header version only
                End If
                'version is greater than 5
                PScaleHeight = .PScaleHeight
                PScaleWidth = .PScaleWidth
                'Note: V4 & V5 values differ
                '      Newer V5 is 00/10 000 (like %),
                '      older V5 and V4 are px
                ScrollX = .ScrollX
                ScrollY = .ScrollY
                'new to version 8
                UnitScale = .UnitScale
                '.Bool1
                '.Tag1
            End With
        ElseIf Version = 1 Then
            'no extra header information provided
        Else
            Version = 0
        End If
    Else
        Version = 0 'no version on file
    End If
    
    'V1,2,3 File continuation
    If Version > 0 And Version < 4 Then 'readable
        Do While Not EOF(F)
            AddFromLine
        Loop
    ElseIf Version >= 4 And Version < 6 Then 'V4 & 5
        Do While Not Loc(F) > LOF(F)
            'ClearEntry typEntry
            Get #F, , typEntry
            If Loc(F) <= LOF(F) Then
                AddFromEntry45 typEntry
            End If
        Loop
    ElseIf Version >= 6 And Version < 7 Then 'Version 6
        Do While Not Loc(F) > LOF(F)
            'ClearEntry6 typEntry6
            Get #F, , typEntry6
            If Loc(F) <= LOF(F) Then
                AddFromEntry6 typEntry6
            End If
        Loop
    ElseIf Version >= 7 And Version < 9 Then 'Version 7 and 8
        Do While Not Loc(F) > LOF(F)
            'ClearEntry7 typEntry7
            Get #F, , typEntry7
            If Loc(F) <= LOF(F) Then
                AddFromEntry7 typEntry7
            End If
        Loop
    Else
        MsgBox "The file version of the flow chart, version " & Version & ", cannot be opened using this version (" & conCurrentVersion & ") of Flow Chart.", vbExclamation
        Load = conFail
    End If
    m_Convert.ConvToTwips Me
    
    Close #F
    
    Changed = False
    
    'refresh to a view if there is one
    Refresh View
    
    Exit Function
Handler:
    Load = Err.Number
EndThis:
    Close
End Function

Public Function LoadSettingsToPrinter() As Boolean
    Dim lngPrinter As Long
    
    'TRUE=GOOD
    
    'load settings to the printer
    If Printer.DeviceName = Header1PDevName Then
        Printer.Orientation = Orientation
        Printer.PaperSize = PaperSize
        If Printer.ScaleWidth < PScaleWidth Or Printer.ScaleHeight < PScaleHeight Then
            MsgBox "The print area has changed and is smaller than before.  Go to Printer Setup and verify settings first before printing.", vbInformation
        Else
            LoadSettingsToPrinter = True
        End If
    Else
        MsgBox "Go to Printer Setup to select the printer you wish to use.", vbInformation, "Print"
    End If
End Function

'Saves printer properties; does not change Printer properties.
Public Sub PrintFile(ByVal Range As EnumPrintRange, ByVal FromLayer As Long, ByVal ToLayer As Long, ByVal PrintOnSamePage As Boolean)
    Dim objItem     As FlowItem
    Dim lngZoom     As Long
    Dim blnBold     As Boolean
    Dim lngOrder    As Long
    Dim lngLayer    As Long
    Dim Prn         As Printer
    Dim lngL As Long, lngU As Long
    
    'NOTE: Must query layers first before using PrintFile().
    
    On Error Resume Next
    
    Set Prn = Printer 'match printer
    Refresh Prn
    
    If LoadSettingsToPrinter = False Then
        Printer.KillDoc
        Exit Sub
    End If
    
    gstrErrorMsg = "" 'clear error msg buffer
    
    'printer page adjustments
    Prn.ScaleMode = vbTwips
    Prn.CurrentX = 0
    Prn.CurrentY = 0
    
    Screen.MousePointer = vbHourglass
    
    Prn.Print 'initialize printing
    
    lngZoom = ZoomPercent
    blnBold = ForceBold
    
    ZoomPercent = 100
    ForceBold = False
    ShowSelected = False
    
    If CustomBack Then 'shaded box, not the whole page
        Prn.Line (0, 0)-(PScaleWidth, PScaleHeight), mlngBackColour, BF
    End If
    
    On Error GoTo Handler
    If Range = conPrintAll Then
        Layers.EnableAll = False
        Layers.MinMaxLayer lngL, lngU
        For lngLayer = lngL To lngU
            If Layers(lngLayer).Count > 0 Then
                Layers.EnableAll = False
                Layers.Layer(lngLayer).Enabled = True
                
                GoSub PrintEnabled
                'new page
                If Not PrintOnSamePage Then
                    Printer.NewPage
                    If CustomBack Then 'shaded box, not the whole page
                        Prn.Line (0, 0)-(PScaleWidth, PScaleHeight), mlngBackColour, BF
                    End If
                End If 'PrintOnSamePage
            End If 'Count > 0
        Next lngLayer
        
        
    ElseIf Range = conPrintRange Then
        For lngLayer = FromLayer To ToLayer
            If Layers.Layer(lngLayer).Count > 0 Then
                Layers.EnableAll = False
                Layers.Layer(lngLayer).Enabled = True
                GoSub PrintEnabled
                'new page
                If Not PrintOnSamePage Then
                    Printer.NewPage
                    If CustomBack Then 'shaded box, not the whole page
                        Prn.Line (0, 0)-(PScaleWidth, PScaleHeight), mlngBackColour, BF
                    End If
                End If
            End If
        Next lngLayer
    ElseIf Range = conPrintSelection Then
        GoSub PrintEnabled
    End If
    

    
    Prn.EndDoc 'finish doc so it'll print
    
    If Len(gstrErrorMsg) Then 'display error message
        MsgBox gstrErrorMsg, vbExclamation, "Print Log"
        gstrErrorMsg = "" 'clear message
    End If
    
    ZoomPercent = lngZoom
    ForceBold = blnBold
    Screen.MousePointer = vbDefault
    
    Exit Sub
Handler:
    If MsgBox("There was a problem while printing.  " & Err.Description & " (" & Err.Number & ").  Do you want to continue with the next item?", vbExclamation Or vbRetryCancel Or vbDefaultButton2, "Print") = vbCancel Then
        On Error Resume Next
        Prn.KillDoc 'clear the document from Windows spooling
        ZoomPercent = lngZoom
        Screen.MousePointer = vbDefault
        Exit Sub
    Else
        gstrErrorMsg = gstrErrorMsg & Err.Description & " (" & Err.Number & ")" & vbNewLine
        Resume Next
    End If
    Exit Sub
PrintEnabled:
    For lngOrder = conBottom To conTop
        For Each objItem In Me
            If objItem.P.DrawOrder = lngOrder And objItem.P.Enabled Then
                objItem.Draw Prn, Me
            End If
        Next objItem
    Next lngOrder
    Return
End Sub

Public Property Get Percentage() As Single
    Percentage = ZoomPercent / 100
End Property

Public Sub Refresh(ByVal View As Object)
    'the View is for calculations only - not for drawing
    Dim objFlowItem As FlowItem
    
    For Each objFlowItem In Me
        If objFlowItem.CanRefresh Then 'only Refresh items that can be refreshed
            objFlowItem.Refresh Me, View
        End If
    Next objFlowItem
End Sub

#If Prev = 0 Then
Public Sub Remove(ByVal Index As Variant)
    'used when removing an element from the collection
    'vntIndexKey contains either the Index or Key, which is why
    'it is declared as a Variant
    'Syntax: x.Remove(xyz)

    mCol.Remove Index
End Sub
#End If

Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    'this property allows you to enumerate
    'this collection with the For...Each syntax
    Set NewEnum = mCol.[_NewEnum]
End Property

#If Prev = 0 Then
Public Sub RemoveObj(ByVal Item As FlowItem)
    Dim objItem  As FlowItem
    Dim lngIndex As Long
    
    lngIndex = GetIndex(Item)
    If lngIndex > 0 Then
        If Item.P.GroupNo <> 0 Then 'remove this from the group too
            Groups.Remove Item
        End If
        'delete its index
        Remove lngIndex
    End If
End Sub

Public Function Save(FileName As String) As Long
    Dim F           As Long
    Dim objItem     As FlowItem
    Dim strType     As String
    Dim bytHeader() As Byte
    Dim typHeader   As FTHeaderData 'V4 to V7
    Dim typHeader8  As FTHeaderDataV8 'V8
    Dim typEntry    As FTEntryV45
    Dim typEntry6   As FTEntryV6
    Dim typEntry7   As FTEntryV7
    Dim lngItem     As Long
    
    On Error GoTo Handler
    Me.FileName = FileName
    Changed = False
    
    m_Convert.ConvFromTwips Me
    
    F = FreeFile
    
    If Version >= 2 And Version < 4 Then '2-3 V3 extra data
        Open FileName For Output As #F
        Print #F, conHeader & Format(Version, "#0.0#")
        Write #F, FontName, FontSize, ZoomPercent
    ElseIf Version < 2 Then 'less than 2
        Open FileName For Output As #F
        Print #F, conHeader & Format(Version, "#0.0#")
    ElseIf Version >= 4 And Version < 9 Then  'V4 or greater
        Open FileName For Output As #F 'purge old file
        Close #F
        Open FileName For Binary As #F
        bytHeader = StrConv(conHeader & Format(Version, "#0.0#") & vbNewLine, vbFromUnicode)
        Put #F, , bytHeader
    Else
        MsgBox "Cannot save version " & Version & " of flow chart.", vbExclamation, "Save"
        Save = 5 'invalid procedure or call
        Exit Function
    End If
    
    If Version < 4 Then
        For Each objItem In Me
            strType = GetFItemStr(objItem)
            'For text boxes, only save if text is in it
            If (TypeOf objItem Is FText) And (Len(objItem.P.Text) = 0) Then
                'Skip, removes text boxes
            Else
                Write #F, strType, objItem.P.Left, objItem.P.Top, objItem.P.Width, objItem.P.Height, _
                      objItem.P.TextAlign & "¿" & DiverseReplace(objItem.P.Text, vbBinaryCompare, """", "''", vbTab, "", vbCr, "", vbLf, ""), _
                      objItem.P.DrawOrder
            End If
        Next objItem
    ElseIf Version >= 4 And Version < 8 Then 'version 4 or greater, access with 5
        'write header
        With typHeader
            .FontName = FontName
            .FontSize = FontSize
            .Header1PDevName = Header1PDevName
            .Orientation = Orientation
            .PaperSize = PaperSize
            If Version >= 5 Then 'V5
                .PScaleHeight = PScaleHeight
                .PScaleWidth = PScaleWidth
            Else
                .PScaleHeight = PScaleHeight / conPx
                .PScaleWidth = PScaleWidth / conPx
            End If
            .ScrollX = ScrollX
            .ScrollY = ScrollY
            .Version = Version
            .Zoom = ZoomPercent
        End With
        Put #F, , typHeader
        
        If Version < 6 Then 'V4 to V5
            'write other data
            For Each objItem In Me
                If (TypeOf objItem Is FText) And (Len(objItem.P.Text) = 0) Then
                    'no text in text box, skip
                Else
                    SetToEntry45 typEntry, objItem
                    Put #F, , typEntry
                End If
            Next objItem
        ElseIf Version >= 6 And Version < 7 Then
            'write other data
            For Each objItem In Me
                If (TypeOf objItem Is FText) And (Len(objItem.P.Text) = 0) Then
                    'no text in text box, skip
                Else
                    SetToEntry6 typEntry6, objItem
                    Put #F, , typEntry6
                End If
            Next objItem
        ElseIf Version >= 7 And Version < 8 Then
            'write other data
            For Each objItem In Me
                If (TypeOf objItem Is FText) And (Len(objItem.P.Text) = 0) Then
                    'no text in text box, skip
                Else
                    SetToEntry7 typEntry7, objItem
                    Put #F, , typEntry7
                End If
            Next objItem
            
            GoSub SaveDefaultLayers
        Else
            MsgBox "This file was NOT saved because the version is improper.", vbExclamation
            Changed = True
            Save = conFail
        End If
    ElseIf Version >= 8 And Version < 9 Then 'Version 8
        'write header
        With typHeader8
            .FontName = FontName
            .FontSize = FontSize
            .DeviceName = Header1PDevName
            .Orientation = Orientation
            .PaperSize = PaperSize
            .PScaleHeight = PScaleHeight
            .PScaleWidth = PScaleWidth
            .ScrollX = ScrollX
            .ScrollY = ScrollY
            .Version = Version
            .Zoom = ZoomPercent
            'new for version 8
            .UnitScale = UnitScale
            '.Bool1
            '.Header2
            '.Tag1
        End With
        Put #F, , typHeader8
        
        'write other data
        For Each objItem In Me
            If (TypeOf objItem Is FText) And (Len(objItem.P.Text) = 0) Then
                'no text in text box, skip
            Else
                SetToEntry7 typEntry7, objItem
                Put #F, , typEntry7
            End If
        Next objItem
        
        GoSub SaveDefaultLayers
    Else
        MsgBox "Although the original file was erased, there was a MAJOR problem saving the new file.  Save this file in an older version from the Properties dialog box.", vbExclamation, "Save"
        Save = 5 'invalid procedure or call
    End If
    Close #F
    
    m_Convert.ConvToTwips Me
    
Handler:
    Save = Err.Number
    Close
    Exit Function
    
SaveDefaultLayers:
    For lngItem = 0 To Layers.MaxDefaultShape
        If Not Layers.DefaultShape(lngItem) Is Nothing Then
            SetToEntry7 typEntry7, Layers.DefaultShape(lngItem)
            'note specially that this is layer information
            'and this is associated with layer (not provided
            'in the pointer object)
            typEntry7.FlowItemNo = conAddLayerDefaults
            typEntry7.Layer = lngItem
            Put #F, , typEntry7
        End If
    Next lngItem
    Return
    
End Function

Public Sub SaveSettingsFromPrinter()
    'save settings from the printer
    Header1PDevName = Printer.DeviceName
    Orientation = Printer.Orientation
    PaperSize = Printer.PaperSize
    If Not Customized Then
        PScaleHeight = Printer.ScaleHeight
        PScaleWidth = Printer.ScaleWidth
    End If
End Sub

Private Sub SetToEntry45(Entry As FTEntryV45, ByVal FlowItem As FlowItem)   'V4, SAVE
    Dim strType As String
    
    'entry 4,5
    strType = GetFItemStr(FlowItem)
    With Entry
        .FlowItemType = strType
        .Height = FlowItem.P.Height
        .Left = FlowItem.P.Left
'        If Len(FlowItem.Tag1) = 0 Then 'new style values
            'convert new style to old style
            m_Convert.SetTextFormat FlowItem, FlowItem.P.TextBold, FlowItem.P.TextItalic, (FlowItem.P.TextSize - FontSize) / 2
            .Tag1 = FlowItem.P.Tag1 'write Tag1 to file structure
            FlowItem.P.Tag1 = "" 'clear this
'        Else
'            .Tag1 = FlowItem.Tag1 'save old style
'        End If
        .Tag2 = FlowItem.P.FontFace
        .Tag3 = FlowItem.P.Tag3
        .Text = FlowItem.P.TextAlign & "¿" & FlowItem.P.Text
        '.Text = FlowItem.P.Text
        .Top = FlowItem.P.Top
        .Width = FlowItem.P.Width
        .DrawOrder = FlowItem.P.DrawOrder
    End With
End Sub

Private Sub SetToEntry6(Entry As FTEntryV6, ByVal FlowItem As FlowItem) 'V4, SAVE
    With Entry
        .FlowItemNo = GetFItemNo(FlowItem)
        .Height = FlowItem.P.Height
        .Left = FlowItem.P.Left
        .Tag1 = FlowItem.P.Tag1
        'removed .Tag2 = FlowItem.Tag2
        .Tag3 = FlowItem.P.Tag3
        .Text = FlowItem.P.Text
        .Top = FlowItem.P.Top
        .Width = FlowItem.P.Width
        .DrawOrder = FlowItem.P.DrawOrder
        
        .BackColour = FlowItem.P.BackColour
        .FillStyle = FlowItem.P.FillStyle
        .FontFace = FlowItem.P.FontFace
        .ForeColour = FlowItem.P.ForeColour
        .LineStyle = FlowItem.P.LineStyle
        .LineWidth = FlowItem.P.LineWidth
        .Name = FlowItem.P.Name
        .TextAlign = FlowItem.P.TextAlign
        .TextBold = FlowItem.P.TextBold
        .TextItalic = FlowItem.P.TextItalic
        .TextSize = FlowItem.P.TextSize
        .TextUnderline = FlowItem.P.TextUnderline
        .TextColour = FlowItem.P.TextColour '20 items
    End With
End Sub


Private Sub SetToEntry7(Entry As FTEntryV7, ByVal FlowItem As FlowItem)  'V7, SAVE
    Dim objP As Properties
    
    Set objP = FlowItem.P
    
    With Entry
        .FlowItemNo = FlowItem.Number
        .Height = objP.Height
        .Left = objP.Left
        .Tag1 = objP.Tag1
        'removed .Tag2 = FlowItem.Tag2
        .Tag3 = objP.Tag3
        .Text = objP.Text
        .Top = objP.Top
        .Width = objP.Width
        .DrawOrder = objP.DrawOrder
        
        .ArrowEngg = objP.ArrowEngg
        .ArrowSize = objP.ArrowSize
        .BackColour = objP.BackColour
        .FillStyle = objP.FillStyle
        .FontFace = objP.FontFace
        .ForeColour = objP.ForeColour
        .GroupNo = objP.GroupNo
        .LineStyle = objP.LineStyle
        .LineWidth = objP.LineWidth
        .Name = objP.Name
        .TextAlign = objP.TextAlign
        .TextBold = objP.TextBold
        .TextItalic = objP.TextItalic
        .TextSize = objP.TextSize
        .TextUnderline = objP.TextUnderline
        .TextColour = objP.TextColour '20 items
        
        .ZBase = objP.ZBase
        .ZHeight = objP.ZHeight
        .Layer = objP.Layer
    End With
End Sub
#End If


Private Sub Class_Initialize()
    'creates the collection when this class is created
    'Error handling for problems concerning
    'a network printer (VB5) or no printer at all.
10  Set mCol = New Collection
20  Version = conCurrentVersion
30  On Error GoTo Handler
40  Printer.ScaleMode = vbTwips
50  Orientation = Printer.Orientation
60  PScaleWidth = Printer.ScaleWidth
70  PScaleHeight = Printer.ScaleHeight
80  PaperSize = Printer.PaperSize
90  FontName = conFontName
100 FontSize = conFontSize
110 ZoomPercent = 100 'default is 100%
130 Set Groups = New Groups
140 AskToSave = True
150 UnitScale = vbCentimeters
160 mlngBackColour = -1
170 Set Layers = New Layers
180 Layers.Initialize Me
190 Set m_Convert = New FlowOldFile
    Exit Sub
Handler:
    If Err Then
        If gblnPrnErrMsg = False Then   'message not shown
            MsgBox "There was an error loading printer information.  If you do not have a printer installed, install one from Control Panel.  You need a printer to use this program.", vbExclamation
            gblnPrnErrMsg = True
        End If
        PrinterError = True
        Select Case Erl
            Case 50: Orientation = 1 'portrait
            Case 60: PScaleWidth = 11520 'HP laserjet 4l twips
            Case 70: PScaleHeight = 15178
            Case 80: PaperSize = 1 ' 8" x 11"
        End Select
        Err.Clear
        Resume Next
    End If
End Sub


Private Sub Class_Terminate()
    'destroys collection when this class is terminated
    Set mCol = Nothing
End Sub


