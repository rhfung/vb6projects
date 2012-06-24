VERSION 5.00
Begin VB.UserControl EListBox 
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4770
   ForeColor       =   &H80000008&
   KeyPreview      =   -1  'True
   PropertyPages   =   "EListBox.ctx":0000
   ScaleHeight     =   240
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   318
   ToolboxBitmap   =   "EListBox.ctx":0015
   Begin VB.PictureBox picList 
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   3495
      Left            =   0
      ScaleHeight     =   229
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   309
      TabIndex        =   0
      Top             =   0
      Width           =   4695
      Begin VB.TextBox txtEdit 
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   1080
         TabIndex        =   3
         Top             =   1200
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.HScrollBar hsbScroll 
         Height          =   255
         Left            =   0
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   3240
         Visible         =   0   'False
         Width           =   4335
      End
      Begin VB.VScrollBar vsbScroll 
         Height          =   3495
         Left            =   4320
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   0
         Visible         =   0   'False
         Width           =   255
      End
   End
End
Attribute VB_Name = "EListBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "Enhanced List Box"
Attribute VB_Ext_KEY = "PropPageWizardRun" ,"Yes"
'/// Enhanced List Box
Option Explicit
Option Base 1 '1 to N
'// Developed January 2, 2003
'// Last updated May 13, 2003
'// Enhanced List Box developed by Richard H Fung
'// Copyright (C) 2003, Richard H Fung for Hai Consultants Inc.

'assumption: font must be set before text is added

'Default Property Values:
Const m_def_RowDivisions = False
Const m_def_ColumnDivisions = False
Const m_def_DivisionColor = vbButtonFace
Const m_def_SelectedBackColor = vbHighlight
Const m_def_SelectedForeColor = vbHighlightText
Const m_def_Columns = 1
Const m_def_CheckBoxes = False
Const m_def_DblClickCheck = True

'Property Variables:
Dim m_RowDivisions      As Boolean
Dim m_ColumnDivisions   As Boolean
Dim m_CheckBoxes        As Boolean
Dim m_DblClickCheck     As Boolean
Dim m_SelectedBackColor As OLE_COLOR
Dim m_SelectedForeColor As OLE_COLOR
Dim m_DivisionColor     As OLE_COLOR
Dim m_Columns As Long
Dim m_Column  As ESingleArray

'run time variables

Private Type EListData
    strText        As String
    blnSelected    As Boolean
    lngItemData    As Long
    strTextSub()   As String
End Type

'edit box
Private m_lngColX()   As Long
Attribute m_lngColX.VB_VarDescription = "Occurs when an item is selected by keyboard or mouse.  One-based array."
Private m_lngEditRow  As Long
Attribute m_lngEditRow.VB_VarDescription = "Occurs when an item is selected by keyboard or mouse.  One-based array."
Private m_lngEditCol  As Long
Attribute m_lngEditCol.VB_VarDescription = "Occurs when an item is selected by keyboard or mouse.  One-based array."

Private m_ListData() As EListData
Attribute m_ListData.VB_VarDescription = "Occurs when an item is selected by keyboard or mouse.  One-based array."
Private m_ListCount  As Long
Attribute m_ListCount.VB_VarDescription = "Occurs when an item is selected by keyboard or mouse.  One-based array."
Private m_ListUBound As Long
Attribute m_ListUBound.VB_VarDescription = "Occurs when an item is selected by keyboard or mouse.  One-based array."
Private m_ListIndex  As Long
Attribute m_ListIndex.VB_VarDescription = "Occurs when an item is selected by keyboard or mouse.  One-based array."
Private m_blnFocus   As Boolean
Attribute m_blnFocus.VB_VarDescription = "Occurs when an item is selected by keyboard or mouse.  One-based array."
Private m_lngTextWidth() As Long 'width of string
Attribute m_lngTextWidth.VB_VarDescription = "Occurs when an item is selected by keyboard or mouse.  One-based array."

Private m_strSearch     As String
Attribute m_strSearch.VB_VarDescription = "Occurs when an item is selected by keyboard or mouse.  One-based array."
Private m_blnClickOnWhiteSpace As Boolean
Attribute m_blnClickOnWhiteSpace.VB_VarDescription = "Occurs when an item is selected by keyboard or mouse.  One-based array."
Private m_blnFullRectShown As Boolean
Attribute m_blnFullRectShown.VB_VarDescription = "Occurs when an item is selected by keyboard or mouse.  One-based array."

'Event Declarations:
Event Click(ByVal ListIndex As Long)  'MappingInfo=UserControl,UserControl,-1,Click
Attribute Click.VB_Description = "Occurs when an item is selected by keyboard or mouse.  One-based array."
Attribute Click.VB_UserMemId = -600
Attribute Click.VB_MemberFlags = "200"
'/// Occurs when an item is selected by keyboard or mouse.  One-based array.
Event DblClick() 'MappingInfo=UserControl,UserControl,-1,DblClick
Attribute DblClick.VB_Description = "Occurs when the user presses and releases a mouse button and then presses and releases it again over an object."
Attribute DblClick.VB_UserMemId = -601
'/// Occurs when the user presses and releases a mouse button and then presses and releases it again over an object.
Event KeyDown(KeyCode As Integer, Shift As Integer) 'MappingInfo=UserControl,UserControl,-1,KeyDown
Attribute KeyDown.VB_Description = "Occurs when the user presses a key while an object has the focus."
Attribute KeyDown.VB_UserMemId = -602
'/// Occurs when the user presses a key while an object has the focus.
Event KeyPress(KeyAscii As Integer) 'MappingInfo=UserControl,UserControl,-1,KeyPress
Attribute KeyPress.VB_Description = "Occurs when the user presses and releases an ANSI key."
Attribute KeyPress.VB_UserMemId = -603
'/// Occurs when the user presses and releases an ANSI key.
Event KeyUp(KeyCode As Integer, Shift As Integer) 'MappingInfo=UserControl,UserControl,-1,KeyUp
Attribute KeyUp.VB_Description = "Occurs when the user releases a key while an object has the focus."
Attribute KeyUp.VB_UserMemId = -604
'/// Occurs when the user releases a key while an object has the focus.
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseDown
Attribute MouseDown.VB_Description = "Occurs when the user presses the mouse button while an object has the focus."
Attribute MouseDown.VB_UserMemId = -605
'/// Occurs when the user presses the mouse button while an object has the focus.
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseMove
Attribute MouseMove.VB_Description = "Occurs when the user moves the mouse."
Attribute MouseMove.VB_UserMemId = -606
'/// Occurs when the user moves the mouse.
Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseUp
Attribute MouseUp.VB_Description = "Occurs when the user releases the mouse button while an object has the focus."
Attribute MouseUp.VB_UserMemId = -607
'/// Occurs when the user releases the mouse button while an object has the focus.
Event ItemChecked(ByVal ListIndex As Long, ByVal Checked As Boolean, ByRef Cancel As Boolean)
Attribute ItemChecked.VB_Description = "Occurs when an item in the list box is checked and CheckBoxes variable is True."
'/// Occurs when an item in the list box is checked and CheckBoxes variable is True.
Event ValidateEdit(ByVal Row As Long, ByVal Col As Long, ByVal EditText As String, ByRef Cancel As Boolean)
Attribute ValidateEdit.VB_Description = "Occurs when the edit text box hides."
'/// Occurs when the edit text box hides.

'Windows API:
Private Declare Function DrawFocusRect Lib "user32" (ByVal hdc As Long, lpRect As RECT) As Long
Private Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hdc As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long

Private Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type

Private Const DT_BOTTOM = &H8
Attribute DT_BOTTOM.VB_VarDescription = "Adds an item to the list."
Private Const DT_CALCRECT = &H400
Attribute DT_CALCRECT.VB_VarDescription = "Adds an item to the list."
Private Const DT_CENTER = &H1
Attribute DT_CENTER.VB_VarDescription = "Adds an item to the list."
Private Const DT_CHARSTREAM = 4
Attribute DT_CHARSTREAM.VB_VarDescription = "Adds an item to the list."
Private Const DT_DISPFILE = 6
Attribute DT_DISPFILE.VB_VarDescription = "Adds an item to the list."
Private Const DT_EXPANDTABS = &H40
Attribute DT_EXPANDTABS.VB_VarDescription = "Adds an item to the list."
Private Const DT_EXTERNALLEADING = &H200
Attribute DT_EXTERNALLEADING.VB_VarDescription = "Adds an item to the list."
Private Const DT_INTERNAL = &H1000
Attribute DT_INTERNAL.VB_VarDescription = "Adds an item to the list."
Private Const DT_LEFT = &H0
Attribute DT_LEFT.VB_VarDescription = "Adds an item to the list."
Private Const DT_METAFILE = 5
Attribute DT_METAFILE.VB_VarDescription = "Adds an item to the list."
Private Const DT_NOCLIP = &H100
Attribute DT_NOCLIP.VB_VarDescription = "Adds an item to the list."
Private Const DT_NOPREFIX = &H800
Attribute DT_NOPREFIX.VB_VarDescription = "Adds an item to the list."
Private Const DT_PLOTTER = 0
Attribute DT_PLOTTER.VB_VarDescription = "Adds an item to the list."
Private Const DT_RASCAMERA = 3
Attribute DT_RASCAMERA.VB_VarDescription = "Adds an item to the list."
Private Const DT_RASDISPLAY = 1
Attribute DT_RASDISPLAY.VB_VarDescription = "Adds an item to the list."
Private Const DT_RASPRINTER = 2
Attribute DT_RASPRINTER.VB_VarDescription = "Adds an item to the list."
Private Const DT_RIGHT = &H2
Attribute DT_RIGHT.VB_VarDescription = "Adds an item to the list."
Private Const DT_SINGLELINE = &H20
Attribute DT_SINGLELINE.VB_VarDescription = "Adds an item to the list."
Private Const DT_TABSTOP = &H80
Attribute DT_TABSTOP.VB_VarDescription = "Adds an item to the list."
Private Const DT_TOP = &H0
Attribute DT_TOP.VB_VarDescription = "Adds an item to the list."
Private Const DT_VCENTER = &H4
Attribute DT_VCENTER.VB_VarDescription = "Adds an item to the list."
Private Const DT_WORDBREAK = &H10
Attribute DT_WORDBREAK.VB_VarDescription = "Adds an item to the list."

Public Sub AddItem(ByVal Item As String)
Attribute AddItem.VB_Description = "Adds an item to the list."
'/// Adds an item to the list.
    m_ListCount = m_ListCount + 1
    If m_ListCount > m_ListUBound Then
        m_ListUBound = m_ListUBound + 10
        ReDim Preserve m_ListData(m_ListUBound)
    End If
    
    m_ListData(m_ListCount).strText = Item
    If m_Columns > 1 Then ReDim m_ListData(m_ListCount).strTextSub(m_Columns - 1)
    
    'set max column width
    If picList.TextWidth(Item) > m_lngTextWidth(1) Then
        m_lngTextWidth(1) = picList.TextWidth(Item)
    End If
    
    UpdateList
End Sub

Public Sub AddItem2(ByVal Item As String, ByVal ItemData As Integer, ParamArray args() As Variant)
Attribute AddItem2.VB_Description = "Adds an item, its item data, and its associated column strings to the list."
'/// Adds an item, its item data, and its associated column strings to the list.
    m_ListCount = m_ListCount + 1
    If m_ListCount > m_ListUBound Then
        m_ListUBound = m_ListUBound + 10
        ReDim Preserve m_ListData(m_ListUBound)
    End If
    
    With m_ListData(m_ListCount)
        .strText = Item
        'set max column width
        If picList.TextWidth(Item) > m_lngTextWidth(1) Then
            m_lngTextWidth(1) = picList.TextWidth(Item)
        End If
        
        .lngItemData = ItemData
        
        If m_Columns > 1 Then
            ReDim .strTextSub(m_Columns - 1)
            
            Dim lngItem     As Long
            Dim lngIndex    As Long
            
            
            For lngItem = LBound(args) To UBound(args)
                lngIndex = lngIndex + 1 'increment index counter
                If lngIndex > m_Columns - 1 Then Exit For 'read until filled up space - trash remaining column data
                'fill in text data
                .strTextSub(lngIndex) = args(lngItem)
                'set max column width
                If picList.TextWidth(.strTextSub(lngIndex)) > m_lngTextWidth(lngIndex + 1) Then
                    m_lngTextWidth(lngIndex + 1) = picList.TextWidth(.strTextSub(lngIndex))
                End If
            Next lngItem
        End If
    End With
    
    UpdateList
End Sub



'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,BackColor
Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Returns/sets the background color used to display text and graphics in an object."
Attribute BackColor.VB_ProcData.VB_Invoke_Property = ";Appearance"
Attribute BackColor.VB_UserMemId = -501
'/// Returns/sets the background color used to display text and graphics in an object.
    BackColor = picList.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    picList.BackColor() = New_BackColor
    PropertyChanged "BackColor"
End Property

Public Property Let CheckBoxes(ByVal pCheck As Boolean)
Attribute CheckBoxes.VB_Description = "Determines if check boxes are shown in the list box."
Attribute CheckBoxes.VB_ProcData.VB_Invoke_PropertyPut = "Grid"
    m_CheckBoxes = pCheck
    picList.Refresh
    PropertyChanged "CheckBoxes"
End Property

Public Property Get CheckBoxes() As Boolean
'/// Determines if check boxes are shown in the list box.
    CheckBoxes = m_CheckBoxes
End Property


Public Sub Clear()
Attribute Clear.VB_Description = "Clears all elements from the list."
'/// Clears all elements from the list.
    ReDim m_ListData(1)
    m_ListCount = 0
    m_ListUBound = 1
    m_strSearch = ""
    
    UpdateList
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,0
Public Property Get ColumnDivisions() As Boolean
Attribute ColumnDivisions.VB_Description = "Shows/hides the column division line."
'/// Shows/hides the column division line.
    ColumnDivisions = m_ColumnDivisions
End Property

Public Property Let ColumnDivisions(ByVal New_ColumnDivisions As Boolean)
    m_ColumnDivisions = New_ColumnDivisions
    PropertyChanged "ColumnDivisions"
End Property

Public Property Get DivisionColor() As OLE_COLOR
Attribute DivisionColor.VB_Description = "Gets/sets the colour of the division lines."
'/// Gets/sets the colour of the division lines.
    DivisionColor = m_DivisionColor
End Property

Public Property Let DivisionColor(ByVal vNewValue As OLE_COLOR)
    m_DivisionColor = vNewValue
    PropertyChanged "DivisionColor"
End Property

Public Property Get DblClickCheck() As Boolean
Attribute DblClickCheck.VB_Description = "When double-click on an item, is the item checked?"
Attribute DblClickCheck.VB_ProcData.VB_Invoke_Property = "Grid"
'/// When double-click on an item, is the item checked?
    DblClickCheck = m_DblClickCheck
End Property

Public Property Let DblClickCheck(ByVal pVal As Boolean)
    m_DblClickCheck = pVal
    PropertyChanged "DblClickCheck"
End Property

Private Sub DrawFocus(ByVal Visible As Boolean)
    Static tRect As RECT
            
    If Visible <> m_blnFullRectShown Then
        tRect.Right = picList.ScaleWidth - IIf(vsbScroll.Visible, vsbScroll.Width, 0)
        tRect.Bottom = picList.ScaleHeight - IIf(hsbScroll.Visible, vsbScroll.Height, 0)
        DrawFocusRect picList.hdc, tRect
        m_blnFullRectShown = Not m_blnFullRectShown
    End If
End Sub

Private Sub DrawRows(ByVal FromIndex As Long, ByVal ToIndex As Long, ByVal BackFill As Boolean)
    Dim lngItem     As Long
    Dim lngSub      As Long
    Dim lngHeight   As Long
    Dim lngWidth    As Long
    Dim lngCheck    As Long 'check box adjustment
    Dim lngWidthCheck As Long
    Dim lngLastY    As Long
    Dim tRect       As RECT
    Static lngStack As Long 'protect from recursion
    
    Const conSpacer = 10
    
    On Error Resume Next
    
    If lngStack > 2 Then Exit Sub
    
    'absolute column positions
    lngHeight = GetRowHeight()
    lngWidth = picList.ScaleWidth - IIf(vsbScroll.Visible, vsbScroll.Width, 0)
    If m_CheckBoxes Then
        lngCheck = lngHeight - hsbScroll
    End If
    lngWidthCheck = IIf(m_CheckBoxes, lngHeight, 0)
    
    ReDim m_lngColX(m_Columns)
    
    If m_Column(1) = 0 Then
        m_lngColX(1) = m_lngTextWidth(1) + lngWidthCheck + conSpacer
    Else
        m_lngColX(1) = m_Column(1) + lngWidthCheck + conSpacer
    End If

    For lngSub = 2 To m_Columns
        If m_Column(lngSub) <= -1 Then
            m_lngColX(lngSub) = m_lngColX(lngSub - 1)
        ElseIf m_Column(lngSub) = 0 Then
            m_lngColX(lngSub) = m_lngColX(lngSub - 1) + m_lngTextWidth(lngSub) + conSpacer
        Else
            m_lngColX(lngSub) = m_lngColX(lngSub - 1) + m_Column(lngSub) + conSpacer
        End If
    Next lngSub

    'horizontal scroll bar adjustments
    If m_lngColX(m_Columns) - picList.ScaleWidth > 0 Then
        hsbScroll.Max = m_lngColX(m_Columns) - picList.ScaleWidth + IIf(vsbScroll.Visible, vsbScroll.Width, 0)
        hsbScroll.LargeChange = picList.ScaleWidth - IIf(vsbScroll.Visible, vsbScroll.Width, 0)
        For lngSub = 1 To m_Columns
            m_lngColX(lngSub) = m_lngColX(lngSub) - hsbScroll
        Next lngSub
        If hsbScroll.Visible = False Then
            lngStack = lngStack + 1
            hsbScroll.Visible = True
            UpdateList
            DrawRows FromIndex, ToIndex, BackFill
            lngStack = lngStack - 1
            Exit Sub
        End If
    Else
        If hsbScroll.Visible = True Then
            lngStack = lngStack + 1
            hsbScroll.Visible = False
            hsbScroll = 0
            UpdateList
            DrawRows FromIndex, ToIndex, BackFill
            lngStack = lngStack - 1
            Exit Sub
        End If
    End If
    
    If FromIndex < 1 Then FromIndex = 1 'first element
    If ToIndex > ListCount Then ToIndex = ListCount  'last element
    
    
    'drawing the rows
    For lngItem = FromIndex To ToIndex
        lngLastY = (lngItem - vsbScroll) * lngHeight
        tRect.Left = lngCheck
        tRect.Top = lngLastY + 1
        tRect.Right = lngWidth
        tRect.Bottom = tRect.Top + lngHeight - 1
        
        'display highlight
        If lngItem = m_ListIndex Then
            picList.ForeColor = IIf(picList.Enabled, SelectedBackColor, vbInactiveTitleBar)
            If BackFill Then
                'white background
                picList.Line (0, lngLastY)-Step(lngWidth, lngHeight - 1), picList.BackColor, BF
            End If
            picList.Line (1 + lngCheck, lngLastY + 1)-Step(lngWidth - lngCheck - 1, lngHeight - 2), SelectedBackColor, BF
            'picList.Line (1 + lngCheck, lngLastY + 2)-Step(lngWidth - lngCheck - 1, lngHeight - 4), vbGreen, BF 'SelectedBackColor, BF
            If m_blnFocus Then
                DrawFocusRect picList.hdc, tRect
            End If
            picList.ForeColor = IIf(picList.Enabled, SelectedForeColor, vbGrayText)
            picList.Print "";
        Else
            picList.ForeColor = IIf(picList.Enabled, UserControl.ForeColor, vbGrayText)
            If BackFill Then
                'white background
                picList.Line (0, lngLastY)-Step(lngWidth, lngHeight - 1), picList.BackColor, BF
            End If
        End If
        
        If m_CheckBoxes Then
            'display check box
            picList.Line (2 - hsbScroll, lngLastY + 2)-Step(lngHeight - 4, lngHeight - 4), UserControl.ForeColor, B
            'display X
            If m_ListData(lngItem).blnSelected Then
                picList.Line (2 - hsbScroll, lngLastY + 2)-Step(lngHeight - 4, lngHeight - 4), UserControl.ForeColor
                picList.Line (2 - hsbScroll, lngLastY + lngHeight - 2)-Step(lngHeight - 4, -lngHeight + 4), UserControl.ForeColor
            End If
        End If
'        picList.CurrentY = lngLastY
'        picList.CurrentX = ln
        
        'display first column
        tRect.Left = lngWidthCheck + conSpacer \ 2 - IIf(hsbScroll.Visible, hsbScroll, 0)
        tRect.Top = lngLastY + 1
        tRect.Right = m_lngColX(1)
        tRect.Bottom = tRect.Top + lngHeight - 1
        
        DrawText picList.hdc, m_ListData(lngItem).strText, -1, tRect, DT_SINGLELINE Or DT_TOP
        
        'display other columns
        For lngSub = 1 To m_Columns - 1
            If m_Column(lngSub + 1) >= 0 Then
                tRect.Left = m_lngColX(lngSub) + conSpacer \ 2
                tRect.Right = m_lngColX(lngSub + 1)
                DrawText picList.hdc, m_ListData(lngItem).strTextSub(lngSub), -1, tRect, DT_SINGLELINE Or DT_TOP
            End If
        Next lngSub
        
        'display division line top
        If m_RowDivisions Then
            picList.Line (0, lngLastY)-Step(lngWidth, 0), m_DivisionColor
        End If
        'move down
        picList.CurrentY = picList.CurrentY + lngHeight
        picList.CurrentX = 0
        
        'division line bottom
        If m_RowDivisions Then
            picList.Line -Step(lngWidth, 0), DivisionColor
        End If
    Next lngItem
    
    'draw columns
    If m_ColumnDivisions Then
        For lngSub = 1 To m_Columns - 1
            picList.Line (m_lngColX(lngSub), 0)-Step(0, picList.ScaleHeight), m_DivisionColor
        Next lngSub
    End If
    
    'move text box
    If m_lngEditCol > 1 Then
        txtEdit.Left = m_lngColX(m_lngEditCol - 1) + conSpacer / 2
    Else
        txtEdit.Left = lngWidthCheck + conSpacer \ 2 - IIf(hsbScroll.Visible, hsbScroll, 0)
    End If
    If m_lngEditCol > 0 And txtEdit.Visible Then
        txtEdit.Width = m_lngColX(m_lngEditCol) - txtEdit.Left - conSpacer / 2
    End If
    
End Sub

Public Sub EditBoxOff(Optional ByVal Cancel As Boolean)
Attribute EditBoxOff.VB_Description = "Hides the edit box.  If Cancel = False, the data is saved."
'/// Hides the edit box.  If Cancel = False, the data is saved.
    If Not txtEdit.Visible Then Exit Sub

    RaiseEvent ValidateEdit(m_lngEditRow, m_lngEditCol, txtEdit.Text, Cancel)

    If Not Cancel Then
        ListSub(m_lngEditRow, m_lngEditCol) = txtEdit.Text
    End If
    
    txtEdit.Visible = False
    txtEdit.Text = vbNullString
End Sub

Public Sub EditBoxOn(ByVal Row As Long, ByVal Col As Long)
Attribute EditBoxOn.VB_Description = "Shows the edit box."
'/// Shows the edit box.
    Dim lngHeight   As Long
    Dim lngWidthCheck As Long
    
    If txtEdit.Visible Then
        EditBoxOff
    End If
    
    m_lngEditRow = Row
    m_lngEditCol = Col
    
    lngHeight = GetRowHeight
    
    'lngWidthCheck + conSpacer \ 2 - IIf(hsbScroll.Visible, hsbScroll, 0)
    lngWidthCheck = IIf(m_CheckBoxes, lngHeight, 0)
    
    Const conSpacer = 10
    
    txtEdit.Top = lngHeight * (Row - TopIndex) + 1
    txtEdit.Height = lngHeight
    If Col > 1 Then
        txtEdit.Left = m_lngColX(Col - 1) + conSpacer / 2
    Else
        txtEdit.Left = lngWidthCheck + conSpacer \ 2 - IIf(hsbScroll.Visible, hsbScroll, 0)
    End If
    
    txtEdit.Width = m_lngColX(Col) - txtEdit.Left - conSpacer / 2
    If txtEdit.Width < 10 Then 'fill whole space
        If picList.ScaleWidth - vsbScroll.Width - txtEdit.Left > 0 Then
            txtEdit.Width = picList.ScaleWidth - vsbScroll.Width - txtEdit.Left
        Else
            txtEdit.Left = 0
            txtEdit.Width = picList.ScaleWidth - vsbScroll.Width
        End If
    End If
    
    txtEdit.Text = ListSub(Row, Col)
    txtEdit.SelStart = 0
    txtEdit.SelLength = Len(txtEdit.Text)
    txtEdit.ToolTipText = "Row:" & Row & " Col:" & Col
    
    txtEdit.Visible = True
    On Error Resume Next
    txtEdit.SetFocus
End Sub

Public Property Get EditText() As String
Attribute EditText.VB_Description = "Gets/sets the text in the edit box, called by EditBoxOn() method."
Attribute EditText.VB_MemberFlags = "400"
'/// Gets/sets the text in the edit box, called by EditBoxOn() method.
    EditText = txtEdit.Text
End Property


Public Property Let EditText(ByVal pEditText As String)
    txtEdit.Text = pEditText
End Property


'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,ForeColor
Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_Description = "Returns/sets the foreground color used to display text and graphics in an object."
Attribute ForeColor.VB_ProcData.VB_Invoke_Property = ";Appearance"
Attribute ForeColor.VB_UserMemId = -513
'/// Returns/sets the foreground color used to display text and graphics in an object.
    ForeColor = UserControl.ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
    UserControl.ForeColor() = New_ForeColor
    PropertyChanged "ForeColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Enabled
Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Returns/sets a value that determines whether an object can respond to user-generated events."
Attribute Enabled.VB_ProcData.VB_Invoke_Property = ";Behavior"
Attribute Enabled.VB_UserMemId = -514
'/// Returns/sets a value that determines whether an object can respond to user-generated events.
    Enabled = picList.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    picList.Enabled() = New_Enabled
    vsbScroll.Enabled = New_Enabled
    PropertyChanged "Enabled"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Font
Public Property Get Font() As Font
Attribute Font.VB_Description = "Returns a Font object."
Attribute Font.VB_ProcData.VB_Invoke_Property = "StandardFont;Appearance"
Attribute Font.VB_UserMemId = -512
'/// Returns a Font object.
    Set Font = picList.Font
End Property

Public Property Set Font(ByVal New_Font As Font)
    Set picList.Font = New_Font
    Set txtEdit.Font = New_Font
    Set UserControl.Font = New_Font
    PropertyChanged "Font"
End Property

Public Function GetRowHeight() As Long
Attribute GetRowHeight.VB_Description = "Height of a row is returned in pixels."
'/// Height of a row is returned in pixels.
    GetRowHeight = picList.TextHeight("O") * 1.1
End Function

Public Function GetItemsVisible() As Integer
Attribute GetItemsVisible.VB_Description = "Number of items visible in the list box."
'/// Number of items visible in the list box.
    GetItemsVisible = (picList.ScaleHeight - IIf(hsbScroll.Visible, hsbScroll.Height, 0)) / GetRowHeight()
    '+ 1 - IIf(hsbScroll, TextHeight("U") / hsbScroll.Height, 1)
End Function

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=13,0,2,
Public Property Get List(ByVal Index As Long) As String
Attribute List.VB_Description = "One-based array list of strings."
Attribute List.VB_ProcData.VB_Invoke_Property = ";List"
Attribute List.VB_MemberFlags = "400"
'/// One-based array list of strings.
    List = m_ListData(Index).strText
End Property

Public Property Let List(ByVal Index As Long, ByVal New_Text As String)
    If Ambient.UserMode = False Then Err.Raise 387
    m_ListData(Index).strText = New_Text
    
    'update column size if needs be
    If m_Column(1) = 0 Then
        If picList.TextWidth(New_Text) > m_lngTextWidth(1) Then
            m_lngTextWidth(1) = picList.TextWidth(New_Text)
        End If
    End If
    
    
    picList.Refresh
    PropertyChanged "List"
End Property

Public Property Get ListCount() As Long
'/// Returns count of items in the list.  One-based array
    ListCount = m_ListCount
End Property

Public Property Let ListCount(ByVal pNewValue As Long)
Attribute ListCount.VB_Description = "Returns count of items in the list.  One-based array"
Attribute ListCount.VB_ProcData.VB_Invoke_PropertyPut = ";List"
Attribute ListCount.VB_MemberFlags = "400"
'/// Returns count of items in the list.  One-based array
    Dim lngItem As Long
    
    If Ambient.UserMode = False Then Err.Raise 387
    
    If pNewValue > 0 Then
        ReDim Preserve m_ListData(pNewValue)
        For lngItem = m_ListCount + 1 To pNewValue
            ReDim m_ListData(lngItem).strTextSub(m_Columns)
        Next lngItem
        
        m_ListUBound = pNewValue
        m_ListCount = pNewValue
        UpdateList
        picList.Refresh
    Else
        Err.Raise 9
    End If
    PropertyChanged "ListCount"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=13,0,2,
Public Property Get ListSub(ByVal Row As Long, ByVal Col As Long) As String
Attribute ListSub.VB_Description = "One-based array.  Column 1 is the same as List(n)."
Attribute ListSub.VB_ProcData.VB_Invoke_Property = ";List"
Attribute ListSub.VB_MemberFlags = "400"
'/// One-based array.  Column 1 is the same as List(n).
    If Col > 1 Then
        ListSub = m_ListData(Row).strTextSub(Col - 1)
    ElseIf Col = 1 Or Col = 0 Then
        ListSub = m_ListData(Row).strText
    Else
        Err.Raise 9
    End If
End Property

Public Property Let ListSub(ByVal Row As Long, ByVal Col As Long, ByVal New_TextSub As String)
    If Ambient.UserMode = False Then Err.Raise 387
    If Col > 1 Then
        m_ListData(Row).strTextSub(Col - 1) = New_TextSub
        
        'update column size if needs be
        If m_Column(Col) = 0 Then
            If picList.TextWidth(New_TextSub) > m_lngTextWidth(Col) Then
                m_lngTextWidth(Col) = picList.TextWidth(New_TextSub)
            End If
        End If
        
        picList.Refresh
    ElseIf Col = 1 Or Col = 0 Then
        m_ListData(Row).strText = New_TextSub
        'update column size if needs be
        If m_Column(1) = 0 Then
            If picList.TextWidth(New_TextSub) > m_lngTextWidth(1) Then
                m_lngTextWidth(1) = picList.TextWidth(New_TextSub)
            End If
        End If
        picList.Refresh
    Else
        Err.Raise 9
    End If
    PropertyChanged "ListSub"
End Property


Public Property Let ListIndex(ByVal New_Index As Long)
    If 1 <= New_Index And New_Index <= m_ListCount Then
        
        If Not (vsbScroll <= New_Index And New_Index <= CLng(vsbScroll) + CLng(GetItemsVisible()) - CLng(IIf(hsbScroll.Visible, 3, 1))) Then
            Dim lngTop As Long
            
            'determine if scroll up or down
            If New_Index > m_ListIndex Then
                lngTop = New_Index - GetItemsVisible() + IIf(hsbScroll.Visible, 3, 1)
            ElseIf New_Index < m_ListIndex Then
                lngTop = New_Index
            Else
                lngTop = New_Index
            End If
            If lngTop > vsbScroll.Max Then lngTop = vsbScroll.Max
            If lngTop < 1 Then lngTop = 1
            m_ListIndex = New_Index 'change selected item
            vsbScroll = lngTop
        Else
            'm_ListIndex = New_Index 'change selected item
            'picList.Refresh
            MoveSelection New_Index
        End If
        
        RaiseEvent Click(New_Index)
    Else
        Err.Raise 9 'subscript out of range
    End If
End Property

Public Property Get ListIndex() As Long
Attribute ListIndex.VB_Description = "One-based array.  The ListIndex is between 1 and ListCount."
Attribute ListIndex.VB_ProcData.VB_Invoke_Property = ";List"
Attribute ListIndex.VB_MemberFlags = "400"
'/// One-based array.  The ListIndex is between 1 and ListCount.
    ListIndex = m_ListIndex
End Property


Private Sub MoveSelection(ByVal New_Index As Long)
'draws the new selection position
    Dim lngOld As Long
    
    lngOld = m_ListIndex
    m_ListIndex = New_Index
    If 0 < lngOld And lngOld <= m_ListCount Then
        DrawRows lngOld, lngOld, True
    End If
    If 0 < New_Index And New_Index <= m_ListCount Then
        DrawRows New_Index, New_Index, True
    End If
    
    If New_Index <> m_lngEditRow And txtEdit.Visible Then EditBoxOff
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Refresh
Public Sub Refresh()
Attribute Refresh.VB_Description = "Forces a complete repaint of a object."
Attribute Refresh.VB_UserMemId = -550
'/// Forces a complete repaint of a object.
    picList.Refresh
End Sub

Public Sub RemoveItem(ByVal Index As Long)
Attribute RemoveItem.VB_Description = "One-based array.  Once an item is removed, all items thereafter are re-indexed."
'/// One-based array.  Once an item is removed, all items thereafter are re-indexed.
    Dim lngAfter  As Long
    
    For lngAfter = Index To m_ListCount - 1
        m_ListData(lngAfter) = m_ListData(lngAfter + 1)
    Next lngAfter
    
    m_ListCount = m_ListCount - 1
    If m_ListIndex > m_ListCount Then
        m_ListIndex = m_ListCount
    End If
    
    m_strSearch = ""
    
    UpdateList
    picList.Refresh
End Sub

Public Sub ShowAbout()
Attribute ShowAbout.VB_Description = "Shows the About dialog box."
Attribute ShowAbout.VB_UserMemId = -552
Attribute ShowAbout.VB_MemberFlags = "40"
'/// Shows the About dialog box.
    Load frmAbout
    frmAbout.lblDescription = "EListBox Thread:" & App.ThreadID & " " & App.LegalCopyright
    If txtEdit.Visible Then
        frmAbout.lblDescription = frmAbout.lblDescription & vbNewLine & "Currently Editing: " & txtEdit.ToolTipText
    End If
    frmAbout.Show vbModal
End Sub

Public Property Get Text() As String
Attribute Text.VB_Description = "[Get] Text relates to the item at ListIndex. [Let] Searches each list item's text for the given value, or returns error if failed."
Attribute Text.VB_ProcData.VB_Invoke_Property = ";List"
Attribute Text.VB_UserMemId = 0
Attribute Text.VB_MemberFlags = "400"
'/// [Get] Text relates to the item at ListIndex. [Let] Searches each list item's text for the given value, or returns error if failed.
    If m_ListIndex > 0 Then
        Text = m_ListData(m_ListIndex).strText
    End If
End Property

Public Property Let Text(ByVal pStr As String)
    Dim lngItem As Long
    
    If Ambient.UserMode = False Then Err.Raise 387
    
    m_strSearch = ""
    
    For lngItem = 1 To m_ListCount
        If StrComp(pStr, m_ListData(lngItem).strText, vbTextCompare) = 0 Then
            ListIndex = lngItem
            Exit Property
        End If
    Next lngItem
    
    Err.Raise 5
    
'Changes text of the selected item - not the normal behaviour.
'    If m_ListIndex > 0 Then
'        m_ListData(m_ListIndex).strText = str
'    Else
'        Err.Raise 9
'    End If

    PropertyChanged "Text"
End Property

Private Sub UpdateList()
    Dim lngMax As Long
    
    lngMax = m_ListCount - GetItemsVisible() / 2
    
    If lngMax <= 1 Then lngMax = 1

    vsbScroll.Min = 1
    If vsbScroll > lngMax Then
        vsbScroll = lngMax
    End If
    If lngMax <= 32767 Then
        vsbScroll.Max = lngMax
    Else
        vsbScroll.Max = 32767
    End If
    
    vsbScroll.SmallChange = 1
    vsbScroll.LargeChange = GetItemsVisible()
    vsbScroll.Visible = vsbScroll.Min <> vsbScroll.Max
    hsbScroll.Move 0, picList.ScaleHeight - hsbScroll.Height, picList.ScaleWidth - IIf(vsbScroll.Visible, vsbScroll.Width, 0)
End Sub

Private Sub HsbScroll_Change()
    picList.Refresh
End Sub

Private Sub HsbScroll_Scroll()
    picList.Refresh
End Sub


Private Sub picList_Click()
    m_strSearch = ""
    If m_ListIndex > 0 And Not m_blnClickOnWhiteSpace Then RaiseEvent Click(m_ListIndex)
End Sub

Private Sub picList_DblClick()
    Dim lngItem As Long
    Dim blnCancel As Boolean
    
    If m_ListIndex <= 0 Or m_blnClickOnWhiteSpace Then Exit Sub
    lngItem = m_ListIndex
    'see picList_MouseUp
    If 0 < lngItem And lngItem <= m_ListCount Then
        'check box region
        If m_CheckBoxes And m_DblClickCheck Then
            RaiseEvent ItemChecked(lngItem, Not m_ListData(lngItem).blnSelected, blnCancel)
            If Not blnCancel Then
                m_ListData(lngItem).blnSelected = Not m_ListData(lngItem).blnSelected
            End If
        End If
        MoveSelection lngItem
    End If
    
    RaiseEvent DblClick
End Sub

Private Sub picList_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeySpace And m_CheckBoxes And m_ListIndex > 0 Then
        m_ListData(m_ListIndex).blnSelected = Not m_ListData(m_ListIndex).blnSelected
        MoveSelection m_ListIndex
    ElseIf KeyAscii >= vbKeySpace Then
        Dim i As Long, L As Long
        
        m_strSearch = m_strSearch & Chr$(KeyAscii)
        L = Len(m_strSearch)
        
        For i = 1 To m_ListCount
            If StrComp(m_strSearch, Left$(m_ListData(i).strText, L), vbTextCompare) = 0 Then
                ListIndex = i
                Exit For
            End If
        Next i
    Else
        m_strSearch = ""
    End If
    RaiseEvent KeyPress(KeyAscii)
End Sub

Private Sub picList_Paint()
    picList.Cls
    m_blnFullRectShown = False

    If m_ListCount Then
        Dim lngTo As Long
        
        lngTo = CLng(vsbScroll) + CLng(GetItemsVisible())
        If lngTo > m_ListCount Then lngTo = m_ListCount
        DrawRows vsbScroll, lngTo, False
    Else
        If Ambient.UserMode = False Then
            picList.Print Ambient.DisplayName
        End If
    End If
    
    If m_ListIndex = 0 Then DrawFocus m_blnFocus
End Sub


Private Sub UserControl_AmbientChanged(PropertyName As String)
    picList.Refresh
End Sub

Private Sub UserControl_EnterFocus()
    m_blnFocus = True
    If m_ListIndex Then
        DrawRows m_ListIndex, m_ListIndex, True
    Else
        DrawFocus True
    End If
End Sub

Private Sub UserControl_ExitFocus()
    m_blnFocus = False
    m_strSearch = ""
    If m_ListIndex Then
        DrawRows m_ListIndex, m_ListIndex, True
    Else
        DrawFocus False
    End If
End Sub


Private Sub UserControl_Initialize()
    Set m_Column = New ESingleArray
End Sub

Private Sub picList_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case vbKeyDown
        If m_ListIndex < m_ListCount Then
            ListIndex = m_ListIndex + 1
        End If
    Case vbKeyUp
        If m_ListIndex > 1 Then
            ListIndex = m_ListIndex - 1
        End If
    Case vbKeyPageDown
        If CLng(vsbScroll) + CLng(vsbScroll.LargeChange) <= CLng(vsbScroll.Max) Then
            vsbScroll = vsbScroll + vsbScroll.LargeChange
        Else
            vsbScroll = vsbScroll.Max
        End If
    Case vbKeyPageUp
        If vsbScroll - vsbScroll.LargeChange >= vsbScroll.Min Then
            vsbScroll = vsbScroll - vsbScroll.LargeChange
        Else
            vsbScroll = vsbScroll.Min
        End If
    Case vbKeyLeft
        If hsbScroll.Visible And hsbScroll - 1 >= hsbScroll.Min Then
            hsbScroll = hsbScroll - 1
        End If
    Case vbKeyRight
        If hsbScroll.Visible And hsbScroll + 1 <= hsbScroll.Max Then
            hsbScroll = hsbScroll + 1
        End If
    Case vbKeyA
        If (Shift And (vbCtrlMask Or vbShiftMask)) = (vbCtrlMask Or vbShiftMask) Then
            ShowAbout
        End If
    End Select
    
    
    RaiseEvent KeyDown(KeyCode, Shift)
End Sub


Private Sub picList_KeyUp(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyUp(KeyCode, Shift)
End Sub

Private Sub picList_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim lngItem As Long
    
    lngItem = Y \ GetRowHeight() + vsbScroll.Value
    If 0 < lngItem And lngItem <= m_ListCount Then
        MoveSelection lngItem
        m_blnClickOnWhiteSpace = False
    Else
        m_blnClickOnWhiteSpace = True
    End If
    
    On Error GoTo Handler
    RaiseEvent MouseDown(Button, Shift, ScaleX(X, ScaleMode, vbContainerPosition), ScaleY(Y, ScaleMode, vbContainerPosition))
    Exit Sub
Handler:
    'site does not contain a scale mode
    RaiseEvent MouseDown(Button, Shift, X, Y)
    Exit Sub
End Sub

Private Sub picList_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error GoTo Handler
    RaiseEvent MouseMove(Button, Shift, ScaleX(X, ScaleMode, vbContainerPosition), ScaleY(Y, ScaleMode, vbContainerPosition))
    Exit Sub
Handler:
    'site does not contain a scale mode
    RaiseEvent MouseMove(Button, Shift, X, Y)
    Exit Sub
End Sub

Private Sub picList_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim lngItem As Long
    Dim blnCancel As Boolean
    
    'see picList_DblClick
    lngItem = Y \ GetRowHeight() + vsbScroll.Value
    If 0 < lngItem And lngItem <= m_ListCount Then
        'check box region
        If X < GetRowHeight() And m_CheckBoxes Then
            RaiseEvent ItemChecked(lngItem, Not m_ListData(lngItem).blnSelected, blnCancel)
            If Not blnCancel Then
                m_ListData(lngItem).blnSelected = Not m_ListData(lngItem).blnSelected
            End If
        End If
        MoveSelection lngItem
        m_blnClickOnWhiteSpace = False
    Else
        m_blnClickOnWhiteSpace = True
    End If
    
    On Error GoTo Handler
    RaiseEvent MouseUp(Button, Shift, ScaleX(X, ScaleMode, vbContainerPosition), ScaleY(Y, ScaleMode, vbContainerPosition))
    Exit Sub
Handler:
    'site does not contain a scale mode
    RaiseEvent MouseUp(Button, Shift, X, Y)
    Exit Sub
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=vsbScroll,vsbScroll,-1,LargeChange
Public Property Get LargeChange() As Integer
Attribute LargeChange.VB_Description = "Returns/sets amount of change to Value property in a scroll bar when user clicks the scroll bar area."
Attribute LargeChange.VB_ProcData.VB_Invoke_Property = ";Scroll"
Attribute LargeChange.VB_MemberFlags = "400"
'/// Returns/sets amount of change to Value property in a scroll bar when user clicks the scroll bar area.
    LargeChange = vsbScroll.LargeChange
End Property

Public Property Let LargeChange(ByVal New_LargeChange As Integer)
    vsbScroll.LargeChange() = New_LargeChange
    PropertyChanged "LargeChange"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=vsbScroll,vsbScroll,-1,SmallChange
Public Property Get SmallChange() As Integer
Attribute SmallChange.VB_Description = "Returns/sets amount of change to Value property in a scroll bar when user clicks a scroll arrow."
Attribute SmallChange.VB_ProcData.VB_Invoke_Property = ";Scroll"
'/// Returns/sets amount of change to Value property in a scroll bar when user clicks a scroll arrow.
    SmallChange = vsbScroll.SmallChange
End Property

Public Property Let SmallChange(ByVal New_SmallChange As Integer)
    vsbScroll.SmallChange() = New_SmallChange
    PropertyChanged "SmallChange"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,0
Public Property Get SelectedBackColor() As OLE_COLOR
Attribute SelectedBackColor.VB_ProcData.VB_Invoke_Property = ";Appearance"
    SelectedBackColor = m_SelectedBackColor
End Property

Public Property Let SelectedBackColor(ByVal New_SelectedBackColor As OLE_COLOR)
    m_SelectedBackColor = New_SelectedBackColor
    PropertyChanged "SelectedBackColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,0
Public Property Get SelectedForeColor() As OLE_COLOR
Attribute SelectedForeColor.VB_ProcData.VB_Invoke_Property = ";Appearance"
    SelectedForeColor = m_SelectedForeColor
End Property

Public Property Let SelectedForeColor(ByVal New_SelectedForeColor As OLE_COLOR)
    m_SelectedForeColor = New_SelectedForeColor
    PropertyChanged "SelectedForeColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=7,0,0,1
Public Property Get Columns() As Integer
Attribute Columns.VB_Description = "Gets/sets the number of columns in the list box, which range from 1 to MAX_INT."
Attribute Columns.VB_ProcData.VB_Invoke_Property = "Columns"
'/// Gets/sets the number of columns in the list box, which range from 1 to MAX_INT.
    Columns = m_Columns
End Property

Public Property Let Columns(ByVal New_Columns As Integer)
    'columns 0 or 1 are the same thing meaning 1 column
    Dim lngItem As Long
    
    If New_Columns >= 0 Then
        m_Columns = New_Columns
        If New_Columns > 0 Then
            m_Column.Count = New_Columns
        Else
            m_Column.Count = 1
        End If
        ReDim m_lngTextWidth(New_Columns)
        
        For lngItem = 1 To m_ListCount
            ReDim Preserve m_ListData(lngItem).strTextSub(m_Columns)
        Next lngItem
        
        PropertyChanged "Columns"
    Else
        Err.Raise vbObjectError + 1, , "Column value must be 0 or a positive value"
    End If
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=12,0,0,0
Public Property Get Column(ByVal Col As Integer) As Single
Attribute Column.VB_Description = "Specify the width of the column in pixels.  Returns the width of the column set by user if non-zero, or the width of the largest item otherwise.  One-based array."
Attribute Column.VB_ProcData.VB_Invoke_Property = ";Behavior"
Attribute Column.VB_MemberFlags = "400"
'/// Specify the width of the column in pixels.  Returns the width of the column set by user if non-zero, or the width of the largest item otherwise.  One-based array.
    If m_Column(Col) = 0 Then
        Column = picList.TextWidth("W") * m_lngTextWidth(Col)
    Else
        Column = m_Column(Col)
    End If
End Property

Public Property Let Column(ByVal Col As Integer, ByVal New_Column As Single)
    If Col = 1 And New_Column < 0 Then 'column 1 cannot be hidden
        m_Column(Col) = 0!
    ElseIf Col > 1 And New_Column <= -1 Then 'other columns hide when -1
        m_Column(Col) = -1!
    Else
        m_Column(Col) = New_Column
    End If
    PropertyChanged "Column"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=vsbScroll,vsbScroll,-1,Value
Public Property Get TopIndex() As Integer
Attribute TopIndex.VB_Description = "Returns/sets the top visible item of a list box.  One-based array."
Attribute TopIndex.VB_ProcData.VB_Invoke_Property = ";Scroll"
Attribute TopIndex.VB_MemberFlags = "400"
'/// Returns/sets the top visible item of a list box.  One-based array.
    TopIndex = vsbScroll.Value
End Property

Public Property Let TopIndex(ByVal New_TopIndex As Integer)
    If Ambient.UserMode = False Then Err.Raise 387
    If New_TopIndex >= vsbScroll Then
        vsbScroll = vsbScroll.Max
    Else
        vsbScroll = New_TopIndex
    End If
    PropertyChanged "TopIndex"
End Property

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
    Set picList.Font = Ambient.Font
    m_SelectedBackColor = m_def_SelectedBackColor
    m_SelectedForeColor = m_def_SelectedForeColor
    m_Columns = m_def_Columns
    Set m_Column = New ESingleArray
    ReDim m_ListData(1)
    ReDim m_lngTextWidth(1)
    m_RowDivisions = m_def_RowDivisions
    m_ColumnDivisions = m_def_ColumnDivisions
    m_DivisionColor = m_def_DivisionColor
    m_CheckBoxes = m_def_CheckBoxes
    m_DblClickCheck = m_def_DblClickCheck
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    Set m_Column = New ESingleArray

    On Error GoTo Handler
    picList.BackColor = PropBag.ReadProperty("BackColor", &H8000000F)
    picList.ForeColor = PropBag.ReadProperty("ForeColor", &H80000012)
    picList.Enabled = PropBag.ReadProperty("Enabled", True)
    Set picList.Font = PropBag.ReadProperty("Font", Ambient.Font)
    Set txtEdit.Font = picList.Font
    
    vsbScroll.LargeChange = PropBag.ReadProperty("LargeChange", 1)
    vsbScroll.SmallChange = PropBag.ReadProperty("SmallChange", 1)
    m_SelectedBackColor = PropBag.ReadProperty("SelectedBackColor", m_def_SelectedBackColor)
    m_SelectedForeColor = PropBag.ReadProperty("SelectedForeColor", m_def_SelectedForeColor)
    m_CheckBoxes = PropBag.ReadProperty("CheckBoxes", m_def_CheckBoxes)
    vsbScroll.Value = PropBag.ReadProperty("TopIndex", 0)
    m_RowDivisions = PropBag.ReadProperty("RowDivisions", m_def_RowDivisions)
    m_ColumnDivisions = PropBag.ReadProperty("ColumnDivisions", m_def_ColumnDivisions)
    m_DivisionColor = PropBag.ReadProperty("DivisionColor", m_def_DivisionColor)
    m_DblClickCheck = PropBag.ReadProperty("DblClickCheck", m_def_DblClickCheck)
    
    m_Columns = PropBag.ReadProperty("Columns", m_def_Columns)
    ReDim m_lngTextWidth(m_Columns)
    
    m_Column.Count = m_Columns
    Set m_Column = PropBag.ReadProperty("Column")
    
    Exit Sub
Handler:
    Exit Sub
End Sub

Private Sub UserControl_Resize()
    On Error GoTo Catch
    picList.Move 0, 0, ScaleWidth, ScaleHeight
    vsbScroll.Move picList.ScaleWidth - vsbScroll.Width, 0, vsbScroll.Width, picList.ScaleHeight
    hsbScroll.Move 0, picList.ScaleHeight - hsbScroll.Height, picList.ScaleWidth - IIf(vsbScroll.Visible, vsbScroll.Width, 0)
    Exit Sub
Catch:
    Err.Clear
    Exit Sub
End Sub

Private Sub UserControl_Terminate()
    Set m_Column = Nothing
    Erase m_ListData
    Erase m_lngTextWidth
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    On Error GoTo Handler
    Call PropBag.WriteProperty("BackColor", picList.BackColor, &H8000000F)
    Call PropBag.WriteProperty("ForeColor", picList.ForeColor, &H80000012)
    Call PropBag.WriteProperty("Enabled", picList.Enabled, True)
    Call PropBag.WriteProperty("Font", picList.Font, Ambient.Font)
    Call PropBag.WriteProperty("LargeChange", vsbScroll.LargeChange, 1)
    Call PropBag.WriteProperty("SmallChange", vsbScroll.SmallChange, 1)
    Call PropBag.WriteProperty("SelectedBackColor", m_SelectedBackColor, m_def_SelectedBackColor)
    Call PropBag.WriteProperty("SelectedForeColor", m_SelectedForeColor, m_def_SelectedForeColor)
    Call PropBag.WriteProperty("RowDivisions", m_RowDivisions, m_def_RowDivisions)
    Call PropBag.WriteProperty("ColumnDivisions", m_ColumnDivisions, m_def_ColumnDivisions)
    Call PropBag.WriteProperty("DivisionColor", m_DivisionColor, m_def_DivisionColor)
    Call PropBag.WriteProperty("CheckBoxes", m_CheckBoxes, m_def_CheckBoxes)
    Call PropBag.WriteProperty("DblClickCheck", m_DblClickCheck, m_def_DblClickCheck)
    
    Call PropBag.WriteProperty("Columns", m_Columns, m_def_Columns)
    Call PropBag.WriteProperty("Column", m_Column)
    
    Exit Sub
Handler:
    Exit Sub
End Sub


'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,2,0
Public Property Get Selected(ByVal Index As Long) As Boolean
Attribute Selected.VB_Description = "One-based array.  Determines if an item is checked."
Attribute Selected.VB_ProcData.VB_Invoke_Property = ";List"
Attribute Selected.VB_MemberFlags = "400"
'/// One-based array.  Determines if an item is checked.
    Selected = m_ListData(Index).blnSelected
End Property

Public Property Let Selected(ByVal Index As Long, ByVal New_Selected As Boolean)
    If Ambient.UserMode = False Then Err.Raise 387
    m_ListData(Index).blnSelected = New_Selected
    If Index > 0 Then DrawRows Index, Index, True
    PropertyChanged "Selected"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=8,0,2,0
Public Property Get ItemData(ByVal Index As Long) As Long
Attribute ItemData.VB_Description = "One-based array.  Data is not shown to the user."
Attribute ItemData.VB_ProcData.VB_Invoke_Property = ";List"
Attribute ItemData.VB_MemberFlags = "400"
'/// One-based array.  Data is not shown to the user.
    ItemData = m_ListData(Index).lngItemData
End Property

Public Property Let ItemData(ByVal Index As Long, ByVal New_ItemData As Long)
    If Ambient.UserMode = False Then Err.Raise 387
    m_ListData(Index).lngItemData = New_ItemData
End Property


'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,0
Public Property Get RowDivisions() As Boolean
Attribute RowDivisions.VB_Description = "Shows/hides the row division line."
Attribute RowDivisions.VB_ProcData.VB_Invoke_Property = "Columns"
'/// Shows/hides the row division line.
    RowDivisions = m_RowDivisions
End Property

Public Property Let RowDivisions(ByVal New_RowDivisions As Boolean)
    m_RowDivisions = New_RowDivisions
    PropertyChanged "RowDivisions"
End Property

Private Sub vsbScroll_Change()
    picList.Refresh
    If txtEdit.Visible Then EditBoxOff
End Sub

Private Sub vsbScroll_Scroll()
    picList.Refresh
    If txtEdit.Visible Then EditBoxOff
End Sub

