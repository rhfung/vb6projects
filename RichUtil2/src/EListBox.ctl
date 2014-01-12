VERSION 5.00
Begin VB.UserControl EListBox 
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4770
   ForeColor       =   &H80000008&
   KeyPreview      =   -1  'True
   LockControls    =   -1  'True
   PropertyPages   =   "EListBox.ctx":0000
   ScaleHeight     =   240
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   318
   ToolboxBitmap   =   "EListBox.ctx":001C
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
      Begin VB.PictureBox picGray 
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   1800
         ScaleHeight     =   17
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   17
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   2160
         Width           =   255
      End
      Begin VB.TextBox txtEdit 
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   0
         TabIndex        =   3
         Top             =   0
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
   Begin VB.Image imgAbout 
      Height          =   750
      Left            =   1800
      Picture         =   "EListBox.ctx":0116
      Top             =   1560
      Visible         =   0   'False
      Width           =   2100
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
'// Started January 2, 2003
'// Modified December 22, 2003
'//
'// Enhanced List Box developed by Richard H Fung
'// Copyright (C) 2003-2007, Richard H Fung for Hai Consultants Inc.
'//
'// * This version corrects mistakes in RichUtil1.ocx and RichUtil.dll
'// * Edit box now shows properly even if there is no column width
'// * The Columns property page has bug fixes
'// * List and ItemData are now saved with the control at design time
'// * Fixes to scroll bar calculations, scrolling, and extra white space
'//   at the end of the list
'// * Added AddItem3() that uses string tokens to separate columns
'// * Column(Col) returns correct width set by user
'// * ColumnRealWidth(Col) returns run-time width
'// * ReadyForResize(Max_Size) allows the array to resized quite large
'//   before calling AddItemX()
'// * Restrictions removed for length of list box to 32,767 items
'//
'// Modified May 23, 2004 for
'//     - full width selection of the last edit field
'//     - RowData() run-time variable
'//     - fixes to dynamic resizing of list
'//     - AddMode to quickly redraw the list
'//     - KeyDown, Press, and Up events for edit box
'//     - RowCol click events
'//     - adding columns and maintaining auto column width
'//     - corner gray box and corner events
'//     - mouse events always in pixels
'//     - changed ListIndex = 0 focus rectangle to only highlight first row
'//     - Locked property for edit box
'//     - AddItem3Assign() function
'//     - ListChange2() function uses TokenChar
'//     - TokenChar() property
'//     - Can load tokenized List property at Read and Write properties.
'//     - Allows the columns to be filled in at design-time.
'//     - Changes to List property page to enable this feature.
'//
'// Modified August 14-18, 2007 for
'//     - using in Design 3 (PLTP)
'//     - new AutoSizeColumns
'//     - new AddRow
'//     - changed Set Font
'//     - new SelectRowText
'//     - render the focus rectangle better
'//     - render the check mark darker
'//     - default font is Tahoma in About box (Windows 95 and NT are not supported)
'//     - new EditMaxLength for text box
'//     - new EditSelLength for text box
'//     - new EditSelStart for text box
'//     - new EditSelText for text box
'//     - changed RemoveItem to handle boundary conditions
'//     - changed RemoveItem to paint the visible changes only
'//     - changed Enabled to actually disable the real control
'//     - changed how the list box appears when there are no items
'//     - observed SmallChange and LargeChange do not work
'//     - new TypeToSearch property
'//     - changed KeyPress will now ask ItemChecked for check box list
'//     - arrow keys will reset the TypeToSearch query
'//
'// Modified January 12, 2014:
'//     - OcxHandshake() no longer required to use the list box

'Default Property Values:
Const m_def_RowDivisions = False
Const m_def_ColumnDivisions = False
Const m_def_DivisionColor = vbButtonFace
Const m_def_SelectedBackColor = vbHighlight
Const m_def_SelectedForeColor = vbHighlightText
Const m_def_Columns = 1
Const m_def_CheckBoxes = False
Const m_def_DblClickCheck = True
Const m_def_TypeToSearch = True

'Property Variables:
Dim m_RowDivisions      As Boolean
Dim m_ColumnDivisions   As Boolean
Dim m_CheckBoxes        As Boolean
Dim m_DblClickCheck     As Boolean
Dim m_SelectedBackColor As OLE_COLOR
Dim m_SelectedForeColor As OLE_COLOR
Dim m_DivisionColor     As OLE_COLOR
Dim m_Columns           As Long
Dim m_Column            As ESingleArray
Private m_blnTypeToSearch As Boolean 'True: type to search; False: intercept to edit

'run time variables
Private m_strTokenChar As String

Private Type EListData
    strText        As String
    blnSelected    As Boolean
    lngItemData    As Long
    strTextSub()   As String 'from 1 to Columns, but last element is not used
    vntData        As Variant 'user property RowData
End Type

'edit box
Private m_lngColX()   As Long 'left position of a column 1 to Count
Private m_lngEditRow  As Long
Attribute m_lngEditRow.VB_VarDescription = "Occurs when an item is selected by keyboard or mouse.  One-based array."
Private m_lngEditCol  As Long
Attribute m_lngEditCol.VB_VarDescription = "Occurs when an item is selected by keyboard or mouse.  One-based array."

'list
Private m_ListData() As EListData
Private m_ListCount  As Long
Attribute m_ListCount.VB_VarDescription = "Occurs when an item is selected by keyboard or mouse.  One-based array."
Private m_ListUBound As Long
Attribute m_ListUBound.VB_VarDescription = "Occurs when an item is selected by keyboard or mouse.  One-based array."
Private m_ListIndex  As Long 'between 1 and ListCount
Attribute m_ListIndex.VB_VarDescription = "Occurs when an item is selected by keyboard or mouse.  One-based array."

'visual
Private m_blnFocus   As Boolean
Attribute m_blnFocus.VB_VarDescription = "Occurs when an item is selected by keyboard or mouse.  One-based array."
Private m_lngTextWidth() As Long 'run-time width of string for column 1 to Count
Private m_lngMultiplier As Long 'multiplying factor for scroll bar

'behaviour
Private m_strSearch             As String
Attribute m_strSearch.VB_VarDescription = "Occurs when an item is selected by keyboard or mouse.  One-based array."
Private m_blnClickOnWhiteSpace  As Boolean
Attribute m_blnClickOnWhiteSpace.VB_VarDescription = "Occurs when an item is selected by keyboard or mouse.  One-based array."
Private m_blnFullRectShown      As Boolean
Attribute m_blnFullRectShown.VB_VarDescription = "Occurs when an item is selected by keyboard or mouse.  One-based array."

'behaviour: 23 May 2004
Private m_blnAddMode            As Boolean 'supress redrawing list box
Private m_lngHandshake1         As Long 'to supress the logo when more than 4 columns in the list
Private m_lngClickedOnRow       As Long 'for RowColClick events
Private m_lngClickedOnCol       As Long 'for RowColClick events
Private m_lngHandshake2         As Long 'to supress the logo when more than 4 columns in the list


'Event Declarations:
Event Click(ByVal ListIndex As Long)  '/// Occurs when an item is selected by keyboard or mouse.  One-based array.
Attribute Click.VB_Description = "Occurs when an item is selected by keyboard or mouse.  One-based array."
Attribute Click.VB_UserMemId = -600
Event DblClick() '/// Occurs when the user presses and releases a mouse button and then presses and releases it again over an object.
Attribute DblClick.VB_Description = "Occurs when the user presses and releases a mouse button and then presses and releases it again over an object."
Attribute DblClick.VB_UserMemId = -601
Event KeyDown(KeyCode As Integer, Shift As Integer) '/// Occurs when the user presses a key while an object has the focus.
Attribute KeyDown.VB_Description = "Occurs when the user presses a key while an object has the focus."
Attribute KeyDown.VB_UserMemId = -602
Event KeyPress(KeyAscii As Integer) '/// Occurs when the user presses and releases an ANSI key.
Attribute KeyPress.VB_Description = "Occurs when the user presses and releases an ANSI key."
Attribute KeyPress.VB_UserMemId = -603
Event KeyUp(KeyCode As Integer, Shift As Integer) '/// Occurs when the user releases a key while an object has the focus.
Attribute KeyUp.VB_Description = "Occurs when the user releases a key while an object has the focus."
Attribute KeyUp.VB_UserMemId = -604
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single) '/// Occurs when the user presses the mouse button while an object has the focus.
Attribute MouseDown.VB_Description = "Occurs when the user presses the mouse button while an object has the focus."
Attribute MouseDown.VB_UserMemId = -605
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single) '/// Occurs when the user moves the mouse.
Attribute MouseMove.VB_Description = "Occurs when the user moves the mouse."
Attribute MouseMove.VB_UserMemId = -606
Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single) '/// Occurs when the user releases the mouse button while an object has the focus.
Attribute MouseUp.VB_Description = "Occurs when the user releases the mouse button while an object has the focus."
Attribute MouseUp.VB_UserMemId = -607

Event ItemChecked(ByVal ListIndex As Long, ByVal Checked As Boolean, ByRef Cancel As Boolean) '/// Occurs when an item in the list box is checked and CheckBoxes variable is True.
Attribute ItemChecked.VB_Description = "Occurs when an item in the list box is checked and CheckBoxes variable is True."
Event ValidateEdit(ByVal Row As Long, ByVal Col As Long, ByVal EditText As String, ByRef Cancel As Boolean) '/// Occurs when the edit text box hides.
Attribute ValidateEdit.VB_Description = "Occurs when the edit text box hides."

'added May 23, 2004
Event BeforeEdit(ByRef Row As Long, ByRef Col As Long, ByRef Cancel As Boolean) '/// Occurs before the edit box is shown.
Attribute BeforeEdit.VB_Description = "Occurs before the edit box is shown."
Event AfterEdit(ByVal Row As Long, ByVal Col As Long) '/// Occurs after editing when the edit box has hidden.
Attribute AfterEdit.VB_Description = "Occurs after editing when the edit box has hidden."
'
Event KeyDownEdit(KeyCode As Integer, Shift As Integer) '/// The key was pressed inside the edit box.
Attribute KeyDownEdit.VB_Description = "The key was pressed inside the edit box."
Event KeyPressEdit(KeyAscii As Integer) '/// The key was pressed inside the edit box.
Attribute KeyPressEdit.VB_Description = "The key was pressed inside the edit box."
Event KeyUpEdit(KeyCode As Integer, Shift As Integer) '/// The key was pressed inside the edit box.
Attribute KeyUpEdit.VB_Description = "The key was pressed inside the edit box."
'
Event RowColClick(ByVal Row As Long, ByVal Col As Long) '/// A click event.
Attribute RowColClick.VB_Description = "A click event."
Attribute RowColClick.VB_MemberFlags = "200"
Event RowColDblClick(ByVal Row As Long, ByVal Col As Long) '/// A double click event.
Attribute RowColDblClick.VB_Description = "A double click event."
'
Event CornerClick() '/// Lower right hand corner box when both scroll bars are visible.
Attribute CornerClick.VB_Description = "Lower right hand corner box when both scroll bars are visible."
Event CornerDblClick() '/// Lower right hand corner box when both scroll bars are visible.
Attribute CornerDblClick.VB_Description = "Lower right hand corner box when both scroll bars are visible."
Event CornerMouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single) '/// Lower right hand corner box when both scroll bars are visible.
Attribute CornerMouseDown.VB_Description = "Lower right hand corner box when both scroll bars are visible."
Event CornerMouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single) '/// Lower right hand corner box when both scroll bars are visible.
Attribute CornerMouseMove.VB_Description = "Lower right hand corner box when both scroll bars are visible."
Event CornerMouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single) '/// Lower right hand corner box when both scroll bars are visible.
Attribute CornerMouseUp.VB_Description = "Lower right hand corner box when both scroll bars are visible."


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


Private Sub Add()
'Increases the size of the array.
    If m_ListCount > m_ListUBound Then
        On Error GoTo Handler
        m_ListUBound = m_ListUBound * 2
        ReDim Preserve m_ListData(m_ListUBound)
    End If
    
    Exit Sub
Handler:
    m_ListUBound = &H7FFFFFFF
    Resume Next
End Sub

Public Property Get AddMode() As Boolean
Attribute AddMode.VB_Description = "Determines if list is in AddMode.  In add mode, the list is not redrawn."
Attribute AddMode.VB_MemberFlags = "400"
'/// Determines if list is in AddMode.  In add mode, the list is not redrawn.
    AddMode = m_blnAddMode
End Property

Public Property Let AddMode(ByVal pAddMode As Boolean)
    m_blnAddMode = pAddMode
    If pAddMode = False Then
        Refresh
    End If
End Property

Public Sub AddItem(ByVal Item As String)
Attribute AddItem.VB_Description = "Adds an item to the list."
'/// Adds an item to the list.
    m_ListCount = m_ListCount + 1
    Add
    
    m_ListData(m_ListCount).strText = Item
    If m_Columns >= 1 Then ReDim m_ListData(m_ListCount).strTextSub(m_Columns)
    
    'set max column width
    If picList.TextWidth(Item) > m_lngTextWidth(1) Then
        m_lngTextWidth(1) = picList.TextWidth(Item)
    End If
    
    RefreshListIfNotAdding
    PropertyChanged "ListCount"
End Sub

Public Sub AddItem2(ByVal Item As String, ByVal ItemData As Long, ParamArray args() As Variant)
Attribute AddItem2.VB_Description = "Adds an item, its item data, and its associated column strings to the list.  Use AddRow() instead."
Attribute AddItem2.VB_MemberFlags = "40"
'/// Adds an item, its item data, and its associated column strings to the list.  Use AddRow() instead.
    m_ListCount = m_ListCount + 1
    Add
    
    With m_ListData(m_ListCount)
        'assign text
        .strText = Item
        'set max column width
        If picList.TextWidth(Item) > m_lngTextWidth(1) Then
            m_lngTextWidth(1) = picList.TextWidth(Item)
        End If
        
        'assign item data
        .lngItemData = ItemData
        
        'assign extra columns
        If m_Columns > 1 Then
            ReDim .strTextSub(m_Columns)
            
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
    RefreshListIfNotAdding
    PropertyChanged "ListCount"
End Sub

'Added a new method to insert elements into the list control.  (RF, 14-Aug-2007)
Public Sub AddRow(ParamArray Columns() As Variant)
Attribute AddRow.VB_Description = "Adds an item to the end of the list with ItemData set to 0.  Extra columns are ignored."
'/// Adds an item to the end of the list with ItemData set to 0.  Extra columns are ignored.
    m_ListCount = m_ListCount + 1
    Add
    
    With m_ListData(m_ListCount)
        'assign text
        .strText = Columns(LBound(Columns))
        'set max column width
        If picList.TextWidth(.strText) > m_lngTextWidth(1) Then
            m_lngTextWidth(1) = picList.TextWidth(.strText)
        End If
        
        'assign item data to a default value of 0
        .lngItemData = 0
        
        'assign extra columns
        If UBound(Columns) > LBound(Columns) Then
            ReDim .strTextSub(m_Columns)
            
            Dim lngItem     As Long
            Dim lngIndex    As Long
            
            For lngItem = LBound(Columns) + 1 To UBound(Columns)
                lngIndex = lngIndex + 1 'increment index counter
                If lngIndex > m_Columns - 1 Then Exit For 'read until filled up space - trash remaining column data
                'fill in text data
                .strTextSub(lngIndex) = Columns(lngItem)
                'set max column width
                If picList.TextWidth(.strTextSub(lngIndex)) > m_lngTextWidth(lngIndex + 1) Then
                    m_lngTextWidth(lngIndex + 1) = picList.TextWidth(.strTextSub(lngIndex))
                End If
            Next lngItem
        End If
    End With
    RefreshListIfNotAdding
    PropertyChanged "ListCount"
End Sub

Public Sub AddItem3(ByVal TokenizedString As String, Optional Token As String = vbTab, Optional ItemData As Long = 0&)
Attribute AddItem3.VB_Description = "Adds an item to the list where the TokenizedString's columns are split using Token.  Use AddRow() instead."
Attribute AddItem3.VB_MemberFlags = "40"
'/// Adds an item to the list where the TokenizedString's columns are split using Token.  Use AddRow() instead.
'NOTE: Token and ItemData should be passed ByVal
    m_ListCount = m_ListCount + 1
    Add
    
    'See also AddItem3Assign()
    
    'create parameters
    Dim args()          As String
    Dim lngArgCount     As Long
    
    args = Split(TokenizedString, Token, Count)
    lngArgCount = UBound(args) - LBound(args) + 1
    
    With m_ListData(m_ListCount)
        If lngArgCount > 0 Then
            'assign text
            .strText = args(0)
            'set max column width
            If picList.TextWidth(.strText) > m_lngTextWidth(1) Then
                m_lngTextWidth(1) = picList.TextWidth(.strText)
            End If
        End If
        
        'assign item data
        .lngItemData = ItemData
        
        'assign extra columns
        If m_Columns > 1 Then
            ReDim .strTextSub(m_Columns)
            
            Dim lngItem     As Long
            Dim lngIndex    As Long
            
            For lngItem = 1 To lngArgCount - 1
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
    RefreshListIfNotAdding
    PropertyChanged "ListCount"
End Sub

Private Sub AddItem3Assign(ByVal ListIndex As Long, ByVal TokenizedString As String, Optional ByVal Token As String = vbTab, Optional ByVal ItemData As Long = 0&)
'Does not add an item - modifies existing item.
'See Also : AddItem3
    'create parameters
    Dim args()          As String
    Dim lngArgCount     As Long
    
    args = Split(TokenizedString, Token, Count)
    lngArgCount = UBound(args) - LBound(args) + 1
    
    With m_ListData(ListIndex)
        If lngArgCount > 0 Then
            'assign text
            .strText = args(0)
            'set max column width
            If picList.TextWidth(.strText) > m_lngTextWidth(1) Then
                m_lngTextWidth(1) = picList.TextWidth(.strText)
            End If
        End If
        
        'assign item data
        .lngItemData = ItemData
        
        'assign extra columns
        If m_Columns > 1 Then
            ReDim .strTextSub(m_Columns)
            
            Dim lngItem     As Long
            Dim lngIndex    As Long
            
            For lngItem = 1 To lngArgCount - 1
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

End Sub

'Added RF, 14-Aug-2007
Public Sub AutoSizeColumns(ByVal ForceFixedColumns As Boolean)
Attribute AutoSizeColumns.VB_Description = "Resize columns to best fit that are auto-sized and also fixed columns if flag is set."
'/// Resize columns to best fit that are auto-sized and also fixed columns if flag is set.
    Dim lngCol As Long
    Dim lngRow As Long
    Dim strSample As String
    
    For lngCol = 1 To m_Columns
        'm_Column(lngCol) = 0 --> auto-size
        'm_Column(lngCol) > 0 --> fixed width
        'm_Column(lngCol) < 0 --> hidden
        If (m_Column(lngCol) = 0) Or (ForceFixedColumns And (m_Column(lngCol) >= 0)) Then
            m_lngTextWidth(lngCol) = 0
            For lngRow = 1 To m_ListCount
                'select the column text
                If lngCol = 1 Then
                    strSample = m_ListData(lngRow).strText
                Else
                    strSample = m_ListData(lngRow).strTextSub(lngCol - 1)
                End If
                
                'maximize function
                If picList.TextWidth(strSample) > m_lngTextWidth(lngCol) Then
                    m_lngTextWidth(lngCol) = picList.TextWidth(strSample)
                End If
            Next lngRow
        End If
    Next lngCol
    
    RefreshListIfNotAdding
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
Attribute CheckBoxes.VB_ProcData.VB_Invoke_PropertyPut = "Grid;Behavior"
    m_CheckBoxes = pCheck
    RefreshListIfNotAdding
    PropertyChanged "CheckBoxes"
End Property

Public Property Get CheckBoxes() As Boolean
'/// Determines if check boxes are shown in the list box.
    CheckBoxes = m_CheckBoxes
End Property

Public Sub Clear()
Attribute Clear.VB_Description = "Clears all elements from the list."
'/// Clears all elements from the list.
    ReDim m_ListData(10)
    m_ListCount = 0
    m_ListUBound = 10
    m_strSearch = ""
    
    UpdateList
    PropertyChanged "ListCount"
    PropertyChanged "List"
    PropertyChanged "ItemData"
End Sub


'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=7,0,0,1
Public Property Get Columns() As Integer
Attribute Columns.VB_Description = "Gets/sets the number of columns in the list box, which range from 1 to MAX_INT."
Attribute Columns.VB_ProcData.VB_Invoke_Property = "Columns;Behavior"
'/// Gets/sets the number of columns in the list box, which range from 1 to MAX_INT.
    Columns = m_Columns
End Property

Public Property Let Columns(ByVal New_Columns As Integer)
    'columns 0 or 1 are the same thing meaning 1 column
    Dim lngItem As Long
    
    If New_Columns >= 0 Then
        If New_Columns = 0 Then 'assign correct value
            New_Columns = 1
        End If
        
        m_Columns = New_Columns
        m_Column.Count = New_Columns
        ReDim Preserve m_lngTextWidth(New_Columns) 'fixed 23 May 2004
        
        For lngItem = 1 To m_ListCount
            'If UBound(m_ListData(lngItem).strTextSub) < LBound(m_ListData(lngItem).strTextSub) Then
            '    ReDim m_ListData(lngItem).strTextSub(m_Columns )
            'ElseIf UBound(m_ListData(lngItem).strTextSub) < m_Columns  Then
            ReDim Preserve m_ListData(lngItem).strTextSub(m_Columns)
            'End If
        Next lngItem
        
        RefreshListIfNotAdding
        PropertyChanged "Columns"
    Else
        Err.Raise vbObjectError + 1, , "Column value must be 0 or a positive value"
    End If
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=12,0,0,0
Public Property Get Column(ByVal Col As Integer) As Single
Attribute Column.VB_Description = "Gets the width of the column set by user.  Sets the width of the column in pixels.   If 0, the column width is fit to the text width; use ColumnRealWidth to get that value.   If negative, then the column is hidden, except for the first column."
Attribute Column.VB_MemberFlags = "400"
'/// Gets the width of the column set by user.  Sets the width of the column in pixels.
'/// If 0, the column width is fit to the text width; use ColumnRealWidth to get that value.
'/// If negative, then the column is hidden, except for the first column.
    'If m_Column(Col) = 0 Then
    '    Column = picList.TextWidth("W") * m_lngTextWidth(Col)
    'Else
    Column = m_Column(Col)
    'End If
End Property

Public Property Let Column(ByVal Col As Integer, ByVal New_Column As Single)
    If Col = 1 And New_Column < 0! Then 'column 1 cannot be hidden
        m_Column(Col) = 0!
    ElseIf Col > 1 And New_Column <= -1! Then 'other columns hide when -1
        m_Column(Col) = -1!
    Else
        m_Column(Col) = New_Column
    End If
    RefreshListIfNotAdding
    PropertyChanged "Column"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,0
Public Property Get ColumnDivisions() As Boolean
Attribute ColumnDivisions.VB_Description = "Shows/hides the column division line."
Attribute ColumnDivisions.VB_ProcData.VB_Invoke_Property = ";Appearance"
'/// Shows/hides the column division line.
    ColumnDivisions = m_ColumnDivisions
End Property

Public Property Let ColumnDivisions(ByVal New_ColumnDivisions As Boolean)
    m_ColumnDivisions = New_ColumnDivisions
    RefreshListIfNotAdding
    PropertyChanged "ColumnDivisions"
End Property

Public Property Get DivisionColor() As OLE_COLOR
Attribute DivisionColor.VB_Description = "Gets/sets the colour of the division lines."
Attribute DivisionColor.VB_ProcData.VB_Invoke_Property = ";Appearance"
'/// Gets/sets the colour of the division lines.
    DivisionColor = m_DivisionColor
End Property

Public Property Let DivisionColor(ByVal vNewValue As OLE_COLOR)
    m_DivisionColor = vNewValue
    RefreshListIfNotAdding
    PropertyChanged "DivisionColor"
End Property

Public Property Get DblClickCheck() As Boolean
Attribute DblClickCheck.VB_Description = "When double-click on an item, is the item checked?"
Attribute DblClickCheck.VB_ProcData.VB_Invoke_Property = "Grid;Behavior"
'/// When double-click on an item, is the item checked?
    DblClickCheck = m_DblClickCheck
End Property

Public Property Let DblClickCheck(ByVal pVal As Boolean)
    m_DblClickCheck = pVal
    PropertyChanged "DblClickCheck"
End Property

Private Sub DrawFocus(ByVal Visible As Boolean)
    Static tRect As RECT
    Dim lngColour As Long
            
    If Visible <> m_blnFullRectShown Then
        tRect.Right = picList.ScaleWidth - IIf(vsbScroll.Visible, vsbScroll.Width, 0)
        tRect.Bottom = GetRowHeight  'picList.ScaleHeight - IIf(hsbScroll.Visible, hsbScroll.Height, 0)
        
        'use the same fore colour for the focus rectangle every time (RF, 16-Aug-2007)
        lngColour = picList.ForeColor
        picList.ForeColor = SelectedBackColor
        DrawFocusRect picList.hdc, tRect
        picList.ForeColor = lngColour
        
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
    Dim lngRowSpacer As Long 'add in a row spacer that should be 0 or 1 px
    Dim tRect       As RECT
    Static lngStack As Long 'protect from recursion
    
    Const conSpacer = 10
    
    On Error Resume Next
    
    If lngStack > 2 Then Exit Sub
    
    'absolute column positions
    lngHeight = GetRowHeight()
    lngWidth = picList.ScaleWidth - IIf(vsbScroll.Visible, vsbScroll.Width, 0)
    If m_CheckBoxes Then lngCheck = lngHeight - hsbScroll 'width of check box adjusted with scroll bar X
    lngWidthCheck = IIf(m_CheckBoxes, lngHeight, 0)
    
    ReDim m_lngColX(m_Columns)
    
    If m_Column(1) <= 0 Then 'first column must show
        m_lngColX(1) = m_lngTextWidth(1) + lngWidthCheck + conSpacer
    Else
        m_lngColX(1) = m_Column(1) + lngWidthCheck + conSpacer
    End If

    For lngSub = 2 To m_Columns
        If m_Column(lngSub) <= -1 Then 'hide column
            m_lngColX(lngSub) = m_lngColX(lngSub - 1)
        ElseIf m_Column(lngSub) = 0 Then 'auto-size
            m_lngColX(lngSub) = m_lngColX(lngSub - 1) + m_lngTextWidth(lngSub) + conSpacer
        Else 'fixed size
            m_lngColX(lngSub) = m_lngColX(lngSub - 1) + m_Column(lngSub) + conSpacer
        End If
    Next lngSub

    'horizontal scroll bar adjustments -> slide all the items left
    If m_lngColX(m_Columns) - picList.ScaleWidth > 0 Then
        hsbScroll.Max = m_lngColX(m_Columns) - picList.ScaleWidth + IIf(vsbScroll.Visible, vsbScroll.Width, 0)
        hsbScroll.LargeChange = picList.ScaleWidth - IIf(vsbScroll.Visible, vsbScroll.Width, 0)
        For lngSub = 1 To m_Columns
            m_lngColX(lngSub) = m_lngColX(lngSub) - hsbScroll
        Next lngSub
        'well now I know that the scroll bar needs to be shown -
        'have to redo all the calculations
        If hsbScroll.Visible = False Then
            lngStack = lngStack + 1
            hsbScroll.Visible = True
            UpdateList
            DrawRows FromIndex, ToIndex, BackFill
            lngStack = lngStack - 1
            Exit Sub
        End If
    Else
        'since removing the scroll bar, have to redo the calculations
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
    
    'adjust indices
    If FromIndex < 1 Then FromIndex = 1 'first element
    If ToIndex > ListCount Then ToIndex = m_ListCount  'last element
    
    'row spacer to make the focus rect look nice in either mode (RF, 14-Aug-2007)
    If m_RowDivisions Then
        lngRowSpacer = 1
    Else
        lngRowSpacer = 0
    End If
    
    'drawing the rows
    For lngItem = FromIndex To ToIndex
        lngLastY = (lngItem - vsbScroll * m_lngMultiplier) * lngHeight
        tRect.Left = lngCheck
        tRect.Top = lngLastY + 1
        tRect.Right = lngWidth
        tRect.Bottom = tRect.Top + lngHeight - 1
        
        'display highlight
        If lngItem = m_ListIndex Then
            'for the selected item
            '- modify the colour first to do some rendering and restore colour before exiting the function
            
            If BackFill Then
                'white background
                picList.Line (0, lngLastY)-(lngWidth, lngLastY + lngHeight - 1), picList.BackColor, BF
            End If
            picList.ForeColor = SelectedBackColor 'make this look like standard text box (RF, 14-Aug-2007)
            If m_blnFocus Then
                picList.Line (tRect.Left, tRect.Top)-(tRect.Right, tRect.Bottom - 1), SelectedBackColor, BF 'make this look like standard text box (RF, 14-Aug-2007)
                DrawFocusRect picList.hdc, tRect
                picList.ForeColor = IIf(picList.Enabled, SelectedForeColor, vbGrayText)
            Else
                picList.Line (tRect.Left, tRect.Top)-(tRect.Right, tRect.Bottom - 1), SelectedBackColor, BF 'make this look like standard text box (RF, 14-Aug-2007)
                picList.ForeColor = IIf(picList.Enabled, SelectedForeColor, vbGrayText)
            End If
            'picList.Line (1 + lngCheck, lngLastY + 2)-Step(lngWidth - lngCheck - 1, lngHeight - 4), vbGreen, BF 'SelectedBackColor, BF
            picList.Print "";
        Else
            'not the selected item
            picList.ForeColor = IIf(picList.Enabled, UserControl.ForeColor, vbGrayText)
            If BackFill Then
                'white background
                picList.Line (0, lngLastY)-(lngWidth, lngLastY + lngHeight - 1), picList.BackColor, BF
            End If
        End If
        
        If m_CheckBoxes Then
            'display check box
            picList.Line (2 - hsbScroll, lngLastY + 2)-Step(lngHeight - 4, lngHeight - 4), UserControl.ForeColor, B
            'display X -- darkened to make it more visible (RF, 14-Aug-2007)
            If m_ListData(lngItem).blnSelected Then
                'picList.Line (2 - hsbScroll, lngLastY + 3)-Step(lngHeight - 5, lngHeight - 5), UserControl.ForeColor
                picList.Line (2 - hsbScroll, lngLastY + 2)-Step(lngHeight - 4, lngHeight - 4), UserControl.ForeColor
                picList.Line (3 - hsbScroll, lngLastY + 2)-Step(lngHeight - 4, lngHeight - 4), UserControl.ForeColor
                
                'picList.Line (1 - hsbScroll, lngLastY + lngHeight - 2)-Step(lngHeight - 5, -lngHeight + 5), UserControl.ForeColor
                picList.Line (2 - hsbScroll, lngLastY + lngHeight - 2)-Step(lngHeight - 4, -lngHeight + 4), UserControl.ForeColor
                picList.Line (3 - hsbScroll, lngLastY + lngHeight - 2)-Step(lngHeight - 4, -lngHeight + 4), UserControl.ForeColor
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
            If m_Column(lngSub + 1) >= 0 Then 'show other columns if chosen
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
    
    'draw column lines if needed
    If m_ColumnDivisions Then
        For lngSub = 1 To m_Columns - 1
            picList.Line (m_lngColX(lngSub), 0)-Step(0, picList.ScaleHeight), m_DivisionColor
        Next lngSub
    End If
    
    'reposition text box
    'see also EditBoxOn()
    If txtEdit.Tag = "F" Then
        txtEdit.Left = lngWidthCheck + conSpacer \ 2 - IIf(hsbScroll.Visible, hsbScroll, 0)
        txtEdit.Width = picList.ScaleWidth - vsbScroll.Width
    Else
        
        If m_lngEditCol = 1 Then
            txtEdit.Left = lngWidthCheck + conSpacer \ 2 - IIf(hsbScroll.Visible, hsbScroll, 0)
            'txtEdit.Width = m_lngColX(Col) - txtEdit.Left - conSpacer / 2
        ElseIf m_lngEditCol = m_Columns Then 'last column
            txtEdit.Left = m_lngColX(m_lngEditCol - 1) + conSpacer / 2
            'txtEdit.Width = picList.ScaleWidth - txtEdit.Left - conSpacer / 2
        ElseIf m_lngEditCol > 1 Then 'somewhere in between
            txtEdit.Left = m_lngColX(m_lngEditCol - 1) + conSpacer / 2
            'txtEdit.Width = m_lngColX(Col) - txtEdit.Left - conSpacer / 2
        End If
        
    End If


End Sub

Public Sub EditBoxOff(Optional ByVal Cancel As Boolean)
Attribute EditBoxOff.VB_Description = "Hides the edit box.  If Cancel = False, the data is saved."
'/// Hides the edit box.  If Cancel = False, the data is saved.
    If Not txtEdit.Visible Then Exit Sub

    If Not Cancel Then
        RaiseEvent ValidateEdit(m_lngEditRow, m_lngEditCol, txtEdit.Text, Cancel)
    End If

    If Not Cancel Then
        ListSub(m_lngEditRow, m_lngEditCol) = txtEdit.Text
    End If
    
    txtEdit.Visible = False
    txtEdit.Text = ""
    
    RaiseEvent AfterEdit(m_lngEditRow, m_lngEditCol)
    
    m_lngEditRow = 0
    m_lngEditCol = 0
End Sub

Public Sub EditBoxOn(ByVal Row As Long, ByVal Col As Long)
Attribute EditBoxOn.VB_Description = "Shows the edit box."
'/// Shows the edit box.
    Dim lngHeight   As Long
    Dim lngWidthCheck As Long
    Dim blnCancel   As Boolean
    
    If txtEdit.Visible Then
        EditBoxOff
    End If
    
    m_lngEditRow = 0
    m_lngEditCol = 0
    
    If Row < 1 Or Row > m_ListCount Then Err.Raise 9, "EListBox", "Row subscript out of range (starts at 1)"
    If Col < 1 Or Col > m_Columns Then Err.Raise 9, "EListBox", "Col subscript out of range (starts at 1)"
    
    RaiseEvent BeforeEdit(Row, Col, blnCancel)
    If blnCancel Then Exit Sub
    
    m_lngEditRow = Row
    m_lngEditCol = Col
    ListIndex = Row 'move the current selected row to the box that is being edited (RF, 16-Aug-2007)
    
    lngHeight = GetRowHeight
    
    'lngWidthCheck + conSpacer \ 2 - IIf(hsbScroll.Visible, hsbScroll, 0)
    lngWidthCheck = IIf(m_CheckBoxes, lngHeight, 0)
    
    Const conSpacer = 10
    
    'reposition the edit box
    'see also DrawRows() which will move text box after this code
    txtEdit.Top = lngHeight * (Row - TopIndex) + 1
    txtEdit.Height = lngHeight
    
    If m_lngTextWidth(Col) > 0 Then
        If Col = 1 Then
            txtEdit.Left = lngWidthCheck + conSpacer \ 2 - IIf(hsbScroll.Visible, hsbScroll, 0)
            txtEdit.Width = m_lngColX(Col) - txtEdit.Left - conSpacer / 2
        ElseIf Col = m_Columns Then 'last column
            txtEdit.Left = m_lngColX(Col - 1) + conSpacer / 2
            txtEdit.Width = picList.ScaleWidth - txtEdit.Left - conSpacer / 2
        ElseIf Col > 1 Then 'somewhere in between
            txtEdit.Left = m_lngColX(Col - 1) + conSpacer / 2
            txtEdit.Width = m_lngColX(Col) - txtEdit.Left - conSpacer / 2
        End If
        
        txtEdit.Tag = ""
    Else
'    If txtEdit.Width < 10 Then 'fill the whole space
        If Col = m_Columns Then 'last column
            txtEdit.Left = m_lngColX(Col - 1) + conSpacer / 2
            txtEdit.Width = picList.ScaleWidth - txtEdit.Left - conSpacer / 2
        Else 'not last column - must choose the whole row to fill up
            txtEdit.Left = lngWidthCheck + conSpacer \ 2 - IIf(hsbScroll.Visible, hsbScroll, 0)
            txtEdit.Width = picList.ScaleWidth - vsbScroll.Width
            txtEdit.Tag = "F"
        End If
    End If
    
    txtEdit.Text = ListSub(Row, Col)
    txtEdit.SelStart = 0
    txtEdit.SelLength = Len(txtEdit.Text)
    txtEdit.ToolTipText = "Row:" & Row & " Column:" & Col
    
    txtEdit.Visible = True
    On Error Resume Next
    txtEdit.SetFocus
End Sub

Public Property Get EditSelStart() As Long
Attribute EditSelStart.VB_Description = "Returns/sets the starting point of text selected."
Attribute EditSelStart.VB_MemberFlags = "400"
'/// Returns/sets the starting point of text selected.
    EditSelStart = txtEdit.SelStart
End Property

Public Property Let EditSelStart(ByVal RHS As Long)
    txtEdit.SelStart = RHS
End Property

Public Property Get EditSelLength() As Long
Attribute EditSelLength.VB_Description = "Returns/sets the number of characters selected."
Attribute EditSelLength.VB_MemberFlags = "400"
'///    Returns/sets the number of characters selected.
    EditSelLength = txtEdit.SelLength
End Property

Public Property Let EditSelLength(ByVal RHS As Long)
    txtEdit.SelLength = RHS
End Property

Public Property Get EditSelText() As String
Attribute EditSelText.VB_Description = "Returns/sets the string containing the currently selected text."
'///  Returns/sets the string containing the currently selected text.
    EditSelText = txtEdit.SelText
End Property

Public Property Let EditSelText(ByVal RHS As String)
    txtEdit.SelText = RHS
End Property

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
    UserControl.Enabled = New_Enabled 'disable the control itself (RF, 16-Aug-2007)
    RefreshListIfNotAdding
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
    AutoSizeColumns False 'after changing font, resize all the columns (RF, 14-Aug-2007)
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
    GetItemsVisible = (picList.ScaleHeight - IIf(hsbScroll.Visible, hsbScroll.Height, 0)) \ GetRowHeight()
    'Debug.Print "GetItemsVisible returns "; GetItemsVisible
End Function

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=13,0,2,
Public Property Get List(ByVal Index As Long) As String
Attribute List.VB_Description = "One-based array list of strings."
Attribute List.VB_ProcData.VB_Invoke_Property = "List;List"
Attribute List.VB_MemberFlags = "200"
'/// One-based array list of strings.
    List = m_ListData(Index).strText
End Property

Public Property Let List(ByVal Index As Long, ByVal New_Text As String)
'    'If Ambient.UserMode = False Then Err.Raise 387
'    m_ListData(Index).strText = New_Text
'
'    'update column size if needs be
'    If m_Column(1) = 0 Then
'        If picList.TextWidth(New_Text) > m_lngTextWidth(1) Then
'            m_lngTextWidth(1) = picList.TextWidth(New_Text)
'        End If
'    End If
'
'    picList.Refresh
'    PropertyChanged "List"
    ListChange Index, New_Text
    RefreshListIfNotAdding
End Property

Friend Function ListToken(ByVal Row As Long) As String
    Dim lngCol  As Long
    Dim strToken As String
    
    strToken = m_ListData(Row).strText
    
    If Len(m_strTokenChar) > 0 Then
        For lngCol = 1 To m_Columns - 1 'other columns
            strToken = strToken & m_strTokenChar & m_ListData(Row).strTextSub(lngCol)
        Next lngCol
    End If
    
    ListToken = strToken
End Function

Friend Sub ListChange(ByVal Index As Long, ByVal New_Text As String)
    m_ListData(Index).strText = New_Text
    
    'update column size if needs be
    If m_Column(1) = 0 Then
        If picList.TextWidth(New_Text) > m_lngTextWidth(1) Then
            m_lngTextWidth(1) = picList.TextWidth(New_Text)
        End If
    End If
    
    PropertyChanged "List"
End Sub

Friend Sub ListChange2(ByVal Index As Long, ByVal New_Text As String)
'Always uses TokenChar.
    AddItem3Assign Index, New_Text, m_strTokenChar
    PropertyChanged "List"
End Sub

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
    
    'If Ambient.UserMode = False Then Err.Raise 387
    
    If pNewValue = m_ListCount Then Exit Property
    If pNewValue > 0 Then
        If pNewValue < &H7FFFFFEF Then
            m_ListUBound = pNewValue + &H10
        Else
            m_ListUBound = &H7FFFFFFF
        End If
        
        ReDim Preserve m_ListData(m_ListUBound)
        For lngItem = m_ListCount + 1 To pNewValue
            'If UBound(m_ListData(lngItem).strTextSub) < LBound(m_ListData(lngItem).strTextSub) Then
            '    ReDim m_ListData(lngItem).strTextSub(m_Columns - 1)
            'Else
            ReDim Preserve m_ListData(lngItem).strTextSub(m_Columns)
            'End If
        Next lngItem
        
        m_ListCount = pNewValue
        UpdateList
        RefreshListIfNotAdding
    ElseIf pNewValue = 0 Then
        Clear
    Else
        Err.Raise 9, , "Subscript out of range.  The list cannot have a negative amount of items."
    End If
    PropertyChanged "ListCount"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=13,0,2,
Public Property Get ListSub(ByVal Row As Long, ByVal Col As Long) As String
Attribute ListSub.VB_Description = "One-based array.  Column 0/1 is the same as List(n)."
Attribute ListSub.VB_ProcData.VB_Invoke_Property = ";List"
Attribute ListSub.VB_MemberFlags = "400"
'/// One-based array.  Column 0/1 is the same as List(n).
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
        
        RefreshListIfNotAdding
    ElseIf Col = 1 Or Col = 0 Then
        m_ListData(Row).strText = New_TextSub
        'update column size if needs be
        If m_Column(1) = 0 Then
            If picList.TextWidth(New_TextSub) > m_lngTextWidth(1) Then
                m_lngTextWidth(1) = picList.TextWidth(New_TextSub)
            End If
        End If
        RefreshListIfNotAdding
    Else
        Err.Raise 9
    End If
    PropertyChanged "ListSub"
End Property


Public Property Let ListIndex(ByVal New_Index As Long)
    If 1 <= New_Index And New_Index <= m_ListCount Then
        
        If Not (vsbScroll * m_lngMultiplier <= New_Index And _
        New_Index <= CLng(vsbScroll) * m_lngMultiplier + CLng(GetItemsVisible()) - CLng(IIf(hsbScroll.Visible, hsbScroll.Height / GetRowHeight(), 1))) Then
            Dim lngTop As Long
            Dim intOld%
            
            'Debug.Print "out of range"
            
            'determine if scroll up or down
            If New_Index > m_ListIndex Then
                lngTop = New_Index - GetItemsVisible() + IIf(hsbScroll.Visible, TextHeight("O") / GetRowHeight(), 1)
            ElseIf New_Index < m_ListIndex Then
                lngTop = New_Index
            Else
                lngTop = New_Index
            End If
            If lngTop > vsbScroll.Max * m_lngMultiplier Then lngTop = vsbScroll.Max * m_lngMultiplier
            If lngTop < 1 Then lngTop = 1
            m_ListIndex = New_Index 'change selected item
            
            intOld = vsbScroll
            'Debug.Print "Moved from "; vsbScroll; " to ";
            If lngTop / m_lngMultiplier < vsbScroll.Min Then
                vsbScroll = 1
            Else
                vsbScroll = lngTop / m_lngMultiplier
            End If
            If vsbScroll = intOld Then
                RefreshListIfNotAdding
            End If
            'Debug.Print vsbScroll
        Else
            'Debug.Print "in range"
            'm_ListIndex = New_Index 'change selected item
            'picList.Refresh
            MoveSelection New_Index
        End If
        
        RaiseEvent Click(New_Index)
    Else
        Err.Raise 9, , "Subscript out of range.  ListIndex starts at 1." 'subscript out of range
    End If
End Property

Public Property Get ListIndex() As Long
Attribute ListIndex.VB_Description = "One-based array.  The ListIndex is between 1 and ListCount."
Attribute ListIndex.VB_ProcData.VB_Invoke_Property = ";List"
Attribute ListIndex.VB_MemberFlags = "400"
'/// One-based array.  The ListIndex is between 1 and ListCount.
    ListIndex = m_ListIndex
End Property

Public Property Get Locked() As Boolean
Attribute Locked.VB_Description = "Gets/sets the Locked property of the edit box."
Attribute Locked.VB_ProcData.VB_Invoke_Property = ";Behavior"
'/// Gets/sets the Locked property of the edit box.
    Locked = txtEdit.Locked
End Property

Public Property Let Locked(ByVal pLocked As Boolean)
    txtEdit.Locked = pLocked
    PropertyChanged "Locked"
End Property

Public Property Get EditMaxLength() As Long
Attribute EditMaxLength.VB_Description = "Returns/sets the maximum number of characters that can be entered in a control."
Attribute EditMaxLength.VB_ProcData.VB_Invoke_Property = ";Behavior"
Attribute EditMaxLength.VB_MemberFlags = "400"
'/// Returns/sets the maximum number of characters that can be entered in a control.
    EditMaxLength = txtEdit.MaxLength
End Property

Public Property Let EditMaxLength(ByVal RHS As Long)
'Requested feature for Design 3 (RF, 15-Aug-2007)
    txtEdit.MaxLength = RHS
    PropertyChanged "EditMaxLength"
End Property

Private Sub MoveSelection(ByVal New_Index As Long)
'draws the new selection position
    Dim lngOld As Long
    
    lngOld = m_ListIndex 'old list index
    m_ListIndex = New_Index 'new list index
    
    If lngOld = 0 Then DrawFocus False 'turn off full list box selector
    
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

'Added in a helper refresh method. (RF, 14-Aug-2007)
Private Sub RefreshListIfNotAdding()
    If Not m_blnAddMode Then picList.Refresh
End Sub

Public Sub RemoveItem(ByVal Index As Long)
Attribute RemoveItem.VB_Description = "One-based array.  Once an item is removed, all items thereafter are re-indexed."
'/// One-based array.  Once an item is removed, all items thereafter are re-indexed.
    Dim lngAfter  As Long
    
    'guard against removing no items (RF, 16-Aug-2007)
    If Index < 0 Or Index > m_ListCount Then
        Err.Raise 9
        Exit Sub
    End If
    
    For lngAfter = Index To m_ListCount - 1
        m_ListData(lngAfter) = m_ListData(lngAfter + 1)
    Next lngAfter
    
    If m_ListCount > 0 Then 'guard against removing no items (RF, 16-Aug-2007)
        m_ListCount = m_ListCount - 1
        If m_ListIndex > m_ListCount Then
            m_ListIndex = m_ListCount
        End If
    End If
    
    m_strSearch = ""
    
    If Not UpdateList() Then
        RefreshAfterRemove Index, Index
    End If
End Sub

Public Sub RemoveRows(ByVal Row As Long, Optional ByVal Rows As Long = 1)
Attribute RemoveRows.VB_Description = "One-based array.  Removes RowSel number of rows starting at Row."
'/// One-based array.  Removes RowSel number of rows starting at Row.
    Dim lngAfter  As Long
    
    If Rows = 0 Then Err.Raise 9, , "RowSel must be positive or negative"
    If Rows < 0 Then Row = Row + Rows: Rows = -Rows 'make positive
    
    If Row + Rows > m_ListCount Then Err.Raise 9, , "Number of rows to remove is out of range because Row + Rows > ListCount"
    
    For lngAfter = Row To m_ListCount - Rows
        m_ListData(lngAfter) = m_ListData(lngAfter + Rows)
    Next lngAfter
    
    m_ListCount = m_ListCount - Rows
    If m_ListIndex > m_ListCount Then 'move selected item
        m_ListIndex = m_ListCount
    End If
    
    m_strSearch = ""
    
    If Not UpdateList() Then
        RefreshAfterRemove Row, Rows
    End If
End Sub

Private Sub RefreshAfterRemove(ByVal FromIndex As Long, ByVal ToIndex As Long)
'Added to make the remove look better (RF, 16-Aug-2007)
'See Also: picList_Paint

    Dim lngFrom As Long 'viewport from
    Dim lngTo As Long   'viewport to
    Dim lngLastY As Long
    
    'we are still changing lots of items in the list so don't redraw now...
    If m_blnAddMode Then Exit Sub
        
    lngFrom = vsbScroll * m_lngMultiplier
    lngTo = CLng(vsbScroll * m_lngMultiplier) + CLng(GetItemsVisible())
        
    If FromIndex > lngTo Then Exit Sub 'no visible items have changed
    If FromIndex > lngFrom Then lngFrom = FromIndex 'don't redraw some items at the top
    
    If lngTo > m_ListCount Then
        lngTo = m_ListCount
        'blank out the bottom
        lngLastY = ((m_ListCount + 1) - vsbScroll * m_lngMultiplier) * GetRowHeight()
        picList.Line (0, lngLastY)-(picList.ScaleWidth, picList.ScaleHeight), picList.BackColor, BF
    End If
    
    'redraw items that were moved
    DrawRows lngFrom, lngTo, True
    
    'draw the focus rect if there are no items
    If m_ListIndex = 0 Then DrawFocus m_blnFocus
End Sub

Public Sub OcxHandshake(ByVal IntNum As Integer)
Attribute OcxHandshake.VB_MemberFlags = "40"
    'Does nothing
End Sub


Public Sub ReadyForResize(ByVal Size As Long)
Attribute ReadyForResize.VB_Description = "Prepares the array for resizing, but ListCount does not change."
'/// Prepares the array for resizing, but ListCount does not change.
    If Size > 0 And m_ListUBound < Size Then
        ReDim Preserve m_ListData(Size)
        m_ListUBound = Size
    End If
End Sub

Public Sub ShowAbout()
Attribute ShowAbout.VB_Description = "Shows the About dialog box."
Attribute ShowAbout.VB_UserMemId = -552
Attribute ShowAbout.VB_MemberFlags = "40"
'/// Shows the About dialog box.
    Load frmAbout
    frmAbout.lblDescription = "EListBox Rel:2 " & App.LegalCopyright
    If txtEdit.Visible Then
        frmAbout.lblDescription = frmAbout.lblDescription & vbNewLine & "Currently Editing: " & txtEdit.ToolTipText
    End If
    frmAbout.Show vbModal
End Sub

'Requested to search for items other than in first column (RF, 14-Aug-2007)
Public Sub SelectRowText(ByVal RowText As String, ByVal Column As Integer, ByVal CaseSensitive As Boolean)
Attribute SelectRowText.VB_Description = "Similar to assigning Text a value to select a row.  This method allows you to select rows from data in any column."
'/// Similar to assigning Text a value to select a row.  This method allows you to select rows from data in any column.
    Dim lngItem As Long
    Dim lngMode As Long  'case sensitive mode
        
    m_strSearch = ""
    
    If CaseSensitive Then
        lngMode = vbBinaryCompare
    Else
        lngMode = vbTextCompare
    End If
    
    If Column = 1 Then
        For lngItem = 1 To m_ListCount
            If StrComp(RowText, m_ListData(lngItem).strText, lngMode) = 0 Then
                ListIndex = lngItem
                Exit Sub
            End If
        Next lngItem
    ElseIf Column > 1 And Column <= m_Columns Then
        For lngItem = 1 To m_ListCount
            If StrComp(RowText, m_ListData(lngItem).strTextSub(Column - 1), lngMode) = 0 Then
                ListIndex = lngItem
                Exit Sub
            End If
        Next lngItem
    Else
        Err.Raise 9, , "Subscript out of range.  The subscript must be between 1 and the number of columns."
    End If
    
    Err.Raise 5, , "Invalid procedure call or argument.  The RowText was not found in the list at the given column."
    
    PropertyChanged "Text"
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

Public Property Get TokenChar() As String
Attribute TokenChar.VB_Description = "Gets/sets token character for loading the List."
Attribute TokenChar.VB_MemberFlags = "400"
'/// Gets/sets token character for loading the List.
    TokenChar = m_strTokenChar
End Property


Public Property Let TokenChar(ByVal pTokenChar As String)
    m_strTokenChar = Left$(pTokenChar, 1)
    PropertyChanged "TokenChar"
End Property

Public Property Get TypeToSearch() As Boolean
Attribute TypeToSearch.VB_Description = "If True then a KeyPress event triggers an automatic searching of the list for a matching string. True is the default setting.   If False then a custom handler for key events should be implemented."
Attribute TypeToSearch.VB_ProcData.VB_Invoke_Property = ";Behavior"
'/// If True then a KeyPress event triggers an automatic searching of the list for a matching string. True is the default setting.
'/// If False then a custom handler for key events should be implemented.
    TypeToSearch = m_blnTypeToSearch
End Property

Public Property Let TypeToSearch(ByVal RHS As Boolean)
    m_blnTypeToSearch = RHS
    m_strSearch = ""
    PropertyChanged "TypeToSearch"
End Property

Private Function UpdateList() As Boolean
'This will return True when the current scroll value changes.
    Dim lngMax As Long
    Dim lngItemsVisible As Long
    
    lngItemsVisible = GetItemsVisible()
    
    lngMax = m_ListCount - lngItemsVisible
    'Debug.Print "UpdateList lngMax is "; lngMax
    
    If lngMax <= 1 Then
        lngMax = 1
        m_lngMultiplier = 1
    Else
        m_lngMultiplier = lngMax \ &H3FFF 'half of integer maximum
        If m_lngMultiplier < 1 Then m_lngMultiplier = 1
        lngMax = lngMax + IIf(m_lngMultiplier <= lngItemsVisible, m_lngMultiplier, lngItemsVisible)
    End If
    
    vsbScroll.Min = 1
    If vsbScroll * m_lngMultiplier > lngMax Then
        vsbScroll.Value = lngMax \ m_lngMultiplier
        UpdateList = True
    End If
    vsbScroll.Max = lngMax \ m_lngMultiplier
    
    vsbScroll.SmallChange = 1
    If lngItemsVisible \ m_lngMultiplier > 0 Then
        vsbScroll.LargeChange = lngItemsVisible \ m_lngMultiplier
    Else
        vsbScroll.LargeChange = 1
    End If
    vsbScroll.Visible = vsbScroll.Min <> vsbScroll.Max
    vsbScroll.Height = IIf(hsbScroll.Visible, picList.ScaleHeight - hsbScroll.Height, picList.ScaleHeight)
    hsbScroll.Move 0, picList.ScaleHeight - hsbScroll.Height, picList.ScaleWidth - IIf(vsbScroll.Visible, vsbScroll.Width, 0)
    'show gray corner and vertical scroll
    picGray.Visible = vsbScroll.Visible And hsbScroll.Visible
End Function


'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=vsbScroll,vsbScroll,-1,LargeChange
Public Property Get LargeChange() As Integer
Attribute LargeChange.VB_Description = "Returns amount of change to Value property in a scroll bar when user clicks the scroll bar area."
Attribute LargeChange.VB_ProcData.VB_Invoke_Property = ";Behavior"
Attribute LargeChange.VB_MemberFlags = "400"
'/// Returns amount of change to Value property in a scroll bar when user clicks the scroll bar area.
    LargeChange = vsbScroll.LargeChange
End Property

Public Property Let LargeChange(ByVal New_LargeChange As Integer)
    vsbScroll.LargeChange() = New_LargeChange
    PropertyChanged "LargeChange"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=vsbScroll,vsbScroll,-1,SmallChange
Public Property Get SmallChange() As Integer
Attribute SmallChange.VB_Description = "Returns amount of change to Value property in a scroll bar when user clicks a scroll arrow."
Attribute SmallChange.VB_ProcData.VB_Invoke_Property = ";Behavior"
Attribute SmallChange.VB_MemberFlags = "400"
'/// Returns amount of change to Value property in a scroll bar when user clicks a scroll arrow.
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

Public Property Get RowData(ByVal Row As Long) As Variant
Attribute RowData.VB_Description = "Gets/sets any user-defined data for the Row in the list."
Attribute RowData.VB_MemberFlags = "400"
'/// Gets/sets any user-defined data for the Row in the list.
'Add your code here.
    If IsObject(m_ListData(Row).vntData) Then
        Set RowData = m_ListData(Row).vntData
    Else
        RowData = m_ListData(Row).vntData
    End If
End Property


Public Property Let RowData(ByVal Row As Long, ByVal pRowData As Variant)
'UNDONE: Change Variant with the default data type
'Add your code here.
    m_ListData(Row).vntData = pRowData
End Property


Public Property Set RowData(ByVal Row As Long, ByVal pRowData As Variant)
'Add your code here.
    Set m_ListData(Row).vntData = pRowData
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

Public Property Get ColumnRealWidth(ByVal Column As Integer) As Single
Attribute ColumnRealWidth.VB_Description = "Returns the actual column width if the default column width is not specified."
Attribute ColumnRealWidth.VB_MemberFlags = "400"
'/// Returns the actual column width if the default column width is not specified.
    If m_Column(Column) = 0! Then
        ColumnRealWidth = m_lngTextWidth(Column)
    Else
        ColumnRealWidth = m_Column(Column)
    End If
End Property


'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=vsbScroll,vsbScroll,-1,Value
Public Property Get TopIndex() As Integer
Attribute TopIndex.VB_Description = "Returns/sets the top visible item of a list box.  One-based array."
Attribute TopIndex.VB_MemberFlags = "400"
'/// Returns/sets the top visible item of a list box.  One-based array.
    TopIndex = vsbScroll.Value
End Property

Public Property Let TopIndex(ByVal New_TopIndex As Integer)
    If Ambient.UserMode = False Then Err.Raise 387
    If New_TopIndex >= vsbScroll.Max * m_lngMultiplier Then
        vsbScroll = vsbScroll.Max \ m_lngMultiplier
    ElseIf New_TopIndex <= vsbScroll.Min * m_lngMultiplier Then
        vsbScroll = vsbScroll.Min \ m_lngMultiplier
    Else
        vsbScroll = New_TopIndex \ m_lngMultiplier
    End If
    PropertyChanged "TopIndex"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,2,0
Public Property Get Selected(ByVal Index As Long) As Boolean
Attribute Selected.VB_Description = "One-based array.  Determines if an item is checked."
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
Attribute ItemData.VB_MemberFlags = "400"
'/// One-based array.  Data is not shown to the user.
    ItemData = m_ListData(Index).lngItemData
End Property

Public Property Let ItemData(ByVal Index As Long, ByVal New_ItemData As Long)
    'If Ambient.UserMode = False Then Err.Raise 387
    m_ListData(Index).lngItemData = New_ItemData
    PropertyChanged "ItemData"
End Property


'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,0
Public Property Get RowDivisions() As Boolean
Attribute RowDivisions.VB_Description = "Shows/hides the row division line."
Attribute RowDivisions.VB_ProcData.VB_Invoke_Property = ";Appearance"
'/// Shows/hides the row division line.
    RowDivisions = m_RowDivisions
End Property

Public Property Let RowDivisions(ByVal New_RowDivisions As Boolean)
    m_RowDivisions = New_RowDivisions
    PropertyChanged "RowDivisions"
End Property


Private Sub HsbScroll_Change()
    RefreshListIfNotAdding
End Sub

Private Sub HsbScroll_Scroll()
    RefreshListIfNotAdding
End Sub

Private Sub picGray_Click()
    RaiseEvent CornerClick
End Sub

Private Sub picGray_DblClick()
    RaiseEvent CornerDblClick
End Sub

Private Sub picGray_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent CornerMouseDown(Button, Shift, picGray.Left + X, picGray.Top + Y)
End Sub

Private Sub picGray_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent CornerMouseMove(Button, Shift, picGray.Left + X, picGray.Top + Y)
End Sub

Private Sub picGray_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent CornerMouseUp(Button, Shift, picGray.Left + X, picGray.Top + Y)
End Sub

Private Sub picList_Click()
    m_strSearch = ""
    If m_ListIndex > 0 And Not m_blnClickOnWhiteSpace Then RaiseEvent Click(m_ListIndex)
    EditBoxOff
    
    RaiseEvent RowColClick(m_lngClickedOnRow, m_lngClickedOnCol)
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
    RaiseEvent RowColDblClick(m_lngClickedOnRow, m_lngClickedOnCol)
End Sub

Private Sub picList_KeyPress(KeyAscii As Integer)
    Static blnCancel As Boolean
    
    RaiseEvent KeyPress(KeyAscii)
    
    If KeyAscii = vbKeySpace And m_CheckBoxes And m_ListIndex > 0 Then
        'the space bar will check and uncheck a list box
        
        'the key intercept will now ask before changing the state (RF, 18 Aug 2007)
        blnCancel = False
        RaiseEvent ItemChecked(m_ListIndex, Not m_ListData(m_ListIndex).blnSelected, blnCancel)
        If Not blnCancel Then
            'change the check box state and redraw
            m_ListData(m_ListIndex).blnSelected = Not m_ListData(m_ListIndex).blnSelected
            MoveSelection m_ListIndex
        End If
        
        'reset search
        m_strSearch = ""
    ElseIf m_blnTypeToSearch And ((KeyAscii >= &H20 And KeyAscii <= &H7E) Or (KeyAscii > &H80)) Then 'from space character to tilda ... visible text
        Dim i As Long, L As Long
        'we want to type to search for some text and this feature is enabled (via TypeToSearch)
        
        m_strSearch = m_strSearch & Chr$(KeyAscii)
        L = Len(m_strSearch)
        
        For i = 1 To m_ListCount
            If StrComp(m_strSearch, Left$(m_ListData(i).strText, L), vbTextCompare) = 0 Then
                ListIndex = i
                Exit For
            End If
        Next i
    ElseIf m_blnTypeToSearch Then
        'reset search
        m_strSearch = ""
    End If
End Sub

Private Sub picList_Paint()
'see also: RefreshAfterRemove
    UpdateList
    
    picList.Cls
    m_blnFullRectShown = False

    If m_ListCount > 0 Then
        Dim lngTo As Long
        
        lngTo = CLng(vsbScroll * m_lngMultiplier) + CLng(GetItemsVisible())
        If lngTo > m_ListCount Then lngTo = m_ListCount
        DrawRows vsbScroll * m_lngMultiplier, lngTo, False
    Else
        'still draw the gridlines (RF, 16-Aug-2007)
        DrawRows 1, 1, False
        
        If Ambient.UserMode = False Then
            'always draw in the corner
            picList.CurrentX = 0
            picList.CurrentY = 0
            picList.Print Ambient.DisplayName
        End If
    End If
    
    If m_ListIndex = 0 Then DrawFocus m_blnFocus
End Sub

Private Sub txtEdit_KeyDown(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyDownEdit(KeyCode, Shift)
End Sub

Private Sub txtEdit_KeyPress(KeyAscii As Integer)
    RaiseEvent KeyPressEdit(KeyAscii)
End Sub

Private Sub txtEdit_KeyUp(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyUpEdit(KeyCode, Shift)
    
    Select Case KeyCode
    Case vbKeyPageDown, vbKeyPageUp
        EditBoxOff
        picList_KeyDown KeyCode, Shift
    Case vbKeyUp, vbKeyDown
        EditBoxOff
        picList_KeyDown KeyCode, Shift
    Case vbKeyReturn
        EditBoxOff
    End Select
End Sub

Private Sub txtEdit_LostFocus()
    EditBoxOff
End Sub

Private Sub UserControl_AmbientChanged(PropertyName As String)
    RefreshListIfNotAdding
End Sub

Private Sub UserControl_EnterFocus()
    m_blnFocus = True
    If m_ListIndex > 0 Then 'if -1 then no item is selected; > 0 means a real item is selected
        DrawRows m_ListIndex, m_ListIndex, True
    Else
        DrawFocus True
    End If
End Sub

Private Sub UserControl_ExitFocus()
    m_blnFocus = False
    m_strSearch = ""
    EditBoxOff
    If m_ListIndex Then
        DrawRows m_ListIndex, m_ListIndex, True
    Else
        DrawFocus False
    End If
End Sub


Private Sub UserControl_Initialize()
    Set m_Column = New ESingleArray
    m_lngMultiplier = 1
    ReadyForResize 10
End Sub

Private Sub picList_KeyDown(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyDown(KeyCode, Shift)
        
    'pass to others first to override KeyCode commands
    Select Case KeyCode
    Case vbKeyDown
        If m_ListIndex < m_ListCount Then
            ListIndex = m_ListIndex + 1
        End If
        'reset search (RF, 18-Aug-2007)
        m_strSearch = ""
        
    Case vbKeyUp
        If m_ListIndex > 1 Then
            ListIndex = m_ListIndex - 1
        End If
        'reset search (RF, 18-Aug-2007)
        m_strSearch = ""
    
    Case vbKeyPageDown
        If CLng(vsbScroll) + CLng(vsbScroll.LargeChange) <= CLng(vsbScroll.Max) Then
            vsbScroll = vsbScroll + vsbScroll.LargeChange
        Else
            vsbScroll = vsbScroll.Max
        End If
        'reset search (RF, 18-Aug-2007)
        m_strSearch = ""
    
    Case vbKeyPageUp
        If vsbScroll - vsbScroll.LargeChange >= vsbScroll.Min Then
            vsbScroll = vsbScroll - vsbScroll.LargeChange
        Else
            vsbScroll = vsbScroll.Min
        End If
        'reset search (RF, 18-Aug-2007)
        m_strSearch = ""
    
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
    
    
End Sub


Private Sub picList_KeyUp(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyUp(KeyCode, Shift)
End Sub

Private Sub picList_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim lngItem As Long
    Dim lngCols As Long
    Dim lngCurCol As Long
    
    lngItem = Y \ GetRowHeight() + vsbScroll * m_lngMultiplier
    If 0 < lngItem And lngItem <= m_ListCount Then
        MoveSelection lngItem
        For lngCols = 1 To m_Columns
            If X < m_lngColX(lngCols) Then lngCurCol = lngCols: Exit For
        Next lngCols
        
        If lngCurCol = 0 Then lngCurCol = m_Columns 'choose last one if no choice available
        
        m_lngClickedOnRow = lngItem
        m_lngClickedOnCol = lngCurCol
        
        m_blnClickOnWhiteSpace = False
    Else
        m_lngClickedOnRow = 0
        m_lngClickedOnCol = 0
        
        m_blnClickOnWhiteSpace = True
    End If
    
    On Error GoTo Handler
    'RaiseEvent MouseDown(Button, Shift, ScaleX(X, ScaleMode, vbContainerPosition), ScaleY(Y, ScaleMode, vbContainerPosition))
    RaiseEvent MouseDown(Button, Shift, X, Y)
    Exit Sub
Handler:
    'site does not contain a scale mode
    RaiseEvent MouseDown(Button, Shift, X, Y)
    Exit Sub
End Sub

Private Sub picList_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error GoTo Handler
    RaiseEvent MouseMove(Button, Shift, X, Y) ' ScaleX(X, ScaleMode, vbContainerPosition), ScaleY(Y, ScaleMode, vbContainerPosition))
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
    lngItem = Y \ GetRowHeight() + vsbScroll * m_lngMultiplier
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
    RaiseEvent MouseUp(Button, Shift, X, Y) 'ScaleX(X, ScaleMode, vbContainerPosition), ScaleY(Y, ScaleMode, vbContainerPosition))
    Exit Sub
Handler:
    'site does not contain a scale mode
    RaiseEvent MouseUp(Button, Shift, X, Y)
    Exit Sub
End Sub

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
    m_blnTypeToSearch = m_def_TypeToSearch 'feature (RF, 18-Aug-2007)
End Sub

Private Sub UserControl_LostFocus()
    EditBoxOff
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    Dim objList As ETextArray ': Set objList = New ETextArray
    Dim objItemData As ELongArray ': Set objItemData = New ELongArray
    Dim intItem As Long

    Set m_Column = New ESingleArray
    
    m_blnAddMode = True 'don't repaint all the time (RF, 14-Aug-2007)

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
    txtEdit.Locked = PropBag.ReadProperty("Locked", False)
    txtEdit.MaxLength = PropBag.ReadProperty("EditMaxLength", 0) 'requested feature (RF, 15-Aug-2007)
    m_strTokenChar = PropBag.ReadProperty("TokenChar", "")
    m_blnTypeToSearch = PropBag.ReadProperty("TypeToSearch", m_def_TypeToSearch) 'feature (RF, 18-Aug-2007)
    
    On Error GoTo Handler2
    'columns
    m_Columns = PropBag.ReadProperty("Columns", m_def_Columns)
    ReDim m_lngTextWidth(m_Columns)
    'm_Column.Count = m_Columns
    Set m_Column = PropBag.ReadProperty("Column")
    
    'rows
    ListCount = PropBag.ReadProperty("ListCount", 0)
    Set objList = PropBag.ReadProperty("List")
    Set objItemData = PropBag.ReadProperty("ItemData")
    
    If Len(m_strTokenChar) = 0 Then
        If Not objList Is Nothing Then
            For intItem = 1 To m_ListCount
                If intItem > objList.Count Then Exit For
                'AddItem objList(intItem)
                ListChange intItem, objList(intItem)
            Next intItem
        End If
    Else
        If Not objList Is Nothing Then
            For intItem = 1 To m_ListCount
                If intItem > objList.Count Then Exit For
                'AddItem objList(intItem)
                ListChange2 intItem, objList(intItem)
            Next intItem
        End If
    End If
        
    If Not objItemData Is Nothing Then
        For intItem = 1 To m_ListCount
            If intItem > objItemData.Count Then Exit For
            m_ListData(intItem).lngItemData = objItemData(intItem)
        Next intItem
    End If
    
    m_blnAddMode = False 'don't repaint all the time (RF, 14-Aug-2007)
    
    Exit Sub
Handler:
    Debug.Print "ReadProperties : "; Err.Description
    Resume Next
Handler2:
    Debug.Print "ReadProperties in Row Col data: "; Err.Description
    m_blnAddMode = False 'don't repaint all the time (RF, 14-Aug-2007)
    Exit Sub
End Sub

Private Sub UserControl_Resize()
    On Error GoTo Catch
    picList.Move 0, 0, ScaleWidth, ScaleHeight
    vsbScroll.Move picList.ScaleWidth - vsbScroll.Width, 0, vsbScroll.Width, IIf(hsbScroll.Visible, picList.ScaleHeight - hsbScroll.Height, picList.ScaleHeight)
    hsbScroll.Move 0, picList.ScaleHeight - hsbScroll.Height, picList.ScaleWidth - IIf(vsbScroll.Visible, vsbScroll.Width, 0)
    picGray.Move vsbScroll.Left, hsbScroll.Top, vsbScroll.Width, hsbScroll.Height
    
    UpdateList
    RefreshListIfNotAdding
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
    Call PropBag.WriteProperty("Locked", txtEdit.Locked, False)
    Call PropBag.WriteProperty("EditMaxLength", txtEdit.MaxLength, 0) 'requested feature (RF, 15-Aug-2007)
    Call PropBag.WriteProperty("TokenChar", m_strTokenChar, "")
    Call PropBag.WriteProperty("TypeToSearch", m_blnTypeToSearch, m_def_TypeToSearch) 'feature (RF, 18-Aug-2007)
    
    Call PropBag.WriteProperty("Columns", m_Columns, m_def_Columns)
    Call PropBag.WriteProperty("Column", m_Column)
    
    On Error GoTo Handler2
    
    'list data - only first column
    Dim objList As ETextArray: Set objList = New ETextArray
    Dim objItemData As ELongArray: Set objItemData = New ELongArray
    Dim intItem As Long
    
    Call PropBag.WriteProperty("ListCount", m_ListCount)
    
    objList.Count = m_ListCount
    objItemData.Count = m_ListCount
    
    If Len(m_strTokenChar) = 0 Then
        For intItem = 1 To m_ListCount
            objList(intItem) = m_ListData(intItem).strText
            objItemData(intItem) = m_ListData(intItem).lngItemData
        Next intItem
    Else 'convert columns to tokenized string
        For intItem = 1 To m_ListCount
            objList(intItem) = ListToken(intItem)
            objItemData(intItem) = m_ListData(intItem).lngItemData
        Next intItem
    End If
    
    Call PropBag.WriteProperty("List", objList)
    Call PropBag.WriteProperty("ItemData", objItemData)
    
    Exit Sub
Handler:
    Resume Next
Handler2:
    Exit Sub
End Sub



Private Sub vsbScroll_Change()
    RefreshListIfNotAdding
    If txtEdit.Visible Then EditBoxOff
End Sub

Private Sub vsbScroll_Scroll()
    RefreshListIfNotAdding
    If txtEdit.Visible Then EditBoxOff
End Sub

