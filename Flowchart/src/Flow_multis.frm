VERSION 5.00
Begin VB.Form frmIMulti 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Multiple Selection"
   ClientHeight    =   300
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4200
   Icon            =   "Flow_multis.frx":0000
   KeyPreview      =   -1  'True
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   300
   ScaleWidth      =   4200
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdTools 
      Caption         =   "Tools"
      Height          =   300
      Left            =   3360
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   0
      Width           =   840
   End
   Begin VB.CommandButton cmdCopy 
      Caption         =   "Copy"
      Height          =   300
      Left            =   2520
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   0
      Width           =   840
   End
   Begin VB.CommandButton cmdAlign 
      Caption         =   "Align"
      Height          =   300
      Left            =   1680
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   0
      Width           =   840
   End
   Begin VB.CommandButton cmdSize 
      Caption         =   "Same Size"
      Height          =   300
      Left            =   600
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   1080
   End
   Begin VB.Label lblStatus 
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   525
   End
   Begin VB.Menu mnuSize 
      Caption         =   ""
      Visible         =   0   'False
      Begin VB.Menu mnuSizeW 
         Caption         =   "Width"
      End
      Begin VB.Menu mnuSizeH 
         Caption         =   "Height"
      End
      Begin VB.Menu mnuSizeWH 
         Caption         =   "Both"
      End
      Begin VB.Menu mnuSizeSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSizeCol 
         Caption         =   "&Columns"
      End
      Begin VB.Menu mnuSizeRows 
         Caption         =   "&Rows"
      End
   End
   Begin VB.Menu mnuAlign 
      Caption         =   ""
      Visible         =   0   'False
      Begin VB.Menu mnuAlignL 
         Caption         =   "Left"
      End
      Begin VB.Menu mnuAlignC 
         Caption         =   "Centre"
      End
      Begin VB.Menu mnuAlignR 
         Caption         =   "Right"
      End
      Begin VB.Menu mnuAlignSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAlignT 
         Caption         =   "Top"
      End
      Begin VB.Menu mnuAlignM 
         Caption         =   "Middle"
      End
      Begin VB.Menu mnuAlignB 
         Caption         =   "Bottom"
      End
      Begin VB.Menu mnuAlignSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAlignGrid 
         Caption         =   "To Grid"
      End
   End
   Begin VB.Menu mnuCopy 
      Caption         =   ""
      Visible         =   0   'False
      Begin VB.Menu mnuCopyF 
         Caption         =   "Font"
      End
      Begin VB.Menu mnuCopyA 
         Caption         =   "Apperance"
      End
   End
   Begin VB.Menu mnuTools 
      Caption         =   ""
      Visible         =   0   'False
      Begin VB.Menu mnuToolsDuplicate 
         Caption         =   "Duplicate"
      End
      Begin VB.Menu mnuToolsDelete 
         Caption         =   "Delete..."
      End
      Begin VB.Menu mnuToolsSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuToolsGroup 
         Caption         =   "Group"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuToolsUngroup 
         Caption         =   "Ungroup"
         Enabled         =   0   'False
      End
   End
End
Attribute VB_Name = "frmIMulti"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'1 February 2002.

Public mCollection  As Collection
Private mSelected   As FlowItem
Private mblnSizing  As Boolean
Private mblnMoving  As Boolean
Private mRect       As Rect 'for sizing and moving

Public Sub AddItem(ByVal Ptr As FlowItem)
    If Not Ptr Is Nothing Then
        mCollection.Add Ptr
        Ptr.P.Selected = True
        ItemSelected
        Update
    End If
End Sub

Public Sub AddItem2(ByVal Ptr As FlowItem)
    If Not Ptr Is Nothing Then
    'Adds but does not Update
        mCollection.Add Ptr
        Ptr.P.Selected = True
    End If
End Sub

Public Sub BeginMove()
    mblnMoving = True
    
    frmMain.ToolBoxChanged conUndoMultiple
    
    mRect.X1 = mSelected.P.Left
    mRect.Y1 = mSelected.P.Top
    Update
End Sub

Public Sub BeginSize()
    mblnSizing = True
    
    frmMain.ToolBoxChanged conUndoMultiple
    
    mRect.X2 = mSelected.P.Width
    mRect.Y2 = mSelected.P.Height
    Update
End Sub

Public Sub Clear()
    Dim objItem As FlowItem
    
    If Not mCollection Is Nothing Then
        ItemUnSelected
        
        For Each objItem In mCollection
            If Not objItem Is Nothing Then
                objItem.P.Selected = False
            End If
        Next objItem
    End If
    
    Set mCollection = New Collection
    Set mSelected = Nothing
    Update
End Sub

Public Sub DoAlign()
    mnuAlignGrid_Click
End Sub

Public Sub DoCopy()
    Set frmMain.mClipboard.Selection = mCollection
End Sub

Public Sub DoCopyAppearance(ByVal FillStyle As Boolean, ByVal LineStyle As Boolean, ByVal LineWidth As Boolean, _
ByVal ArrowEngg As Boolean, ByVal ArrowSize As Boolean, ByVal BackColour As Boolean, ByVal ForeColour As Boolean, ByVal TextColour As Boolean)
    Dim objItem As FlowItem
    Dim objProp As Properties
    
    If IsSelected Then
       frmMain.ToolBoxChanged conUndoMultiple
        
        Set objProp = New Properties
        With mSelected.P
            objProp.FillStyle = .FillStyle
            objProp.LineStyle = .LineStyle
            objProp.LineWidth = .LineWidth
            objProp.ArrowEngg = .ArrowEngg
            objProp.ArrowSize = .ArrowSize
            'objProp.ArrowSolid = .ArrowSolid
            objProp.BackColour = .BackColour
            objProp.ForeColour = .ForeColour
            objProp.TextColour = .TextColour
        End With
        
        For Each objItem In mCollection
            If Not objItem Is Nothing Then
                With objItem.P
                    If FillStyle Then .FillStyle = objProp.FillStyle
                    If LineStyle Then .LineStyle = objProp.LineStyle
                    If LineWidth Then .LineWidth = objProp.LineWidth
                    If ArrowEngg Then .ArrowEngg = objProp.ArrowEngg
                    If ArrowSize Then .ArrowSize = objProp.ArrowSize
                    
                    If BackColour Then .BackColour = objProp.BackColour
                    If ForeColour Then .ForeColour = objProp.ForeColour
                    If TextColour Then .TextColour = objProp.TextColour
                    '.ArrowSolid = objProp.ArrowSolid
                End With
            End If
        Next objItem
        ItemChanged
    End If
    Set objProp = Nothing


End Sub

Public Sub DoCopyFont(ByVal FontFace As Boolean, ByVal TextAlign As Boolean, _
ByVal TextBold As Boolean, ByVal TextItalic As Boolean, ByVal TextUnderline As Boolean, ByVal TextSize As Boolean)
    Dim objItem As FlowItem
    Dim objProp As Properties
    
    If IsSelected Then
        frmMain.ToolBoxChanged conUndoMultiple
    
        Set objProp = New Properties
        'get properties to copy
        With mSelected.P
            objProp.FontFace = .FontFace
            objProp.TextAlign = .TextAlign
            objProp.TextBold = .TextBold
            'objProp.TextColour = .TextColour
            objProp.TextItalic = .TextItalic
            objProp.TextUnderline = .TextUnderline
            objProp.TextSize = .TextSize
        End With
        
        For Each objItem In mCollection
            If Not objItem Is Nothing Then
                With objItem.P
                    If FontFace Then .FontFace = objProp.FontFace
                    If TextAlign Then .TextAlign = objProp.TextAlign
                    If TextBold Then .TextBold = objProp.TextBold
                    'If TextColour Then .TextColour = objProp.TextColour
                    If TextItalic Then .TextItalic = objProp.TextItalic
                    If TextUnderline Then .TextUnderline = objProp.TextUnderline
                    If TextSize Then .TextSize = objProp.TextSize
                End With
            End If
        Next objItem
        ItemChanged
    End If
    Set objProp = Nothing

End Sub

Public Sub DoCut()
    Dim colPrevious As Collection
    Dim objItem     As FlowItem

    Set frmMain.mClipboard.Selection = mCollection
    'copied from mnuToolsDelete_Click()
    Set colPrevious = mCollection
    Clear 'selection
    Set frmMain.mSelected = Nothing

    For Each objItem In colPrevious
        If Not objItem Is Nothing Then mFlowChart.RemoveObj objItem
    Next objItem
    Update
    ItemChanged
End Sub

Public Sub DoDuplicate()
    mnuToolsDuplicate_Click
End Sub

Public Sub DoPaste()
    Dim objItem As FlowItem
    Dim objPaste As FlowItem

    Clear
    For Each objItem In frmMain.mClipboard.Selection
        If Not objItem Is Nothing Then
            Set objPaste = mFlowChart.Add(Duplicate(objItem))
            AddItem2 objPaste
            CheckObjBounds objPaste, frmMain.picBay.ScaleWidth, frmMain.picBay.ScaleHeight
        End If
    Next objItem
    
    mnuToolsGroup_Click
    
    Set frmMain.mSelected = objPaste
    Set objPaste = Nothing

    Update
    ItemChanged
End Sub

Public Sub EndMove(ByVal Cancel As Boolean)
    Dim objItem As FlowItem
    Dim sngValue1 As Single
    Dim sngValue2 As Single
    
    If mblnMoving And Not Cancel Then
        sngValue1 = mSelected.P.Left - mRect.X1
        sngValue2 = mSelected.P.Top - mRect.Y1
        For Each objItem In mCollection
            If Not objItem Is mSelected And Not objItem Is Nothing Then
                objItem.P.Left = objItem.P.Left + sngValue1
                objItem.P.Top = objItem.P.Top + sngValue2
                CheckObjBounds objItem, frmMain.picBay.ScaleWidth, frmMain.picBay.ScaleHeight
            End If
        Next objItem
        ItemChanged
    End If
    mblnMoving = False
    Update
End Sub

Public Sub EndSize(ByVal Cancel As Boolean)
    Dim objItem As FlowItem
    Dim sngValue1 As Single
    Dim sngValue2 As Single
    
    
    If mblnSizing And Not Cancel And mRect.X2 > 0 And mRect.Y2 > 0 Then
        If mSelected.P.Width <> mRect.X2 Then
            sngValue1 = mSelected.P.Width / mRect.X2
        Else
            sngValue1 = 1
        End If
        If mSelected.P.Height <> mRect.Y2 Then
            sngValue2 = mSelected.P.Height / mRect.Y2
        Else
            sngValue2 = 1
        End If
        'Debug.Print sngValue1, sngValue2
        For Each objItem In mCollection
            If Not objItem Is mSelected And Not objItem Is Nothing Then
                objItem.P.Width = objItem.P.Width * sngValue1
                objItem.P.Height = objItem.P.Height * sngValue2
                If objItem.Number = conAddPicture Then
                    GetPicture(objItem).SetCustomSize
                End If
            End If
        Next objItem
        ItemChanged
    End If
    mblnSizing = False
    
    Update
End Sub

Public Function GetCount() As Long
    GetCount = mCollection.Count
End Function

Public Function GetIndex(ByVal Ptr As FlowItem) As Long
    Static lngLastIndex As Long
    Dim lngIndex As Long
    Dim objItem  As FlowItem
    
    If lngLastIndex <> 0 Then 'default value
        If 1 <= lngLastIndex And lngLastIndex + 1 <= mCollection.Count Then
            'quick search by increment of last index
            If mCollection(lngLastIndex + 1) Is Ptr Then
                GetIndex = lngLastIndex + 1
                lngLastIndex = lngIndex
                'Debug.Print "Quick search successful"
                Exit Function
            End If
        Else
            lngLastIndex = 0 'at the end, go back to beginning
        End If
    End If
    
    For lngIndex = 1 To mCollection.Count
        Set objItem = mCollection(lngIndex)
        If objItem Is Ptr And Not objItem Is Nothing Then
            GetIndex = lngIndex
            lngLastIndex = lngIndex
            Exit Function
        End If
    Next lngIndex
End Function

Public Sub HoldItem(ByVal Ptr As FlowItem)
    If GetIndex(Ptr) = 0 Then
        AddItem Ptr
    End If
End Sub

Private Function IsSelected() As Boolean
    IsSelected = Not mSelected Is Nothing
End Function

Private Sub ItemChanged()
    'the item in this window has changed
    frmMain.Redraw
End Sub

Private Sub ItemSelected()
    frmMain.RedrawSelection True
End Sub

Private Sub ItemUnSelected()
    frmMain.RedrawSelection False
End Sub

Public Property Get mFlowChart() As FlowChart
    Set mFlowChart = frmMain.mFlowChart
End Property

Public Sub AddGroup(GroupNo As Long)
    Dim objItem As FlowItem
    Dim blnAdded As Boolean
    
    For Each objItem In mFlowChart.Groups(GroupNo)
        If GetIndex(objItem) = 0 Then
            AddItem2 objItem
            blnAdded = True
        End If
    Next objItem
    
    If blnAdded Then
        ItemSelected
        Update
    End If
End Sub

Public Sub RemoveGroup(GroupNo As Long)
    Dim lngItem As Long
    Dim objItem As FlowItem
    Dim blnRemoved As Boolean
    
    ItemUnSelected
    'remove every group item
    For lngItem = mCollection.Count To 1 Step -1
        Set objItem = mCollection(lngItem)
        If objItem.P.GroupNo = GroupNo Then
            objItem.P.Selected = False
            mCollection.Remove lngItem 'remove this
            blnRemoved = True
        End If
    Next lngItem
    
    ItemSelected
    
    If blnRemoved Then
        SelectLast
        Update
    End If
    
    Set objItem = Nothing
End Sub

Public Sub RemoveItem(ByVal Ptr As FlowItem)
    Dim lngI As Long
    
    ItemUnSelected
    
    lngI = GetIndex(Ptr)
    
    If lngI > 0 Then
        Ptr.P.Selected = False
        mCollection.Remove lngI
    End If
    
    If Ptr Is mSelected Then
        Set mSelected = Nothing
    End If
    
    If lngI > 0 Then
        SelectLast
        Update
    End If
    ItemSelected
End Sub

Private Sub SelectLast()
    If Not mCollection Is Nothing Then
        If mCollection.Count > 0 Then
            Set frmMain.mSelected = mCollection(mCollection.Count)
        End If
    End If
End Sub

Public Sub SwitchItemGroup(ByVal Ptr As FlowItem)
    'adds or removes an item/group
    If GetIndex(Ptr) = 0 Then
        If Ptr.P.GroupNo <> 0 Then
            AddGroup Ptr.P.GroupNo
        Else
            AddItem Ptr
        End If
    Else
        If Ptr.P.GroupNo = 0 Then
            RemoveItem Ptr
        Else
            RemoveGroup Ptr.P.GroupNo
        End If
    End If
End Sub

Public Sub Update()
    Dim blnEnabled As Boolean
    
    'update from main form
    Set mSelected = frmMain.mSelected
    
    If IsSelected And mCollection.Count > 0 Then
        blnEnabled = Not mblnMoving And Not mblnSizing
        If mblnMoving Then
            lblStatus = "Moving..."
        ElseIf mblnSizing Then
            lblStatus = "Sizing..."
        Else
            lblStatus = mCollection.Count & " items"
        End If
    Else
        lblStatus = mCollection.Count & " items"
    End If
    cmdSize.Enabled = blnEnabled
    cmdAlign.Enabled = blnEnabled
    cmdCopy.Enabled = blnEnabled
    cmdTools.Enabled = blnEnabled
End Sub


Private Sub cmdAlign_Click()
    PopupMenu mnuAlign, vbPopupMenuLeftAlign, cmdAlign.Left, cmdAlign.Top + cmdAlign.Height
End Sub

Private Sub cmdCopy_Click()
    PopupMenu mnuCopy, vbPopupMenuLeftAlign, cmdCopy.Left, cmdCopy.Top + cmdCopy.Height
End Sub

Private Sub cmdSize_Click()
    PopupMenu mnuSize, vbPopupMenuLeftAlign, cmdSize.Left, cmdSize.Top + cmdSize.Height
End Sub

Private Sub cmdTools_Click()
    PopupMenu mnuTools, vbPopupMenuLeftAlign, cmdTools.Left, cmdTools.Top + cmdTools.Height
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case vbKeyLeft, vbKeyRight
        'nothing to do
    Case Else
        frmMain.SetFocus
    End Select
    KeyCode = 0
End Sub

Private Sub mnuAlignB_Click()
    Dim objItem As FlowItem
    Dim sngValue As Single

    If IsSelected Then
        frmMain.ToolBoxChanged conUndoMultiple
    
        sngValue = mSelected.P.Top + mSelected.P.Height
        For Each objItem In mCollection
            If Not objItem Is Nothing Then objItem.P.Top = sngValue - objItem.P.Height
        Next objItem
        ItemChanged
    End If
End Sub

Private Sub mnuAlignC_Click() 'vertical axis
    Dim objItem As FlowItem
    Dim sngValue As Single

    If IsSelected Then
        frmMain.ToolBoxChanged conUndoMultiple
        
        sngValue = mSelected.P.CenterX 'Left
        For Each objItem In mCollection
            If Not objItem Is Nothing Then objItem.P.Left = sngValue - objItem.P.Width / 2
        Next objItem
        ItemChanged
    End If
End Sub

Private Sub mnuAlignGrid_Click()
    Dim objItem As FlowItem

    For Each objItem In mCollection
        frmMain.ToolBoxChanged conUndoMultiple
        
        If Not objItem Is Nothing Then frmMain.Align objItem
    Next objItem
    ItemChanged
End Sub

Private Sub mnuAlignL_Click()
    Dim objItem As FlowItem
    Dim sngValue As Single

    If IsSelected Then
        frmMain.ToolBoxChanged conUndoMultiple
        
        sngValue = mSelected.P.Left
        For Each objItem In mCollection
            If Not objItem Is Nothing Then objItem.P.Left = sngValue
        Next objItem
        ItemChanged
    End If
End Sub

Private Sub mnuAlignM_Click() 'horizontal axis
    Dim objItem As FlowItem
    Dim sngValue As Single

    If IsSelected Then
        frmMain.ToolBoxChanged conUndoMultiple
        
        sngValue = mSelected.P.CenterY 'modifies Tops
        For Each objItem In mCollection
            If Not objItem Is Nothing Then objItem.P.Top = sngValue - objItem.P.Height / 2
        Next objItem
        ItemChanged
    End If
End Sub

Private Sub mnuAlignR_Click()
    Dim objItem As FlowItem
    Dim sngValue As Single

    If IsSelected Then
        frmMain.ToolBoxChanged conUndoMultiple
        
        sngValue = mSelected.P.Left + mSelected.P.Width
        For Each objItem In mCollection
            If Not objItem Is Nothing Then objItem.P.Left = sngValue - objItem.P.Width
        Next objItem
        ItemChanged
    End If
End Sub

Private Sub mnuAlignT_Click()
    Dim objItem As FlowItem
    Dim sngValue As Single

    If IsSelected Then
        frmMain.ToolBoxChanged conUndoMultiple
    
        sngValue = mSelected.P.Top
        For Each objItem In mCollection
            If Not objItem Is Nothing Then objItem.P.Top = sngValue
        Next objItem
        ItemChanged
    End If
End Sub




Private Sub mnuCopyA_Click() 'appearance
    DoCopyAppearance True, True, True, True, True, True, True, True
End Sub

Private Sub mnuCopyF_Click()
    DoCopyFont True, True, True, True, True, True
End Sub



Private Sub mnuSizeCol_Click()
    Dim objItem As FlowItem
    Dim colOrder As Collection
    Dim i As Integer, j As Integer, k As Integer
    Dim sngMinLeft As Single
    Dim sngMaxLeft As Single
    Dim sngColumnW As Single 'column width
    
    
    frmMain.ToolBoxChanged conUndoMultiple
    
    Set colOrder = New Collection
    
    For Each objItem In mCollection
        If colOrder.Count = 0 Then
            colOrder.Add objItem
            sngMinLeft = objItem.P.Left
            sngMaxLeft = sngMinLeft
        Else
            'Get the range of the colleciton
            i = 1
            j = colOrder.Count
            'Determine wheter this should be inserted before the first item
            If objItem.P.Left <= colOrder(i).P.Left Then
                colOrder.Add objItem, , i
            'Determine whether should be inserted after last item
            ElseIf objItem.P.Left >= colOrder(j).P.Left Then
                colOrder.Add objItem, , , j
            'Conduct binary search
            Else
                Do Until j - i <= 1
                    k = (i + j) \ 2
                    If colOrder(k).P.Left < objItem.P.Left Then
                        i = k
                    Else
                        j = k
                    End If
                Loop
                'insert the item where it belongs
                colOrder.Add objItem, , j
            End If
            'keep track of min and max left
            If objItem.P.Left < sngMinLeft Then sngMinLeft = objItem.P.Left
            If objItem.P.Left > sngMaxLeft Then sngMaxLeft = objItem.P.Left
        End If
    Next objItem
    
'    For Each objItem In colOrder
'        Debug.Print "Item "; mFlowChart.GetIndex(objItem); " has text: " & objItem.P.Text
'    Next objItem
    
    sngColumnW = (sngMaxLeft - sngMinLeft) / (colOrder.Count - 1)

'    Debug.Print "Average column width of "; sngColumnW
'    Debug.Print "Min left "; sngMinLeft
'    Debug.Print "Max Left "; sngMaxLeft
    
    For Each objItem In colOrder
        objItem.P.Left = sngMinLeft
        sngMinLeft = sngMinLeft + sngColumnW
    Next objItem
    
    Set colOrder = Nothing
    
    ItemChanged
End Sub


Private Sub mnuSizeH_Click()
    Dim objItem As FlowItem
    Dim sngValue As Single

    If IsSelected Then
        frmMain.ToolBoxChanged conUndoMultiple
    
        sngValue = mSelected.P.Height
        For Each objItem In mCollection
            If Not objItem Is Nothing Then objItem.P.Height = sngValue
        Next objItem
        ItemChanged
    End If
End Sub

Private Sub mnuSizeRows_Click()
    Dim objItem As FlowItem
    Dim colOrder As Collection
    Dim i As Integer, j As Integer, k As Integer
    Dim sngMinTop As Single
    Dim sngMaxTop As Single
    Dim sngRowH  As Single 'row height
    
    frmMain.ToolBoxChanged conUndoMultiple
    
    Set colOrder = New Collection
    
    For Each objItem In mCollection
        If colOrder.Count = 0 Then
            colOrder.Add objItem
            sngMinTop = objItem.P.Top
            sngMaxTop = sngMinTop
        Else
            'Get the range of the colleciton
            i = 1
            j = colOrder.Count
            'Determine wheter this should be inserted before the first item
            If objItem.P.Top <= colOrder(i).P.Top Then
                colOrder.Add objItem, , i
            'Determine whether should be inserted after last item
            ElseIf objItem.P.Top >= colOrder(j).P.Top Then
                colOrder.Add objItem, , , j
            'Conduct binary search
            Else
                Do Until j - i <= 1
                    k = (i + j) \ 2
                    If colOrder(k).P.Top < objItem.P.Top Then
                        i = k
                    Else
                        j = k
                    End If
                Loop
                'insert the item where it belongs
                colOrder.Add objItem, , j
            End If
            'keep track of min and max top
            If objItem.P.Top < sngMinTop Then sngMinTop = objItem.P.Top
            If objItem.P.Top > sngMaxTop Then sngMaxTop = objItem.P.Top
        End If
    Next objItem
    
'    For Each objItem In colOrder
'        Debug.Print "Item "; mFlowChart.GetIndex(objItem); " has text: " & objItem.P.Text
'    Next objItem
    
    sngRowH = (sngMaxTop - sngMinTop) / (colOrder.Count - 1)

'    Debug.Print "Average row height of "; sngRowH
'    Debug.Print "Min top "; sngMinTop
'    Debug.Print "Max top "; sngMaxTop
    
    For Each objItem In colOrder
        objItem.P.Top = sngMinTop
        sngMinTop = sngMinTop + sngRowH
    Next objItem
    
    Set colOrder = Nothing
    
    ItemChanged
End Sub

Private Sub mnuSizeW_Click()
    Dim objItem As FlowItem
    Dim sngValue As Single

    If IsSelected Then
        frmMain.ToolBoxChanged conUndoMultiple
    
        sngValue = mSelected.P.Width
        For Each objItem In mCollection
            If Not objItem Is Nothing Then objItem.P.Width = sngValue
        Next objItem
        ItemChanged
    End If
End Sub


Private Sub mnuSizeWH_Click()
    Dim objItem As FlowItem
    Dim sngValue1 As Single
    Dim sngValue2 As Single

    If IsSelected Then
        frmMain.ToolBoxChanged conUndoMultiple
        
        sngValue1 = mSelected.P.Width
        sngValue2 = mSelected.P.Height
        For Each objItem In mCollection
            If Not objItem Is Nothing Then
                objItem.P.Width = sngValue1
                objItem.P.Height = sngValue2
            End If
        Next objItem
        ItemChanged
    End If
End Sub

Private Sub Form_Activate()
    Update
End Sub

Private Sub Form_Load()
    mFlowChart.ShowSelected = True
    gblnWindowMultis = True
    If WindowState <> vbMinimized Then
        Move Screen.Width * 0.9 - Width, Screen.Height * 0.03
    End If
    Clear
'    Debug.Print "Load"
End Sub


Private Sub Form_Unload(Cancel As Integer)
    Clear
    Set mCollection = Nothing 'release both pointers
    Set mSelected = Nothing
    gblnWindowMultis = False
    mFlowChart.ShowSelected = False
'    Debug.Print "Unload"
End Sub

Private Sub mnuTools_Click()
    Dim lngLastGroupNo  As Long
    Dim objItem         As FlowItem
    Dim blnDifferent    As Boolean
    
    lngLastGroupNo = -1
    
    If mCollection.Count < 2 Then
        mnuToolsGroup.Enabled = False
        mnuToolsUngroup.Enabled = False
    Else 'verify each item belongs to the same group
        For Each objItem In mCollection
            If lngLastGroupNo = -1 Then
                lngLastGroupNo = objItem.P.GroupNo
            ElseIf lngLastGroupNo <> objItem.P.GroupNo Then
                blnDifferent = True
            End If
        Next objItem
        If blnDifferent Then '2 different groups selected
            mnuToolsGroup.Enabled = True 'can combine into larger groups
            mnuToolsUngroup.Enabled = True
        ElseIf lngLastGroupNo = 0 Then 'not in group
            mnuToolsGroup.Enabled = True
            mnuToolsUngroup.Enabled = False
        Else 'all in the same group
            mnuToolsGroup.Enabled = False
            mnuToolsUngroup.Enabled = True
        End If
    End If
End Sub

Private Sub mnuToolsDuplicate_Click()
    Dim colPrevious As Collection
    Dim objItem     As FlowItem
    Dim objNew      As FlowItem
    
    Set colPrevious = mCollection
    Clear 'selection
    Set frmMain.mSelected = Nothing
    
    For Each objItem In colPrevious
        If Not objItem Is Nothing Then
            Set objNew = Duplicate(objItem)
            mFlowChart.Add objNew
            AddItem2 objNew
        End If
    Next objItem
    Update
    ItemChanged
End Sub


Public Sub mnuToolsDelete_Click()
    Dim colPrevious As Collection
    Dim objItem     As FlowItem

    If MsgBox("Delete all " & GetCount & " selected items?", vbQuestion Or vbYesNo, "Delete") = vbYes Then
        frmMain.ToolBoxChanged conUndoDelete, True
        
        Set colPrevious = mCollection
        Clear 'selection
        Set frmMain.mSelected = Nothing
        
        For Each objItem In colPrevious
            If Not objItem Is Nothing Then mFlowChart.RemoveObj objItem
        Next objItem
        Update
        ItemChanged
    End If
End Sub

Public Sub mnuToolsGroup_Click()
    Dim objItem     As FlowItem
    Dim lngMin As Long, lngMax As Long
    Dim lngGroup    As Long
    Dim objGroups   As Groups
    
    frmMain.UpgradeVersion 8
    frmMain.ToolBoxChanged conUndoMultiple
    
    Set objGroups = mFlowChart.Groups
    
    objGroups.MinMaxGroup lngMin, lngMax
    lngGroup = lngMax + 1
    
    For Each objItem In mCollection
        If Not objItem Is Nothing Then
            objGroups.Add objItem, lngGroup
        End If
    Next objItem
End Sub

Public Sub mnuToolsUngroup_Click()
    Dim objGroups As Groups
    Dim objItem   As FlowItem
    
    frmMain.ToolBoxChanged conUndoMultiple
    
    'attempt to remove first group
    Set objGroups = mFlowChart.Groups
    'objGroups.RemoveGroup mSelected.P.GroupNo
    objGroups.Group(mSelected.P.GroupNo).RemoveAll
    
    'go through and ensure each item is not in a group
    For Each objItem In mCollection
        If Not objItem Is Nothing Then
            If objItem.P.GroupNo <> 0 Then
                objGroups.Group(objItem.P.GroupNo).RemoveAll
            End If
        End If
    Next objItem
    
    Clear
End Sub


