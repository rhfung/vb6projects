VERSION 5.00
Object = "{B0F710F8-474B-4DE4-BE20-2C604C6801EE}#9.0#0"; "RichUtil1.ocx"
Begin VB.Form frmLayer 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Layer"
   ClientHeight    =   375
   ClientLeft      =   7275
   ClientTop       =   690
   ClientWidth     =   4335
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   375
   ScaleWidth      =   4335
   ShowInTaskbar   =   0   'False
   Begin VB.Frame fraButton 
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   2280
      TabIndex        =   2
      Top             =   0
      Width           =   1815
      Begin VB.CommandButton cmdRem 
         Caption         =   "-"
         Height          =   375
         Left            =   720
         TabIndex        =   4
         Top             =   0
         Width           =   495
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "+"
         Height          =   375
         Left            =   120
         TabIndex        =   3
         Top             =   0
         Width           =   495
      End
      Begin VB.CommandButton cmdDef 
         Caption         =   "Def"
         Enabled         =   0   'False
         Height          =   375
         Left            =   1320
         TabIndex        =   5
         ToolTipText     =   "Set this object as default object properties"
         Top             =   0
         Width           =   495
      End
   End
   Begin VB.ComboBox cboLayer 
      Height          =   315
      Left            =   120
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   0
      Width           =   1935
   End
   Begin RichUtil1.EListBox lstLayerItems 
      Height          =   3615
      Left            =   240
      TabIndex        =   0
      Top             =   600
      Visible         =   0   'False
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   6376
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnDivisions =   -1  'True
      CheckBoxes      =   -1  'True
      Columns         =   3
      BeginProperty Column {276D682E-2140-4AB5-8D2C-F105DB576FFF} 
         Count           =   3
      EndProperty
   End
End
Attribute VB_Name = "frmLayer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_objFile As FlowChart
Private m_blnUpdating As Boolean
Public Property Get ActiveLayer() As Long
    If cboLayer.ListIndex > -1 Then
        ActiveLayer = cboLayer.ItemData(cboLayer.ListIndex)
    Else
        ActiveLayer = -1
    End If
End Property

Public Sub LoadData(ByVal File As FlowChart)
    Dim objItem As FlowItem
    Dim blnLayer(0 To 255) As Boolean
    
    Set m_objFile = File
    
    For Each objItem In File
        On Error Resume Next
        blnLayer(objItem.P.Layer) = True
    Next objItem
    
    Dim lngLayer As Long
    
    cboLayer.AddItem "All Layers"
    cboLayer.ItemData(0) = -1
    For lngLayer = 0 To 255
        If blnLayer(lngLayer) Then
            cboLayer.AddItem "Layer " & lngLayer
            cboLayer.ItemData(cboLayer.NewIndex) = lngLayer
        End If
    Next lngLayer

    LoadItems
    cboLayer.ListIndex = 0
End Sub

Public Sub LoadItems()
    Dim lngLayer As Long
    Dim objItem As FlowItem
    
    lstLayerItems.Clear
    
    For Each objItem In m_objFile
        lstLayerItems.AddItem2 Replace$(Left$(objItem.P.Text, 30), vbCrLf, "¶"), 0, objItem.Description, objItem.DescriptionF
    Next objItem
End Sub


Public Sub Update() 'receive messages from other windows
    m_blnUpdating = True
    If lstLayerItems.ListCount <> m_objFile.Count Then
        LoadItems
    End If
    If frmMain.IsSelected Then
        On Error Resume Next
        lstLayerItems.ListIndex = m_objFile.GetIndex(frmMain.mSelected)
    End If
    m_blnUpdating = False
End Sub


Public Sub UpdateDrawing()
    Dim lngIndex As Long
    Dim blnFound As Boolean
    
    For lngIndex = 1 To lstLayerItems.ListCount
        If lstLayerItems.Selected(lngIndex) And Not blnFound Then
            Set frmMain.mSelected = m_objFile(lngIndex)
            blnFound = True
        End If
        m_objFile(lngIndex).P.Enabled = lstLayerItems.Selected(lngIndex)
    Next lngIndex
    
    frmMain.Redraw
End Sub
Public Sub UpdateDrawingShowAll()
    Dim lngIndex As Long
    
    For lngIndex = 1 To lstLayerItems.ListCount
        m_objFile(lngIndex).P.Enabled = True
    Next lngIndex
    
    frmMain.Redraw
End Sub


Private Sub cboLayer_Click()
    Dim lngLayer As Long
    Dim lngIndex As Long
    Dim blnFirst As Boolean
    
    If cboLayer.ListIndex > -1 And ActiveLayer <> -1 Then
        lngLayer = ActiveLayer
        
        For lngIndex = 1 To lstLayerItems.ListCount
            lstLayerItems.Selected(lngIndex) = (m_objFile(lngIndex).P.Layer = lngLayer)
            If (m_objFile(lngIndex).P.Layer = lngLayer) And Not blnFirst Then
                blnFirst = True
                lstLayerItems_Click lngIndex
            End If
        Next lngIndex
        
        If Not blnFirst Then
            Set frmMain.mSelected = Nothing
        End If
        
        lstLayerItems.Visible = True
        
        Height = Height - ScaleHeight + cboLayer.Height * 10
        UpdateDrawing
        
        m_objFile.DefaultLayer = ActiveLayer
        
        cmdRem.Enabled = True
        If lstLayerItems.ListCount > 0 And lstLayerItems.ListIndex > 0 Then
            cmdDef.Enabled = lstLayerItems.Selected(lstLayerItems.ListIndex)
        Else
            cmdDef.Enabled = False
        End If
    Else
        lstLayerItems.Visible = False
        
        UpdateDrawingShowAll
        Height = Height - ScaleHeight + cboLayer.Height
        
        m_objFile.DefaultLayer = 0
        
        cmdRem.Enabled = False
        cmdDef.Enabled = False
    End If
End Sub



Private Sub cmdAdd_Click()
    Dim lngL As Long

    lngL = cboLayer.ItemData(cboLayer.ListCount - 1) + 1
    If lngL < 255 Then
        cboLayer.AddItem "Layer " & lngL
        cboLayer.ItemData(cboLayer.NewIndex) = lngL
        cboLayer.ListIndex = cboLayer.NewIndex
    Else
        MsgBox "Maximum number of layers has been added.", vbExclamation
    End If
End Sub

Private Sub cmdDef_Click()
    If frmMain.IsSelected Then
        'make a new item
        Set m_objFile.Layers.DefaultShape(ActiveLayer) = New FlowItem
        CopyProperties frmMain.mSelected.P, m_objFile.Layers.DefaultShape(ActiveLayer).P
        With m_objFile.Layers.DefaultShape(ActiveLayer).P
            .Text = vbNullString 'clear text
        End With
    Else
        MsgBox "Select an object before setting Default Shape.", vbInformation
    End If
    
End Sub

Private Sub cmdRem_Click()
    Dim lngItem As Long
    
    For lngItem = 1 To lstLayerItems.ListCount
        If lstLayerItems.Selected(lngItem) Then
            m_objFile(lngItem).P.Layer = 0
        End If
    Next lngItem
    cboLayer.RemoveItem cboLayer.ListIndex
    cboLayer.ListIndex = cboLayer.ListCount - 1
End Sub

Private Sub Form_DblClick()
    Dim lngLayer As Long
    Dim objItem  As FlowItem
    
    If MsgBox("This will overwrite existing default shapes?  Continue?", vbQuestion Or vbYesNo) = vbNo Then Exit Sub
        
    cboLayer.Clear
    cboLayer.AddItem "All Layers"
    cboLayer.ItemData(cboLayer.NewIndex) = -1
    
    For lngLayer = 0 To 255
        cboLayer.AddItem "Layer " & lngLayer
        cboLayer.ItemData(cboLayer.NewIndex) = lngLayer
        
        Set objItem = New FlowItem
        objItem.P.ForeColour = QBColor(lngLayer Mod 15)
        Set m_objFile.Layers.DefaultShape(lngLayer) = objItem
    Next lngLayer
    
    cboLayer.ListIndex = 0
End Sub

Private Sub Form_Load()
    gblnWindowLayer = True
    If WindowState <> vbMinimized Then
        Move Screen.Width * 0.9 - Width, Screen.Height * 0.1
    End If
    
End Sub


Private Sub Form_Resize()
    If WindowState <> vbMinimized And ScaleWidth - fraButton.Width > 0 Then
        
        cboLayer.Move 0, 0, ScaleWidth - fraButton.Width
        
        fraButton.Move cboLayer.Width, 0, fraButton.Width, cboLayer.Height
        cmdAdd.Height = cboLayer.Height
        cmdRem.Height = cboLayer.Height
        cmdDef.Height = cboLayer.Height
        
        On Error Resume Next
        If lstLayerItems.Visible Then lstLayerItems.Move 0, cboLayer.Height + 100, ScaleWidth, ScaleHeight - cboLayer.Height - 100
    End If
End Sub


Private Sub Form_Unload(Cancel As Integer)
    gblnWindowLayer = False
End Sub

Private Sub lstLayerItems_Click(ListIndex As Long)
    If lstLayerItems.Selected(ListIndex) And Not m_blnUpdating Then
        cmdDef.Enabled = True
        frmMain.SetSelected m_objFile(ListIndex)
        frmMain.ToolBoxDoneNoCopy
    End If
End Sub


Private Sub lstLayerItems_ItemChecked(ByVal ListIndex As Long, ByVal Checked As Boolean, Cancel As Boolean)
    If m_blnUpdating Then Exit Sub
    If cboLayer.ItemData(cboLayer.ListIndex) = 0 Then
        If Checked = False Then Cancel = True: Exit Sub
    End If
    
    If Checked And cboLayer.ListIndex > -1 Then
        m_objFile(ListIndex).P.Layer = cboLayer.ItemData(cboLayer.ListIndex)
        m_objFile(ListIndex).P.Enabled = True
        Set frmMain.mSelected = m_objFile(ListIndex)
        frmMain.RedrawSingle frmMain.mSelected
    ElseIf Not Checked Then
        m_objFile(ListIndex).P.Layer = 0
        m_objFile(ListIndex).P.Enabled = (m_objFile(ListIndex).P.Layer = ActiveLayer)
        Set frmMain.mSelected = Nothing
    End If
End Sub


