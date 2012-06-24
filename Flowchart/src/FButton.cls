VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit 'using Properties
'28 Jan 2002

Implements FlowItem

Private mobjProp    As Properties
Private mobjItem    As FlowItem
Private msngOne     As Single
Private mblnDown    As Boolean

Public Sub Click(ByVal View As Object, ByVal Parent As FlowChart)
    Dim vntSplit As Variant
    Dim blnFail  As Boolean
    Dim strPath  As String
    
    mblnDown = True
    Screen.MousePointer = vbHourglass
    If Not View Is Nothing Then
        FlowItem_Draw View, Parent
    End If
    If Len(mobjProp.Tag3) Then
        vntSplit = Split(LCase$(mobjProp.Tag3), " ", 2)
    Else
        ReDim vntSplit(0)
    End If
    On Error GoTo Handler
    
    strPath = Parent.GetPath()
    If Len(strPath) Then
        ChDrive strPath
        ChDir strPath
    End If
    
    If vntSplit(0) = "shell" Then
        On Error Resume Next
        Shell vntSplit(1), vbNormalFocus
        If Err Then
            MsgBox "Error running program: " & Err.Description & " (" & Err.Number & ")", vbExclamation
            Err.Clear
        End If
    ElseIf vntSplit(0) = "open" Then
        Load frmForm
        With frmForm
            .Caption = "Flow Chart"
            .mnuFileOpen.Visible = False
            Set .mFlowChart = New FlowChart
            
            If .mFlowChart.Load(vntSplit(1), View) <> 0 Then
                MsgBox "Problem opening the flow chart file named:" & vbNewLine & vntSplit(1), vbExclamation
                Set .mFlowChart = Nothing
                blnFail = True
            Else
                .SetZoom .mFlowChart.ZoomPercent
                .Show
            End If
        End With
        If blnFail Then
            Unload frmForm
        Else
            If View Is frmForm.picPreview Then 'do not draw over the new file
                mblnDown = False
                Screen.MousePointer = vbDefault
                Exit Sub
            End If
        End If
    ElseIf vntSplit(0) = "pause" Then
        MsgBox "Paused."
    Else
        Sleep 250
    End If
    
    mblnDown = False
    If Not View Is Nothing Then
        FlowItem_Draw View, Parent
    End If
    Screen.MousePointer = vbDefault
    Exit Sub
Handler:
    mblnDown = False
    If Not View Is Nothing Then
        FlowItem_Draw View, Parent
    End If
    Screen.MousePointer = vbDefault
    Exit Sub
End Sub


Private Sub Class_Initialize()
    Set mobjProp = New Properties
    Set mobjItem = New FlowItem
    mobjProp.FontFace = "MS Sans Serif"
    mobjProp.TextSize = 8
End Sub

Private Sub Class_Terminate()
    Set mobjProp = Nothing
    Set mobjItem = Nothing
End Sub


Private Property Get FlowItem_CanRefresh() As Boolean
    'nothing to do
End Property

Private Property Get FlowItem_Description() As String
    FlowItem_Description = "Button"
End Property


Private Property Get FlowItem_DescriptionF() As String
    'nothing to do
End Property


Private Function FlowItem_Draw(ByVal View As Object, ByVal Parent As FlowChart) As Long
    Dim lngColour(0 To 4) As Long
    '0 face, 1 highlight, 2 light, 3 dark shadow, 4 darkest shadow
    Dim lngPixels         As Long
    
    If TypeOf View Is Printer Then
        lngColour(0) = &HC0C0C0  'normal gray QBColor(7)
        If mblnDown Then
            lngColour(1) = vbBlack
            lngColour(2) = &H808080 'QBColor(8)
            lngColour(3) = &HE0E0E0
            lngColour(4) = vbWhite
        Else
            lngColour(1) = vbWhite
            lngColour(2) = &HE0E0E0
            lngColour(3) = &H808080 'QBColor(8)
            lngColour(4) = vbBlack
        End If
        View.FillColor = &HC0C0C0 'QBColor(7)
        View.ForeColor = vbBlack
    Else
        lngColour(0) = vbButtonFace
        If mblnDown Then
            lngColour(1) = vb3DDKShadow
            lngColour(2) = vb3DShadow
            lngColour(3) = vb3DLight
            lngColour(4) = vb3DHighlight
        Else
            lngColour(1) = vb3DHighlight
            lngColour(2) = vb3DLight
            lngColour(3) = vb3DShadow
            lngColour(4) = vb3DDKShadow
        End If
        View.FillColor = vbButtonFace
        View.ForeColor = vbButtonText
    End If
    
    msngOne = View.ScaleX(15, vbTwips, View.ScaleMode) * IIf(Parent.Percentage < 1, 1, Parent.Percentage)
    lngPixels = IIf(Parent.Percentage < 1, 1, Parent.Percentage)
    'different set of properties
    View.DrawStyle = vbSolid
    View.DrawWidth = View.ScaleX(lngPixels * 15, vbTwips, vbPixels)
    View.FillStyle = vbFSTransparent
    'changed above
    'View.FillColor = vbButtonFace
    'View.ForeColor = vbButtonText
    'draw button
    View.Line (mobjProp.Left, mobjProp.Top)-Step(mobjProp.Width, mobjProp.Height), lngColour(0), BF 'box
    View.Line (mobjProp.Left, mobjProp.Top)-Step(mobjProp.Width, 0), lngColour(1) '1
    View.Line (mobjProp.Left, mobjProp.Top)-Step(0, mobjProp.Height), lngColour(1)
    View.Line (mobjProp.Left + msngOne, mobjProp.Top + msngOne)-Step(mobjProp.Width - msngOne, 0), lngColour(2) '2
    View.Line (mobjProp.Left + msngOne, mobjProp.Top + msngOne)-Step(0, mobjProp.Height - msngOne), lngColour(2)
    If lngPixels = 1 Then
        View.Line (mobjProp.Left + mobjProp.Width, mobjProp.Top)-Step(0, mobjProp.Height + msngOne), lngColour(4) '4
        View.Line (mobjProp.Left, mobjProp.Top + mobjProp.Height)-Step(mobjProp.Width + msngOne, 0), lngColour(4)
        View.Line (mobjProp.Left + mobjProp.Width - msngOne, mobjProp.Top + msngOne)-Step(0, mobjProp.Height - msngOne), lngColour(3)  '3
        View.Line (mobjProp.Left + msngOne, mobjProp.Top + mobjProp.Height - msngOne)-Step(mobjProp.Width - msngOne, 0), lngColour(3)
    Else
        View.Line (mobjProp.Left + mobjProp.Width, mobjProp.Top)-(mobjProp.Left + mobjProp.Width, mobjProp.Top + mobjProp.Height), lngColour(4)   '4
        View.Line (mobjProp.Left, mobjProp.Top + mobjProp.Height)-(mobjProp.Left + mobjProp.Width, mobjProp.Top + mobjProp.Height), lngColour(4)
        View.Line (mobjProp.Left + mobjProp.Width - msngOne, mobjProp.Top + msngOne)-Step(0, mobjProp.Height - msngOne * 2), lngColour(3) '3
        View.Line (mobjProp.Left + msngOne, mobjProp.Top + mobjProp.Height - msngOne)-Step(mobjProp.Width - msngOne * 2, 0), lngColour(3)
    End If
    
    SetFontProps View, Me, Parent
    CopyProperties mobjProp, mobjItem.P
    With mobjItem.P
        .Left = mobjProp.Left 'mobjProp.CenterX - View.TextWidth(mobjProp.Text) / 2
        .Top = mobjProp.CenterY - View.TextHeight(mobjProp.Text) / 2
        .Width = mobjProp.Width '- mobjProp.CenterX
        .Height = mobjProp.Top + mobjProp.Height - .Top
        If mblnDown Then
            .Left = .Left + msngOne
            .Top = .Top + msngOne
        End If
        .TextAlign = DT_CENTER
'        View.Line (.Left, .Top)-Step(.Width, .Height), , B
    End With
    
    DrawText View, mobjItem, Parent
End Function

Private Property Get FlowItem_Number() As FAddType
    FlowItem_Number = conAddButton
End Property


Private Property Set FlowItem_P(ByVal RHS As Properties)
    Set mobjProp = RHS
End Property


Private Property Get FlowItem_P() As Properties
    Set FlowItem_P = mobjProp
End Property


Private Sub FlowItem_Refresh(ByVal Parent As FlowChart, ByVal View As Object)
    'nothing to do
End Sub

Private Property Get FlowItem_TextBottomMrg() As Single
    FlowItem_TextBottomMrg = msngOne * 2
End Property


Private Property Get FlowItem_TextLeftMrg() As Single
    FlowItem_TextLeftMrg = msngOne * 2
End Property


Private Property Get FlowItem_TextRightMrg() As Single
    FlowItem_TextRightMrg = msngOne * 2
End Property

Private Property Get FlowItem_TextTopMrg() As Single
    FlowItem_TextTopMrg = msngOne * 2
End Property


