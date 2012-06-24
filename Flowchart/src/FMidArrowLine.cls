VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FMidArrowLine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
'Richard Fung.  August 12, 2000.

Implements FlowItem
Private mobjProp As Properties

Const conSFlow = 220 '300 'flow line
'Const conSPointed = 180 'pointed arrow

Const conStdArrow = 3600 'twips
Const conLinExSpace = 30 'extra space for line text
Const conMidArrow = 45 * conR 'degrees conv to radians

Private Sub Class_Initialize()
    Set mobjProp = New Properties
End Sub

Private Sub Class_Terminate()
    Set mobjProp = Nothing
End Sub

Private Property Get FlowItem_CanRefresh() As Boolean
    'nothing to do
End Property

Private Property Get FlowItem_Description() As String
    FlowItem_Description = "Mid Arrow Line"
End Property


Private Property Get FlowItem_DescriptionF() As String
    FlowItem_DescriptionF = "Flow Line"
End Property


'Private Property Get FlowItem_CenterX() As Single
'    FlowItem_CenterX = mobjProp.CenterX
'End Property
'
'Private Property Get FlowItem_CenterY() As Single
'    FlowItem_CenterY = mobjProp.CenterY
'End Property


Private Function FlowItem_Draw(ByVal View As Object, ByVal Parent As FlowChart) As Long
    Dim angle As Double, b As Double, c As Double
    Dim sX As Single, sY As Single, tRect As Rect
    Dim sngS As Single 'use to be a sngStant, but now single
    
    tRect = GetRect(Me)
    SetDrawProps View, Me, Parent
    
    View.Line (tRect.X1, tRect.Y1)-(tRect.X2, tRect.Y2) 'draw line
    
    'calculate angle
    If (tRect.X2 - tRect.X1) = 0 Then
        If tRect.Y2 > tRect.Y1 Then
            angle = 270 * conR
        Else
            angle = 90 * conR
        End If
    Else
        angle = Atn((tRect.Y1 - tRect.Y2) / (tRect.X2 - tRect.X1))
    End If
    'convert arrow point to proper side
    If (tRect.X2 - tRect.X1) < 0 Then
        angle = angle + 180 * conR
    End If
    
    'calc new angles
    b = angle + conMidArrow
    c = angle - conMidArrow
    'calculate length of line
'    If Sqr((tRect.y2 - tRect.y1) ^ 2 + (tRect.x2 - tRect.x1) ^ 2) > sngStdArrow Then
        sngS = conSFlow * mobjProp.ArrowSize / 100 'size of arrow
        sX = mobjProp.CenterX 'too long, put in middle
        sY = mobjProp.CenterY
'    Else
'        sngS = sngSPointed 'size of arrow
'        sX = tRect.x2 'short enough for end
'        sY = tRect.y2
'        View.Line (sX + Cos(b) * sngS * -1, sY + Sin(b) * sngS)-(sX + Cos(c) * sngS * -1, sY + Sin(c) * sngS)
'    End If
    'draw the lines
    View.Line (sX, sY)-(sX + Cos(b) * sngS * -1, sY + Sin(b) * sngS)
    View.Line (sX, sY)-(sX + Cos(c) * sngS * -1, sY + Sin(c) * sngS)
    'View.Line (sX + Cos(b) * sngS * -1, sY + Sin(b) * sngS)-(sX + Cos(c) * sngS * -1, sY + Sin(c) * sngS)
    
    DrawTextForLines View, Me, Parent
    'calculate position of X, Y for text
'    If mobjProp.Width = 0 Then
'        View.CurrentY = mobjProp.CenterY - View.TextHeight(mobjProp.Text) / 2
'    Else
'        If mobjProp.Height < 0 Then 'Y
'            View.CurrentY = mobjProp.CenterY + conLinExSpace
'        Else
'            View.CurrentY = mobjProp.CenterY - View.TextHeight(mobjProp.Text) - conLinExSpace
'        End If
'    End If
'    If mobjProp.Height = 0 Then
'        View.CurrentX = mobjProp.CenterX - View.TextWidth(mobjProp.Text) / 2
'    Else
'        If mobjProp.Width < 0 Then 'X
'            View.CurrentX = mobjProp.CenterX - View.TextWidth(mobjProp.Text) - conLinExSpace
'        Else
'            View.CurrentX = mobjProp.CenterX + conLinExSpace
'        End If
'    End If
'    'draw text
'    SetFontProps View, Me, Parent
'    'View.ForeColor = mobjProp.TextColour
'    View.Print mobjProp.Text
End Function

'Private Property Let FlowItem_DrawOrder(ByVal RHS As FOrder)
'    FlowItem.DrawOrder = RHS
'End Property
'
'Private Property Get FlowItem_DrawOrder() As FOrder
'    FlowItem_DrawOrder = FlowItem.DrawOrder
'End Property
'
'Private Property Let FlowItem_Height(ByVal RHS As Single)
'    mobjProp.Height = RHS
'End Property
'
'Private Property Get FlowItem_Height() As Single
'    FlowItem_Height = mobjProp.Height
'End Property
'
'Private Property Let FlowItem_Left$(ByVal RHS As Single)
'    mobjProp.Left = RHS
'End Property
'
'Private Property Get FlowItem_Left$() As Single
'    FlowItem_Left = mobjProp.Left
'End Property

Private Property Get FlowItem_Number() As FAddType
    FlowItem_Number = conAddMidArrowLine
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

'
'Private Property Let FlowItem_TextBottomMrg(ByVal RHS As Single)
'    FlowItem.TextBottomMrg = RHS
'End Property

Private Property Get FlowItem_TextBottomMrg() As Single
    FlowItem_TextBottomMrg = FlowItem.TextBottomMrg
End Property

'Private Property Get mobjProp.TextColour() As Long
'    mobjProp.TextColour = FlowItem.TextColour
'End Property
'
'Private Property Let mobjProp.TextColour(ByVal RHS As Long)
'    FlowItem.TextColour = RHS
'End Property

'
'Private Property Let FlowItem_TextLeftMrg(ByVal RHS As Single)
'    FlowItem.TextLeftMrg = RHS
'End Property

Private Property Get FlowItem_TextLeftMrg() As Single
    FlowItem_TextLeftMrg = FlowItem.TextLeftMrg
End Property


'Private Property Let FlowItem_TextRightMrg(ByVal RHS As Single)
'    FlowItem.TextRightMrg = RHS
'End Property

Private Property Get FlowItem_TextRightMrg() As Single
    FlowItem_TextRightMrg = FlowItem.TextRightMrg
End Property


'Private Property Let FlowItem_TextTopMrg(ByVal RHS As Single)
'    FlowItem.TextTopMrg = RHS
'End Property

Private Property Get FlowItem_TextTopMrg() As Single
    FlowItem_TextTopMrg = FlowItem.TextTopMrg
End Property

