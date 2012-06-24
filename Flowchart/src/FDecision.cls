VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FDecision"
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


Private Sub Class_Initialize()
    Set mobjProp = New Properties
End Sub

Private Sub Class_Terminate()
    Set mobjProp = Nothing
End Sub


Private Property Get FlowItem_CanRefresh() As Boolean
    'nothing to do
End Property
'
'Private Property Get FlowItem_CenterX() As Single
'    FlowItem_CenterX = mobjProp.CenterX
'End Property
'
'Private Property Get FlowItem_CenterY() As Single
'    FlowItem_CenterY = mobjProp.CenterY
'End Property


Private Property Get FlowItem_Description() As String
    FlowItem_Description = "Rounded Rectangle"
End Property


Private Property Get FlowItem_DescriptionF() As String
    FlowItem_DescriptionF = "Decision"
End Property


Private Function FlowItem_Draw(ByVal View As Object, ByVal Parent As FlowChart) As Long
    Dim tPoints(0 To 4) As POINTAPI
    Dim lngLoop         As Long
    
    SetDrawProps View, Me, Parent
    'Override command here.
    'View.Line (mobjProp.Left, mobjProp.CenterY)-(mobjProp.CenterX, mobjProp.Top)
    With mobjProp
        tPoints(0).X = .Left
        tPoints(0).Y = .CenterY
        tPoints(1).X = .CenterX
        tPoints(1).Y = .Top
        tPoints(2).X = .Left + .Width
        tPoints(2).Y = .CenterY
        tPoints(3).X = .CenterX
        tPoints(3).Y = .Top + .Height
        tPoints(4).X = .Left
        tPoints(4).Y = .CenterY
        For lngLoop = 0 To 4
            tPoints(lngLoop).X = View.ScaleX(tPoints(lngLoop).X, View.ScaleMode, vbPixels)
            tPoints(lngLoop).Y = View.ScaleY(tPoints(lngLoop).Y, View.ScaleMode, vbPixels)
        Next lngLoop
    End With
    Polygon View.hdc, tPoints(0), 5
    'View.Line -(mobjProp.Left + mobjProp.Width, mobjProp.CenterY)
    'View.Line -(mobjProp.CenterX, mobjProp.Top + mobjProp.Height)
    'View.Line -(mobjProp.Left, mobjProp.CenterY)
    DrawText View, Me, Parent
End Function
'
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
    FlowItem_Number = conAddDecision
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
    FlowItem_TextBottomMrg = mobjProp.Height / 4
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
    FlowItem_TextLeftMrg = mobjProp.Width / 4
End Property

'
'Private Property Let FlowItem_TextRightMrg(ByVal RHS As Single)
'    FlowItem.TextRightMrg = RHS
'End Property

Private Property Get FlowItem_TextRightMrg() As Single
    FlowItem_TextRightMrg = mobjProp.Width / 4
End Property


''Private Property Let FlowItem_TextTopMrg(ByVal RHS As Single)
''    FlowItem.TextTopMrg = RHS
''End Property

Private Property Get FlowItem_TextTopMrg() As Single
    FlowItem_TextTopMrg = mobjProp.Height / 4
End Property

