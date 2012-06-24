VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FTerminator"
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

Private Property Get FlowItem_Description() As String
    FlowItem_Description = "Rounded Rectangle"
End Property


Private Property Get FlowItem_DescriptionF() As String
    FlowItem_DescriptionF = "Terminator"
End Property


'Private Property Get FlowItem_CenterX() As Single
'    FlowItem_CenterX = mobjProp.CenterX
'End Property
'
'Private Property Get FlowItem_CenterY() As Single
'    FlowItem_CenterY = mobjProp.CenterY
'End Property


Private Function FlowItem_Draw(ByVal View As Object, ByVal Parent As FlowChart) As Long
    Dim Radius As Double
    Dim objItem As FlowItem
    
    Radius = mobjProp.Height / 2
    If Radius < 0 Then Exit Function
    
    If mobjProp.Width <= mobjProp.Height Then
        Set objItem = New FlowItem
        Set objItem.P = mobjProp
        objItem.Draw View, Parent
        Set objItem = Nothing
    Else
        If mobjProp.FillStyle <> vbFSTransparent Then
            SetDrawProps View, Me, Parent
            View.DrawStyle = vbInvisible
            View.Circle (mobjProp.Left + Radius, mobjProp.Top + Radius), Radius
            View.Circle (mobjProp.Left + mobjProp.Width - Radius, mobjProp.Top + Radius), Radius
            View.Line (mobjProp.Left + Radius, mobjProp.Top)-(mobjProp.Left + mobjProp.Width - Radius, mobjProp.Top + mobjProp.Height), , B
        End If
        SetDrawProps View, Me, Parent
        View.Circle (mobjProp.Left + Radius, mobjProp.Top + Radius), Radius, , 90 * conR, 270 * conR
        View.Circle (mobjProp.Left + mobjProp.Width - Radius, mobjProp.Top + Radius), Radius, , 270 * conR, 90 * conR
        View.Line (mobjProp.Left + Radius, mobjProp.Top)-(mobjProp.Left + mobjProp.Width - Radius, mobjProp.Top)
        View.Line (mobjProp.Left + Radius, mobjProp.Top + mobjProp.Height)-(mobjProp.Left + mobjProp.Width - Radius, mobjProp.Top + mobjProp.Height)
        DrawText View, Me, Parent
    End If
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
    FlowItem_Number = conAddTerminator
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
    FlowItem_TextBottomMrg = 0
End Property


'Private Property Get mobjProp.TextColour() As Long
'    mobjProp.TextColour = FlowItem.TextColour
'End Property
'
'Private Property Let mobjProp.TextColour(ByVal RHS As Long)
'    FlowItem.TextColour = RHS
'End Property

'Private Property Let FlowItem_TextLeftMrg(ByVal RHS As Single)
'    FlowItem.TextLeftMrg = RHS
'End Property

Private Property Get FlowItem_TextLeftMrg() As Single
    FlowItem_TextLeftMrg = mobjProp.Height / 2
End Property


'Private Property Let FlowItem_TextRightMrg(ByVal RHS As Single)
'    FlowItem.TextRightMrg = RHS
'End Property

Private Property Get FlowItem_TextRightMrg() As Single
    FlowItem_TextRightMrg = mobjProp.Height / 2
End Property


'Private Property Let FlowItem_TextTopMrg(ByVal RHS As Single)
'    FlowItem.TextTopMrg = RHS
'End Property

Private Property Get FlowItem_TextTopMrg() As Single
    FlowItem_TextTopMrg = 0
End Property

