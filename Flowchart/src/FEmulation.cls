VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FEmulation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'February 13, 2002

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
    Select Case mobjProp.EmulationNumber
    Case conAddEllipse
        FlowItem_Description = "Ellipse"
    Case conAddShapeA To conAddShapeF
        FlowItem_Description = "Shape"
    Case conAddSecurity
        FlowItem_Description = "Security Module"
    Case conAddExtra2 To conAddExtra3
        FlowItem_Description = "Extra"
    Case conAddExtra1
        FlowItem_Description = "Area"
    Case Else
        FlowItem_Description = "Generic Object"
    End Select
End Property


Private Property Get FlowItem_DescriptionF() As String
    'nothing
End Property

Private Function FlowItem_Draw(ByVal View As Object, ByVal Parent As FlowChart) As Long
    Dim sngWidthQ As Single
    Dim sngHeightQ As Single
    
    If mobjProp.EmulationNumber <> conAddExtra1 Then
        SetDrawProps View, Me, Parent
        sngWidthQ = mobjProp.Width / 4
        sngHeightQ = mobjProp.Height / 4
        View.Line (mobjProp.Left, mobjProp.Top)-Step(mobjProp.Width, mobjProp.Height), , B
        View.Line (mobjProp.Left, mobjProp.Top)-Step(mobjProp.Width, mobjProp.Height)
        View.Line (mobjProp.Left, mobjProp.Top + mobjProp.Height)-(mobjProp.Left + mobjProp.Width, mobjProp.Top)
        View.Line (mobjProp.Left + sngWidthQ, mobjProp.Top + sngHeightQ)-Step(mobjProp.Width - sngWidthQ * 2, mobjProp.Height - sngHeightQ * 2), , B
        View.CurrentX = mobjProp.Left + conIndent
        View.CurrentY = mobjProp.Top + conIndent
        SetFontProps View, Me, Parent
        View.Print "#" & FlowItem_Description
'        Select Case mobjProp.EmulationNumber
'        Case conAddEllipse
'            View.Print "#Ellipse"
'        Case conAddShapeA To conAddShapeF
'            View.Print "#Shape"
'        Case conAddSecurity
'            View.Print "#Security Module"
'        Case conAddExtra1 To conAddExtra3
'            View.Print "#Extra"
'        Case Else
'            View.Print "#Generic object"
'        End Select
    End If
End Function

Private Property Get FlowItem_Number() As FAddType
    FlowItem_Number = mobjProp.EmulationNumber
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
    'nothing to do
End Property

Private Property Get FlowItem_TextLeftMrg() As Single
    'nothing to do
End Property

Private Property Get FlowItem_TextRightMrg() As Single
    'nothing to do
End Property

Private Property Get FlowItem_TextTopMrg() As Single
    'nothing to do
End Property
