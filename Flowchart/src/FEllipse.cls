VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FEllipse"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'FEllipse
'August 24, 2002
Implements FlowItem
Private mobjProp As Properties

Private Declare Function Ellipse Lib "gdi32" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long


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
    FlowItem_Description = "Ellipse"
End Property

Private Property Get FlowItem_DescriptionF() As String
    'nothing to do
End Property

Private Function FlowItem_Draw(ByVal View As Object, ByVal Parent As FlowChart) As Long
    Dim X1 As Long
    Dim Y1 As Long
    Dim X2 As Long
    Dim Y2 As Long
    
    X1 = View.ScaleX(mobjProp.Left, View.ScaleMode, vbPixels)
    Y1 = View.ScaleY(mobjProp.Top, View.ScaleMode, vbPixels)
    X2 = View.ScaleX(mobjProp.Left + mobjProp.Width, View.ScaleMode, vbPixels)
    Y2 = View.ScaleY(mobjProp.Top + mobjProp.Height, View.ScaleMode, vbPixels)
    
    SetDrawProps View, Me, Parent
    Ellipse View.hdc, X1, Y1, X2, Y2
    DrawText View, Me, Parent
End Function

Private Property Get FlowItem_Number() As FAddType
    FlowItem_Number = conAddEllipse
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
    FlowItem_TextBottomMrg = mobjProp.Height / 2 - Sqr((mobjProp.Height / 2) ^ 2 / 2) + conIndent
End Property

Private Property Get FlowItem_TextLeftMrg() As Single
    FlowItem_TextLeftMrg = mobjProp.Width / 2 - Sqr((mobjProp.Width / 2) ^ 2 / 2) + conIndent
End Property

Private Property Get FlowItem_TextRightMrg() As Single
    FlowItem_TextRightMrg = mobjProp.Width / 2 - Sqr((mobjProp.Width / 2) ^ 2 / 2) + conIndent
End Property

Private Property Get FlowItem_TextTopMrg() As Single
    FlowItem_TextTopMrg = mobjProp.Height / 2 - Sqr((mobjProp.Height / 2) ^ 2 / 2) + conIndent
End Property
