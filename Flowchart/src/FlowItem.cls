VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FlowItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Option Explicit
'Richard Fung.  August 12, 2000.

Public P        As Properties

Public Property Get CanRefresh() As Boolean
    'nothing to do
End Property


Public Property Get Description() As String
    Description = "Item" 'must have normal description
End Property

Public Property Get DescriptionF() As String
    'nothing to do - optional description for flow chart items specifically
End Property


Public Function Draw(ByVal View As Object, ByVal Parent As FlowChart) As Long
    SetDrawProps View, Me, Parent
    SetFontProps View, Me, Parent
    
    View.Line (P.Left, P.Top)-Step(P.Width, P.Height), , B
    View.Line (P.Left, P.Top)-Step(P.Width, P.Height)
    View.Line (P.Left, P.Top + P.Height)-(P.Left + P.Width, P.Top)
    View.CurrentX = P.Left + conIndent
    View.CurrentY = P.Top + conIndent
    View.Print "#Error"
End Function

Public Property Get Number() As FAddType
    Number = 0
End Property

Public Sub Refresh(ByVal Parent As FlowChart, ByVal View As Object)
    'the View is for calculations only - not for drawing
    'nothing to refresh in this case
End Sub

Public Property Get TextBottomMrg() As Single
'
End Property

Public Property Get TextLeftMrg() As Single
'
End Property

Public Property Get TextRightMrg() As Single
'
End Property


Public Property Get TextTopMrg() As Single
'
End Property


Private Sub Class_Initialize()
'    DrawOrder = conMiddle
'    LineStyle = vbSolid
'    FillStyle = vbFSTransparent
'    BackColour = vbWhite
'    ForeColour = vbBlack
'    TextColour = vbBlack
    Set P = New Properties
End Sub

Private Sub Class_Terminate()
    Set P = Nothing
End Sub


