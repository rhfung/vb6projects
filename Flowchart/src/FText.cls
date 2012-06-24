VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FText"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit 'Using Properties
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
    FlowItem_CanRefresh = True
End Property

'Private Property Get FlowItem_CenterX() As Single
'    FlowItem_CenterX = mobjProp.CenterX
'End Property

Private Property Get FlowItem_CenterY() As Single
    FlowItem_CenterY = mobjProp.CenterY
End Property


Private Property Get FlowItem_Description() As String
    FlowItem_Description = "Text"
End Property

Private Property Get FlowItem_DescriptionF() As String
    'nothing to do
End Property


Private Function FlowItem_Draw(ByVal View As Object, ByVal Parent As FlowChart) As Long
    'Override command here.
    SetDrawProps View, Me, Parent
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
    FlowItem_Number = conAddText
End Property

Private Property Set FlowItem_P(ByVal RHS As Properties)
    Set mobjProp = RHS
End Property

Private Property Get FlowItem_P() As Properties
    Set FlowItem_P = mobjProp
End Property

Private Sub FlowItem_Refresh(ByVal Parent As FlowChart, ByVal View As Object)
    Dim vntSplit As Variant
    
    If Len(mobjProp.Tag3) Then
        vntSplit = Split(LCase$(mobjProp.Tag3), " ", 2)
    Else
        ReDim vntSplit(0)
    End If
    
    On Error GoTo Handler
    If vntSplit(0) = "type" And UBound(vntSplit) = 1 Then
        Select Case vntSplit(1)
        Case "today"
            mobjProp.Text = Format(Now, "General Date")
        Case "date"
            mobjProp.Text = Format(Now, "Medium Date")
        Case "time"
            mobjProp.Text = Format(Now, "Medium Time")
        Case "filename"
            mobjProp.Text = Parent.FileName
        Case "file"
            mobjProp.Text = Parent.GetFilePart
        Case "path"
            mobjProp.Text = Parent.GetPath
        Case "info"
            mobjProp.Text = "This document was created by Flow Chart version " & App.Major & "." & App.Minor & "." & App.Revision
        Case Else 'date and time formats
            mobjProp.Text = Format(Now, vntSplit(1))
        End Select
        mobjProp.CanEdit = False
    Else
        mobjProp.CanEdit = True
    End If
    
    Exit Sub
Handler:
End Sub


Private Property Get FlowItem_TextBottomMrg() As Single
    FlowItem_TextBottomMrg = 0
End Property

'
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
    FlowItem_TextLeftMrg = 0
End Property


'Private Property Let FlowItem_TextRightMrg(ByVal RHS As Single)
'    FlowItem.TextRightMrg = RHS
'End Property

Private Property Get FlowItem_TextRightMrg() As Single
    FlowItem_TextRightMrg = 0
End Property


'Private Property Let FlowItem_TextTopMrg(ByVal RHS As Single)
'    FlowItem.TextTopMrg = RHS
'End Property

Private Property Get FlowItem_TextTopMrg() As Single
    FlowItem_TextTopMrg = 0
End Property
