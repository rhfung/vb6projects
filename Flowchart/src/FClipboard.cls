VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PClipboard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"FlowItem"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
'FClipboard March 2, 2001

'for popup menus
Public TempX As Long, TempY As Long

'local variable(s) to hold property value(s)
Private mvarSnapshot  As FlowItem 'local copy
Private mvarLastAdded As FlowItem 'local copy
Private mcolSelection As Collection

Public Sub Clear()
    Set mvarLastAdded = Nothing
    Set mvarSnapshot = Nothing
    Set mcolSelection = Nothing
End Sub


Public Sub ClearLink() 'For FlowItem
    Set mvarLastAdded = Nothing
End Sub

Public Function IsItem() As Boolean
    IsItem = Not mvarSnapshot Is Nothing
End Function


Public Function IsSelection() As Boolean
    If mcolSelection Is Nothing Then
        'nothing to collection
        IsSelection = False
    Else
        'something in collection
        IsSelection = mcolSelection.Count > 0
    End If
End Function

Public Function IsText() As Boolean
    IsText = Clipboard.GetFormat(vbCFText)
End Function

Public Property Set FlowItem(ByVal vData As FlowItem)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.FlowItem = Form1
    Clear

    Set mvarSnapshot = Duplicate(vData)
    Set mvarLastAdded = vData
End Property


Public Property Get FlowItem() As FlowItem
    Set FlowItem = Duplicate(mvarSnapshot)  'always follow last item
    If Not mvarLastAdded Is Nothing Then
        FlowItem.P.Left = mvarLastAdded.P.Left + mvarLastAdded.P.Width * 0.25
        FlowItem.P.Top = mvarLastAdded.P.Top + mvarLastAdded.P.Height * 0.25
    End If
    Set mvarLastAdded = FlowItem
End Property



Public Property Get Selection() As Collection
    Set Selection = mcolSelection
End Property

Public Property Set Selection(ByVal col As Collection)
    Dim objItem As FlowItem
    
    Clear
    Set mcolSelection = New Collection
    For Each objItem In col
        mcolSelection.Add Duplicate(objItem)
    Next objItem
End Property

'An individual flow item differs from a get FlowItem()
'because the last added item is not followed.
Public Sub SetFreeFlowItem(ByVal FlowItem As FlowItem)
    Set mvarSnapshot = Duplicate(FlowItem)
    Set mvarLastAdded = Nothing
    Set mcolSelection = Nothing
End Sub


Public Sub TempClear()
    TempX = 0
    TempY = 0
End Sub

Public Property Let Text(ByVal pText As String)
    Clipboard.Clear
    Clipboard.SetText pText
End Property

Public Property Get Text() As String
    Text = Clipboard.GetText()
End Property

Private Sub Class_Terminate()
    Clear
End Sub


