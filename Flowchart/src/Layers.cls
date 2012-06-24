VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Layers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'April 26 2003
'Create this object only when you want to use
'information about all the layers.

Private m_Layers(0 To 255)  As Layer
Private m_File              As FlowChart
Private m_DefaultShapes(0 To 255) As FlowItem
Private m_MaxDefaultShape         As Long

Public Sub CloseQuery()
    Erase m_Layers
End Sub

Public Property Set DefaultShape(ByVal Layer As Long, ByVal pVal As FlowItem)
    If 0 <= Layer And Layer <= 255 Then
        Set m_DefaultShapes(Layer) = pVal
        If Layer > m_MaxDefaultShape Then m_MaxDefaultShape = Layer
    End If
End Property

Public Property Get DefaultShape(ByVal Layer As Long) As FlowItem
    If m_DefaultShapes(Layer) Is Nothing Then
       Set DefaultShape = New FlowItem
    Else
        Set DefaultShape = m_DefaultShapes(Layer)
    End If
End Property


Public Property Let EnableAll(ByVal pEA As Boolean)
    Dim objA As FlowItem
    
    For Each objA In m_File
        objA.P.Enabled = pEA
    Next objA
End Property


Public Sub Initialize(ByVal File As FlowChart)
    Set m_File = File
End Sub

Public Property Get MaxDefaultShape() As Long
    MaxDefaultShape = m_MaxDefaultShape
End Property

Public Sub Requery()
    Dim lngI As Long
    
    For lngI = 0 To 255
        Set m_Layers(lngI) = New Layer
        m_Layers(lngI).LayerNo = lngI
    Next lngI
    
    Dim objItem  As FlowItem
    
    For Each objItem In m_File
        If 0 <= objItem.P.Layer And objItem.P.Layer <= 255 Then
             m_Layers(objItem.P.Layer).Add objItem
        End If
    Next objItem
End Sub




Public Function Layer(ByVal LayerNo As Long) As Layer
Attribute Layer.VB_UserMemId = 0
    If m_Layers(LayerNo) Is Nothing Then 'not query
        Dim objLayer As Layer
        Dim objItem  As FlowItem
        
        Set objLayer = New Layer 'find the layer from file
        
        For Each objItem In m_File
            If objItem.P.Layer = LayerNo Then
                objLayer.Add objItem
            End If
        Next objItem
        
        Set m_Layers(LayerNo) = objLayer 'store query
        Set Layer = objLayer
    Else 'in query
        Set Layer = m_Layers(LayerNo)
    End If
End Function

'Public Function Item(n As Long) As Group
'    Set Item = mCol(n)
'End Function


Public Sub MinMaxLayer(ptrMin As Long, ptrMax As Long)
    Dim objLayer As Layer
    Dim lngItem  As Long
    
    ptrMin = 0 'means that there is nothing in the group
    ptrMax = 0
    
  
    For lngItem = 0 To 255
        Set objLayer = m_Layers(lngItem)
        If Not objLayer Is Nothing Then
            If objLayer.LayerNo < ptrMin And objLayer.Count > 0 Then ptrMin = objLayer.LayerNo
            If objLayer.LayerNo > ptrMax And objLayer.Count > 0 Then ptrMax = objLayer.LayerNo
        End If
    Next lngItem
End Sub








Private Sub Class_Initialize()
    Dim lngI As Long
    
    For lngI = 0 To 14
        Set m_DefaultShapes(lngI) = New FlowItem
        m_DefaultShapes(lngI).P.ForeColour = QBColor(lngI)
    Next lngI
End Sub


Private Sub Class_Terminate()
    Set m_File = Nothing
End Sub


