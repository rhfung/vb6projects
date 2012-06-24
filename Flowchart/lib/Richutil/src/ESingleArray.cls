VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 1  'Persistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ESingleArray"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Option Base 1
'January 2, 2002
'1 based array

Public LastError    As Long

'local variable(s) to hold property value(s)
Private m_lngCount       As Long 'local copy
Private m_Item()         As Single 'local copy

Private m_lngUpperBound  As Long
Private Const conDefaultValue As Single = 0!
Public Property Let Item(ByVal Index As Long, ByVal vData As Single)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Item = 5
    m_Item(Index) = vData
    PropertyChanged "Item" & Index
End Property


Public Property Get Item(ByVal Index As Long) As Single
Attribute Item.VB_UserMemId = 0
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Item
    Item = m_Item(Index)
End Property



Public Property Let Count(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Count = 5
    If vData > 0 Then
        'increase upper bound if necessary
        If vData > m_lngUpperBound Then
            m_lngUpperBound = vData
        End If
        'resize array
        m_lngCount = vData
        ReDim Preserve m_Item(vData)
    ElseIf vData = 0 Then
        ReDim m_Item(1)
        m_lngCount = 0
        m_lngUpperBound = 1
    Else
        Err.Raise vbObjectError + 1, , "Count must be a non-negative number"
    End If
    PropertyChanged "Count"
End Property


Public Property Get Count() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Count
    Count = m_lngCount
End Property



Public Sub Remove(ByVal Index As Long)
    Dim lngAfter  As Long
    
    For lngAfter = Index To m_lngCount - 1
        m_Item(lngAfter) = m_Item(lngAfter + 1)
        PropertyChanged "Item" & lngAfter
    Next lngAfter
    PropertyChanged "Item" & (m_lngCount)
    
    m_lngCount = m_lngCount - 1
End Sub


Public Sub Add(ByVal Item As Single)
    m_lngCount = m_lngCount + 1
    If m_lngCount > m_lngUpperBound Then
        m_lngUpperBound = m_lngUpperBound + 10
        ReDim Preserve m_Item(m_lngUpperBound)
    End If
    
    m_Item(m_lngCount) = Item
    PropertyChanged "Item" & m_lngCount
End Sub



Private Sub Class_Initialize()
    Count = 0
End Sub

Private Sub Class_InitProperties()
    Count = 0
End Sub


Private Sub Class_ReadProperties(PropBag As PropertyBag)
    Dim lngI As Long
    
    On Error GoTo Handler
    Count = PropBag.ReadProperty("Count", 0)
    
    For lngI = 1 To m_lngCount
        m_Item(lngI) = PropBag.ReadProperty("Item" & lngI, conDefaultValue)
    Next lngI
    
    Exit Sub
Handler:
    LastError = Err
    Exit Sub
End Sub


Private Sub Class_Terminate()
    Erase m_Item
End Sub

Private Sub Class_WriteProperties(PropBag As PropertyBag)
    Dim lngI As Long
    
    On Error GoTo Handler
    
    PropBag.WriteProperty "Count", m_lngCount, 0
    For lngI = 1 To m_lngCount
        PropBag.WriteProperty "Item" & lngI, m_Item(lngI), conDefaultValue
    Next lngI
    Exit Sub
Handler:
    LastError = Err
    Exit Sub
End Sub


