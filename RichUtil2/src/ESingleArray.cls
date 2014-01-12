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
Attribute VB_Description = "Richard Fung Utilities 2 Single Array"
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'/// Richard Fung Utilities 2 Single Array
Option Explicit
Option Base 1
'December 22, 2003
'1 based array

Public LastError    As Long

'local variable(s) to hold property value(s)
Private m_lngCount       As Long 'local copy
Private m_Item()         As Single 'local copy

Private m_lngUpperBound  As Long
Private Const conDefaultValue! = 0!
Private Const conItem = "I"

Public Sub Add(ByVal Item As Single)
Attribute Add.VB_Description = "Adds a member to the array."
'/// Adds a member to the array.
    m_lngCount = m_lngCount + 1
    If m_lngCount > m_lngUpperBound Then
        On Error GoTo SizeErr
        m_lngUpperBound = m_lngUpperBound * 2
        ReDim Preserve m_Item(m_lngUpperBound)
    End If
    
    m_Item(m_lngCount) = Item
    PropertyChanged conItem & m_lngCount
    Exit Sub
SizeErr:
    m_lngUpperBound = &H7FFFFFFF
    Resume Next
End Sub

Public Property Let Count(ByVal vData As Long)
Attribute Count.VB_Description = "Returns the number of elements in the list.  Also can set Count to resize the array."
    Resize vData
End Property


Public Property Get Count() As Long
'/// Returns the number of elements in the list.  Also can set Count to resize the array.
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Count
    Count = m_lngCount
End Property


Public Property Let Item(ByVal Index As Long, ByVal vData As Single)
Attribute Item.VB_Description = "Gets/sets an element of the array, starting at 1."
Attribute Item.VB_UserMemId = 0
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Item = 5
    m_Item(Index) = vData
    PropertyChanged conItem & Index
End Property


Public Property Get Item(ByVal Index As Long) As Single
'/// Gets/sets an element of the array, starting at 1.
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Item
    Item = m_Item(Index)
End Property

Public Sub Remove(ByVal Index As Long)
Attribute Remove.VB_Description = "Removes an item from the array at the given index."
'/// Removes an item from the array at the given index.
    Dim lngAfter  As Long
    
    For lngAfter = Index To m_lngCount - 1
        m_Item(lngAfter) = m_Item(lngAfter + 1)
        PropertyChanged conItem & lngAfter
    Next lngAfter
    PropertyChanged conItem & (m_lngCount)
    
    m_lngCount = m_lngCount - 1
End Sub

Public Sub Resize(ByVal NewSize As Long)
Attribute Resize.VB_Description = "Resizes the array and preserves existing contents."
'/// Resizes the array and preserves existing contents.
    If NewSize > 0 Then
        'increase upper bound if necessary
        If NewSize > m_lngUpperBound Then
            m_lngUpperBound = NewSize
            ReDim Preserve m_Item(NewSize)
        End If
        'resize array
        m_lngCount = NewSize
        
    ElseIf NewSize = 0 Then
        ReDim m_Item(10)
        m_lngCount = 0
        m_lngUpperBound = 10
    Else
        Err.Raise vbObjectError + 1, , "Count must be a non-negative number"
    End If
    PropertyChanged "Count"
End Sub

'Public Function IndexOf(ByVal FindThis As String, Optional CompareMethod As VbCompareMethod = vbBinaryCompare) As Long
''/// Finds a specific String in an array of Strings.  Returns 0 if not found.
'    Dim lngItem As Long
'
'    For lngItem = 1 To m_lngCount
'        If StrComp(FindThis, m_Item(lngItem), CompareMethod) = 0 Then
'            IndexOf = lngItem
'            Exit Function
'        End If
'    Next lngItem
'    IndexOf = 0 'could not find
'End Function

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
        m_Item(lngI) = PropBag.ReadProperty(conItem & lngI, conDefaultValue)
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
        PropBag.WriteProperty conItem & lngI, m_Item(lngI), conDefaultValue
    Next lngI
    Exit Sub
Handler:
    LastError = Err
    Exit Sub
End Sub




