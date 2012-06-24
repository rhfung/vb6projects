VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Groups"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'25 August 2002
'n - index

Private mCol As Collection

Public Sub Add(ByVal Ptr As FlowItem, ByVal GroupNo As Long)
    Dim objGroup As Group
    
    'try to get existing group with the same number
    Set objGroup = Group(GroupNo)
    
    'if cannot find a group, create a new group
    If objGroup Is Nothing Then
        Set objGroup = New Group
        mCol.Add objGroup
        objGroup.GroupNo = GroupNo
    End If
    
    'if already in an existing group, and that group is not this group
    If Ptr.P.GroupNo <> 0 And Ptr.P.GroupNo <> GroupNo Then
        'remove this from the other group
        Dim objOtherGroup As Group
        
        Set objOtherGroup = Group(Ptr.P.GroupNo)
        'found the other group
        If Not objOtherGroup Is Nothing Then
            objOtherGroup.Remove Ptr
        Else
        'invalid group number - does not reference to anything
            Ptr.P.GroupNo = 0 'reset group number to 0
        End If
    End If
    'add the object to the new group
    objGroup.Add Ptr
End Sub

Public Function Count() As Long
    Count = mCol.Count
End Function


Public Function Group(ByVal GroupNo As Long) As Group
Attribute Group.VB_UserMemId = 0
    Dim objGroup As Group
    
    For Each objGroup In Me
        If objGroup.GroupNo = GroupNo Then
            Set Group = objGroup
            Exit Function
        End If
    Next objGroup
End Function

'Public Function Item(n As Long) As Group
'    Set Item = mCol(n)
'End Function


Public Sub MinMaxGroup(ptrMin As Long, ptrMax As Long)
    Dim objGroup As Group
    
    ptrMin = 0 'means that there is nothing in the group
    ptrMax = 0
    
    For Each objGroup In Me
        If objGroup.GroupNo < ptrMin Then ptrMin = objGroup.GroupNo
        If objGroup.GroupNo > ptrMax Then ptrMax = objGroup.GroupNo
    Next objGroup
End Sub

Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    Set NewEnum = mCol.[_NewEnum]
End Function

Public Sub Remove(Ptr As FlowItem)
    Dim objGroup As Group
    
    Set objGroup = Group(Ptr.P.GroupNo)
    If Not objGroup Is Nothing Then
        objGroup.Remove Ptr
    End If
End Sub

'Public Sub RemoveGroup(GroupNo As Long)
'    Dim objGroup    As Group
'    Dim lngIndex    As Long
'
'    Set objGroup = Group(GroupNo) 'find group
'
'    If Not objGroup Is Nothing Then
'        objGroup.GroupNo = 0 'clear group number
'        objGroup.Selected = False
'        'search for the group index in this collection & remove it
'        For lngIndex = Count To 1 Step -1
'            If objGroup Is mCol(lngIndex) Then
'                mCol.Remove lngIndex 'remove index from this collection
'            End If
'        Next lngIndex
'    End If
'End Sub

Private Sub Class_Initialize()
    Set mCol = New Collection
End Sub


Private Sub Class_Terminate()
    Set mCol = Nothing
End Sub


