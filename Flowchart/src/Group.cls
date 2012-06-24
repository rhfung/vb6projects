VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Group"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'25 August 2002

Private mlngGroupNo As Long
Private mCol        As Collection

Public Sub Add(Ptr As FlowItem)
    'in this group or not in a group
    If Ptr.P.GroupNo = GroupNo Or Ptr.P.GroupNo = 0 Then
        Ptr.P.GroupNo = GroupNo
        mCol.Add Ptr
    End If
End Sub

Public Function Count() As Long
    Count = mCol.Count
End Function

Public Property Get GroupNo() As Long
    GroupNo = mlngGroupNo
End Property

Public Property Let GroupNo(ByVal pVal As Long)
    Dim objItem As FlowItem
    
    mlngGroupNo = pVal
    
    For Each objItem In Me
        objItem.P.GroupNo = pVal
    Next objItem
End Property
    
Public Function Item(n As Long) As FlowItem
Attribute Item.VB_UserMemId = 0
    Set Item = mCol(n)
End Function


Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    Set NewEnum = mCol.[_NewEnum]
End Function

Public Sub Remove(Ptr As FlowItem)
    Dim lngIndex As Long
    
    For lngIndex = mCol.Count To 1 Step -1
        If mCol(lngIndex) Is Ptr Then
            Item(lngIndex).P.GroupNo = 0 'clear group number
            Item(lngIndex).P.Selected = False
            mCol.Remove lngIndex
        End If
    Next lngIndex
End Sub


Public Sub RemoveAll()
    Dim lngIndex As Long
    
    For lngIndex = mCol.Count To 1 Step -1
        Item(lngIndex).P.GroupNo = 0 'clear group number
        Item(lngIndex).P.Selected = False
        mCol.Remove lngIndex
    Next lngIndex
End Sub

Public Property Let Selected(ByVal pSel As Boolean)
    Dim objItem As FlowItem
    
    For Each objItem In Me
        objItem.P.Selected = pSel
    Next objItem
End Property

Private Sub Class_Initialize()
    Set mCol = New Collection
End Sub


Private Sub Class_Terminate()
    Set mCol = Nothing
End Sub




