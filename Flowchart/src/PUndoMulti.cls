VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PUndoMulti"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'26 August 2002

Implements IUndoInterface

Private mCol As Collection

Public Sub Add(Ptr As FlowItem, theActionWas As PUndoAction)
    Dim objNew As PUndoItem
    
    Set objNew = New PUndoItem
    Set objNew.Ptr = Ptr
    CopyProperties Ptr.P, objNew.P
    objNew.P.GroupNo = Ptr.P.GroupNo 'also keep the group number
    objNew.theActionWas = theActionWas
    
    mCol.Add objNew
End Sub

Public Function Count() As Long
    Count = mCol.Count
End Function

Public Function Item(n As Long) As PUndoItem
Attribute Item.VB_UserMemId = 0
    Set Item = mCol(n)
End Function

Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    Set NewEnum = mCol.[_NewEnum]
End Function

Private Sub Class_Initialize()
    Set mCol = New Collection
End Sub

Private Sub Class_Terminate()
    Set mCol = Nothing
End Sub


Private Property Get IUndoInterface_IsMultiple() As Boolean
    IUndoInterface_IsMultiple = True
End Property


