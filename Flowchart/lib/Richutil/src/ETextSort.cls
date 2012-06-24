VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ETextSort"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Option Compare Text 'for sorting without case preference
'August 20 2002
'Retreived from FlowChart
'Sorts text or anything you want.

Private colWords As Collection

Public Function Count() As Long
    Count = colWords.Count
End Function

Public Sub Insert(ByVal V As Variant)
    Dim i As Long, j As Long, k As Long
    
    If colWords.Count = 0 Then
        colWords.Add V
        Exit Sub
    End If
    
    i = 1
    j = colWords.Count
    
    If V <= colWords(i) Then
        colWords.Add V, , i
        Exit Sub
    End If
    
    If V >= colWords(j) Then
        colWords.Add V, , , j
        Exit Sub
    End If
    
    Do Until j - i <= 1
        k = (i + j) / 2
        If colWords(k) < V Then
            i = k
        Else
            j = k
        End If
    Loop
    
    colWords.Add V, , j
End Sub


Public Function Item(ByVal Index As Long) As String
    Item = colWords(Index)
End Function


Public Function NewEnum() As IUnknown
    Set NewEnum = colWords.[_NewEnum]
End Function


Private Sub Class_Initialize()
    Set colWords = New Collection
End Sub



Private Sub Class_Terminate()
    Set colWords = Nothing
End Sub


