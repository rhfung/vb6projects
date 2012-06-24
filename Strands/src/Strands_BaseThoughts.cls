VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "BaseThoughts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"Thought"
Attribute VB_Ext_KEY = "Member0" ,"Header"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "Member1" ,"Thought"
Option Explicit
'Richard Fung.
'March 30, 2000.
'BaseThoughts cls

'local variable to hold collection
Private mCol As Collection
Private mvarHeader As Header


Public Property Get Header() As Header
    Set Header = mvarHeader
End Property


Public Property Set Header(vData As Header)
    Set mvarHeader = vData
End Property

Public Function Find(ByVal Obj As Object) As Thought
    Dim objI As Thought
    
    For Each objI In mCol
        If Obj Is objI Then
            Set Find = objI
        End If
    Next objI
End Function

Public Function FindNum(ByVal Index As Integer) As Thought
    Dim objI As Thought
    
    For Each objI In mCol
        If objI.Index = Index Then
            Set FindNum = objI
        End If
    Next objI
End Function

Public Function Add(Idea As String, Text As String, Attachment As String, AttachmentTag As String, CenterX As Single, CenterY As Single, Optional Obj As Object, Optional sKey As String) As Thought
    'create a new object
    Dim objNewMember As Thought
    Set objNewMember = New Thought

    'set the properties passed into the method
    objNewMember.Idea = Idea
    objNewMember.Text = Text
    objNewMember.Attachment = Attachment
    objNewMember.AttachmentTag = AttachmentTag
    objNewMember.CenterX = CenterX
    objNewMember.CenterY = CenterY
    objNewMember.Index = Count + 1
    Set objNewMember.Links = New Thoughts

    If Not Obj Is Nothing Then
        objNewMember.Refresh Obj
    End If

    If Len(sKey) = 0 Then
        mCol.Add objNewMember
    Else
        mCol.Add objNewMember, sKey
    End If

    'return the object created
    Set Add = objNewMember
    Set objNewMember = Nothing
End Function





Public Property Get Item(vntIndexKey As Variant) As Thought
Attribute Item.VB_UserMemId = 0
    'used when referencing an element in the collection
    'vntIndexKey contains either the Index or Key to the collection,
    'this is why it is declared as a Variant
    'Syntax: Set foo = x.Item(xyz) or Set foo = x.Item(5)
  Set Item = mCol(vntIndexKey)
End Property



Public Property Get Count() As Long
    'used when retrieving the number of elements in the
    'collection. Syntax: Debug.Print x.Count
    Count = mCol.Count
End Property



Public Sub Remove(vntIndexKey As Variant)
    'used when removing an element from the collection
    'vntIndexKey contains either the Index or Key, which is why
    'it is declared as a Variant
    'Syntax: x.Remove(xyz)
    Dim objItem As Thought
    Dim intCount As Integer
    
    mCol.Remove vntIndexKey
        
    'renumber all items
    intCount = 0
    For Each objItem In mCol
        intCount = intCount + 1
        objItem.Index = intCount
    Next objItem
End Sub


Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    'this property allows you to enumerate
    'this collection with the For...Each syntax
    Set NewEnum = mCol.[_NewEnum]
End Property


Private Sub Class_Initialize()
    'creates the collection when this class is created
    Set mCol = New Collection
    'create the mHeader object when the Thoughts class is created
    Set mvarHeader = New Header
End Sub


Private Sub Class_Terminate()
    Set mvarHeader = Nothing
    Set mCol = Nothing
End Sub


