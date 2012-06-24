VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Thought"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Option Explicit
'Richard Fung.
'March 30, 2000.
'Thought

Public Idea     As String
Attribute Idea.VB_VarUserMemId = 0
Public Text     As String
Public Attachment As String
Public AttachmentTag As String
Public Width    As Single
Public Height   As Single
Public Links    As Thoughts
Public Index    As Integer
Public Left     As Single
Public Top      As Single
Public Picture  As String
Public Tag As String
'local variable(s) to hold property value(s)
Private mvarCenterX As Single 'local copy
Private mvarCenterY As Single 'local copy
Public Property Let CenterY(ByVal vData As Single)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.CenterY = 5
    mvarCenterY = vData
End Property


Public Property Get CenterY() As Single
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.CenterY
    CenterY = mvarCenterY
    If mvarCenterY < 0 Then mvarCenterY = 0
    If mvarCenterY > conHeight Then mvarCenterY = conHeight
End Property



Public Property Let CenterX(ByVal vData As Single)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.CenterX = 5
    mvarCenterX = vData
    If mvarCenterX < 0 Then mvarCenterX = 0
    If mvarCenterX > conWidth Then mvarCenterX = conWidth
End Property


Public Property Get CenterX() As Single
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.CenterX
    CenterX = mvarCenterX
End Property





'Public Sub Destroy()
'    Set Links = Nothing
'End Sub

Public Property Get LinkList() As String
    Dim objI As Thought
    
    For Each objI In Links
        If LinkList = "" Then
            LinkList = objI.Index
        Else
            LinkList = LinkList & "," & objI.Index
        End If
    Next objI
End Property

Public Sub Refresh(ByVal Obj As Object)
    Width = Obj.TextWidth(Idea)
    Height = Obj.TextHeight(Idea)
    Left = CenterX - Width / 2
    Top = CenterY - Height / 2
End Sub


Public Function OverThis(ByVal x As Single, ByVal y As Single) As Boolean
    OverThis = (Left <= x And x <= Left + Width) And (Top <= y And y <= Top + Height)
End Function



Private Sub Class_Terminate()
    'Debug.Print "Destroyed class - " & Idea & " : " & Text
    Set Links = Nothing
End Sub

