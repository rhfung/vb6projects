VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Header"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Attribute VB_Ext_KEY = "Member0" ,"Thoughts"
Option Explicit

Public ColourScheme As Integer
Public Author As String
Public EditTextWindow As Boolean
Public Comment As String
Public Circled As Boolean

'local variable(s) to hold property value(s)
Private mvarDateModified As Date 'local copy
Private mvarLastSelected As Integer 'local copy
'local variable(s) to hold property value(s)
Private mvarTag As String 'local copy










Public Property Get Tag() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Tag
    Tag = mvarTag
End Property





Public Sub ExtraData(ByVal LastSelected As Integer, ByVal DateModified As Date, ByVal Tag As String)
    mvarLastSelected = LastSelected
    mvarDateModified = DateModified
End Sub

Public Property Get LastSelected() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.LastSelected
    LastSelected = mvarLastSelected
End Property



Public Property Get DateModified() As Date
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.DateModified
    DateModified = mvarDateModified
End Property


