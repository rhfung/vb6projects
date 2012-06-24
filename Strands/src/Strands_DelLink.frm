VERSION 5.00
Begin VB.Form frmDelLink 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Delete Link"
   ClientHeight    =   3480
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   Icon            =   "Strands_DelLink.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3480
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Default         =   -1  'True
      Height          =   375
      Left            =   2760
      TabIndex        =   2
      Top             =   3000
      Width           =   1095
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Delete"
      Height          =   375
      Left            =   600
      TabIndex        =   1
      Top             =   3000
      Width           =   1095
   End
   Begin VB.ListBox lstLink 
      Height          =   2760
      IntegralHeight  =   0   'False
      ItemData        =   "Strands_DelLink.frx":000C
      Left            =   120
      List            =   "Strands_DelLink.frx":000E
      Style           =   1  'Checkbox
      TabIndex        =   0
      Top             =   120
      Width           =   4455
   End
End
Attribute VB_Name = "frmDelLink"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'April 15, 2000

Public Thoughts As BaseThoughts
Public LinkItem As Thought
Public Links    As Thoughts
Public Sub ReadLinks()
    Dim objItem As Thought
    Dim objConnect As Thought
    
    'read links & put in Links
    Set Links = New Thoughts
    For Each objItem In Thoughts
        For Each objConnect In objItem.Links
            If objConnect Is LinkItem Then
                Links.Add objItem
            ElseIf objItem Is LinkItem Then
                Links.Add objConnect
            End If
        Next objConnect
    Next objItem
    
    'show in list
    lstLink.Clear
    For Each objItem In Links
        If Len(objItem.Text) Then
            lstLink.AddItem objItem & ": " & objItem.Text
        Else
            lstLink.AddItem objItem
        End If
    Next objItem
    
    cmdDelete.Enabled = lstLink.ListCount
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub




Private Sub cmdDelete_Click()
    Dim intIndex    As Integer
    Dim colDelete   As Thoughts
    Dim objObject   As Thought
    Dim objItem     As Thought
    Dim objRemove   As Thought

    'based on Links
    Set colDelete = New Thoughts
    For intIndex = 1 To Links.Count
        If lstLink.Selected(intIndex - 1) Then
            'to be deleted
            colDelete.Add Links(intIndex)
        End If
    Next intIndex
    
    'hunt down those links to delete
    For Each objObject In Thoughts
        If objObject Is LinkItem Then 'object is me, good
            For Each objItem In objObject.Links 'go through link list
                Set objRemove = colDelete.Find(objItem) 'find this item if need be deleted
                If Not objRemove Is Nothing Then
                    objObject.Links.Remove Object:=objRemove
                End If
            Next objItem
        Else
            Set objItem = colDelete.Find(objObject) 'is this having reference to me?
            If Not objItem Is Nothing Then
                Set objRemove = objObject.Links.Find(LinkItem) 'find where I'm attached (because sometimes, I won't be attached)
                If Not objRemove Is Nothing Then
                    objObject.Links.Remove Object:=objRemove 'objRemove is also LinkItem
                End If
            End If
        End If
    Next objObject

    Set colDelete = Nothing
    cmdClose.Caption = "Close" 'change button face
    ReadLinks
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set LinkItem = Nothing
    Set Thoughts = Nothing
    Set Links = Nothing
End Sub


