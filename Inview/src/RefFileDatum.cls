VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "RefFileDatum"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Explicit
'RefFileDatum
'Reference File Data
'For use with the RefFile class.
'Richard Fung, Programming Applications 1

Public Keyword      As String
Public Syntax       As String
Public Description  As String

'Removes extraneous characters.
Public Sub RTrim()
    On Error Resume Next
    'remove null characters
    Keyword = Left(Keyword, InStr(1, Keyword, vbNullChar, vbBinaryCompare) - 1)
    Description = Left(Description, InStr(1, Description, vbNullChar, vbBinaryCompare) - 1)
    Syntax = Left(Syntax, InStr(1, Syntax, vbNullChar, vbBinaryCompare) - 1)
    'remove extra spaces
    Keyword = VBA.RTrim(Keyword)
    Description = VBA.RTrim(Description)
    Syntax = VBA.RTrim(Syntax)
End Sub


