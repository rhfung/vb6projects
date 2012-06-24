VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"IVData"
Option Explicit
'Interface class, IFile
'Other class objects implement this class.
'Richard Fung, Programming Applications 1


Public FileName As String
Public FilePos    As Long 'File position

'Enumeration for function Data().
'Enum IVData
'    conFileNextLine = -1
'    conFileNextByte = 0
'End Enum

Public Sub CloseFile()
    'Closes a file.
End Sub

'Public Function Data(ByVal Number As IVData) As Variant
'    'Data reads sequentially from a file.
'    'Binary: from temporary file, returns binary data
'    'Text:   from main file, returns text
'    'Ref:    from temporary file, returns object RefFileDatum
'End Function

Public Property Get EOF() As Boolean
    'Returns True if the end of file has been reached.
End Property

Public Property Get FileDate() As Date
    'Returns the FileDate.
End Property

Public Property Get FileLength() As Long
    'Returns the length of the file, in bytes.
End Property

Public Property Get FS() As Long
    'Returns the opened File Stream
End Property

Public Property Let Item(ByVal Pos As Long, ByVal Val As Byte)
    'Sets the byte at a specific position.
    'Binary:  Puts a byte into the temporary file.
    'Text:    Sets a value to the loaded text stream.
    'Ref:     Puts a RefFileDatum object to a fixed-width type structure in a file.
End Property

Public Property Get Item(ByVal Pos As Long) As Byte
    'Gets the byte at a specific position.
    'Binary:  Seeks for a byte and returns it from the temporary file.
    'Text:    Reads from the loaded text stream.
    'Ref:     Gets a fixed-width type structure from a file into an object RefFileDatum.
End Property

Public Function SaveFile() As Long
    'Saves the file.
End Function


Public Function OpenFile(ByVal FileName As String) As Long
    'Opens a file, and creates the file stream.
End Function

Private Property Let TemporaryFile(ByVal pTemp As String)
    'For Binary & reference files
End Property

Public Property Get TemporaryFile() As String
    'For Binary & reference files
End Property




