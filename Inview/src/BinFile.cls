VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BinFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'An Inquisite View of Files
'Binary File
'Richard Fung, Programming Applications 1

'Note: Data in this class is read into the buffer.
'See Rebuffer() for size of buffer.

'Implements IFile

'storage variables
Private mstrFileName        As String
Private mstrTemporaryFile   As String
Private mlngFS              As Long
Private mlngBufferFirstByte As Long
Private mlngBufferLen       As Long
Private mbytBuffer()        As Byte
Private mlngPos             As Long

Private Sub Rebuffer(ByVal Pos As Long)
    Dim lngLen As Long
    
    lngLen = 30720  '30 KB
    If Pos + lngLen > FileLength Then
        lngLen = FileLength - Pos
    End If
    If lngLen >= 0 Then
        mlngBufferLen = lngLen
        ReDim mbytBuffer(Pos To Pos + mlngBufferLen) As Byte
        mlngBufferFirstByte = Pos
        Get #mlngFS, Pos, mbytBuffer
        'Debug.Print "Rebuffered at "; Pos; " to "; mlngBufferFirstByte + mlngBufferLen
    End If
End Sub

''creates a random number of characters
'Private Function RndChr(ByVal Characters As Long) As String
'    Dim lngPos As Long
'
'    For lngPos = 1 To Characters
'        Randomize Timer
'        RndChr = RndChr & Chr(Int(Rnd * 26) + 97)
'    Next lngPos
'End Function

Public Sub Class_Initialize()
    mlngBufferLen = -1024
End Sub

'when the class is terminated, close any opened files
Public Sub Class_Terminate()
    If mlngFS Then
        CloseFile
    End If
End Sub

'closes an opened file
Public Sub CloseFile()
    On Error Resume Next
    Close #mlngFS
    mlngFS = 0
    Kill mstrTemporaryFile
End Sub

'Returns the next "data" from the file stream.
'Public Function Data(ByVal Number As IVData) As Variant
'    Select Case Number
'        Case conFileNextByte 'binary/text
'            Dim bytDatum As Byte
'            If mlngBufferFirstByte >= FilePos And FilePos <= mlngBufferFirstByte + mlngBufferLen Then
'                bytDatum = mbytBuffer(FilePos)
'            Else
'                'Get #mlngFS, , bytDatum
'                If FilePos = 0 Then Seek #mlngFS, 1
'                bytDatum = Item(FilePos)
'            End If
'            Data = bytDatum
'            Seek #mlngFS, FilePos + 1
'        Case conFileNextLine 'only for text
'            Err.Raise 5
'        Case Else
'            Dim bytData() As Byte 'binary
'            ReDim bytData(1 To Number)
'            Get #mlngFS, , bytData
'            Data = bytData
'    End Select
'End Function

Public Property Get EOF() As Boolean
    'if the file position is greater than the length
    'of the file, the end of the file has been reached
    EOF = (FilePos >= FileLength)
End Property

'The following are described quite well in the interface class, IFile.
Public Property Get FileDate() As Date
    FileDate = FileDateTime(mstrFileName)
End Property

Public Property Get FileLength() As Long
    If mlngFS Then
        FileLength = LOF(mlngFS)
    Else
        FileLength = FileLen(mstrFileName)
    End If
End Property

Public Property Let Filename(ByVal RHS As String)
    mstrFileName = RHS
End Property

Public Property Get Filename() As String
    Filename = mstrFileName
End Property

Public Property Let FilePos(ByVal RHS As Long)
    'Seek #mlngFS, RHS
    mlngPos = RHS
End Property

Public Property Get FilePos() As Long
    FilePos = mlngPos 'Loc(mlngFS)
End Property

Public Property Get FS() As Long
    FS = mlngFS
End Property

'Item() gets/sets bytes in the binary file.
Public Property Let Item(ByVal Pos As Long, ByVal RHS As Byte)
    Put #mlngFS, Pos, CByte(RHS)
    If Pos >= mlngBufferFirstByte And Pos <= mlngBufferFirstByte + mlngBufferLen Then 'in buffer
        mbytBuffer(Pos) = RHS
    End If
End Property

Public Property Get Item(ByVal Pos As Long) As Byte
Attribute Item.VB_UserMemId = 0
Attribute Item.VB_MemberFlags = "200"
'    Dim bytDatum As Byte
    
    
    If Pos >= mlngBufferFirstByte And Pos <= mlngBufferFirstByte + mlngBufferLen Then
        Item = mbytBuffer(Pos)
    Else
        Rebuffer Pos
        If Pos <= FileLength Then Item = mbytBuffer(Pos)
    End If
    mlngPos = Pos
'    Get #mlngFS, Pos, bytDatum
'    Item = bytDatum
End Property

'Opens a temporary file that is a copy from the original file.
Public Function OpenFile(ByVal Filename As String) As Long
    If mlngFS Then
        OpenFile = 5
    ElseIf Filename = "" Then
        On Error Resume Next
        mlngFS = FreeFile()
        mstrTemporaryFile = GetTempFile("Untitled") 'Filename & "." & RndChr(5) & ".iv~"
        mstrFileName = Filename
        
        Open mstrTemporaryFile For Binary As #mlngFS
        OpenFile = Err
    Else
        On Error Resume Next
        mlngFS = FreeFile()
        mstrTemporaryFile = GetTempFile(Filename) 'Filename & "." & RndChr(5) & ".iv~"
        mstrFileName = Filename
        FileCopy Filename, mstrTemporaryFile
        Open mstrTemporaryFile For Binary As #mlngFS
        OpenFile = Err
    End If
End Function

'Copies the temporary file back to the original file while
'the temporary file is closed.  Then, the temporary file is reopened.
Public Function SaveFile() As Long
    Dim lngPos As Long
    
    lngPos = FilePos()
    Close #mlngFS
    FileCopy mstrTemporaryFile, mstrFileName

    Open mstrTemporaryFile For Binary As #mlngFS
    FilePos = lngPos
End Function

Public Property Get TemporaryFile() As String
    TemporaryFile = mstrTemporaryFile
End Property

