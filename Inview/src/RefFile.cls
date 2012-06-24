VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "RefFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Explicit
'An Inquisite View of Files
'Note: Data in this class is
'      read on demand.
'Richard Fung, Programming Applications 1

Implements IFile

'storage data
Private mstrFileName        As String
Private mstrTemporaryFile   As String
Private mlngFS              As Long

'generates random characters
Private Function RndChr(ByVal Characters As Long) As String
    Dim lngPos As Long
    
    For lngPos = 1 To Characters
        Randomize Timer
        RndChr = RndChr & Chr(Int(Rnd * 26) + 97)
    Next lngPos
End Function

'when terminating, closes any opened file
Private Sub Class_Terminate()
    If mlngFS Then
        IFile_CloseFile
    End If
End Sub

'closes any opened file
Private Sub IFile_CloseFile()
    On Error Resume Next
    Close #mlngFS
    mlngFS = 0
    Kill mstrTemporaryFile
End Sub

'returns the next record, in a type structure, as
'a object RefFileDatum
Private Function IFile_Data(ByVal Number As IVData) As Variant
    Dim tRead As RefFileStruct
    Dim objRead As New RefFileDatum
    
    Select Case Number
        Case conFileNextByte 'binary/text
            Get #mlngFS, , tRead
            objRead.Description = tRead.Description
            objRead.Keyword = tRead.Keyword
            objRead.Syntax = tRead.Syntax
            objRead.RTrim 'clean out extra spaces
            Set IFile_Data = objRead
        Case conFileNextLine 'only for text
            Err.Raise 5
        Case Else
            Err.Raise 5
    End Select
End Function

Private Property Get IFile_EOF() As Boolean
    'end of file is detected by EOF() for Random access mode
    IFile_EOF = EOF(mlngFS)
End Property

'The following are described quite well in the interface class, IFile.
Private Property Get IFile_FileDate() As Date
    IFile_FileDate = FileDateTime(mstrFileName)
End Property

Private Property Get IFile_FileLength() As Long
    If mlngFS Then
        IFile_FileLength = LOF(mlngFS)
    Else
        IFile_FileLength = FileLen(mstrFileName)
    End If
End Property

Private Property Let IFile_FileName(ByVal RHS As String)
    mstrFileName = RHS
End Property

Private Property Get IFile_FileName() As String
    IFile_FileName = mstrFileName
End Property

Private Property Let IFile_FilePos(ByVal RHS As Long)
    Seek #mlngFS, RHS
End Property

Private Property Get IFile_FilePos() As Long
    IFile_FilePos = Loc(mlngFS)
End Property


Private Property Get IFile_FS() As Long
    IFile_FS = mlngFS
End Property

'Returns/sets a RefFileDatum object.
'To do this, internally the object uses a type structure.
'However, the type structure cannot be Public in a file,
'so the object has to be used.
Private Property Let IFile_Item(ByVal Pos As Long, ByVal RHS As Variant)
    Dim tWrite As RefFileStruct
    Dim objWrite As RefFileDatum
    
    Set objWrite = RHS
    tWrite.Description = objWrite.Description
    tWrite.Keyword = objWrite.Keyword
    tWrite.Syntax = objWrite.Syntax
    Put #mlngFS, Pos, tWrite
End Property

Private Property Get IFile_Item(ByVal Pos As Long) As Variant
    Dim tRead As RefFileStruct
    Dim objRead As New RefFileDatum
    
    Get #mlngFS, Pos, tRead
    objRead.Description = tRead.Description
    objRead.Keyword = tRead.Keyword
    objRead.Syntax = tRead.Syntax
    objRead.RTrim
    Set IFile_Item = objRead
End Property

'Opens a copy of the original file, so the original file
'is not modified if SaveFile() is not invoked.
'Also, it only opens files with the extension .refiv
Private Function IFile_OpenFile(ByVal FileName As String) As Long
    Dim tSample As RefFileStruct
    
    If mlngFS Then
        IFile_OpenFile = 5
    Else
        'only opens files with the extension .refiv
        If LCase(Right(FileName, 6)) = ".refiv" Then
            On Error Resume Next
            'create a temporary file and open it
            'in a free file stream
            mlngFS = FreeFile()
            mstrTemporaryFile = FileName & "." & RndChr(5) & ".iv~"
            mstrFileName = FileName
            FileCopy FileName, mstrTemporaryFile
            If Err Then 'file not there
                Err.Clear
            End If
            Open mstrTemporaryFile For Random As #mlngFS Len = Len(tSample)
            IFile_OpenFile = Err
        ElseIf FileName = "" Then 'new file
            On Error Resume Next
            mlngFS = FreeFile()
            mstrTemporaryFile = RndChr(8) & "." & RndChr(5) & ".iv~"
            mstrFileName = FileName
            If Err Then 'file not there
                Err.Clear
            End If
            Open mstrTemporaryFile For Random As #mlngFS Len = Len(tSample)
            IFile_OpenFile = Err
        Else
            IFile_OpenFile = 5 'cannot open other kinds of files
        End If
    End If
End Function

'Copies the temporary file back to the original file while
'the temporary file is closed.  Then, the temporary file is reopened.
Private Function IFile_SaveFile() As Long
    Dim lngPos As Long
    Dim tSample As RefFileStruct
    
    lngPos = IFile_FilePos()
    Close #mlngFS
    FileCopy mstrTemporaryFile, mstrFileName
    Open mstrTemporaryFile For Random As #mlngFS Len = Len(tSample)
    If lngPos > 0 Then 'record value should be good
        IFile_FilePos = lngPos
    End If
End Function

Private Property Get IFile_TemporaryFile() As String
    IFile_TemporaryFile = mstrTemporaryFile
End Property

'These are not implemented in the Reference file records.
Private Property Let IFile_TextStream(ByVal RHS As String)
    'blank
End Property

Private Property Get IFile_TextStream() As String
    'blank
End Property

