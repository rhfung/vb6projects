VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TextFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'An Inquisite View of Files
'Note: data in this class is loaded on OpenFile.
'      data saved in mstrText variable
'Richard Fung, Programming Applications 1

Public CanSave              As Boolean

'storage variables
Private mstrFileName        As String
Private mstrTemporaryFile   As String
Private mlngFS              As Long
Private mblnNotUnicode      As Boolean
'text stream variable
Private mstrText            As String
Private mlngBlocks          As Long
Private mlngBlockLen        As Long
Private mlngBlock           As Long

Public Property Get Block() As Long
    Block = mlngBlock
End Property

Public Property Let Block(ByVal Block As Long)
    Dim bytRead()   As Byte
    Dim lngFirst    As Long
    Dim lngLast     As Long
    Dim lngLen      As Long
    
    mlngBlock = Block
    lngFirst = (Block - 1) * mlngBlockLen + 1
    'Seek #FS, lngFirst
    lngLast = lngFirst + mlngBlockLen
    If lngLast > FileLength Then
        lngLast = FileLength
        lngLen = lngLast - lngFirst + 1
    Else
        lngLen = mlngBlockLen
    End If
    ReDim bytRead(1 To lngLen)
    Get #FS, lngFirst, bytRead
    If mblnNotUnicode Then
        Text = StrConv(bytRead, vbUnicode)
    Else 'unicode
        Text = CStr(bytRead)
    End If
End Property



Public Property Get Blocks() As Long
    Blocks = mlngBlocks
End Property

Public Function GetRealPos(ByVal BlockPos As Long) As Long
    GetRealPos = (Block - 1) * mlngBlockLen + BlockPos
End Function

Public Property Let IsUnicode(ByVal pValue As Boolean)
    mblnNotUnicode = Not pValue
End Property

'closes any opened files
Private Sub Class_Terminate()
    If mlngFS Then
        CloseFile
    End If
End Sub

'releases the file stream
Public Sub CloseFile()
    Close #mlngFS
    mlngFS = 0
    Kill mstrTemporaryFile
End Sub

''Returns the next set of strings from the file,
''not including Chr(0).
'Public Function Data(ByVal Number As IVData) As Variant
'    Dim strData As String
'    Dim bytDatum As Byte
'    Dim lngLoop As Long
'
'    Select Case Number
'        Case conFileNextByte 'binary/text
'            'Get #mlngFS, , bytDatum
'            'Data = Chr(bytDatum)
'        Case conFileNextLine 'only for text
'            'Do
'            '    Get #mlngFS, , bytDatum
'            '    If bytDatum = 13 Then 'new line
'            '        Exit Do
'            '    Else
'            '        strData = strData & Chr(bytDatum)
'            '    End If
'            'Loop
'            Line Input #mlngFS, strData
'            Data = strData
'        Case Else 'Returns the specified number of characters
'            Data = Input(Number, #mlngFS)
''            Dim bytData() As Byte 'binary
''            ReDim bytData(1 To Number)
''            Get #mlngFS, , bytData
''            For lngLoop = LBound(bytData) To UBound(bytData)
''                strData = strData & Chr(bytData(lngLoop))
''            Next lngLoop
''            Data = strData
'    End Select
'End Function
'
'The following are described quite well in the interface class, IFile.
Public Property Get EOF() As Boolean
    'the EOF() function returns True if the
    'end of file has been reached
    EOF = VBA.EOF(mlngFS)
End Property

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


Public Property Get FS() As Long
    FS = mlngFS
End Property


Public Property Get IsUnicode() As Boolean
    IsUnicode = Not mblnNotUnicode
End Property


'Opens the original file and copies all the data in the
'file into a String, which is the text stream.
'All changes are made through the text stream.
Public Function OpenFile(ByVal Filename As String) As Long
'    Dim lngChunk    As Long
    Dim lngFileLen  As Long
    Dim lngPos      As Long
    Dim lngIs0      As Long
    Dim bytRead()   As Byte
    Dim blnASCII    As Boolean
    
    If mlngFS Then
        OpenFile = 5
    ElseIf Filename = "" Then
        On Error Resume Next
        mlngFS = FreeFile()
        mstrTemporaryFile = GetTempFile("Untitled")
        mstrText = ""
        mstrFileName = Filename
        mblnNotUnicode = False 'so it is unicode by default
        Open mstrTemporaryFile For Binary As #mlngFS
        'FilePos = 1 'set to start
    Else
        On Error Resume Next
        mlngFS = FreeFile()
        mstrTemporaryFile = GetTempFile(Filename)
        mstrText = ""
        mstrFileName = Filename
        
        FileCopy Filename, mstrTemporaryFile
        Open mstrTemporaryFile For Binary As #mlngFS
        lngFileLen = LOF(mlngFS)
'        If lngFileLen > 65536 Then  '64 KB
'            'OpenFile = -4 'too large
'            'CloseFile
'            mlngBlocks = FileLength \ 65536
'            mlngBlockLen = 65536
'            If FileLength Mod 65536 > 0 Then
'                mlngBlocks = mlngBlocks + 1
'            End If
'            mblnNotUnicode = False
        If lngFileLen > 0 Then
            If lngFileLen <= 65536 Then '64 KB
                ReDim bytRead(1 To lngFileLen)
                Get #mlngFS, , bytRead
                'test the file
                For lngPos = 2 To lngFileLen Step 2
                    If bytRead(lngPos) = 0 Then lngIs0 = lngIs0 + 1
                Next lngPos
                blnASCII = lngIs0 < lngFileLen / 3
            Else
                ReDim bytRead(1 To 65536)
                Get #mlngFS, , bytRead
                'test the file
                For lngPos = 2 To 65536 Step 2
                    If bytRead(lngPos) = 0 Then lngIs0 = lngIs0 + 1
                Next lngPos
                blnASCII = lngIs0 < 65536 / 3
            End If
            
            If blnASCII Then 'is ASCII?
                If MsgBox("Inside File Viewer thinks that the file '" & Filename & "' is an ASCII file.  Attempt to open as ASCII?", vbQuestion Or vbYesNo) = vbYes Then
                    If lngFileLen > 32768 Then '32 KB
                        mlngBlocks = FileLength \ 32768
                        mlngBlockLen = 32768
                        If FileLength Mod 32768 > 0 Then
                            mlngBlocks = mlngBlocks + 1
                        End If
                        mblnNotUnicode = True
                        Block = 1
                        CanSave = False
                        'OpenFile = -4
                        'CloseFile
                    Else
                        'ASCII File
                        mstrText = StrConv(bytRead, vbUnicode)
                        mblnNotUnicode = True
                    End If
                Else
                    GoSub UnicodeFile
                End If
            Else 'lots of 0
                GoSub UnicodeFile
            End If
        End If
'        Open mstrFileName For Input As #mlngFS
'        lngFileLen = LOF(mlngFS)
'        For lngChunk = 1 To lngFileLen Step 10240   'reads 10 KB at a time
'            If lngChunk + 10240 < lngFileLen Then
'                mstrText = mstrText & Input(10240, #mlngFS)
'            Else
'                mstrText = mstrText & Input(lngFileLen - lngChunk - 1, #mlngFS)
'            End If
'        Next lngChunk
    End If
    If Err Then OpenFile = Err
    Exit Function
UnicodeFile:
    If lngFileLen > 65536 Then
        mlngBlocks = FileLength \ 65536
        mlngBlockLen = 65536
        If FileLength Mod 65536 > 0 Then
            mlngBlocks = mlngBlocks + 1
        End If
        mblnNotUnicode = False
        Block = 1 'text loaded from changing block
        CanSave = False
    Else
        'Unicode file
        mstrText = bytRead
        mblnNotUnicode = False
    End If
    Return
End Function

'On save, the opened file is closed.
Public Function SaveFile() As Long
    Dim bytWrite() As Byte
    
    If Blocks = 0 Then
        'close the existing file
        If mlngFS Then
            Close #mlngFS
        End If
        
        If mblnNotUnicode Then
            bytWrite = StrConv(Text, vbFromUnicode)
        Else
            bytWrite = Text
        End If
        
        Open mstrFileName For Output As #mlngFS 'purge existing file
        Close #mlngFS
        'write binary file
        Open mstrFileName For Binary As #mlngFS
        'dump the binary array
        Put #mlngFS, , bytWrite
        'close the file & reopen it for Binary
        Close #mlngFS
        
        Open mstrTemporaryFile For Binary As #mlngFS
    Else
        MsgBox "Cannot save this file because it is loaded in blocks.", vbExclamation, "Save File"
    End If
End Function

Public Property Get TemporaryFile() As String
    TemporaryFile = mstrTemporaryFile
End Property

'Text stream is implemented for all file access
'to the file, except for OpenFile() and SaveFile().
Public Property Get Text() As String
    Text = mstrText
End Property

Public Property Let Text(ByVal RHS As String)
    mstrText = RHS
End Property
