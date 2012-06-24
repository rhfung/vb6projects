VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FileDlg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'FileDlg
'13 February 2002
'Update         changes window
'ItemChanged    changes variables/objects
'By Richard Fung
'From:      Flow Chart V 1.7
'Added:     Replace
'Changed:   Update
'Type:      ASCII
'Last modified: March 4, 2002 Win95

Public hwndOwner    As Long 'must define
Public DefaultExt   As String 'must define
Public FileName     As String
Public Filter       As String 'must define
Public FilterIndex  As Long
Public FileTitle    As String
Public InitialDir   As String
Public Title        As String
Public Flags        As Long

Public RetFile      As String
Public RetExt       As String
Public RetPath      As String

Private Type OPENFILENAME
        lStructSize As Long
        hwndOwner As Long
        hInstance As Long
        lpstrFilter As String
        lpstrCustomFilter As String
        nMaxCustFilter As Long
        nFilterIndex As Long
        lpstrFile As String
        nMaxFile As Long
        lpstrFileTitle As String
        nMaxFileTitle As Long
        lpstrInitialDir As String
        lpstrTitle As String
        Flags As Long
        nFileOffset As Integer
        nFileExtension As Integer
        lpstrDefExt As String
        lCustData As Long
        lpfnHook As Long
        lpTemplateName As String
End Type
Private tFile   As OPENFILENAME

Public Enum CDErr
    CDERR_OK = -&H400
    CDERR_CANCELLED = 0
    CDERR_DIALOGFAILURE = &HFFFF
    CDERR_FINDRESFAILURE = &H6
'    CDERR_GENERALCODES = &H0
    CDERR_INITIALIZATION = &H2
    CDERR_LOADRESFAILURE = &H7
    CDERR_LOADSTRFAILURE = &H5
    CDERR_LOCKRESFAILURE = &H8
    CDERR_MEMALLOCFAILURE = &H9
    CDERR_MEMLOCKFAILURE = &HA
    CDERR_NOHINSTANCE = &H4
    CDERR_NOHOOK = &HB
    CDERR_NOTEMPLATE = &H3
    CDERR_REGISTERMSGFAIL = &HC
    CDERR_STRUCTSIZE = &H1
    FNERR_BUFFERTOOSMALL = &H3003
    FNERR_FILENAMECODES = &H3000
    FNERR_INVALIDFILENAME = &H3002
    FNERR_SUBCLASSFAILURE = &H3001
End Enum

Private Declare Function CommDlgExtendedError Lib "comdlg32.dll" () As Long
Private Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long
Private Declare Function GetSaveFileName Lib "comdlg32.dll" Alias "GetSaveFileNameA" (pOpenfilename As OPENFILENAME) As Long

'Private Const OFN_ALLOWMULTISELECT = &H200
'Private Const OFN_CREATEPROMPT = &H2000
'Private Const OFN_ENABLEHOOK = &H20
'Private Const OFN_ENABLETEMPLATE = &H40
'Private Const OFN_ENABLETEMPLATEHANDLE = &H80
'Private Const OFN_EXPLORER = &H80000
'Private Const OFN_EXTENSIONDIFFERENT = &H400
Private Const OFN_FILEMUSTEXIST = &H1000
'Private Const OFN_HIDEREADONLY = &H4
'Private Const OFN_LONGNAMES = &H200000
'Private Const OFN_NOCHANGEDIR = &H8
'Private Const OFN_NODEREFERENCELINKS = &H100000
'Private Const OFN_NOLONGNAMES = &H40000
'Private Const OFN_NONETWORKBUTTON = &H20000
'Private Const OFN_NOREADONLYRETURN = &H8000
'Private Const OFN_NOTESTFILECREATE = &H10000
'Private Const OFN_NOVALIDATE = &H100
Private Const OFN_OVERWRITEPROMPT = &H2
Private Const OFN_PATHMUSTEXIST = &H800
'Private Const OFN_READONLY = &H1
'Private Const OFN_SHAREAWARE = &H4000
'Private Const OFN_SHAREFALLTHROUGH = 2
'Private Const OFN_SHOWHELP = &H10

Private Function GetString(Str As String) As String
    Dim n As Long 'null char
    
    n = InStr(1, Str, vbNullChar)
    If n > 0 Then
        GetString = Left$(Str, n - 1)
    Else
        GetString = RTrim(Str)
    End If
End Function

Public Sub Initialize(ByVal hwndOwner As Long, ByVal DefaultExt As String, ByVal Filter As String, ByVal FileName As String, ByVal OpenFile As Boolean, ByVal SaveFile As Boolean, ByVal Flags As Long)
    With Me
        .hwndOwner = hwndOwner
        .DefaultExt = DefaultExt
        .Filter = Filter
        .FileName = FileName
        .Flags = Flags
        If OpenFile Then .Flags = .Flags Or OFN_FILEMUSTEXIST
        If SaveFile Then .Flags = .Flags Or OFN_PATHMUSTEXIST Or OFN_OVERWRITEPROMPT
    End With
End Sub

Private Sub ItemChanged()
    Flags = tFile.Flags
    FileName = GetString(tFile.lpstrFile)
    FileTitle = GetString(tFile.lpstrFileTitle)
    FilterIndex = tFile.nFilterIndex
    If tFile.nFileOffset > 0 Then
        RetFile = Mid$(FileName, tFile.nFileOffset + 1)
        RetPath = Left$(FileName, tFile.nFileOffset)
    Else
        RetFile = ""
        RetPath = ""
    End If
    If tFile.nFileExtension > 0 Then
        RetExt = Mid$(FileName, tFile.nFileExtension + 1)
    Else
        RetExt = ""
    End If
End Sub

'Added for VB5
#If VB6 = 0 Then
Private Function Replace(ByVal Str As String, ByVal FromChr As String, ByVal ToChr As String) As String
    Dim lngI As Long
    
    FromChr = Left$(FromChr, 1)
    ToChr = Left$(ToChr, 1)
    For lngI = 1 To Len(Str)
        If Mid$(Str, lngI, 1) = FromChr Then
            Mid$(Str, lngI, 1) = ToChr 'replace
        End If
    Next lngI
    Replace = Str
End Function
#End If

Public Function ShowOpen() As Long
    Dim r As Long

    Update
    r = GetOpenFileName(tFile)
    If r = 0 Then
        ShowOpen = CommDlgExtendedError()
    Else
        ShowOpen = CDERR_OK
    End If
    ItemChanged
End Function

Public Function ShowSave() As Long
    Dim r As Long
    
    Update
    r = GetSaveFileName(tFile)
    If r = 0 Then
        ShowSave = CommDlgExtendedError()
    Else
        ShowSave = CDERR_OK
    End If
    ItemChanged
End Function


Private Sub Update()
    Dim strFile As String * 1024
    
    On Error Resume Next
    Call GetAttr(FileName) 'see if file exists before passed to common dialog box
    If Err Then
        strFile = vbNullChar
        Err.Clear
    Else
        strFile = FileName & vbNullChar
    End If
    
    tFile.Flags = Flags
    tFile.hwndOwner = hwndOwner
    tFile.lpstrDefExt = DefaultExt & vbNullChar
    
    tFile.lpstrFile = strFile 'already has null char
    tFile.nMaxFile = 1024
    
    tFile.lpstrFileTitle = Space(512) & vbNullChar
    tFile.nMaxFileTitle = 512
    
    tFile.lpstrFilter = Replace(Filter, "|", vbNullChar) & vbNullChar & vbNullChar
    tFile.nFilterIndex = FilterIndex
    
    tFile.lpstrTitle = Title & vbNullChar
    tFile.lStructSize = Len(tFile)
End Sub

