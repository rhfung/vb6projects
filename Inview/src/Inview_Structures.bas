Attribute VB_Name = "modS"
Option Explicit
'Richard Fung, Programming Applications 1


''Structures for file formats
''This file format is for reference files.
'Public Type RefFileStruct
'    Keyword As String * 15
'    Syntax  As String * 50
'    Description As String * 100
'End Type

'Constants for frmMain::CloseFile()
Public Const Yes = True
Public Const No = False
Public Const conTitle = "In View"
Public gData    As IVRegistry
Public gstrPath As String 'for Temp files

Private Declare Function GetDriveType Lib "kernel32" Alias "GetDriveTypeA" (ByVal nDrive As String) As Long
Public Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Private Const DRIVE_CDROM = 5
Private Const DRIVE_FIXED = 3
Private Const DRIVE_RAMDISK = 6
Private Const DRIVE_REMOTE = 4
Private Const DRIVE_REMOVABLE = 2

Public Function GetPath(ByVal Pathname As String) As String
    Dim lngPrevious As Long
    Dim lngNext     As Long
    Dim strExtra    As String
    Dim strPath     As String
    
    'get path name
    Do
        lngPrevious = lngNext
        lngNext = InStr(lngPrevious + 1, Pathname, "\")
    Loop Until lngNext = 0
    If lngPrevious > 0 Then strPath = Left$(Pathname, lngPrevious)
    
    GetPath = strPath
End Function

Public Function GetTempFile(ByVal Filename As String) As String
    'for determining drive type
    Static lngDrive As Long
    Dim strDrive    As String
    Dim r           As Long
    'for getting path and filename
    Dim lngPrevious As Long
    Dim lngNext     As Long
    Dim strExtra    As String
    Dim strFile     As String
    
        
    'first determine if this has been called before
    If lngDrive = 0 And Len(gstrPath) = 0 Then
        'get drive type
        strDrive = Left$(App.Path, 3) & vbNullChar
        lngDrive = GetDriveType(strDrive)
        'choose where to place temp files
        Select Case lngDrive
        Case DRIVE_FIXED, DRIVE_RAMDISK, DRIVE_REMOTE   'writable
            gstrPath = App.Path 'put temp files in application directory
        Case Else
            gstrPath = Space(1024) 'put temp files in Temp directory if possible
            r = GetTempPath(1024, gstrPath)
            If r > 1024 Or r = 0 Then 'error - place temp file in App.Path
                gstrPath = App.Path
            Else 'good
                gstrPath = Left$(gstrPath, InStr(1, gstrPath, vbNullChar) - 1)
            End If
        End Select
    End If
    
    'get path name
    Do
        lngPrevious = lngNext
        lngNext = InStr(lngPrevious + 1, Filename, "\")
    Loop Until lngNext = 0
    If Right$(gstrPath, 1) <> "\" Then strExtra = "\"
    strFile = Mid$(Filename, lngPrevious + 1)
    
    lngPrevious = 0
    lngNext = 0
    Do
        lngPrevious = lngNext
        lngNext = InStr(lngPrevious + 1, strFile, ".")
        If lngNext > 0 Then Mid$(strFile, lngNext, 1) = "_" 'replace (.) with (_)
    Loop Until lngNext = 0
    If lngPrevious > 0 Then
        strFile = Left$(strFile, lngPrevious - 1)
    End If
    
    GetTempFile = gstrPath & strExtra & strFile & "_" & RndChr(5) & ".tmp"
End Function

'creates a random number of characters
Private Function RndChr(ByVal Characters As Long) As String
    Dim lngPos As Long
    
    For lngPos = 1 To Characters
        Randomize Timer
        RndChr = RndChr & Chr(Int(Rnd * 26) + 97)
    Next lngPos
End Function


