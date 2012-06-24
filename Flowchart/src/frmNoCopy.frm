VERSION 5.00
Begin VB.Form frmNocopy 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Design 3"
   ClientHeight    =   1695
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3465
   ControlBox      =   0   'False
   Icon            =   "frmNoCopy.frx":0000
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1695
   ScaleWidth      =   3465
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdButton 
      Caption         =   "OK"
      Height          =   375
      Left            =   1080
      TabIndex        =   1
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Label lblInfo 
      Height          =   855
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   3015
   End
End
Attribute VB_Name = "frmNocopy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Feb 8, 2001
Private Declare Function GetVolumeInformation Lib "kernel32" Alias "GetVolumeInformationA" (ByVal lpRootPathName As String, ByVal lpVolumeNameBuffer As String, ByVal nVolumeNameSize As Long, lpVolumeSerialNumber As Long, lpMaximumComponentLength As Long, lpFileSystemFlags As Long, ByVal lpFileSystemNameBuffer As String, ByVal nFileSystemNameSize As Long) As Long

Private mByte As Byte, int1 As Integer, int2 As Integer, int3 As Integer, int4 As Integer, int5 As Integer, int6 As Integer, int7 As Integer, int8 As Integer
Private intS1 As Integer, intS2 As Integer, intS3 As Integer

Private Sub AgathaChristie()
    'Agatha Chrisite function
End Sub

Sub Cut(ByRef Str As String)
    On Error Resume Next
    Str = Left(Str, InStr(1, Str, vbNullChar) - 1)
    If Err Then Str = ""
End Sub

Private Sub Determine() 'determines how the window will respond
    Const kingcaesarID = &H2D651BDE 'used for protection
    Const requestedID = &H17E4143D 'real ID
    Const caesarsID = &H18D81F49
    Dim strRootPathName$
    Dim strVolumeName$ 'user assigned name
    Dim lngSerialNo& 'in hexadecimal
    Dim lngMaxComponentLength&
    Dim lngSystemFlags&
    Dim strSysName$ 'FAT, NTFS
    Dim r&
    
    strRootPathName = Left(App.Path, 1) & ":\"
    strVolumeName = String(255, " ")
    strSysName = String(5, " ")
    
    r = GetVolumeInformation(strRootPathName, strVolumeName, LenL(strVolumeName), lngSerialNo, lngMaxComponentLength, lngSystemFlags, strSysName, LenL(strSysName))
    
    Cut strVolumeName
    Cut strSysName
    
    lblInfo = "The program on the disk"
    #If CopyProtect Then '--------------------------
        Randomize Timer
        int1 = Int(Rnd * 10) + 1
        int2 = Int(Rnd * 10) + 1
        Randomize Timer
        int3 = Int(Rnd * 10) + 1
        int4 = Int(Rnd * 10) + 1
        Randomize Timer
        int5 = Int(Rnd * 10) + 1
        int6 = Int(Rnd * 10) + 1
        Randomize Timer
        int7 = Int(Rnd * 10) + 1
        int8 = Int(Rnd * 10) + 1
        
        If lngSerialNo = requestedID And strVolumeName = "DESIGN3" And r = 1 Then
            AgathaChristie
            lblInfo = lblInfo & " is OK."
            mByte = 83
            intS1 = int1 + int4 + int6 + int2
            intS3 = int8 + int2 + int1 - int5
            cmdButton_Click
        Else
            If lngSerialNo = kingcaesarID And strVolumeName = "DESIGN2" And r = 9 Then
                AgathaChristie
            End If
            lblInfo = lblInfo & " is not OK -- it has been copied.  Program developed by Richard Fung."
            mByte = 17
            intS2 = int8 + int2 + int1 - int5
            intS3 = int2 + int3 + int5 + int7
            If lngSerialNo = caesarsID And strVolumeName = "FULLCOPY" And r = 4 Then
                AgathaChristie
            End If
        End If
    #Else '-------------------------
        lblInfo = lblInfo & " is not checked."
        mByte = 83
        intS1 = 0
        intS3 = 0
        cmdButton_Click  'pass by window
    #End If
End Sub


Private Function LenL(Text As String) As Long
    LenL = CLng(Len(Text))
End Function


Private Sub cmdButton_Click()
    If mByte = 83 And mByte = 62 + 21 And _
     int1 + int4 + int6 + int2 = intS1 And _
     intS3 = int8 + int2 + int1 - int5 Then
        'Load Window
        CPlusPlus
        Unload Me
    ElseIf mByte = 83 Then
        MsgBox "Program has been illegally modified!", vbCritical
        End
    Else
        Unload Me
    End If
End Sub

Private Sub Form_Load()
    Determine
End Sub


