VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "RegEntry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
'22 Jan 2002
'GetSetting returns Value, Default, or Error
'SaveSettings returns 0=good or non-zero=err#

' Reg Key Security Options...
Const READ_CONTROL = &H20000
Const KEY_QUERY_VALUE = &H1
Const KEY_SET_VALUE = &H2
Const KEY_CREATE_SUB_KEY = &H4
Const KEY_ENUMERATE_SUB_KEYS = &H8
Const KEY_NOTIFY = &H10
Const KEY_CREATE_LINK = &H20
Const KEY_ALL_ACCESS = KEY_QUERY_VALUE + KEY_SET_VALUE + _
                       KEY_CREATE_SUB_KEY + KEY_ENUMERATE_SUB_KEYS + _
                       KEY_NOTIFY + KEY_CREATE_LINK + READ_CONTROL
                     
' Reg Key ROOT Types...
Enum RegEntryEnum
    HKEY_CLASSES_ROOT = &H80000000
    HKEY_CURRENT_CONFIG = &H80000005
    HKEY_CURRENT_USER = &H80000001
    HKEY_DYN_DATA = &H80000006
    HKEY_LOCAL_MACHINE = &H80000002
    HKEY_PERFORMANCE_DATA = &H80000004
    HKEY_USERS = &H80000003
End Enum
Const ERROR_SUCCESS = 0
Const REG_SZ = 1                         ' Unicode nul terminated string
Const REG_DWORD = 4                      ' 32-bit number

'Const gREGKEYSYSINFOLOC = "SOFTWARE\Microsoft\Shared Tools Location"
'Const gREGVALSYSINFOLOC = "MSINFO"
'Const gREGKEYSYSINFO = "SOFTWARE\Microsoft\Shared Tools\MSINFO"
'Const gREGVALSYSINFO = "PATH"

Private Declare Function RegOpenKeyEx Lib "advapi32" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, ByRef phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, ByRef lpType As Long, ByVal lpData As String, ByRef lpcbData As Long) As Long
Private Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32" (ByVal hKey As Long) As Long
Private Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Private Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal hKey As Long, ByVal lpSubKey As String) As Long
Private Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal hKey As Long, ByVal lpValueName As String) As Long
Private Declare Function RegSetValueExString Lib "advapi32.dll" Alias _
   "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, _
   ByVal Reserved As Long, ByVal dwType As Long, ByVal lpValue As _
   String, ByVal cbData As Long) As Long
   

Public KeyRoot As Long
Public KeyName As String
Public KeySubRef As String
Private mlngKeyType As Long
'
'KeyRoot must be pre-defined.
'Other variables can be assigned separately,
' or accessed using the functions GetSetting
' and SaveSetting.





Public Function DeleteSetting(ByVal KeyName As String, ByVal SubKeyName As String) As Boolean
    DeleteSetting = DelKeyValue(KeyRoot, KeyName, SubKeyName)
End Function


Private Function DelKeyValue(KeyRoot As Long, KeyName As String, SubKeyName As String) As Boolean
    Dim i As Long                                           ' Loop Counter
    Dim rc As Long                                          ' Return Code
    Dim hKey As Long                                        ' Handle To An Open Registry Key
    Dim hDepth As Long                                      '
    Dim KeyValType As Long                                  ' Data Type Of A Registry Key
    Dim tmpVal As String                                    ' Tempory Storage For A Registry Key Value
    Dim KeyValSize As Long                                  ' Size Of Registry Key Variable
    '------------------------------------------------------------
    ' Open RegKey Under KeyRoot {HKEY_LOCAL_MACHINE...}
    '------------------------------------------------------------
    rc = RegOpenKeyEx(KeyRoot, KeyName, 0, KEY_ALL_ACCESS, hKey) ' Open Registry Key
    
    If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError          ' Handle Error...
    
    rc = RegDeleteValue(hKey, SubKeyName)
                        
    If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError          ' Handle Errors
    
    rc = RegCloseKey(hKey)                                  ' Close Registry Key
    DelKeyValue = True
    Exit Function                                           ' Exit
    
GetKeyError:      ' Cleanup After An Error Has Occured...
    DelKeyValue = False                                     ' Return Failure
    rc = RegCloseKey(hKey)
End Function

Private Function GetKeyValue(KeyRoot As Long, KeyName As String, SubKeyRef As String, ByRef KeyVal As Variant) As Boolean
    Dim i As Long                                           ' Loop Counter
    Dim rc As Long                                          ' Return Code
    Dim hKey As Long                                        ' Handle To An Open Registry Key
    Dim hDepth As Long                                      '
    Dim KeyValType As Long                                  ' Data Type Of A Registry Key
    Dim tmpVal As String                                    ' Tempory Storage For A Registry Key Value
    Dim KeyValSize As Long                                  ' Size Of Registry Key Variable
    '------------------------------------------------------------
    ' Open RegKey Under KeyRoot {HKEY_LOCAL_MACHINE...}
    '------------------------------------------------------------
    rc = RegOpenKeyEx(KeyRoot, KeyName, 0, KEY_ALL_ACCESS, hKey) ' Open Registry Key
    
    If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError          ' Handle Error...
    
    tmpVal = String$(1024, 0)                             ' Allocate Variable Space
    KeyValSize = 1024                                       ' Mark Variable Size
    
    '------------------------------------------------------------
    ' Retrieve Registry Key Value...
    '------------------------------------------------------------
    rc = RegQueryValueEx(hKey, SubKeyRef, 0, _
                         KeyValType, tmpVal, KeyValSize)    ' Get/Create Key Value
                        
    If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError          ' Handle Errors
    
    If (Asc(Mid$(tmpVal, KeyValSize, 1)) = 0) Then           ' Win95 Adds Null Terminated String...
        tmpVal = Left$(tmpVal, KeyValSize - 1)               ' Null Found, Extract From String
    Else                                                    ' WinNT Does NOT Null Terminate String...
        tmpVal = Left$(tmpVal, KeyValSize)                   ' Null Not Found, Extract String Only
    End If
    '------------------------------------------------------------
    ' Determine Key Value Type For Conversion...
    '------------------------------------------------------------
    Select Case KeyValType                                  ' Search Data Types...
    Case REG_SZ                                             ' String Registry Key Data Type
        KeyVal = tmpVal                                     ' Copy String Value
    Case REG_DWORD                                          ' Double Word Registry Key Data Type
        For i = Len(tmpVal) To 1 Step -1                    ' Convert Each Bit
            KeyVal = KeyVal + Hex(Asc(Mid$(tmpVal, i, 1)))   ' Build Value Char. By Char.
        Next
        KeyVal = CLng("&h" + KeyVal)                     ' Convert Double Word To long
    End Select
    
    mlngKeyType = KeyValType 'copy format of loaded key
    
    GetKeyValue = True                                      ' Return Success
    rc = RegCloseKey(hKey)                                  ' Close Registry Key
    Exit Function                                           ' Exit
    
GetKeyError:      ' Cleanup After An Error Has Occured...
    GetKeyValue = False                                     ' Return Failure
    rc = RegCloseKey(hKey)                                  ' Close Registry Key
End Function

Public Function GetSetting(ByVal KeyName As String, ByVal KeySubRef As String, Optional ByVal Default) As Variant
    If IsMissing(Default) Then
        Me.KeyName = KeyName
        Me.KeySubRef = KeySubRef
        GetSetting = Me.KeyVal
    Else
        On Error Resume Next
        Me.KeyName = KeyName
        Me.KeySubRef = KeySubRef
        GetSetting = Me.KeyVal
        If Err Then
            Err.Clear
            GetSetting = Default
        End If
    End If
End Function

Public Property Get KeyType() As Long
    KeyType = mlngKeyType
End Property

Public Function SaveSetting(ByVal KeyName As String, ByVal KeySubRef As String, ByVal Value As Variant) As Long
    Me.KeyName = KeyName
    Me.KeySubRef = KeySubRef
    
    On Error Resume Next
    Me.KeyVal = Value
    If Err Then
        SaveSetting = Err 'return error #
        Err.Clear
    End If
End Function

Private Function SetKeyValue(KeyRoot As Long, KeyName As String, SubKeyRef As String, ByRef KeyVal As Variant) As Boolean
    Dim i As Long                                           ' Loop Counter
    Dim rc As Long                                          ' Return Code
    Dim hKey As Long                                        ' Handle To An Open Registry Key
    Dim hDepth As Long                                      '
    Dim tmpVal As String                                    ' Tempory Storage For A Registry Key Value
    Dim tmpVal2 As Long
    Dim KeyValSize As Long                                  ' Size Of Registry Key Variable
    '------------------------------------------------------------
    ' Open RegKey Under KeyRoot {HKEY_LOCAL_MACHINE...}
    '------------------------------------------------------------
    rc = RegOpenKeyEx(KeyRoot, KeyName, 0, KEY_ALL_ACCESS, hKey)    ' Open Registry Key
    
    If (rc <> ERROR_SUCCESS) Then
        'GoTo GetKeyError          ' Handle Error...
        'doesn't exist, create key
        rc = RegCreateKey(KeyRoot, KeyName, hKey)
        If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError
    End If
    
    '------------------------------------------------------------
    ' Determine Key Value Type For Conversion...
    '------------------------------------------------------------
    Select Case KeyType                                  ' Search Data Types...
    Case REG_SZ                                             ' String Registry Key Data Type
        tmpVal = KeyVal & Chr(0)                                ' Allocate Variable Space
        KeyValSize = Len(tmpVal) - 1                          ' Mark Variable Size
        rc = RegSetValueExString(hKey, SubKeyRef, 0&, _
                             KeyType, tmpVal, KeyValSize)  ' Set/Create Key Value
    Case REG_DWORD                                          ' Double Word Registry Key Data Type
        tmpVal2 = CLng(KeyVal)
        KeyValSize = Len(tmpVal2)
        rc = RegSetValueEx(hKey, SubKeyRef, 0&, _
                             KeyType, tmpVal2, KeyValSize)     ' Set/Create Key Value
'        tmpVal = Hex(KeyVal)
'        tmpVal2 = ""
'        ReDim Preserve tmpVal2(Len(tmpVal) \ 2)
'        For i = Len(tmpVal) To 1 Step -2
'            tmpVal2 = tmpVal2 + Chr(Val("&h" & Mid$(tmpVal, i - 1, 2)))
'        Next i
'        tmpVal = tmpVal2
'        KeyValSize = Len(tmpVal)
'        For i = Len(KeyVal) To 1 Step -1                    ' Convert Each Bit
'            KeyVal = KeyVal + Hex(Asc(Mid$(tmpVal, i, 1)))   ' Build Value Char. By Char.
'        Next
'        tmpVal = Format$("&h" + KeyVal)                    ' Convert Double Word To String
    Case Else
        GoTo GetKeyError
    End Select
    
    '------------------------------------------------------------
    ' Save Registry Key Value...
    '--------------- ---------------------------------------------
'    rc = RegSetValueEx(hKey, SubKeyRef, 0, _
'                         KeyType, tmpVal, KeyValSize)    ' Set/Create Key Value
                        
    If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError          ' Handle Errors
    
'    If (Asc(Mid$(tmpVal, KeyValSize, 1)) = 0) Then           ' Win95 Adds Null Terminated String...
'        tmpVal = Left$(tmpVal, KeyValSize - 1)               ' Null Found, Extract From String
'    Else                                                    ' WinNT Does NOT Null Terminate String...
'        tmpVal = Left$(tmpVal, KeyValSize)                   ' Null Not Found, Extract String Only
'    End If
    
    SetKeyValue = True                                      ' Return Success
    rc = RegCloseKey(hKey)                                  ' Close Registry Key
    Exit Function                                           ' Exit
    
GetKeyError:      ' Cleanup After An Error Has Occured...
    SetKeyValue = False                                     ' Return Failure
    rc = RegCloseKey(hKey)                                  ' Close Registry Key
End Function


Public Property Let KeyVal(ByVal pKeyVal As Variant)
    Select Case VarType(pKeyVal)
        Case 2 To 3, 11, 17
            mlngKeyType = REG_DWORD
            If Not SetKeyValue(KeyRoot, KeyName, KeySubRef, pKeyVal) Then
                Err.Raise 5
            End If
        Case 4 To 8, 14
            mlngKeyType = REG_SZ
            If Not SetKeyValue(KeyRoot, KeyName, KeySubRef, pKeyVal) Then
                Err.Raise 5
            End If
        Case Else
            Err.Raise 13 'type mismatch
    End Select
End Property

Public Property Get KeyVal() As Variant
Attribute KeyVal.VB_UserMemId = 0
    ' Try To Get System Info Program Path\Name From Registry...
    If Not GetKeyValue(KeyRoot, KeyName, KeySubRef, KeyVal) Then
        'error
        Err.Raise 5
    End If
End Property


