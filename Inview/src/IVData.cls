VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IVRegistry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
'Data class
'Richard Fung, Programming Applications 1

Public HexCols      As Long
Public ChCols       As Long 'character columns
Public DecCols      As Long

Public TextFontName As String   'as the user wishes
Public TextFontSize As Currency
Public TextBold     As Boolean

Public BinFontName  As String
Public BinFontSize  As Currency
Public BinBold      As Boolean

Public NoSaveSetting As Boolean

Private Const conAppName = "InView"
Private Const conSText = "Text" 'section
Private Const conSBin = "Binary"
Private Const conSGeneral = "General"

Private Const conKName = "FontName" 'key
Private Const conKSize = "FontSize"
Private Const conKBold = "Bold"

Private Const conKAppPath = "AppPath"

Private Const conMinSize = 4
Private Const conMaxSize = 20
Private Const conDefaultSize = 10

Public Sub ClearData()
    On Error Resume Next
    DeleteSetting conAppName
End Sub


Private Function CSize(ByVal Str As String) As Currency
    If IsNumeric(Str) Then
        CSize = CCur(Str)
        If CSize < conMinSize Then CSize = conMinSize 'keep font size in bounds 4 to 20
        If CSize > conMaxSize Then CSize = conMaxSize
    Else
        CSize = conDefaultSize 'default
    End If
End Function

'At some time in the future, the data could
'be read from the registry instead of being
'assigned preset values.
Public Sub ReadData()
    ChCols = 80 'default number of columns
    HexCols = 30
    DecCols = 20
    
    On Error Resume Next
    
    TextFontName = GetSetting(conAppName, conSText, conKName, "Courier New")
    TextFontSize = CSize(GetSetting(conAppName, conSText, conKSize, conDefaultSize))
    TextBold = CBool(GetSetting(conAppName, conSText, conKBold, "0"))
    
    BinFontName = "Courier New"
    BinFontSize = CSize(GetSetting(conAppName, conSBin, conKSize, conDefaultSize))
    BinBold = CBool(GetSetting(conAppName, conSBin, conKBold, "0"))
End Sub

Public Sub SaveData()
    If Not NoSaveSetting Then
        'Saves data to the registry.
        SaveSetting conAppName, conSText, conKName, TextFontName
        SaveSetting conAppName, conSText, conKSize, TextFontSize
        SaveSetting conAppName, conSText, conKBold, TextBold
        
        SaveSetting conAppName, conSBin, conKSize, BinFontSize
        SaveSetting conAppName, conSBin, conKBold, BinBold
        
        SaveSetting conAppName, conSGeneral, conKAppPath, App.Path
    End If
End Sub


Private Sub Class_Initialize()
    ReadData
End Sub


Private Sub Class_Terminate()
    SaveData
End Sub


