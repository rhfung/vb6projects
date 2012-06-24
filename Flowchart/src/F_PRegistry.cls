VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PRegistry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
'Richard Fung.  August 16, 2000.

'Interface
Public GridSnap As Boolean
Public GridShow As Boolean
Public Margin   As Single 'in Inches
Public AutoEditText As Boolean
'Public StatusBar    As Boolean
Public FontBar      As Boolean

Const conApp = "Flow Chart"
Const conSect = "Interface"

Public Function ReadRegistry() As Long
    On Error GoTo Handler
    
    GridSnap = GetSetting(conApp, conSect, "GridSnap", False)
    GridShow = GetSetting(conApp, conSect, "GridShow", False)
    Margin = CSng(GetSetting(conApp, conSect, "Margin", 1))
    AutoEditText = GetSetting(conApp, conSect, "AutoEditText", True)
    FontBar = GetSetting(conApp, conSect, "FontBar", False)
    'StatusBar = GetSetting(conApp, conSect, "StatusBar", False)
    
    Exit Function
Handler:
    ReadRegistry = Err.Number
End Function

Public Function SaveRegistry() As Long
#If Prev = 0 Then
    On Error GoTo Handler
    
    SaveSetting conApp, conSect, "GridSnap", GridSnap
    SaveSetting conApp, conSect, "GridShow", GridShow
    SaveSetting conApp, conSect, "Margin", Margin
    SaveSetting conApp, conSect, "AutoEditText", AutoEditText
    SaveSetting conApp, conSect, "FontBar", FontBar
    'SaveSetting conApp, conSect, "StatusBar", StatusBar
    
    Exit Function
Handler:
    SaveRegistry = Err.Number
#End If
End Function
