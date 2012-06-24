VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FlowOldFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const conPx = 4.8

Public mblnConverted As Boolean

Public Sub ConvToTwips(ByVal File As FlowChart)  'LOAD
    Dim objItem As FlowItem
    
    If File.Version < 5 And Not mblnConverted Then
        For Each objItem In File
           With objItem.P
                .Left = .Left * conPx
                .Top = .Top * conPx
                .Width = .Width * conPx
                .Height = .Height * conPx
            End With
        Next objItem
        mblnConverted = True
    End If
End Sub

#If Prev = 0 Then
Public Sub ConvFromTwips(ByVal File As FlowChart)  'SAVE
    Dim objItem As FlowItem
    
    'The commented line below is suppose
    'to have this code run once, but it could
    'run more than once (for a reason
    'I do not know of).  Since this problem
    'probably won't happen, I leave the flag
    'check commented.
    'Without commenting mblnConverted, some
    'saving to version 4 will not work if it
    'was not loaded as version 4.
    

    
    'runs only once for old version of files
    If File.Version < 5 And mblnConverted Then
        For Each objItem In File
           With objItem.P
                .Left = .Left / conPx
                .Top = .Top / conPx
                .Width = .Width / conPx
                .Height = .Height / conPx
            End With
        Next objItem
        mblnConverted = False
    End If
    
End Sub
#End If

Public Function GetText(ByVal Obj As FlowItem) As String
    'split at ¿ chr(191)
    If InStr(1, Obj.P.Text, "¿", vbBinaryCompare) < 3 Then 'at position 1 or 2
        GetText = Mid$(Obj.P.Text, InStr(1, Obj.P.Text, "¿", vbBinaryCompare) + 1)
    Else
        GetText = Obj.P.Text
    End If
End Function

Public Function GetTextFlags(ByVal Obj As FlowItem) As Long
    If InStr(1, Obj.P.Text, "¿", vbBinaryCompare) Then
        GetTextFlags = Left$(Obj.P.Text, InStr(1, Obj.P.Text, "¿", vbBinaryCompare) - 1)
    Else
        GetTextFlags = 0 'default
    End If
End Function

Public Function GetTextFormatBold(ByVal Obj As FlowItem) As Boolean
    GetTextFormatBold = InStr(1, Obj.P.Tag1, "b", vbBinaryCompare)
End Function

Public Function GetTextFormatItalic(ByVal Obj As FlowItem) As Boolean
    GetTextFormatItalic = InStr(1, Obj.P.Tag1, "i", vbBinaryCompare)
End Function

Public Function GetTextFormatSizeFactor(ByVal Obj As FlowItem) As Integer
'asc("+")=43
'asc("-")=45
    GetTextFormatSizeFactor = StrCount(Obj.P.Tag1, 43) - StrCount(Obj.P.Tag1, 45)
End Function


#If Prev = 0 Then
Public Sub SetTextFormat(Obj As FlowItem, Optional Bold, Optional Italic, Optional SizeFactor)
    Dim blnBold         As Boolean
    Dim blnItalic       As Boolean
    Dim intSizeFactor   As Integer
    
    'set flags
    If Not IsMissing(Bold) Then blnBold = Bold Else blnBold = GetTextFormatBold(Obj)
    If Not IsMissing(Italic) Then blnItalic = Italic Else blnItalic = GetTextFormatItalic(Obj)
    If Not IsMissing(SizeFactor) Then intSizeFactor = SizeFactor Else intSizeFactor = GetTextFormatSizeFactor(Obj)
    'save flags
    Obj.P.Tag1 = IIf(blnBold, "b", vbNullString) & IIf(blnItalic, "i", vbNullString)
    If intSizeFactor > 0 Then
        Obj.P.Tag1 = Obj.P.Tag1 & String$(intSizeFactor, "+")
    ElseIf intSizeFactor < 0 Then
        Obj.P.Tag1 = Obj.P.Tag1 & String$(-intSizeFactor, "-")
    End If
End Sub
#End If

Private Function StrCount(ByRef Text As String, ByVal Find As Integer) As Long
    Dim lngCh As Long
    
    
    For lngCh = 1 To Len(Text)
        If Asc(Mid$(Text, lngCh, 1)) = Find Then
            StrCount = StrCount + 1
        End If
    Next lngCh
End Function
