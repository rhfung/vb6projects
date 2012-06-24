Attribute VB_Name = "modFile"
Option Explicit
'Richard Fung.
'March 31, 2000.
'File mod

Private Type FThought
    Idea As String
    Text As String
    Attachment As String
    AttachmentTag As String
    CenterX As Single
    CenterY As Single
    LinkList As String
    Picture As String
    Tag As String
End Type

Private Type FHeader
    ColourScheme As Integer
    Author As String
    DateModified As Date
    LastSelected As Integer
    EditTextWin As Boolean
    Comment As String
    Tag As String
End Type

Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Type POINTAPI
        x As Long
        y As Long
End Type
'Declare Function WinExec Lib "kernel32" (ByVal lpCmdLine As String, ByVal nCmdShow As Long) As Long

Public Const conWidth = 1000
Public Const conHeight = 1000
Public Const Good = True
Public Const Fail = False
Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
'Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
'
'Public Const WS_THICKFRAME = &H40000
'Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
'
'Public Const SWP_NOMOVE = &H2
'
'Public Const SWP_NOSIZE = &H1
'
'Public Const SWP_NOACTIVATE = &H10
'
'Public Const SWP_NOZORDER = &H4
'Public Const HWND_BOTTOM = 1
'Public Const SWP_FRAMECHANGED = &H20
'Public Const SWP_DRAWFRAME = SWP_FRAMECHANGED
'Public Const WS_CAPTION = &HC00000
'
'Public Const WS_MAXIMIZEBOX = &H10000
'
'Public Const WS_MINIMIZEBOX = &H20000
'
'Public Const WS_DLGFRAME = &H400000
'Public Const WS_SYSMENU = &H80000
'Public Const WS_OVERLAPPED = &H0&
'Public Const WS_OVERLAPPEDWINDOW = (WS_OVERLAPPED Or WS_CAPTION Or WS_SYSMENU Or WS_THICKFRAME Or WS_MINIMIZEBOX Or WS_MAXIMIZEBOX)
'Public Const WS_BORDER = &H800000
'
'CDL
'Type OPENFILENAME
'        lStructSize As Long
'        hwndOwner As Long
'        hInstance As Long
'        lpstrFilter As String
'        lpstrCustomFilter As String
'        nMaxCustFilter As Long
'        nFilterIndex As Long
'        lpstrFile As String
'        nMaxFile As Long
'        lpstrFileTitle As String
'        nMaxFileTitle As Long
'        lpstrInitialDir As String
'        lpstrTitle As String
'        flags As Long
'        nFileOffset As Integer
'        nFileExtension As Integer
'        lpstrDefExt As String
'        lCustData As Long
'        lpfnHook As Long
'        lpTemplateName As String
'End Type
'
'Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long
'
'Public Const OFN_READONLY = &H1
'Public Const OFN_OVERWRITEPROMPT = &H2
'Public Const OFN_HIDEREADONLY = &H4
'Public Const OFN_NOCHANGEDIR = &H8
'Public Const OFN_SHOWHELP = &H10
'Public Const OFN_ENABLEHOOK = &H20
'Public Const OFN_ENABLETEMPLATE = &H40
'Public Const OFN_ENABLETEMPLATEHANDLE = &H80
'Public Const OFN_NOVALIDATE = &H100
'Public Const OFN_ALLOWMULTISELECT = &H200
'Public Const OFN_EXTENSIONDIFFERENT = &H400
'Public Const OFN_PATHMUSTEXIST = &H800
'Public Const OFN_FILEMUSTEXIST = &H1000
'Public Const OFN_CREATEPROMPT = &H2000
'Public Const OFN_SHAREAWARE = &H4000
'Public Const OFN_NOREADONLYRETURN = &H8000
'Public Const OFN_NOTESTFILECREATE = &H10000
'Public Const OFN_NONETWORKBUTTON = &H20000
'Public Const OFN_NOLONGNAMES = &H40000                      '  force no long names for 4.x modules
'Public Const OFN_EXPLORER = &H80000                         '  new look commdlg
'Public Const OFN_NODEREFERENCELINKS = &H100000
'Public Const OFN_LONGNAMES = &H200000                       '  force long names for 3.x modules
'
'
Declare Function Ellipse Lib "gdi32" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long



'From G_Ins3.bas: GroupCharVB
Public Function Sort(ByVal Text As String, ByVal SplitChar As String, Optional ByVal QuoteSensitive As Boolean = False) As Variant
    Dim lngPosition As Long
    Dim lngSection  As Long
    Dim strReturn() As String
'    Dim blnSymbol   As Boolean
    Dim blnInQuotes As Boolean
    
'    SplitChar = Mid(SplitChar, 1, 1)
    SplitChar = Left(SplitChar, 1)
    lngSection = 1
    ReDim strReturn(1 To 1)
    For lngPosition = 1 To Len(Text)
        Select Case Mid(Text, lngPosition, 1)
            Case SplitChar
                If Not blnInQuotes Then
                    If SplitChar = " " Then
                        'if multiple spaces, ignore
                        If Len(strReturn(lngSection)) Then
                            lngSection = lngSection + 1
                            ReDim Preserve strReturn(1 To lngSection)
                        End If
                    Else
                        lngSection = lngSection + 1
                        ReDim Preserve strReturn(1 To lngSection)
                    End If
                Else
                    strReturn(lngSection) = strReturn(lngSection) & Mid(Text, lngPosition, 1)
                End If
            Case """"
                If QuoteSensitive Then
                    If Mid(Text, lngPosition, 2) = """" & SplitChar Or Mid(Text, lngPosition, 2) = """" Then
                        blnInQuotes = False
    '                    If blnInQuotes Then
    '                        blnInQuotes = False
                    ElseIf Len(strReturn(lngSection)) Then
                        strReturn(lngSection) = strReturn(lngSection) & """"
                    Else
                        blnInQuotes = True
                    End If
                Else
                    strReturn(lngSection) = strReturn(lngSection) & """"
                End If
            Case Else
                strReturn(lngSection) = strReturn(lngSection) & Mid(Text, lngPosition, 1)
        End Select
    Next lngPosition
    Sort = strReturn
End Function


Public Function OpenFile(ByVal FileName As String) As BaseThoughts
    Dim objOpen As BaseThoughts
    Dim tRead   As FThought
    Dim tHead(1 To 2) As FHeader
    Dim intCount As Integer
    'links
    Dim objUse  As Thought
    Dim vntUse  As Variant
    Dim intUse  As Integer
    
    
    Set objOpen = New BaseThoughts
    
    Open FileName For Binary As #1
    
    'get header
    Get #1, , tHead(1)
    With objOpen.Header
        .Author = tHead(1).Author
        .ColourScheme = tHead(1).ColourScheme
        .Comment = tHead(1).Comment
        .EditTextWindow = tHead(1).EditTextWin
        .ExtraData tHead(1).LastSelected, tHead(1).DateModified, tHead(1).Tag
        vntUse = Sort(tHead(1).Tag, "/") 'for future tags
        If vntUse(1) <> "" Then
            .Circled = vntUse(1)
        End If
    End With
    
    'get thoughts
    Do While Loc(1) <= LOF(1)
        tRead.Attachment = ""
        tRead.AttachmentTag = ""
        tRead.CenterX = 0
        tRead.CenterY = 0
        tRead.Idea = ""
        tRead.LinkList = ""
        tRead.Picture = ""
        tRead.Tag = ""
        tRead.Text = ""
        Get #1, , tRead
        If Loc(1) <= LOF(1) Then
            With objOpen.Add(tRead.Idea, tRead.Text, tRead.Attachment, tRead.AttachmentTag, tRead.CenterX, tRead.CenterY)
                'not included with Add
                .Picture = tRead.Picture
                .Tag = tRead.Tag
            End With
        End If
    Loop
    
    'get links
    Seek #1, 1
    Get #1, , tHead(2)
    Do While Loc(1) <= LOF(1)
        tRead.Attachment = ""
        tRead.AttachmentTag = ""
        tRead.CenterX = 0
        tRead.CenterY = 0
        tRead.Idea = ""
        tRead.LinkList = ""
        tRead.Picture = ""
        tRead.Tag = ""
        tRead.Text = ""
        intCount = intCount + 1
        Get #1, , tRead
        If Loc(1) <= LOF(1) Then
            vntUse = Sort(tRead.LinkList, ",")
            If vntUse(1) <> "" Then
                For intUse = LBound(vntUse) To UBound(vntUse)
                    Set objUse = objOpen.FindNum(vntUse(intUse))
                    If objUse Is Nothing Then
                        MsgBox "Link reference in " & intCount & " to " & vntUse(intUse) & " is invalid.", vbExclamation, "Open File"
                    Else
                        objOpen(intCount).Links.Add objUse
                    End If
                Next intUse
            End If
        End If
    Loop
    Close #1
    Set OpenFile = objOpen
    Set objOpen = Nothing
End Function

Public Sub SaveFile(ByVal FileName As String, ByVal objSave As BaseThoughts, Optional ByVal EditWindow As Boolean = False, Optional ByVal Default As Thought)
    Dim tRead   As FThought
    Dim tHead   As FHeader
    Dim intCount As Integer
    
    'delete a temp file
    'Why a temp file?  Because if cannot
    'save properly, prog shouldn't delete
    'the good file.  That's why put to temp,
    'errors affect temp file.  Orginal file
    'still good and can be reverted if needed.
    Open FileName & ".tmp" For Output As #3  'erase file
    Close #3
    
    'save to temp file
    Open FileName & ".tmp" For Binary As #2
    
    'renumber all items FIRST
    For intCount = 1 To objSave.Count
        objSave(intCount).Index = intCount
    Next intCount
    
    'header
    With objSave.Header
        tHead.Author = .Author
        tHead.ColourScheme = .ColourScheme
        tHead.Comment = .Comment
        tHead.DateModified = Now
        tHead.EditTextWin = EditWindow
        If Not Default Is Nothing Then
            tHead.LastSelected = Default.Index
        End If
        tHead.Tag = .Circled
    End With
    Put #2, , tHead
    
    'save thoughts
    For intCount = 1 To objSave.Count
        With objSave(intCount)
            tRead.Attachment = .Attachment
            tRead.AttachmentTag = .AttachmentTag
            tRead.CenterX = .CenterX
            tRead.CenterY = .CenterY
            tRead.Idea = .Idea
            tRead.LinkList = .LinkList
            tRead.Picture = .Picture
            tRead.Tag = .Tag
            tRead.Text = .Text
            Put #2, , tRead
        End With
    Next intCount
    Close #2
    
    'delete org file (if any)
    On Error Resume Next
    Kill FileName
    On Error GoTo 0
    
    'move temp file
    Name FileName & ".tmp" As FileName
End Sub



