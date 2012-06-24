VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Application"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "Code Comment Add-In"
'/// Code Comment Add-In
DefStr S
DefLng L
DefBool B
DefInt I
Option Explicit
'May 12, 2003

Implements IDTExtensibility

Private newMenu   As CommandBarControl
Private VBE       As VBE
Private WithEvents mnuCom2Browser As CommandBarEvents
Attribute mnuCom2Browser.VB_VarHelpID = -1
Private WithEvents mnuBrowser2Com As CommandBarEvents
Attribute mnuBrowser2Com.VB_VarHelpID = -1
Private WithEvents mnuDocument    As CommandBarEvents
Attribute mnuDocument.VB_VarHelpID = -1
Private WithEvents mnuEraser      As CommandBarEvents
Attribute mnuEraser.VB_VarHelpID = -1

Private Declare Function WritePrivateProfileString& Lib _
"kernel32" Alias "WritePrivateProfileStringA" _
(ByVal AppName$, ByVal KeyName$, ByVal _
keydefault$, ByVal Filename$)

Private Const conComment = "'///"
Private Const conNewLine = "   " '3 spaces

Private Sub AddToINI()
    Dim rc As Long
    rc = WritePrivateProfileString("Add-Ins32", _
    "CodeComment.Application", "0", "VBADDIN.INI")
    MsgBox "Add-in is now entered in VBADDIN.INI file."
End Sub


Private Function AskOverwrite(OldComment As String, NewComment As String) As Integer
    If StrComp(NewComment, OldComment, vbTextCompare) = 0 Then AskOverwrite = 2: Exit Function
    
    If MsgBox("Overwrite '" & OldComment & "' with '" & NewComment & "' ?", vbYesNo Or vbQuestion) = vbYes Then _
    AskOverwrite = 1 Else AskOverwrite = 0
End Function

Private Function FormatComment(Description As String) As String
    FormatComment = conComment & " " & Replace(Description, conNewLine, vbNewLine & conComment & " ", , vbTextCompare)
End Function


Private Function GetDeclaration(Parent As CodeModule, objMember As Member) As String
    Dim lngLines, lngLocation, lngLoop
    On Error Resume Next
    
    If objMember Is Nothing Then Exit Function
    
    Select Case objMember.Type
    Case vbext_mt_Method
        lngLocation = Parent.ProcStartLine(objMember.Name, vbext_pk_Proc)
        lngLines = Parent.ProcCountLines(objMember.Name, vbext_pk_Proc)
    Case vbext_mt_Property
        
        'get
        lngLocation = Parent.ProcStartLine(objMember.Name, vbext_pk_Get)
        lngLines = Parent.ProcCountLines(objMember.Name, vbext_pk_Get)
        If Err Then
            'let
            Err.Clear
            lngLocation = Parent.ProcStartLine(objMember.Name, vbext_pk_Let)
            lngLines = Parent.ProcCountLines(objMember.Name, vbext_pk_Let)
            If Err Then
                'set
                Err.Clear
                lngLocation = Parent.ProcStartLine(objMember.Name, vbext_pk_Set)
                lngLines = Parent.ProcCountLines(objMember.Name, vbext_pk_Set)
                If Err Then
                    Err.Clear
                    Exit Function
                End If
            End If
        End If
    Case vbext_mt_Event, vbext_mt_Const, vbext_mt_Variable
        lngLocation = objMember.CodeLocation 'Parent.ProcStartLine(objMember.Name, vbext_pk_Proc)
        lngLines = 1
    End Select
    
    If lngLocation = 1 Then Exit Function 'no declaration to be retreived
    
    Dim strRead
    Dim vntSplit, vntItem
    Dim blnReadNext
    
    For lngLoop = lngLocation To lngLocation + lngLines
        If lngLoop > Parent.CountOfLines Then Exit For
        strRead = Trim$(Parent.Lines(lngLoop, 1))
        
        If blnReadNext Then
            blnReadNext = False
            GetDeclaration = GetDeclaration & strRead
        Else
            vntSplit = Split(Trim$(strRead), " ")
        
            For Each vntItem In vntSplit
                'commands related to the declaration of a statement
                Select Case LCase(vntItem)
                Case "sub", "function", "property", "enum", "type"
                    GetDeclaration = GetDeclaration & strRead
                    blnReadNext = IsLineContinued(strRead)
                    Exit For
                Case "end"
                    Exit For
                Case "public", "friend", "private"
                    'keep going
                End Select
            Next vntItem
        End If
    Next lngLoop
End Function

Private Function GetNameOfVBType(ByVal T As vbext_ComponentType) As String
    Select Case T
    Case vbext_ComponentType.vbext_ct_ActiveXDesigner
        GetNameOfVBType = "ActiveX Designer"
    Case vbext_ComponentType.vbext_ct_ClassModule
        GetNameOfVBType = "Class"
    Case vbext_ComponentType.vbext_ct_DocObject
        GetNameOfVBType = "Document Object"
    Case vbext_ComponentType.vbext_ct_MSForm
        GetNameOfVBType = "MS Form"
    Case vbext_ComponentType.vbext_ct_PropPage
        GetNameOfVBType = "Property Page"
    Case vbext_ComponentType.vbext_ct_RelatedDocument
        GetNameOfVBType = "Related Document"
    Case vbext_ComponentType.vbext_ct_ResFile
        GetNameOfVBType = "Resource File"
    Case vbext_ComponentType.vbext_ct_StdModule
        GetNameOfVBType = "Standard Module"
    Case vbext_ComponentType.vbext_ct_UserControl
        GetNameOfVBType = "User Control"
    Case vbext_ComponentType.vbext_ct_VBForm
        GetNameOfVBType = "Form"
    Case vbext_ComponentType.vbext_ct_VBMDIForm
        GetNameOfVBType = "MDI Form"
    Case Else
        GetNameOfVBType = "Visual Basic Component"
    End Select
End Function

Private Function GetPath(Filename As String) As String
    Dim lngI As Long
    
    lngI = InStrRev(Filename, "\")
    If lngI > 0 Then
        GetPath = Left$(Filename, lngI)
    End If
End Function

Private Function IsGood(Str As String, lngEncounter As Long) As Boolean
    Dim blnGood As Boolean
    Dim vntSplit
    Dim vntItem
    
    If Left$(LTrim$(Str), 1) = "'" Then IsGood = True: Exit Function
    
    If Len(Trim$(Str)) > 0 Then
        vntSplit = Split(Trim$(Str), " ")
        
        For Each vntItem In vntSplit
            'commands related to the declaration of a statement
            Select Case LCase(vntItem)
            Case "sub", "function", "property", "enum", "type"
                If lngEncounter < 2 Then
                    lngEncounter = lngEncounter + 1
                    blnGood = True
                Else
                    blnGood = False
                End If
                Exit For
            Case "end"
                blnGood = False
                Exit For
            Case "public", "friend", "private"
                If lngEncounter < 2 Then
                    lngEncounter = lngEncounter + 1
                    blnGood = True
                Else
                    blnGood = False
                End If
                Exit For
            End Select
        Next vntItem
    End If
    
    
    IsGood = (Left$(Str, Len(conComment)) = conComment) Or _
            Len(Str) = 0 Or _
            blnGood
End Function

Private Function IsLineContinued(Line As String) As Boolean
    Dim strS
    strS = Trim$(Line)
    IsLineContinued = Left$(strS, 1) <> "'" And Right$(strS, 1) = "_"
End Function


Private Function ReadComment(ByVal Parent As CodeModule, ByVal objMember As Member, ByRef lngCountOfCommentLines As Long, ByRef lngModuleStart As Long, ByRef lngCommentStart As Long) As String
    Dim strRead     As String 'read
    Dim lngLocation As Long 'location of module
    Dim lngLines    As Long 'length of module
    Dim lngCount    As Long 'max count of lines in code
    Dim strNew      As String '= ""
    Dim lngGood     As Long
    Dim lngLine     As Long 'line being processed
    Dim blnCont
    Dim blnLast
    
    lngCommentStart = 0
    lngCountOfCommentLines = 0
    lngModuleStart = 0
    lngCount = Parent.CountOfLines
    
    On Error Resume Next
    
    If objMember Is Nothing Then
        'read top line
        lngLocation = 1
        lngLines = 1
    Else
        Select Case objMember.Type
        Case vbext_mt_Method
            lngLocation = Parent.ProcStartLine(objMember.Name, vbext_pk_Proc)
            lngLines = Parent.ProcCountLines(objMember.Name, vbext_pk_Proc)
        Case vbext_mt_Property
            
            'get
            lngLocation = Parent.ProcStartLine(objMember.Name, vbext_pk_Get)
            lngLines = Parent.ProcCountLines(objMember.Name, vbext_pk_Get)
            If Err Then
                'let
                Err.Clear
                lngLocation = Parent.ProcStartLine(objMember.Name, vbext_pk_Let)
                lngLines = Parent.ProcCountLines(objMember.Name, vbext_pk_Let)
                If Err Then
                    'set
                    Err.Clear
                    lngLocation = Parent.ProcStartLine(objMember.Name, vbext_pk_Set)
                    lngLines = Parent.ProcCountLines(objMember.Name, vbext_pk_Set)
                    If Err Then
                        Err.Clear
                        Exit Function
                    End If
                End If
            End If
        Case vbext_mt_Event, vbext_mt_Const, vbext_mt_Variable
            lngLocation = objMember.CodeLocation 'Parent.ProcStartLine(objMember.Name, vbext_pk_Proc)
            
            lngLines = 1
        End Select
    End If
    
    
    If Not objMember Is Nothing And lngLocation = 1 Then lngLocation = 0
    
    lngModuleStart = lngLocation
    
    If lngLocation = 0 Or Err <> 0 Then Exit Function
    
    lngLine = lngLocation
    
    Debug.Assert StrComp(objMember.Name, "MsgBoxChoices", vbTextCompare)
    
    With Parent
        Do
            strRead = Trim$(.Lines(lngLine, 1))
            blnLast = blnCont
            blnCont = IsLineContinued(strRead)
            
            If Left$(strRead, Len(conComment)) = conComment Then
                If Len(strNew) Then strNew = strNew & conNewLine  'new line
                strNew = strNew & LTrim$(Mid$(strRead, Len(conComment) + 1))
                
                lngCountOfCommentLines = lngCountOfCommentLines + 1
                If lngCommentStart = 0 Then lngCommentStart = lngLine
            End If
            
            lngLine = lngLine + 1 'increment line
        Loop While (Len(strRead) = 0 Or IsGood(strRead, lngGood) Or blnLast) _
        And lngLine <= lngCount And lngLine <= lngLocation + lngLines
    End With
    
    #If DebugMode Then
    If Not objMember Is Nothing Then
        Debug.Print objMember.Name, conComment & strNew
    End If
    #End If
    ReadComment = strNew
End Function




Private Sub Class_Terminate()
    Unload frmUpdate
    Set frmUpdate = Nothing
End Sub


Private Sub IDTExtensibility_OnAddInsUpdate(custom() As Variant)
    'nothing
End Sub

Private Sub IDTExtensibility_OnConnection(ByVal VBInst As Object, ByVal ConnectMode As VBIDE.vbext_ConnectMode, ByVal AddInInst As VBIDE.AddIn, custom() As Variant)
    Dim myMenuBar As CommandBar
    Dim com2browser As CommandBarButton
    Dim browser2com  As CommandBarButton
    Dim document     As CommandBarButton
    Dim eraser       As CommandBarButton
    
    Set VBE = AddInInst.VBE

    Set myMenuBar = AddInInst.VBE.CommandBars("Add-Ins")
    
    Set newMenu = myMenuBar.Controls.Add(Type:=msoControlPopup, Temporary:=True)
    newMenu.Caption = "&CodeComment"
    
    Set com2browser = newMenu.Controls.Add(Type:=msoControlButton, Id:=1)
    com2browser.Caption = "&1 Comment to Object Browser"
    com2browser.ToolTipText = "Comment to Object Browser"
    
    Set browser2com = newMenu.Controls.Add(Type:=msoControlButton, Id:=1)
    browser2com.Caption = "&2 Object Browser to Comment"
    browser2com.ToolTipText = "Object Browser to Comment"
    
    Set document = newMenu.Controls.Add(Type:=msoControlButton, Id:=1)
    document.Caption = "&3 Document this module"
    document.ToolTipText = "Document this module"
    
    Set eraser = newMenu.Controls.Add(Type:=msoControlButton, Id:=1)
    eraser.Caption = "&4 Erase all comments"
    eraser.ToolTipText = "Erase all comments"
    
    
     Set mnuBrowser2Com = VBE.Events.CommandBarEvents(browser2com)
     Set mnuCom2Browser = VBE.Events.CommandBarEvents(com2browser)
     Set mnuDocument = VBE.Events.CommandBarEvents(document)
     Set mnuEraser = VBE.Events.CommandBarEvents(eraser)
End Sub


Private Sub IDTExtensibility_OnDisconnection(ByVal RemoveMode As VBIDE.vbext_DisconnectMode, custom() As Variant)
    If Not newMenu Is Nothing Then newMenu.Delete
End Sub


Private Sub IDTExtensibility_OnStartupComplete(custom() As Variant)
    'nothing
End Sub

Private Sub mnuBrowser2Com_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
    Dim objMember As Member
    Dim lngLines    As Long
    Dim strRead     As String
    Dim lngCounter  As Long
    Dim lngUpdates  As Long
    Dim lngModStart As Long
    Dim lngCommentStart As Long
    Dim intReturn
    
    On Error GoTo Handler
    
    If VBE.ActiveCodePane Is Nothing Then
        MsgBox "Open a code pane first.", vbInformation
        Exit Sub
    End If
    
    Load frmUpdate
    
    With VBE.ActiveCodePane.CodeModule
        'write top line
        If Len(.Parent.Description) > 0 Then
            strRead = ReadComment(VBE.ActiveCodePane.CodeModule, Nothing, lngCounter, 0, 0)
            
            'if previous comment
            If Len(strRead) Then
                'if previous comment bigger than new comment
                If Len(strRead) >= Len(.Parent.Description) Then
                    intReturn = AskOverwrite(strRead, .Parent.Description)
                    If intReturn Then
                        .ReplaceLine 1, FormatComment(.Parent.Description)
                        frmUpdate.AddItem "(CodeModule)", "Replace" & IIf(intReturn = 1, " Asked", "")
                        lngUpdates = lngUpdates + 1 'counter
                    End If
                Else
                'put bigger comment without asking
                    .ReplaceLine 1, FormatComment(.Parent.Description)
                    frmUpdate.AddItem "(CodeModule)", "Replace"
                    lngUpdates = lngUpdates + 1 'counter
                End If
            Else
                .InsertLines 1, FormatComment(.Parent.Description)
                frmUpdate.AddItem "(CodeModule)", "Insert"
                lngUpdates = lngUpdates + 1 'counter
            End If
        End If
        
        
        'write members
        For Each objMember In .Members
            If Len(objMember.Description) > 0 Then
                strRead = ReadComment(VBE.ActiveCodePane.CodeModule, objMember, lngCounter, lngModStart, lngCommentStart)
                
                'if previous comment
                If lngModStart > 0 And lngCommentStart > 0 Then
                    'if previous comment bigger than new comment
                    If Len(strRead) >= Len(objMember.Description) Then
                        intReturn = AskOverwrite(strRead, objMember.Description)
                        If intReturn Then
                            .ReplaceLine lngCommentStart, FormatComment(objMember.Description)
                            frmUpdate.AddItem objMember.Name, "Replace" & IIf(intReturn = 1, " Asked", "")
                            lngUpdates = lngUpdates + 1 'counter
                        End If
                    Else
                    'put bigger comment without asking
                        .ReplaceLine lngCommentStart, FormatComment(objMember.Description)
                        frmUpdate.AddItem objMember.Name, "Replace"
                        lngUpdates = lngUpdates + 1 'counter
                    End If
                ElseIf lngModStart > 0 Then
                    'add comment
                    .InsertLines lngModStart + 1, FormatComment(objMember.Description)
                    frmUpdate.AddItem objMember.Name, "Insert"
                    lngUpdates = lngUpdates + 1 'counter
                End If
            End If
        Next objMember
    End With
    
    If lngUpdates Then
        frmUpdate.Show vbModal
    Else
        MsgBox "No descriptions from the object browser were placed in the code.", vbInformation
    End If
    Unload frmUpdate
    
    Exit Sub
Handler:
    MsgBox Err.Description & " " & Err.Number, vbExclamation
    Exit Sub
End Sub

Private Sub mnuCom2Browser_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
    Dim objMember As Member
    Dim lngLocation As Long
    Dim strRead     As String
    Dim lngCounter  As Long
    Dim lngUpdates  As Long
    Dim lngLines    As Long
    Dim intReturn
    
    On Error GoTo Handler
    
    If VBE.ActiveCodePane Is Nothing Then
        MsgBox "Open a code pane first.", vbInformation
        Exit Sub
    End If
    
    Load frmUpdate
    With VBE.ActiveCodePane.CodeModule
        'top line in file is component's description
        strRead = ReadComment(VBE.ActiveCodePane.CodeModule, Nothing, lngCounter, 0, 0)
        
        If Len(strRead) Then
            If Len(.Parent.Description) = 0 Then
                .Parent.Description = strRead
                frmUpdate.AddItem "(CodeModule)", "Added"
                lngUpdates = lngUpdates + 1 'counter
            Else
                intReturn = AskOverwrite(.Parent.Description, strRead)
                If intReturn Then
                    .Parent.Description = strRead
                    frmUpdate.AddItem "(CodeModule)", "Changed" & IIf(intReturn = 1, " Asked", "")
                    lngUpdates = lngUpdates + 1 'counter
                End If
            End If
        End If
    
        'each member's description is below its declaration line
        For Each objMember In .Members
            strRead = ""
            strRead = ReadComment(VBE.ActiveCodePane.CodeModule, objMember, lngCounter, 0, 0)
            
            If Len(strRead) Then
                If Len(objMember.Description) = 0 Then
                    objMember.Description = strRead
                    frmUpdate.AddItem objMember.Name, "Added"
                    lngUpdates = lngUpdates + 1 'counter
                Else
                    intReturn = AskOverwrite(objMember.Description, strRead)
                    If intReturn Then
                        objMember.Description = strRead
                        frmUpdate.AddItem objMember.Name, "Change" & IIf(intReturn = 1, " Asked", "")
                        lngUpdates = lngUpdates + 1 'counter
                    End If
                End If
            End If
SkipThis:
        Next objMember
    End With
    
    If lngUpdates Then
        frmUpdate.Show vbModal
    Else
    
        MsgBox lngUpdates & " comments were placed in the object browser.", vbInformation
    End If
    Unload frmUpdate
    
    Exit Sub
Handler:
    MsgBox Err.Description & " " & Err.Number, vbExclamation
    Exit Sub
End Sub


Private Sub mnuDocument_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
    Dim F As Long
    Dim objComponent As VBComponent
    Dim objCode      As CodeModule
    Dim strLine      As String
    Dim lngLine      As Long
    Dim strComments  As String
    Dim strExposure  As String
    Dim objMember    As Member
    Dim strDir       As String
    Dim frmDoc       As frmDoc
    Dim strScope     As String
    Dim blnPublic    As Boolean '1
    Dim blnFriend    As Boolean '2
    Dim blnPrivate   As Boolean '3
    Dim blnDim      As Boolean '4
    Dim blnOK       As Boolean
    Dim blnBlank    As Boolean
    Dim vntSplit    As Variant
    Dim blnWriteThis As Boolean
        
    
    Set objComponent = VBE.SelectedVBComponent
    
    If objComponent.CodeModule Is Nothing Then
        MsgBox "There is no code associated with the " & GetNameOfVBType(objComponent.Type) & "  component.", vbInformation
        Exit Sub
    End If
    
    If Len(objComponent.FileNames(1)) = 0 Then
        MsgBox "Save the module first.", vbInformation
        Exit Sub
    End If
    
    Set frmDoc = New frmDoc
    Load frmDoc
    frmDoc.lblName = "Active module: " & objComponent.Name
    frmDoc.Show vbModal
    
    blnOK = frmDoc.m_blnOK
    blnPublic = frmDoc.lstMembers.Selected(1)
    blnFriend = frmDoc.lstMembers.Selected(2)
    blnPrivate = frmDoc.lstMembers.Selected(3)
    blnDim = frmDoc.lstMembers.Selected(4)
    blnBlank = frmDoc.chkWriteBlank
    
    Unload frmDoc
    Set frmDoc = Nothing
    
    
    'if cancelled, unload window and form and exit sub
    If Not blnOK Then Exit Sub
    
    On Error GoTo Handler1
    
    Set objCode = objComponent.CodeModule
    
    strDir = GetPath(objComponent.FileNames(1))
    ChDrive strDir
    ChDir strDir
    
    'display module information
    F = FreeFile
    Open "CodeComment_" & objComponent.Name & ".htm" For Output As #F
    
    On Error GoTo Handler2
    
    Print #F, "<!--Generated by Code Comment " & Format$(Now, "General Date") & "-->"
    Print #F, "<html>"
    Print #F, "<head><title>Documentation for "; objComponent.Name; "</title></head>"
    Print #F, "<body>"
    Print #F, "<p><b>Module name:</b> "; objComponent.Name
    Print #F, "<p><b>Description:</b> "; GetNameOfVBType(objComponent.Type)
    Print #F, "<p><b>Other notes:</b>"
    
    'read declaration section for extra notes
    For lngLine = 1 To objCode.CountOfDeclarationLines
        strLine = objCode.Lines(lngLine, 1)
        
        If Left$(strLine, 3) = "'//" And Left$(strLine, Len(conComment)) <> conComment Then
            'normal quotation
            strComments = strComments & Mid$(strLine, 4) & "<br>"
        End If
    Next lngLine
    
    Print #F, "<blockquote>"; strComments; "</blockquote></p><hr>"
    
    'display each Member in module
    For Each objMember In objCode.Members
        blnWriteThis = False
        strLine = GetDeclaration(objCode, objMember)
        vntSplit = Split(strLine, " ", 2)
        
        If UBound(vntSplit) > 0 Then
            Select Case LCase$(CStr(vntSplit(0)))
            Case "public"
                blnWriteThis = blnPublic
                strExposure = StrConv(vntSplit(0), vbProperCase)
            Case "friend"
                blnWriteThis = blnFriend
                strExposure = StrConv(vntSplit(0), vbProperCase)
            Case "private"
                blnWriteThis = blnPrivate
                strExposure = StrConv(vntSplit(0), vbProperCase)
            Case "dim"
                blnWriteThis = blnDim
                strExposure = "Private"
            Case Else 'public
                blnWriteThis = True
                strExposure = ""
            End Select
        Else
            blnWriteThis = True 'do not know what it is
            strExposure = ""
        End If
        
        If Not blnBlank And Len(objMember.Description) = 0 Then blnWriteThis = False
        
        If blnWriteThis Then
            Print #F, "<p><table border=1 cellspacing=0 cellpadding=0 cols=2 width=90%>"
            Print #F, Tab; "<tr><td bgcolor=AAFFFF><b>"; objMember.Name; "</b>  <font size=-4>"; strExposure; "</font></td><td bgcolor=C8C8C8>"; strLine; "</td></tr>"
            Print #F, Tab; "<tr><td colspan=2>"; objMember.Description; "</td></tr></table></p>"
            Print #F, "<p>"
        End If
    Next objMember
    
    
    Print #F, "</body>"
    Print #F, "</html>"
    Close #F
    
    

    Exit Sub
Handler1:
    MsgBox "Problem while attempting to write documentation.", vbInformation
    
    If Not frmDoc Is Nothing Then
        Unload frmDoc
        Set frmDoc = Nothing
    End If
    Exit Sub
    
Handler2:
    MsgBox "Problem while writing documentation.", vbInformation
    
    If Not frmDoc Is Nothing Then
        Unload frmDoc
        Set frmDoc = Nothing
    End If
    
    Close #F
    
    Exit Sub
End Sub

    
Private Sub mnuEraser_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
    Dim objMember As Member
    Dim lngLocation As Long
    Dim strRead     As String
    Dim lngCounter  As Long
    Dim lngUpdates  As Long
    
    On Error GoTo Handler
    
    If VBE.ActiveCodePane Is Nothing Then
        MsgBox "Open a code pane first.", vbInformation
        Exit Sub
    ElseIf MsgBox("Are you sure you want to erase all the comments from the object browser?", vbQuestion Or vbYesNo) <> vbYes Then
        Exit Sub
    End If
    
    Load frmUpdate
    
    With VBE.ActiveCodePane.CodeModule
        'write top line
        .Parent.Description = ""
        frmUpdate.AddItem "(CodeModule)", "Erased"
        lngUpdates = lngUpdates + 1 'counter
        
        
        'write members
        For Each objMember In .Members
            lngLocation = objMember.CodeLocation
            
            objMember.Description = ""
            frmUpdate.AddItem objMember.Name, "Erased"
            lngUpdates = lngUpdates + 1 'counter
        Next objMember
    End With
    
    If lngUpdates Then
        frmUpdate.Show vbModal
    Else
        MsgBox "No descriptions from the object browser were erased.", vbInformation
    End If
    Unload frmUpdate
    
    Exit Sub
Handler:
    MsgBox Err.Description & " " & Err.Number, vbExclamation
    Exit Sub

End Sub


