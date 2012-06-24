VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3195
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   4680
   ControlBox      =   0   'False
   FillColor       =   &H80000012&
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   12
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   FontTransparent =   0   'False
   Icon            =   "Strands_Main.frx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1000
   ScaleMode       =   0  'User
   ScaleWidth      =   1000
   Begin MSComDlg.CommonDialog dlgFile 
      Left            =   2160
      Top             =   1320
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin VB.TextBox txtEdit 
      Alignment       =   2  'Center
      BorderStyle     =   0  'None
      Height          =   735
      Left            =   1680
      TabIndex        =   1
      Top             =   1200
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Shape shpDrag 
      DrawMode        =   6  'Mask Pen Not
      FillStyle       =   0  'Solid
      Height          =   375
      Left            =   1680
      Top             =   600
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label lblText 
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   1815
      WordWrap        =   -1  'True
   End
   Begin VB.Menu mnuBlank 
      Caption         =   ""
      Visible         =   0   'False
      Begin VB.Menu mnuBlankAdd 
         Caption         =   "&Add Thought"
      End
      Begin VB.Menu mnuBlankSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuBlankNew 
         Caption         =   "&New Mind"
         Begin VB.Menu mnuBlankNewBlank 
            Caption         =   "&Blank"
         End
         Begin VB.Menu mnuBlankNewFolder 
            Caption         =   "From &folders..."
            Index           =   0
         End
         Begin VB.Menu mnuBlankNewFolder 
            Caption         =   "From &files..."
            Index           =   1
         End
         Begin VB.Menu mnuBlankNewFolder 
            Caption         =   "From &extension..."
            Index           =   2
         End
         Begin VB.Menu mnuBlankNewCopy 
            Caption         =   "&Copy from last"
         End
      End
      Begin VB.Menu mnuBlankOpen 
         Caption         =   "&Open Mind..."
      End
      Begin VB.Menu mnuBlankSave 
         Caption         =   "&Save Mind"
      End
      Begin VB.Menu mnuBlankRevert 
         Caption         =   "&Revert"
      End
      Begin VB.Menu mnuBlankSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuBlankScheme 
         Caption         =   "S&cheme"
         Begin VB.Menu mnuBlankSchemeItem 
            Caption         =   "&Sky"
            Index           =   0
         End
         Begin VB.Menu mnuBlankSchemeItem 
            Caption         =   "W&ater"
            Index           =   1
         End
         Begin VB.Menu mnuBlankSchemeItem 
            Caption         =   "&Desktop"
            Index           =   2
         End
         Begin VB.Menu mnuBlankSchemeItem 
            Caption         =   "&Grass"
            Index           =   3
         End
         Begin VB.Menu mnuBlankSchemeItem 
            Caption         =   "&Beige"
            Index           =   4
         End
         Begin VB.Menu mnuBlankSchemeItem 
            Caption         =   "&Overcast"
            Index           =   5
         End
         Begin VB.Menu mnuBlankSchemeItem 
            Caption         =   "&Wood"
            Index           =   6
         End
         Begin VB.Menu mnuBlankSchemeSep1 
            Caption         =   "-"
         End
         Begin VB.Menu mnuBlankSchemeCircle 
            Caption         =   "Circle"
         End
      End
      Begin VB.Menu mnuBlankFont 
         Caption         =   "&Font"
         Begin VB.Menu mnuBlankFontSize 
            Caption         =   "&Extra Small"
            Index           =   8
         End
         Begin VB.Menu mnuBlankFontSize 
            Caption         =   "&Small"
            Index           =   10
         End
         Begin VB.Menu mnuBlankFontSize 
            Caption         =   "&Medium"
            Index           =   12
         End
         Begin VB.Menu mnuBlankFontSize 
            Caption         =   "&Large"
            Index           =   14
         End
         Begin VB.Menu mnuBlankFontSize 
            Caption         =   "&Extra Large"
            Index           =   16
         End
         Begin VB.Menu mnuBlankFontSep1 
            Caption         =   "-"
         End
         Begin VB.Menu mnuBlankFontHighlight 
            Caption         =   "&2-fold highlight"
            Index           =   2
         End
         Begin VB.Menu mnuBlankFontHighlight 
            Caption         =   "&4-fold highlight"
            Index           =   4
         End
         Begin VB.Menu mnuBlankFontHighlight 
            Caption         =   "&6-fold highlight"
            Index           =   6
         End
         Begin VB.Menu mnuBlankFontHighlight 
            Caption         =   "&8-fold highlight"
            Index           =   8
         End
      End
      Begin VB.Menu mnuBlankChange 
         Caption         =   "&Change"
         Begin VB.Menu mnuBlankChangeAuthor 
            Caption         =   "&Author"
         End
         Begin VB.Menu mnuBlankChangeComment 
            Caption         =   "&Comment"
         End
      End
      Begin VB.Menu mnuBlankFind 
         Caption         =   "F&ind"
         Begin VB.Menu mnuBlankFindIdea 
            Caption         =   "&Idea"
         End
         Begin VB.Menu mnuBlankFindText 
            Caption         =   "&Text"
         End
         Begin VB.Menu mnuBlankFindMissing 
            Caption         =   "&Missing Attachment"
         End
      End
      Begin VB.Menu mnuBlankSep3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuBlankExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuItem 
      Caption         =   ""
      Visible         =   0   'False
      Begin VB.Menu mnuItemConnect 
         Caption         =   "&Connect Line"
      End
      Begin VB.Menu mnuItemSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuItemRun 
         Caption         =   "&Run Thought"
      End
      Begin VB.Menu mnuItemEdit 
         Caption         =   "&Edit Thought"
      End
      Begin VB.Menu mnuItemDelete 
         Caption         =   "&Delete Thought"
      End
      Begin VB.Menu mnuItemSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuItemText 
         Caption         =   "Edit &Text"
         Index           =   1
      End
      Begin VB.Menu mnuItemAttach 
         Caption         =   "&Attach File..."
      End
      Begin VB.Menu mnuItemDeattach 
         Caption         =   "Dea&ttach"
         Begin VB.Menu mnuItemDeattachFile 
            Caption         =   "&File"
         End
         Begin VB.Menu mnuItemDeattachPicture 
            Caption         =   "&Picture"
         End
      End
      Begin VB.Menu mnuItemConnections 
         Caption         =   "&Connections"
         Begin VB.Menu mnuItemConnectionsFind 
            Caption         =   "&Find"
         End
         Begin VB.Menu mnuItemConnectionsDel 
            Caption         =   "&Delete"
         End
      End
   End
   Begin VB.Menu mnuSpace 
      Caption         =   ""
      Visible         =   0   'False
      Begin VB.Menu mnuSpaceAdd 
         Caption         =   "&Add Thought"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Richard Fung.
'March 30, 2000.
'Main frm

Public Mind As BaseThoughts
Public FileName As String
'Click
Public Over As Thought
Public MouseX  As Single
Public MouseY  As Single
Public Mousebtn As MouseButtonConstants
Public Selected As Thought
Public CancelDblClick  As Boolean
Public MouseDown As Boolean
'Drag
Public DownX   As Single
Public DownY   As Single
Public DifX    As Single
Public DifY    As Single
'Colour schemes,etc.
Private mintNormalSize As Integer '= 12
Private mlngNormalFore As Long '= vbBlack
Private mlngNormalBack As Long '= &HFFFFD0
Private mintHighlightSize As Integer '= 14
Private mlngHighlightFore As Long '= vbBlue
Private mintTextSize As Integer '= 10
Private mstrTextFace As String '= "Tahoma"
Private mlngTextFore As Long '= vbBlack
Private mintLevel    As Integer 'hints
Private mstrInput    As String 'hints
Private mblnEllipse  As Boolean



Private Sub DrawMind(Optional ByVal FirstTime As Boolean)
    Dim intT As Integer
    Dim intR As Integer
    Dim tConnect As Thought
    Dim sngX(1) As Single
    Dim sngY(1) As Single
    Static blnNextTime As Boolean
    
    Cls
    'set colors
    BackColor = mlngNormalBack
    ForeColor = mlngNormalFore
    Font.Size = mintNormalSize
    DrawWidth = 1
    
    'picture
    lblText.Visible = False
    If (Not Selected Is Nothing) Then
        On Error Resume Next
        If Len(Selected.Picture) Then
            Set Picture = LoadPicture(Selected.Picture)
            If Err Then
                Err.Clear
                Set Picture = Nothing
            End If
        Else
            Set Picture = Nothing
        End If
        On Error GoTo 0
    End If
    
    'cheat - stars
    If Mind.Header.ColourScheme = 8501 Then
        Dim intX As Integer
        Dim intY As Integer
        Dim intC As Integer
        Dim intS As Integer
        
        MousePointer = vbArrowHourglass
        Randomize Timer
        For intS = 1 To (Width / 1000) * (Height / 1000)
            intX = Int(Rnd * (conWidth + 1))
            intY = Int(Rnd * (conHeight + 1))
            intC = Int(Rnd * 3) + 1
            Select Case intC
                Case 1: ForeColor = vbWhite
                Case 2: ForeColor = vbBlack
                Case 3: ForeColor = &HC0C0C0
            End Select
            PSet (intX, intY), ForeColor
        Next intS
        ForeColor = mlngNormalFore
        DrawWidth = 1
        MousePointer = vbDefault
    End If
    
    'connect lines
    For intT = 1 To Mind.Count
        With Mind(intT)
            .Refresh Me
            For Each tConnect In .Links
                'determine if it NEED BE deleted
                If Mind.Find(tConnect) Is Nothing Then
                    .Links.Remove Object:=tConnect
                Else
                    Line (.CenterX, .CenterY)-(tConnect.CenterX, tConnect.CenterY)
                End If
            Next tConnect
            
            If mnuBlankSchemeCircle.Checked Then
                sngX(0) = ScaleX(.Left - .Width / 2, ScaleMode, vbPixels)
                sngX(1) = sngX(0) + ScaleX(.Width * 2, ScaleMode, vbPixels)
                sngY(0) = ScaleY(.Top - .Height / 2, ScaleMode, vbPixels)
                sngY(1) = sngY(0) + ScaleY(.Height * 2, ScaleMode, vbPixels)
                If Not Mind(intT) Is Selected Then
                    FillColor = BackColor
                    FillStyle = vbFSSolid
                    Call Ellipse(hdc, sngX(0), sngY(0), sngX(1), sngY(1))
                    FillStyle = vbFSTransparent
                End If
            End If
        End With
    Next intT
    
    'text
    For intT = 1 To Mind.Count
'        If Mind(intT) Is Selected Then
'            ForeColor = mlngHighlightFore
'            Font.Size = mintHighlightSize
'            Selected.Refresh Me
'        ElseIf Font.Size <> mintNormalSize Then
'            ForeColor = mlngNormalFore
'            Font.Size = mintNormalSize
'        End If
        With Mind(intT)
            CurrentX = .Left
            CurrentY = .Top
            Print .Idea
        End With
        If Len(Mind(intT).Attachment) Then
            Line (Mind(intT).Left, Mind(intT).Top + Mind(intT).Height)-Step(Mind(intT).Width, 0)
        End If
    Next intT
    
    mnuItemText_Click 0
    ShowSelected
    If (FirstTime Or blnNextTime) And Not Selected Is Nothing Then
        If Not Visible Then
            blnNextTime = True 'show next time
        Else
            ForeColor = mlngTextFore
            Font.Size = mintTextSize
            CurrentX = Selected.Left
            CurrentY = Selected.Top + Selected.Height
            Print "Author: "; Mind.Header.Author
            CurrentX = Selected.Left
            Print "Last Modified: "; Mind.Header.DateModified
            blnNextTime = False
        End If
    End If
    
    'other information
    If Not Selected Is Nothing Then
        If Picture.Handle = 0 And Len(Selected.Picture) Then
            lblText = "Attached picture is missing: """ & Selected.Picture & """"
            lblText.Visible = True
        ElseIf Picture.Handle Then
            lblText.Visible = False
        Else
            If Len(Mind.Header.Comment) Then
                lblText = Mind.Header.Comment & vbNewLine & Selected.Text
            Else
                lblText = Selected.Text
            End If
            lblText.Visible = Len(lblText)
        End If
    Else
        lblText.Visible = False
        lblText = ""
    End If
End Sub


Private Sub DrawOther()
    If BorderStyle <> 1 Then
        If Len(Mind.Header.Author) Then
            Caption = "Strands (" & Mind.Header.Author & ")"
        Else
            Caption = "Strands"
        End If
    End If
End Sub

Public Sub Highlight(ByVal Item As Thought, Optional ByVal DoubleBold As Boolean)
    ForeColor = mlngHighlightFore
    If DoubleBold Then
        DrawWidth = 2
    Else
        DrawWidth = 1
    End If
    Line (Item.Left, Item.Top)-Step(Item.Width, Item.Height), , B
End Sub

Private Function Hit(ByVal x As Single, ByVal y As Single) As Thought
    Dim objT As Thought
    For Each objT In Mind
        If objT.OverThis(x, y) Then
            Set Hit = objT
        End If
    Next objT
End Function

Private Sub ImportFile()
'    Dim objI As Thought
'
    MousePointer = vbHourglass
    
    Set Mind = OpenFile(FileName)
    If Mind.Header.LastSelected = 0 Or Mind.Header.LastSelected > Mind.Count Then
        Set Selected = Nothing
    Else
        Set Selected = Mind(Mind.Header.LastSelected)
    End If
    Set Over = Selected
    
    mnuBlankSchemeCircle.Checked = Mind.Header.Circled
    Scheme Mind.Header.ColourScheme
    DrawOther
    DrawMind True
    If Mind.Header.EditTextWindow Then
        mnuItemText_Click 1
    End If

    MousePointer = vbDefault
End Sub

Private Sub NewMind()
    Mind.Add "My Mind", "", "", "", ScaleWidth / 2, ScaleHeight / 2, Me
    Set Selected = Mind(1)
End Sub



Private Sub LoadRegistry()
    On Error GoTo Handler
    mintHighlightSize = GetSetting("Strands", "Size", "Highlight", mintHighlightSize)
    mintNormalSize = GetSetting("Strands", "Size", "Normal", mintNormalSize)
    If mintHighlightSize <= mintNormalSize Then
        mintHighlightSize = mintNormalSize + 2
    End If
    mintTextSize = mintNormalSize - 2
    Exit Sub
Handler:
End Sub

Private Function QuerySave() As Boolean
    On Error GoTo SaveError
    If Len(FileName) Then
        Save
        QuerySave = Good
    Else
        Select Case MsgBox("Do you want to save thoughts?", vbExclamation Or vbYesNoCancel, "Save Mind")
            Case vbYes
                Save
                QuerySave = Good
            Case vbCancel
                QuerySave = Fail
            Case vbNo
                QuerySave = Good
        End Select
    End If
    Exit Function
SaveError:
    If Err <> cdlCancel Then
        MsgBox "There is a problem saving the thoughts because " & Err.Description & ".", vbExclamation, "Save Mind"
    End If
    QuerySave = Fail
End Function

Private Sub RunAssociation(ByVal File As String)
    Dim r As Long 'return
    
    MousePointer = vbHourglass
    r = ShellExecute(hwnd, vbNullString, File, vbNullString, vbNullString, vbNormalFocus)
    If r <= 32 Then
        MsgBox "Cannot run associated file: """ & File & """.", vbExclamation, "Run Association"
    End If
    MousePointer = vbDefault
End Sub

Private Sub RunThisProgram(ByVal ExeName As String)
    MousePointer = vbHourglass
    On Error Resume Next
    Shell ExeName, vbNormalFocus
    If Err Then
        lblText = "Problem running the file """ & ExeName & """ because " & Err.Description & "."
        lblText.Visible = True
        Err.Clear
    End If
    MousePointer = vbDefault
End Sub

Private Sub Save()
    If Len(FileName) = 0 Then
        With dlgFile
            .DialogTitle = "Save Thoughts"
            .FileName = FileName
            .Filter = "Thought Files - Fibre|*.fib;*.fibre"
            .FilterIndex = 1
            .Flags = cdlOFNOverwritePrompt Or cdlOFNHideReadOnly Or cdlOFNPathMustExist
            .ShowSave
            FileName = .FileName
        End With
    End If
    Mind.Header.Circled = mnuBlankSchemeCircle.Checked
    SaveFile FileName, Mind, frmText.Visible, Selected
End Sub

Private Sub SaveRegistry()
    On Error GoTo Handler
    SaveSetting "Strands", "Size", "Highlight", mintHighlightSize
    SaveSetting "Strands", "Size", "Normal", mintNormalSize
    Exit Sub
Handler:
    MsgBox "Error saving registry setting", vbExclamation, "Save Settings"
    Resume Next
End Sub

Private Sub Scheme(ByVal Scheme As Integer)
    If Not Mind Is Nothing Then
        Mind.Header.ColourScheme = Scheme
    End If
    Select Case Scheme
        Case 0 'sky
            mlngNormalFore = vbBlack
            mlngNormalBack = &HFFFFD0
            mlngHighlightFore = vbBlue
            mlngTextFore = vbBlack
        Case 1 'water
            mlngNormalFore = vbWhite
            mlngNormalBack = &H800000
            mlngHighlightFore = vbYellow
            mlngTextFore = vbYellow
        Case 2 'desktop
            mlngNormalFore = vbWhite
            mlngNormalBack = &H808000
            mlngHighlightFore = vbYellow
            mlngTextFore = vbYellow
        Case 3 'grass
            mlngNormalFore = &H800000
            mlngNormalBack = &HC0FFC0
            mlngHighlightFore = &H7000&
            mlngTextFore = &H7000&
        Case 4 'beige
            mlngNormalFore = vbBlack
            mlngNormalBack = &HC0E0FF
            mlngHighlightFore = 32896
            mlngTextFore = 16512
        Case 5 'overcast
            mlngNormalFore = &H808080
            mlngNormalBack = &HE0E0E0
            mlngHighlightFore = vbBlack
            mlngTextFore = &H808080
        Case 6 'wood
            mlngNormalFore = vbBlack
            mlngNormalBack = &H80C0FF
            mlngHighlightFore = 16512
            mlngTextFore = 16512
        Case 8501 'born yr 85/01 RF
            mlngNormalFore = vbWhite
            mlngNormalBack = vbBlack
            mlngHighlightFore = vbYellow
            mlngTextFore = &HC0C0C0
    End Select

    'text
    lblText.Font.Name = mstrTextFace
    lblText.Font.Size = mintTextSize
    lblText.ForeColor = mlngTextFore
    lblText.BackColor = mlngNormalBack
End Sub

Private Sub ShowSelected()
    Static objLast  As Thought
    Dim l           As Single
    Dim sngX(1)     As Single
    Dim sngY(1)     As Single
    Dim sngSize     As Single
    
    If (Not Selected Is Nothing) Then
        ForeColor = mlngHighlightFore
        
        If (Not objLast Is Selected) Then
            l = mintNormalSize
        Else
            l = mintHighlightSize
        End If
        CurrentY = Selected.Top
        
        For sngSize = l To mintHighlightSize Step 0.5
            Font.Size = sngSize
            Selected.Width = TextWidth(Selected)
            Selected.Left = Selected.CenterX - Selected.Width / 2
            Selected.Height = TextHeight(Selected)
            
            If mnuBlankSchemeCircle.Checked Then
                With Selected
                    sngX(0) = ScaleX(.Left - .Width / 2, ScaleMode, vbPixels)
                    sngX(1) = sngX(0) + ScaleX(.Width * 2, ScaleMode, vbPixels)
                    sngY(0) = ScaleY(.Top - .Height / 2, ScaleMode, vbPixels)
                    sngY(1) = sngY(0) + ScaleY(.Height * 2, ScaleMode, vbPixels)
                End With
                FillColor = BackColor
                FillStyle = vbFSSolid
                Call Ellipse(hdc, sngX(0), sngY(0), sngX(1), sngY(1))
                FillStyle = vbFSTransparent
            End If
            
            CurrentX = Selected.Left
            Print Selected;
            Refresh
        Next sngSize
        
        If Len(Selected.Attachment) Then
            Line (Selected.Left, Selected.Top + Selected.Height)-Step(Selected.Width, 0)
        End If
        Refresh
    End If
    Set objLast = Selected
End Sub

Private Sub Form_Click()
    If txtEdit.Visible Then
        txtEdit.Visible = False
        CancelDblClick = True
        Exit Sub
    End If
    
    If Mousebtn = vbLeftButton Then
        If Not Over Is Nothing Then
            Set Selected = Over
            DrawMind
        End If
    ElseIf Mousebtn = vbRightButton Then
        If Over Is Nothing Then
            PopupMenu mnuBlank, vbPopupMenuLeftButton Or vbPopupMenuRightButton
        Else
            PopupMenu mnuItem, vbPopupMenuLeftButton Or vbPopupMenuRightButton
        End If
    End If
End Sub

Private Sub Form_DblClick()
    If Mousebtn = vbLeftButton And (Not Over Is Nothing) And (Not CancelDblClick) Then
        If Len(Over.Attachment) Then
            mnuItemRun_Click
        Else
            mnuItemEdit_Click
        End If
    ElseIf Mousebtn = vbLeftButton And (Over Is Nothing) And (Not CancelDblClick) Then
        PopupMenu Menu:=mnuSpace, defaultmenu:=mnuSpaceAdd
    End If
    CancelDblClick = False
End Sub


Private Sub Form_Initialize()
    mintNormalSize = 12
    mintHighlightSize = 14
    mintTextSize = 10
    mstrTextFace = "Tahoma"
    Scheme 0
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 93 Then
        If Selected Is Nothing Or (Shift And vbCtrlMask) Then
            PopupMenu mnuBlank, , 0, 0
        Else
            Set Over = Selected
            If Len(Over.Attachment) Then
                PopupMenu mnuItem, , Selected.CenterX, Selected.CenterY, mnuItemRun
            Else
                PopupMenu mnuItem, , Selected.CenterX, Selected.CenterY
            End If
        End If
    ElseIf KeyCode = vbKeyTab And Shift = 0 Then
        If Not Selected Is Nothing Then
            If Selected.Index = 1 Then
                Set Selected = Mind(Mind.Count)
            Else
                Set Selected = Mind(Selected.Index - 1)
            End If
        ElseIf Mind.Count > 0 Then
            Set Selected = Mind(Mind.Count)
        End If
        Set Over = Selected
        DrawMind
    ElseIf KeyCode = vbKeyTab And (Shift And vbShiftMask) Then
        If Not Selected Is Nothing Then
            If Selected.Index = Mind.Count Then
                Set Selected = Mind(1)
            Else
                Set Selected = Mind(Selected.Index + 1)
            End If
        ElseIf Mind.Count > 0 Then
            Set Selected = Mind(1)
        End If
        Set Over = Selected
        DrawMind
    ElseIf KeyCode = vbKeyF1 Then
        frmAbout1.Show vbModal, Me
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    On Error GoTo Handler
    If mintLevel = 4 And KeyAscii = vbKeyReturn Then
        Dim objItem As Thought
        Dim objLast As Thought
        Dim objObj  As Object
        Dim lngNum  As Long
        Dim lngNo2  As Long
        #If Cheats = 2 Then
        Dim blnFail As Boolean
        #End If
        
        MousePointer = vbHourglass
        Randomize Timer
        Select Case mstrInput
            Case "black"
                Mind.Header.ColourScheme = 8501
                Scheme 8501
                DrawMind
            Case "beep"
                Beep
            Case "explorer"
                RunAssociation "Explorer.exe"
            Case "notepad"
                RunAssociation "Notepad.exe"
            Case "command.com"
                RunAssociation "Command.com"
            Case "quoth"
                MsgBox "To be, or not to be: that is the question:" & vbNewLine & "Whether 'tis nobler in the mind to suffer" & vbNewLine & "The slings and arrows of outrageous fortune," & vbNewLine & "Or to take arms against a sea of troubles," & vbNewLine & "And by opposing end them.  To die: to sleep;" & vbNewLine & "No more; and by a sleep to say we end" & vbNewLine & "The heart-ache, and the thousand natural shocks" & vbNewLine & "That flesh is heir to, 'tis a consummation" & vbNewLine & "Devoutly to be wish'd.  To die, to sleep;" & vbNewLine & "To sleep: perchance to dream: ay, there's the rub;" & vbNewLine & "For in that sleep of death what dreams may come..." & vbNewLine & vbNewLine & "- Act III, Scene I soliloquy, William Shakespeare", , "Hamlet"
            Case "quote"
                MsgBox "If you fail to plan, plan to fail."
            Case "quoto"
                MsgBox "We may attempt to achieve that [goals] at all costs, even if it carries possible punishment, even death."
            Case "quoteme", "quoteMe"
                MsgBox "Created by Richard Fung, April 2000, at Calgary AB, on the IPC CS-FX Multimedia computer.  Using Microsoft® Visual Basic 5.0 (SP3).  Queen Elizabeth Jr/Sr High School."
            Case "eeks"
                Set objLast = Selected
                For Each objItem In Mind
                    Set Selected = objItem
                    DrawMind
                    Refresh
                Next objItem
                Set Selected = objLast
                Set objLast = Nothing
                DrawMind
            Case "restore"
                WindowState = vbNormal
            Case "maximize"
                WindowState = vbMaximized
            Case "icon_out"
                Set Icon = Nothing
            Case "randomize"
                Set objLast = Selected
                For Each objItem In Mind
                    objItem.CenterX = Int(Rnd * (conWidth + 1))
                    objItem.CenterY = Int(Rnd * (conHeight + 1))
                    DrawMind
                    Refresh
                Next objItem
                Set Selected = objLast
                Set objLast = Nothing
                DrawMind
            Case "saucer"
                If Not Selected Is Nothing Then
                    lngNo2 = Selected.CenterX
                    For lngNum = 0 To conWidth Step (conWidth / 50)
                        Selected.CenterX = lngNum
                        DrawMind
                    Next lngNum
                    Selected.CenterX = lngNo2
                    DrawMind
                End If
            Case "aboutme", "aboutMe"
                frmAbout.Show 1, Me
            Case "winword"
                Set objObj = CreateObject("Word.Application")
                objObj.Visible = True
                Set objObj = Nothing
            Case "time"
                MsgBox Format(Time, "Long Time")
            Case "date"
                MsgBox Format(Date, "Long Date")
            Case "AST"
                MsgBox "AST Computer DX/33 4 w/ Mb ram & 200 Mb hard drive"
            Case "romeo", "romeo and juliet"
                Set objObj = frmAbout.imgShk.Picture
                Unload frmAbout
                PaintPicture objObj, ScaleWidth / 2 - ScaleX(objObj.Width, vbHimetric, ScaleMode) / 2, ScaleHeight / 2 - ScaleY(objObj.Height, vbHimetric, ScaleMode) / 2
                Set objObj = Nothing
                MsgBox "Romeo and Juliet by William Shakespeare"
            Case "telephone"
                If Not Selected Is Nothing Then
                    For lngNum = 0 To 50
                        Select Case Int(Rnd * 4)
                            Case 0 'north
                                Selected.CenterX = Selected.CenterX - 10
                            Case 1 'south
                                Selected.CenterX = Selected.CenterX + 10
                            Case 2 'west
                                Selected.CenterY = Selected.CenterY - 10
                            Case 3 'east
                                Selected.CenterY = Selected.CenterY + 10
                        End Select
                        DrawMind
                    Next lngNum
                End If
            Case "rich"
                If Not Selected Is Nothing Then
                    CurrentX = Selected.CenterX
                    CurrentY = Selected.CenterY
                    DrawWidth = 2
                    For lngNum = 1 To 300
                        ForeColor = QBColor(Int(Rnd * 16))
                        Select Case Int(Rnd * 4)
                            Case 0 'north
                                Line -(CurrentX - 15, CurrentY), ForeColor
                            Case 1 'south
                                Line -(CurrentX + 15, CurrentY), ForeColor
                            Case 2 'west
                                Line -(CurrentX, CurrentY - 15), ForeColor
                            Case 3 'east
                                Line -(CurrentX, CurrentY + 15), ForeColor
                        End Select
                        Refresh
                    Next lngNum
                End If
            Case "communication revolution"
                Set objObj = frmAbout.imgEarth.Picture
                Unload frmAbout
                PaintPicture objObj, ScaleWidth / 2 - ScaleX(objObj.Width, vbHimetric, ScaleMode) / 2, ScaleHeight / 2 - ScaleY(objObj.Height, vbHimetric, ScaleMode) / 2
                Set objObj = Nothing
            Case "quit", "exit"
                Unload Me
                Exit Sub 'don't process rest of Sub
            Case "I see", "i see", "I see!", "I see with my own eyes"
                Set objLast = Selected
                For Each objItem In Mind
                    CurrentX = objItem.Left
                    CurrentY = objItem.Top
                    PaintPicture Icon, CurrentX, CurrentY, ScaleX(Icon.Width, vbHimetric, ScaleMode), ScaleY(Icon.Height, vbHimetric, ScaleMode)
                Next objItem
                Set Selected = objLast
                Set objLast = Nothing
            Case "about"
                frmAbout1.Show vbModal, Me
            Case "refresh"
                DrawOther
                DrawMind
            #If Cheats = 2 Then
            Case Else
                blnFail = True
            #End If
        End Select
        #If Cheats = 2 Then
            Font.Size = 6
            Print ">" & mstrInput
            If blnFail Then
                Print "Command not accepted."
            Else
                Print "Done."
            End If
        #End If
        mstrInput = ""
        MousePointer = vbDefault
    ElseIf mintLevel = 4 And KeyAscii = vbKeyEscape Then
        mstrInput = ""
    ElseIf mintLevel = 4 Then
        mstrInput = mstrInput & Chr(KeyAscii)
    End If
    Exit Sub
Handler:
    MsgBox "?"
End Sub

Private Sub Form_Load()
'    lblSelected.Font.Size = mintHighlightSize
'    lblSelected.ForeColor = mlngHighlightFore
'    lblSelected.BackColor = mlngNormalBack
    
    Set Mind = New BaseThoughts
    
    LoadRegistry
    
    WindowState = vbMaximized
    
    If Len(Command) = 0 Then
        NewMind
    Else
        FileName = Command
        ImportFile
    End If
    
    #If Cheats Then
        mintLevel = 4 'cheats enabled
    #End If
End Sub


Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If txtEdit.Visible Then Exit Sub
    
    If Button Then
        Mousebtn = Button 'click
        MouseDown = True
        MouseX = x 'click
        MouseY = y
        DownX = x 'drag-drop
        DownY = y
        Set Over = Hit(x, y)
        If Not Over Is Nothing Then 'for drag
            DifX = x - Over.Left
            DifY = y - Over.Top
            shpDrag.Width = Over.Width
            shpDrag.Height = Over.Height
        End If
    End If
End Sub

'Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    Dim t As Thought
'
'    If Mousebtn = vbRightButton Then
'        Set t = Hit(X, Y)
'        If Not t Is Nothing Then
'            Set Selected = t
'        End If
'        Set t = Nothing
'    End If
'End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If txtEdit.Visible Then Exit Sub
    
    If Button And MouseDown Then
        Mousebtn = Button 'click
        MouseX = x
        MouseY = y
        'detect if drag needed
        If Not Over Is Nothing Then
            'bound checks
            If MouseX < DifX Then MouseX = DifX
            If MouseX > conWidth - shpDrag.Width + DifX Then MouseX = conWidth - shpDrag.Width + DifX
            If MouseY < DifY Then MouseY = DifY
            If MouseY > conHeight - shpDrag.Height + DifY Then MouseY = conHeight - shpDrag.Height + DifY
            'second line for sensitivity to mouse
            If shpDrag.Visible Then
                shpDrag.Move MouseX - DifX, MouseY - DifY
            ElseIf ((MouseX < DownX - 25) Or (MouseX > DownX + 25)) Or _
              ((MouseY < DownY - 25) Or (MouseY > DownY + 25)) Or (Shift And vbShiftMask) Then
                shpDrag.Move MouseX - DifX, MouseY - DifY
                shpDrag.Visible = True
            End If
        End If
    End If
End Sub


Private Sub Form_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If txtEdit.Visible Then Exit Sub
    
    If Button And MouseDown Then
        Mousebtn = Button
        MouseX = x
        MouseY = y
        
        'drag ONLY handled here
        If shpDrag.Visible Then
            'bound checks
            If MouseX < DifX Then MouseX = DifX
            If MouseX > conWidth - shpDrag.Width + DifX Then MouseX = conWidth - shpDrag.Width + DifX
            If MouseY < DifY Then MouseY = DifY
            If MouseY > conHeight - shpDrag.Height + DifY Then MouseY = conHeight - shpDrag.Height + DifY
            
            Over.CenterX = MouseX - DifX + Over.Width / 2
            Over.CenterY = MouseY - DifY + Over.Height / 2
            Mousebtn = 0
            shpDrag.Visible = False
            DrawMind
        Else
            Set Over = Hit(x, y)
        End If
    End If
    MouseDown = False
End Sub




Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If QuerySave = Fail Then
        Cancel = 1
    End If
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    If WindowState <> vbMinimized Then
        Scale (0, 0)-(1000, 1000)
        DrawMind
    End If
End Sub



Private Sub Form_Unload(Cancel As Integer)
    'just in case
    Unload frmAbout
    Unload frmAbout1
    Unload frmFolder
    Unload frmText
    'save registry
    SaveRegistry
End Sub


'
'
'
'
'
'Private Sub lblText_Click()
'    Form_Click
'End Sub
'
'Private Sub lblText_DblClick()
'    Form_DblClick
'End Sub
'
'
'Private Sub lblText_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    Form_MouseDown Button, Shift, X, Y
'End Sub
'
'
'Private Sub lblText_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    Form_MouseMove Button, Shift, X, Y
'End Sub
'
'
'Private Sub lblText_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    Form_MouseUp Button, Shift, X, Y
'End Sub
'

Private Sub mnuBlank_Click()
    Dim intItem As Integer
    
    mnuBlankRevert.Enabled = Len(FileName)
    
    'mnuBlankScheme
    
    For intItem = mnuBlankSchemeItem.LBound To mnuBlankSchemeItem.UBound
        mnuBlankSchemeItem(intItem).Checked = (intItem = Mind.Header.ColourScheme)
    Next intItem
    
    'mnuBlankFont
    
    For intItem = mnuBlankFontSize.LBound To mnuBlankFontSize.UBound Step 2
        mnuBlankFontSize(intItem).Checked = (mintNormalSize = intItem)
    Next intItem
    For intItem = mnuBlankFontHighlight.LBound To mnuBlankFontHighlight.UBound Step 2
        mnuBlankFontHighlight(intItem).Checked = ((mintHighlightSize - mintNormalSize) = intItem)
    Next intItem
End Sub


Private Sub mnuBlankAdd_Click()
    Set Over = Mind.Add("New idea", "", "", "", MouseX, MouseY, Me)
    DrawMind
    mnuItemEdit_Click
End Sub



Private Sub mnuBlankChangeAuthor_Click()
    Dim strText As String
    
    strText = InputBox("Enter in your full name.", "Modify Author", Mind.Header.Author)
    If Len(strText) = 0 Then
        If MsgBox("Remove author from this set of thoughts?", vbQuestion Or vbYesNo Or vbDefaultButton2, "Modify Author") = vbYes Then
            Mind.Header.Author = strText
        End If
    Else
        Mind.Header.Author = strText
    End If
    DrawOther
End Sub

Private Sub mnuBlankChangeComment_Click()
    Dim strText As String
    
    strText = InputBox("Enter in a comment to be shown in every item.", "Modify Comment", Mind.Header.Comment)
    If Len(strText) = 0 Then
        If MsgBox("Remove comment from this set of thoughts?", vbQuestion Or vbYesNo Or vbDefaultButton2, "Modify Author") = vbYes Then
            Mind.Header.Comment = strText
        End If
    Else
        Mind.Header.Comment = strText
    End If
    DrawMind
End Sub


Private Sub mnuBlankExit_Click()
    Unload Me
End Sub

Private Sub mnuBlankFindIdea_Click()
    Dim objItem As Thought
    Dim strText As String
    
    strText = InputBox("Enter in the Idea you want to find.", "Find Idea")
    If Len(strText) = 0 Then Exit Sub
    
    For Each objItem In Mind
        If InStr(1, objItem.Idea, strText, vbTextCompare) > 0 Then
            Highlight objItem
        End If
    Next objItem
End Sub

Private Sub mnuBlankFindMissing_Click()
    Dim objItem As Thought
    Dim r As Integer 'response
    
    On Error Resume Next
    For Each objItem In Mind
        If Len(objItem.Attachment) Then
            r = GetAttr(objItem.Attachment)
            If Err Then
                Err.Clear
                Highlight objItem
            End If
        End If
        If Len(objItem.Picture) Then
            r = GetAttr(objItem.Picture)
            If Err Then
                Err.Clear
                Highlight objItem
            End If
        End If
    Next objItem
End Sub

Private Sub mnuBlankFindText_Click()
    Dim objItem As Thought
    Dim strText As String
    
    strText = InputBox("Enter in the Text you want to find.", "Find Text")
    If Len(strText) = 0 Then Exit Sub
    
    For Each objItem In Mind
        If InStr(1, objItem.Text, strText, vbTextCompare) > 0 Then
            Highlight objItem
        End If
    Next objItem
End Sub




Private Sub mnuBlankNewBlank_Click()
    If QuerySave = False Then Exit Sub
    
    MousePointer = vbHourglass
    
    Set Mind = New BaseThoughts
    NewMind
    
    FileName = ""
    Set Over = Nothing
    
    Scheme Mind.Header.ColourScheme
    DrawOther
    DrawMind
    MousePointer = vbDefault
End Sub



Private Sub mnuBlankNewCopy_Click()
    FileName = ""
End Sub

'Default changes go here.
Private Sub mnuBlankNewFolder_Click(Index As Integer)
    Dim strFile     As String
    Dim strFolder   As String
    Dim strExt      As String
    Dim strFTitle   As String
    
    
    If QuerySave = False Then Exit Sub
    
    MousePointer = vbHourglass
   
    If Not frmFolder.GetFolder((Index = 2), Me, strFolder, strExt) Then
        MousePointer = vbDefault
        Exit Sub
    End If
    
    Set Mind = New BaseThoughts
    
    FileName = ""
    Set Over = Nothing
    Set Selected = Nothing
    
    Scheme Mind.Header.ColourScheme
    DrawOther
    DrawMind
    If Right(strFolder, 1) <> "\" Then
        strFolder = strFolder & "\"
    End If
    
    If Index = 1 Then
        strFTitle = Dir(strFolder)
    ElseIf Index = 2 Then
'        Dim vntFTitle As Variant
'        Dim strExt  As String
'
'        vntFTitle = GroupCharVB(dlgFile.FileTitle, ".", False)
'        strExt = vntFTitle(UBound(vntFTitle))
        strFTitle = Dir(strFolder & strExt)
    Else
        strFTitle = Dir(strFolder, vbDirectory)
    End If
    strFile = strFolder & strFTitle
    
    Do While strFTitle <> ""
        If UCase(strFTitle) = strFTitle Then
            strFTitle = StrConv(strFTitle, vbProperCase)
        End If
        If Index Then 'file
            If Not (GetAttr(strFile) And vbDirectory) Then
                Mind.Add strFTitle, "", strFile, "file", 1000, 1000, Me ' Int(Rnd * 1001), Int(Rnd * 1001), Me
            End If
        Else 'folder
            If (GetAttr(strFile) And vbDirectory) And strFTitle <> "." And strFTitle <> ".." Then
                Mind.Add strFTitle, "", strFile, "file", 1000, 1000, Me ' Int(Rnd * 1001), Int(Rnd * 1001), Me
            End If
        End If
        If Mind.Count > 50 Then
            MsgBox "The number of items is restricted to 50 due to memory considerations.", vbExclamation, "New Mind"
            Exit Do
        End If
        strFTitle = Dir
        strFile = strFolder & strFTitle
    Loop
    
    If Mind.Count = 1 Then
        Mind(1).CenterX = conWidth / 2
        Mind(1).CenterY = conHeight / 2
        GoTo OutHere
    ElseIf Mind.Count = 2 Then
        Mind(1).CenterX = conWidth / 2
        Mind(1).CenterY = conHeight * (1 / 3)
        Mind(2).CenterX = Mind(1).CenterX
        Mind(2).CenterY = conHeight * (2 / 3)
        GoTo OutHere
    ElseIf Mind.Count = 3 Then
        Mind(1).CenterX = conWidth / 2
        Mind(1).CenterY = conHeight * (1 / 4)
        Mind(2).CenterX = Mind(1).CenterX
        Mind(2).CenterY = conHeight * (2 / 4)
        Mind(3).CenterX = Mind(1).CenterX
        Mind(3).CenterY = conHeight * (3 / 4)
        GoTo OutHere
    ElseIf Mind.Count = 0 Then
        NewMind
        GoTo OutHere
    End If
    
    'order the items
    Dim intSqrCount As Integer
    Dim sngXStep As Single
    Dim sngYStep As Single
    Dim intX     As Integer
    Dim intY     As Integer
    Dim intRunCount As Integer
    
    intSqrCount = CInt(Sqr(Mind.Count))
    sngXStep = (conWidth - 70) \ intSqrCount
    sngYStep = (conHeight - 70) \ intSqrCount
    intRunCount = 0
    For intY = 1 To (intSqrCount + 1)
        For intX = 1 To intSqrCount
            intRunCount = intRunCount + 1
            If intRunCount > Mind.Count Then GoTo OutHere
            If intX = 1 And intY = 1 Then 'don't put in first square
                intX = intX + 1
            End If
            Mind(intRunCount).CenterX = (intX - 1) * sngXStep + 35
            Mind(intRunCount).CenterY = (intY - 1) * sngYStep + 35
        Next intX
    Next intY
OutHere:
    
    DrawMind
    MousePointer = vbDefault
End Sub


Private Sub mnuBlankSchemeCircle_Click()
    mnuBlankSchemeCircle.Checked = Not mnuBlankSchemeCircle.Checked
    DrawMind
End Sub


Private Sub mnuItemConnectionsDel_Click()
    Set frmDelLink.Thoughts = Mind
    Set frmDelLink.LinkItem = Over
    frmDelLink.ReadLinks
    frmDelLink.Show vbModal, Me
    DrawMind 'changes might have been made
End Sub


Private Sub mnuItemConnectionsFind_Click()
    Dim objItem As Thought
    Dim objConnect As Thought
    
    For Each objItem In Mind
        For Each objConnect In objItem.Links
            If objConnect Is Over Then
                Highlight objItem
            ElseIf objItem Is Over Then
                Highlight objConnect
            End If
        Next objConnect
    Next objItem
    Highlight Over, True
End Sub

Private Sub mnuBlankFontHighlight_Click(Index As Integer)
    mintHighlightSize = mintNormalSize + Index
    DrawMind
End Sub

Private Sub mnuBlankFontSize_Click(Index As Integer)
    mintHighlightSize = mintHighlightSize - mintNormalSize + Index
    mintNormalSize = Index
    mintTextSize = Index - 2
    Scheme Mind.Header.ColourScheme
    DrawMind
End Sub


Private Sub mnuBlankOpen_Click()
    On Error Resume Next
    With dlgFile
        .DialogTitle = "Open Thoughts"
        .FileName = FileName
        .Filter = "Thought Files - Fibre|*.fib;*.fibre"
        .FilterIndex = 1
        .Flags = cdlOFNFileMustExist Or cdlOFNHideReadOnly
        .ShowOpen
        If Err <> cdlCancel Then
            FileName = .FileName
            On Error GoTo 0
            ImportFile
        End If
    End With
End Sub

Private Sub mnuBlankRevert_Click()
    If MsgBox("Do you want to reload the thoughts from the last saved instance?", vbQuestion Or vbYesNo, "Revert") = vbYes Then
        ImportFile
    End If
End Sub


Private Sub mnuBlankSave_Click()
    On Error Resume Next
    Save
    If Err <> cdlCancel And Err <> 0 Then
        MsgBox "There is a problem saving the thoughts because " & Err.Description & ".", vbExclamation, "Save Mind"
        Err.Clear
    End If
End Sub

Private Sub mnuBlankSchemeItem_Click(Index As Integer)
    Scheme Index
    DrawMind
End Sub

Private Sub mnuItem_Click()
    mnuItemConnect.Enabled = False
    If (Not Selected Is Nothing) And (Not Over Is Nothing) Then
        If Selected.Index < Over.Index Then
            If Selected.Links.Find(Over) Is Nothing Then
                mnuItemConnect.Enabled = True
            End If
        ElseIf Selected.Index > Over.Index Then
            If Over.Links.Find(Selected) Is Nothing Then
                mnuItemConnect.Enabled = True
            End If
        End If
    End If
    
    'mnuItemDeattach
    
    mnuItemRun.Enabled = Len(Over.Attachment)
    mnuItemDeattachFile.Enabled = Len(Over.Attachment)
    mnuItemDeattachPicture.Enabled = Len(Over.Picture)
End Sub

Private Sub mnuItemAttach_Click()
    Dim strAttach As String
    
    On Error Resume Next
    With dlgFile
        .DialogTitle = "Attach File"
        .Filter = "Executable Files|*.exe;*.bat;*.com;*.pif|Picture Files|*.bmp;*.gif;*.jpg;*.ico;*.wmf"
        .FilterIndex = 1
        .FileName = Over.Attachment
        If Len(Over.Picture) <> 0 And Len(Over.Attachment) = 0 Then
            .FilterIndex = 2 'change to 2
            .FileName = Over.Picture
        End If
        .Flags = cdlOFNFileMustExist Or cdlOFNHideReadOnly
        .ShowOpen
        If Err = cdlCancel Then
            Exit Sub
        End If
        strAttach = .FileName
    End With
    
    If Len(strAttach) > 0 Then
        Select Case LCase(Right(strAttach, 4))
            Case ".exe", ".bat", ".com", ".pif"
                Over.Attachment = strAttach
                Over.AttachmentTag = "exe"
            Case ".bmp", ".gif", ".jpg", ".ico", ".wmf"
                Over.Picture = strAttach
            Case Else
                Over.Attachment = strAttach
                Over.AttachmentTag = "file"
        End Select
        DrawMind
    End If
End Sub

Private Sub mnuItemConnect_Click()
    If (Not Selected Is Nothing) And (Not Over Is Nothing) And (Not Over Is Selected) Then
        If Selected.Index = Over.Index Or Selected Is Over Then 'just in case
            'do nothing
        ElseIf Selected.Index < Over.Index Then
            If Selected.Links.Find(Over) Is Nothing Then
                Selected.Links.Add Over
            End If
        ElseIf Selected.Index > Over.Index Then
            If Over.Links.Find(Selected) Is Nothing Then
                Over.Links.Add Selected
            End If
        End If
        DrawMind
    End If
End Sub


Private Sub mnuItemDeattachFile_Click()
    Over.Attachment = ""
    Over.AttachmentTag = ""
    DrawMind
End Sub


Private Sub mnuItemDeattachPicture_Click()
    Over.Picture = ""
    DrawMind
End Sub


Private Sub mnuItemDelete_Click()
    If Not Over Is Nothing Then
        If mintLevel = 3 And Over = "Bertolt Brecht" Then
            mintLevel = 4
        End If
        Mind.Remove Over.Index
        If Over Is Selected Then
            Set Selected = Nothing
        End If
        Set Over = Nothing
        DrawMind
    End If
End Sub

Private Sub mnuItemEdit_Click()
    If Not Over Is Nothing Then
        If Over Is Selected Then
            Font.Size = mintHighlightSize
            txtEdit.Font.Size = mintHighlightSize
        Else
            Font.Size = mintNormalSize
            txtEdit.Font.Size = mintNormalSize
        End If
        txtEdit.Move Over.Left, Over.Top, Over.Width + TextWidth("O"), Over.Height
        txtEdit.Tag = Over.Idea
        txtEdit = Over.Idea
        txtEdit.Visible = True
    End If
End Sub


Private Sub mnuItemRun_Click()
    If Over.AttachmentTag = "exe" Then
        RunThisProgram Over.Attachment
    ElseIf Over.AttachmentTag = "file" Then
        RunAssociation Over.Attachment
    End If
End Sub

'Also used by DrawMind.
Private Sub mnuItemText_Click(Index As Integer)
    Set frmText.Attached = Over
    If Over Is Nothing Then
        frmText.Hide
    ElseIf Index Then
        frmText.txtEdit.Font.Name = mstrTextFace
        frmText.txtEdit.Font.Size = mintTextSize
        If Index Then
            Dim cursor As POINTAPI
            
            Call GetCursorPos(cursor)
            cursor.x = ScaleX(cursor.x, vbPixels, vbTwips)
            cursor.y = ScaleY(cursor.y, vbPixels, vbTwips)
            frmText.Move cursor.x - frmText.Width / 2, cursor.y - frmText.Height / 2
        End If
        frmText.Show OwnerForm:=Me
    End If
End Sub


Private Sub mnuSpaceAdd_Click()
    mnuBlankAdd_Click
End Sub

Private Sub txtEdit_Change()
    Dim s As Single
    
    Over.Idea = txtEdit
    Over.Refresh Me
    's = txtEdit.Left + txtEdit.Width / 2 - TextWidth(txtEdit) / 2
    txtEdit.Width = Over.Width + TextWidth("O")
    txtEdit.Left = Over.Left
End Sub


Private Sub txtEdit_GotFocus()
    SendKeys "{Home}+{End}"
End Sub


Private Sub txtEdit_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        txtEdit.Visible = False
        'validate text
        If mintNormalSize = 8 Then
            If txtEdit = "Our aspirations and goals are the foundation to an interesting and unique lifetime." And txtEdit.Tag = "New idea" Then
                'from 26-Mar-00 document by Richard Fung
                mintLevel = 1
            ElseIf txtEdit = "The Life of Galileo" And mintLevel = 1 And txtEdit.Tag = "Our aspirations and goals are the foundation to an interesting and unique lifetime." Then
                'Title of the play
                mintLevel = 2
            ElseIf txtEdit = "And do you know that this discovery of yours, which you have described as the fruit of 17 years research" And mintLevel = 2 And txtEdit.Tag = "The Life of Galileo" Then
                'Richard, playing Curator, speaks this line
                txtEdit = "Bertolt Brecht"
                mintLevel = 3
            End If
        ElseIf mintLevel Then
            mintLevel = 0 'clear level
        End If
    ElseIf KeyAscii = vbKeyEscape Then 'cancel changes
        KeyAscii = 0
        txtEdit = txtEdit.Tag 'revert changes
        txtEdit.Visible = False
    End If
End Sub

Private Sub txtEdit_LostFocus()
    DrawMind
End Sub


