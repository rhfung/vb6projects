VERSION 5.00
Begin VB.Form frmObject 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Format"
   ClientHeight    =   4620
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8505
   Icon            =   "Flow_Object.frx":0000
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4620
   ScaleWidth      =   8505
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraPic2 
      Caption         =   "More Picture..."
      Height          =   3855
      Left            =   5520
      TabIndex        =   54
      Top             =   120
      Visible         =   0   'False
      Width           =   2895
      Begin VB.OptionButton opPicDef 
         Caption         =   "Use &default size"
         Height          =   255
         Left            =   120
         TabIndex        =   57
         Top             =   960
         Width           =   1695
      End
      Begin VB.ComboBox cboPicRatio 
         Height          =   315
         ItemData        =   "Flow_Object.frx":000C
         Left            =   1800
         List            =   "Flow_Object.frx":0028
         Style           =   2  'Dropdown List
         TabIndex        =   56
         Top             =   360
         Width           =   855
      End
      Begin VB.OptionButton opPicRatio 
         Caption         =   "Use &scaled size"
         Height          =   255
         Left            =   120
         TabIndex        =   55
         Top             =   360
         Width           =   1575
      End
      Begin VB.CommandButton cmdPicMoreScales 
         Caption         =   "More scales..."
         Height          =   375
         Left            =   480
         TabIndex        =   58
         Top             =   1560
         Width           =   1575
      End
      Begin VB.Label lblPicHint 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   1  'Fixed Single
         Caption         =   $"Flow_Object.frx":0051
         ForeColor       =   &H80000017&
         Height          =   1455
         Left            =   360
         TabIndex        =   59
         Top             =   2160
         Width           =   1935
      End
   End
   Begin VB.Frame fraGroup 
      Caption         =   "Group"
      Height          =   3855
      Left            =   5520
      TabIndex        =   49
      Top             =   120
      Visible         =   0   'False
      Width           =   2895
      Begin VB.OptionButton opUngroup 
         Caption         =   "Ungroup"
         Height          =   255
         Left            =   240
         TabIndex        =   53
         Top             =   1800
         Width           =   2535
      End
      Begin VB.OptionButton opGroupRemove 
         Caption         =   "Remove this from group"
         Height          =   255
         Left            =   240
         TabIndex        =   52
         Top             =   1320
         Width           =   2415
      End
      Begin VB.OptionButton opGroupKeep 
         Caption         =   "Keep group"
         Height          =   255
         Left            =   240
         TabIndex        =   51
         Top             =   840
         Value           =   -1  'True
         Width           =   2415
      End
      Begin VB.Label lblGroupCount 
         Height          =   255
         Left            =   120
         TabIndex        =   50
         Top             =   360
         Width           =   2655
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3120
      TabIndex        =   42
      Top             =   4080
      Width           =   1095
   End
   Begin VB.Frame fraInfo 
      Caption         =   "Information"
      Height          =   1095
      Left            =   120
      TabIndex        =   22
      Top             =   1680
      Width           =   5295
      Begin VB.Label lblGroup 
         Height          =   255
         Left            =   4080
         TabIndex        =   46
         Top             =   240
         Width           =   975
      End
      Begin VB.Label lblInfoG 
         Caption         =   "Group Count:"
         Height          =   255
         Left            =   2880
         TabIndex        =   45
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label lblDescript 
         Height          =   255
         Index           =   1
         Left            =   1440
         TabIndex        =   48
         Top             =   600
         Width           =   3735
      End
      Begin VB.Label lblDescript 
         Caption         =   "Description:"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   47
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label lblIndex 
         Height          =   255
         Index           =   1
         Left            =   1440
         TabIndex        =   44
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label lblIndex 
         Caption         =   "Item number:"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   43
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.Frame fraArrow 
      Caption         =   "&Arrow"
      Height          =   3855
      Left            =   5520
      TabIndex        =   27
      Top             =   120
      Visible         =   0   'False
      Width           =   2895
      Begin VB.TextBox txtArrowCaption 
         Height          =   285
         Left            =   1320
         TabIndex        =   35
         Top             =   1800
         Width           =   1335
      End
      Begin VB.ComboBox cboArrowType 
         Height          =   315
         ItemData        =   "Flow_Object.frx":00EB
         Left            =   1320
         List            =   "Flow_Object.frx":00F5
         Style           =   2  'Dropdown List
         TabIndex        =   33
         Top             =   1200
         Width           =   1335
      End
      Begin VB.ComboBox cboArrowHead 
         Height          =   315
         ItemData        =   "Flow_Object.frx":010E
         Left            =   1320
         List            =   "Flow_Object.frx":0118
         Style           =   2  'Dropdown List
         TabIndex        =   31
         Top             =   720
         Width           =   1335
      End
      Begin VB.ComboBox cboArrowSize 
         Height          =   315
         ItemData        =   "Flow_Object.frx":012E
         Left            =   1320
         List            =   "Flow_Object.frx":014A
         Style           =   2  'Dropdown List
         TabIndex        =   29
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label lblArrowCaption 
         Caption         =   "&Caption:"
         Height          =   255
         Left            =   120
         TabIndex        =   34
         Top             =   1800
         Width           =   1215
      End
      Begin VB.Label lblArrowType 
         Caption         =   "Arrow type:"
         Height          =   255
         Left            =   120
         TabIndex        =   32
         Top             =   1200
         Width           =   1215
      End
      Begin VB.Label lblArrowHead 
         Caption         =   "Arrow head:"
         Height          =   255
         Left            =   120
         TabIndex        =   30
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label lblArrowSize 
         Caption         =   "Arrow size: (%)"
         Height          =   255
         Left            =   120
         TabIndex        =   28
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Frame fraName 
      Caption         =   "&Name"
      Height          =   1095
      Left            =   120
      TabIndex        =   23
      Top             =   2880
      Width           =   5295
      Begin VB.ComboBox cboName 
         Height          =   315
         ItemData        =   "Flow_Object.frx":0173
         Left            =   840
         List            =   "Flow_Object.frx":017A
         TabIndex        =   25
         Top             =   360
         Width           =   2535
      End
      Begin VB.CommandButton cmdNameUpdate 
         Caption         =   "Update Now"
         Height          =   375
         Left            =   3480
         TabIndex        =   26
         Top             =   360
         Width           =   1575
      End
      Begin VB.Label lblName 
         Caption         =   "Name:"
         Height          =   255
         Left            =   120
         TabIndex        =   24
         Top             =   360
         Width           =   855
      End
   End
   Begin VB.Frame fraField 
      Caption         =   "&Field"
      Height          =   1095
      Left            =   120
      TabIndex        =   19
      Top             =   1680
      Visible         =   0   'False
      Width           =   5295
      Begin VB.ComboBox cboField 
         Height          =   315
         ItemData        =   "Flow_Object.frx":0186
         Left            =   960
         List            =   "Flow_Object.frx":01B4
         Style           =   2  'Dropdown List
         TabIndex        =   21
         Top             =   240
         Width           =   2295
      End
      Begin VB.Label lblField 
         Caption         =   "Field for"
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.Frame fraButton 
      Caption         =   "&Button"
      Height          =   1095
      Left            =   120
      TabIndex        =   13
      Top             =   1680
      Visible         =   0   'False
      Width           =   5295
      Begin VB.CommandButton cmdMacroFile 
         Caption         =   "..."
         Height          =   255
         Left            =   4680
         TabIndex        =   18
         Top             =   720
         Width           =   375
      End
      Begin VB.TextBox txtMacroFile 
         Height          =   285
         Left            =   1080
         Locked          =   -1  'True
         TabIndex        =   17
         Top             =   720
         Width           =   3495
      End
      Begin VB.ComboBox cboMacro 
         Height          =   315
         ItemData        =   "Flow_Object.frx":0237
         Left            =   1080
         List            =   "Flow_Object.frx":0247
         Style           =   2  'Dropdown List
         TabIndex        =   15
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label lblMacroFile 
         Caption         =   "Filename:"
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   720
         Width           =   975
      End
      Begin VB.Label lblMacro 
         Caption         =   "Macro:"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.Frame fraPic 
      Caption         =   "&Picture"
      Height          =   1095
      Left            =   120
      TabIndex        =   9
      Top             =   1680
      Visible         =   0   'False
      Width           =   5295
      Begin VB.CommandButton cmdPicture 
         Caption         =   "..."
         Height          =   255
         Left            =   4680
         TabIndex        =   12
         Top             =   240
         Width           =   375
      End
      Begin VB.TextBox txtPicture 
         Height          =   285
         Left            =   1080
         Locked          =   -1  'True
         TabIndex        =   11
         Top             =   240
         Width           =   3495
      End
      Begin VB.Label lblPicture 
         Caption         =   "Filename:"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.Frame fraLocation 
      Caption         =   "Location"
      Height          =   1455
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5295
      Begin VB.TextBox txtZBase 
         Height          =   285
         Left            =   1080
         TabIndex        =   61
         Top             =   960
         Width           =   1095
      End
      Begin VB.TextBox txtZHeight 
         Height          =   285
         Left            =   3600
         TabIndex        =   60
         Top             =   960
         Width           =   1095
      End
      Begin VB.TextBox txtHeight 
         Height          =   285
         Left            =   3600
         TabIndex        =   8
         Top             =   600
         Width           =   1095
      End
      Begin VB.TextBox txtWidth 
         Height          =   285
         Left            =   1080
         TabIndex        =   6
         Top             =   600
         Width           =   1095
      End
      Begin VB.TextBox txtTop 
         Height          =   285
         Left            =   3600
         TabIndex        =   4
         Top             =   240
         Width           =   1095
      End
      Begin VB.TextBox txtLeft 
         Height          =   285
         Left            =   1080
         TabIndex        =   2
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "Z Base:"
         Height          =   255
         Left            =   120
         TabIndex        =   63
         Top             =   960
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Z elevation:"
         Height          =   255
         Left            =   2640
         TabIndex        =   62
         Top             =   960
         Width           =   975
      End
      Begin VB.Label lblHeight 
         Caption         =   "&Height:"
         Height          =   255
         Left            =   2640
         TabIndex        =   7
         Top             =   600
         Width           =   975
      End
      Begin VB.Label lblWidth 
         Caption         =   "&Width:"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   600
         Width           =   975
      End
      Begin VB.Label lblTop 
         Caption         =   "&Top:"
         Height          =   255
         Left            =   2640
         TabIndex        =   3
         Top             =   240
         Width           =   975
      End
      Begin VB.Label lblLeft 
         Caption         =   "&Left:"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.Frame fraCD 
      Caption         =   "&CD List"
      Height          =   3855
      Left            =   5520
      TabIndex        =   36
      Top             =   120
      Visible         =   0   'False
      Width           =   2895
      Begin VB.CheckBox chkCDFolder1 
         Caption         =   "Read first folder level"
         Height          =   255
         Left            =   240
         TabIndex        =   39
         Top             =   1200
         Width           =   2295
      End
      Begin VB.TextBox txtCDList 
         Height          =   2175
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   40
         Top             =   1560
         Width           =   2655
      End
      Begin VB.DriveListBox drvCDList 
         Height          =   315
         Left            =   600
         TabIndex        =   38
         Top             =   720
         Width           =   1695
      End
      Begin VB.Label lblCDInfo 
         Caption         =   "Choose the drive which you want to catalog."
         Height          =   495
         Left            =   120
         TabIndex        =   37
         Top             =   240
         Width           =   2655
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   1440
      TabIndex        =   41
      Top             =   4080
      Width           =   975
   End
End
Attribute VB_Name = "frmObject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'August 20, 2002

Private mFile   As FlowChart
Private mObject As FlowItem

Public Sub LoadObj(File As FlowChart, Object As FlowItem)
    Dim strPart()   As String
    Dim blnSideBar  As Boolean
    Dim objGroup    As Group

    Const conDigits = 2 'nearest hundredth
    
    On Error Resume Next
    
    Set mFile = File
    Set mObject = Object
    
    Caption = "Format " & Object.Description
    fraLocation.Caption = "Location --> "
    Select Case File.UnitScale
        Case vbCentimeters: fraLocation.Caption = fraLocation.Caption & "centimeters"
        Case vbInches: fraLocation.Caption = fraLocation.Caption & "inches"
        Case vbMillimeters: fraLocation.Caption = fraLocation.Caption & "millimeters"
        Case vbPixels: fraLocation.Caption = fraLocation.Caption & "pixels"
        Case vbCharacters: fraLocation.Caption = fraLocation.Caption & "characters"
        Case Else
            'correct units here
            File.UnitScale = vbCentimeters
            fraLocation.Caption = fraLocation.Caption & "unknown"
    End Select
    
    
    With Object.P
        'location
        txtLeft = Round(ScaleX(.Left, vbTwips, File.UnitScale), conDigits)
        txtTop = Round(ScaleY(.Top, vbTwips, File.UnitScale), conDigits)
        txtWidth = Round(ScaleX(.Width, vbTwips, File.UnitScale), conDigits)
        txtHeight = Round(ScaleY(.Height, vbTwips, File.UnitScale), conDigits)
        txtZBase = Round(ScaleX(.ZBase, vbTwips, File.UnitScale), conDigits)
        txtZHeight = Round(ScaleX(.ZHeight, vbTwips, File.UnitScale), conDigits)
        
        'information
        lblIndex(1) = File.GetIndex(Object)
        lblDescript(1) = Object.Description
        If Len(Object.DescriptionF) Then lblDescript(1) = lblDescript(1) & " / " & Object.DescriptionF
        If Object.P.GroupNo <> 0 Then
            lblIndex(0) = "Group number:"
            lblIndex(1) = Object.P.GroupNo
            lblGroup = mFile.Groups(Object.P.GroupNo).Count
            fraGroup.Visible = True
            lblGroupCount = "Group " & Object.P.GroupNo & " has " & lblGroup & " item(s)."
            blnSideBar = True
        Else
            lblGroup = "not in group"
        End If
        
        'name
        cboName = Object.P.Name
        
        'specialities
        Select Case Object.Number
        Case conAddPicture 'picture
            Dim lngR As Long
        
            fraInfo.Visible = False
            fraPic.Visible = True
            fraPic2.Visible = True
            
            blnSideBar = True
            
            txtPicture = Object.P.Text
            
            txtWidth.Visible = False 'hide text box
            lblWidth = lblWidth & "  " & txtWidth 'add to label
            lblWidth.Width = lblWidth.Width + txtWidth.Width 'enlarge label
            
            txtHeight.Visible = False
            lblHeight = lblHeight & "  " & txtHeight
            lblHeight.Width = lblHeight.Width + txtWidth.Width
            
            If GetPicture(Object).IsDefaultSize(Me) Then
                cboPicRatio = "100"
                opPicDef = True
            Else
                lngR = GetPicture(Object).EstimatedZoom(Me)
                
                cboPicRatio = lngR
                If Err Then
                    Err.Clear
                    cboPicRatio.AddItem lngR
                    cboPicRatio.ListIndex = cboPicRatio.NewIndex
                End If
                
                opPicRatio = True
            End If
            
        Case conAddButton
            fraInfo.Visible = False
            fraButton.Visible = True
            
            strPart = Split(Object.P.Tag3, " ", 2)
            If UBound(strPart) >= 0 Then cboMacro = strPart(0) Else cboMacro = "(None)"
            If UBound(strPart) = 1 Then txtMacroFile = strPart(1)
        Case conAddText
            fraInfo.Visible = False
            fraField.Visible = True
            strPart = Split(Object.P.Tag3, " ", 2)
            If UBound(strPart) = 1 Then
                If LCase$(strPart(0)) = "type" Then
                    cboField = strPart(1)
                End If
            Else
                cboField = "(None)"
            End If
        Case conAddEndArrowLine
            fraArrow.Visible = True
            blnSideBar = True
            cboArrowHead.ListIndex = Object.P.FillStyle
            cboArrowSize = Object.P.ArrowSize
            cboArrowType.ListIndex = Abs(Object.P.ArrowEngg)
        Case conAddMidArrowLine
            fraArrow.Visible = True
            blnSideBar = True
            cboArrowHead.Enabled = False
            cboArrowType.Enabled = False
            cboArrowHead.ListIndex = Object.P.FillStyle
            cboArrowSize = Object.P.ArrowSize
            cboArrowType.ListIndex = Abs(Object.P.ArrowEngg)
        End Select
        
        'See Also: cmdNameUpdate
        
        Select Case LCase$(cboName)
        Case "cdlist"
            fraCD.Visible = True
            blnSideBar = True
            drvCDList.ListIndex = -1
        End Select
    End With
    
    If blnSideBar Then
        'enlarge window size
        Width = Width - ScaleWidth + fraArrow.Left + fraArrow.Width + 10 * Screen.TwipsPerPixelX
    Else
        'shrink window size
        Width = Width - ScaleWidth + fraLocation.Left + fraLocation.Width + 10 * Screen.TwipsPerPixelX
    End If
End Sub






Private Sub SaveObj()
    Dim objPicture      As FPicture
    Dim objSet          As FlowItem

    On Error Resume Next
    
    'location - is number
    
    If Not IsNumeric(txtLeft) Then MsgBox "The Left coordinate is not a number.", vbExclamation, "Location": Exit Sub
    If Not IsNumeric(txtTop) Then MsgBox "The Top coordinate is not a number.", vbExclamation, "Location": Exit Sub
    If Not IsNumeric(txtWidth) Then MsgBox "The Width dimension is not a number.", vbExclamation, "Location": Exit Sub
    If Not IsNumeric(txtHeight) Then MsgBox "The Height dimension is not a number.", vbExclamation, "Location": Exit Sub
    If Not IsNumeric(txtZBase) Then MsgBox "The Z Base dimension is not a number.", vbExclamation, "Location": Exit Sub
    If Not IsNumeric(txtZHeight) Then MsgBox "The Z Elevation dimension is not a number.", vbExclamation, "Location": Exit Sub
    
    'location - can convert to units
    
    Set objSet = New FlowItem
    
    objSet.P.Left = ScaleX(txtLeft, mFile.UnitScale, vbTwips)
    objSet.P.Top = ScaleY(txtTop, mFile.UnitScale, vbTwips)
    objSet.P.Width = ScaleX(txtWidth, mFile.UnitScale, vbTwips)
    objSet.P.Height = ScaleY(txtHeight, mFile.UnitScale, vbTwips)
    objSet.P.ZBase = ScaleX(txtZBase, mFile.UnitScale, vbTwips)
    objSet.P.ZHeight = ScaleX(txtZHeight, mFile.UnitScale, vbTwips)
    
    
    If Err Then
        MsgBox "The location units are invalid numbers.", vbExclamation, "Location"
        Exit Sub
    End If
    
    'location - is correct sign
    If objSet.P.Left < 0 Then MsgBox "The Left coordinate cannot be a negative number.", vbExclamation, "Location": Exit Sub
    If objSet.P.Top < 0 Then MsgBox "The Top coordinate cannot be a negative number.", vbExclamation, "Location": Exit Sub
    
    If Not IsObjLine(mObject) Then
        If objSet.P.Width < 0 Then MsgBox "The Width dimension cannot be a negative number.", vbExclamation, "Location": Exit Sub
        If objSet.P.Height < 0 Then MsgBox "The Height dimension cannot be a negative number.", vbExclamation, "Location": Exit Sub
    End If
    
    'location - in bounds
    CheckObjBounds objSet, mFile.PScaleWidth, mFile.PScaleHeight
    
    If Err Then
        MsgBox "The location units are out of bounds.", vbExclamation, "Location"
        Exit Sub
    End If
    
    On Error GoTo 0
    
    'save data and undo information
    
    frmMain.ToolBoxChanged conUndoObject
    
    mFile.Changed = True
    mObject.P.Left = objSet.P.Left
    mObject.P.Top = objSet.P.Top
    mObject.P.ZBase = objSet.P.ZBase
    mObject.P.ZHeight = objSet.P.ZHeight
    If txtWidth.Visible Then mObject.P.Width = objSet.P.Width
    If txtHeight.Visible Then mObject.P.Height = objSet.P.Height
    
    cboName_Validate False
    mObject.P.Name = cboName
        
    'special cases
    If fraArrow.Visible Then
        mObject.P.FillStyle = cboArrowHead.ListIndex
        mObject.P.ArrowSize = Val(cboArrowSize)
        mObject.P.ArrowEngg = cboArrowType.ListIndex
    End If
    
    If fraButton.Visible Then
        If LCase$(cboMacro) <> "(none)" Then
            If Len(txtMacroFile) > 0 And txtMacroFile.Enabled Then
                mObject.P.Tag3 = cboMacro & " " & txtMacroFile
            Else
                mObject.P.Tag3 = cboMacro
            End If
        Else
            mObject.P.Tag3 = ""
        End If
    End If
    
    If fraCD.Visible And Len(txtCDList) > 0 Then
        mObject.P.Text = txtCDList
    End If
    
    If fraField.Visible Then
        If LCase$(cboField) <> "(none)" Then
            mObject.P.Tag3 = "Type " & cboField
        Else
            mObject.P.Tag3 = ""
        End If
        mObject.Refresh mFile, Nothing
    End If
    
    If fraPic.Visible Then 'fraPic2 should also be visible
        'pointer to the same object
        Set objPicture = mObject
        
        If mObject.P.Text <> txtPicture Then 'new/different picture
            mObject.P.Text = txtPicture
            
            objPicture.LoadPicture2 txtPicture, mFile.GetPath
            
            If opPicDef Then
                objPicture.SetDefaultSize Me 'the View is only used for scaling
            Else
                objPicture.SetZoom CLng(cboPicRatio), Me
            End If
            
        Else
            'refresh loading picture
            objPicture.LoadPicture2 txtPicture, mFile.GetPath
        
            If opPicDef Then
                objPicture.SetDefaultSize Me 'the View is only used for scaling
            Else
                objPicture.SetZoom CLng(cboPicRatio), Me
            End If
        End If
    End If
    
    If fraGroup.Visible Then
        If opGroupRemove Then
            mFile.Groups.Remove mObject 'remove this item from group
        ElseIf opUngroup Then
            mFile.Groups(mObject.P.GroupNo).RemoveAll
        End If
    End If
    
    frmMain.Redraw
    Unload Me
End Sub

Private Sub SetHomeEnd(Text As TextBox)
    Text.SelStart = 0
    Text.SelLength = Len(Text)
End Sub



Private Sub cboMacro_Click()
    txtMacroFile.Enabled = cboMacro.ItemData(cboMacro.ListIndex)
    cmdMacroFile.Enabled = cboMacro.ItemData(cboMacro.ListIndex)
End Sub


Private Sub cboPicRatio_Click()
    opPicRatio = True
End Sub


Private Sub chkCDFolder1_Click()
    drvCDList_Change
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub


Private Sub cmdMacroFile_Click()
    Dim objOpen     As FileDlg
    Dim strSetting  As String
    Dim strPathPart As String
    
    Set objOpen = New FileDlg
    Select Case LCase$(cboMacro)
    Case "open"
        objOpen.Initialize hwnd, "flc", "Flow Chart Files (*.flc)|*.flc", txtMacroFile, True, False, cdlOFNHideReadOnly
    Case "shell"
        objOpen.Initialize hwnd, "", "Executable Files (*.exe;*.bat;*.com)|*.exe;*.bat;*.com", txtMacroFile, True, False, cdlOFNHideReadOnly
    Case Else
        MsgBox "This macro does not require any file specification.", vbInformation, "Open"
        Exit Sub
    End Select
    
    Select Case objOpen.ShowOpen
    Case CDERR_CANCELLED
        'do nothing
    Case CDERR_OK
        strSetting = objOpen.FileName
        strPathPart = objOpen.RetPath
        
        If LCase$(strPathPart) = LCase$(mFile.GetPath()) Then
            strSetting = objOpen.RetFile
        End If
        
        txtMacroFile = strSetting
    Case Else
        MsgBox "Problem showing the Open dialog box.", vbExclamation, "Open"
    End Select
End Sub

Private Sub cmdNameUpdate_Click()
    Dim blnSideBar As Boolean
    
    Select Case LCase$(cboName)
    Case "cdlist"
        fraCD.Visible = True
        blnSideBar = True
        drvCDList.ListIndex = -1
        drvCDList.SetFocus
    Case Else
        fraCD.Visible = False
    End Select
    
    If blnSideBar Then
        'enlarge window size
        Width = Width - ScaleWidth + fraArrow.Left + fraArrow.Width + 10 * Screen.TwipsPerPixelX
    Else
        'shrink window size
        Width = Width - ScaleWidth + fraLocation.Left + fraLocation.Width + 10 * Screen.TwipsPerPixelX
    End If
End Sub

Private Sub cmdOK_Click()
    SaveObj
End Sub

Private Sub cmdPicMoreScales_Click()
    Dim objPic As FPicture
    
    'For pictures only
    Set objPic = mObject
    If objPic.IsLoaded() Then
        Load frmZoom
        frmZoom.Zoom = objPic.EstimatedZoom(Me)
        frmZoom.Caption = "Format Picture"
        frmZoom.Show vbModal, Me
        
        If frmZoom.Changed Then
            On Error Resume Next
            cboPicRatio = frmZoom.Zoom
            If Err Then
                Err.Clear
                cboPicRatio.AddItem frmZoom.Zoom
                cboPicRatio.ListIndex = cboPicRatio.NewIndex
            End If
        End If
        Unload frmZoom
    Else
        MsgBox "No picture has been loaded.", vbInformation
    End If
    Set objPic = Nothing
End Sub

Private Sub cmdPicture_Click()
    Dim objFile     As FileDlg
    Dim strPathPart As String
    Dim strPathFLC  As String 'of flow chart
    
    On Error Resume Next
    'Change Pictures
    Set objFile = New FileDlg
    
    objFile.Title = "Open Picture"
    objFile.Initialize hwnd, "", "Picture Files|*.bmp;*.jpg;*.gif;*.wmf;*.emf;.ico;*.cur", txtPicture, True, False, cdlOFNHideReadOnly
    
    Select Case objFile.ShowOpen
    Case CDERR_OK
        strPathPart = objFile.RetPath  'Left$(objFile.Filename, InStrRev(objFile.Filename, "\"))
        strPathFLC = mFile.GetPath()
        
        If LCase$(strPathPart) = LCase$(strPathFLC) Then
            txtPicture = objFile.RetFile
        Else
            txtPicture = objFile.FileName
        End If
    Case CDERR_CANCELLED
        'do nothing
    Case Else
        MsgBox "Problem showing Open Picture dialog box.", vbExclamation, "Open Picture"
    End Select
End Sub

Private Sub drvCDList_Change()
    Dim strFile     As String
    Dim strBuffer   As String
    Dim objFolders  As TextSort
    Dim objFiles    As TextSort
    Dim objSubFolders As TextSort
    Dim objSubFiles As TextSort
    Dim lngAttr     As Long
    Dim vntItem     As Variant
    Dim vntSubItem  As Variant
    Dim lngCount    As Long
    Dim lngSubCount As Long
    
    Const conTab = "     "
    
    If drvCDList.ListIndex = -1 Then Exit Sub
    
    MousePointer = vbHourglass
    
    Set objFolders = New TextSort
    Set objFiles = New TextSort
    
    On Error GoTo HandlerNoDisk
    'go to root directory
    ChDrive Left$(drvCDList, 1)
    ChDir Left$(drvCDList, 1) & ":\"
    
    On Error GoTo HandlerRead
    strFile = Dir(Left$(drvCDList, 1) & ":\*.*", vbDirectory Or vbHidden Or vbReadOnly Or vbSystem Or vbArchive)
    
    Do While Len(strFile) > 0
        If strFile <> "." And strFile <> ".." And LCase$(strFile) <> "pagefile.sys" Then
            lngAttr = GetAttr(strFile)
            If UCase$(strFile) = strFile Then 'uppercase
                'convert to Proper Case
                strFile = StrConv(strFile, vbProperCase)
            End If
            If (lngAttr And vbDirectory) = vbDirectory Then
                objFolders.Insert strFile
            Else
                objFiles.Insert strFile
            End If
        End If
        strFile = Dir
    Loop
    
    For Each vntItem In objFolders
        lngCount = lngCount + 1
        strBuffer = strBuffer & lngCount & ". " & vntItem & vbNewLine
        
        'read 1st folder level too
        If chkCDFolder1 = vbChecked Then
            strFile = Dir(Left$(drvCDList, 1) & ":\" & vntItem & "\*.*", vbNormal Or vbReadOnly Or vbHidden Or vbSystem Or vbDirectory)
            
            Set objSubFolders = New TextSort
            Set objSubFiles = New TextSort
            lngSubCount = 0
            
            ChDir Left$(drvCDList, 1) & ":\" & vntItem
            
            Do While Len(strFile) > 0
                If strFile <> "." And strFile <> ".." Then
                    lngAttr = GetAttr(strFile)
                    
                    If UCase$(strFile) = strFile Then 'uppercase
                        'convert to Proper Case
                        strFile = StrConv(strFile, vbProperCase)
                    End If
                    
                    If (lngAttr And vbDirectory) = vbDirectory Then
                        objSubFolders.Insert strFile
                    Else
                        objSubFiles.Insert strFile
                    End If
                End If
                strFile = Dir
            Loop
            
            For Each vntSubItem In objSubFolders
                lngSubCount = lngSubCount + 1
                strBuffer = strBuffer & conTab & lngCount & "." & lngSubCount & ".  " & vntSubItem & vbNewLine
            Next vntSubItem
        
            For Each vntSubItem In objSubFiles
                lngSubCount = lngSubCount + 1
                strBuffer = strBuffer & conTab & lngCount & "." & lngSubCount & ")  " & vntSubItem & vbNewLine
            Next vntSubItem
        End If
    Next vntItem
    
    If objFiles.Count > 0 Then
        'add in a divider (hyphen so far)
        If Len(strBuffer) > 0 Then strBuffer = strBuffer & "-" & vbNewLine
    
        'add in list of files
        For Each vntItem In objFiles
            lngCount = lngCount + 1
            strBuffer = strBuffer & lngCount & ") " & vntItem & vbNewLine
        Next vntItem
    End If
    
    If Len(strBuffer) <= 32767 Then
        txtCDList = strBuffer
        lblCDInfo = "Choose the drive which you want to catalog."
        cmdOK.Enabled = True
    Else
        lblCDInfo = "The drive listing is too large, and is truncated at 32 KB."
        txtCDList = Left$(strBuffer, 32767)
        cmdOK.Enabled = False
    End If
    
    MousePointer = vbDefault
    
    Exit Sub
HandlerNoDisk:
    txtCDList = ""
    lblCDInfo = "There is no disk in this drive."
    cmdOK.Enabled = False
    MousePointer = vbDefault
    Exit Sub
HandlerRead:
    lblCDInfo = "Problem reading this disk's contents.  " & Err.Description
    
    MsgBox lblCDInfo, vbExclamation, "CD List"
    
    If Len(strBuffer) <= 32767 Then
        txtCDList = strBuffer
        cmdOK.Enabled = True
    Else
        txtCDList = Left$(strBuffer, 32767) 'truncation
        cmdOK.Enabled = False
    End If
    
    MousePointer = vbDefault
    Exit Sub
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set mFile = Nothing
    Set mObject = Nothing
End Sub


Private Sub txtHeight_GotFocus()
    SetHomeEnd txtHeight
End Sub

Private Sub txtLeft_GotFocus()
    SetHomeEnd txtLeft
End Sub

Private Sub cboName_Validate(Cancel As Boolean)
    Select Case LCase$(cboName)
    Case "cdlist": cboName = "CDList"
    End Select
End Sub


Private Sub txtTop_GotFocus()
    SetHomeEnd txtTop
End Sub

Private Sub txtWidth_GotFocus()
    SetHomeEnd txtWidth
End Sub


