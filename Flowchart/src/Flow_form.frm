VERSION 5.00
Begin VB.Form frmForm 
   BackColor       =   &H8000000C&
   Caption         =   "Form"
   ClientHeight    =   4950
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   5565
   Icon            =   "Flow_form.frx":0000
   LockControls    =   -1  'True
   ScaleHeight     =   4950
   ScaleWidth      =   5565
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picPreview 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ClipControls    =   0   'False
      ForeColor       =   &H00000000&
      Height          =   4935
      Left            =   0
      MouseIcon       =   "Flow_form.frx":0442
      ScaleHeight     =   4905
      ScaleWidth      =   5505
      TabIndex        =   0
      Top             =   0
      Width           =   5535
      Begin VB.Timer timLayer 
         Enabled         =   0   'False
         Interval        =   2000
         Left            =   2040
         Top             =   480
      End
      Begin VB.Label lblCaption 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Layer"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   27.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   855
         Left            =   0
         TabIndex        =   1
         Top             =   0
         Visible         =   0   'False
         Width           =   3735
      End
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileOpen 
         Caption         =   "&Open"
         Shortcut        =   ^O
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFileBold 
         Caption         =   "&Bold"
         Shortcut        =   ^B
      End
      Begin VB.Menu mnuFileSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileAbout 
         Caption         =   "&About..."
      End
      Begin VB.Menu mnuFilePrint 
         Caption         =   "&Print..."
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFilePrintSetup 
         Caption         =   "Print Set&up..."
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFileSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileClose 
         Caption         =   "&Close"
      End
   End
   Begin VB.Menu mnuZoom 
      Caption         =   "&Zoom"
      Begin VB.Menu mnuZoom25 
         Caption         =   "25%"
         Shortcut        =   ^A
      End
      Begin VB.Menu mnuZoom50 
         Caption         =   "50%"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuZoom75 
         Caption         =   "75%"
         Shortcut        =   ^D
      End
      Begin VB.Menu mnuZoom100 
         Caption         =   "100%"
         Shortcut        =   ^F
      End
      Begin VB.Menu mnuZoom150 
         Caption         =   "150%"
         Shortcut        =   ^G
      End
      Begin VB.Menu mnuZoom200 
         Caption         =   "200%"
         Shortcut        =   ^H
      End
   End
   Begin VB.Menu mnuLayer 
      Caption         =   "&Layer"
      Begin VB.Menu mnuLayerAll 
         Caption         =   "All Layers"
      End
      Begin VB.Menu mnuLayerItem 
         Caption         =   "0"
         Index           =   0
      End
   End
End
Attribute VB_Name = "frmForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public mFlowChart  As FlowChart
'for interfaces
Private mMousePos  As Rect
Private mblnDown   As Boolean
Private mWindow    As Rect

#If Prev Then
    #If active Then
        Private Const conTitle = "Flow Chart - Viewer [Unified Vector Space release]"
    #Else
        Private Const conTitle = "Flow Chart - Viewer"
    #End If
#End If

#If Prev Then
Public Sub AskOpen(Optional ByVal FileName As String)
    Dim lngErr As Long
    Dim objOpen As FileDlg
    
    On Error Resume Next
    
    
    Set objOpen = New FileDlg
    If Len(FileName) Then
        objOpen.FileName = FileName
        GoSub OpenFile
    Else
        
        objOpen.Initialize hwnd, conDefaultExt, conDefaultFilter, "", True, False, &H4& 'cdlOFNHideReadOnly
        Select Case objOpen.ShowOpen
        Case CDERR_CANCELLED
            'nothing
        Case CDERR_OK
            GoSub OpenFile
        Case Else
            MsgBox "Unable to show Open dialog box.", vbExclamation
        End Select
    End If
    
    Screen.MousePointer = vbDefault
    'With dlgFile
'        .Flags = cdlOFNFileMustExist Or cdlOFNHideReadOnly
'        .Filter = conDefaultFilter
'        .FilterIndex = 1
'        .ShowOpen
'        If Err = cdlCancel Then 'cancelled
'            Exit Sub
'        ElseIf Err <> 0 Then 'error
'            MsgBox "Unable to show Open dialog box.  (" & Err & ")", vbExclamation
'            .FileName = ""
'            .InitDir = ""
'        Else 'good
'            Set mFlowChart = New FlowChart
'            lngErr = mFlowChart.Load(.FileName, picPreview)
'            Select Case lngErr
'            Case conFail
'                MsgBox "Problem opening the flow chart.  The flow chart may be corrupted or the file is too new for this version of the program.", vbExclamation
'                mFlowChart.FileName = "" 'erase filename
'            Case 0
'                'no problem
'            Case Else
'                MsgBox "Failed opening file because of the following problem: " & Error(lngErr), vbExclamation
'                mFlowChart.FileName = "" 'erase filename
'            End Select
'            SetZoom mFlowChart.ZoomPercent
'        End If
'    End With

Exit Sub
OpenFile:
        Screen.MousePointer = vbHourglass
        Set mFlowChart = New FlowChart
        
        lngErr = mFlowChart.Load(objOpen.FileName, picPreview)
        Select Case lngErr
        Case conFail
            MsgBox "Problem opening the flow chart.  The flow chart may be corrupted or the file is too new for this version of the program.", vbExclamation
            mFlowChart.FileName = "" 'erase filename
        Case 0
            'no problem
        Case Else
            MsgBox "Failed opening file because of the following problem: " & Error(lngErr), vbExclamation
            mFlowChart.FileName = "" 'erase filename
        End Select
        
        mFlowChart.DefaultLayer = -1 'all layers shown
        
        SetLayers
        SetZoom mFlowChart.ZoomPercent
    Return
End Sub
#End If

Private Function CalcView() As Rect
    CalcView.X1 = picPreview.ScaleX(-picPreview.Left, ScaleMode, picPreview.ScaleMode)
    CalcView.Y1 = picPreview.ScaleY(-picPreview.Top, ScaleMode, picPreview.ScaleMode)
    CalcView.X2 = CalcView.X1 + picPreview.ScaleX(ScaleWidth, ScaleMode, picPreview.ScaleMode)
    CalcView.Y2 = CalcView.Y1 + picPreview.ScaleY(ScaleHeight, ScaleMode, picPreview.ScaleMode)
End Function


#If active Then
Public Sub EnableTimer(ByVal Caption As String)
#Else
Private Sub EnableTimer(ByVal Caption As String)
#End If
    lblCaption = Caption
    lblCaption.Move -picPreview.Left, -picPreview.Top
    lblCaption.Visible = True
    timLayer.Enabled = False
    timLayer.Enabled = True
End Sub

Private Function Hit(ByVal X As Single, ByVal Y As Single) As Boolean
    Dim objItem As FlowItem
    
    'only Hit on a button
    For Each objItem In mFlowChart
        If X >= objItem.P.Left And Y >= objItem.P.Top And _
           X <= objItem.P.Left + objItem.P.Width And Y <= objItem.P.Top + objItem.P.Height _
           And objItem.Number = conAddButton Then
            Hit = True
            mFlowChart.ForceBold = mnuFileBold.Checked
            GetButton(objItem).Click picPreview, mFlowChart
            mFlowChart.ForceBold = False
            picPreview.Refresh
            Exit Function
        End If
    Next objItem
End Function

Public Property Let MenusVisible(ByVal pVis As Boolean)
    mnuFile.Visible = pVis
    mnuLayer.Visible = pVis
    mnuZoom.Visible = pVis
End Property

Private Sub SetLayers()
    Dim intIndex As Integer
    
    'remove old menu items
    For intIndex = mnuLayerItem.UBound To 1 Step -1
        Unload mnuLayerItem(intIndex)
    Next intIndex
    
    'load new ones
    With mFlowChart.Layers
        .Requery
        For intIndex = 1 To 255
            If .Layer(intIndex).Count > 0 Then
                Load mnuLayerItem(intIndex)
                mnuLayerItem(intIndex).Caption = intIndex
                mnuLayerItem(intIndex).Visible = True
            End If
        Next intIndex
    End With
End Sub

Public Sub SetZoom(ByVal Percent As Integer)
    Dim sngScale As Single
    
    sngScale = Percent / 100
    mFlowChart.ZoomPercent = Percent
'    Debug.Print "Before", picPreview.ScaleWidth, picPreview.ScaleHeight
    If mFlowChart.PrinterError Then
        picPreview.Move 0, 0, mFlowChart.PScaleWidth * sngScale, mFlowChart.PScaleHeight * sngScale
        picPreview.Scale (0, 0)-(mFlowChart.PScaleWidth, mFlowChart.PScaleHeight)
    Else
        picPreview.Move 0, 0, Printer.ScaleWidth * sngScale, Printer.ScaleHeight * sngScale
        picPreview.Scale (0, 0)-(Printer.ScaleWidth, Printer.ScaleHeight)
    End If
'    Debug.Print "After", picPreview.ScaleWidth, picPreview.ScaleHeight
    If Not mFlowChart.CustomBack Then
        picPreview.BackColor = vbWhite
    End If
    Form_Resize
    picPreview.Cls
    picPreview_Paint
End Sub


#If Prev Then
Private Sub Form_Load()
    Caption = conTitle
    mnuFilePrint.Visible = True
    mnuFilePrintSetup.Visible = True
    mnuFileOpen.Visible = True
    mnuFileClose.Caption = "E&xit"
End Sub
#End If

#If active Then
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode <> vbFormCode Then
        If g_lngAutomation > 0 Then
            MsgBox "This program is being used by another application at this time.", vbInformation
            Cancel = 1
        End If
    End If
End Sub
#End If

Private Sub Form_Resize()
    If WindowState <> vbMinimized Then
        mWindow.X1 = -picPreview.Width + ScaleWidth 'Left
        mWindow.Y1 = -picPreview.Height + ScaleHeight 'Top
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set mFlowChart = Nothing
End Sub


Private Sub mnuFileAbout_Click()
    Load frmAbout
    frmAbout.Show vbModal, Me
End Sub

Private Sub mnuFileBold_Click()
    mnuFileBold.Checked = Not mnuFileBold.Checked
    picPreview.Refresh
End Sub

Private Sub mnuFileClose_Click()
    #If active Then
    If g_lngAutomation > 0 Then
        MsgBox "This program is being used by another application at this time.", vbInformation
        Exit Sub
    End If
    #End If
    Unload Me
End Sub

Private Sub mnuFileOpen_Click()
    #If Prev = 0 Then
    Hide
    Set frmMain = New frmIMain
    Load frmMain
    frmMain.MainShow
    If Len(mFlowChart.FileName) Then
        frmMain.OpenFile mFlowChart.FileName
    End If
    Unload Me
    #Else
        AskOpen
    #End If
End Sub

#If Prev Then
Private Sub mnuFilePrint_Click()
    Dim lngLower As Long, lngUpper As Long
    
    On Error Resume Next
    Printer.KillDoc 'release printer object handle
    Load frmPrint
    
    mFlowChart.Layers.Requery
    mFlowChart.Layers.MinMaxLayer lngLower, lngUpper
    
    If frmPrint.Initialize(conPrintAll Or conPrintRange Or conPrintSelection, conPrintSelection, lngLower, lngUpper, lngLower, lngUpper) Then
        If frmPrint.ShowModal(Me) Then
            'restore previous orientation because that can ONLY be changed through Print Setup, not Print
            mFlowChart.Header1PDevName = Printer.DeviceName 'user okay this printer
            
            mFlowChart.PrintFile frmPrint.PrintRange, frmPrint.FromPage, frmPrint.ToPage, frmPrint.PrintOnSamePage
            mFlowChart.Changed = True
            picPreview.Refresh
        End If
    Else
        MsgBox "No printers are installed.", vbExclamation
    End If
    
    'removed, layers does not change in preview window
    'mFlowChart.Layers.CloseQuery
    Unload frmPrint
    Set frmPrint = Nothing

End Sub

Private Sub mnuFilePrintSetup_Click()
    On Error Resume Next
    'Do I want to remove this and
    'use my own window?

    Printer.KillDoc 'release printer object handle
    Load frmPrint
    If frmPrint.InitializeSetup(Printer.Orientation) Then
        If frmPrint.ShowModal(Me) Then
            SetZoom mFlowChart.ZoomPercent
        End If
    Else
        MsgBox "No printers are installed.", vbExclamation
    End If
    Unload frmPrint
    Set frmPrint = Nothing
End Sub
#End If


Private Sub mnuLayerAll_Click()
    mFlowChart.Layers.EnableAll = True
    mFlowChart.DefaultLayer = -1
    picPreview.Refresh
    
    EnableTimer "All Layers"
End Sub

Private Sub mnuLayerItem_Click(Index As Integer)
    mFlowChart.Layers.EnableAll = False
    mFlowChart.DefaultLayer = Index
    mFlowChart.Layers(Index).Enabled = True
    picPreview.Refresh
    
    EnableTimer "Layer " & Index
End Sub

Private Sub mnuZoom100_Click()
    SetZoom 100
End Sub

Private Sub mnuZoom150_Click()
    SetZoom 150
End Sub


Private Sub mnuZoom200_Click()
    SetZoom 200
End Sub


Private Sub mnuZoom25_Click()
    SetZoom 25
End Sub

Private Sub mnuZoom50_Click()
    SetZoom 50
End Sub


Private Sub mnuZoom75_Click()
    SetZoom 75
End Sub

Private Sub picPreview_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim lngLoop As Long
    
    If Shift = vbAltMask Then
        Select Case KeyCode
        Case vbKeyDown, vbKeyRight
            mFlowChart.Layers.EnableAll = False
            mFlowChart.DefaultLayer = mFlowChart.DefaultLayer + 1
            
            If mFlowChart.DefaultLayer + 1 > 255 Then mFlowChart.DefaultLayer = 255
            
            For lngLoop = mFlowChart.DefaultLayer To 255
                If mFlowChart.Layers(lngLoop).Count > 0 Then
                    mFlowChart.Layers(lngLoop).Enabled = True
                    mFlowChart.DefaultLayer = lngLoop
                    Exit For
                End If
            Next lngLoop
            
            picPreview.Cls
            picPreview.Refresh
            
            EnableTimer "Layer " & mFlowChart.DefaultLayer
        Case vbKeyUp, vbKeyLeft
            mFlowChart.Layers.EnableAll = False
            mFlowChart.DefaultLayer = mFlowChart.DefaultLayer - 1
            
            If mFlowChart.DefaultLayer <= -1 Then
                mFlowChart.DefaultLayer = -1
                mFlowChart.Layers.EnableAll = True
                picPreview.Cls
                picPreview.Refresh
                
                EnableTimer "All Layers"
            Else
                For lngLoop = mFlowChart.DefaultLayer To 0 Step -1
                    If mFlowChart.Layers(lngLoop).Count > 0 Then
                        mFlowChart.Layers(lngLoop).Enabled = True
                        mFlowChart.DefaultLayer = lngLoop
                        Exit For
                    End If
                Next lngLoop
                
                picPreview.Cls
                picPreview.Refresh
                
                EnableTimer "Layer " & mFlowChart.DefaultLayer
            End If
        End Select
    Else
        Select Case KeyCode
        Case vbKeyDown
            If picPreview.Top - ScaleHeight / 2 >= mWindow.Y1 Then
                picPreview.Top = picPreview.Top - ScaleHeight / 2
            Else
                picPreview.Top = mWindow.Y1
            End If
        Case vbKeyUp
            If picPreview.Top + ScaleHeight / 2 <= 0 Then
                picPreview.Top = picPreview.Top + ScaleHeight / 2
            Else
                picPreview.Top = 0
            End If
        Case vbKeyLeft
            If picPreview.Left + ScaleWidth / 2 <= 0 Then
                picPreview.Left = picPreview.Left + ScaleWidth / 2
            Else
                picPreview.Left = 0
            End If
        Case vbKeyRight
            If picPreview.Left - ScaleWidth / 2 >= mWindow.X1 Then
                picPreview.Left = picPreview.Left - ScaleWidth / 2
            Else
                picPreview.Left = mWindow.X1
            End If
        Case vbKeyF5
            picPreview.Cls
            mFlowChart.DrawFile picPreview, mnuFileBold.Checked, True, 0, 0, 0, 0
        End Select
    End If
End Sub


Private Sub picPreview_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbLeftButton Then
        If Not Hit(X, Y) Then
            mMousePos.X1 = picPreview.ScaleX(X, picPreview.ScaleMode, ScaleMode)
            mMousePos.Y1 = picPreview.ScaleY(Y, picPreview.ScaleMode, ScaleMode)
'            mMousePos.x2 = X
'            mMousePos.y2 = Y
            picPreview.MousePointer = vbCustom
            mblnDown = True
        Else
            mblnDown = False
        End If
    End If
End Sub

Private Sub picPreview_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If mblnDown Then
        mMousePos.X2 = picPreview.Left + picPreview.ScaleX(X, picPreview.ScaleMode, ScaleMode) - mMousePos.X1
        mMousePos.Y2 = picPreview.Top + picPreview.ScaleY(Y, picPreview.ScaleMode, ScaleMode) - mMousePos.Y1
        If mMousePos.X2 > 0 Then
            mMousePos.X2 = 0
        ElseIf Not (mMousePos.X2 + picPreview.Width >= ScaleWidth) Then
            If picPreview.Width > ScaleWidth Then
                mMousePos.X2 = ScaleWidth - picPreview.Width
            Else
                mMousePos.X2 = 0
            End If
        End If
        If mMousePos.Y2 > 0 Then
            mMousePos.Y2 = 0 'picPreview.Top
        ElseIf Not (mMousePos.Y2 + picPreview.Height >= ScaleHeight) Then
            If picPreview.Height > ScaleHeight Then
                mMousePos.Y2 = ScaleHeight - picPreview.Height
            Else
                mMousePos.Y2 = 0
            End If
        End If
        picPreview.Move mMousePos.X2, mMousePos.Y2
    End If
End Sub


Private Sub picPreview_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If mblnDown Then
        picPreview_MouseMove Button, Shift, X, Y
        On Error Resume Next
        'hsbBar.Value = ScaleX(-mMousePos.x1, vbTwips, vbPixels)
        'vsbBar.Value = ScaleY(-mMousePos.y1, vbTwips, vbPixels)
        picPreview.MousePointer = vbDefault
        mblnDown = False
    End If
End Sub

Private Sub picPreview_Paint()
    Dim tView As Rect

    tView = CalcView()
    mFlowChart.DrawFile picPreview, mnuFileBold.Checked, False, tView.X1, tView.Y1, tView.X2, tView.Y2
End Sub


Private Sub timLayer_Timer()
    timLayer.Enabled = False
    lblCaption.Visible = False
End Sub


