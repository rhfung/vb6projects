VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FPicture"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit 'Using Properties
'Richard Fung.  February 24, 2001.

'LoadPicture() does not Resize.
'Refresh()      loads and resizes.
'SetDefaultSize() also resizes.

#Const ForceLargePictures = 0

Implements FlowItem

Private mstrFilename    As String
Private mstrPath        As String
Private mobjProp        As Properties
Private mobjPicture     As StdPicture
Private mstrError       As String
Private mblnScaleError  As Boolean
Private mblnChecked     As Boolean

Private Const conDefSize = "default-size"


Public Function EstimatedZoom(View As Object) As Long
    EstimatedZoom = (mobjProp.Width / GetDefWidth(View) + _
                mobjProp.Height / GetDefHeight(View)) * 50 '/2 halves * 100 per cent
End Function

Public Function FullFilename(Optional ByVal File As FlowChart) As String
    If Not File Is Nothing Then mstrPath = File.GetPath
    
    If Len(mstrPath) Then
        FullFilename = IIf(Right$(mstrPath, 1) = "\", mstrPath & mstrFilename, mstrPath & "\" & mstrFilename)
    Else
        FullFilename = mstrFilename
    End If
End Function

Public Function GetDefWidth(View As Object)
    GetDefWidth = View.ScaleX(mobjPicture.Width, vbHimetric, vbTwips)
End Function

Public Function GetDefHeight(View As Object)
    GetDefHeight = View.ScaleY(mobjPicture.Height, vbHimetric, vbTwips)
End Function

Public Function IsDefaultSize(View As Object) As Boolean
'Checks the file's width and height and compares that
'to the stored width and height.  Sometimes these values
'will not match because user is using high resolution
'screen mode with large text.
    If Round(GetDefWidth(View)) = Round(mobjProp.Width) And _
    Round(GetDefHeight(View)) = Round(mobjProp.Height) Then
        IsDefaultSize = True
        mblnScaleError = False
    Else
        IsDefaultSize = False
        mblnScaleError = True
    End If
    mblnChecked = True
End Function


Public Function IsLoaded() As Boolean
    IsLoaded = Not mobjPicture Is Nothing
End Function


Public Sub LoadPicture()
    LoadPicture2 mobjProp.Text, ""
End Sub

Private Property Let mblnDefSize(ByVal pVal As Boolean)
    mobjProp.Tag1 = IIf(pVal, conDefSize, "")
End Property

Private Property Get mblnDefSize() As Boolean
    mblnDefSize = (mobjProp.Tag1 = conDefSize)
End Property


Public Sub SetDefaultSize(ByVal View As Object)
    If Not mobjPicture Is Nothing Then
        mobjProp.Width = View.ScaleX(mobjPicture.Width, vbHimetric, vbTwips)
        mobjProp.Height = View.ScaleY(mobjPicture.Height, vbHimetric, vbTwips)
    End If
    mblnDefSize = True
End Sub


Public Sub SetCustomSize()
    mblnDefSize = False
End Sub

Public Sub SetZoom(ByVal ZoomRatio As Long, ByVal View As Object)
    SetCustomSize

    mobjProp.Width = GetDefWidth(View) * ZoomRatio / 100
    mobjProp.Height = GetDefHeight(View) * ZoomRatio / 100
End Sub


Private Sub Class_Initialize()
    Set mobjProp = New Properties
End Sub

Private Sub Class_Terminate()
    Set mobjPicture = Nothing
    Set mobjProp = Nothing
End Sub


Private Property Get FlowItem_CanRefresh() As Boolean
    FlowItem_CanRefresh = True
End Property

Private Property Get FlowItem_Description() As String
    FlowItem_Description = "Picture"
End Property

Private Property Get FlowItem_DescriptionF() As String
    'nothing to do
End Property


'Private Property Get FlowItem_CenterX() As Single
'    FlowItem_CenterX = mobjProp.CenterX
'End Property
''
'Private Property Get FlowItem_CenterY() As Single
'    FlowItem_CenterY = mobjProp.CenterY
'End Property


Private Function FlowItem_Draw(ByVal View As Object, ByVal Parent As FlowChart) As Long
'    If mstrFilename <> mobjProp.Text Then
'        LoadPicture
'    End If
    If IsLoaded() Then
        If Not mblnChecked Then IsDefaultSize View 'update, IsDefaultSize() will report mblnScaleError
        If mblnDefSize And Not mblnScaleError And Parent.ZoomPercent = 100 Then
            View.PaintPicture mobjPicture, mobjProp.Left, mobjProp.Top
        Else
            View.PaintPicture mobjPicture, mobjProp.Left, mobjProp.Top, mobjProp.Width, mobjProp.Height
        End If
    Else
        SetDrawProps View, Me, Parent
        SetFontProps View, Me, Parent
        
        View.Line (mobjProp.Left, mobjProp.Top)-Step(mobjProp.Width, mobjProp.Height), , B
        View.Line (mobjProp.Left, mobjProp.Top)-Step(mobjProp.Width, mobjProp.Height)
        View.Line (mobjProp.Left, mobjProp.Top + mobjProp.Height)-(mobjProp.Left + mobjProp.Width, mobjProp.Top)
        View.CurrentX = mobjProp.Left + conIndent
        View.CurrentY = mobjProp.Top + conIndent
        View.Print "#" & mstrError
    End If
End Function
'
'Private Property Let FlowItem_DrawOrder(ByVal RHS As FOrder)
'    mobjFlowItem.DrawOrder = RHS
'End Property
'
'Private Property Get FlowItem_DrawOrder() As FOrder
'    FlowItem_DrawOrder = mobjFlowItem.DrawOrder
'End Property
'
'Private Property Let FlowItem_Height(ByVal RHS As Single)
'    mobjProp.Height = RHS
'End Property
'
'Private Property Get FlowItem_Height() As Single
'    FlowItem_Height = mobjProp.Height
'End Property
'
'Private Property Let FlowItem_Left$(ByVal RHS As Single)
'    mobjProp.Left = RHS
'End Property
'
'Private Property Get FlowItem_Left$() As Single
'    FlowItem_Left = mobjProp.Left
'End Property

Private Property Get FlowItem_Number() As FAddType
    FlowItem_Number = conAddPicture
End Property

Private Property Set FlowItem_P(ByVal RHS As Properties)
    Set mobjProp = RHS
End Property

Private Property Get FlowItem_P() As Properties
    Set FlowItem_P = mobjProp
End Property

Public Sub LoadPicture2(FileName As String, Path As String)
    On Error Resume Next
    Dim lngFileSize As Long

    mstrFilename = FileName
    mstrPath = Path
     
    If Len(Path) Then 'change to correct path
       
        ChDrive Path
        ChDir Path
    End If
        
    If Len(mstrFilename) And Err = 0 Then
        lngFileSize = FileLen(mstrFilename) 'vbLP... for constants
        If Err Then
            mstrError = "Picture missing"
            Err.Clear
            Set mobjPicture = Nothing
        Else
            #If ForceLargePictures = 0 Then
            If lngFileSize <= 1048576 Then  '1 MB
            #End If
            Set mobjPicture = VB.LoadPicture(mstrFilename)
            If Err Then
                Err.Clear
                Set mobjPicture = Nothing
                mstrError = "Unable to load picture"
            End If
            #If ForceLargePictures = 0 Then
            Else 'too large
                Set mobjPicture = Nothing
                mstrError = "Picture size is over 1 MB - Too large"
            End If
            #End If
        End If
    Else
        mstrError = "No picture"
        Set mobjPicture = Nothing
    End If
    mblnChecked = False
    mblnScaleError = False
    
    If mobjProp.Text <> FileName Then
        mobjProp.Text = FileName 'save filename
    End If
    
    Err.Clear
End Sub

Private Sub FlowItem_Refresh(ByVal Parent As FlowChart, ByVal View As Object)
    If Not Parent Is Nothing Then
        'loads picture
        LoadPicture
    Else
        LoadPicture2 mobjProp.Text, Parent.GetPath
    End If
'    If Not View Is Nothing And Not mobjPicture Is Nothing And mblnDefSize Then
'        mobjProp.Width = GetDefWidth(View)
'        mobjProp.Height = GetDefHeight(View)
'    End If
End Sub


Private Property Get FlowItem_TextBottomMrg() As Single
    FlowItem_TextBottomMrg = 0
End Property

'
'Private Property Get FlowItem.P.TextColour() As Long
'    FlowItem.P.TextColour = mobjFlowItem.TextColour
'End Property
'
'Private Property Let FlowItem.P.TextColour(ByVal RHS As Long)
'    mobjFlowItem.TextColour = RHS
'End Property

'Private Property Let FlowItem_TextLeftMrg(ByVal RHS As Single)
'    mobjFlowItem.TextLeftMrg = RHS
'End Property

Private Property Get FlowItem_TextLeftMrg() As Single
    FlowItem_TextLeftMrg = 0
End Property


'Private Property Let FlowItem_TextRightMrg(ByVal RHS As Single)
'    mobjFlowItem.TextRightMrg = RHS
'End Property

Private Property Get FlowItem_TextRightMrg() As Single
    FlowItem_TextRightMrg = 0
End Property


'Private Property Let FlowItem_TextTopMrg(ByVal RHS As Single)
'    mobjFlowItem.TextTopMrg = RHS
'End Property

Private Property Get FlowItem_TextTopMrg() As Single
    FlowItem_TextTopMrg = 0
End Property

