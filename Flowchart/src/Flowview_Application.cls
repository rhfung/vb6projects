VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FlowchartApplication"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements IApplication

Private m_Drawing As Drawing

Private Sub Class_Initialize()
    If Not gblnMainStarted Then CPlusPlus
    g_lngAutomation = g_lngAutomation + 1
End Sub

Private Sub Class_Terminate()
    g_lngAutomation = g_lngAutomation - 1
    If g_lngAutomation = 0 And frmForm.Visible = False Then
        'close transactions and quit
        Unload frmForm
        Set frmForm = Nothing
        
        Unload g_frmSplash
    End If
End Sub


Private Property Get IApplication_Application() As Object
    Err.Raise 91, , "This application does not allow third parties to use its components."
End Property

Private Property Get IApplication_Application2() As UniVector.IWindow
    Err.Raise 91, , "This application does not allow third parties to use its components."
End Property

Private Property Get IApplication_AppName() As String
    IApplication_AppName = "Flow Chart with Unified Vector Space"
End Property


Private Property Set IApplication_Drawing(RHS As UniVector.Drawing)
    Set m_Drawing = RHS
End Property

Private Property Get IApplication_Drawing() As UniVector.Drawing
    Set IApplication_Drawing = m_Drawing
End Property

Private Sub IApplication_Load(ByVal FileName As String)
    frmForm.Show
    frmForm.EnableTimer "Load..."
    frmForm.AskOpen FileName
    
    IApplication_LoadActiveWindow
End Sub


Private Sub IApplication_LoadActiveWindow()
    Dim objFile As FlowChart
    Dim objItem As FlowItem
    Dim lngC    As Long
    Dim objStart As Uni3DPoint
    Dim objEnd   As Uni3DPoint
    Dim sngHeight As Single
    Const conError = "Calling application must set Drawing object before Load or LoadActiveWindow because ActiveX calls from out-of-process component is much slower.  Drawing should be kept in the memory space of the calling object."
    
    Set g_frmSplash.mFlowChart = g_Flowchart
    Load g_frmSplash
    ConditionSplash g_frmSplash
    g_frmSplash.Show vbModeless, frmForm
    
    frmForm.Show
    
    Set objFile = frmForm.mFlowChart
    If m_Drawing Is Nothing Then
        Err.Raise 91, , conError
        Exit Sub
    ElseIf m_Drawing.Elements.Count <> 0 Then
        Err.Raise 91, , conError
        Exit Sub
    End If
    
    'Set m_Drawing = New Drawing
    
    frmForm.EnableTimer "Sending"
    
    If objFile Is Nothing Then
        Err.Raise 91, , "There is no window loaded."
    End If
    
    sngHeight = frmForm.ScaleY(objFile.PScaleHeight, vbTwips, objFile.UnitScale)
    
    m_Drawing.Properties.Add "ScaleMode", objFile.UnitScale
    m_Drawing.Properties.Add "FlowchartVersion", objFile.Version
    m_Drawing.Properties.Add "ZoomPercent", objFile.ZoomPercent
    m_Drawing.Properties.Add "PageHeight", CStr(sngHeight)
    m_Drawing.Properties.Add "PageWidth", frmForm.ScaleX(objFile.PScaleWidth, vbTwips, objFile.UnitScale)
   
    
    For Each objItem In objFile
        lngC = lngC + 1
        
        If IsObjLine(objItem) Then
            Set objStart = GetUniPoint(frmForm.ScaleX(objItem.P.Left, vbTwips, objFile.UnitScale), -frmForm.ScaleY(objItem.P.Top, vbTwips, objFile.UnitScale) + sngHeight, frmForm.ScaleX(objItem.P.ZBase, vbTwips, objFile.UnitScale))
            Set objEnd = GetUniPoint(frmForm.ScaleX(objItem.P.Left + objItem.P.Width, vbTwips, objFile.UnitScale), -frmForm.ScaleY(objItem.P.Top + objItem.P.Height, vbTwips, objFile.UnitScale) + sngHeight, frmForm.ScaleX(objItem.P.ZBase + objItem.P.ZHeight, vbTwips, objFile.UnitScale))
        
            With m_Drawing.Elements.Add(objItem.P.Name, lngC, UniGetUnitVector(objStart, objEnd).X, objStart, objEnd)
                .Tags.Add "ForeColour", objItem.P.ForeColour
                .Tags.Add "Text", objItem.P.Text
                .Tags.Add "FlowItemNumber", objItem.Number
                .Tags.Add "LineWidth", objItem.P.LineWidth
            End With
        End If
    Next objItem
    
    Unload g_frmSplash
End Sub



Private Property Get IApplication_SatelliteDeveloper() As String
    IApplication_SatelliteDeveloper = "Richard Fung"
End Property


Private Sub IApplication_Save(ByVal FileName As String)
    Dim objFile As FlowChart
    Dim objE    As Element
    Dim objNew  As FlowItem
    Dim lngP    As Long
    Dim intS    As Integer
    Dim sngHeight As Single
    Dim lngType As Long
    Dim lngColour As Long
    Dim lngWidth   As Long
    Dim blnNativeFile As Boolean
    
    Set g_frmSplash.mFlowChart = g_Flowchart
    Load g_frmSplash
    ConditionSplash g_frmSplash
    g_frmSplash.Show vbModeless, frmForm
    
    Set objFile = New FlowChart
    Set frmForm.mFlowChart = objFile
    
    frmForm.EnableTimer "Receiving"
    
    objFile.FileName = FileName
    
    On Error Resume Next
    objFile.Version = m_Drawing.Properties("FlowchartVersion")
    If Err Then
        Err.Clear
        objFile.Version = conCurrentVersion
        With objFile.AddParam(New FText, 0, 0, objFile.PScaleWidth, _
        frmForm.picPreview.TextHeight("Hi") * 3, "This file has been translated into Flow Chart using Unified Vector Space.", conBottom)
            .P.TextColour = QBColor(8)
        End With
    Else
        blnNativeFile = True
    End If
    
    intS = m_Drawing.Properties("ScaleMode")
    If Err Then Err.Clear: intS = vbInches
    
    sngHeight = frmForm.ScaleY(m_Drawing.Properties("PageHeight"), intS, vbTwips)
    If Err Then Err.Clear: sngHeight = frmForm.ScaleY(Printer.ScaleHeight, Printer.ScaleMode, vbTwips)
        
    objFile.ZoomPercent = m_Drawing.Properties("ZoomPercent")
    If Err Then Err.Clear: objFile.ZoomPercent = 100
    
    For Each objE In m_Drawing.Elements
        For lngP = 1 To objE.Points.Count - 1
            If blnNativeFile Then
                lngType = Val(objE.Tags("FlowItemNumber"))
                lngColour = Val(objE.Tags("ForeColour"))
                lngWidth = Val(objE.Tags("LineWidth")): If lngWidth <= 0 Then lngWidth = 1
            Else
                lngType = conAddLine
                lngColour = vbBlack
                lngWidth = 1 'default line width?
            End If
            
            Set objNew = Duplicate(, lngType)
            
            objNew.P.ForeColour = lngColour
            objNew.P.Text = objE.Tags("Text")
            objNew.P.LineWidth = lngWidth
        
            objFile.AddParam objNew, _
            frmForm.ScaleX(objE.Points(lngP).X, intS, vbTwips), _
            sngHeight - frmForm.ScaleY(objE.Points(lngP).Y, intS, vbTwips), _
            frmForm.ScaleX(objE.Points(lngP + 1).X - objE.Points(lngP).X, intS, vbTwips), _
            -frmForm.ScaleY(objE.Points(lngP + 1).Y - objE.Points(lngP).Y, intS, vbTwips), ""
        
        Next lngP
    Next objE
    
    frmForm.SetZoom 100
    Unload g_frmSplash
End Sub


Private Property Get IApplication_VersionCompatible() As Single
    IApplication_VersionCompatible = 1.008014!
End Property


Private Property Let IApplication_Visible(ByVal RHS As Boolean)
    frmForm.Visible = RHS
End Property

Private Property Get IApplication_Visible() As Boolean
    IApplication_Visible = frmForm.Visible
End Property
