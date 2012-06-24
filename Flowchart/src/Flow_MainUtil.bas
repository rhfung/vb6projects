Attribute VB_Name = "modMainUtil"
Option Explicit
'August 21 2002

#If Prev = 0 Then
Public frmMain  As frmIMain
#End If


#If Prev = 0 Then
Public Sub CheckFileBounds(File As FlowChart, ScaleWidth As Single, ScaleHeight As Single, col As Collection)
'Called from OpenFile().
'Checks to see if any item is off the page.
    Dim objItem     As FlowItem
    Dim strText     As String
    Dim lngOutCount As Long 'count of off-boundary items
    
    For Each objItem In File
        If CLng(objItem.P.Left + objItem.P.Width) > ScaleWidth Or _
           CLng(objItem.P.Top + objItem.P.Height) > ScaleHeight Then
            
'            strText = objItem.P.Text
'            frmLog.AddLine "Object with text """ & IIf(Len(strText) > 20, Left$(strText, 20) & "...", strText) & """ is off the paper size."
            col.Add objItem
            lngOutCount = lngOutCount + 1
        ElseIf objItem.P.Left < 0 Or objItem.P.Top < 0 Then
'            strText = objItem.P.Text
'            frmLog.AddLine "Object with text """ & IIf(Len(strText) > 20, Left$(strText, 20) & "...", strText) & """ is off of the page."
            col.Add objItem
            lngOutCount = lngOutCount + 1
        End If
'        If TypeOf objItem Is FPicture Then
'            If Not GetPicture(objItem).IsLoaded Then
'                frmLog.AddLine "The picture with the file """ & objItem.P.Text & """ was not loaded."
'            End If
'        End If
    Next objItem
    
    If lngOutCount > 0 Then
        If MsgBox("There are " & lngOutCount & " object(s) that are not on the page." & vbNewLine & "Do you want to move them onto the page?" & vbNewLine & "[ Right-click on the layer bar (left side) to reveal contents of the hidden object(s). ]", vbQuestion Or vbYesNo) = vbYes Then
            For Each objItem In File
                CheckObjBounds objItem, ScaleWidth, ScaleHeight
            Next objItem
            'frmLog.AddLine "Move objects around the perimeter of the page to their proper locations."
            File.Changed = True
        'Else
            'frmLog.AddLine "There are " & lngOutCount & " item(s) that are off the view of the page.  These items may not print properly."
        End If
    End If
End Sub

Public Sub CheckObjBounds(Item As FlowItem, ScaleWidth As Single, ScaleHeight As Single)
    With Item.P
        If .Left < 0 Then .Left = 0
        If .Left > ScaleWidth Then .Left = ScaleWidth
        If .Top < 0 Then .Top = 0
        If .Top > ScaleHeight Then .Top = ScaleHeight
        
        If .Width >= 0 Then 'positive width
            If .Left + .Width > ScaleWidth Then
                .Left = frmMain.Grid(ScaleWidth - .Width)
                'too big to fit
                If .Left < 0 Then .Left = 0: .Width = ScaleWidth
            End If
        Else 'negative width
            If .Left + .Width < 0 Then .Left = -.Width
        End If
        
        If .Height >= 0 Then
            If .Top + .Height > ScaleHeight Then
                .Top = frmMain.Grid(ScaleHeight - .Height)
                'too big to fit
                If .Top < 0 Then .Top = 0: .Height = ScaleHeight
            End If
        Else
            If .Top + .Height < 0 Then .Top = -.Height
        End If
    End With
End Sub


Public Sub CheckRectBounds(Rect As Rect, ScaleWidth As Single, ScaleHeight As Single)
    Dim objItem As FlowItem
    
    Set objItem = New FlowItem
    With objItem.P
        .Left = Rect.X1
        .Top = Rect.Y1
        .Width = Rect.X2 - Rect.X1
        .Height = Rect.Y2 - Rect.Y1
    End With
    
    CheckObjBounds objItem, ScaleWidth, ScaleHeight
    
    With objItem.P
        Rect.X1 = .Left
        Rect.Y1 = .Top
        Rect.X2 = .Width + Rect.X1
        Rect.Y2 = .Height + Rect.Y1
    End With
End Sub

Public Function ConvCRtoCRLF(Text As String) As String
'Used by spell checker
    Dim lngCh As Long
    
    For lngCh = 1 To Len(Text)
        If Mid$(Text, lngCh, 2) = vbCrLf Then
            ConvCRtoCRLF = ConvCRtoCRLF & vbCrLf
            lngCh = lngCh + 1 'loop will advance another
        ElseIf Mid$(Text, lngCh, 1) = vbCr Then 'return ch
            ConvCRtoCRLF = ConvCRtoCRLF & vbCrLf
        Else
            ConvCRtoCRLF = ConvCRtoCRLF & Mid$(Text, lngCh, 1)
        End If
    Next lngCh
End Function
#End If


Public Function GetButton(ByVal FI As FlowItem) As FButton
    Set GetButton = FI
End Function

Public Function GetPicture(ByVal FI As FlowItem) As FPicture
    Set GetPicture = FI
End Function

Public Function IsObjLine(ByVal Obj As FlowItem) As Boolean
'Three kinds of lines.
    IsObjLine = (TypeOf Obj Is FArrowLine) Or (TypeOf Obj Is FLine) Or (TypeOf Obj Is FMidArrowLine)
End Function

#If Prev = 0 Then
Public Sub ExportWord(ByVal File As FlowChart, ByVal Converter As Form)
    Dim objCanvas As Word.CanvasShapes
    Dim objDoc    As Word.Document
    Dim objItem   As FlowItem
    Dim objShape  As Word.Shape
    Dim sngLeft%, sngTop%, sngWidth%, sngHeight%
    Dim strTemp   As String
    
    On Error GoTo Handler
       
    Screen.MousePointer = vbHourglass
    
    Set objDoc = CreateObject("Word.Document")
    objDoc.Application.Visible = False
    Set objDoc = objDoc.Application.Documents.Add()
    
    Set objCanvas = objDoc.Shapes.AddCanvas(0, 0, Converter.ScaleX(File.PScaleWidth, vbTwips, vbPoints), Converter.ScaleX(File.PScaleHeight, vbTwips, vbPoints)).CanvasItems
    With objCanvas
        For Each objItem In File
            sngLeft = Converter.ScaleX(objItem.P.Left, vbTwips, vbPoints)
            sngTop = Converter.ScaleX(objItem.P.Top, vbTwips, vbPoints)
            sngWidth = Converter.ScaleX(objItem.P.Width, vbTwips, vbPoints)
            sngHeight = Converter.ScaleX(objItem.P.Height, vbTwips, vbPoints)
            Select Case objItem.Number
            Case conAddLine, conAddEndArrowLine, conAddMidArrowLine
                Set objShape = .AddLine(sngLeft, sngTop, sngLeft + sngWidth, sngTop + sngHeight)
            Case conAddCircle
                Set objShape = .AddShape(msoShapeFlowchartConnector, sngLeft, sngTop, sngWidth, sngHeight)
            Case conAddTerminator
                Set objShape = .AddShape(msoShapeFlowchartTerminator, sngLeft, sngTop, sngWidth, sngHeight)
            Case conAddRect
                Set objShape = .AddShape(msoShapeRectangle, sngLeft, sngTop, sngWidth, sngHeight)
            Case conAddDecision
                Set objShape = .AddShape(msoShapeFlowchartDecision, sngLeft, sngTop, sngWidth, sngHeight)
            Case conAddEllipse
                Set objShape = .AddShape(msoShapeOval, sngLeft, sngTop, sngWidth, sngHeight)
            Case conAddPicture
            
                Set objShape = .AddPicture(GetPicture(objItem).FullFilename(File), False, True, sngLeft, sngTop, sngWidth, sngHeight)
                
            Case conAddInOut
                Set objShape = .AddShape(msoShapeTrapezoid, sngLeft, sngTop, sngWidth, sngHeight)
            
            Case conAddText
                Set objShape = .AddShape(msoShapeRectangle, sngLeft, sngTop, sngWidth, sngHeight)
                objShape.Line.Visible = msoFalse
                objShape.Fill.Visible = msoFalse
            End Select
            
            If objItem.Number <> conAddPicture Then
                objShape.TextFrame.MarginLeft = 2
                objShape.TextFrame.MarginRight = 2
                objShape.TextFrame.MarginTop = 2
                objShape.TextFrame.MarginBottom = 2
                objShape.TextEffect.FontName = objItem.P.FontFace
                objShape.TextEffect.FontSize = objItem.P.TextSize
                objShape.AlternativeText = "FlowChart"
                objShape.TextFrame.TextRange.Text = objItem.P.Text
            End If
        Next objItem
    
    End With
    
    objDoc.Application.Visible = True
    Screen.MousePointer = vbDefault
    
    Exit Sub
Handler:
    MsgBox "Failed to export Flow Chart to Microsoft Word application.", vbExclamation, "Export Word"
    If Not objDoc Is Nothing Then
        objDoc.Close
        objDoc.Application.Quit
    End If
    
    Screen.MousePointer = vbDefault
    Exit Sub
End Sub
#End If

