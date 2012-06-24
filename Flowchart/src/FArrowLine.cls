VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FArrowLine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit 'Using Properties
'Richard Fung.  August 12, 2000.
'FArrowLine (FEndArrowLine)
'Search and replace FlowItem.

Implements FlowItem

'Const conSFlow = 300 'flow line
Const conSPointed = 180 'pointed arrow

Const conStdArrow = 3600 'twips
Const conLinExSpace = 30 'extra space for line text
Const conNormArrow = 30 * conR 'degrees conv to radians
Const conEnggArrow = 15 * conR 'degrees conv to radians

Private mobjProp As Properties


Private Sub Class_Initialize()
    Set mobjProp = New Properties
    mobjProp.FillStyle = vbFSSolid 'solid arrow head
End Sub

Private Sub Class_Terminate()
    Set mobjProp = Nothing
End Sub


Private Property Get FlowItem_CanRefresh() As Boolean
    'nothing to do
End Property

Private Property Get FlowItem_Description() As String
    FlowItem_Description = "End Arrow Line"
End Property


Private Property Get FlowItem_DescriptionF() As String
    FlowItem_DescriptionF = "Arrow Line"
End Property

'Private Property Get FlowItem_CenterX() As Single
'    FlowItem_CenterX = mobjProp.CenterX
'End Property
'
'Private Property Get FlowItem_CenterY() As Single
'    FlowItem_CenterY = mobjProp.CenterY
'End Property


Private Function FlowItem_Draw(ByVal View As Object, ByVal Parent As FlowChart) As Long
    Dim angle As Double, b As Double, c As Double
    Dim sX As Single, sY As Single, tRect As Rect
    Dim sngS As Single 'was a sngStant, but now single
    Dim tPoints(0 To 3) As POINTAPI, lngLoop As Long 'arrow head code
    
    SetDrawProps View, Me, Parent
    
    tRect = GetRect(Me)
    
    View.Line (tRect.x1, tRect.y1)-(tRect.x2, tRect.y2) 'draw line
    
    'calculate angle
    If (tRect.x2 - tRect.x1) = 0 Then
        If tRect.y2 > tRect.y1 Then
            angle = 270 * conR
        Else
            angle = 90 * conR
        End If
    Else
        angle = Atn((tRect.y1 - tRect.y2) / (tRect.x2 - tRect.x1))
    End If
    'convert arrow point to proper side
    If (tRect.x2 - tRect.x1) < 0 Then
        angle = angle + 180 * conR
    End If
    
    'calc new angles
    If mobjProp.ArrowEngg Then
        b = angle + conEnggArrow
        c = angle - conEnggArrow
    Else
        b = angle + conNormArrow
        c = angle - conNormArrow
    End If
    'calculate length of line
'    If Sqr((tRect.y2 - tRect.y1) ^ 2 + (tRect.x2 - tRect.x1) ^ 2) > sngStdArrow Then
'        sngS = sngSFlow 'size of arrow
'        sX = mobjProp.CenterX 'too long, put in middle
'        sY = mobjProp.CenterY
'    Else
        sngS = conSPointed * mobjProp.ArrowSize / 100 'size of arrow
        sX = tRect.x2 'short enough for end
        sY = tRect.y2
        'View.Line (sX + Cos(b) * sngS * -1, sY + Sin(b) * sngS)-(sX + Cos(c) * sngS * -1, sY + Sin(c) * sngS)
        tPoints(0).X = sX + Cos(b) * sngS * -1
        tPoints(0).Y = sY + Sin(b) * sngS
        tPoints(1).X = sX
        tPoints(1).Y = sY
        tPoints(2).X = sX + Cos(c) * sngS * -1
        tPoints(2).Y = sY + Sin(c) * sngS
        tPoints(3).X = tPoints(0).X 'return to point 0
        tPoints(3).Y = tPoints(0).Y
        
        If mobjProp.FillStyle <> vbFSTransparent Then  'ArrowSolid
            For lngLoop = 0 To 3
                With tPoints(lngLoop)
                    .X = View.ScaleX(.X, View.ScaleMode, vbPixels)
                    .Y = View.ScaleY(.Y, View.ScaleMode, vbPixels)
                End With
            Next lngLoop
        
            View.FillStyle = vbFSSolid
            View.FillColor = View.ForeColor
            Call Polygon(View.hdc, tPoints(0), 4)
        Else 'Transparent arrow - skip pixel conversion and DrawPolygon
            View.FillStyle = vbFSTransparent
            View.Line (tPoints(0).X, tPoints(0).Y)-(tPoints(1).X, tPoints(1).Y)
            View.Line (tPoints(1).X, tPoints(1).Y)-(tPoints(2).X, tPoints(2).Y)
        End If
        View.FillStyle = vbFSTransparent
'    End If
    'draw the lines
    'View.Line (sX, sY)-(sX + Cos(b) * sngS * -1, sY + Sin(b) * sngS)
    'View.Line (sX, sY)-(sX + Cos(c) * sngS * -1, sY + Sin(c) * sngS)
    'View.Line (sX + Cos(b) * sngS * -1, sY + Sin(b) * sngS)-(sX + Cos(c) * sngS * -1, sY + Sin(c) * sngS)
    
    'calculate position of X, Y for text
    DrawTextForLines View, Me, Parent
'    If mobjProp.Width = 0 Then
'        View.CurrentY = mobjProp.CenterY - View.TextHeight(mobjProp.Text) / 2
'    Else
'        If mobjProp.Height < 0 Then 'Y
'            View.CurrentY = mobjProp.CenterY + conLinExSpace
'        Else
'            View.CurrentY = mobjProp.CenterY - View.TextHeight(mobjProp.Text) - conLinExSpace
'        End If
'    End If
'    If mobjProp.Height = 0 Then
'        View.CurrentX = mobjProp.CenterX - View.TextWidth(mobjProp.Text) / 2
'    Else
'        If mobjProp.Width < 0 Then 'X
'            View.CurrentX = mobjProp.CenterX - View.TextWidth(mobjProp.Text) - conLinExSpace
'        Else
'            View.CurrentX = mobjProp.CenterX + conLinExSpace
'        End If
'    End If
    'draw text
    'SetFontProps View, Me, Parent
    'View.ForeColor = mobjProp.TextColour
    'View.Print mobjProp.Text
End Function

Private Property Get FlowItem_Number() As FAddType
    FlowItem_Number = conAddEndArrowLine
End Property

Private Property Set FlowItem_P(ByVal RHS As Properties)
    Set mobjProp = RHS
End Property

Private Property Get FlowItem_P() As Properties
    Set FlowItem_P = mobjProp
End Property



Private Sub FlowItem_Refresh(ByVal Parent As FlowChart, ByVal View As Object)
    'nothing to do
End Sub


Private Property Get FlowItem_TextBottomMrg() As Single
    FlowItem_TextBottomMrg = 0
End Property

Private Property Get FlowItem_TextLeftMrg() As Single
    FlowItem_TextLeftMrg = 0
End Property

Private Property Get FlowItem_TextRightMrg() As Single
    FlowItem_TextRightMrg = 0
End Property

Private Property Get FlowItem_TextTopMrg() As Single
    FlowItem_TextTopMrg = 0
End Property



