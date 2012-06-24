VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PUndo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"PUndoClass"
Attribute VB_Ext_KEY = "Member0" ,"PUndoClass"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
'March 23, 2001

Public Enum PUndoAction
    conUndoNothing = -100
    conUndoSize = 1
    conUndoMove
    conUndoTextChange
    conUndoDelete
    conUndoCut
    conUndoAutoSize
    conUndoAlignment
    conUndoPicReset
    conUndoPicZoom
    conUndoFontChanged
    conUndoForeColourChange
    conUndoBackColourChange 'includes FillStyle
    conUndoTextColourChange
    conUndoObjectFormatFillStyle
    conUndoObjectFormatLineStyle
    conUndoObjectFormatLineWidth
    conUndoChangeTag3Macro 'Tag3 also modifies Text
    conUndoChangeTag3Field
    conUndoChangeArrowSolid 'Tag10
    conUndoChangeArrowType 'Tag10
    conUndoChangeArrowSize 'Tag11
    conUndoChangeName
    conUndoObject
    conUndoMultiple = 1001
End Enum


Public Event UndoItemChanged(Item As IUndoInterface)

Private Const MAX_UNDO = 7

'local variable to hold collection
Private mobjUndo(1 To MAX_UNDO) As IUndoInterface

Public Function Add(Ptr As FlowItem, theActionWas As PUndoAction) As PUndoItem
    Dim objNewMember As PUndoItem
    Dim lngShift    As Long
    
    'shift the undo items
    For lngShift = MAX_UNDO - 1 To 1 Step -1
        Set mobjUndo(lngShift + 1) = mobjUndo(lngShift)
    Next lngShift

    Set mobjUndo(1) = New PUndoItem
    Set objNewMember = mobjUndo(1)

    'set the properties passed into the method
    CopyProperties Ptr.P, objNewMember.P
    Set objNewMember.Ptr = Ptr
    objNewMember.theActionWas = theActionWas
    
    'return the object created
    Set Add = objNewMember
    RaiseEvent UndoItemChanged(objNewMember)
    Set objNewMember = Nothing
End Function

Public Function AddMany(Many As Collection, Optional theActionWas As PUndoAction = conUndoMultiple) As PUndoMulti
    Dim objNewMember    As PUndoMulti
    Dim lngShift        As Long
    Dim objItem         As FlowItem
    
    'shift the undo items
    For lngShift = MAX_UNDO - 1 To 1 Step -1
        Set mobjUndo(lngShift + 1) = mobjUndo(lngShift)
    Next lngShift

    Set mobjUndo(1) = New PUndoMulti
    Set objNewMember = mobjUndo(1)

    'set the properties passed into the method
    For Each objItem In Many
        objNewMember.Add objItem, theActionWas
    Next objItem
    
    'return the object created
    Set AddMany = objNewMember
    RaiseEvent UndoItemChanged(objNewMember)
    Set objNewMember = Nothing
End Function


Public Sub DoUndoOp(File As FlowChart)
    Dim objItem As PUndoItem
    
    If Last.IsMultiple = False Then
        With LastItem
            Select Case .theActionWas 'choose the undo action
            Case conUndoDelete, conUndoCut
                File.Add .Ptr 'restore lost object
            Case conUndoMove
                .Ptr.P.Left = .P.Left
                .Ptr.P.Top = .P.Top
            Case conUndoSize, conUndoAutoSize, conUndoPicReset, conUndoPicZoom, conUndoAlignment
                If .Ptr.Number = conAddPicture Then
                    GetPicture(.Ptr).SetCustomSize 'allow changes to be made
                End If
                .Ptr.P.Left = .P.Left
                .Ptr.P.Top = .P.Top
                .Ptr.P.Width = .P.Width
                .Ptr.P.Height = .P.Height
            Case conUndoTextChange
                .Ptr.P.Text = .P.Text
            Case conUndoBackColourChange
                .Ptr.P.BackColour = .P.BackColour
                .Ptr.P.FillStyle = .P.FillStyle
            Case conUndoTextColourChange
                .Ptr.P.TextColour = .P.TextColour
            Case conUndoForeColourChange
                .Ptr.P.ForeColour = .P.ForeColour
            Case conUndoFontChanged
                .Ptr.P.FontFace = .P.FontFace
                .Ptr.P.TextSize = .P.TextSize
                .Ptr.P.TextBold = .P.TextBold
                .Ptr.P.TextItalic = .P.TextItalic
                .Ptr.P.TextUnderline = .P.TextUnderline
            Case conUndoObjectFormatFillStyle
                .Ptr.P.FillStyle = .P.FillStyle
            Case conUndoObjectFormatLineStyle
                .Ptr.P.LineStyle = .P.LineStyle
            Case conUndoObjectFormatLineWidth
                .Ptr.P.LineWidth = .P.LineWidth
            Case conUndoChangeTag3Field, conUndoChangeTag3Macro
                .Ptr.P.Tag3 = .P.Tag3
                .Ptr.P.Text = .P.Text
            Case conUndoChangeArrowSize
                .Ptr.P.ArrowSize = .P.ArrowSize
            Case conUndoChangeArrowType
                .Ptr.P.ArrowEngg = .P.ArrowEngg
            Case conUndoChangeArrowSolid
                .Ptr.P.FillStyle = .P.FillStyle
            Case conUndoChangeName
                .Ptr.P.Name = .P.Name
            Case conUndoObject, conUndoMultiple 'copy all the properties
                CopyProperties .P, .Ptr.P
            Case Else
                MsgBox "Undo command cannot be done.  Internal error.", vbExclamation
                Exit Sub
            End Select
        End With
    Else 'multiple
        For Each objItem In LastMany
            If objItem.theActionWas <> conUndoDelete Then
                'restore previous properties
                CopyProperties objItem.P, objItem.Ptr.P
                objItem.Ptr.P.GroupNo = objItem.P.GroupNo 'restore group number
            Else
                'add back the pointer to the object
                objItem.Ptr.P.GroupNo = objItem.P.GroupNo
                File.Add objItem.Ptr
            End If
        Next objItem
    End If
    RemoveLast 'last undo is erased from UNDO queue
End Sub

Public Function Last() As IUndoInterface
Attribute Last.VB_UserMemId = 0
    'used when referencing an element in the collection
    'vntIndexKey contains either the Index or Key to the collection,
    'this is why it is declared as a Variant
    'Syntax: Set foo = x.Item(xyz) or Set foo = x.Item(5)
    Set Last = mobjUndo(1)
End Function

Public Function LastItem() As PUndoItem
    Set LastItem = mobjUndo(1)
End Function

Public Function LastMany() As PUndoMulti
    Set LastMany = mobjUndo(1)
End Function


Public Sub RemoveLast()
    Dim lngShift    As Long
    
    'shift the undo items
    For lngShift = 1 To MAX_UNDO - 1
        Set mobjUndo(lngShift) = mobjUndo(lngShift + 1)
    Next lngShift
    Set mobjUndo(MAX_UNDO) = Nothing
    RaiseEvent UndoItemChanged(mobjUndo(1))
End Sub



Private Sub Class_Terminate()
    'destroys collection when this class is terminated
    Erase mobjUndo
End Sub

