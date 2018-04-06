Attribute VB_Name = "mVimWord"
'mVimWord
'Copyright (c) Chris White 2018
'CC-BY-NC-SA 4.0, or any later version, at your option.

Option Explicit
Option Base 0

Public Sub VimDoCommand()
' Grab and run a Vim command!
    Dim TITLE As String: TITLE = "Do Vim command"
    Dim doc As Document: Set doc = Nothing

    Dim CSET_WS As String: CSET_WS = " " & Chr(9) & Chr(10) & Chr(12) & Chr(13)
        ' NOT comment markers since I've been having problems with those lately
        
    On Error Resume Next: Set doc = ActiveDocument: On Error GoTo 0
    If doc Is Nothing Then Exit Sub

    Dim proczone As Range, coll As Boolean
    Set proczone = GetProczone_V(iswholedoc:=coll, doc:=doc)
    If coll Then    ' coll => collapsed selection
        Set proczone = doc.ActiveWindow.Selection.Range.Duplicate
    End If

    ' Get the command
    Dim frm As frmGrabKeys
    Set frm = New frmGrabKeys
    frm.Show
    Dim cmd As VimOperator
    Dim motion As VimMotion
    Dim cmdc As Long
    Dim motionc As Long
    Dim cmdstr As String
    Dim arg As String
    
    If Not frm.WasCancelled Then
        cmdstr = frm.Keys
        cmd = frm.VCommand
        motion = frm.VMotion
        cmdc = frm.VCommandCount
        motionc = frm.VMotionCount
        arg = frm.VArg
    End If
    
    Unload frm
    Set frm = Nothing
    If cmd = vcUndef Or motion = vmUndef Then
        Exit Sub
    End If
    
    Dim count As Long
    count = cmdc * motionc  ' per motion.txt#operator
    
    Dim undos As UndoRecord
    Set undos = Application.UndoRecord

    ' Run the command
    
    On Error GoTo VDC_Err
    undos.StartCustomRecord TITLE & ": " & cmdstr
    Application.ScreenUpdating = False

    Dim colldir As WdCollapseDirection
    colldir = wdCollapseEnd ' by default
    
    Select Case motion
        Case vmLeft: proczone.MoveStart wdCharacter, -count: colldir = wdCollapseStart
        Case vmRight: proczone.MoveEnd wdCharacter, count: colldir = wdCollapseEnd
        
        'Case vmUp:     ' TODO figure out how to get these
        'Case vmDown:
        'Case vmStartOfLine:
        Case vmStartOfParagraph: proczone.Start = proczone.Paragraphs(1).Range.Start: colldir = wdCollapseStart
        'Case vmEOL:
        'Case vmLine:
        Case vmCharForward:
            colldir = wdCollapseEnd
            If proczone.MoveEndUntil(arg, wdForward) <> 0 Then
                proczone.MoveEnd wdCharacter, 1     ' f => to and including
            End If
            
        Case vmCharBackward:
            colldir = wdCollapseStart
            If proczone.MoveEndUntil(arg, wdBackward) <> 0 Then
                proczone.MoveStart wdCharacter, -1     ' F => to and including
            End If
            
        Case vmTilForward: proczone.MoveEndUntil arg, wdForward: colldir = wdCollapseEnd
        Case vmTilBackward: proczone.MoveStartUntil arg, wdBackward: colldir = wdCollapseStart
        
        Case vmWordForward: proczone.MoveEnd wdWord, 1: colldir = wdCollapseEnd
        Case vmEOWordForward:
            colldir = wdCollapseEnd
            proczone.MoveEnd wdWord, 1
            proczone.MoveEndWhile CSET_WS, wdBackward
            
        Case vmWordBackward: proczone.MoveStart wdWord, -1: colldir = wdCollapseStart
            
        Case vmNonblankForward:
            colldir = wdCollapseEnd
            proczone.MoveEndUntil CSET_WS, wdForward
            proczone.MoveEndWhile CSET_WS, wdForward
            
        Case vmEONonblankForward:
            colldir = wdCollapseEnd
            proczone.MoveEndUntil CSET_WS, wdForward
            
        Case vmNonblankBackward:
            colldir = wdCollapseStart
            proczone.MoveStartUntil CSET_WS, wdBackward
            
        Case vmSentenceForward: proczone.MoveEnd wdSentence, 1: colldir = wdCollapseEnd
        Case vmSentenceBackward: proczone.MoveStart wdSentence, -1: colldir = wdCollapseStart
        Case vmParaForward: proczone.MoveEnd wdParagraph, 1: colldir = wdCollapseEnd
        Case vmParaBackward: proczone.MoveStart wdParagraph, -1: colldir = wdCollapseStart
    
        ' Non-collapsing ones
        Case vmAWord: proczone.Expand wdWord: coll = False
        Case vmIWord: proczone.Expand wdWord: proczone.MoveEndWhile CSET_WS, wdBackward: coll = False
        Case vmANonblank:
            coll = False
            proczone.MoveStartUntil CSET_WS, wdBackward
            proczone.MoveEndUntil CSET_WS, wdForward
            proczone.MoveEndWhile CSET_WS, wdForward    ' aW includes the trailing WS
            
        Case vmINonblank:
            coll = False
            proczone.MoveStartUntil CSET_WS, wdBackward
            proczone.MoveEndUntil CSET_WS, wdForward
            ' NO proczone.MoveEndWhile CSET_WS, wdForward    ' iW excludes the trailing WS
        
        Case vmASentence: proczone.Expand wdSentence: coll = False
        Case vmISentence: proczone.Expand wdSentence: proczone.MoveEndWhile CSET_WS, wdBackward: coll = False
        Case vmAPara: proczone.Expand wdParagraph: coll = False
        Case vmIPara: proczone.Expand wdParagraph: proczone.MoveEndWhile CSET_WS, wdBackward: coll = False
        
        Case Else: GoTo VDC_Finally
        
    End Select

    Select Case cmd
        Case vcDelete:
            proczone.Delete
            GoTo VDC_Finally
        Case vcYank:
            proczone.Copy
            GoTo VDC_Finally
    End Select
    
    If (cmd = vcGo) And coll Then
        proczone.Collapse colldir
    End If
    
    proczone.Select
    
VDC_Finally:
    On Error Resume Next    ' or else errors in the cleanup code cause infinite loops
    Application.ScreenUpdating = True
    undos.EndCustomRecord
    Exit Sub

VDC_Err:
    MsgBox "Got error " & CStr(Err.Number) & ": " & vbCrLf & _
            Err.Description, vbExclamation, TITLE
    Resume VDC_Finally

End Sub 'VimDoCommand

Private Function GetProczone_V(Optional ByRef iswholedoc As Variant, _
                                Optional doc As Variant, _
                                Optional selrange As Range) As Range
' Get the processing zone for _doc_ (or ActiveDocument):
'   - If there is a selection, the selection
'   - If there is no selection, the full story in which the cursor
'       currently appears.
' Returns the proczone.  Sets IsWholeDoc to True iff there is no selection.
' Returns Nothing if _doc_ is invalid.
' Uses selrange instead of Selection if selrange is provided.

    Dim thedoc As Document: Set thedoc = Nothing
    Dim retval As Range
    Dim wdretv As Boolean   'is-Whole-Doc RETurn Value
    wdretv = False
    
    If IsMissing(doc) Then
        On Error Resume Next: Set thedoc = ActiveDocument: On Error GoTo 0
    Else
        Set thedoc = doc
    End If

    If thedoc Is Nothing Then
        Set GetProczone_V = Nothing
        Exit Function
    End If

    If IsMissing(selrange) Then
        Set retval = thedoc.ActiveWindow.Selection.Range.Duplicate
    Else
        If selrange Is Nothing Then
            Set retval = thedoc.ActiveWindow.Selection.Range.Duplicate
        Else
            Set retval = selrange.Duplicate
        End If
    End If
    
    If retval.Start = retval.End Then
        ' Select the whole story the selection is in.  This is my empirical
        ' way of doing so; hopefully there's a better way.
        Dim en As Long
        retval.EndOf wdStory        ' Find the end
        en = retval.End
        retval.StartOf wdStory      ' Anchor the start
        retval.End = en             ' Expand to the whole story
        wdretv = True
    End If

    Set GetProczone_V = retval
    If Not IsMissing(iswholedoc) Then iswholedoc = wdretv

End Function 'GetProczone_V
'


