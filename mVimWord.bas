Attribute VB_Name = "mVimWord"
' mVimWord
' Copyright (c) Chris White 2018
' CC-BY-NC-SA 4.0, or any later version, at your option.
' Thanks to https://glts.github.io/2013/04/28/vim-normal-mode-grammar.html
'   2018-04-06  chrisw  Initial version
'   2018-04-20  chrisw  Split vimRunCommand off VimDoCommand
'   2018-04-24  chrisw  Added counts to tTfF
'   2018-05-01  chrisw  Added pastes, voChange
'   2018-05-02  chrisw  Cleanup; fixed vmNonblankBackward; added ge, gE
'   2018-05-04  chrisw  Changed paste behaviour per Word
'   2018-05-07  chrisw  gp/gP now paste unformatted text; added ninja-feet
'   2018-05-10  chrisw  Fixed whitespace classes used for text objects

' General comment: Word puts the cursor between characters; Vim puts the
' cursor on characters.  This makes quite a difference.  I may need
' to go through later on and regularize the behaviour.

Option Explicit
Option Base 0

Public Sub VimDoCommand_About()
    MsgBox "VimWord version 0.2.10, 2018-05-10.  Copyright (c) 2018 Christopher White.  " & _
            "All Rights Reserved.  Licensed CC-BY-NC-SA 4.0 (or later).", _
            vbOKOnly + vbInformation, "About VimWord"
End Sub 'VimDoCommand_About
'

Public Sub VimDoCommand()       ' Grab and run a Vim command!
    VDCInternal False
End Sub 'VimDoCommand

Public Sub VimCommandLoop()     ' Grab and run a Vim command, and reopen the dialog for another command
    VDCInternal True
End Sub 'VimDoCommand

Private Sub VDCInternal(LoopIt As Boolean)
    Dim doc As Document: Set doc = Nothing

    On Error Resume Next: Set doc = ActiveDocument: On Error GoTo 0
    If doc Is Nothing Then Exit Sub

    Dim stay_in_normal As Boolean: stay_in_normal = False

    Do
        Dim proczone As Range, coll As Boolean, atStart As Variant
        atStart = Empty
        Set proczone = GetProczone_V(doc:=doc, _
                            iswholedoc:=coll, start_is_active:=atStart)

        If coll Then    ' coll => collapsed selection, so don't use the whole doc
                        ' (which is what GetProczone_V gave us)
            Set proczone = doc.ActiveWindow.Selection.Range.Duplicate
            atStart = Empty
        End If

        ' Get the command
        Dim frm As frmGrabKeys
        Set frm = New frmGrabKeys
        frm.Show

        Dim oper As VimOperator: oper = voUndef
        Dim cmd As VimCommand: cmd = vcUndef
        Dim motion As VimMotion: motion = vmUndef
        Dim operc As Long: operc = 0
        Dim motionc As Long: motionc = 0
        Dim cmdstr As String: cmdstr = ""
        Dim arg As String: arg = ""
        Dim ninja As VimNinja: ninja = vnUndef

        If Not frm.WasCancelled Then
            cmdstr = frm.Keys
            oper = frm.VOperator
            cmd = frm.VCommand
            motion = frm.VMotion
            operc = frm.VOperatorCount
            motionc = frm.VMotionCount
            arg = frm.VArg
            ninja = frm.VNinja
        End If

        Unload frm
        Set frm = Nothing
        If (cmd <> vcUndef) Or (oper <> voUndef And motion <> vmUndef) Then
            vimRunCommand doc, proczone, coll, atStart, oper, cmd, motion, _
                operc, motionc, cmdstr, arg, ninja
            Application.ScreenRefresh
        End If

        ' d and y leave you in normal mode; s and c do not.
        ' TODO expand this list in the future.
        stay_in_normal = (oper = voDelete) Or (oper = voYank)
    Loop While LoopIt And stay_in_normal
End Sub 'VDCInternal

Private Sub vimRunCommand( _
    doc As Document, _
    proczone As Range, _
    coll As Boolean, _
    atStart As Variant, _
    oper As VimOperator, _
    cmd As VimCommand, _
    motion As VimMotion, _
    operc As Long, _
    motionc As Long, _
    cmdstr As String, _
    arg As String, _
    ninja As VimNinja _
)
    Dim TITLE As String: TITLE = "Do Vim command"

    Dim CSET_WS As String
    CSET_WS = " " & ChrW(U_TAB) & ChrW(U_LF) & _
        ChrW(W_LINE_BREAK) & ChrW(W_FUNKY_BREAK) & Chr(U_CR) & Chr(W_NBSP) & _
        ChrW(&H1680) & ChrW(&H180E) & ChrW(&H2000) & ChrW(&H2001) & _
        ChrW(&H2002) & ChrW(&H2003) & ChrW(&H2004) & ChrW(&H2005) & _
        ChrW(&H2006) & ChrW(&H2007) & ChrW(&H2008) & ChrW(&H2009) & _
        ChrW(&H200A) & ChrW(&H200B) & ChrW(&H202F) & ChrW(&H205F) & _
        ChrW(&H3000) & ChrW(&HFEFF) & ChrW(&H2028) & ChrW(&H2029)
        ' NOT comment markers since I've been having problems with those lately

    Dim CSET_WS_ONELINE As String ' Whitespace on a single line
    CSET_WS_ONELINE = " " & ChrW(U_TAB) & Chr(W_NBSP) & _
        ChrW(&H1680) & ChrW(&H180E) & ChrW(&H2000) & ChrW(&H2001) & _
        ChrW(&H2002) & ChrW(&H2003) & ChrW(&H2004) & ChrW(&H2005) & _
        ChrW(&H2006) & ChrW(&H2007) & ChrW(&H2008) & ChrW(&H2009) & _
        ChrW(&H200A) & ChrW(&H200B) & ChrW(&H202F) & ChrW(&H205F) & _
        ChrW(&H3000) & ChrW(&HFEFF)

    Dim CSET_WS_BREAKS As String
    CSET_WS_BREAKS = ChrW(U_LF) & _
        ChrW(W_LINE_BREAK) & ChrW(W_FUNKY_BREAK) & Chr(U_CR) & _
        ChrW(&H2028) & ChrW(&H2029)

    ' Sanity check: we can have an operator or a command, but not both.
    If (oper <> voUndef And cmd <> vcUndef) Or _
            (oper = voUndef And cmd = vcUndef) Then
        MsgBox "Error (contact Chris White): operator " & CStr(oper) & _
                " with command " & CStr(cmd), vbExclamation, TITLE
        Exit Sub
    End If

    ' Save the incoming proczone for ninja-feet
    Dim origpzstart As Long, origpzend As Long
    origpzstart = proczone.Start
    origpzend = proczone.End

    Dim undos As UndoRecord
    Set undos = Application.UndoRecord

    ' Run the command

    On Error GoTo VRC_Err
    undos.StartCustomRecord TITLE & ": " & cmdstr
    Application.ScreenUpdating = False

    Dim count As Long
    count = operc * motionc  ' per motion.txt#operator

    Dim colldir As WdCollapseDirection
    colldir = wdCollapseEnd ' by default

    Dim idx As Long, result As Long, ltmp As Long

    ' === Motion ============================================================

    Select Case motion
        Case vmLeft: proczone.MoveStart wdCharacter, -count: colldir = wdCollapseStart
        Case vmRight: proczone.MoveEnd wdCharacter, count: colldir = wdCollapseEnd

        Case vmUp, vmDown:
            Set proczone = MoveVertical_( _
                motion = vmUp, _
                count, _
                (Not IsEmpty(atStart)) And atStart, _
                doc, _
                proczone, _
                colldir)

        Case vmStartOfLine, vmEOL:
            Set proczone = moveHorizontal_( _
                motion = vmStartOfLine, doc, proczone, colldir)

            If motion = vmEOL And count > 1 Then
                proczone.MoveEnd wdParagraph, count - 1
            End If

        Case vmStartOfParagraph: proczone.Start = proczone.Paragraphs(1).Range.Start: colldir = wdCollapseStart

        Case vmEOParagraph:
            proczone.Start = proczone.Paragraphs(1).Range.Start
            colldir = wdCollapseEnd
            If count > 1 Then
                proczone.MoveEnd wdParagraph, count - 1
            End If

        'TODO Case vmLine

        Case vmCharForward:
            arg = ExpandCSet_(arg)
            colldir = wdCollapseEnd
            For idx = 1 To count
                If proczone.MoveEndUntil(arg, wdForward) = 0 Then Exit For
                proczone.MoveEnd wdCharacter, 1     ' f => to and including
            Next idx

        Case vmCharBackward:
            arg = ExpandCSet_(arg)
            colldir = wdCollapseStart
            For idx = 1 To count
                If proczone.MoveStartUntil(arg, wdBackward) = 0 Then Exit For
                proczone.MoveStart wdCharacter, -1      ' F => to and including
            Next idx

        Case vmTilForward:
            arg = ExpandCSet_(arg)
            colldir = wdCollapseEnd
            result = proczone.MoveEndUntil(arg, wdForward)
            For idx = 2 To count
                If result = 0 Then Exit For
                proczone.MoveEnd wdCharacter, 1
                result = proczone.MoveEndUntil(arg, wdForward)
            Next idx

        Case vmTilBackward:
            arg = ExpandCSet_(arg)
            colldir = wdCollapseStart
            result = proczone.MoveStartUntil(arg, wdBackward)
            For idx = 2 To count
                If result = 0 Then Exit For
                proczone.MoveStart wdCharacter, -1
                result = proczone.MoveStartUntil(arg, wdBackward)
            Next idx

        Case vmWordForward:
            colldir = wdCollapseEnd
            proczone.MoveEnd wdWord, count

        Case vmEOWordForward:
            colldir = wdCollapseEnd
            proczone.MoveEnd wdWord, count
            proczone.MoveEndWhile CSET_WS, wdBackward

        Case vmWordBackward:
            colldir = wdCollapseStart
            proczone.MoveStart wdWord, -count

        Case vmEOWordBackward:      ' ge
            colldir = wdCollapseEnd
            proczone.MoveStart wdWord, -count
            ltmp = proczone.End
            proczone.Collapse wdCollapseStart
            proczone.Expand wdWord
            proczone.Collapse wdCollapseEnd
            proczone.MoveStartWhile CSET_WS, wdBackward
            proczone.End = ltmp

        Case vmNonblankForward:
            colldir = wdCollapseEnd
            For idx = 1 To count
                proczone.MoveEndUntil CSET_WS, wdForward
                proczone.MoveEndWhile CSET_WS, wdForward
            Next idx

        Case vmEONonblankForward:
            colldir = wdCollapseEnd
            proczone.MoveEndUntil CSET_WS, wdForward
            For idx = 2 To count
                proczone.MoveEndWhile CSET_WS, wdForward
                proczone.MoveEndUntil CSET_WS, wdForward
            Next idx

        Case vmNonblankBackward:
            colldir = wdCollapseStart

            ' TODO handle failures (MoveStartUntil returns 0).  Move to
            ' beginning of paragraph?  Likewise, handle errors in
            ' all other MoveStart/MoveEnd calls throughout.
            ' Test case: in the first nonblank in the file.
            For idx = 1 To count
                proczone.MoveStartWhile CSET_WS, wdBackward
                    ' In case the cursor is already on whitespace.
                    ' TODO adjust similarly throughout if necessary
                proczone.MoveStartUntil CSET_WS, wdBackward
            Next idx

        Case vmEONonblankBackward:  'gE
            colldir = wdCollapseEnd

            For idx = 1 To count - 1
                proczone.MoveStartWhile CSET_WS, wdBackward
                proczone.MoveStartUntil CSET_WS, wdBackward
            Next idx

            proczone.MoveStartWhile CSET_WS, wdBackward

        Case vmSentenceForward: proczone.MoveEnd wdSentence, count: colldir = wdCollapseEnd
        Case vmSentenceBackward: proczone.MoveStart wdSentence, -count: colldir = wdCollapseStart
        Case vmParaForward: proczone.MoveEnd wdParagraph, count: colldir = wdCollapseEnd
        Case vmParaBackward: proczone.MoveStart wdParagraph, -count: colldir = wdCollapseStart

        ' H, M, L
        ' -------
        ' TODO only in wdMainTextStory:
        ' Use selection.Move{Up,Down} wdScreen to get the screen size implicitly.
        ' May need to use ranges elsewhere in the document if we're too close to the end.
        ' Use ActiveWindow.ActivePane.VerticalPercentScrolled to find out where the top is
        ' in percentage, and then binary-search the range to get close.
        ' Jump there, then ScrollIntoView just to be on the safe side.

        'Case vmScreenTop
        'Case vmScreenMiddle
        'Case vmScreenBottom

        ' Text objects
        ' ------------

        ' Non-collapsing ones
        Case vmAWord:
            proczone.Expand wdWord
            coll = False
            If count > 1 Then proczone.MoveEnd wdWord, count - 1

        Case vmIWord:
            proczone.Expand wdWord
            coll = False
            If count > 1 Then proczone.MoveEnd wdWord, count - 1
            proczone.MoveEndWhile CSET_WS, wdBackward

        Case vmANonblank:
            coll = False
            proczone.MoveStartUntil CSET_WS, wdBackward
            For idx = 1 To count
                proczone.MoveEndUntil CSET_WS, wdForward
                proczone.MoveEndWhile CSET_WS, wdForward
            Next idx
            proczone.MoveEndWhile CSET_WS_BREAKS, wdBackward
                ' aW includes the trailing WS, but don't grab the paragraph mark as well.

        Case vmINonblank:
            coll = False
            proczone.MoveStartUntil CSET_WS, wdBackward
            proczone.MoveEndUntil CSET_WS, wdForward
            For idx = 2 To count
                proczone.MoveEndWhile CSET_WS, wdForward
                proczone.MoveEndUntil CSET_WS, wdForward    ' iW excludes the trailing WS
            Next idx

        Case vmASentence:
            proczone.Expand wdSentence
            coll = False
            If count > 1 Then proczone.MoveEnd wdSentence, count - 1
            proczone.MoveEndWhile CSET_WS_BREAKS, wdBackward
                ' So deleting the last sentence in the paragraph doesn't
                ' remove the break between that paragraph and the next

        Case vmISentence:
            proczone.Expand wdSentence
            If count > 1 Then proczone.MoveEnd wdSentence, count - 1
            proczone.MoveEndWhile CSET_WS, wdBackward
            coll = False

        Case vmAPara:
            proczone.Expand wdParagraph
            coll = False
            If count > 1 Then proczone.MoveEnd wdParagraph, count - 1

        Case vmIPara
            proczone.Expand wdParagraph
            If count > 1 Then proczone.MoveEnd wdParagraph, count - 1
            proczone.MoveEndWhile Chr(13), -1   ' Only the last Chr(13)
            coll = False

        Case Else:
            If cmd = vcUndef Then GoTo VRC_Finally     ' Unimplemented is not an error
    End Select ' motion

    ' Ninja-feet
    Select Case ninja
        Case vnLeft     ' [
            proczone.End = origpzend
        Case vnRight
            proczone.Start = origpzstart
    End Select

    ' === Operator/Command ==================================================

    ' Process it.  We have either an operator or a command.
    ' TODO merge operator and command enums, since they are mutually exclusive.

    If oper <> vcUndef Then     ' Operators
        Select Case oper
            Case voYank:
                If proczone.Start <> proczone.End Then proczone.Copy
                GoTo VRC_Finally

            Case voDelete:
                If proczone.Start <> proczone.End Then proczone.Cut
                GoTo VRC_Finally

            Case voChange:
                If proczone.Start <> proczone.End Then proczone.Cut

            ' voGo, voSelect handled below
        End Select 'operator

        If (oper = voGo) And coll Then
            proczone.Collapse colldir
        End If

        proczone.Select     ' Handles voSelect

    Else                        ' Commands

        Dim issearch As Boolean: issearch = False
        Dim searchforward As Boolean
        Dim searchwholeword As Boolean

        Dim ispaste As Boolean: ispaste = False
        Dim paste_plain As Boolean
        Dim paste_backup As Boolean

        Select Case cmd

            ' Searches
            Case vcSearchWholeItemForward:
                issearch = True: searchforward = True: searchwholeword = True

            Case vcSearchWholeItemBackward:
                issearch = True: searchforward = False: searchwholeword = True

            Case vcSearchItemForward:
                issearch = True: searchforward = True: searchwholeword = False

            Case vcSearchItemBackward:
                issearch = True: searchforward = False: searchwholeword = False

            ' Pastes
            Case vcPutAfter:
                ispaste = True: paste_backup = False: paste_plain = False

            Case vcPutAfterG:
                ispaste = True: paste_backup = False: paste_plain = True

            Case vcPutBefore:
                ispaste = True: paste_backup = True: paste_plain = False

            Case vcPutBeforeG:
                ispaste = True: paste_backup = True: paste_plain = True

            Case Else: GoTo VRC_Finally
        End Select

        If issearch Then
            proczone.Select
            With doc.ActiveWindow.Selection
                .Collapse IIf(searchforward, wdCollapseEnd, wdCollapseStart)
                .Find.Execute proczone.Text, MatchWholeWord:=searchwholeword, Forward:=searchforward, Wrap:=wdFindContinue
            End With

        ElseIf ispaste Then
            ' TODO implement counts
            If Not paste_plain Then
                proczone.Paste

            Else
                ' PasteSpecial leaves the range collapsed at the end of
                ' what was pasted.  Therefore, save/restore the start.
                Dim pzstart As Long
                pzstart = proczone.Start

                proczone.PasteSpecial Link:=False, DataType:=20, Placement:=wdInLine, _
                    DisplayAsIcon:=False    ' 20 = unformatted Unicode text

                proczone.Start = pzstart
            End If

            proczone.Collapse IIf(paste_backup, wdCollapseStart, wdCollapseEnd)
            proczone.Select

        End If

    End If

VRC_Finally:
    On Error Resume Next    ' or else errors in the cleanup code cause infinite loops
    Application.ScreenUpdating = True
    undos.EndCustomRecord
    Exit Sub

VRC_Err:
    MsgBox "Got error " & CStr(Err.Number) & ": " & vbCrLf & _
            Err.Description, vbExclamation, TITLE
    Resume VRC_Finally

End Sub 'vimRunCommand

Private Function MoveVertical_(isUp As Boolean, count As Long, _
                            atStart As Boolean, doc As Document, _
                            proczone As Range, _
                            ByRef colldir As WdCollapseDirection) As Range

    Dim r As Range
    Set r = proczone.Duplicate
    r.Select

    colldir = IIf(atStart, wdCollapseStart, wdCollapseEnd)
    With doc.ActiveWindow.Selection
        .Collapse colldir
        If isUp Then .MoveUp wdLine, count Else .MoveDown wdLine, count
        If atStart Then r.Start = .Start Else r.End = .End
    End With

    Set MoveVertical_ = r
End Function 'MoveVertical_
'

Private Function moveHorizontal_( _
                goToStartOfLine As Boolean, doc As Document, _
                proczone As Range, ByRef colldir As WdCollapseDirection)

    Dim r As Range
    Set r = proczone.Duplicate
    r.Select

    colldir = IIf(goToStartOfLine, wdCollapseStart, wdCollapseEnd)
    With doc.ActiveWindow.Selection
        .Collapse colldir
        If goToStartOfLine Then .HomeKey wdLine Else .EndKey wdLine
        If goToStartOfLine Then r.Start = .Start Else r.End = .End
    End With

    Set moveHorizontal_ = r
End Function 'MoveHorizontal_
'

Private Function GetProczone_V(Optional ByRef iswholedoc As Variant, _
                                Optional doc As Variant, _
                                Optional selrange As Range, _
                                Optional ByRef start_is_active As Variant) As Range
' Get the processing zone for _doc_ (or ActiveDocument):
'   - If there is a selection, the selection
'   - If there is no selection, the full story in which the cursor
'       currently appears.
' Returns the proczone.  Sets IsWholeDoc to True iff there is no selection.
' Returns Nothing if _doc_ is invalid.
' Uses selrange instead of Selection if selrange is provided.
' If not iswholedoc, and selrange is not provided, sets start_is_active
' to Selection.StartIsActive.

    Dim thedoc As Document: Set thedoc = Nothing
    Dim retval As Range
    Dim wdretv As Boolean   'is-Whole-Doc RETurn Value
    wdretv = False
    Dim using_sel As Boolean: using_sel = False     ' Whether we read selection
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
        using_sel = True
    Else
        If selrange Is Nothing Then ' duplicate because no short-circuit eval
            Set retval = thedoc.ActiveWindow.Selection.Range.Duplicate
            using_sel = True
        Else
            Set retval = selrange.Duplicate
        End If
    End If

    If using_sel And Not IsMissing(start_is_active) Then
        start_is_active = thedoc.ActiveWindow.Selection.StartIsActive
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

Private Function ExpandCSet_(arg As String) As String
' Expand characters to include Unicode equivalents
    ExpandCSet_ = arg
    Select Case arg
        Case " ": ExpandCSet_ = " " & ChrW(W_NBSP) & ChrW(U_TAB)

        Case "'": ExpandCSet_ = "'" & ChrW(U_CURLY_APOS) & ChrW(U_CURLY_BACKQUOTE)

        Case """": ExpandCSet_ = """" & _
                    ChrW(U_CURLY_OPENDQUOTE) & _
                    ChrW(U_CURLY_CLOSEDQUOTE)

        Case "-": ExpandCSet_ = "-" & _
                    ChrW(U_OPT_HYPHEN) & _
                    ChrW(U_REAL_HYPHEN) & _
                    ChrW(U_NONBREAK_HYPHEN) & _
                    ChrW(U_FIGURE_DASH) & _
                    ChrW(U_EN_DASH) & _
                    ChrW(U_EM_DASH) & _
                    ChrW(U_WAVE_DASH) & _
                    ChrW(U_FULLWIDTH_TILDE) & _
                    ChrW(W_NBHYPHEN) & _
                    ChrW(W_OPTHYPHEN)
    End Select
End Function 'ExpandCSet_
'

