Attribute VB_Name = "mVimWord"
' mVimWord
' Copyright (c) Chris White 2018--2019
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
'   2018-06-07  chrisw  Hack in voDelete/voChange for strange Word behaviour.
'   2018-06-26  chrisw  ip: don't select whole table cell.  It still doesn't
'                       work with a count, though, and I'm not sure why.  E.g., 2vip
'   2018-06-29  chrisw  Added voDrop; bugfix in ip
'   2018-07-16  chrisw  Fixed g$ to be (in effect) an abbreviation for ]ip.
'                       (I think it's easier to remember than ]ip.)
'                       Fixed the count on $ to operate line-by-line.
'                       Added vmLine support.
'   2018-08-18  chrisw  Added basic register support to c, d, y, p
'   2019-02-14  chrisw  Switched Undo to cUndoWrapper for 2007 compatibility
'   2019-02-08  chrisw  Added VIMWORD_* version constants.
'                       Added LastViewType_ material: switch to Normal view
'                       before any use of Range.Move{Start,End}{While,Until} (#7).
'   2019-02-19  chrisw  Fixed logic for vmEOWordBackward (ge) and vmEONonblankBackward (gE).
'                       Be more selective about ViewNormal_ calls.
'   2019-03-05  chrisw  Implemented H, M, L: added HandleHML_ material

' General comment: Word puts the cursor between characters; Vim puts the
' cursor on characters.  This makes quite a difference.  I may need
' to go through later on and regularize the behaviour.

Option Explicit
Option Base 0

' Version info
Private Const VIMWORD_VERSION = "0.3.4-pre.1"
Private Const VIMWORD_DATE = "2019-03-05"

' Scratchpad filename, lower case for comparison
Private Const SCRATCHPAD_FN_LC = "vimwordscratchpad.dotm"

' Cache for the document index of the open scratchpad.  Trust but verify.
Private ScratchpadDocIndex_ As Long

' Storage for the last command, to support `.`.  NOTE: This only lasts until
' the next time the VBA project is reset.
' Public since it's used by frmGrabKeys; here so that it will stick around
' after a frmGrabKeys instance is unloaded.
Public VimLastCommand_ As String
'

' Storage for the view, in case we need to switch to Normal.
Dim LastViewType_ As WdViewType
Dim DidStashView_ As Boolean
'

Public Sub VimDoCommand_About()
    MsgBox SPrintF("VimWord version %s (%s).  Copyright (c) 2018--2019 Christopher White.  " & _
            "All Rights Reserved.  Licensed CC-BY-NC-SA 4.0 (or later)." & vbCrLf & _
            "Uses code by Phlip Bradbury <phlipping@yahoo.com>.", _
            VIMWORD_VERSION, VIMWORD_DATE), _
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

    DidStashView_ = False
    
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

        Dim reg As String: reg = vbNullString
        Dim oper As VimOperator: oper = voUndef
        Dim cmd As VimCommand: cmd = vcUndef
        Dim motion As VimMotion: motion = vmUndef
        Dim operc As Long: operc = 0
        Dim motionc As Long: motionc = 0
        Dim cmdstr As String: cmdstr = vbNullString
        Dim arg As String: arg = vbNullString
        Dim ninja As VimNinja: ninja = vnUndef
        Dim space As Boolean: space = False
        Dim has_count As Boolean: has_count = False
        
        If Not frm.WasCancelled Then
            cmdstr = frm.Keys
            reg = frm.VRegister
            oper = frm.VOperator
            cmd = frm.VCommand
            motion = frm.VMotion
            operc = frm.VOperatorCount
            motionc = frm.VMotionCount
            arg = frm.VArg
            ninja = frm.VNinja
            space = frm.VSpace
            has_count = (frm.VHasOperatorCount Or frm.VHasMotionCount)
        End If

        Unload frm
        Set frm = Nothing
        If (cmd <> vcUndef) Or (oper <> voUndef And motion <> vmUndef) Then
            vimRunCommand doc, proczone, coll, atStart, reg, oper, cmd, motion, _
                operc, motionc, cmdstr, arg, ninja, space, has_count
            Application.ScreenRefresh
        End If

        ' d and y leave you in normal mode; s and c do not.
        ' TODO expand this list in the future.
        stay_in_normal = (oper = voDelete) Or (oper = voYank)
    Loop While LoopIt And stay_in_normal
    
    ' Restore the view if we changed it
    If DidStashView_ Then
        doc.ActiveWindow.View = LastViewType_
    End If
    
End Sub 'VDCInternal

Private Sub vimRunCommand( _
    doc As Document, _
    proczone As Range, _
    coll As Boolean, _
    atStart As Variant, _
    reg As String, _
    oper As VimOperator, _
    cmd As VimCommand, _
    motion As VimMotion, _
    operc As Long, _
    motionc As Long, _
    cmdstr As String, _
    arg As String, _
    ninja As VimNinja, _
    space As Boolean, _
    has_count As Boolean _
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

    Dim CSET_WS_ONELINE As String ' Whitespace on a single line, WITHOUT comment markers
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

    Dim undo As New cUndoWrapper

    ' Run the command

    On Error GoTo VRC_Err
    undo.Start TITLE & ": " & cmdstr
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
                Dim dummy_colldir As WdCollapseDirection
                Set proczone = MoveVertical_(False, count - 1, False, doc, proczone, dummy_colldir)
                Set proczone = moveHorizontal_(False, doc, proczone, dummy_colldir)
                    ' Make sure we're at the end, since moving down doesn't necessarily
                    ' keep us at EOL
            End If

        Case vmStartOfParagraph:    ' 0 - special command
            proczone.Start = proczone.Paragraphs(1).Range.Start
            colldir = wdCollapseStart

        Case vmLine
            If Not has_count Then   ' G without a count => last line
                proczone.MoveEnd wdStory
                colldir = wdCollapseEnd
                
            Else
                Dim linerange As Range
                Set linerange = proczone.GoTo(What:=wdGoToLine, which:=wdGoToFirst, count:=count)
                If linerange.Start < proczone.Start Then
                    proczone.Start = linerange.Start
                    colldir = wdCollapseStart
                Else
                    proczone.End = linerange.End
                    colldir = wdCollapseEnd
                End If
            End If
            
        Case vmCharForward:
            'ViewToNormal_ doc  ' Instead, use the following precondition check.
            If InStr(arg, ChrW(W_COMMENT)) > 0 Then GoTo VRC_Finally
            
            arg = ExpandCSet_(arg)  ' Never adds W_COMMENT, by invariant.
            
            colldir = wdCollapseEnd
            For idx = 1 To count
                If proczone.MoveEndUntil(arg, wdForward) = 0 Then Exit For
                proczone.MoveEnd wdCharacter, 1     ' f => to and including
            Next idx

        Case vmCharBackward:
            'ViewToNormal_ doc  ' Instead, use the following precondition check.
            If InStr(arg, ChrW(W_COMMENT)) > 0 Then GoTo VRC_Finally
            arg = ExpandCSet_(arg)
            
            colldir = wdCollapseStart
            For idx = 1 To count
                If proczone.MoveStartUntil(arg, wdBackward) = 0 Then Exit For
                proczone.MoveStart wdCharacter, -1      ' F => to and including
            Next idx

        Case vmTilForward:          ' t
            'ViewToNormal_ doc  ' Instead, use the following precondition check.
            If InStr(arg, ChrW(W_COMMENT)) > 0 Then GoTo VRC_Finally
            
            arg = ExpandCSet_(arg)
            colldir = wdCollapseEnd
            result = proczone.MoveEndUntil(arg, wdForward)
            For idx = 2 To count
                If result = 0 Then Exit For
                proczone.MoveEnd wdCharacter, 1
                result = proczone.MoveEndUntil(arg, wdForward)
            Next idx

        Case vmTilBackward:         ' T
            'ViewToNormal_ doc  ' Instead, use the following precondition check.
            If InStr(arg, ChrW(W_COMMENT)) > 0 Then GoTo VRC_Finally
            
            arg = ExpandCSet_(arg)
            colldir = wdCollapseStart
            result = proczone.MoveStartUntil(arg, wdBackward)
            For idx = 2 To count
                If result = 0 Then Exit For
                proczone.MoveStart wdCharacter, -1
                result = proczone.MoveStartUntil(arg, wdBackward)
            Next idx

        Case vmWordForward:         ' w
            colldir = wdCollapseEnd
            proczone.MoveEnd wdWord, count

        Case vmEOWordForward:       ' e
            colldir = wdCollapseEnd
            proczone.MoveEnd wdWord, count
            If proczone.Characters.count <> Len(proczone.Text) Then
                proczone.End = proczone.Start + Len(proczone.Text)
                ' If the range ends with a comment character, exclude it to
                ' avoid lockups (see #7).
            End If
            If proczone.Comments.count > 0 Then ViewToNormal_ doc
            proczone.MoveEndWhile CSET_WS, wdBackward

        Case vmWordBackward:        ' b
            colldir = wdCollapseStart
            proczone.MoveStart wdWord, -count

        Case vmEOWordBackward:      ' ge
            colldir = wdCollapseEnd
            
            proczone.Expand wdWord  ' Get to the start of the current word
            proczone.MoveStart wdWord, -count   ' Back up the right number of words
            
            proczone.Collapse wdCollapseStart   ' Get just that word
            proczone.Expand wdWord
            If proczone.Comments.count > 0 Then ViewToNormal_ doc   ' #7
            
            proczone.MoveEndWhile CSET_WS, wdBackward

        Case vmNonblankForward:     ' W
            ViewToNormal_ doc   ' TODO do we need this?  Maybe rewrite in terms of MoveEnd, with a manual WS check
            colldir = wdCollapseEnd
            For idx = 1 To count
                proczone.MoveEndUntil CSET_WS, wdForward
                proczone.MoveEndWhile CSET_WS, wdForward
            Next idx

        Case vmEONonblankForward:   ' E
            ViewToNormal_ doc   ' TODO do we need this?
            colldir = wdCollapseEnd
            proczone.MoveEndUntil CSET_WS, wdForward
            For idx = 2 To count
                proczone.MoveEndWhile CSET_WS, wdForward
                proczone.MoveEndUntil CSET_WS, wdForward
            Next idx

        Case vmNonblankBackward:    ' B
            ViewToNormal_ doc   ' TODO do we need this?
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
            ViewToNormal_ doc   ' TODO do we need this?
            colldir = wdCollapseEnd
            
            proczone.Expand wdWord

            For idx = 1 To count - 1
                proczone.MoveStartWhile CSET_WS, wdBackward
                proczone.MoveStartUntil CSET_WS, wdBackward
            Next idx

            proczone.MoveStartWhile CSET_WS, wdBackward
            proczone.Collapse wdCollapseStart

        Case vmSentenceForward: proczone.MoveEnd wdSentence, count: colldir = wdCollapseEnd
        Case vmSentenceBackward: proczone.MoveStart wdSentence, -count: colldir = wdCollapseStart
        Case vmParaForward: proczone.MoveEnd wdParagraph, count: colldir = wdCollapseEnd
        Case vmParaBackward: proczone.MoveStart wdParagraph, -count: colldir = wdCollapseStart

        ' H, M, L
        ' -------

        Case vmScreenTop, vmScreenMiddle, vmScreenBottom:
            Dim hmlrange As Range
            Set hmlrange = HandleHML_(motion, doc)
            If Not (hmlrange Is Nothing) Then Set proczone = hmlrange
            
        ' Text objects
        ' ------------

        ' Non-collapsing ones
        Case vmAWord:
            proczone.Expand wdWord
            coll = False
            If count > 1 Then proczone.MoveEnd wdWord, count - 1

        Case vmIWord:
            'ViewToNormal_ doc
            proczone.Expand wdWord
            coll = False
            If count > 1 Then proczone.MoveEnd wdWord, count - 1
                
            If proczone.Characters.count <> Len(proczone.Text) Then
                proczone.End = proczone.Start + Len(proczone.Text)
                ' If the range ends with a comment character, exclude it to
                ' avoid lockups (see #7).
            End If
            If proczone.Comments.count > 0 Then ViewToNormal_ doc

            proczone.MoveEndWhile CSET_WS, wdBackward

        Case vmANonblank:
            ViewToNormal_ doc
            coll = False
            proczone.MoveStartUntil CSET_WS, wdBackward
            For idx = 1 To count
                proczone.MoveEndUntil CSET_WS, wdForward
                proczone.MoveEndWhile CSET_WS, wdForward
            Next idx
            proczone.MoveEndWhile CSET_WS_BREAKS, wdBackward
                ' aW includes the trailing WS, but don't grab the paragraph mark as well.

        Case vmINonblank:
            ViewToNormal_ doc
            coll = False
            
            ' TODO FIXME: MoveStartUntil fails if the first word in the
            ' document starts at the first character of the document.
            proczone.MoveStartUntil CSET_WS, wdBackward
            
            proczone.MoveEndUntil CSET_WS, wdForward
            For idx = 2 To count
                proczone.MoveEndWhile CSET_WS, wdForward
                proczone.MoveEndUntil CSET_WS, wdForward    ' iW excludes the trailing WS
            Next idx

        Case vmASentence:
            'ViewToNormal_ doc
            proczone.Expand wdSentence
            coll = False
            If count > 1 Then proczone.MoveEnd wdSentence, count - 1
            If proczone.Comments.count > 0 Then ViewToNormal_ doc
            proczone.MoveEndWhile CSET_WS_BREAKS, wdBackward
                ' So deleting the last sentence in the paragraph doesn't
                ' remove the break between that paragraph and the next

        Case vmISentence:
            'ViewToNormal_ doc
            proczone.Expand wdSentence
            If count > 1 Then proczone.MoveEnd wdSentence, count - 1
            If proczone.Comments.count > 0 Then ViewToNormal_ doc
            proczone.MoveEndWhile CSET_WS, wdBackward
            coll = False

        Case vmAPara:
            proczone.Expand wdParagraph
            coll = False
            If count > 1 Then proczone.MoveEnd wdParagraph, count - 1

        Case vmIPara, vmEOParagraph
            ' Include EOParagraph (g$) here to avoid duplicating code
            'ViewToNormal_ doc
            
            proczone.Expand wdParagraph
            If motion = vmEOParagraph Then proczone.Start = origpzstart
            
            If count > 1 Then proczone.MoveEnd wdParagraph, count - 1
            
            ' If we were in a table, the end marker is now selected.
            ' The marker is a Chr(13) & Chr(7), but MoveEndWhile won't move
            ' over it.  Therefore, test for it and skip it manually.
            If proczone.Tables.count >= 1 Then
                If proczone.Cells.count >= 1 Then
                    If Right(proczone.Text, 1) = ChrW(7) Then
                        proczone.MoveEnd wdCharacter, -1
                    End If
                End If
            End If
            If proczone.Comments.count > 0 Then ViewToNormal_ doc
            proczone.MoveEndWhile Chr(13), -1   ' Only the last Chr(13)
            coll = (motion = vmEOParagraph)

        Case vmMSWordSelection
            If proczone.Start = proczone.End Then GoTo VRC_Finally
                ' If nothing is selected, don't try to take action.
            
        Case Else
            If cmd = vcUndef Then GoTo VRC_Finally     ' Unimplemented is not an error
    End Select ' motion

    ' Ninja-feet
    Select Case ninja
        Case vnLeft     ' [
            proczone.End = origpzend
        Case vnRight
            proczone.Start = origpzstart
    End Select
    
    ' Extra whitespace.  Takes effect after ninja-feet.
    If space And (colldir = wdCollapseEnd) Then
        'ViewToNormal_ doc
        If proczone.Comments.count > 0 Then ViewToNormal_ doc
        proczone.MoveEndWhile CSET_WS_ONELINE, wdForward
    ElseIf space And (colldir = wdCollapseStart) Then
        'ViewToNormal_ doc
        If proczone.Comments.count > 0 Then ViewToNormal_ doc
        proczone.MoveStartWhile CSET_WS_ONELINE, wdBackward
    End If
    
    ' === Operator/Command ==================================================

    ' Process it.  We have either an operator or a command.
    ' TODO merge operator and command enums, since they are mutually exclusive.

    If oper <> vcUndef Then     ' Operators
    
        If proczone.Start <> proczone.End Then  ' c d y X
            Select Case oper
                Case voYank:
                    If reg = vbNullString Then
                        proczone.Copy
                    Else
                        SaveRangeToRegister_ proczone, reg
                    End If
                    
                    GoTo VRC_Finally
    
                Case voDelete, voChange, voDrop:
                    ' Word doesn't always delete the whole selection!
                    Dim endr As Range
                    Set endr = proczone.Duplicate
                    endr.Collapse wdCollapseEnd
                    endr.MoveEnd wdCharacter, 1
                    
                    If oper = voDrop Then
                        proczone.Delete
                    Else
                        If reg = vbNullString Then
                            proczone.Cut
                        Else
                            SaveRangeToRegister_ proczone, reg
                            proczone.Delete
                        End If
                    End If
                    
                    If endr.Characters.count > 1 Then     ' something strange happened
                        If endr.Characters.First = ChrW(13) Then
                            endr.End = endr.Start + 1
                            endr.Delete
                        End If
                    End If
                            
                    GoTo VRC_Finally
    
                ' voGo, voSelect handled below
            End Select 'operator
        End If 'something is selected

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
        End Select 'cmd

        If issearch Then
            proczone.Select
            With doc.ActiveWindow.Selection
                .Collapse IIf(searchforward, wdCollapseEnd, wdCollapseStart)
                .Find.Execute proczone.Text, MatchWholeWord:=searchwholeword, Forward:=searchforward, Wrap:=wdFindContinue
            End With

        ElseIf ispaste Then
            ' TODO implement counts
            
            If reg <> vbNullString Then
                ReplaceRangeFromRegister_ proczone, reg, (Not paste_plain)
            Else
                If Not paste_plain Then
                    proczone.Paste
                        ' Removes whatever is in the proczone, then adds the
                        ' pasted text after the proczone.  Leaves the
                        ' selection at the end of the pasted text.
                    
                Else
                    ' PasteSpecial leaves the range collapsed at the end of
                    ' what was pasted.  Therefore, save/restore the start.
                    Dim pzstart As Long
                    pzstart = proczone.Start
    
                    proczone.PasteSpecial Link:=False, DataType:=20, Placement:=wdInLine, _
                        DisplayAsIcon:=False    ' 20 = unformatted Unicode text
    
                    proczone.Start = pzstart
                End If
                
            End If
            
            proczone.Collapse IIf(paste_backup, wdCollapseStart, wdCollapseEnd)
            proczone.Select

        End If

    End If

VRC_Finally:
    On Error Resume Next    ' or else errors in the cleanup code cause infinite loops
    Application.ScreenUpdating = True
    Exit Sub    ' undo Class_Terminate automatically closes the custom undo record

VRC_Err:
    MsgBox "Got error " & CStr(Err.Number) & ": " & vbCrLf & _
            Err.Description, vbExclamation, TITLE
    Resume VRC_Finally

End Sub 'vimRunCommand

Private Function MoveVertical_(isUp As Boolean, count As Long, _
                            atStart As Boolean, doc As Document, _
                            proczone As Range, _
                            ByRef colldir_out As WdCollapseDirection) As Range

    Dim r As Range
    Set r = proczone.Duplicate
    r.Select

    colldir_out = IIf(atStart, wdCollapseStart, wdCollapseEnd)
    With doc.ActiveWindow.Selection
        .Collapse colldir_out
        If isUp Then .MoveUp wdLine, count Else .MoveDown wdLine, count
        If atStart Then r.Start = .Start Else r.End = .End
    End With

    Set MoveVertical_ = r
End Function 'MoveVertical_
'

Private Function moveHorizontal_( _
                goToStartOfLine As Boolean, doc As Document, _
                proczone As Range, ByRef colldir_out As WdCollapseDirection)

    Dim r As Range
    Set r = proczone.Duplicate
    r.Select

    colldir_out = IIf(goToStartOfLine, wdCollapseStart, wdCollapseEnd)
    With doc.ActiveWindow.Selection
        .Collapse colldir_out
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
' Expand characters to include Unicode equivalents.
' Invariant: ExpandCSet_ will never add W_COMMENT to a set.

    ExpandCSet_ = arg
    Select Case arg
        Case " ": ExpandCSet_ = " " & ChrW(W_NBSP) & ChrW(U_TAB)
            ' Not W_COMMENT

        Case "'": ExpandCSet_ = "'" & ChrW(U_CURLY_APOS) & ChrW(U_CURLY_BACKQUOTE)
            ' But not ` (single backquote).

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

Public Sub MarkScratchpadAsSaved_(Optional AlsoClose = False)
    Dim idx As Long, t As Template, d As Document
    
    On Error Resume Next
    
    Set d = FindScratchpad_(True)
    If Not (d Is Nothing) Then
        d.Saved = True
        If AlsoClose Then d.Close wdDoNotSaveChanges
    End If
    
    ' Now handle the template itself
    For idx = 1 To Templates.count
        Set t = Templates(idx)
        If LCase(t.Name) = SCRATCHPAD_FN_LC Then
            t.Saved = True
            ' I don't see a way to close the template itself.
            Exit For
        End If
    Next idx
End Sub 'MarkScratchpadAsSaved_
'

Public Function FindScratchpad_(Optional noerror As Boolean = False) As Document
' The Templates collection is indexed by full name rather than name, so we have to iterate it.
    Dim idx As Long: idx = -1
    Dim t As Template, d As Document
    
    ' Check the cached document index
    Set d = Nothing
    On Error Resume Next
    Set d = Documents(ScratchpadDocIndex_)
    On Error GoTo 0
    
    If Not (d Is Nothing) Then
        If LCase(d.Name) = SCRATCHPAD_FN_LC Then GoTo FS_GotDoc
    End If
    
    ' If no luck, check all the open documents
    For idx = 1 To Documents.count
        Set d = Documents(idx)
        If LCase(d.Name) = SCRATCHPAD_FN_LC Then GoTo FS_GotDoc
    Next idx
    
    ' Failing that, check all the templates
    For idx = 1 To Templates.count
        Set t = Templates(idx)
        If LCase(t.Name) = SCRATCHPAD_FN_LC Then GoTo FS_GotTemplate
    Next idx
    
FS_Oops:

    ' If we got here, we can't find the scratchpad.  Flag it so calling code is simpler.
    If noerror Then
        Set FindScratchpad_ = Nothing
    Else
        Err.Raise vbObjectError + 512 + 513, "VimWord.dotm", "Could not locate " & SCRATCHPAD_FN_LC & " - try reinstalling"
    End If
    
    Exit Function

FS_GotTemplate:     ' We found a template and need to open it as a document

    Set d = t.OpenAsDocument
    d.ActiveWindow.Visible = False      ' Hidden
    d.Saved = True
    
    ' Just to be safe, don't assume the index of the new document is Documents.Count.
    Dim d2 As Document
    For idx = 1 To Documents.count
        Set d2 = Documents(idx)
        If d2.FullName = d.FullName Then GoTo FS_GotDoc
    Next idx
    
    ' If we got here, we opened the document, but somehow couldn't find it.  Bail.
    d.Close False
    GoTo FS_Oops
    
FS_GotDoc:          ' We found an already-open document
    Set FindScratchpad_ = d
    If idx <> -1 Then ScratchpadDocIndex_ = idx
    
End Function 'FindScratchpad_
'

Private Function GetRegister_(reg As String) As Range
    Set GetRegister_ = Nothing
    
    Dim sp As Document
    Set sp = FindScratchpad_
    
    Dim paranum As Long
    paranum = Asc(LCase(reg)) + 1
    If paranum > 128 Then
        Err.Raise vbObjectError + 512 + 514, "VimWord.dotm", _
            "I don't understand register '" & reg & "' (para. number " & CStr(paranum) & ")."
        Exit Function
    End If

    Dim retval As Range
    Set retval = sp.Paragraphs(paranum).Range.Duplicate
    retval.MoveEnd wdCharacter, -1      ' Not the Chr(13) at the end
    
    Set GetRegister_ = retval
End Function 'GetRegister_
'

Private Sub SaveRangeToRegister_(proczone As Range, reg As String)
    Dim source As Range
    Set source = proczone.Duplicate
    
    Dim dest As Range
    Set dest = GetRegister_(reg)
    
    If source.Characters.last = Chr(13) Then
        ' Remember that the proczone ended with a Chr(13) since we can't
        ' store the Chr(13) directly in the single paragraph the register
        ' is allocated.
        source.MoveEnd wdCharacter, -1
        dest.FormattedText = source
        dest.InsertAfter ChrW(U_PU1)
        dest.MoveEnd wdCharacter, 1
    Else
        dest.FormattedText = source
    End If
    
    MarkScratchpadAsSaved_  ' Don't save changes to disk
End Sub 'SaveRangeToRegister_
'

Private Sub ReplaceRangeFromRegister_(proczone As Range, reg As String, formatted As Boolean)
    Dim source As Range
    Set source = GetRegister_(reg)
    
    Dim need_para As Boolean: need_para = False
    
    If source.Characters.last = ChrW(U_PU1) Then    ' The register ends with a paragraph mark
        need_para = True
        source.MoveEnd wdCharacter, -1      ' TODO check this for robustness in tables
    End If
    
    If formatted Then
        proczone.FormattedText = source
    Else
        proczone.Text = source.Text
    End If
    
    If need_para Then
        proczone.InsertAfter Chr(13)
        proczone.MoveEnd wdCharacter, 1
    End If
End Sub 'ReplaceRangeFromRegister_
'

Private Sub ViewToNormal_(doc As Document)
    DidStashView_ = True
    LastViewType_ = doc.ActiveWindow.View
    doc.ActiveWindow.View = wdNormalView
End Sub 'ViewToNormal_

Private Function HandleHML_(motion As VimMotion, doc As Document) As Range
    Set HandleHML_ = Nothing
    
    Dim visible_range As Range, proczone As Range
    
    Set visible_range = GetVisibleRange(doc)
    If visible_range Is Nothing Then
        Exit Function   ' *** EXIT POINT ***
    End If
    
    Set proczone = visible_range.Duplicate
    
    If motion <> vmScreenMiddle Then
        proczone.Collapse IIf(motion = vmScreenTop, wdCollapseStart, wdCollapseEnd)
        
    Else    ' vmScreenMiddle
        ' TODO improve this hack - for now, go halfway down the range by Chr(13)s.
        Dim RE13 As VBScript_RegExp_55.RegExp
        Dim matches As VBScript_RegExp_55.MatchCollection
        
        Set RE13 = New VBScript_RegExp_55.RegExp
        RE13.Pattern = Chr(13)
        RE13.Global = True
        Set matches = RE13.Execute(proczone.Text)
        
        If matches.count < 5 Then
            ' Also a hack - for screens with few ^p's visible, just use the middle by characters.
            proczone.End = proczone.Start + (proczone.End - proczone.Start) / 2
            proczone.Collapse wdCollapseEnd
            
        Else
            Dim half13s As Long
            half13s = matches.count / 2
            proczone.Collapse wdCollapseStart
            Do While half13s > 0
                proczone.MoveEndUntil Chr(13)       ' TODO do I need to switch to Print Layout here?
                                                    ' I don't think so, because ^p's are visible in every view.
                proczone.MoveEnd wdCharacter, 1     ' Past the Chr(13) so the next MoveEndUntil will work
                half13s = half13s - 1
            Loop
            proczone.Collapse wdCollapseEnd
        End If
    End If
    
    Set HandleHML_ = proczone
End Function 'HandleHML_
'
