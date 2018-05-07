VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmGrabKeys 
   Caption         =   "Run Vim command"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   OleObjectBlob   =   "frmGrabKeys.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmGrabKeys"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' frmGrabKeys: collect keypresses and parse Vim commands
' Copyright (c) 2018 Chris White.  All rights reserved.
' See changelog in mVimWord for other changes
'   2018-04-06  chrisw  Initial version
'   2018-04-20  chrisw  Major expansion/rewrite
'   2018-04-24  chrisw  Change "s" to "v" (visual selection)
'   2018-04-26  chrisw  Changed regex from hand-generated to re2vba.pl
'   2018-05-02  chrisw  Refactored motion code into ProcessMotion_
'   2018-05-07  chrisw  Added ninja-feet; refactored regex

' NOTE: the consolidated reference is in :help normal-index

Option Explicit
Option Base 0

' Vim support
Public Enum VimRegister     ' from :help registers
    vrUndef
    vrUnnamed   ' "
    vr0
    vr1
    vr2
    vr3
    vr4
    vr5
    vr6
    vr7
    vr8
    vr9
    vrSmallDelete   ' -
    vrA
    vrB
    vrC
    vrD
    vrE
    vrF
    vrG
    vrH
    vrI
    vrJ
    vrK
    vrL
    vrM
    vrN
    vrO
    vrP
    vrQ
    vrR
    vrS
    vrT
    vrU
    vrV
    vrW
    vrX
    vrY
    vrZ
    vrColon
    vrDot
    vrFilename      ' %
    vrLastFilename  ' #
    vrExpression    ' =
    vrClipboard     ' * Win clipboard; X selection
    vrPlus          ' + Win clipboard; X clipboard
    vrTilde         ' dropped text from last drag-and-drop
    vrBlackHole     ' underscore
    vrLastPattern   ' /
End Enum 'VimRegister

Public Enum VimCommand      ' Intransitive commands
' thanks to https://www.fprintf.net/vimCheatSheet.html and :help change.txt

    vcUndef

    ' Note: intransitive motions (e.g., 0, ^, $) are handled with a fake operator voGo.
    'vcAppend       ' a
    'vcAppendEOL    ' A
    'vcInsert       ' i
    'vcInsertSOT    ' I
    'vcReplace      ' R
    'vcOpen         ' o
    'vcOpenAbove    ' O

    'vcDelAfter     ' x
    'vcDelBefore    ' X
    ' TODO D, C, s, S

    ' TODO implement registers

    ' Pastes.  NOTE: These behave a bit differently from Vim because
    ' Word puts the cursor between letters instead of on letters.
    ' PutAfter: Word's paste (paste at cursor; cursor to end of pasted text)
    ' PutBefore: paste at cursor; leave cursor at start of pasted text
    ' PutAfterG and PutBeforeG: same, but paste as unformatted text.
    vcPutAfter     ' p
    vcPutBefore    ' P
    vcPutAfterG     'gp
    vcPutBeforeG    'gP

    ' TODO gp, gP, ]p, [p, ]P, [P

    'vcUndo         ' u
    'vcRedo         ' Ctl+R
    'vcUndoLine     ' U     No plans to implement this - I don't think Word gives me the necessary control.

    ' TODO z., zt, zb, z+, z-

    ' TODO /, ?
    'vcSearchNext    ' n
    'vcSearchPrev    ' N

    ' *, #, g*, g#          Search based on Word's idea of a word
    ' gW*, gW#, gWg*, gWg#  Search based on a WORD (not in Vim).
    '                       In Vim, gW is unused.  I am using it by analogy with W.
    ' However, sometimes gW* === gWg*, and gW# === gWg#, because Word
    ' doesn't respect SearchWholeWord if the search text has multiple
    ' words as defined by Word.

    vcSearchWholeItemForward    ' *, gW*
    vcSearchWholeItemBackward   ' #, gW#
    vcSearchItemForward         ' g*, gWg*
    vcSearchItemBackward        ' g#, gWg#

    ' TODO J, gJ
End Enum 'VimCommand

Public Enum VimOperator
    voUndef
    voGo            ' Placeholder for motions, which don't have an operator.

    'voReplaceChar  ' r (more like an operator than anything else) (maybe?)
    ' TODO m, ', ` here?

    ' Operators from :help operator
    voChange        ' c
    voDelete        ' d
    voYank          ' y
    voSelect        ' v Select <motion>.

    'voSwitchCase    ' ~/g~ unimpl
    ' TODO Maybe a custom titlecase on g~?
    'voLowercase     ' gu unimpl
    'voUppercase     ' gU unimpl

    'voFilter       ' ! No plans to implement this.
    'voEqualPrg     ' = No plans to implement this.
    'voFormatLines  ' gq No plans to implement this.
    'voFormatLinesNoMove    ' gw No plans to implement this.
    'voROT13        ' g? No plans to implement this.
    'voSHR          ' > No plans to implement this.
    'voSHL          ' < No plans to implement this.
    'voDefineFold   ' zf No plans to implement this.
    'voCallFunc     ' g@ No plans to implement this.

End Enum 'VimOperator

'Public Enum VimForce    ' adverb - No plans to implement this
'    vfUndef
'    vfCharacterwise
'    vfLinewise
'End Enum 'VimForce

Public Enum VimMotion   ' Motions/objects/direct objects of transitive operators
    vmUndef

    ' Marks - TODO the rest of :help mark-motions
    'vmMark             ' '
    'vmMarkExact        ' `

    ' Searches
    'vmSearchDown       ' /
    'vmSearchUp         ' ?
    'vmSearchNext       ' n
    'vmSearchPrev       ' N

    ' Objects from :help visual-operators
    vmAWord             ' aw
    vmIWord             ' iw
    vmANonblank         ' aW
    vmINonblank         ' iW
    vmASentence         ' as
    vmISentence         ' is
    vmAPara             ' ap (includes Chr(13))
    vmIPara             ' ip (not Chr(13))
    ' ab/ib () Not yet implemented
    ' aB/iB {} "
    ' at/it    " <tag></tag>
    ' a</i< <> "
    ' a[/i[ [] "
    ' a"/i"    "  Double-quoted string
    ' a'/i'    "  Single-quoted string
    ' a`/i`    "  Backtick-quoted string

    ' Motions from :help left-right-motions
    vmLeft              ' h
    vmRight             ' l
    vmStartOfParagraph  ' 0     Note: using slightly differently than in Vim.
    vmStartOfLine       ' ^
    vmEOL               ' $     Note: can take a count, in which case it goes count-1 lines downward inclusive
    vmEOParagraph       ' g$    Note: using slightly differently than in Vim.  Can take a count.

    ' g_, g0, g^, gm: Not yet implemented.
    vmColumn            ' |
    vmCharForward       ' f
    vmCharBackward      ' F
    vmTilForward        ' t
    vmTilBackward       ' T
    ' TODO? ; ,

    ' :help up-down-motions
    vmUp                ' k
    vmDown              ' j
    ' gk, gj, -, +: Not yet implemented
    vmUnderscore        ' for dd, yy, cc === d_, y_, c_
    vmLine              ' G ' TODO decide how to implement this; implement goto para/page
    ' gg (G, but not a jump) unimpl
    ' % (goto percentage), go (byte offset) unimpl

    ' :help word-motions
    vmWordForward           ' w exclusive
        ' TODO: operator+w and last word is at end of line => stop at end of word
    vmNonblankForward       ' W excl

    vmEOWordForward         ' e inclusive
    vmEONonblankForward     ' E incl

    vmWordBackward          ' b excl
    vmNonblankBackward      ' B excl

    vmEOWordBackward        ' ge incl
    vmEONonblankBackward    ' gE incl

    ' :help object-motions
    vmSentenceForward   ' )
    vmSentenceBackward  ' (
    vmParaForward       ' }
    vmParaBackward      ' {
    vmSectionForward    ' ]] or ][ (NOTE: not distinguishing the two)
    vmSectionBackward   ' [[ or [] (NOTE: not distinguishing the two)

    ' :help various-motions
    'vmPercent          ' %
    'vmPrevParen        ' [(
    'vmPrevBrace        ' [{
    'vmNextParen        ' ])
    'vmNextBrace        ' ]}
    ' ]m ]M [m [M next/prev start/END of method
    ' [# prev #if/#else
    ' ]# next #else/#endif
    ' [* [/ ]( ]/ C comment jumps

    ' H, M, L - move within currently visible text.  Count is line number for H and L.
    vmScreenTop
    vmScreenMiddle
    vmScreenBottom

    ' Custom (not in Vim) (TODO)
    'vmRevisionForward
    'vmRevisionBackward
    'vmARevision
    'vmIRevision

End Enum 'VimMotion
'

Public Enum VimNinja    ' https://github.com/tommcdo/vim-ninja-feet
    vnUndef
    vnLeft  ' ninja-feet [
    vnRight ' ninja-feet ]
End Enum 'VimNinja
'

' TODO: cw/cW is like ce/cE if cursor is on a non-blank

' TODO implement marks
' TODO implement jumps

Public WasCancelled As Boolean
Public Keys As String
Public VOperator As VimOperator
Public VCommand As VimCommand
Public VMotion As VimMotion
Public VOperatorCount As Long
Public VMotionCount As Long
Public VArg As String
Public VNinja As VimNinja

' Regexp
Private RE_ACT As VBScript_RegExp_55.RegExp

' Submatch numbers - see vim-regex.txt
Private RESM_COUNT1 As Long
Private RESM_IVERB As Long
Private RESM_IMOTION As Long
Private RESM_ITEXT As Long
Private RESM_TVERB As Long
Private RESM_COUNT2 As Long
Private RESM_TARGET As Long
Private RESM_TOBJ_RANGE As Long
Private RESM_NINJA As Long
Private RESM_OBJTYPE As Long
Private RESM_TTEXT As Long
Private RE_PAT As String
'

Private Sub UserForm_Initialize()
    WasCancelled = False
    Keys = ""
    VOperator = voUndef
    VCommand = vcUndef
    VMotion = vmUndef
    VOperatorCount = 1
    VMotionCount = 1
    VArg = ""
    VNinja = vnUndef

    Dim RE_PAT As String

    ' === Build up the regex ===
    ' The following code is from the output of `re2vba.pl --nodim vim-regex.txt`.
    ' DO NOT MODIFY HERE.  If you need to change it, modify vim-regex.txt
    ' and re-run re2vba.pl.

    RE_PAT = _
        "^(([1-9][0-9]*)?((([HMLGhjklwbWB\x28\x29\x7b\x7d]|g?[eE0\^\$" & _
        "]|[fFtT](.))|(gW)?g?[\*#]|g?[pP])|([cdyv])([1-9][0-9]*)?(([\" & _
        "[\]])?([ai])([wWsp])|[fFtT](.)|[HMLGhjklwbWB\x28\x29\x7b\x7d" & _
        "]|g?[eE0\^\$])))$" & _
        ""
    RESM_COUNT1 = 1
    RESM_IVERB = 3
    RESM_IMOTION = 4
    RESM_ITEXT = 5
    RESM_TVERB = 7
    RESM_COUNT2 = 8
    RESM_TARGET = 9
    RESM_NINJA = 10
    RESM_TOBJ_RANGE = 11
    RESM_OBJTYPE = 12
    RESM_TTEXT = 13

    ' === End of generated code ===

    Set RE_ACT = New VBScript_RegExp_55.RegExp
    RE_ACT.IgnoreCase = False
    RE_ACT.Pattern = RE_PAT

End Sub 'UserForm_Initialize
'

Private Sub btnCancel_Click()
    WasCancelled = True
    Me.Hide
End Sub

Private Sub UserForm_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If KeyAscii = vbKeyReturn Then
        Keys = Keys & Chr(13)
        Update
        Me.Hide
    ElseIf KeyAscii = vbKeyBack Then
        If Len(Keys) > 0 Then
            Keys = Left(Keys, Len(Keys) - 1)
            Update
        End If
    ElseIf KeyAscii >= 32 And KeyAscii <= 127 Then
        Keys = Keys & Chr(KeyAscii)
        Update
    End If
End Sub 'KeyPress

' Main workers ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Private Function ProcessMotion_(motion As String) As Boolean
' Return true iff #motion is an understood no-argument motion.
' Sets VMotion on success.  Does not handle text objects or motions
' with arguments (fFtT).

    ProcessMotion_ = True

    Select Case motion
        Case "H": VMotion = vmScreenTop
        Case "M": VMotion = vmScreenMiddle
        Case "L": VMotion = vmScreenBottom
        Case "G": VMotion = vmLine

        Case "h": VMotion = vmLeft
        Case "j": VMotion = vmDown
        Case "k": VMotion = vmUp
        Case "l": VMotion = vmRight

        Case "w": VMotion = vmWordForward
        Case "b": VMotion = vmWordBackward
        Case "W": VMotion = vmNonblankForward
        Case "B": VMotion = vmNonblankBackward

        Case ")": VMotion = vmSentenceForward
        Case "(": VMotion = vmSentenceBackward
        Case "}": VMotion = vmParaForward
        Case "{": VMotion = vmParaBackward

        Case "e": VMotion = vmEOWordForward
        Case "E": VMotion = vmEONonblankForward
        Case "ge": VMotion = vmEOWordBackward
        Case "gE": VMotion = vmEONonblankBackward

        Case "0": VMotion = vmStartOfParagraph
        Case "^": VMotion = vmStartOfLine
        Case "$": VMotion = vmEOL

        'Case "g0": VMotion = vmStartOfParagraph     ' TODO decide what this should do
        'Case "g^": VMotion = vmStartOfLine          ' TODO decide what this should do
        Case "g$": VMotion = vmEOParagraph

        Case Else: ProcessMotion_ = False
    End Select
End Function 'ProcessMotion_
'

Private Function ProcessHit_(hit As VBScript_RegExp_55.Match) As Boolean
' Returns true if parse succeeded
    ProcessHit_ = False

    ' Consistent values at the start
    VOperator = voUndef
    VCommand = vcUndef
    VMotion = vmUndef
    VOperatorCount = 1
    VMotionCount = 1
    VArg = ""
    VNinja = vnUndef

    ' Special-case "0" in code so that I don't have to special-case it in the
    ' regex.  A non-empty count preceding a "0" command means that the "0"
    ' should actually be part of the count, so wait for more keys.

    If (Not IsEmpty(hit.SubMatches(RESM_COUNT1))) And _
        (hit.SubMatches(RESM_IMOTION) = "0") _
    Then        ' ^ Empty decays to ""
        Exit Function
    End If

    ' Count before the command, if any
    If Len(hit.SubMatches(RESM_COUNT1)) = 0 Then
        ' Note: works because Len(Empty) returns 0
        VOperatorCount = 1
    Else
        VOperatorCount = CLng(hit.SubMatches(RESM_COUNT1))
    End If

    If Not IsEmpty(hit.SubMatches(RESM_IVERB)) Then     ' intransitive

        'Debug.Print "Intransit.", IIf(IsEmpty(hit.SubMatches(RESM_IMOTION)), "-", hit.SubMatches(RESM_IMOTION)), _
        '                            hit.SubMatches(RESM_IVERB), hit.SubMatches(RESM_ITEXT)

        If Not IsEmpty(hit.SubMatches(RESM_IMOTION)) Then
            If ProcessMotion_(hit.SubMatches(RESM_IMOTION)) Then
                VOperator = voGo
            Else
                Select Case Left(hit.SubMatches(RESM_IVERB), 1)
                    Case "f": VOperator = voGo: VMotion = vmCharForward: VArg = hit.SubMatches(RESM_ITEXT)
                    Case "F": VOperator = voGo: VMotion = vmCharBackward: VArg = hit.SubMatches(RESM_ITEXT)
                    Case "t": VOperator = voGo: VMotion = vmTilForward: VArg = hit.SubMatches(RESM_ITEXT)
                    Case "T": VOperator = voGo: VMotion = vmTilBackward: VArg = hit.SubMatches(RESM_ITEXT)
                    Case Else: Exit Function
                End Select
            End If

        Else    ' Not a motion, so it's a search or a paste

            Select Case hit.SubMatches(RESM_IVERB)
                ' Searches
                Case "*": VCommand = vcSearchWholeItemForward: VMotion = vmIWord
                Case "#": VCommand = vcSearchWholeItemBackward: VMotion = vmIWord
                Case "g*": VCommand = vcSearchItemForward: VMotion = vmIWord
                Case "g#": VCommand = vcSearchItemBackward: VMotion = vmIWord

                Case "gW*": VCommand = vcSearchWholeItemForward: VMotion = vmINonblank
                Case "gW#": VCommand = vcSearchWholeItemBackward: VMotion = vmINonblank
                Case "gWg*": VCommand = vcSearchItemForward: VMotion = vmINonblank
                Case "gWg#": VCommand = vcSearchItemBackward: VMotion = vmINonblank

                ' Pastes
                Case "p": VCommand = vcPutAfter
                Case "P": VCommand = vcPutBefore
                Case "gp": VCommand = vcPutAfterG
                Case "gP": VCommand = vcPutBeforeG

                Case Else: Exit Function
            End Select
        End If 'a motion else

    ElseIf Not IsEmpty(hit.SubMatches(RESM_TVERB)) Then     ' transitive

        'Debug.Print "Transitive", hit.SubMatches(RESM_TVERB), hit.SubMatches(RESM_COUNT2), Left(hit.SubMatches(RESM_TARGET), 1), hit.SubMatches(RESM_OBJTYPE), hit.SubMatches(RESM_TTEXT)

        ' Operator
        Select Case hit.SubMatches(RESM_TVERB)
            Case "c": VOperator = voChange
            Case "d": VOperator = voDelete
            Case "y": VOperator = voYank
            Case "v": VOperator = voSelect      ' V, i.e., visual selection - just like in Vim.
            Case Else: Exit Function
        End Select

        ' Post-operator count
        If Len(hit.SubMatches(RESM_COUNT2)) = 0 Then
            VMotionCount = 1
        Else
            VMotionCount = CLng(hit.SubMatches(RESM_COUNT2))
        End If

        ' Process the motion
        If ProcessMotion_(hit.SubMatches(RESM_TARGET)) Then         ' Motion without argument
            ' Nothing more to do

        ElseIf Not IsEmpty(hit.SubMatches(RESM_TOBJ_RANGE)) Then    ' Text object
            Select Case hit.SubMatches(RESM_TOBJ_RANGE)
                Case "a":
                    Select Case hit.SubMatches(RESM_OBJTYPE)
                        Case "w": VMotion = vmAWord
                        Case "W": VMotion = vmANonblank
                        Case "s": VMotion = vmASentence
                        Case "p": VMotion = vmAPara
                        Case Else: Exit Function
                    End Select

                Case "i":
                    Select Case hit.SubMatches(RESM_OBJTYPE)
                        Case "w": VMotion = vmIWord
                        Case "W": VMotion = vmINonblank
                        Case "s": VMotion = vmISentence
                        Case "p": VMotion = vmIPara
                        Case Else: Exit Function
                    End Select

                Case Else: Exit Function
            End Select

            If Not IsEmpty(hit.SubMatches(RESM_NINJA)) Then
                VNinja = IIf(hit.SubMatches(RESM_NINJA) = "[", vnLeft, vnRight)
            End If

        Else ' Not a text object, so it's a motion with argument
            Select Case Left(hit.SubMatches(RESM_TARGET), 1)
                Case "f": VMotion = vmCharForward: VArg = hit.SubMatches(RESM_TTEXT)
                Case "F": VMotion = vmCharBackward: VArg = hit.SubMatches(RESM_TTEXT)
                Case "t": VMotion = vmTilForward: VArg = hit.SubMatches(RESM_TTEXT)
                Case "T": VMotion = vmTilBackward: VArg = hit.SubMatches(RESM_TTEXT)

                Case Else: Exit Function
            End Select
        End If

    Else    ' Neither intransitive nor transitive
        MsgBox "This shouldn't happen - no iverb or tverb.  Send a screenshot of this box to Chris White: -" & hit.Value & "-", vbExclamation, "VimWord"
        Exit Function

    End If 'intransitive else transitive else

    ProcessHit_ = True    ' If we made it here, the parse was successful

End Function ' ProcessHit_
'

Private Sub Update()
    Dim done As Boolean: done = False
    lblKeys.Caption = Keys

    ' parse Vim commands to see if one is done
    Dim matches As VBScript_RegExp_55.MatchCollection
    Dim hit As VBScript_RegExp_55.Match

    Set matches = RE_ACT.Execute(Keys)
    If matches.count > 0 Then
        Do 'Once
            Set hit = matches.Item(0)
            If hit.SubMatches.count < 1 Then Exit Do
            Debug.Print "Matched:", hit.Value

            done = ProcessHit_(hit)
            'If done Then Debug.Print "", "operator count:", VOperatorCount
        Loop While False
    End If

    If done Then Me.Hide
End Sub 'Update

