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
'   2018-04-06  chrisw  Initial version
'   2018-04-20  chrisw  Major expansion/rewrite
'   2018-04-24  chrisw  Change "s" to "v" (visual selection)
'   2018-04-26  chrisw  Changed regex from hand-generated to re2vba.pl

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

    vcundef

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

    ' TODO implement these once I implement registers
    'vcPutAfter     ' p
    'vcPutBefore    ' P
    ' TODO gp, gP, ]p, [p, ]P, [P

    'vcUndo         ' u
    'vcRedo         ' Ctl+R
    'vcUndoLine     ' U     No plans to implement this - I don't think Word gives me the necessary control.

    ' TODO z., zt, zb, z+, z-

    ' TODO /, ?, *, #, g*, g# (search based on Word's idea of a word)
    'vcSearchNext    ' n
    'vcSearchPrev    ' N
    vcSearchWholeWordForward    ' *
    vcSearchWholeWordBackward   ' #
    vcSearchWordForward         ' g*
    vcSearchWordBackward        ' g#


    ' TODO gW*, gW#, gWg*, gWg# to search based on a WORD (not in Vim).
    ' In Vim, gW is unused.  This is by analogy with W.
    'vcSearchNonblankForward    ' z*
    vcSearchWholeNonblankForward    ' gW*
    vcSearchWholeNonblankBackward   ' gW#
    vcSearchNonblankForward         ' gWg*
    vcSearchNonblankBackward        ' gWg#

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

    'vmEOWordBackward        ' ge incl not yet impl
    'vmEONonblankBackward    ' gE incl "

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

    ' Custom (not in Vim) (TODO)
    'vmRevisionForward
    'vmRevisionBackward
    'vmARevision
    'vmIRevision

End Enum 'VimMotion
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

' Regexp
Private RE_ACT As VBScript_RegExp_55.RegExp
' Submatch numbers

'Private RESM_REGISTER As Long  ' Not yet implemented

' For no-count sentences
Private RESM_NOCOUNT As Long

Private RESM_COUNT1 As Long

' For intransitive sentences
Private RESM_IVERB As Long
Private RESM_ITEXT As Long      ' Text for motions such as t and f

' For transitive sentences
Private RESM_TVERB As Long      ' a transitive verb
Private RESM_COUNT2 As Long     ' Only for transitive sentences
Private RESM_TOBJ As Long       ' " - a motion or text object
Private RESM_OBJTYPE As Long    ' When selecting text objects, which kind of object
Private RESM_TTEXT As Long      ' Text for motions such as t and f
'

Private Sub UserForm_Initialize()
    WasCancelled = False
    Keys = ""
    VOperator = voUndef
    VCommand = vcundef
    VMotion = vmUndef
    VOperatorCount = 1
    VMotionCount = 1
    VArg = ""

    Dim RE_PAT As String

    ' === Build up the regex ===
    ' The following code is from the output of `re2vba.pl --nodim vim-regex.txt`.
    ' DO NOT MODIFY HERE.  If you need to change it, modify vim-regex.txt
    ' and re-run re2vba.pl.

    RE_PAT = _
        "^(([0\^])|(([1-9][0-9]*)?(([$wWeEbB]|g\$|[fFtT](.)|\*|[#]|g\" & _
        "*|g#|gW\*|gW#|gWg\*|gWg#)|([cdyv])?([1-9][0-9]*)?([ai]([wWsp" & _
        "])|[fFtT](.)|[hjklGwebWEB\x28\x29\x7b\x7d]))))$" & _
        ""
    RESM_NOCOUNT = 1
    RESM_COUNT1 = 3
    RESM_IVERB = 5
    RESM_ITEXT = 6
    RESM_TVERB = 7
    RESM_COUNT2 = 8
    RESM_TOBJ = 9
    RESM_OBJTYPE = 10
    RESM_TTEXT = 11

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

Private Function ProcessHit_(hit As VBScript_RegExp_55.Match) As Boolean
' Returns true if parse succeeded
    ProcessHit_ = False

    If Not IsEmpty(hit.SubMatches(RESM_NOCOUNT)) Then   ' no-count
        Select Case Left(hit.SubMatches(RESM_NOCOUNT), 1)
            Case "0": VOperator = voGo: VMotion = vmStartOfParagraph
            Case "^": VOperator = voGo: VMotion = vmStartOfLine
            Case Else: Exit Function
        End Select

    ElseIf IsEmpty(hit.SubMatches(RESM_TVERB)) Then    ' intransitive

        'Debug.Print "Intransit.", Left(hit.SubMatches(RESM_IVERB), 1), hit.SubMatches(RESM_ITEXT)

        Select Case Left(hit.SubMatches(RESM_IVERB), 1)
            Case "$": VOperator = voGo: VMotion = vmEOL
            Case "w": VOperator = voGo: VMotion = vmWordForward
            Case "W": VOperator = voGo: VMotion = vmNonblankForward
            Case "e": VOperator = voGo: VMotion = vmEOWordForward
            Case "E": VOperator = voGo: VMotion = vmEONonblankForward
            Case "b": VOperator = voGo: VMotion = vmWordBackward
            Case "B": VOperator = voGo: VMotion = vmNonblankBackward

            Case "f": VOperator = voGo: VMotion = vmCharForward: VArg = hit.SubMatches(RESM_ITEXT)
            Case "F": VOperator = voGo: VMotion = vmCharBackward: VArg = hit.SubMatches(RESM_ITEXT)
            Case "t": VOperator = voGo: VMotion = vmTilForward: VArg = hit.SubMatches(RESM_ITEXT)
            Case "T": VOperator = voGo: VMotion = vmTilBackward: VArg = hit.SubMatches(RESM_ITEXT)

            Case Else:  ' Check the whole iverb, since the first character isn't enough
                Select Case hit.SubMatches(RESM_IVERB)
                    ' Motions
                    Case "g$": VOperator = voGo: VMotion = vmEOParagraph
                    
                    ' Searches
                    Case "*": VCommand = vcSearchWholeWordForward: VMotion = vmIWord
                    Case "#": VCommand = vcSearchWholeWordBackward: VMotion = vmIWord
                    Case "g*": VCommand = vcSearchWordForward: VMotion = vmIWord
                    Case "g#": VCommand = vcSearchWordBackward: VMotion = vmIWord
                    
                    Case "gW*": VCommand = vcSearchWholeNonblankForward: VMotion = vmINonblank
                    Case "gW#": VCommand = vcSearchWholeNonblankBackward: VMotion = vmINonblank
                    Case "gWg*": VCommand = vcSearchNonblankForward: VMotion = vmINonblank
                    Case "gWg#": VCommand = vcSearchNonblankBackward: VMotion = vmINonblank
                    
                    Case Else: Exit Function
                End Select
        End Select

    Else                                            ' transitive

        'Debug.Print "Transitive", hit.SubMatches(RESM_TVERB), hit.SubMatches(RESM_COUNT2), Left(hit.SubMatches(RESM_TOBJ), 1), hit.SubMatches(RESM_OBJTYPE), hit.SubMatches(RESM_TTEXT)

        Select Case hit.SubMatches(RESM_TVERB)
            Case "c": VOperator = voChange
            Case "d": VOperator = voDelete
            Case "y": VOperator = voYank
            Case "v": VOperator = voSelect      ' V, i.e., visual selection - just like in Vim.
            Case Else: Exit Function
        End Select

        If Len(hit.SubMatches(RESM_COUNT2)) = 0 Then
            VMotionCount = 1
        Else
            VMotionCount = CLng(hit.SubMatches(RESM_COUNT2))
        End If

        Select Case Left(hit.SubMatches(RESM_TOBJ), 1)
            Case "a":
                Select Case hit.SubMatches(RESM_OBJTYPE)
                    Case "w": VMotion = vmAWord
                    Case "W": VMotion = vmANonblank
                    Case "s": VMotion = vmASentence
                    Case "p": VMotion = vmAPara
                End Select

            Case "i":
                Select Case hit.SubMatches(RESM_OBJTYPE)
                    Case "w": VMotion = vmIWord
                    Case "W": VMotion = vmINonblank
                    Case "s": VMotion = vmISentence
                    Case "p": VMotion = vmIPara
                End Select

            Case "f": VMotion = vmCharForward: VArg = hit.SubMatches(RESM_TTEXT)
            Case "F": VMotion = vmCharBackward: VArg = hit.SubMatches(RESM_TTEXT)
            Case "t": VMotion = vmTilForward: VArg = hit.SubMatches(RESM_TTEXT)
            Case "T": VMotion = vmTilBackward: VArg = hit.SubMatches(RESM_TTEXT)

            Case "h": VMotion = vmLeft
            Case "j": VMotion = vmDown
            Case "k": VMotion = vmUp
            Case "l": VMotion = vmRight
            Case "G": VMotion = vmLine

            Case "w": VMotion = vmWordForward
            Case "e": VMotion = vmEOWordForward
            Case "b": VMotion = vmWordBackward
            Case "W": VMotion = vmNonblankForward
            Case "E": VMotion = vmEONonblankForward
            Case "B": VMotion = vmNonblankBackward
            Case ")": VMotion = vmSentenceForward
            Case "(": VMotion = vmSentenceBackward
            Case "}": VMotion = vmParaForward
            Case "{": VMotion = vmParaBackward

            Case Else: Exit Function
        End Select

    End If ' Intransitive else

    ProcessHit_ = True    ' If we made it here, the parse was successful

End Function ' ProcessHit_
'

Private Sub Update()
    Dim done As Boolean: done = False
    lblKeys.Caption = Keys

    VArg = ""   ' empty unless assigned below (fFtT)

    ' parse Vim commands to see if one is done
    Dim matches As VBScript_RegExp_55.MatchCollection
    Dim hit As VBScript_RegExp_55.Match

    Set matches = RE_ACT.Execute(Keys)
    If matches.count > 0 Then
        Do 'Once
            Set hit = matches.Item(0)
            If hit.SubMatches.count < 1 Then Exit Do

            If Len(hit.SubMatches(RESM_COUNT1)) = 0 Then
                ' Note: works because Len(Empty) returns 0
                VOperatorCount = 1
            Else
                VOperatorCount = CLng(hit.SubMatches(RESM_COUNT1))
            End If

            done = ProcessHit_(hit)

            'If done Then Debug.Print "", "operator count:", VOperatorCount

        Loop While False
    End If

    If done Then Me.Hide
End Sub 'Update

