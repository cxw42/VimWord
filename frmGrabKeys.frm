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
' frmGrabKeys: collect keypresses.
' Copyright (c) 2018 Chris White.  All rights reserved.
'   2018-04-06  chrisw  Initial version
'   2018-04-20  chrisw  Major expansion/rewrite

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
    
    ' TODO implement these once I implement registers
    'vcPutAfter     ' p
    'vcPutBefore    ' P
    ' TODO gp, gP, ]p, [p, ]P, [P
    
    'vcUndo         ' u
    'vcRedo         ' Ctl+R
    'vcUndoLine     ' U
    
    ' TODO z., zt, zb, z+, z-
    
    ' TODO /, ?
    'vcSearchNext    ' n
    'vcSearchPrev    ' N
    
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
    
    ' Custom (not in Vim)
    voSelect        ' s Select <motion>.  Mostly for use as a debugging aid.
End Enum 'VimOperator

'Public Enum VimForce    ' adverb - No plans to implement this
'    vfUndef
'    vfCharacterwise
'    vfLinewise
'End Enum 'VimForce

Public Enum VimMotion   ' Motions/objects/direct objects of transitive operators
    vmUndef
    
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
    vmEOL               ' $
    ' g_, g0, g^, gm, g$: Not yet implemented.
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
Public VMotion As VimMotion
Public VOperatorCount As Long
Public VMotionCount As Long
Public VArg As String

' Regexp
Private RE_ACT As VBScript_RegExp_55.RegExp
' Submatch numbers
'Private RESM_REGISTER As Long  ' Not yet implemented
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
    VOperator = vcUndef
    VMotion = vmUndef
    VOperatorCount = 1
    VMotionCount = 1
    VArg = ""
    
    Dim PC As Long: PC = 0      ' Paren counter
    Dim PC_INTRANS As Long, PC_TRANS As Long
    
    Set RE_ACT = New VBScript_RegExp_55.RegExp
    
    Dim PAT_INTRANS As String   ' Pattern for an intransitive sentence, after goal and quantifier
        ' Not yet impl
    Dim PAT_TRANS As String     ' Pattern for a transitive sentence, after goal and quantifier
        ' TODO figure out ^0$
    
    ' Note: /^"./ (register/goal) not yet implemented
    
    ' === Build up the regex ===
    ' We do this a piece at a time to make it easier to change later.
    ' Also, this prevents you from going insane trying to manually track
    ' submatch numbers between pieces of the regex.
    
    PAT_INTRANS = _
                    "([0\^$wWeEbB]|[fFtT](.))"     ' includes motions
    '                |                 |
    RESM_IVERB = 0 '-^                 |
    RESM_ITEXT = 1 ' ------------------'
    PC_INTRANS = 2
    
    PAT_TRANS = _
                    "([cdys])?([1-9][0-9]*)?([ai]([wWsp])|[fFtT](.)|[hjklGwebWEB\)\(\}\{])"
    '                |        |             |    |              |
    RESM_TVERB = 0 '-^        |             |    |              |
    RESM_COUNT2 = 1 ' --------'             |    |              |
    RESM_TOBJ = 2   ' ----------------------'    |              |
    RESM_OBJTYPE = 3    ' -----------------------'              |
    RESM_TTEXT = 4  ' ------------------------------------------'
    PC_TRANS = 5
    
    RE_ACT.Pattern = "^([1-9][0-9]*)?(" & _
                                "(" & PAT_INTRANS & ")" & _
                                "|" & _
                                "(" & PAT_TRANS & ")" & _
                            ")$"
    'RESM_REGISTER     |     not yet implemented
    '      RESM_COUNT1-^         |
    RESM_COUNT1 = 0     '        |565
    ' Grouping parens numbered --' but we don't use them
    
    ' Get submatch numbers for intransitive
    PC = 3  ' RESM_COUNT1, and two following open parens
    RESM_IVERB = RESM_IVERB + PC
    RESM_ITEXT = RESM_ITEXT + PC
    
    ' Get submatch numbers for transitive
    PC = PC + PC_INTRANS
    PC = PC + 1     ' open paren before PAT_TRANS
    RESM_TVERB = RESM_TVERB + PC
    RESM_COUNT2 = RESM_COUNT2 + PC
    RESM_TOBJ = RESM_TOBJ + PC
    RESM_OBJTYPE = RESM_OBJTYPE + PC
    RESM_TTEXT = RESM_TTEXT + PC
    PC = PC + PC_TRANS
    
End Sub 'UserForm_Initialize
'

Private Sub btnCancel_Click()
    WasCancelled = True
    Me.Hide
End Sub

Private Sub UserForm_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If KeyAscii = vbKeyReturn Then
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
    
    If IsEmpty(hit.SubMatches(RESM_TVERB)) Then    ' intransitive
    
        'Debug.Print "Intransit.", Left(hit.SubMatches(RESM_IVERB), 1), hit.SubMatches(RESM_ITEXT)
        
        Select Case Left(hit.SubMatches(RESM_IVERB), 1)
            Case "0": VOperator = voGo: VMotion = vmStartOfParagraph
            Case "^": VOperator = voGo: VMotion = vmStartOfLine
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
            
            Case Else: Exit Function
        End Select
    
    Else                                            ' transitive
    
        'Debug.Print "Transitive", hit.SubMatches(RESM_TVERB), hit.SubMatches(RESM_COUNT2), Left(hit.SubMatches(RESM_TOBJ), 1), hit.SubMatches(RESM_OBJTYPE), hit.SubMatches(RESM_TTEXT)
        
        Select Case hit.SubMatches(RESM_TVERB)
            Case "c": VOperator = voChange
            Case "d": VOperator = voDelete
            Case "y": VOperator = voYank
            Case "s": VOperator = voSelect
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

