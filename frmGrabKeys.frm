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

Option Explicit
Option Base 0

' Vim support
Public Enum VimOperator
    vcUndef
    vcChange        ' c
    vcDelete        ' d
    vcYank          ' y
    vcGo            ' ' - NOT IN VIM.  To use motions as cursor motions.
End Enum 'VimOperator

Public Enum VimMotion
    vmUndef
    'vmWholeLine         ' for dd, yy, cc.
    vmLeft              ' h
    vmRight             ' l
    vmUp                ' k
    vmDown              ' j
    vmStartOfLine       ' ^     ' TODO
    vmStartOfParagraph  ' 0     ' TODO this gets mixed in with the count
    vmEOL               ' $     ' TODO
    vmLine              ' G
    vmCharForward       ' f
    vmCharBackward      ' F
    vmTilForward        ' t
    vmTilBackward       ' T
    vmWordForward       ' w
    vmEOWordForward     ' e
    vmWordBackward      ' b
    vmNonblankForward   ' W
    vmEONonblankForward ' E
    vmNonblankBackward  ' B
    vmSentenceForward   ' )
    vmSentenceBackward  ' (
    vmParaForward       ' }
    vmParaBackward      ' {
    
    vmAWord             ' aw
    vmIWord             ' iw
    vmANonblank         ' aW
    vmINonblank         ' iW
    vmASentence         ' as
    vmISentence         ' is
    vmAPara             ' ap (includes Chr(13))
    vmIPara             ' ip (not Chr(13))
End Enum 'VimMotion
'

Public WasCancelled As Boolean
Public Keys As String
Public VCommand As VimOperator
Public VMotion As VimMotion
Public VCommandCount As Long
Public VMotionCount As Long
Public VArg As String

Private RE_EXCMD As VBScript_RegExp_55.RegExp

Private Sub Update()
    Dim done As Boolean: done = False
    lblKeys.Caption = Keys
    
    ' parse Vim commands to see if one is done
    Dim matches As VBScript_RegExp_55.MatchCollection
    Dim hit As VBScript_RegExp_55.Match
    
    Set matches = RE_EXCMD.Execute(Keys)
    If matches.count > 0 Then
        Set hit = matches.Item(0)
        On Error GoTo BadParse
        
        If Len(hit.submatches(0)) = 0 Then
            VCommandCount = 1
        Else
            VCommandCount = CLng(hit.submatches(0))
        End If
        
        If Len(hit.submatches(2)) = 0 Then
            VMotionCount = 1
        Else
            VMotionCount = CLng(hit.submatches(2))
        End If
        
        VArg = ""   ' empty unless assigned below (fFtT)
        
        Select Case hit.submatches(1)
            Case "c": VCommand = vcChange
            Case "d": VCommand = vcDelete
            Case "y": VCommand = vcYank
            Case "'", ";": VCommand = vcGo
            Case Else: Err.Raise vbObjectError
        End Select
        
        Select Case Left(hit.submatches(3), 1)
            Case "h": VMotion = vmLeft
            Case "l": VMotion = vmRight
            Case "k": VMotion = vmUp
            Case "j": VMotion = vmDown
            Case "^": VMotion = vmStartOfLine
            Case "0": VMotion = vmStartOfParagraph
            Case "$": VMotion = vmEOL
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
            
            Case "f", "F", "t", "T":
                VArg = Mid(hit.submatches(3), 2, 1)
                Select Case Left(hit.submatches(3), 1)
                    Case "f": VMotion = vmCharForward
                    Case "F": VMotion = vmCharBackward
                    Case "t": VMotion = vmTilForward
                    Case "T": VMotion = vmTilBackward
                End Select
                
            Case "a", "i":
                Select Case hit.submatches(3)
                    Case "aw": VMotion = vmAWord
                    Case "iw": VMotion = vmIWord
                    Case "aW": VMotion = vmANonblank
                    Case "iW": VMotion = vmINonblank
                    Case "as": VMotion = vmASentence
                    Case "is": VMotion = vmISentence
                    Case "ap": VMotion = vmAPara
                    Case "ip": VMotion = vmIPara
                End Select
                
            Case Else: Err.Raise vbObjectError
        End Select
         
        done = True     ' If we made it here, the parse was successful
    End If
    
Update_Finally:
    On Error Resume Next
    If done Then
        Me.Hide
    End If
    Exit Sub
BadParse:
    done = False
    Resume Update_Finally
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Private Sub UserForm_Initialize()
    WasCancelled = False
    Keys = ""
    VCommand = vcUndef
    VMotion = vmUndef
    VCommandCount = 1
    VMotionCount = 1
    VArg = ""
    
    Set RE_EXCMD = New VBScript_RegExp_55.RegExp
    RE_EXCMD.pattern = "^([0-9]*)([dcy;'])([0-9]*)([ai][wWsp]|[fFtT].|[hjklGwebWEB\)\(\}\{])$"
        ' TODO figure out ^0$
End Sub

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
End Sub
