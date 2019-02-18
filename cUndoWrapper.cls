VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cUndoWrapper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' cUndoWrapper: wrapper for custom Undo routines
' Copyright (c) 2016--2019 Chris White
'   2016/08/05  chrisw  Initial version
'   2019-01-31  chrisw  Added am_i_recording_
'   2019-02-14  chrisw  Added conditional-compilation guard for
'                       Word 2007 compatibility.  On 2007,
'                       this class is a no-op.
'   2019-02-18  chrisw  nothrow in Class_Terminate; bugfix in Start

Option Explicit
Option Base 0

Dim undos_ As Object            ' an UndoRecord, late-bound
Dim am_i_recording_ As Boolean  ' Safety check in case other code starts recording
'

Private Sub Class_Initialize()
    If Application.Version >= 14 Then   ' 2010+
        Dim o As Object             ' A late-bound interface to Application
        Set o = Application
        Set undos_ = o.UndoRecord   ' Late-bound call
    Else
        Set undos_ = Nothing
    End If
    am_i_recording_ = False
End Sub

Public Sub Start(TITLE As String)
    If Not (undos_ Is Nothing) Then
        undos_.StartCustomRecord TITLE
        am_i_recording_ = True
    End If
End Sub

Private Sub Class_Terminate()
    Dim was_i_recording As Boolean
    
    ' Reset am_i_recording_ first so we don't ever try to
    ' close the group twice (not sure how we would, but just
    ' to be safe...).
    was_i_recording = am_i_recording_
    am_i_recording_ = False
    
    On Error Resume Next
    If Not (undos_ Is Nothing) Then
        If was_i_recording And undos_.IsRecordingCustomRecord Then
            undos_.EndCustomRecord
        End If
    End If
End Sub

' vi: set ts=4 sts=4 sw=4 et ai: '

