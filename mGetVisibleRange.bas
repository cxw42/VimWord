Attribute VB_Name = "mGetVisibleRange"
' mWin32: Win32 routines
' Copyright (c) 2019 Chris White.  All rights reserved.
' History:
'   2019-03-05  chrisw  Initial version
'   2019-04-29  chrisw  Added GetCornerRange.  In GCR, added loop and
'                       test for wdMainTextStory.

Option Explicit
Option Base 0


' See http://msdn.microsoft.com/en-us/library/office/aa164901%28v=office.10%29.aspx

' --- Types ---------------------------------------------

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type 'RECT

Private Declare Function EnumWindows Lib "user32" (ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long
Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Private Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWndParent As Long, ByVal hWndChildAfter As Long, ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Private Declare Function EnumChildWindows Lib "user32" (ByVal hWndParent As Long, ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long
Private Declare Function SendMessageStr Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare Function GetForegroundWindow Lib "user32" () As Long

' --- Constants -----------------------------------------

' Identifying information
Private Const Word2007_Document_Class As String = "_WwG"
Private Const Word2007_Document_Title As String = "Microsoft Word Document"
Private Const Word2013_Document_Class As String = "_WwG"
Private Const Word2013_Document_Title As String = "Microsoft Word Document"

' *** Caution: I don't know if these are actually correct ***
Private Const Word2010_Document_Class As String = "_WwG"
Private Const Word2010_Document_Title As String = "Microsoft Word Document"
Private Const Word2016_Document_Class As String = "_WwG"
Private Const Word2016_Document_Title As String = "Microsoft Word Document"

' --- Private data --------------------------------------
Private gDocHwnd_ As Long   ' The EnumChildWindows callback's return value

' Static variables holding class and window names.  Used for Office multi-version compatibility.
Private gTargetDocumentClass As String
Private gTargetDocumentTitle As String
Private gTargetInitialized As Boolean   'defaults to False
'

' --- Get current viewport ------------------------------

Public Function GetVisibleRange(doc As Document) As Range
' Returns the range visible in the viewport of doc's active window, or Nothing.

    Dim TITLE As String: TITLE = "Get visible range"
    If doc Is Nothing Then Exit Function
    Dim win As Window: Set win = doc.ActiveWindow

    Set GetVisibleRange = Nothing
    Dim viewport_hwnd As Long
    Dim win_hwnd As Long
    Dim winrect As RECT

    InitDimensions  'in case this is the first time this function has been called

    If CLng(Application.Version) > 12 Then      ' Office 2007 doesn't have Window.HWnd
        Dim owin As Object
        Set owin = win
        win_hwnd = owin.hwnd
    Else
        doc.Activate
        Application.Activate
        win_hwnd = GetForegroundWindow
    End If
    
    'DEBUGFindDocument win_hwnd      ' Uncomment to print all child windows
    viewport_hwnd = FindDocument(win_hwnd)
    If viewport_hwnd = 0 Then Exit Function

    If GetWindowRect(viewport_hwnd, winrect) = 0 Then Exit Function

    ' Get the corners
    Dim ul As Range, lr As Range
    Set ul = GetCornerRange(win, winrect, True)
    Set lr = GetCornerRange(win, winrect, False)

    If (ul Is Nothing) Or (lr Is Nothing) Then
        Set GetVisibleRange = Nothing
    Else
        Set GetVisibleRange = doc.Range(ul.Start, lr.End)
    End If
End Function 'GetVisibleRange

' ===========================================================================
' Internals

Private Sub InitDimensions()
' Initialize the global variables holding class names.
    If gTargetInitialized Then Exit Sub

    If CLng(Application.Version) = 16 Then
        gTargetDocumentClass = Word2016_Document_Class
        gTargetDocumentTitle = Word2016_Document_Title
    ElseIf CLng(Application.Version) = 15 Then
        gTargetDocumentClass = Word2013_Document_Class
        gTargetDocumentTitle = Word2013_Document_Title
    ElseIf CLng(Application.Version) = 14 Then
        gTargetDocumentClass = Word2010_Document_Class
        gTargetDocumentTitle = Word2010_Document_Title
    ' v13 = nonexistent
    ElseIf CLng(Application.Version) = 12 Then
        gTargetDocumentClass = Word2007_Document_Class
        gTargetDocumentTitle = Word2007_Document_Title
    Else    ' unknown or unsupported version - let the methods fail but don't annoy the user otherwise.
        gTargetDocumentClass = ""
        gTargetDocumentTitle = ""
    End If
    gTargetInitialized = True
End Sub      'InitDimensions

' --- Find Document window ---------------------------------

Private Function FindDocument(tophwnd As Long) As Long
' Return the HWND of the document window under tophwnd, or 0 on failure
    gDocHwnd_ = 0
    EnumChildWindows tophwnd, AddressOf FindDocument_Callback, 0
    FindDocument = gDocHwnd_
End Function

Function FindDocument_Callback(ByVal hwnd As Long, ByVal lParam As Long) As Long
    Dim sClass As String
    Dim sTitle As String
    'Dim result As VbMsgBoxResult    'DEBUG

    FindDocument_Callback = 1  'continue unless we find it

    sClass = String(255, 0)
    GetClassName hwnd, sClass, 255

    If (gTargetDocumentClass = "") Or (InStr(sClass, gTargetDocumentClass) > 0) Then
        sTitle = String(255, 0)
        GetWindowText hwnd, sTitle, 255
        If (gTargetDocumentTitle = "") Or (InStr(sTitle, gTargetDocumentTitle) > 0) Then
            gDocHwnd_ = hwnd
            FindDocument_Callback = 0    'Done looping
        End If
    End If
End Function

Private Function GetCornerRange(win As Window, winrect As RECT, _
                                isUL As Boolean) As Range
' Get the corners.  Only looks in the wdMainTextStory.

    ' How many steps to check in the loop
    Dim NSTEPS As Long: NSTEPS = 10

    Dim retval As Range: Set retval = Nothing

    ' Set up for the loop.  Loop by Step while abs(Curr-Start)<Limit.
    Dim xStart As Long, yStart As Long, xLimit As Long, yLimit As Long
    Dim xCurr As Long, yCurr As Long, xStep As Long, yStep As Long

    xLimit = winrect.Right - winrect.Left
    yLimit = winrect.Bottom - winrect.Top

    ' +/-1s are an attempt to deal with the fact that
    ' scroll positions that are partial lines may lead to unexpected results,
    ' e.g., jumping to a line that used to be just off the top of the window.

    If isUL Then    ' Start at the UL; work right, then down.
        xStart = winrect.Left + 1
        xStep = xLimit / NSTEPS
        yStart = winrect.Top + 1
        yStep = yLimit / NSTEPS

    Else    ' LR
        xStart = winrect.Right - 1
        xStep = -xLimit / NSTEPS
        yStart = winrect.Bottom - 1
        yStep = -yLimit / NSTEPS

    End If

    ' The following loop is to deal with an odd case: if the last line of the
    ' window is a deleted paragraph with a non-deleted comment at the
    ' beginning, no range is present, so lr becomes Nothing.

    ' Go across rows, then up/down the page.
    yCurr = yStart
    Do While Abs(yCurr - yStart) < yLimit
        xCurr = xStart  ' Start the row

        Do While Abs(xCurr - xStart) < xLimit
            Set retval = win.RangeFromPoint(xCurr, yCurr)

            If Not (retval Is Nothing) Then
                If retval.StoryType = wdMainTextStory Then
                    GoTo GCR_Done   'Found it
                End If
            End If

            xCurr = xCurr + xStep
        Loop 'X

        yCurr = yCurr + yStep   ' Set up for next iter
    Loop 'Y

GCR_Done:
    Set GetCornerRange = retval
End Function 'GetCornerRange

' ===========================================================================
' Debugging helpers

Private Sub DEBUGFindDocument(tophwnd As Long)
' Dumps info about the windows under tophwnd
    EnumChildWindows tophwnd, AddressOf DEBUGFindDocument_Callback, tophwnd
End Sub

Private Function DEBUGFindDocument_Callback(ByVal hwnd As Long, ByVal lParam As Long) As Long
    Dim sClass As String
    Dim sTitle As String
    Dim sRect As String: sRect = ""
    Dim winrect As RECT

    DEBUGFindDocument_Callback = 1  'always continue

    sClass = String(255, " ")
    GetClassName hwnd, sClass, 255
    sClass = RTrim(sClass)
    sClass = Left(sClass, Len(sClass) - 1) ' trim trailing null, I think

    sTitle = String(255, " ")
    GetWindowText hwnd, sTitle, 255
    sTitle = RTrim(sTitle)
    sTitle = Left(sTitle, Len(sTitle) - 1)

    If GetWindowRect(hwnd, winrect) <> 0 Then
        sRect = " (" & CStr(winrect.Left) & "," & CStr(winrect.Top) & ")->(" & CStr(winrect.Right) & "," & CStr(winrect.Bottom) & ")"
    End If

    Debug.Print Hex(hwnd) & ": class {" & sClass & "}; title " & "{" & sTitle & "}" & sRect
End Function

Private Sub DEBUG_getwindowrect(hwnd As Long)
' Dump a single window's coords
    Dim winrect As RECT
    Dim sRect As String
    If GetWindowRect(hwnd, winrect) <> 0 Then
        sRect = " (" & CStr(winrect.Left) & "," & CStr(winrect.Top) & ")->(" & CStr(winrect.Right) & "," & CStr(winrect.Bottom) & ")"
    End If
    Debug.Print Hex(hwnd) & ": " & sRect

End Sub 'DEBUG_getwindowrect

Private Sub DEBUG_showselectioncoords(doc As Document)
' Dump the coordinates of the current selection in \p doc's active window.
' Adapted from https://docs.microsoft.com/en-us/office/vba/api/word.window.getpoint
    Dim pLeft As Long
    Dim pTop As Long
    Dim pWidth As Long
    Dim pHeight As Long
    Dim pRight As Long
    Dim pBottom As Long

    doc.ActiveWindow.GetPoint pLeft, pTop, pWidth, pHeight, doc.ActiveWindow.Selection.Range
    pRight = pLeft + pWidth - 1
    pBottom = pRight + pHeight - 1
    Debug.Print " (" & CStr(pLeft) & "," & CStr(pTop) & ")->(" & CStr(pRight) & "," & CStr(pBottom) & ")"
End Sub

Private Sub DEBUG_showscreentext(doc As Document)
' Print the text currently on screen.  A useful sanity check of GetVisibleRange.
    Dim r As Range
    Set r = GetVisibleRange(doc)
    Debug.Print ">>>" & vbCrLf & r.Text & vbCrLf & "<<<"
End Sub

