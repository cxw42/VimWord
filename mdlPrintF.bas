Attribute VB_Name = "mdlPrintF"
' From http://www.freevbcode.com/ShowCode.asp?ID=5014

'PrintF family commands in VB
'written by Phlip Bradbury <phlipping@yahoo.com>
'You may use this in your program as long as credit is
'given to me.

'All functions interpret the format as in C, and then
' PrintF calls .Print of forms and picture boxes
'SPrintF returns the string
'FPrintF writes the value to file
'VPrintF, VSPrintF and VFPrintF simulate the similar
'functions in C which take va_list as a parameter, these
'omit the ParamArray keyword, allowing you to pass an array
'of parameters, rather than listing the parameters.
'SPrintF shows a common use of this

'Although there are six functions offerring similar functionality,
'Only VSPrintF contains the actual code. The other 5 simply
'call VSPrintF and handle the returned string.

'It handles all the escape sequences in C
'except \" and \' as they would not help in VB
'It also handles all the parameters except things like
'%ld, %lf, etc, as VB handles that sort of thing internally
'Format handles escape sequences
'  \a   Alert (Bel)
'  \b   Backspace
'  \f   Form Feed
'  \n   Newline (Line Feed)
'  \r   Carriage Return
'  \t   Horizontal Tab
'  \v   Verical Tab
'  \ddd Octal character
'  \xdd Hexadecimal character
'and parameters
'%[flags][width][.precision]formattype
'flags:
'- left justify   + prefix with sign   # prefixes o,x,X with 0 or 0x
'format types handled:
'  %d, %i signed number
'  %u     unsigned number
'  %o     unsigned octal number
'  %x, %X unsigned hexadecimal number
'  %f     floating point number without exponent
'  %e, %E scientific floating point number (with exponent)
'  %g, %G %f or %e, whichever is shorter
'  %c     single character (from ASCII value)
'  %s     String
'so eg %-6.3d is a number with a minimum of 3 digits
'left justified in a field a minimum of 6 characters wide
'use \\ to type a backslash and either \% or %% to type a
'percent sign
'note: %u treats as short (VB As Integer) when converting
'negative numbers. values below -32768 will look odd.
'%o, %x and %X, however, treat as long.
'finally %c is sent the ascii value of the character to
'print, if you want to send a single-character string
'use Asc() or %s, so
'SPrintF("%s", Char) = SPrintF("%c", Asc(Char)) and
'SPrintF("%s", Chr(Num)) = SPrintF("%c", Num)
Option Explicit
Option Compare Binary

'String PRINT Formatted
'equiv to sprintf() in C
Public Function SPrintF(ByVal FormatString As String, ParamArray Parms() As Variant) As String
  Dim A() As Variant, i As Integer
  'create a copy of Parms because we can't pass a ParamArray to a function
  If UBound(Parms) >= LBound(Parms) Then
    ReDim A(LBound(Parms) To UBound(Parms))
    For i = LBound(Parms) To UBound(Parms)
      A(i) = Parms(i)
    Next
  Else
    ReDim A(0 To 0)
  End If
  SPrintF = VSPrintF(FormatString, A)
End Function

'PRINT Formatted
'equiv to printf() in C
Public Sub PrintF(PrintTo As Object, ByVal FormatString As String, ParamArray Parms() As Variant)
  Dim A() As Variant, i As Integer
  'create a copy of Parms because we can't pass a ParamArray to a function
  If UBound(Parms) >= LBound(Parms) Then
    ReDim A(LBound(Parms) To UBound(Parms))
    For i = LBound(Parms) To UBound(Parms)
      A(i) = Parms(i)
    Next
  Else
    ReDim A(0 To 0)
  End If
  PrintTo.Print VSPrintF(FormatString, A);
End Sub

'File PRINT Formatted
'equiv to fprintf() in C
Public Sub FPrintF(FileNum As Integer, ByVal FormatString As String, ParamArray Parms() As Variant)
  Dim A() As Variant, i As Integer
  'create a copy of Parms because we can't pass a ParamArray to a function
  If UBound(Parms) >= LBound(Parms) Then
    ReDim A(LBound(Parms) To UBound(Parms))
    For i = LBound(Parms) To UBound(Parms)
      A(i) = Parms(i)
    Next
  Else
    ReDim A(0 To 0)
  End If
  Print #FileNum, VSPrintF(FormatString, A);
End Sub

'Variable-argument PRINT Formatted
'equiv to vprintf() in C
Public Sub VPrintF(PrintTo As Object, ByVal FormatString As String, Parms() As Variant)
  PrintTo.Print VSPrintF(FormatString, Parms);
End Sub

'Variable-argument File PRINT Formatted
'equiv to vfprintf() in C
Public Sub VFPrintF(FileNum As Integer, ByVal FormatString As String, Parms() As Variant)
  Print #FileNum, VSPrintF(FormatString, Parms);
End Sub

'Variable-argument String PRINT Formatted
'equiv to vsprintf() in C
'this is where the actual work is done
Public Function VSPrintF(ByVal FormatString As String, Parms() As Variant) As String
  'general
  Dim Ret As String
  Dim Char As String
  'escape
  Dim NumberBuffer As String
  'parameters
  Dim ParamUpTo As Integer
  Dim Flags As String
  Dim Width As String
  Dim Precision As String
  Dim Value As Variant
  Dim AddStr As String
  'for calculating %e and %g
  Dim Mantissa As Double, Exponent As Long
  'for calculating %g
  Dim AddStrPercentF As String, AddStrPercentE As String
  ParamUpTo = LBound(Parms)
  Ret = ""
  While FormatString <> ""
    Char = NextChar(FormatString)
    Select Case Char
      Case "\"
        Char = NextChar(FormatString)
        Select Case Char
          Case "a" 'alert (bell)
            Ret = Ret & Chr(7)
          Case "b" 'backspace
            Ret = Ret & vbBack
          Case "f" 'formfeed
            Ret = Ret & vbFormFeed
          Case "n" 'newline (linefeed)
            Ret = Ret & vbLf
          Case "r" 'carriage return
            Ret = Ret & vbCr
          Case "t" 'horizontal tab
            Ret = Ret & vbTab
          Case "v" 'vertical tab
            Ret = Ret & vbVerticalTab
          Case "0" To "9" 'octal character
            NumberBuffer = Char
            While InStr("01234567", Left(FormatString, 1)) And Len(FormatString) > 0
              NumberBuffer = NumberBuffer & NextChar(FormatString)
            Wend
            Ret = Ret & Chr(Oct2Dec(NumberBuffer))
          Case "x" 'hexadecimal character
            NumberBuffer = ""
            While InStr("0123456789ABCDEFabcdef", Left(FormatString, 1)) And Len(FormatString) > 0
              NumberBuffer = NumberBuffer & NextChar(FormatString)
            Wend
            Ret = Ret & Chr(Hex2Dec(NumberBuffer))
          Case "\" 'backslash
            Ret = Ret & "\"
          Case "%" 'percent
            Ret = Ret & "%"
          Case Else 'unrecognised
            Ret = Ret & Char
            Debug.Print "WARNING: Unrecognised escape sequence: \" & Char
        End Select
      Case "%"
        Char = NextChar(FormatString)
        If Char = "%" Then
          Ret = Ret & "%"
        Else
          Flags = ""
          Width = ""
          Precision = ""
          While Char = "-" Or Char = "+" Or Char = "#"
            Flags = Flags & Char
            Char = NextChar(FormatString)
          Wend
          While IsNumeric(Char)
            Width = Width & Char
            Char = NextChar(FormatString)
          Wend
          If Char = "." Then
            Char = NextChar(FormatString)
            While IsNumeric(Char)
              Precision = Precision & Char
              Char = NextChar(FormatString)
            Wend
          End If
          Select Case Char
            Case "d", "i" 'signed decimal
              Value = CLng(Parms(ParamUpTo))
              AddStr = CStr(Abs(Value))
              If Precision <> "" Then
                If val(Precision) > Len(AddStr) Then
                  AddStr = String(val(Precision) - Len(AddStr), "0") & AddStr
                End If
              End If
              If Value < 0 Then
                AddStr = "-" & AddStr
              ElseIf InStr(Flags, "+") Then
                AddStr = "+" & AddStr
              ElseIf InStr(Flags, "-") = 0 Then
                AddStr = " " & AddStr
              End If
            Case "u" 'unsigned decimal
              Value = CLng(Parms(ParamUpTo))
              If Value < 0 Then Value = Value + 65536
              AddStr = CStr(Value)
              If Precision <> "" Then
                If val(Precision) > Len(AddStr) Then
                  AddStr = String(val(Precision) - Len(AddStr), "0") & AddStr
                End If
              End If
            Case "o" 'unsigned octal value
              Value = CLng(Parms(ParamUpTo))
              AddStr = Oct(Value)
              If Precision <> "" Then
                If val(Precision) > Len(AddStr) Then
                  AddStr = String(val(Precision) - Len(AddStr), "0") & AddStr
                End If
              End If
              If InStr(Flags, "#") Then AddStr = "0" & AddStr
            Case "x", "X" 'unsigned hexadecimal value
              Value = CLng(Parms(ParamUpTo))
              AddStr = Hex(Value)
              If Char = "x" Then AddStr = LCase(AddStr)
              If Precision <> "" Then
                If val(Precision) > Len(AddStr) Then
                  AddStr = String(val(Precision) - Len(AddStr), "0") & AddStr
                End If
              End If
              If InStr(Flags, "#") Then AddStr = "0x" & AddStr
            Case "f" 'float w/o exponent
              Value = CDbl(Parms(ParamUpTo))
              If Precision = "" Then Precision = "6"
              AddStr = Format(Abs(Value), "0." & String(val(Precision), "0"))
              If Value < 0 Then
                AddStr = "-" & AddStr
              ElseIf InStr(Flags, "+") Then
                AddStr = "+" & AddStr
              ElseIf InStr(Flags, "-") = 0 Then
                AddStr = " " & AddStr
              End If
            Case "e", "E" 'float w/ exponent
              Value = CDbl(Parms(ParamUpTo))
              Mantissa = Abs(Value)
              Exponent = 0
              If Mantissa > 10 Then
                While Mantissa >= 10
                  Mantissa = Mantissa / 10
                  Exponent = Exponent + 1
                Wend
              Else
                While Mantissa < 1
                  Mantissa = Mantissa * 10
                  Exponent = Exponent - 1
                Wend
              End If
              If Precision = "" Then Precision = "6"
              AddStr = Format(Mantissa, "0." & String(val(Precision), "0"))
              If Right(AddStr, 1) = "." Then AddStr = Left(AddStr, Len(AddStr) - 1)
              AddStr = AddStr & Char & IIf(Exponent < 0, "-", "+") & Format(Exponent, "000")
              If Value < 0 Then
                AddStr = "-" & AddStr
              ElseIf InStr(Flags, "+") Then
                AddStr = "+" & AddStr
              ElseIf InStr(Flags, "-") = 0 Then
                AddStr = " " & AddStr
              End If
            Case "g", "G" 'float w/ or w/o exponent, shorter
              'first calculate without
              Value = CDbl(Parms(ParamUpTo))
              If Precision = "" Then Precision = "6"
              AddStrPercentF = Format(Abs(Value), "0." & String(val(Precision), "#"))
              If Value < 0 Then
                AddStrPercentF = "-" & AddStrPercentF
              ElseIf InStr(Flags, "+") Then
                AddStrPercentF = "+" & AddStrPercentF
              ElseIf InStr(Flags, "-") = 0 Then
                AddStrPercentF = " " & AddStrPercentF
              End If
              'then calculate with
              Value = CDbl(Parms(ParamUpTo))
              Mantissa = Abs(Value)
              Exponent = 0
              If Mantissa > 10 Then
                While Mantissa >= 10
                  Mantissa = Mantissa / 10
                  Exponent = Exponent + 1
                Wend
              Else
                While Mantissa < 1
                  Mantissa = Mantissa * 10
                  Exponent = Exponent - 1
                Wend
              End If
              If Precision = "" Then Precision = "6"
              AddStrPercentE = Format(Mantissa, "0." & String(val(Precision), "#"))
              If Right(AddStrPercentE, 1) = "." Then AddStrPercentE = Left(AddStrPercentE, Len(AddStrPercentE) - 1)
              AddStrPercentE = AddStrPercentE & IIf(Char = "G", "E", "e") & IIf(Exponent < 0, "-", "+") & Format(Exponent, "000")
              If Value < 0 Then
                AddStrPercentE = "-" & AddStrPercentE
              ElseIf InStr(Flags, "+") Then
                AddStrPercentE = "+" & AddStrPercentE
              ElseIf InStr(Flags, "-") = 0 Then
                AddStrPercentE = " " & AddStrPercentE
              End If
              'find shortest
              AddStr = IIf(Len(AddStrPercentF) > Len(AddStrPercentE), AddStrPercentE, AddStrPercentF)
            Case "c" 'single character, passed ASCII value
              Value = CByte(Parms(ParamUpTo))
              AddStr = Chr(Value)
            Case "s" 'string
              Value = CStr(Parms(ParamUpTo))
              AddStr = Value
            Case Else
              Debug.Print "WARNING: unrecognised parameter sequence: %" & Flags & Width & IIf(Precision <> "", "." & Precision, "") & Char
              AddStr = "%" & Flags & Width & IIf(Precision <> "", "." & Precision, "") & Char
          End Select
          If Width <> "" Then
            If val(Width) > Len(AddStr) Then
              If InStr(Flags, "-") Then
                AddStr = AddStr & space(val(Width) - Len(AddStr))
              Else
                AddStr = space(val(Width) - Len(AddStr)) & AddStr
              End If
            End If
          End If
          ParamUpTo = ParamUpTo + 1
          Ret = Ret & AddStr
        End If
      Case Else
        Ret = Ret & Char
    End Select
  Wend
  VSPrintF = Ret
End Function

'Various helper functions

'returns the first character from a buffer and removes
'it from the buffer
Private Function NextChar(ByRef Buffer As String) As String
  NextChar = Mid(Buffer, 1, 1)
  Buffer = Mid(Buffer, 2)
End Function

'convert octal to decimal
Private Function Oct2Dec(ByVal Octal As String) As Long
  Dim i As Integer
  i = 0
  While Octal <> ""
    i = i * 8 + val(NextChar(Octal))
  Wend
  Oct2Dec = i
End Function

'convert hexadecimal to decimal
Private Function Hex2Dec(ByVal Hexadecimal As String) As Long
  Hex2Dec = CLng("&H" & Hexadecimal)
End Function

