Attribute VB_Name = "mConstants"
' mConstants: Copyright (c) Chris White 2016--2018.
'   2018/04/27  chrisw  Imported from Normal.dotm
'   2019-02-18  chrisw  Removed W_COMMENT from whitespace groups

Option Explicit
Option Base 0

Public Enum CHARCODES
    U_TAB = &H9
    U_LF = &HA              ' In Word, appears to be smashed to Chr(13).
    U_CR = &HD              ' In Word, the end of a paragraph
    U_SPACE = &H20
    U_HYPHEN = &H2D         ' A regular hyphen (HYPHEN-MINUS)
    U_OPT_HYPHEN = &HAD     ' Optional hyphen, in Unicode (SOFT HYPHEN)
    U_ZWJ = &H200D          ' zero-width joiner

    U_REAL_HYPHEN = &H2010  ' Unicode HYPHEN
    U_NONBREAK_HYPHEN = &H2011  ' Unicode non-breaking hyphen
    U_FIGURE_DASH = &H2012
    U_EN_DASH = &H2013
    U_EM_DASH = &H2014
    U_CURLY_BACKQUOTE = &H2018      ' Typographic single open-quote
    U_CURLY_APOS = &H2019           ' Typographic apostrophe (RIGHT SINGLE QUOTATION MARK)
    U_CURLY_OPENDQUOTE = &H201C     ' Typographic open double-quote
    U_CURLY_CLOSEDQUOTE = &H201D    ' Typographic close double-quote
    U_NUMERO_SIGN = &H2116  ' which I use to hide numbers from frmRebuildPNXRefs
    U_WAVE_DASH = &H301C    ' Japanese figure dash
    U_FULLWIDTH_TILDE = &HFF5E  ' Alternative JP figure dash

    U_PU1 = &H91        ' PRIVATE USE ONE
    
    ' Word-specific constants
    W_FOOTNOTE_MARK = 2 ' marker in the text for the footnote number - same code in
                        ' both the main-text and footnote stories.
    W_COMMENT = 5       ' the marker in the body of the text where a comment is
    W_LINE_BREAK = 11   ' manual line-break
    W_FUNKY_BREAK = 12  ' page break or section break - both are represented by the same char :(
    W_FIELD_START = 19
    W_FIELD_END = 21
    
    W_NBHYPHEN = &H1E   ' non-breaking hyphen, MS Word
    W_OPTHYPHEN = &H1F  ' optional hyphen
    W_NBSP = &HA0

End Enum

Public Const WC_WS_CHARS_NOEOL As String = " ^t^s"
    ' For wildcards.  Whitespace and comment flags, omitting \r\n.
    ' Note: Comment flags (^05) MUST NOT be include in charsets,
    ' because Word locks when attempting to traverse them in
    ' Print Layout view (where they are not visible as characters).

Public Const PAT_EN_DASH As String = "\u2013"
Public Const PAT_EM_DASH As String = "\u2014"

Public Const PAT_W_NBHYPHEN As String = "\x1e"
Public Const PAT_W_OPTHYPHEN As String = "\x1f"

Public Const SN_CLAIM = "Claim"        ' Style name
Public Const SN_SUBCLAIM = "SubClaim"

Public Const DOCVAR_STASHED_BM = "STASHED_BM"
Public Const MAGIC_FLAG_NOT_REAL_TEXT As String = "#.quux.#"
        ' Used in OAMCAA and related for "Withdrawn and currently amended".
        ' Should be text that doesn't occur anywhere else in the paragraph.
Public Const MAGIC_FLAG2_NOT_REAL_TEXT As String = "#.flarp.#"
        ' Used in OAMCAA and related for "Cancelled".
        ' Should be text that doesn't occur anywhere else in the paragraph.

' Whitespace
Public Const PAT_WS_CHARS As String = "\x09\x0a\x0d\x20\xa0" & _
    "\u1680\u180e\u2000\u2001\u2002\u2003\u2004\u2005\u2006" & _
    "\u2007\u2008\u2009\u200a\u200b\u202f\u205f\u3000\ufeff" & _
    "\u2028\u2029"
    ' The guts of a whitespace character class including NBSP.
    ' Thanks to https://www.cs.tut.fi/~jkorpela/chars/spaces.html for a
    ' consolidated list of unicode spaces.
    ' Also:
    '   http://www.fileformat.info/info/unicode/category/Zs/list.htm
    '   http://www.fileformat.info/info/unicode/category/Zl/list.htm
    '   http://www.fileformat.info/info/unicode/category/Zp/list.htm
    ' NOTE: W_COMMENT is not included, for consistency with Range
    ' character sets.

Public Const PAT_WS As String = "[" & PAT_WS_CHARS & "]"
    ' The corresponding character class

Public Const PAT_WS_STAR As String = PAT_WS & "*"
    ' Sugar for zero or more whitespace chars
Public Const PAT_WS_PLUS As String = PAT_WS & "+"
    ' 100% organic, fair trade, non-GMO cane sugar.  Oh so sweet!

Public Const S_SPACE = " "

