VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisDocument"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
' VimWord.ThisDocument
' Copyright (c) Chris White 2018
' CC-BY-NC-SA 4.0, or any later version, at your option.

Private Sub Document_Close()
    ' When prjVimWord closes, also close the scratchpad.
    MarkScratchpadAsSaved_ True
        ' Don't save changes to disk.
        ' True => close the document if it's open.
End Sub

