Attribute VB_Name = "modColorize"
Option Explicit

Public Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long
Public Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function SendMessageByNum Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Long) As Long
Public Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Public Declare Function GetTextMetrics Lib "gdi32" Alias "GetTextMetricsA" (ByVal hdc As Long, lpMetrics As TEXTMETRIC) As Long
Public Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hdc As Long) As Long

Private Const EM_GETFIRSTVISIBLELINE = &HCE
Private Const EM_LINEINDEX = &HBB
Private Const EM_GETRECT = &HB2
Private Const WM_GETFONT = &H31

Const KeyWords = "|screenout|screenput|screenin|set|if|newvar|title|endif|else|pause|delete|box|do|while|loop|end|screenpass|"

Dim FirstVisibleLine As Long
Dim LastVisibleLine As Long


Private Type RECT
  Left As Long
  Top As Long
  Right As Long
  Bottom As Long
End Type

Private Type TEXTMETRIC
  tmHeight As Long
  tmAscent As Long
  tmDescent As Long
  tmInternalLeading As Long
  tmExternalLeading As Long
  tmAveCharWidth As Long
  tmMaxCharWidth As Long
  tmWeight As Long
  tmOverhang As Long
  tmDigitizedAspectX As Long
  tmDigitizedAspectY As Long
  tmFirstChar As Byte
  tmLastChar As Byte
  tmDefaultChar As Byte
  tmBreakChar As Byte
  tmItalic As Byte
  tmUnderlined As Byte
  tmStruckOut As Byte
  tmPitchAndFamily As Byte
  tmCharSet As Byte
End Type

Public Sub Colorize(RTFBox As RichTextBox, CommentColor, StringColor, KeysColor)

Dim lTextSelPos As Long, lTextSelLen As Long
'// Save the cursor position
lTextSelPos = RTFBox.SelStart
lTextSelLen = RTFBox.SelLength

'// Lock the WindowUpdate of the ReichTextBox
LockWindowUpdate RTFBox.hWnd

'On Error GoTo ErrHandler

Dim i As Long
Dim sBuffer As String, lBufferLen As Long
Dim lSelPos As Long, lSelLen As Long
Dim sTempBuffer As String
Dim sSearchChar As String, lSearchCharLen As Long

With RTFBox
    sBuffer = .Text & " "
    lBufferLen = Len(sBuffer)
    sTempBuffer = ""
    
    For i = FirstVisibleChar(RTFBox) To LastVisibleChar(RTFBox, lBufferLen)

      Select Case Asc(Mid(sBuffer, i, 1))
      
        Case 34
          .SelStart = i - 1
          i = InStr(i + 1, sBuffer, """", 1)
          .SelLength = i - .SelStart
          .SelColor = StringColor
        
        Case 62
          
          If Mid(sBuffer, i, 1) = ">" Then
            sSearchChar = vbCrLf
            lSearchCharLen = 0
          Else
            GoTo ExitComment
          End If
          
          sTempBuffer = ""
          
          .SelStart = i - 1
          lSelLen = InStr(i, sBuffer, sSearchChar) + lSearchCharLen
          If lSelLen <> lSearchCharLen Then '// FileEnd ?
            lSelLen = lSelLen - i
          Else
            lSelLen = lBufferLen - i
          End If
          .SelLength = lSelLen
          .SelColor = CommentColor
          i = .SelStart + .SelLength
          
ExitComment:

        Case 97 To 122, 65 To 90, 35
          '// a to  z ,  A to Z , #
          If sTempBuffer = "" Then lSelPos = i
          sTempBuffer = sTempBuffer & Mid(sBuffer, i, 1)
          
        Case Else
          If Trim(sTempBuffer) <> "" Then
            .SelStart = lSelPos - 1
            .SelLength = Len(sTempBuffer)
            If InStr(1, KeyWords, "|" & sTempBuffer & "|", 1) <> 0 Then
             .SelColor = KeysColor
            End If
          End If
        
          sTempBuffer = ""
        End Select
      Next
End With

ErrHandler:

'// Set the Cursor to the old position
RTFBox.SelStart = lTextSelPos
RTFBox.SelLength = lTextSelLen

'// Unlock the WindoUpdate-Lock
LockWindowUpdate 0

End Sub

Private Function FirstVisibleChar(RTFBox As RichTextBox) As Long
  FirstVisibleLine = SendMessage(RTFBox.hWnd, EM_GETFIRSTVISIBLELINE, 0, 0&)
  FirstVisibleChar = SendMessageByNum(RTFBox.hWnd, EM_LINEINDEX, FirstVisibleLine, 0&)
  If FirstVisibleChar = 0 Then FirstVisibleChar = 1
End Function


Private Function LastVisibleChar(RTFBox As RichTextBox, LenFile As Long) As Long
Dim rc As RECT
Dim tm As TEXTMETRIC
Dim hdc As Long
Dim lFont As Long
Dim OldFont As Long
Dim di As Long
Dim lc As Long
Dim VisibleLines As Long

  lc = SendMessage(RTFBox.hWnd, EM_GETRECT, 0, rc)
  lFont = SendMessage(RTFBox.hWnd, WM_GETFONT, 0, 0)
  hdc = GetDC(RTFBox.hWnd)
  If lFont <> 0 Then OldFont = SelectObject(hdc, lFont)
  di = GetTextMetrics(hdc, tm)
  If lFont <> 0 Then lFont = SelectObject(hdc, OldFont)
  VisibleLines = (rc.Bottom - rc.Top) / tm.tmHeight
  di = ReleaseDC(RTFBox.hWnd, hdc)
  
  LastVisibleLine = SendMessage(RTFBox.hWnd, EM_GETFIRSTVISIBLELINE, 0, 0&)
  LastVisibleLine = LastVisibleLine + VisibleLines
  
  LastVisibleChar = SendMessageByNum(RTFBox.hWnd, EM_LINEINDEX, LastVisibleLine, 0&)
  If LastVisibleChar = -1 Or LastVisibleChar = 0 Then LastVisibleChar = LenFile
  
End Function

