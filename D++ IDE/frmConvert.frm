VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmConvert 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Coverting C++ to D++"
   ClientHeight    =   1305
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5280
   ControlBox      =   0   'False
   Icon            =   "frmConvert.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1305
   ScaleWidth      =   5280
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ProgressBar bar 
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   960
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   120
      Picture         =   "frmConvert.frx":030A
      Top             =   120
      Width           =   480
   End
   Begin VB.Label lblStatus 
      AutoSize        =   -1  'True
      Caption         =   "Status"
      Height          =   195
      Left            =   120
      TabIndex        =   2
      Top             =   720
      Width           =   450
   End
   Begin VB.Label lblFile 
      Caption         =   "Converting ..."
      Height          =   195
      Left            =   840
      TabIndex        =   0
      Top             =   240
      Width           =   4305
   End
End
Attribute VB_Name = "frmConvert"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cppCode As String
Dim dppCode As String

Public Sub ConvertToDPP(FileName As String)
Dim StartPos As Integer
Show
frmMain.Enabled = False
DoEvents

lblFile.Caption = "Converting " & GetFileName(FileName) & "...."
cppCode = ReadFile(FileName)
dppCode = cppCode

i = InStr(1, cppCode, " main()")
If i <> 0 Then
    StartPos = InStr(i + 7, cppCode, "{") + 1
    dppCode = Mid(cppCode, StartPos, FindEndBracket(StartPos, cppCode) - StartPos)
Else
    MsgBox "Could not locate main().  Conversion halted.", vbCritical, "Conversion Error"
    frmMain.txtText.Text = cppCode
    Unload Me
End If

bar.Max = Len(dppCode)
For i = 1 To Len(dppCode)
    AddOne
    Pause 0.001
Next i

dppCode = FindReplace(dppCode, "cout << ", "screenput ")
dppCode = FindReplace(dppCode, "cin >> ", "screenin ")

dppCode = FindReplace(dppCode, "return 0;", "finish;")
dppCode = FindReplace(dppCode, "return null;", "finish;")

dppCode = FindReplace(dppCode, "int ", "newvar ")
dppCode = FindReplace(dppCode, "string ", "newvar ")
dppCode = FindReplace(dppCode, "char ", "newvar ")
dppCode = FindReplace(dppCode, "double ", "newvar ")
dppCode = FindReplace(dppCode, "float ", "newvar ")

dppCode = FinddReplace(dppCode, "\n", Chr(34) & " & dpp.crlf & " & Chr(34))
'dppCode = FindReplace(dppCode, "{", "")
'dppCode = FindReplace(dppCode, "}", "")

dppCode = FindReplace(dppCode, "<<", "&")
dppCode = FindReplace(dppCode, "endl", "dpp.crlf")
dppCode = FindReplace(dppCode, vbTab, " ")


frmMain.txtText.Text = dppCode
Unload Me
frmMain.Enabled = True
frmMain.SetFocus

End Sub

Private Function FindEndBracket(Start As Integer, Source As String) As Integer
Dim Depth As Integer
For i = Start To Len(Source)
    sChar = Mid(Source, i, 1)
    If sChar = "}" Then
        If Depth = 0 Then
            FindEndBracket = i
            Exit Function
        Else
            Depth = Depth - 1
        End If
    ElseIf sChar = "{" Then
        Depth = Depth + 1
    End If
Next i
FindEndBracket = i
End Function

Private Function FindReplace(sText As String, sFind As String, sReplace As String) As String
    Dim n%, c%
    Dim sTempR$, sTempL$
    lblStatus.Caption = "Replacing " & sFind & "..."
    DoEvents
    DoEvents
    c = 1
    n = 1
    Do
        c = CodeInStr(n, LCase(sText), LCase(sFind))
        If c% <> 0 Then
            sTempL = Mid$(sText, 1, c - 1)
            sTempR = Mid$(sText, c + Len(sFind))
            sText = sTempL & sReplace & sTempR
        End If
        n = c + 1
    Loop Until c = 0
    FindReplace = sText
End Function

Private Function FinddReplace(sText As String, sFind As String, sReplace As String) As String
    Dim n%, c%
    Dim sTempR$, sTempL$
    lblStatus.Caption = "Replacing " & sFind & "..."
    DoEvents
    DoEvents
    c = 1
    n = 1
    Do
        c = InStr(n, LCase(sText), LCase(sFind))
        If c% <> 0 Then
            sTempL = Mid$(sText, 1, c - 1)
            sTempR = Mid$(sText, c + Len(sFind))
            sText = sTempL & sReplace & sTempR
        End If
        n = c + 1
    Loop Until c = 0
    FinddReplace = sText
End Function

Private Function CodeInStr(StartPos As Integer, SourceText As String, ToFind As String) As Long
'This is just like InStr() only it skips over ""'s
Dim iPend As Long, i As Long
Dim sTemp As String

For i = StartPos To Len(SourceText)
    
    If LCase(Mid(SourceText, i, Len(ToFind))) = LCase(ToFind) Then
        CodeInStr = i
        Exit Function
    End If
                          
    If Mid(SourceText, i, 1) = Chr(34) Then
        i = InStr(i + 1, SourceText, Chr(34))
        If i = 0 Then: Exit Function
    End If
Next i
End Function

Private Sub AddOne()
On Error Resume Next
bar.Value = bar.Value + 1
DoEvents
End Sub

