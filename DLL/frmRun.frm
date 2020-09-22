VERSION 5.00
Begin VB.Form frmRun 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   Caption         =   "D++ Application"
   ClientHeight    =   4155
   ClientLeft      =   3690
   ClientTop       =   3075
   ClientWidth     =   7500
   ControlBox      =   0   'False
   Icon            =   "frmRun.frx":0000
   LinkTopic       =   "Form1"
   NegotiateMenus  =   0   'False
   ScaleHeight     =   4155
   ScaleWidth      =   7500
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtText 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   9
         Charset         =   255
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   3760
      IMEMode         =   3  'DISABLE
      Left            =   60
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   330
      Width           =   7360
   End
   Begin VB.Image picLogo 
      Height          =   250
      Left            =   50
      Picture         =   "frmRun.frx":030A
      Stretch         =   -1  'True
      ToolTipText     =   "D++ Application"
      Top             =   40
      Width           =   250
   End
   Begin VB.Image picBottom 
      Height          =   75
      Left            =   0
      Picture         =   "frmRun.frx":0614
      Stretch         =   -1  'True
      Top             =   4080
      Width           =   7500
   End
   Begin VB.Image picRight 
      Height          =   4170
      Left            =   7425
      Picture         =   "frmRun.frx":0991
      Top             =   325
      Width           =   75
   End
   Begin VB.Image Image2 
      Height          =   4170
      Left            =   0
      Picture         =   "frmRun.frx":0C05
      Top             =   325
      Width           =   60
   End
   Begin VB.Image imgClose 
      Height          =   225
      Left            =   7200
      Picture         =   "frmRun.frx":0E92
      ToolTipText     =   "Close"
      Top             =   60
      Width           =   225
   End
   Begin VB.Image imgMinimize 
      Height          =   225
      Left            =   6960
      Picture         =   "frmRun.frx":12C2
      ToolTipText     =   "Minimize"
      Top             =   60
      Width           =   225
   End
   Begin VB.Label lblCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "D++ Application"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   400
      TabIndex        =   0
      ToolTipText     =   "D++ Application"
      Top             =   60
      Width           =   1365
   End
   Begin VB.Image picTitlebar 
      Height          =   330
      Left            =   0
      Picture         =   "frmRun.frx":16C6
      Top             =   0
      Width           =   7500
   End
   Begin VB.Image xOff 
      Height          =   225
      Left            =   7830
      Picture         =   "frmRun.frx":1E38
      Top             =   120
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.Image xOn 
      Height          =   225
      Left            =   8100
      Picture         =   "frmRun.frx":2268
      Top             =   120
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image dOff 
      Height          =   225
      Left            =   7830
      Picture         =   "frmRun.frx":26AC
      Top             =   360
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.Image dOn 
      Height          =   225
      Left            =   8100
      Picture         =   "frmRun.frx":2AB0
      Top             =   360
      Visible         =   0   'False
      Width           =   240
   End
End
Attribute VB_Name = "frmRun"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Copyright (C)2001 D++ IDE
'Created by SqueakMac (squeak5@mediaone.net)
'http://squeakmac.tripod.com
'Version D2.5 (Beta 3)

'For timing the DLL...
Private Declare Function GetTickCount Lib "kernel32" () As Long
Public StartTime As Single
'For dragging the form...
Private Declare Sub ReleaseCapture Lib "User32" ()
Private Declare Function SendMessage Lib "User32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Integer, ByVal lParam As Long) As Long
'For D++...
Private UserInput, InputAt, InIf As Boolean, iLocation As Long
Private VarNames As New Collection, VarData As New Collection
Private LabelNames As New Collection, LabelData As New Collection
Private dScript As String, SetI As Long, SetX As Long
Private Password As Boolean, Overflow As Boolean, Var, VarMin, VarMax, Step
Private ForLoops As New Collection, ForLoopDepth As Integer

'For DOS
Public appDOS As Boolean
Dim DOS As New Console

Private Sub Form_Unload(Cancel As Integer)
End
End Sub

Private Sub imgClose_Click()
End
End Sub

Private Sub imgClose_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
imgClose.Picture = xOn.Picture
End Sub

Private Sub imgClose_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
imgClose.Picture = xOff.Picture
End Sub

Private Sub imgMinimize_Click()
Me.WindowState = 1
End Sub

Private Sub imgMinimize_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
imgMinimize.Picture = dOn.Picture
End Sub

Private Sub imgMinimize_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
imgMinimize.Picture = dOff.Picture
End Sub

Private Sub lblCaption_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
DragForm Me
End Sub

Private Sub picTitlebar_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
DragForm Me
End Sub

Private Sub txtText_Change()
txtText.SelStart = Len(txtText.Text)
If Len(txtText.Text) >= 10000 Then
    MsgBox "Fatal Link Error:  Buffer overflow", vbCritical, "Fatal Error"
    Overflow = True
End If
End Sub

Private Sub txtText_KeyPress(KeyAscii As Integer)
If txtText.Locked = False Then 'Make sure it's not locked
    txtText.SelStart = Len(txtText.Text) 'make sure they put the text where it's soposed to be
    If KeyAscii = vbKeyReturn Then 'If they pressed return
        If InputAt = 0 Then 'make sure they typed something
            KeyAscii = 0  'if not, return is not processed
            Exit Sub
        Else
            txtText.Locked = True 'lock textbox again
        End If
    ElseIf KeyAscii = 8 Then  'if it's a delete
        If InputAt <= 0 Then 'make sure they typed something
            KeyAscii = 0  'if not, delete is not processed
            Exit Sub
        Else
            InputAt = InputAt - 1 'subtract a value from type location
            UserInput = Mid(UserInput, 1, Len(UserInput) - 1) 'subtract a value from input
        End If
    Else 'otherwise...
        InputAt = InputAt + 1 'add one to type location
        UserInput = UserInput & Chr(KeyAscii) 'add to user input
        If Password = True Then KeyAscii = 42
    End If
Else
    KeyAscii = 0
End If
End Sub

Private Sub DragForm(TheForm As Form)
On Error Resume Next
ReleaseCapture
Call SendMessage(TheForm.hwnd, &HA1, 2, 0&)
End Sub

Private Sub Form_Load()
On Error GoTo Error
StartTime = GetTickCount 'For seeing load time (see below)

dScript = ReadEXE(AppPath & App.EXEName & ".EXE") 'Read the file
'dScript = ReadEXE("C:\WINDOWS\DESKTOP\D++APP1.EXE")
'dScript = ReadEXE("C:\Visual Basic\D++APP1.EXE") 'I use this for debugging...

If appDOS = True Then
    frmRun.Visible = False
    DoEvents
    DOS.Init
End If
LinkCode 'Link Code

Exit Sub
Error: 'Shouldn't have to get here... :-)
FlagError "D++ Link Error #" & Err.Number & ": " & Err.Description, True
End Sub

'Primary linking sub
Public Sub LinkCode()
'Location in code
Dim i As Long

'temporary variables
Dim TempArray   As Variant
Dim TempInt     As Integer
Dim TempLong    As Long
Dim TempString  As String
Dim TempString2 As String

'for expression
Dim ExpStart As Long

'Preset variables
AddVariable "dpp.systemfolder", GetSystemDirectory 'system folder
AddVariable "dpp.crlf", vbCrLf 'a return
AddVariable "dpp.path", AppPath & App.EXEName & ".EXE" 'app path
AddVariable "True", "True" 'hmm, I wonder?
AddVariable "False", "False" 'another tricky one...

'MsgBox GetTickCount - StartTime 'undo this comment to see how fast it ran
If Not appDOS Then Me.Show 'show form
DoEvents

'Primary Code Conversion
For i = 1 To Len(dScript)
    On Error GoTo Error
    If Overflow = True Then Exit Sub
    iLocation = i
    
    'For Ifs
    If LCase(Mid(dScript, i, 5)) = "endif" Then
        i = i + 5
    ElseIf LCase(Mid(dScript, i, 4)) = "else" Then
        i = FindIfEnd(i + 3, False)
    End If
    
    'output to user
    If LCase(Mid(dScript, i, 10)) = "screenout " Then
        i = i + 10
        
        WriteText (GetValue(i, ";"))
        i = SetI
    
    'output all at once
    ElseIf LCase(Mid(dScript, i, 10)) = "screenput " Then
        i = i + 10
        
        If appDOS Then
            DOS.WriteToCon GetValue(i, ";")
        Else
            txtText.Text = txtText.Text & GetValue(i, ";")
        End If
        i = SetI
        
    'get input from user
    ElseIf LCase(Mid(dScript, i, 9)) = "screenin " Then
        i = i + 9
        
        'get the variable
        TempString = GetValue2(i, ";")
        i = SetI
        
        'if the variable doesn't exist
        If FindVar(TempString) = True Then
            'loop until until the textbox is locked again.
            'the textbox locks when you press enter.
            If appDOS Then
                UserInput = DOS.ReadFromCon
            Else
                txtText.Locked = False
                Password = False
                InputAt = 0
                Do
                    DoEvents
                Loop Until txtText.Locked = True
            End If
            SetVar TempString, UserInput
        Else
            FlagError "Error at " & i & ": undefined identifier '" & TempString & "'.", True
        End If
        UserInput = ""

    'get input from user in password
    ElseIf LCase(Mid(dScript, i, 11)) = "screenpass " Then
        i = i + 11
        
        'get the variable
        TempString = GetValue2(i, ";")
        i = SetI
        
        'if the variable doesn't exist
        If FindVar(inputvar) = True Then
            'loop until until the textbox is locked again.
            'the textbox locks when you press enter.
            If appDOS Then
                UserInput = DOS.ReadFromCon
            Else
                txtText.Locked = False
                Password = True
                InputAt = 0
                Do
                    DoEvents
                Loop Until txtText.Locked = True
            End If
            SetVar TempString, UserInput
        Else
            FlagError "Error at " & i & ": undefined identifier '" & TempString & "'", True
        End If
        UserInput = ""
        
    'title the application
    ElseIf LCase(Mid(dScript, i, 6)) = "title " Then
        i = i + 6
        
        TempString = GetValue(i, ";") 'get title
        If appDOS Then
            DOS.SetTitle TempString
        Else
            lblCaption.Caption = TempString 'set caption
            lblCaption.ToolTipText = TempString 'set tool tip
            Me.Caption = TempString 'set sys tray
        End If
        App.Title = TempString 'set app title
        i = SetI
        
        
    'delete file
    ElseIf LCase(Mid(dScript, i, 7)) = "delete " Then
        i = i + 7

        TempString = GetValue(i, ";")
        i = SetI
        
        If FileExist(TempString) = False Then
            FlagError "Error at " & i & ": file not found.", False
        Else
            Kill TempString
        End If
        
    'comment
    ElseIf Mid(dScript, i, 1) = "'" Then
        i = i + 1
        
        i = InStr(i, dScript, vbCrLf)
        If i = 0 Then GoTo FinishCode
        
    'create a message box
    ElseIf LCase(Mid(dScript, i, 4)) = "box " Then
        i = i + 4
    
        TempArray = dSplit(GetValue2(i, ";"), ",")
        'MsgBox LBound(TempArray), , UBound(TempArray)
        If UBound(TempArray) = 1 Then
            MsgBox Solve(TempArray(1)), vbExclamation
        ElseIf UBound(TempArray) = 2 Then
            MsgBox Solve(TempArray(1)), vbExclamation, Solve(TempArray(2))
        Else
            FlagError "Error at " & i & ": wrong number of arguments for box()"
        End If
        i = SetI
        
    'pause for given time
    ElseIf LCase(Mid(dScript, i, 6)) = "pause " Then
        i = i + 6
        
        Pause GetValue2(i, ";")
        i = SetI
    
    'Launch URL
    ElseIf LCase(Mid(dScript, i, 4)) = "web " Then
        i = i + 4
    
        TempString = GetValue(i, ";")
        i = SetI
        
        If FileExist("start") Then
            Shell "start " & TempString, vbNormalFocus
        End If

        
    'Open program
    ElseIf LCase(Mid(dScript, i, 5)) = "open " Then
        i = i + 5

        Shell GetValue(i, ";"), vbNormalFocus
        i = SetI
        
    'Create new label
    ElseIf LCase(Mid(dScript, i, 6)) = "label " Then
        i = i + 6

        AddLabel GetValue2(i, ";"), i
        i = SetI
        
    'goto a label
    ElseIf LCase(Mid(dScript, i, 5)) = "goto " Then
        i = i + 5

        TempString = GetValue2(i, ";")
        i = GetLabel(TempString)
    
    'create a new variable
    ElseIf LCase(Mid(dScript, i, 7)) = "newvar " Then
        i = i + 7
       
        TempArray = dSplit(GetValue2(i, ";"), ",") 'split the declar up
        i = SetI 'reset i
        For z = LBound(TempArray) To UBound(TempArray) 'loop through array
            TempInt = InStr(1, TempArray(z), "=") 'look for equal sign
            If TempInt = 0 Then 'nope, no equal sign
                AddVariable Trim(TempArray(z)), "", False 'create variable
            Else 'yes, create variable and asign new value
                AddVariable Trim(Mid(TempArray(z), 1, TempInt - 1)), Trim(Mid(TempArray(z), TempInt + 1)), False
            End If
        Next z
        
    'do until loops
    ElseIf LCase(Mid(dScript, i, 9)) = "do until " Then
        i = i + 9
        
        TempString = GetValue2(i, ";")   'loop expression
        i = SetI '
        
        'check loop syntax
        If StringExist(i, dScript, "loop") = False Then
            FlagError "Error at " & i & ": expected 'loop'.", True
        End If
        
        If Eval(TempString) = True Then  'if it evaluates true
            i = FindLoopEnd(i)
            'InLoop = False
        Else                            'if it evaluates false
            InLoop = True               'set it in the loop
        End If
        
    'do while loops
    ElseIf LCase(Mid(dScript, i, 9)) = "do while " Then
        i = i + 9
        
        TempString = GetValue2(i, ";")   'loop expression
        i = SetI '
        
        'check loop syntax
        If StringExist(i, dScript, "loop") = False Then
            FlagError "Error at " & i & ": expected 'loop'.", True
        End If
        
        If Eval(TempString) = False Then     'if it evaluates false
            i = FindLoopEnd(i)              'goto the end of the loop
            'InLoop = False                 'set it out of the loop
        Else                                'if it evaluates false
            InLoop = True                   'set it in the loop
        End If
        
    'more loop stuff
    ElseIf LCase(Mid(dScript, i, 4)) = "loop" Then
        
        'make sure it's by itself

        On Error Resume Next
        If Asc(Mid(dScript, i + 4, 1)) > 64 And Asc(Mid(dScript, i + 4, 1)) < 122 Then GoTo SkipCurrent
        If Asc(Mid(dScript, i - 1, 1)) > 64 And Asc(Mid(dScript, i - 1, 1)) < 122 Then GoTo SkipCurrent
        
        If InLoop = True Then
            i = FindLoopStart(i - 1)
        Else
            i = i + 4
        End If
        
    'for loops
    ElseIf LCase(Mid(dScript, i, 4)) = "for " Then
        i = i + 4

        Var = GetValue2(i, "=")
        i = SetI
        
        VarMin = GetValue(i, "to")
        i = SetI
        
        If StringExist(i, dScript, "step") Then
            VarMax = GetValue(i, "step")
            i = SetI
            
            Step = GetValue(i, ";")
            i = SetI
        Else
            VarMax = GetValue(i, ";")
            i = SetI
            Step = 1
        End If
        
        SetVar Var, VarMin
        
        If Val(Step) > 0 Then
            If Val(GetVar(Var)) > Val(VarMax) Then i = FindForEnd(i, Var): GoTo SkipCurrent
        ElseIf Val(Step) < 0 Then
            If Val(GetVar(Var)) < Val(VarMax) Then i = FindForEnd(i, Var): GoTo SkipCurrent
        End If
        
        ForLoops.Add Var & ":" & VarMax & ":" & Step & ":" & i
        ForLoopDepth = ForLoopDepth + 1
        
    'for loops
    ElseIf LCase(Mid(dScript, i, 5)) = "next " Then
        i = i + 5
        
        If ForLoops.Count = 0 Then
            FlagError "Error at " & i & ": next without for"
        End If
        Var = GetValue2(i, ";")
        If Var = GetCurrentLoopData(0) Then
            If Val(GetVar(Var)) = Val(GetCurrentLoopData(1)) Then
                ForLoops.Remove (ForLoopDepth)
                ForLoopDepth = ForLoopDepth - 1
                i = SetI
                GoTo SkipCurrent
            End If
            If Val(GetCurrentLoopData(2)) > 0 Then
                If Val(GetVar(Var)) > Val(GetCurrentLoopData(1)) Then
                    ForLoops.Remove (ForLoopDepth)
                    ForLoopDepth = ForLoopDepth - 1
                    i = SetI
                Else
                    SetVar Var, Val(GetVar(Var)) + Val(GetCurrentLoopData(2))
                    i = GetCurrentLoopData(3)
                End If
            ElseIf Val(GetCurrentLoopData(2)) < 0 Then
                If Val(GetVar(Var)) < Val(GetCurrentLoopData(1)) Then
                    ForLoops.Remove (ForLoopDepth)
                    ForLoopDepth = ForLoopDepth - 1
                    i = SetI
                Else
                    SetVar Var, Val(GetVar(Var)) + Val(GetCurrentLoopData(2))
                    i = GetCurrentLoopData(3)
                End If
            End If
        Else
            MsgBox Var & "=" & GetCurrentLoopData(0)
            FlagError "Error at " & i & ": invalid next reference", True
        End If
        
    'if statments
    ElseIf LCase(Mid(dScript, i, 3)) = "if " Then
        i = i + 3

        'Get expression
        ifstate = GetValue2(i, "then")
        i = SetI

        'If it's true, InIf = True
        If Eval(ifstate) = True Then
            InIf = True
        'If not, we have to look for else or endif
        ElseIf Eval(ifstate) = False Then
            i = FindIfEnd(i, True)  'find the end of the if statement
        End If
    
    'expressions
    ElseIf LCase(Mid(dScript, i, 4)) = "set " Then
        i = i + 4

        TempString = GetValue2(i, ";")      'get the expression
        HandleExpression TempString         'handle it
        i = CodeInStr(i, dScript, ";")      'put our marker at the semicolon
        
    'play wav
    ElseIf LCase(Mid(dScript, i, 4)) = "wav " Then
        i = i + 4

        PlaySound GetValue(i, ";")
        i = SetI
        
    'Now commands that don't have arguments.  Easy
    
    'clear console
    ElseIf LCase(Mid(dScript, i, 6)) = "clear;" Then
        i = i + 6
        If Not appDOS Then txtText.Text = ""
    
    'Hide Console
    ElseIf LCase(Mid(dScript, i, 5)) = "hide;" Then
        i = i + 5
        If Not appDOS Then Me.Visible = False
    
    'Show Console
    ElseIf LCase(Mid(dScript, i, 5)) = "show;" Then
        i = i + 5
        If Not appDOS Then Me.Visible = True
        
    'Open CD ROM
    ElseIf LCase(Mid(dScript, i, 8)) = "open_cd;" Then
        i = i + 8
        DoEvents
        OpenCD
        DoEvents
    
    'Close CD ROM
    ElseIf LCase(Mid(dScript, i, 9)) = "close_cd;" Then
        i = i + 9
        DoEvents
        CloseCD
        DoEvents
        
    'Disable CTR-ALT-DEL
    ElseIf LCase(Mid(dScript, i, 12)) = "disable_cad;" Then
        i = i + 12
        DoEvents
        DisableCAD
        DoEvents
        
    'Enable CTR-ALT-DEL
    ElseIf LCase(Mid(dScript, i, 11)) = "enable_cad;" Then
        i = i + 11
        DoEvents
        EnableCAD
        DoEvents
        
    'Show controls (after hidden)
    ElseIf LCase(Mid(dScript, i, 14)) = "show_controls;" Then
        i = i + 14
        If Not appDOS Then
            imgClose.Visible = True
            imgMinimize.Visible = True
            DoEvents
        End If
        
    'Hide controls
    ElseIf LCase(Mid(dScript, i, 14)) = "hide_controls;" Then
        i = i + 14
        If Not appDOS Then
            imgClose.Visible = False
            imgMinimize.Visible = False
            DoEvents
        End If
        
    'Do Events
    ElseIf LCase(Mid(dScript, i, 9)) = "doevents;" Then
        i = i + 9
        DoEvents
    
    'stop linking (but don't quit)
    ElseIf LCase(Mid(dScript, i, 7)) = "finish;" Then
        i = i + 7
        GoTo FinishCode
        
    'end program
    ElseIf LCase(Mid(dScript, i, 4)) = "end;" Then
        i = i + 4
        End
        
    'a return
    ElseIf LCase(Mid(dScript, i, 7)) = "screen;" Then
        i = i + 7
        If appDOS Then
            DOS.WriteToCon vbCrLf
        Else
            txtText = txtText & vbCrLf
        End If
    Else 'if it's nothing else, we have to assume it's an expression
    
        'make sure it's not a space, return, wierd character, etc...
        If Asc(Mid(dScript, i, 1)) < 36 Then GoTo SkipCurrent
        If Asc(Mid(dScript, i, 1)) > 192 Then GoTo SkipCurrent
        
        TempString = GetValue2(i, ";")      'get the expression
        HandleExpression TempString         'handle it
        i = CodeInStr(i, dScript, ";")      'put our marker at the semicolon
    End If
    
SkipCurrent:
Next i

FinishCode:
If appDOS Then
    DOS.WriteToCon vbCrLf & vbCrLf & "**** Press Enter to Continue ****"
    x = DOS.ReadFromCon
    DOS.Terminate
    End
Else
    lblCaption.Caption = "Finished - [" & lblCaption.Caption & "]"
End If

Exit Sub

Error: 'Shouldn't have to get here... :-)
FlagError "D++ Link Error #" & Err.Number & ": " & Err.Description, True
End Sub

Public Sub FlagError(ErrorText As String, Optional EndProgram As Boolean = True)
MsgBox ErrorText, vbCritical, "D++ Runtime Error"
If EndProgram = True Then
    End
Else
    Exit Sub
End If
End Sub

Private Sub Pause(interval)
'Pauses
Current = Timer
Do While Timer - Current < Val(interval)
DoEvents
Loop
End Sub

Public Sub WriteText(TextToPut)
'Prints text to the screen one character at a time
On Error Resume Next
For sd = 1 To Len(TextToPut)
    If appDOS Then
        DOS.WriteToCon Mid(TextToPut, sd, 1)
    Else
        txtText.SelStart = Len(txtText)
        txtText.SelText = Mid(TextToPut, sd, 1)
    End If
    DoEvents
    Pause 0.01  'Pause for a cool efect
Next sd
txtText.SelStart = Len(txtText)
End Sub

Public Function FileExist(ByVal FileName As String) As Boolean
'Determines if a file exists
On Error Resume Next
If Dir(FileName, vbSystem + vbHidden) = "" Then
    FileExist = False
Else
    FileExist = True
End If
End Function

Private Function FindVar(TheVar As Variant) As Boolean
'Determines if a variable exists
For x = 1 To VarNames.Count
    If VarNames(x) = TheVar Then
        FindVar = True
        Exit Function
    End If
Next x
FindVar = False
End Function

Private Function GetVar(TheVar As Variant) As Variant
'Gets a variables value
If IsNumeric(TheVar) Then GetVar = TheVar: Exit Function
If TheVar = "dpp.tick" Then GetVar = GetTickCount: Exit Function
If FindVar(TheVar) = False Then
    Select Case TheVar
        Case "dpp.ip"
            GetVar = GetIPAddress
            AddVariable "dpp.ip", GetIPAddress
            Exit Function
        Case "dpp.host"
            GetVar = GetIPHostName
            AddVariable "dpp.host", GetIPHostName
            Exit Function
        Case Else
            FlagError "Error at " & iLocation & ": undefined identifier '" & TheVar & "'.", True
    End Select
End If
For x = 1 To VarNames.Count
    If VarNames(x) = TheVar Then
        GetVar = VarData(x)
        Exit Function
    End If
Next x
End Function

Private Sub SetVar(TheVar As Variant, NewVal As Variant)
'Sets the value of a variable

'If a number...
If IsNumeric(TheVar) Then FlagError "Error at " & iLocation & ": cannot modify identifier '" & TheVar & "'"
'If variable doesn't exist...
If FindVar(TheVar) = False Then FlagError "Error at " & iLocation & ": undefined identifier '" & TheVar & "'"

If Mid(TheVar, 1, 4) = "dpp." Or TheVar = "True" Or TheVar = "False" Then
    FlagError "Error at " & iLocation & ": cannot modify identifier '" & TheVar & "'."
End If
For x = VarNames.Count To 1 Step -1
    If VarNames(x) = TheVar Then
        VarNames.Remove x
        VarData.Remove x
        VarNames.Add TheVar
        VarData.Add NewVal
        Exit Sub
    End If
Next x
End Sub

Private Sub AddVariable(VarName As Variant, VariableData As String, Optional PreSet As Boolean = True)
If IsNumeric(VarName) Then FlagError "Error at " & i & ": cannot modify identifier '" & VarName & "'.", True
If FindVar(VarName) = True Then
    FlagError "Error at unkown: cannot create identifier '" & VarName & "'.", True
Else
    VarNames.Add VarName
    'MsgBox VarName & " = " & VariableData
    If PreSet Then
        VarData.Add VariableData
    Else
        VarData.Add Solve(VariableData)
    End If
End If
End Sub

Private Function FindLabel(TheLabel As Variant) As Boolean
'Determines if a variable exists
For x = 1 To LabelNames.Count
    If LabelNames(x) = TheLabel Then
        FindLabel = True
        Exit Function
    End If
Next x
FindLabel = False
End Function

Private Function GetLabel(TheLabel As Variant) As Variant
'Gets a variables value
If IsNumeric(TheLabel) = True Then
    If TheLabel < Len(dScript) Then
        GetLabel = TheLabel
        Exit Function
    End If
End If
If FindLabel(TheLabel) = False Then FlagError "Error at " & i & ": undefined identifier '" & TheLabel & "'.", True
For x = 1 To LabelNames.Count
    If LabelNames(x) = TheLabel Then
        GetLabel = LabelData(x)
        Exit Function
    End If
Next x
End Function

Private Sub AddLabel(TheLabel As Variant, LabelPos As Variant)
If FindLabel(TheLabel) = True Then
    FlagError "Error at " & i & ": cannot create identifier '" & TheLabel & "'.", True
Else
    LabelNames.Add TheLabel
    LabelData.Add LabelPos
End If
End Sub

Private Function Eval(ByVal sFunction As String) As Boolean
'This parses a string into a left value, operator, and right value
Dim LeftVal As String, RightVal As String, Operator
Dim sChar, OpFound As Boolean
OpFound = False

    For x = 1 To Len(sFunction)
        sChar = Mid(sFunction, x, 1)
        If sChar = ">" Or sChar = "<" Or sChar = "=" Then
            Operator = Operator & sChar
            OpFound = True
        Else
            If OpFound = True Then
                RightVal = RightVal & sChar
            Else
                LeftVal = LeftVal & sChar
            End If
        End If
    Next x
    
    LeftVal = Solve(LeftVal)
    RightVal = Solve(RightVal)
    
    Eval = DoOperation2(LeftVal, Operator, RightVal)
End Function

Private Function HandleExpression(ByVal sExpression As String)
'This parses a string into a left value and right value, then
'asings the rightval to leftval
Dim LeftVal As String, RightVal As String
Dim sChar   As String, Operator As Integer
    
    If Right(sExpression, 2) = "++" Then 'incrament one
        LeftVal = Trim(Left(sExpression, Len(sExpression) - 2)) 'get the variable
        CheckIdentifier LeftVal 'make sure it's valid
        SetVar LeftVal, Val(GetVar(LeftVal)) + 1 'add one to vars current value
        Exit Function
    ElseIf Right(sExpression, 2) = "--" Then 'subtract one
        LeftVal = Trim(Left(sExpression, Len(sExpression) - 2)) 'get the variable
        CheckIdentifier LeftVal 'make sure it's valid
        SetVar LeftVal, Val(GetVar(LeftVal)) - 1 'subtract one from vars current value
        Exit Function
    End If

    'it's a normal expression
    Operator = InStr(1, sExpression, "=")   'find the equal sign
    If Operator = 0 Then    'this probably isn't an expression (maybe invalid command?)
        FlagError "Error at " & iLocation & ": syntax error: " & sExpression
        Exit Function
    End If
    
    LeftVal = Trim(Left(sExpression, Operator - 1)) 'leftval is everything before equal sign
    
    CheckIdentifier LeftVal     'check it, make sure it's a valid var (no spaces, etc)
    
    RightVal = Solve(Trim(Mid(sExpression, Operator + 1))) 'rightval is everything after equal sign (solve it)
    SetVar LeftVal, RightVal    'assign the rightval to leftval
End Function

Private Function Equation(ByVal sFunction As String) As Variant
'This sub basically looks for parentheses, and solves what's in them
Dim Paren1 As Integer, Paren2 As Integer, sChar As String
Do
    DoEvents
    For x = 1 To Len(sFunction)
        sChar = Mid(sFunction, x, 1)
        Select Case sChar
            Case Chr(34) 'Character 34 is the "
                x = InStr(x + 1, sFunction, Chr(34))
            Case "("
                Paren1 = x
            Case ")"
                Paren2 = x
                Exit For
        End Select
    Next x
    If Paren1 = 0 Then
        Exit Do
    Else
        sFunction = Mid(sFunction, 1, Paren1 - 1) & " " & Chr(34) & SolveFunction(Mid(sFunction, Paren1 + 1, Paren2 - (Paren1 + 1))) & Chr(34) & " " & Mid(sFunction, Paren2 + 1)
        Paren1 = 0
        Paren2 = 0
    End If
Loop
Equation = sFunction
End Function

Private Function GetValue(Start As Long, Parameter As String, Optional InWhat As String) As Variant
'This will search for the paramter, starting at the
'specified location, skipping things in quotes, then
'evalutate it.
If InWhat = "" Then InWhat = dScript
Dim FinalCode As String, WhereItIs As Long
'Determine the location of the parameter
WhereItIs = CodeInStr(Start, InWhat, Parameter)
If WhereItIs = 0 Then
    'Not there? Flag an error.
    FlagError "Error at " & Start & ":  expected " & Parameter & ".", True
    SetI = Start
Else
    'Get the code between the starting location and the paramter
    FinalCode = Mid(InWhat, Start, WhereItIs - Start)
    'Evalutate it
    GetValue = Solve(FinalCode)
    'Set the location
    SetI = Start + Len(FinalCode) + Len(Parameter)
End If
End Function

Private Function GetValue2(Start As Long, Parameter As String) As Variant
'This will search for the paramter, starting at the
'specified location, skipping things in quotes, but
'DOESN'T evaluate it.
Dim FinalCode As String, WhereItIs As Long
'Determine the location of the parameter
WhereItIs = CodeInStr(Start, dScript, Parameter)
If WhereItIs = 0 Then
    'Not there? Flag an error.
    FlagError "Error at " & Start & ":  expected " & Parameter & ".", True
    SetI = Start
Else
    'Get the code between the starting location and the paramter
    FinalCode = Mid(dScript, Start, WhereItIs - Start)
    GetValue2 = Trim(FinalCode)
    'Set the location
    SetI = Start + Len(FinalCode) + Len(Parameter)
End If
End Function

Private Function GetValue3(Start As Long, Parameter As String, InWhat As String) As Variant
'This will search for the paramter, starting at the
'specified location, skipping things in quotes, but
'DOESN'T evaluate it.
Dim FinalCode As Variant
Dim MakeInt As Integer
If InWhat = "" Then InWhat = dScript
If CodeInStr(Start, InWhat, Parameter) = 0 Then
    'Not there?  Flag an error.
    FlagError "Error at " & Start & ":  expected " & Parameter & ".", True
    SetX = Start
End If
For x = Start To Len(InWhat)
    If Mid(InWhat, x, Len(Parameter)) = Parameter Then
        FinalCode = Mid(InWhat, Start, x - Start)
        GetValue3 = Trim(FinalCode)
        SetX = Start + Len(FinalCode) + 1
        Exit Function
    ElseIf Mid(InWhat, x, 1) = Chr(34) Then
        Quote = InStr(x, InWhat, Chr(34))
        If Quote <> 0 Then
            x = Quote
        End If
    ElseIf Mid(InWhat, x, 1) = "(" Then
        Paren = InStr(x, InWhat, ")")
        If Paren <> 0 Then
            x = Paren
        End If
    End If
Next x
End Function

Private Function GetFunctions(sFunction As String) As Variant
'This sub searches for functions and solves them
On Error Resume Next
Dim TempFunction As String
Dim Arg1 As String, Arg2 As String, Arg3 As String
Dim Args As Variant, TempString As String
Dim x As Long, sChar As String, z As Long
For x = 1 To Len(sFunction)
    sChar = Mid(sFunction, x, 1)
    'Lower Case
    If LCase(Mid(sFunction, x, 6)) = "lcase(" Then
        x = x + 6
        TempString = GetValue3(x, ")", sFunction)
        If TempString <> "" Then
            Args = dSplit(TempString, ",")
            x = SetX
            If UBound(Args) <> 1 Then
                FlagError "Error at " & iLocation & ": wrong number of arguments for lcase()."
            Else
                For z = LBound(Args) To UBound(Args)
                    TempString = Args(z)
                    Args(z) = Solve(TempString)
                Next z
                TempFunction = TempFunction & Chr(34) & LCase(Args(1)) & Chr(34)
            End If
        End If
    'Upper Case
    ElseIf LCase(Mid(sFunction, x, 6)) = "ucase(" Then
        x = x + 6
        TempString = GetValue3(x, ")", sFunction)
        If TempString <> "" Then
            Args = dSplit(TempString, ",")
            x = SetX
            If UBound(Args) <> 1 Then
                FlagError "Error at " & iLocation & ": wrong number of arguments for ucase()."
            Else
                For z = LBound(Args) To UBound(Args)
                    TempString = Args(z)
                    Args(z) = Solve(TempString)
                Next z
                TempFunction = TempFunction & Chr(34) & UCase(Args(1)) & Chr(34)
            End If
        End If
    'Length
    ElseIf LCase(Mid(sFunction, x, 4)) = "len(" Then
        x = x + 4
        TempString = GetValue3(x, ")", sFunction)
        If TempString <> "" Then
            Args = dSplit(TempString, ",")
            x = SetX
            If UBound(Args) <> 1 Then
                FlagError "Error at " & iLocation & ": wrong number of arguments for len()."
            Else
                For z = LBound(Args) To UBound(Args)
                    TempString = Args(z)
                    Args(z) = Solve(TempString)
                Next z
                TempFunction = TempFunction & Chr(34) & Len(Args(1)) & Chr(34)
            End If
        End If
    'Is a number
    ElseIf LCase(Mid(sFunction, x, 10)) = "isnumeric(" Then
        x = x + 10
        TempString = GetValue3(x, ")", sFunction)
        If TempString <> "" Then
            Args = dSplit(TempString, ",")
            x = SetX
            If UBound(Args) <> 1 Then
                FlagError "Error at " & iLocation & ": wrong number of arguments for isnumeric()."
            Else
                For z = LBound(Args) To UBound(Args)
                    TempString = Args(z)
                    Args(z) = Solve(TempString)
                Next z
                TempFunction = TempFunction & Chr(34) & IsNumeric(Args(1)) & Chr(34)
            End If
        End If
    'Is an operator
    ElseIf LCase(Mid(sFunction, x, 5)) = "isop(" Then
        x = x + 5
        TempString = GetValue3(x, ")", sFunction)
        If TempString <> "" Then
            Args = dSplit(TempString, ",")
            x = SetX
            If UBound(Args) <> 1 Then
                FlagError "Error at " & iLocation & ": wrong number of arguments for isop()."
            Else
                For z = LBound(Args) To UBound(Args)
                    TempString = Args(z)
                    Args(z) = Solve(TempString)
                Next z
                TempFunction = TempFunction & Chr(34) & IsOperator(Args(1)) & Chr(34)
            End If
        End If
    'make integer
    ElseIf LCase(Mid(sFunction, x, 4)) = "int(" Then
        x = x + 4
        TempString = GetValue3(x, ")", sFunction)
        If TempString <> "" Then
            Args = dSplit(TempString, ",")
            x = SetX
            If UBound(Args) <> 1 Then
                FlagError "Error at " & iLocation & ": wrong number of arguments for int()."
            Else
                For z = LBound(Args) To UBound(Args)
                    TempString = Args(z)
                    Args(z) = Solve(TempString)
                Next z
                TempFunction = TempFunction & Chr(34) & Int(Args(1)) & Chr(34)
            End If
        End If
    'FileExist
    ElseIf LCase(Mid(sFunction, x, 9)) = "filexist(" Then
        x = x + 9
        TempString = GetValue3(x, ")", sFunction)
        If TempString <> "" Then
            Args = dSplit(TempString, ",")
            x = SetX
            If UBound(Args) <> 1 Then
                FlagError "Error at " & iLocation & ": wrong number of arguments for filexist()."
            Else
                For z = LBound(Args) To UBound(Args)
                    TempString = Args(z)
                    Args(z) = Solve(TempString)
                Next z
                TempFunction = TempFunction & Chr(34) & FileExist(Args(1)) & Chr(34)
            End If
        End If
    'return text length
    ElseIf LCase(Mid(sFunction, x, 8)) = "conlen()" Then
        x = x + 8
        If Not appDOS Then
            TempFunction = TempFunction & Len(txtText.Text)
        Else
            TempFunction = TempFunction & Chr(34) & "Unknown" & Chr(34)
        End If
        
    'Produce Random Number
    ElseIf LCase(Mid(sFunction, x, 7)) = "rndnum(" Then
        x = x + 7
        TempString = GetValue3(x, ")", sFunction)
        If TempString <> "" Then
            Args = dSplit(TempString, ",")
            x = SetX
            If UBound(Args) <> 1 Then
                FlagError "Error at " & iLocation & ": wrong number of arguments for rndnum()."
            Else
                For z = LBound(Args) To UBound(Args)
                    TempString = Args(z)
                    Args(z) = Solve(TempString)
                Next z
                Randomize
                TempFunction = TempFunction & Chr(34) & Int(Rnd * Int(Args(1))) & Chr(34)
            End If
        End If
        
    'val
    ElseIf LCase(Mid(sFunction, x, 4)) = "val(" Then
        x = x + 4
        TempString = GetValue3(x, ")", sFunction)
        If TempString <> "" Then
            Args = dSplit(TempString, ",")
            x = SetX
            If UBound(Args) <> 1 Then
                FlagError "Error at " & iLocation & ": wrong number of arguments for rndnum()."
            Else
                For z = LBound(Args) To UBound(Args)
                    TempString = Args(z)
                    Args(z) = Solve(TempString)
                Next z
                TempFunction = TempFunction & Chr(34) & Val(Args(1)) & Chr(34)
            End If
        End If
        
    'InStr
    ElseIf LCase(Mid(sFunction, x, 6)) = "instr(" Then
        x = x + 6
        TempString = GetValue3(x, ")", sFunction)
        If TempString <> "" Then
            Args = dSplit(TempString, ",")
            x = SetX
            If UBound(Args) <> 3 Then
                FlagError "Error at " & iLocation & ": wrong number of arguments for instr()."
            Else
                For z = LBound(Args) To UBound(Args)
                    TempString = Args(z)
                    Args(z) = Solve(TempString)
                Next z
                TempFunction = TempFunction & Chr(34) & InStr(Args(1), Args(2), Args(3)) & Chr(34)
            End If
        End If
    'Right
    ElseIf LCase(Mid(sFunction, x, 6)) = "right(" Then
        x = x + 6
        TempString = GetValue3(x, ")", sFunction)
        If TempString <> "" Then
            Args = dSplit(TempString, ",")
            x = SetX
            If UBound(Args) <> 2 Then
                FlagError "Error at " & iLocation & ": wrong number of arguments for right()."
            Else
                For z = LBound(Args) To UBound(Args)
                    TempString = Args(z)
                    Args(z) = Solve(TempString)
                Next z
                TempFunction = TempFunction & Chr(34) & Right(Args(1), Args(2)) & Chr(34)
            End If
        End If
    'Left
    ElseIf LCase(Mid(sFunction, x, 5)) = "left(" Then
        x = x + 5
        TempString = GetValue3(x, ")", sFunction)
        If TempString <> "" Then
            Args = dSplit(TempString, ",")
            x = SetX
            If UBound(Args) <> 2 Then
                FlagError "Error at " & iLocation & ": wrong number of arguments for left()."
            Else
                For z = LBound(Args) To UBound(Args)
                    TempString = Args(z)
                    Args(z) = Solve(TempString)
                Next z
                TempFunction = TempFunction & Chr(34) & Left(Args(1), Args(2)) & Chr(34)
            End If
        End If
    'Mid
    ElseIf LCase(Mid(sFunction, x, 4)) = "mid(" Then
        x = x + 4
        TempString = GetValue3(x, ")", sFunction)
        If TempString <> "" Then
            Args = dSplit(TempString, ",")
            x = SetX
            If UBound(Args) <> 3 Then
                FlagError "Error at " & iLocation & ": wrong number of arguments for mid()."
            Else
                For z = LBound(Args) To UBound(Args)
                    TempString = Args(z)
                    Args(z) = Solve(TempString)
                Next z
                TempFunction = TempFunction & Chr(34) & Mid(Args(1), Args(2), Args(3)) & Chr(34)
            End If
        End If
    Else
        TempFunction = TempFunction & sChar
    End If
Next x
GetFunctions = TempFunction
End Function

Private Function Solve(ByVal sFunction As String) As Variant
'This is the base solver
sFunction = GetFunctions(sFunction) 'First clear all functions out
sFunction = Equation(sFunction)
Solve = SolveFunction(sFunction)    'Then solve
End Function

Private Function SolveFunction(sFunction As String) As Variant
'This sub solves equations like 5 + (num1 * 3)
Dim Quote As Integer, sChar As String, variable As String
Dim Num2 As Variant, SolveOp As String, Num1
sFunction = Trim(sFunction)
For x = 1 To Len(sFunction)
    sChar = Mid(sFunction, x, 1)
    If sChar = Chr(34) Then                             'thats the "
        Quote = InStr(x + 1, sFunction, Chr(34))
        Num2 = Mid(sFunction, x + 1, Quote - (x + 1))
        x = Quote
        If SolveOp <> "" Then
            SolveFunction = DoOperation(SolveFunction, SolveOp, Num2)
            SolveOp = ""
        Else
            SolveFunction = Num2
        End If
    ElseIf IsOperator(sChar) = True Then
        If Num1 <> 0 Then
            Num2 = GetVar(Trim(Mid(sFunction, Num1, x - (Num1 + 1))))
            If SolveOp <> "" Then
                SolveFunction = DoOperation(SolveFunction, SolveOp, Num2)
                SolveOp = ""
            Else
                SolveFunction = Num2
            End If
            Num1 = 0
        End If
        SolveOp = sChar
    Else
        If Asc(sChar) <> 32 And Num1 = 0 Then Num1 = x
        If x >= Len(sFunction) Then
            Num2 = GetVar(Trim(Mid(sFunction, Num1, x)))
            If SolveOp <> "" Then
                SolveFunction = DoOperation(SolveFunction, SolveOp, Num2)
                SolveOp = ""
            Else
                SolveFunction = Num2
            End If
            Exit For
        End If
    End If
    GoTo NextX
NextX:
Next x
End Function

Private Function IsOperator(NumVal As Variant) As Boolean
'Determines if a character is an operator
Select Case NumVal
Case "+", "-", "*", "\", "/", "&", ">", "<", "="
    IsOperator = True
End Select
End Function

Private Function DoOperation(ByVal LeftVal As Variant, ByVal Operator As Variant, ByVal RightVal As Variant) As Variant
'Solves an equation
Select Case Operator
    Case "+"
        DoOperation = Val(LeftVal) + Val(RightVal)
    Case "-"
        DoOperation = Val(LeftVal) - Val(RightVal)
    Case "/"
        DoOperation = Val(LeftVal) / Val(RightVal)
    Case "\"
        DoOperation = Val(LeftVal) \ Val(RightVal)
    Case "^"
        DoOperation = Val(LeftVal) ^ Val(RightVal)
    Case "*"
        DoOperation = Val(LeftVal) * Val(RightVal)
    Case "&"
        DoOperation = LeftVal & RightVal
    Case Else
        FlagError "Error at " & iLocation & ": invalid operator, '" & Operator & "'", True
End Select
End Function

Private Function DoOperation2(ByVal LeftVal As Variant, ByVal Operator As Variant, ByVal RightVal As Variant) As Boolean
'Determines if an expression is True or False
Select Case Operator
    Case ">"
        If IsNumeric(LeftVal) Or IsNumeric(RightVal) Then
            If Val(LeftVal) > Val(RightVal) Then DoOperation2 = True
        Else
            If LeftVal > RightVal Then DoOperation2 = True
        End If
    Case "<"
        If IsNumeric(LeftVal) Or IsNumeric(RightVal) Then
            If Val(LeftVal) < Val(RightVal) Then DoOperation2 = True
        Else
            If LeftVal < RightVal Then DoOperation2 = True
        End If
    Case "="
        If IsNumeric(LeftVal) Or IsNumeric(RightVal) Then
            If Val(LeftVal) = Val(RightVal) Then DoOperation2 = True
        Else
            If LeftVal = RightVal Then DoOperation2 = True
        End If
    Case "<>"
        If IsNumeric(LeftVal) Or IsNumeric(RightVal) Then
            If Val(LeftVal) <> Val(RightVal) Then DoOperation2 = True
        Else
            If LeftVal <> RightVal Then DoOperation2 = True
        End If
    Case ">="
        If IsNumeric(LeftVal) Or IsNumeric(RightVal) Then
            If Val(LeftVal) >= Val(RightVal) Then DoOperation2 = True
        Else
            If LeftVal >= RightVal Then DoOperation2 = True
        End If
    Case "<="
        If IsNumeric(LeftVal) Or IsNumeric(RightVal) Then
            If Val(LeftVal) <= Val(RightVal) Then DoOperation2 = True
        Else
            If LeftVal <= RightVal Then DoOperation2 = True
        End If
    Case Else
        FlagError "Error at " & iLocation & ": invalid operator, '" & Operator & "'", True
End Select
End Function

Private Function CodeInStr(StartPos As Long, SourceText As String, ToFind As String) As Long
'This is just like InStr() only it skips over ""'s
Dim iPend As Long, i As Long
Dim sTemp As String

For i = StartPos To Len(SourceText)
    If LCase(Mid(SourceText, i, Len(ToFind))) = LCase(ToFind) Then
        CodeInStr = i
        Exit Function
    End If
                          
    If Mid(SourceText, i, 1) = Chr(34) Then
        If StringExist(i, SourceText, Chr(34)) Then
            i = InStr(i + 1, SourceText, Chr(34))
        Else
            FlagError "Error at " & i & ": Expected '""""'"
        End If
    End If
Next i
End Function

Private Function StringExist(StartPos As Long, SourceText As String, ToFind As String) As Boolean
    'This determines if a string exists in another string
If CodeInStr(StartPos, LCase(SourceText), LCase(ToFind)) = 0 Then
    StringExist = False
Else
    StringExist = True
End If
End Function

Public Function ReadEXE(sPath As String) As String
'Reads an EXE file and gets D++ source
On Error GoTo Errorh
Dim FileData As String
Dim Header As String, EXE_Type As String, Encrypt_Source As String
    
Open sPath For Binary As #1 'open file
    FileData = Space$(LOF(1)) 'set file data length
    Get #1, , FileData 'get file data
Close #1 'close file

If InStr(1, FileData, "dppapp:") = 0 Or InStr(1, FileData, "dpp:") = 0 Then
    ReadEXE = "screenput " & Chr(34) & "Note: No source code found." & Chr(34) & ";"
    Exit Function
End If

Header = Mid(FileData, InStr(1, FileData, "dppapp:"), 35)
EXE_Type = Mid(Header, 13, 3)
Encrypt_Source = Mid(Header, 23, 1)

Select Case EXE_Type
    Case "dos"
        appDOS = True
    Case "dpp"
        appDOS = False
    Case "reg"
        SaveSetting "D++", "Version", "DLL_Version", App.Revision
        End
    Case Else
        MsgBox "Fatal Link Error: Incorrect application type.  Please recompile." & vbCrLf & vbCrLf & "If this message continues, please download the latest D++ compiler at " & vbCrLf & vbCrLf & "http://squeakmac.tripod.com", vbCritical, "Invalid App Type"
        End
End Select

If Encrypt_Source = "t" Then 'Parse file to get source
    ReadEXE = Crypt(Mid(FileData, InStr(1, FileData, "dpp:") + 4))
ElseIf Encrypt_Source = "f" Then
    ReadEXE = Mid(FileData, InStr(1, FileData, "dpp:") + 4)
Else
    MsgBox "Fatal Link Error: Incorrect application type.  Please recompile." & vbCrLf & vbCrLf & "If this message continues, please download the latest D++ compiler at " & vbCrLf & vbCrLf & "http://squeakmac.tripod.com", vbCritical, "Invalid App Type"
    End
End If

Exit Function


Errorh:
If Err.Number <> 0 Then MsgBox "Link Error #" & Err.Number & " has occured: " & Err.Description, vbCritical, "Link Error"
End
End Function

Public Function AppPath() As String
'This gets the app's path
If Right(App.Path, 1) = "\" Then AppPath = App.Path Else AppPath = App.Path & "\"
End Function

Function Crypt(Text) As String
'This is a simple crypter
For i = 1 To Len(Text)
    DoEvents
    Crypt = Crypt & Chr(255 - Asc(Mid(Text, i, 1)))
Next i
End Function

Private Function FindLoopEnd(StartPos As Long) As Long
Dim z As Long, LoopAt As Long, Quote As Long
For z = StartPos To Len(dScript)
    If Mid(dScript, z, 4) = "loop" Then
        If LoopAt = 0 Then
            FindLoopEnd = z + 4
            Exit Function
        Else
            LoopAt = LoopAt + 1
        End If
    ElseIf Mid(dScript, z, 8) = "do until" Then
        LoopAt = LoopAt - 1
    ElseIf Mid(dScript, z, 8) = "do while" Then
        LoopAt = LoopAt - 1
    ElseIf Mid(dScript, z, 1) = Chr(34) Then
        Quote = InStr(z + 1, dScript, Chr(34))
        If Quote <> 0 Then z = Quote
    End If
Next z
FlagError "Error at " & StartPos & ": expected 'loop'."
End Function

Private Function FindLoopStart(StartPos As Long) As Long
Dim z As Long, LoopAt As Long, Quote As Boolean
LoopAt = 0
For z = StartPos To 1 Step -1
    If Mid(dScript, z, 8) = "do until" Then
        If Quote Then GoTo SkipCurrentZ
        If LoopAt = 0 Then
            FindLoopStart = z - 1
            Exit Function
        Else
            LoopAt = LoopAt - 1
        End If
    ElseIf Mid(dScript, z, 8) = "do while" Then
        If Quote Then GoTo SkipCurrentZ
        If LoopAt = 0 Then
            FindLoopStart = z - 1
            Exit Function
        Else
            LoopAt = LoopAt - 1
        End If
    ElseIf Mid(dScript, z, 4) = "loop" Then
        If Quote Then GoTo SkipCurrentZ
        LoopAt = LoopAt + 1
    ElseIf Mid(dScript, z, 1) = Chr(34) Then
        If Quote Then
            Quote = False
        Else
            Quote = True
        End If
    End If
SkipCurrentZ:
Next z
FlagError "Error at " & StartPos & ": loop without do"
End Function

Private Function FindIdentifier(StartPos As Long) As String
Dim i As Long
For i = StartPos To 1 Step -1
    Select Case Mid(dScript, i - 1, 1)
        Case " ", ";", vbTab
            If i > StartPos Then GoTo SkipCurrentI
            FindIdentifier = Trim(Mid(dScript, i, StartPos - (i - 1)))
            If FindIdentifier <> "" Then Exit Function
    End Select
    If Mid(dScript, i - 2, 2) = vbCrLf Then
        If i > StartPos Then GoTo SkipCurrentI
        FindIdentifier = Trim(Mid(dScript, i, StartPos - (i - 1)))
        If FindIdentifier <> "" Then Exit Function
        Exit Function
    End If
SkipCurrentI:
Next i
FindIdentifier = Trim(Mid(dScript, 1, Len(dScript) - StartPos))
End Function

Public Function dSplit(Expression As String, Delimiter As String) As Variant
    ReDim SplitArray(1 To 1) As Variant
    Dim TempLetter As String
    Dim TempSplit As String
    Dim i As Integer
    Dim x As Integer
    Dim StartPos As Integer
    
    Expression = Expression & Delimiter
    For i = 1 To Len(Expression)
        If Mid(Expression, i, 1) = Chr(34) Then
            If InStr(i + 1, Expression, Chr(34)) = 0 Then
                FlagError "Error at " & iLocation & ": expected """
                i = i + 1
            Else
                i = InStr(i + 1, Expression, Chr(34))
            End If
        End If
        TempLetter = Mid(Expression, i, Len(Delimiter))
        If TempLetter = Delimiter Then
            TempSplit = Mid(Expression, (StartPos + 1), (i - StartPos) - 1)
            If TempSplit <> "" Then
                x = x + 1
                ReDim Preserve SplitArray(1 To x) As Variant
                SplitArray(x) = TempSplit
            End If
            StartPos = i
        End If
    Next i
    dSplit = SplitArray
End Function

Private Function FindIfEnd(StartPos As Long, UseElse As Boolean) As Long
Dim z As Long, IfAt As Long, Quote As Long
If UseElse = False Then GoTo EndIfs
For z = StartPos To Len(dScript) 'first look for an else for this if statement
    If LCase(Mid(dScript, z, 4)) = "else" Then
        If IfAt = 0 Then
            FindIfEnd = z + 3
            InIf = True
            Exit Function
        End If
    ElseIf LCase(Mid(dScript, z, 2)) = "if" Then
        If LCase(Mid(dScript, z - 3, 5)) <> "endif" Then 'make sure it's not in other statements
            If LCase(Mid(dScript, z - 4, 6)) <> "elseif" Then
                IfAt = IfAt + 1
            End If
        Else
            If IfAt = 0 Then
                FindIfEnd = z + 1
                InIf = True
                Exit Function
            Else
                IfAt = IfAt - 1
            End If
        End If
    ElseIf Mid(dScript, z, 1) = Chr(34) Then
        Quote = InStr(z + 1, dScript, Chr(34))
        If Quote <> 0 Then z = Quote
    End If
Next z
EndIfs:
IfAt = 0
For z = StartPos To Len(dScript) 'couldn't find an else statement in this if, goto endif
    If LCase(Mid(dScript, z, 5)) = "endif" Then
        If IfAt = 0 Then
            FindIfEnd = z + 6
            InIf = False
            Exit Function
        Else
            IfAt = IfAt - 1
        End If
    ElseIf LCase(Mid(dScript, z, 2)) = "if" Then
        If LCase(Mid(dScript, z - 3, 5)) <> "endif" Then 'make sure it's not in other statements
            If LCase(Mid(dScript, z - 4, 6)) <> "elseif" Then
                IfAt = IfAt + 1
            End If
        End If
    ElseIf Mid(dScript, z, 1) = Chr(34) Then
        Quote = InStr(z + 1, dScript, Chr(34))
        If Quote <> 0 Then z = Quote
    End If
Next z
FlagError "Error at " & StartPos & ": expected 'endif'"
End Function

Private Function GetCurrentLoopData(Data As Integer) As Variant
Dim TempArray As Variant

TempArray = Split(ForLoops(ForLoopDepth), ":")
Select Case Data
    Case 0
        GetCurrentLoopData = TempArray(0)
    Case 1
        GetCurrentLoopData = TempArray(1)
    Case 2
        GetCurrentLoopData = TempArray(2)
    Case 3
        GetCurrentLoopData = TempArray(3)
End Select
End Function


Private Function FindForEnd(StartPos As Long, ByVal Var As String) As Long
Dim x As Long, CurrentVar As String

x = StartPos

search:
x = CodeInStr(x + 1, dScript, "next ")
If x = 0 Then FlagError "Error at " & StartPos & ": expected 'next'.", True
    
CurrentVar = GetValue2(x + 5, ";")
If CurrentVar = Var Then
    FindForEnd = x + Len(Var) + 1
    Exit Function
Else
    GoTo search
End If
End Function

Private Function CheckIdentifier(Identifier As String) As Boolean
Dim Text As String
Text = Trim(Identifier)
If Text = "" Then Exit Function
If _
StringExist(1, Text, "screenout ") = True Or StringExist(1, Text, "screenput ") = True Or _
StringExist(1, Text, "screenin ") = True Or StringExist(1, Text, "screenpass ") = True Or _
StringExist(1, Text, "title ") = True Or StringExist(1, Text, "delete ") = True Or _
StringExist(1, Text, "box ") = True Or StringExist(1, Text, "pause ") = True Or _
StringExist(1, Text, "web ") = True Or StringExist(1, Text, "open ") = True Or _
StringExist(1, Text, "label ") = True Or StringExist(1, Text, "goto ") = True Or _
StringExist(1, Text, "newvar ") = True Or StringExist(1, Text, "do until ") = True Or _
StringExist(1, Text, "if ") = True Or StringExist(1, Text, "set ") = True Or _
StringExist(1, Text, "wav ") = True Or StringExist(1, Text, "clear;") = True Or _
StringExist(1, Text, "hide;") = True Or StringExist(1, Text, "show;") = True Or _
StringExist(1, Text, "open_cd;") = True Or StringExist(1, Text, "close_cd;") = True Or _
StringExist(1, Text, "enable_cad;") = True Or StringExist(1, Text, "disable_cad;") = True Or _
StringExist(1, Text, "show_controls;") = True Or StringExist(1, Text, "hide_controls;") = True Or _
StringExist(1, Text, "end;") = True Or StringExist(1, Text, "finsih;") = True Or _
StringExist(1, Text, "doevents;") = True Or StringExist(1, Text, "screen;") = True Then

    FlagError "Error at " & iLocation & ": expression syntax error"
    Exit Function
ElseIf StringExist(1, Text, " ") = True Or StringExist(1, Text, vbCrLf) = True Or _
       StringExist(1, Text, "(") Then
    FlagError "Error at " & iLocation & ": expression syntax error"
    Exit Function
End If
CheckIdentifier = True
End Function

