VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   BackColor       =   &H00C0C0C0&
   Caption         =   "D++ IDE"
   ClientHeight    =   7065
   ClientLeft      =   3555
   ClientTop       =   2280
   ClientWidth     =   7800
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   Picture         =   "frmMain.frx":030A
   ScaleHeight     =   7065
   ScaleWidth      =   7800
   Begin VB.TextBox txtDebug 
      Height          =   1215
      Left            =   960
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   9
      Top             =   4800
      Visible         =   0   'False
      Width           =   6735
   End
   Begin VB.ListBox lstPos 
      Height          =   645
      Left            =   6360
      TabIndex        =   8
      Top             =   6120
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.TextBox txtText 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3975
      Left            =   960
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   7
      Top             =   720
      Width           =   6735
   End
   Begin VB.TextBox txtGoto 
      Height          =   285
      Left            =   2160
      TabIndex        =   5
      Top             =   360
      Width           =   5535
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   3
      Top             =   6810
      Width           =   7800
      _ExtentX        =   13758
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Bevel           =   2
            Object.Width           =   2541
            MinWidth        =   2541
            Picture         =   "frmMain.frx":064C
            Text            =   "D++ IDE"
            TextSave        =   "D++ IDE"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   7839
            MinWidth        =   7839
            Text            =   "Char: 0"
            TextSave        =   "Char: 0"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            Object.Width           =   1764
            MinWidth        =   1764
            TextSave        =   "1:30 PM"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MouseIcon       =   "frmMain.frx":0B90
   End
   Begin VB.PictureBox picTab 
      Align           =   3  'Align Left
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      Height          =   6480
      Left            =   0
      ScaleHeight     =   6480
      ScaleWidth      =   840
      TabIndex        =   0
      Top             =   330
      Width           =   840
      Begin VB.Image picDPP 
         Height          =   480
         Left            =   230
         Picture         =   "frmMain.frx":0EAA
         Top             =   120
         Width           =   480
      End
      Begin VB.Label lblArea 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "D++ IDE Code Window"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   675
         Left            =   0
         TabIndex        =   1
         Top             =   600
         Width           =   825
         WordWrap        =   -1  'True
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   960
      Top             =   5880
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   11
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":11B4
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":12C8
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":13DC
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":14F0
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1604
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1828
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1A28
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1B3C
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1C40
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1D54
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1E68
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   1560
      Top             =   5880
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   330
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   7800
      _ExtentX        =   13758
      _ExtentY        =   582
      ButtonWidth     =   609
      ButtonHeight    =   582
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   13
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "New"
            Object.ToolTipText     =   "New"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Save"
            Object.ToolTipText     =   "Save"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Open"
            Object.ToolTipText     =   "Open"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Copy"
            Object.ToolTipText     =   "Copy"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Paste"
            Object.ToolTipText     =   "Paste"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Cut"
            Object.ToolTipText     =   "Cut"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Run"
            Object.ToolTipText     =   "Run"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "Stop"
            Object.ToolTipText     =   "Stop"
            ImageIndex      =   11
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Compile"
            Object.ToolTipText     =   "Compile"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Help"
            Object.ToolTipText     =   "Help"
            ImageIndex      =   10
         EndProperty
      EndProperty
      Begin VB.Label Label2 
         BackColor       =   &H00000000&
         Caption         =   "Label1"
         Height          =   85511
         Left            =   0
         TabIndex        =   4
         Top             =   0
         Width           =   1095
      End
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Goto Character:"
      Height          =   195
      Left            =   960
      TabIndex        =   6
      Top             =   405
      Width           =   1125
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileNew 
         Caption         =   "&New"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuFileLine1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileOpen 
         Caption         =   "&Open"
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuFileClose 
         Caption         =   "&Close"
      End
      Begin VB.Menu mnuFileLine2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileSave 
         Caption         =   "&Save"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuFileSaveAS 
         Caption         =   "Save &As..."
      End
      Begin VB.Menu mnuFileLine3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileLastRun 
         Caption         =   "&1 Last Run"
         Shortcut        =   ^R
      End
      Begin VB.Menu mnuFileLastCompile 
         Caption         =   "&2 Last Compile"
      End
      Begin VB.Menu mnuFileLine4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
         Shortcut        =   ^Q
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
      Begin VB.Menu mnuEditCut 
         Caption         =   "C&ut"
      End
      Begin VB.Menu mnuEditCopy 
         Caption         =   "&Copy"
      End
      Begin VB.Menu mnuEditPaste 
         Caption         =   "&Paste"
      End
      Begin VB.Menu mnulne 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditFindReplace 
         Caption         =   "Find/Replace"
         Shortcut        =   ^F
      End
      Begin VB.Menu mnuTimeDate 
         Caption         =   "Time/Date"
      End
      Begin VB.Menu mnuLine5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFindNext 
         Caption         =   "Find Next"
         Shortcut        =   {F3}
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "&View"
      Begin VB.Menu mnuViewDebug 
         Caption         =   "&Debug Window"
         Shortcut        =   {F2}
      End
      Begin VB.Menu mnuViewCalc 
         Caption         =   "&Calculator"
         Shortcut        =   {F9}
      End
      Begin VB.Menu mnuViewDownload 
         Caption         =   "&Download Latest Files"
         Shortcut        =   {F4}
      End
      Begin VB.Menu mnuViewLine1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewOptions 
         Caption         =   "&Options..."
         Shortcut        =   {F12}
      End
   End
   Begin VB.Menu mnuProject 
      Caption         =   "&Project"
      Begin VB.Menu mnuProjectRun 
         Caption         =   "&Run"
         Shortcut        =   {F5}
      End
      Begin VB.Menu mnuProjectCompile 
         Caption         =   "&Make EXE..."
         Shortcut        =   {F6}
      End
      Begin VB.Menu mnuLine 
         Caption         =   "-"
      End
      Begin VB.Menu mnuProjectRunDOS 
         Caption         =   "Run in &DOS"
         Shortcut        =   {F7}
      End
      Begin VB.Menu mnuCompileDOS 
         Caption         =   "Compile in D&OS"
         Shortcut        =   {F8}
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelpHelp 
         Caption         =   "&Help"
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnuHelpSource 
         Caption         =   "&Source Code"
      End
      Begin VB.Menu mnuHelpLine 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "&About D++ IDE..."
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'D++ IDE
'If you use any of this code, please give me credit for what you used.
'By SqueakMac (squeak5@mediaone.net)
'http://squeakmac.tripod.com

Dim DebugActive As Boolean, Save As Boolean
Dim CurrentSearch As String

Private Sub Form_Load()
Me.Show
SetDllLocation 'Set the DLL location (see modMain)
DebugActive = False
Save = False

If GetSetting("D++", "Reg", "UserName") = "" Then 'If user is not registered, register
    frmReg.Show 1
End If

If Command$ = "" Then Exit Sub 'if no command, exit
If LCase(GetExtension(Command$)) = "exe" Then 'if it's an EXE
    txtText.Text = ReadEXE(Command$) 'read source
    Me.Caption = "D++ IDE - [" & GetFileName(Command$) & "]" 'display path
Else
    txtText.Text = ReadFile(Command$) 'Read the command file
    Me.Caption = "D++ IDE - [" & GetFileName(Command$) & "]" 'display path
End If
End Sub

Private Sub Form_Resize()
'Resizes controls
On Error Resume Next
'Line1.X2 = Width
'Line2.X2 = Width
If DebugActive = True Then
    txtDebug.Top = ScaleHeight - txtDebug.Height - 255
    txtDebug.Width = ScaleWidth - 955
    txtText.Width = ScaleWidth - 955
    txtText.Height = ScaleHeight - (txtDebug.Height + 150) - 850
Else
    txtText.Width = ScaleWidth - 955
    txtText.Height = ScaleHeight - (Toolbar1.Height + 50) - 600
End If
txtGoto.Width = ScaleWidth - Label1.Width - 1050
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
End
End Sub

Private Sub txtDebug_Click()
On Error Resume Next
Dim LineStart As Integer
Dim LineArray As Variant
Dim LineNum As Integer
Dim ErrorMesage As String
LineNum = GetLineNum(txtDebug.Text, txtDebug.SelStart)

If lstPos.List(LineNum - 1) = 0 Then Exit Sub
If txtDebug.SelLength > 1 Then Exit Sub

txtText.SetFocus
txtText.SelStart = lstPos.List(LineNum - 1)
txtText.SelLength = 1

LineArray = Split(txtDebug.Text, vbCrLf)
ErrorMessage = Mid(LineArray(LineNum - 1), InStr(1, LineArray(LineNum - 1), ":") + 2)

StatusBar1.Panels(2).Text = ErrorMessage

Exit Sub
'the following highlights the error in the debug window
'however, it's useless because the focus must be on the
'code window. =)  Silly me....

For i = txtDebug.SelStart To 1 Step -1
    If Mid(txtDebug.Text, i, 2) = vbCrLf Then
        LineStart = i
        GoTo HighlightLabel
    End If
Next i
LineStart = 0

HighlightLabel:
LineArray = Split(txtDebug.Text, vbCrLf)
txtDebug.SelStart = LineStart
txtDebug.SelLength = Len(LineArray(LineNum - 1)) + 1
End Sub

Private Sub mnuCompileDOS_Click()
Toolbar1.Buttons(10).Enabled = True
DebugActive = True
txtDebug.Visible = True
mnuViewDebug.Checked = True
Form_Resize 'Resize Form
txtDebug.Text = ""
lstPos.Clear
AddDebug ">D++ IDE DOS Debug"
AddDebug ">"
MakeEXE True 'Make the EXE (see modMain)
Toolbar1.Buttons(10).Enabled = False
End Sub

Private Sub mnuEditCopy_Click()
On Error Resume Next
Clipboard.SetText txtText.SelText
End Sub

Private Sub mnuEditCut_Click()
On Error Resume Next
Clipboard.SetText txtText.SelText
txtText.SelText = ""
End Sub

Private Sub mnuEditFindReplace_Click()
frmFind.Show
End Sub

Private Sub mnuEditPaste_Click()
On Error Resume Next
txtText.SelText = Clipboard.GetText
End Sub

Private Sub mnuFileClose_Click()
Save = False
Me.Caption = "D++ IDE"
txtText.Text = ""
End Sub

Private Sub mnuFileExit_Click()
End
End Sub

Private Sub mnuFileLastCompile_Click()
On Error GoTo Errorh
Dim CompileLocation As String
CompileLocation = GetSetting("D++", "Memory", "Compile") 'Get last compiled file

If FileExist(CompileLocation) = False Then
    MsgBox "Last compiled file not found!", vbExclamation, "Compile File not found!"
    Exit Sub
Else
    txtText.Text = ReadEXE(CompileLocation) 'Read the EXE (see modMain)
    'Colorize txtText, &H8000&, &H808080, &HC00000
    Me.Caption = "D++ IDE - [" & GetFileName(CompileLocation) & "]"
    Save = False
End If
Errorh:
If Err.Number = "0" Then Exit Sub
MsgBox "Error #" & Err.Number & " has occured: " & Err.Description, vbCritical, "Error"
End Sub

Private Sub mnuFileLastRun_Click()
'On Error GoTo Errorh
Dim RunLocation As String
RunLocation = GetSetting("D++", "Memory", "Run") 'Get last run file

If FileExist(RunLocation) = False Then
    MsgBox "D++APP1.EXE could not be found at the specified run location.  Run your program first, then try again.", vbExclamation, "Run File not found!"
    Exit Sub
Else
    txtText.Text = ReadEXE(RunLocation) 'Read EXE (see modMain)
    'Colorize txtText, &H8000&, &H808080, &HC00000
    Me.Caption = "D++ IDE - [D++APP1.EXE]"
    Save = False
End If
Errorh:
If Err.Number = "0" Then Exit Sub
MsgBox "Error #" & Err.Number & " has occured: " & Err.Description, vbCritical, "Error"
End Sub

Private Sub mnuFileNew_Click()
Save = False
Me.Caption = "D++ IDE"
txtText.Text = ""
End Sub

Private Sub mnuFileOpen_Click()
Dim Ext As String, ans
On Error GoTo errh
CommonDialog1.Filter = "D++ Source Files (*.dpp)|*.dpp|D++ EXE Files (*.exe)|*.exe|C++ Files (*.cpp)|*.cpp|All Files (*.*)|*.*"
CommonDialog1.DialogTitle = "Open D++ File"
CommonDialog1.CancelError = True
CommonDialog1.ShowOpen
Ext = GetExtension(CommonDialog1.FileName)
Select Case LCase(Ext)
    Case "exe"
        txtText.Text = ReadEXE(CommonDialog1.FileName)
    'Case "cpp"
    '    ans = MsgBox("You are opening a C++ file.  Would you like it converted to D++?", vbQuestion + vbYesNo, "Convert")
    '    If ans = vbYes Then
    '        'MsgBox "D++ does not support functions, structs, or classes.  The will be removed from the program.", vbExclamation, "Convert"
    '        frmConvert.ConvertToDPP CommonDialog1.FileName
    '    Else
    '        txtText.Text = ReadFile(CommonDialog1.FileName)
    '    End If
    Case "dpp", "txt"
        txtText.Text = ReadFile(CommonDialog1.FileName)
    Case Else
        MsgBox "D++ IDE does not support extension '" & Ext & "'.", vbCritical, "Extension"
End Select
'Colorize txtText, &H8000&, &H808080, &HC00000
Me.Caption = "D++ IDE - [" & CommonDialog1.FileTitle & "]"
Save = True
errh:
End Sub

Private Sub mnuFindNext_Click()
If CurrentSearch = "" Then Exit Sub
FindIt CurrentSearch
End Sub

Private Sub mnuHelpSource_Click()
frmCode.Show
End Sub

Private Sub mnuProjectRunDOS_Click()
Toolbar1.Buttons(10).Enabled = True
DebugActive = True
txtDebug.Visible = True
mnuViewDebug.Checked = True
Form_Resize 'Resize Form
txtDebug.Text = ""
lstPos.Clear
AddDebug ">D++ IDE DOS Debug"
AddDebug ">"
Run True 'run in dos (see modMain)
Toolbar1.Buttons(10).Enabled = False
End Sub

Private Sub mnuFileSave_Click()
On Error Resume Next
If Save = True Then 'If already saved...
    FileNum = FreeFile
    Open CommonDialog1.FileName For Output As #FileNum 'write file
        Print #FileNum, txtText.Text
    Close #FileNum
    Me.Caption = "D++ IDE - [" & CommonDialog1.FileTitle & "]"
Else
    mnuFileSaveAS_Click 'If not saved, save it
End If
End Sub

Private Sub mnuFileSaveAS_Click()
On Error GoTo endit
CommonDialog1.Filter = "D++ Files (*.dpp)|*.dpp|All Files (*.*)|*.*"
CommonDialog1.DialogTitle = "Save D++ File"
CommonDialog1.CancelError = True
CommonDialog1.ShowSave
FileNum = FreeFile
If FileExist(CommonDialog1.FileName) Then
    Overwrite = MsgBox("File Exists!  Overwrite?", 276, "File Found!")
    If Overwrite = 6 Then
        Open CommonDialog1.FileName For Output As #FileNum
        Print #FileNum, txtText.Text
        Close #FileNum
    Else
        Exit Sub
    End If
Else
    Open CommonDialog1.FileName For Output As #FileNum
    Print #FileNum, txtText.Text
    Close #FileNum
End If
Me.Caption = "D++ IDE - [" & CommonDialog1.FileTitle & "]"
Save = True
endit:
End Sub

Private Sub mnuHelpAbout_Click()
frmAbout.Show 1
End Sub

Private Sub mnuHelpHelp_Click()
frmHelp.Show 'Show help
If txtText.SelText <> "" Then 'If text is selected
    frmHelp.txtSearch.Text = txtText.SelText 'Make search text selected text
    frmHelp.search 'Search
End If
End Sub

Private Sub mnuProjectCompile_Click()
Toolbar1.Buttons(10).Enabled = True
DebugActive = True
txtDebug.Visible = True
mnuViewDebug.Checked = True
Form_Resize 'Resize Form
txtDebug.Text = ""
lstPos.Clear
AddDebug ">D++ IDE Debug"
AddDebug ">"
MakeEXE False 'Make the EXE (see modMain)
Toolbar1.Buttons(10).Enabled = False
End Sub

Private Sub mnuProjectRun_Click()
Toolbar1.Buttons(10).Enabled = True
DebugActive = True
txtDebug.Visible = True
mnuViewDebug.Checked = True
Form_Resize 'Resize form
txtDebug.Text = ""
lstPos.Clear
AddDebug ">D++ IDE Debug"
AddDebug ">"
Run False 'Run file (see modMain)
Toolbar1.Buttons(10).Enabled = False
End Sub

Private Sub mnuTimeDate_Click()
txtText.SelText = Time & "/" & Date
End Sub

Private Sub mnuViewCalc_Click()
frmCalc.Show
End Sub

Private Sub mnuViewDebug_Click()
If mnuViewDebug.Checked = True Then
    DebugActive = False
    txtDebug.Visible = False
    Form_Resize
    mnuViewDebug.Checked = False
Else
    DebugActive = True
    txtDebug.Visible = True
    Form_Resize
    mnuViewDebug.Checked = True
End If
End Sub

Private Sub mnuViewDownload_Click()
frmDownload.Show 1
End Sub

Private Sub mnuViewOptions_Click()
frmOptions.Show 1
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
        Case "New"
            mnuFileNew_Click
        Case "Open"
            mnuFileOpen_Click
        Case "Copy"
            mnuEditCopy_Click
        Case "Paste"
            mnuEditPaste_Click
        Case "Cut"
            mnuEditCut_Click
        Case "Help"
            mnuHelpHelp_Click
        Case "Run"
            mnuProjectRun_Click
        Case "Stop"
            StopDebug = True
        Case "Compile"
            mnuProjectCompile_Click
    End Select
End Sub

Private Sub txtGoto_KeyPress(KeyAscii As Integer)
On Error Resume Next
If KeyAscii = vbKeyReturn Then
    txtText.SetFocus
    txtText.SelStart = txtGoto.Text - 1
    txtText.SelLength = 1
End If
End Sub

Private Sub txtText_Change()
If txtText.SelLength > 1 Then
    StatusBar1.Panels(2).Text = "Select Length: " & txtText.SelLength
Else
    StatusBar1.Panels(2).Text = "Char: " & txtText.SelStart + 1
End If
End Sub

Private Sub txtText_Click()
If txtText.SelLength > 1 Then
    StatusBar1.Panels(2).Text = "Select Length: " & txtText.SelLength
Else
    StatusBar1.Panels(2).Text = "Char: " & txtText.SelStart + 1
End If
End Sub

Private Sub txtText_KeyPress(KeyAscii As Integer)
'If KeyAscii = vbKeyReturn Then Colorize txtText, &H8000&, &H808080, &HC00000
'txtText.SelColor = vbBlack
End Sub

Public Sub FindIt(FindWhat As String)
If Len(txtText.Text) = 0 Then
    MsgBox "Search text not found.", vbExclamation, "Search"
    Exit Sub
End If
search:
If txtText.SelLength > 0 Then
    txtText.SelStart = txtText.SelStart + txtText.SelLength
End If
If txtText.SelStart = 0 Then txtText.SelStart = 1
i = InStr(txtText.SelStart, LCase(txtText.Text), LCase(FindWhat))
If i <> 0 Then
    txtText.SelStart = i - 1
    txtText.SelLength = Len(FindWhat)
    CurrentSearch = FindWhat
Else
    If InStr(1, LCase(txtText.Text), LCase(FindWhat)) = 0 Then
        MsgBox "Search text not found.", vbExclamation, "Search"
    Else
        txtText.SelStart = 1
        GoTo search
    End If
End If
End Sub

Public Sub ResumeIDE()
StopDebug = False
Toolbar1.Buttons(10).Enabled = False
End Sub

Private Sub txtText_KeyUp(KeyCode As Integer, Shift As Integer)
If txtText.SelLength > 1 Then
    StatusBar1.Panels(2).Text = "Select Length: " & txtText.SelLength
Else
    StatusBar1.Panels(2).Text = "Char: " & txtText.SelStart + 1
End If
End Sub

Private Sub txtText_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
If txtText.SelLength > 1 Then
    StatusBar1.Panels(2).Text = "Select Length: " & txtText.SelLength
End If
End Sub
