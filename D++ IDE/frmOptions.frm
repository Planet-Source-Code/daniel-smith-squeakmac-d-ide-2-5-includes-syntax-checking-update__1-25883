VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmOptions 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "D++ IDE Options"
   ClientHeight    =   5595
   ClientLeft      =   5490
   ClientTop       =   3600
   ClientWidth     =   5130
   Icon            =   "frmOptions.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5595
   ScaleWidth      =   5130
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTab1 
      Height          =   4935
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   4935
      _ExtentX        =   8705
      _ExtentY        =   8705
      _Version        =   393216
      Style           =   1
      Tab             =   1
      TabHeight       =   520
      TabCaption(0)   =   "Run/Compile"
      TabPicture(0)   =   "frmOptions.frx":030A
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "frRun"
      Tab(0).Control(1)=   "Frame3"
      Tab(0).Control(2)=   "Frame1"
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "Download"
      TabPicture(1)   =   "frmOptions.frx":0326
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Image1"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Frame2"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Frame4"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).ControlCount=   3
      TabCaption(2)   =   "IDE"
      TabPicture(2)   =   "frmOptions.frx":0342
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Image4"
      Tab(2).Control(1)=   "lblCopyright"
      Tab(2).Control(2)=   "lblEncrypt"
      Tab(2).Control(3)=   "lblDLL"
      Tab(2).Control(4)=   "lblDecompile"
      Tab(2).Control(5)=   "chkDebug"
      Tab(2).Control(6)=   "cmdView"
      Tab(2).Control(7)=   "cmdAssociate"
      Tab(2).Control(8)=   "chkEncrypt"
      Tab(2).Control(9)=   "chkDecompile"
      Tab(2).ControlCount=   10
      Begin VB.CheckBox chkDecompile 
         Caption         =   "Prevent Decompile"
         Height          =   195
         Left            =   -73680
         TabIndex        =   26
         Top             =   1320
         Width           =   1695
      End
      Begin VB.CheckBox chkEncrypt 
         Caption         =   "Encrypt EXE"
         Height          =   195
         Left            =   -73680
         TabIndex        =   24
         Top             =   960
         Width           =   1215
      End
      Begin VB.CommandButton cmdAssociate 
         Caption         =   "Associate .dpp files"
         Height          =   375
         Left            =   -73680
         TabIndex        =   23
         Top             =   2015
         Width           =   1815
      End
      Begin VB.CommandButton cmdView 
         Caption         =   "View DLL Information"
         Height          =   375
         Left            =   -73680
         TabIndex        =   22
         Top             =   1560
         Width           =   1815
      End
      Begin VB.CheckBox chkDebug 
         Caption         =   "Enable Debugging"
         Height          =   255
         Left            =   -73680
         TabIndex        =   20
         ToolTipText     =   "Don't worry about this"
         Top             =   600
         Value           =   1  'Checked
         Width           =   1695
      End
      Begin VB.Frame Frame4 
         Caption         =   "When downloaded DLL has older version number,"
         Height          =   1695
         Left            =   240
         TabIndex        =   18
         Top             =   1680
         Width           =   4455
         Begin VB.OptionButton opOriginal 
            Caption         =   "Use original DLL"
            Height          =   255
            Left            =   120
            TabIndex        =   30
            Top             =   1320
            Value           =   -1  'True
            Width           =   1935
         End
         Begin VB.OptionButton opDownload 
            Caption         =   "Use downloaded DLL"
            Height          =   255
            Left            =   120
            TabIndex        =   29
            Top             =   960
            Width           =   1935
         End
         Begin VB.Label Label2 
            Caption         =   "If you are not sure, select use original DLL."
            Height          =   495
            Left            =   2520
            TabIndex        =   31
            Top             =   1080
            Width           =   1815
         End
         Begin VB.Label Label1 
            Caption         =   $"frmOptions.frx":035E
            Height          =   675
            Left            =   120
            TabIndex        =   19
            Top             =   240
            Width           =   4200
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Overwrite During Compile"
         Height          =   1335
         Left            =   -74760
         TabIndex        =   14
         Top             =   480
         Width           =   2175
         Begin VB.OptionButton opcPrompt 
            Caption         =   "Prompt"
            Height          =   255
            Left            =   120
            TabIndex        =   17
            Top             =   960
            Value           =   -1  'True
            Width           =   855
         End
         Begin VB.OptionButton opcNever 
            Caption         =   "Never Overwrite"
            Height          =   255
            Left            =   120
            TabIndex        =   16
            Top             =   600
            Width           =   1575
         End
         Begin VB.OptionButton opcAlways 
            Caption         =   "Always Overwrite"
            Height          =   255
            Left            =   120
            TabIndex        =   15
            Top             =   240
            Width           =   1575
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "After Download"
         Height          =   975
         Left            =   240
         TabIndex        =   11
         Top             =   480
         Width           =   3615
         Begin VB.OptionButton opNoDisplay 
            Caption         =   "Don't Display DLL Information"
            Height          =   255
            Left            =   120
            TabIndex        =   13
            ToolTipText     =   "Don't display information about the DLL when downloading"
            Top             =   600
            Width           =   2415
         End
         Begin VB.OptionButton opDisplay 
            Caption         =   "Display DLL Information"
            Height          =   255
            Left            =   120
            TabIndex        =   12
            ToolTipText     =   "Display information about the DLL when downloading"
            Top             =   240
            Value           =   -1  'True
            Width           =   2055
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Default Run Location"
         Height          =   2775
         Left            =   -74760
         TabIndex        =   8
         Top             =   1920
         Width           =   4455
         Begin VB.DirListBox dirRun 
            Height          =   2115
            Left            =   120
            TabIndex        =   9
            ToolTipText     =   "Default location to put run file"
            Top             =   240
            Width           =   4215
         End
         Begin VB.Label lblLocation 
            AutoSize        =   -1  'True
            Caption         =   "Location"
            Height          =   195
            Left            =   120
            TabIndex        =   10
            Top             =   2440
            Width           =   615
         End
      End
      Begin VB.Frame frRun 
         Caption         =   "Overwrite During Run"
         Height          =   1335
         Left            =   -72480
         TabIndex        =   4
         Top             =   480
         Width           =   2175
         Begin VB.OptionButton oprPrompt 
            Caption         =   "Prompt"
            Height          =   255
            Left            =   120
            TabIndex        =   7
            Top             =   960
            Value           =   -1  'True
            Width           =   855
         End
         Begin VB.OptionButton oprNever 
            Caption         =   "Never Overwrite"
            Height          =   255
            Left            =   120
            TabIndex        =   6
            Top             =   600
            Width           =   1575
         End
         Begin VB.OptionButton oprAlways 
            Caption         =   "Always Overwrite"
            Height          =   255
            Left            =   120
            TabIndex        =   5
            Top             =   240
            Width           =   1575
         End
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   4200
         Picture         =   "frmOptions.frx":0400
         Top             =   600
         Width           =   480
      End
      Begin VB.Label lblDecompile 
         AutoSize        =   -1  'True
         Caption         =   "?"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   -71880
         TabIndex        =   28
         ToolTipText     =   "What's EXE Encrypting?"
         Top             =   1320
         Width           =   120
      End
      Begin VB.Label lblDLL 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "DLL Location"
         Height          =   195
         Left            =   -74760
         TabIndex        =   27
         Top             =   2520
         Width           =   4380
      End
      Begin VB.Label lblEncrypt 
         AutoSize        =   -1  'True
         Caption         =   "?"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   -71880
         TabIndex        =   25
         ToolTipText     =   "What's EXE Encrypting?"
         Top             =   960
         Width           =   120
      End
      Begin VB.Label lblCopyright 
         Alignment       =   2  'Center
         Caption         =   "Copyright 2001 (C)  D++ IDE"
         Height          =   255
         Left            =   -74880
         TabIndex        =   21
         Top             =   4560
         Width           =   4695
      End
      Begin VB.Image Image4 
         BorderStyle     =   1  'Fixed Single
         Height          =   1785
         Left            =   -74760
         Picture         =   "frmOptions.frx":0842
         Top             =   600
         Width           =   885
      End
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   2760
      TabIndex        =   2
      Top             =   5160
      Width           =   1095
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "Apply"
      Enabled         =   0   'False
      Height          =   375
      Left            =   3960
      TabIndex        =   1
      Top             =   5160
      Width           =   1095
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   1560
      TabIndex        =   0
      Top             =   5160
      Width           =   1095
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim loaded As Boolean

Private Sub chkColorize_Click()
If loaded = False Then Exit Sub
cmdApply.Enabled = True
End Sub

Private Sub chkDecompile_Click()
If loaded = False Then Exit Sub
cmdApply.Enabled = True
End Sub

Private Sub chkEncrypt_Click()
If loaded = False Then Exit Sub
cmdApply.Enabled = True
End Sub

Private Sub cmdApply_Click()
cmdApply.Enabled = False
ApplySettings
End Sub

Private Sub cmdAssociate_Click()
AssociateFiles
MsgBox "All .dpp files have been associated with D++ IDE", vbExclamation, "Associate"
End Sub

Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdOk_Click()
ApplySettings
Unload Me
End Sub

Private Sub cmdView_Click()
If FileExist(GetSystemDirectory & "\DLLINF.txt") Then
    ShowInformation ReadFile(GetSystemDirectory & "\DLLINF.txt"), "DLL Information"
Else
    MsgBox "Unable to locate DLL information file!", vbCritical, "File Not Found"
End If
End Sub

Private Sub dirRun_Change()
lblLocation.Caption = dirRun.Path
End Sub

Private Sub Form_Load()
GetSettings
lblDLL.Caption = "D++ DLL: " & DLLFILE
End Sub

Sub ApplySettings()
On Error Resume Next
If oprAlways.Value = True Then SaveSetting "D++", "Options", "Run", "-1"
If oprNever.Value = True Then SaveSetting "D++", "Options", "Run", "0"
If oprPrompt.Value = True Then SaveSetting "D++", "Options", "Run", "1"

If opcAlways.Value = True Then SaveSetting "D++", "Options", "Compile", "-1"
If opcNever.Value = True Then SaveSetting "D++", "Options", "Compile", "0"
If opcPrompt.Value = True Then SaveSetting "D++", "Options", "Compile", "1"

If opDisplay.Value = True Then SaveSetting "D++", "Options", "Download", "-1"
If opNoDisplay.Value = True Then SaveSetting "D++", "Options", "Download", "0"

If chkEncrypt.Value = 1 Then SaveSetting "D++", "Options", "Encrypt", "DPP:$D2>:"
If chkEncrypt.Value = 0 Then SaveSetting "D++", "Options", "Encrypt", "DPP:"

If chkDebug.Value = 1 Then SaveSetting "D++", "Options", "Debugging", "0"
If chkDebug.Value = 0 Then SaveSetting "D++", "Options", "Debugging", "-1"

If chkDecompile.Value = 1 Then SaveSetting "D++", "Options", "Decompile", "-1"
If chkDecompile.Value = 0 Then SaveSetting "D++", "Options", "Decompile", "0"

If opDownload.Value = True Then SaveSetting "D++", "Options", "DownloadDLL", 0
If opOriginal.Value = True Then SaveSetting "D++", "Options", "DownloadDLL", -1

'If chkColorize.Value = True Then SaveSetting "D++", "Options", "Colorize", "1"
'If chkColorize.Value = False Then SaveSetting "D++", "Options", "Colorize", "0"

SaveSetting "D++", "Options", "RunAt", dirRun.Path
End Sub

Sub GetSettings()
On Error Resume Next
loaded = False
If GetSetting("D++", "Options", "Run") = -1 Then oprAlways.Value = True
If GetSetting("D++", "Options", "Run") = 0 Then oprNever.Value = True
If GetSetting("D++", "Options", "Run") = 1 Then oprPrompt.Value = True

If GetSetting("D++", "Options", "Compile") = -1 Then opcAlways.Value = True
If GetSetting("D++", "Options", "Compile") = 0 Then opcNever.Value = True
If GetSetting("D++", "Options", "Compile") = 1 Then opcPrompt.Value = True

If GetSetting("D++", "Options", "Download") = -1 Then opDisplay.Value = True
If GetSetting("D++", "Options", "Download") = 0 Then opNoDisplay.Value = True

If GetSetting("D++", "Options", "Encrypt") = "DPP:$D2>:" Then chkEncrypt.Value = 1
If GetSetting("D++", "Options", "Encrypt") = "DPP:" Then chkEncrypt.Value = 0

If GetSetting("D++", "Options", "Debugging") = "0" Then chkDebug.Value = 1
If GetSetting("D++", "Options", "Debugging") = "-1" Then chkDebug.Value = 0

If GetSetting("D++", "Options", "Decompile") = "-1" Then chkDecompile.Value = 1
If GetSetting("D++", "Options", "Decompile") = "0" Then chkDecompile.Value = 0

If GetSetting("D++", "Options", "DownloadDLL") = 0 Then opDownload.Value = True
If GetSetting("D++", "Options", "DownloadDLL") = -1 Then opOriginal.Value = True

'If GetSetting("D++", "Options", "Colorize") = 1 Then chkColorize.Value = True
'If GetSetting("D++", "Options", "Colorize") = 0 Then chkColorize.Value = False

dirRun.Path = GetSetting("D++", "Options", "RunAt")
lblLocation.Caption = dirRun.Path
loaded = True
End Sub

Private Sub lblCopyright_Click()
frmAbout.Show 1
End Sub

Private Sub lblDecompile_Click()
ShowInformation "Prevent Decompile is a security option.  No one will be able to decompile the EXE's you make, so your source code is protected.  EXE encrypting will help make it more secure.", "Prevent Decompile"
End Sub

Private Sub lblEncrypt_Click()
ShowInformation "EXE Encryptng is a security option.  Your D++ source is encrypted within the EXE, preventing external viewers, such as Notepad, from viewing your source", "EXE Encrypting"
End Sub

Private Sub opcAlways_Click()
If loaded = False Then Exit Sub
cmdApply.Enabled = True
End Sub

Private Sub opcNever_Click()
If loaded = False Then Exit Sub
cmdApply.Enabled = True
End Sub

Private Sub opcPrompt_Click()
If loaded = False Then Exit Sub
cmdApply.Enabled = True
End Sub

Private Sub opDisplay_Click()
If loaded = False Then Exit Sub
cmdApply.Enabled = True
End Sub

Private Sub opNoDisplay_Click()
If loaded = False Then Exit Sub
Dim Message As String
Message = "When downloading the DLL from the internet, you are modifing the D++ language.  Sometimes the language syntax is changed.  "
Message = Message & "All the information is about the DLL is stored in a file in your sytem folder called 'DLLINF.txt'.  "
Message = Message & "After the download, this information normaly is displayed.  If you don't want to read the information, I recomend you read it from the file.  You can do this under 'IDE' in the options."
ShowInformation Message, "D++ Downloading"
cmdApply.Enabled = True
End Sub

Private Sub oprAlways_Click()
If loaded = False Then Exit Sub
cmdApply.Enabled = True
End Sub

Private Sub oprNever_Click()
If loaded = False Then Exit Sub
cmdApply.Enabled = True
End Sub

Private Sub oprPrompt_Click()
If loaded = False Then Exit Sub
cmdApply.Enabled = True
End Sub

