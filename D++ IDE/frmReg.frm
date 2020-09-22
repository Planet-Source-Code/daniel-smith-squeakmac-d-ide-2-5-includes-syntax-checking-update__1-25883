VERSION 5.00
Begin VB.Form frmReg 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Register D++ IDE"
   ClientHeight    =   3855
   ClientLeft      =   6075
   ClientTop       =   4740
   ClientWidth     =   3270
   ControlBox      =   0   'False
   Icon            =   "frmReg.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3855
   ScaleWidth      =   3270
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      Height          =   375
      Left            =   2040
      TabIndex        =   6
      Top             =   3360
      Width           =   1095
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "Ok"
      Default         =   -1  'True
      Height          =   375
      Left            =   840
      TabIndex        =   5
      Top             =   3360
      Width           =   1095
   End
   Begin VB.TextBox txtCompany 
      Height          =   285
      Left            =   120
      TabIndex        =   4
      Top             =   3000
      Width           =   3015
   End
   Begin VB.TextBox txtName 
      Height          =   285
      Left            =   120
      TabIndex        =   2
      Top             =   2400
      Width           =   3015
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Company:"
      Height          =   195
      Left            =   120
      TabIndex        =   3
      Top             =   2760
      Width           =   705
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Name:"
      Height          =   195
      Left            =   120
      TabIndex        =   1
      Top             =   2160
      Width           =   465
   End
   Begin VB.Label Label1 
      Caption         =   "Please register your copy of D++ IDE.  This information is kept on this computer."
      Height          =   435
      Left            =   120
      TabIndex        =   0
      Top             =   1560
      Width           =   3015
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   1395
      Left            =   120
      Picture         =   "frmReg.frx":030A
      Top             =   120
      Width           =   3060
   End
End
Attribute VB_Name = "frmReg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdOk_Click()
If txtName.Text = "" Or txtCompany.Text = "" Then
    MsgBox "You must register to use D++ IDE!", vbExclamation, "Register"
    Exit Sub
End If
CreateSettings
Unload Me
frmMain.SetFocus
frmMain.txtText.Text = frmCode.txt1.Text
frmMain.Caption = "D++ IDE - [Welcome to D2]"
End Sub

Private Sub cmdQuit_Click()
quitval = MsgBox("Are you sure you want quit without registering?", vbYesNo, "Quit with registering?")
If quitval = vbNo Then
    Exit Sub
End If
End
End Sub

Private Sub CreateSettings()
'Registration
SaveSetting "D++", "Reg", "UserName", txtName.Text
SaveSetting "D++", "Reg", "Company", txtCompany.Text
'Default Option Settings
SaveSetting "D++", "Options", "Run", "1"
SaveSetting "D++", "Options", "Compile", "1"
SaveSetting "D++", "Options", "Download", "-1"
SaveSetting "D++", "Options", "Debugging", "0" 'debugging options reversed...(0 = true)
SaveSetting "D++", "Options", "RunAt", GetDesktopDirectory(frmReg.hWnd)
SaveSetting "D++", "Options", "Encrypt", "DPP:$D2>:"
SaveSetting "D++", "Options", "Decompile", "0"
SaveSetting "D++", "Options", "DownloadDLL", "-1"

Result = MsgBox("Would you like to associate .dpp files with D++ IDE?", vbInformation + vbYesNo, "Associate Files")
If Result = vbYes Then
    AssociateFiles
End If
End Sub

