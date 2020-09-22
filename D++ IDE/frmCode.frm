VERSION 5.00
Begin VB.Form frmCode 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "D++ Source Code"
   ClientHeight    =   6330
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7380
   Icon            =   "frmCode.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6330
   ScaleWidth      =   7380
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtCode 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5655
      Left            =   960
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   12
      Top             =   0
      Width           =   6375
   End
   Begin VB.TextBox txt6 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   10320
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   11
      Text            =   "frmCode.frx":030A
      Top             =   1200
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.TextBox txt5 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   10320
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   10
      Text            =   "frmCode.frx":08DF
      Top             =   120
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.TextBox txt4 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   7440
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   9
      Text            =   "frmCode.frx":0A61
      Top             =   4680
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.TextBox txt3 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   7440
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   8
      Text            =   "frmCode.frx":0D40
      Top             =   3480
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.TextBox txt2 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   7440
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   7
      Text            =   "frmCode.frx":0E6B
      Top             =   2400
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.TextBox txt1 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   7440
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   6
      Text            =   "frmCode.frx":107B
      Top             =   1200
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.TextBox txt0 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   7440
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   5
      Text            =   "frmCode.frx":174D
      Top             =   120
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.Frame Frame1 
      Caption         =   "Code"
      Height          =   660
      Left            =   960
      TabIndex        =   2
      Top             =   5640
      Width           =   6375
      Begin VB.CommandButton cmdInsert 
         Caption         =   "Insert"
         Default         =   -1  'True
         Height          =   315
         Left            =   5040
         TabIndex        =   4
         Top             =   240
         Width           =   1215
      End
      Begin VB.ComboBox Code 
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   240
         Width           =   4815
      End
   End
   Begin VB.PictureBox Picture1 
      Align           =   3  'Align Left
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      Height          =   6330
      Left            =   0
      ScaleHeight     =   6330
      ScaleWidth      =   840
      TabIndex        =   0
      Top             =   0
      Width           =   840
      Begin VB.Label lblArea 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "D++ Source"
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
      Begin VB.Image picDPP 
         Height          =   480
         Left            =   230
         Picture         =   "frmCode.frx":17D5
         Top             =   120
         Width           =   480
      End
   End
End
Attribute VB_Name = "frmCode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdInsert_Click()
frmMain.txtText.Text = txtCode.Text
Unload Me
End Sub

Private Sub Code_Click()
DisplayCode Code.ListIndex
End Sub

Private Sub Form_Load()
Code.AddItem "Simple Name Input"
Code.AddItem "D++ Intro Program"
Code.AddItem "Using Loops"
Code.AddItem "Using If's"
Code.AddItem "Binary Tree Example"
Code.AddItem "Nested Loops Example"
Code.AddItem "Guess Number"
DisplayCode 1
Code.Text = "D++ Intro Program"
End Sub

Private Sub DisplayCode(Index As Integer)
Select Case Index
    Case 0
        txtCode.Text = txt0.Text
        Me.Caption = "D++ Source Code - [Simple Input Example]"
    Case 1
        txtCode.Text = txt1.Text
        Me.Caption = "D++ Source Code - [D++ Introduction Program]"
    Case 2
        txtCode.Text = txt2.Text
        Me.Caption = "D++ Source Code - [Using Loops]"
    Case 3
        txtCode.Text = txt3.Text
        Me.Caption = "D++ Source Code - [Using If's]"
    Case 4
        txtCode.Text = txt4.Text
        Me.Caption = "D++ Source Code - [Binary Tree Example]"
    Case 5
        txtCode.Text = txt5.Text
        Me.Caption = "D++ Source Code - [Nested Loops Example]"
    Case 6
        txtCode.Text = txt6.Text
        Me.Caption = "D++ Source Code - [Guess Number Example]"
End Select
End Sub

