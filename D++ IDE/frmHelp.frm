VERSION 5.00
Begin VB.Form frmHelp 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "D++ Help"
   ClientHeight    =   6315
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7380
   Icon            =   "frmHelp.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6315
   ScaleWidth      =   7380
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      Align           =   3  'Align Left
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      Height          =   6315
      Left            =   0
      ScaleHeight     =   6315
      ScaleWidth      =   840
      TabIndex        =   4
      Top             =   0
      Width           =   840
      Begin VB.Image picDPP 
         Height          =   480
         Left            =   230
         Picture         =   "frmHelp.frx":030A
         Top             =   120
         Width           =   480
      End
      Begin VB.Label lblArea 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "D++ Help"
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
         TabIndex        =   5
         Top             =   600
         Width           =   825
         WordWrap        =   -1  'True
      End
   End
   Begin VB.TextBox txtHelp 
      Height          =   5655
      Left            =   960
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   3
      Text            =   "frmHelp.frx":0614
      Top             =   0
      Width           =   6375
   End
   Begin VB.Frame Frame1 
      Caption         =   "Search"
      Height          =   660
      Left            =   960
      TabIndex        =   0
      Top             =   5640
      Width           =   6375
      Begin VB.CommandButton cmdSearch 
         Caption         =   "Search"
         Default         =   -1  'True
         Height          =   315
         Left            =   5040
         TabIndex        =   2
         Top             =   240
         Width           =   1215
      End
      Begin VB.TextBox txtSearch 
         Height          =   315
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   4815
      End
   End
End
Attribute VB_Name = "frmHelp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdSearch_Click()
search
End Sub

Public Sub search()
cmdSearch.Enabled = False
txtHelp.SetFocus
DoEvents
FindIt txtSearch.Text
End Sub

Public Sub FindIt(FindWhat As String)
If txtHelp.SelLength > 0 Then
    txtHelp.SelStart = txtHelp.SelStart + txtHelp.SelLength
End If
If txtHelp.SelStart = 0 Then txtHelp.SelStart = 1
i = InStr(txtHelp.SelStart, txtHelp.Text, FindWhat)
If i <> 0 Then
    txtHelp.SelStart = i - 1
    txtHelp.SelLength = Len(FindWhat)
Else
    MsgBox "Search text not found.", vbExclamation, "Search"
End If
cmdSearch.Enabled = True
End Sub

