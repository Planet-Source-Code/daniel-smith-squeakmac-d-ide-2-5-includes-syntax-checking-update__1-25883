VERSION 5.00
Begin VB.Form frmInformation 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Information"
   ClientHeight    =   5745
   ClientLeft      =   4665
   ClientTop       =   3675
   ClientWidth     =   7095
   Icon            =   "frmInformation.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5745
   ScaleWidth      =   7095
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtText 
      BackColor       =   &H00C0C0C0&
      Height          =   5055
      Left            =   960
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Top             =   120
      Width           =   6015
   End
   Begin VB.PictureBox Picture1 
      Align           =   3  'Align Left
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      Height          =   5745
      Left            =   0
      ScaleHeight     =   5745
      ScaleWidth      =   840
      TabIndex        =   1
      Top             =   0
      Width           =   840
      Begin VB.Image picDPP 
         Height          =   480
         Left            =   230
         Picture         =   "frmInformation.frx":030A
         Top             =   120
         Width           =   480
      End
      Begin VB.Label lblArea 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "D++ Info"
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
         TabIndex        =   2
         Top             =   600
         Width           =   825
         WordWrap        =   -1  'True
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "Ok"
      Default         =   -1  'True
      Height          =   375
      Left            =   3480
      TabIndex        =   0
      Top             =   5280
      Width           =   1215
   End
End
Attribute VB_Name = "frmInformation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdOk_Click()
Unload Me
End Sub

