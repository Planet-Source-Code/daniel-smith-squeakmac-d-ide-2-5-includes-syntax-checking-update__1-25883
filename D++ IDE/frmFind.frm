VERSION 5.00
Begin VB.Form frmFind 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Find"
   ClientHeight    =   1575
   ClientLeft      =   4830
   ClientTop       =   5340
   ClientWidth     =   5295
   Icon            =   "frmFind.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1575
   ScaleWidth      =   5295
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdGo 
      Caption         =   "Goto"
      Height          =   375
      Left            =   3960
      TabIndex        =   5
      Top             =   1080
      Width           =   1215
   End
   Begin VB.CommandButton cmdReplace 
      Caption         =   "Replace"
      Height          =   375
      Left            =   3960
      TabIndex        =   3
      Top             =   600
      Width           =   1215
   End
   Begin VB.CommandButton cmdFind 
      Caption         =   "Find"
      Default         =   -1  'True
      Height          =   375
      Left            =   3960
      TabIndex        =   1
      Top             =   120
      Width           =   1215
   End
   Begin VB.TextBox txtGo 
      Height          =   285
      Left            =   1320
      TabIndex        =   4
      Top             =   1080
      Width           =   1815
   End
   Begin VB.TextBox txtReplace 
      Height          =   285
      Left            =   1320
      TabIndex        =   2
      Top             =   600
      Width           =   2415
   End
   Begin VB.TextBox txtFind 
      Height          =   285
      Left            =   1320
      TabIndex        =   0
      Top             =   120
      Width           =   2415
   End
   Begin VB.Label lblLength 
      AutoSize        =   -1  'True
      Caption         =   "of 0."
      Height          =   195
      Left            =   3240
      TabIndex        =   9
      Top             =   1125
      Width           =   315
   End
   Begin VB.Label lblGo 
      AutoSize        =   -1  'True
      Caption         =   "Goto Charcter:"
      Height          =   195
      Left            =   120
      TabIndex        =   8
      Top             =   1120
      Width           =   1035
   End
   Begin VB.Label lblReplace 
      AutoSize        =   -1  'True
      Caption         =   "Replace With:"
      Height          =   195
      Left            =   120
      TabIndex        =   7
      Top             =   630
      Width           =   1020
   End
   Begin VB.Label lblFind 
      AutoSize        =   -1  'True
      Caption         =   "Find What:"
      Height          =   195
      Left            =   120
      TabIndex        =   6
      Top             =   170
      Width           =   780
   End
End
Attribute VB_Name = "frmFind"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub cmdFind_Click()
frmMain.FindIt txtFind.Text
Unload Me
End Sub

Private Sub cmdGo_Click()
On Error Resume Next
frmMain.SetFocus
If Val(txtGo.Text) > Len(frmMain.txtText.Text) Then
    MsgBox "Invalid character", vbExclamation, "Goto Character"
    Unload Me
Else
    frmMain.txtText.SelStart = Val(txtGo.Text)
End If
Unload Me
End Sub

Private Sub cmdReplace_Click()
frmMain.txtText.Text = FindReplace(frmMain.txtText.Text, txtFind.Text, txtReplace.Text)
Unload Me
End Sub

Private Sub Form_Load()
txtFind.Text = ""
txtReplace.Text = ""
txtGo.Text = ""
lblLength.Caption = "of " & Len(frmMain.txtText.Text) & "."
End Sub

Private Function FindReplace(sText As String, sFind As String, sReplace As String) As String
    Dim n%, c%
    Dim sTempR$, sTempL$
    c = 1
    n = 1
    Do
        c = InStr(n, sText, sFind)
        If c% <> 0 Then
            sTempL = Mid$(sText, 1, c - 1)
            sTempR = Mid$(sText, c + Len(sFind))
            sText = sTempL & sReplace & sTempR
        End If
        n = c + 1
    Loop Until c = 0
    FindReplace = sText
End Function

