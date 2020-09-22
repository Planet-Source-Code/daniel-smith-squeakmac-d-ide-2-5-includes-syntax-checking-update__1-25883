VERSION 5.00
Begin VB.Form frmIcon 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Choose Icon..."
   ClientHeight    =   1800
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3750
   Icon            =   "frmIcon.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1800
   ScaleWidth      =   3750
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "Command3"
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   1320
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command1"
      Height          =   375
      Left            =   1320
      TabIndex        =   3
      Top             =   1320
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   2520
      TabIndex        =   2
      Top             =   1320
      Width           =   1095
   End
   Begin VB.Label lblPah 
      Caption         =   "(Default Icon)"
      Height          =   435
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   3480
   End
   Begin VB.Label lblIcon 
      Caption         =   "Select an Icon you wish to use for your program:"
      Height          =   495
      Left            =   720
      TabIndex        =   0
      Top             =   120
      Width           =   2895
   End
   Begin VB.Image imgIcon 
      BorderStyle     =   1  'Fixed Single
      Height          =   495
      Left            =   120
      Picture         =   "frmIcon.frx":030A
      Top             =   120
      Width           =   495
   End
End
Attribute VB_Name = "frmIcon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
