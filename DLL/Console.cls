VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Console"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Console declarations
Private Declare Function AllocConsole Lib "kernel32" () As Long
Private Declare Function GetStdHandle Lib "kernel32" (ByVal nStdHandle As Long) As Long
Private Declare Function WriteConsole Lib "kernel32" Alias "WriteConsoleA" (ByVal hConsoleOutput As Long, ByVal lpBuffer As Any, ByVal nNumberOfCharsToWrite As Long, lpNumberOfCharsWritten As Long, lpReserved As Any) As Long
Private Declare Function SetConsoleMode Lib "kernel32" (ByVal hConsoleOutput As Long, dwMode As Long) As Long
Private Declare Function ReadConsole Lib "kernel32" Alias "ReadConsoleA" (ByVal hConsoleInput As Long, ByVal lpBuffer As String, ByVal nNumberOfCharsToRead As Long, lpNumberOfCharsRead As Long, lpReserved As Any) As Long
Private Declare Function SetConsoleTitle Lib "kernel32" Alias "SetConsoleTitleA" (ByVal lpConsoleTitle As String) As Long
Private Declare Function FreeConsole Lib "kernel32" () As Long
Private Const STD_INPUT_HANDLE = -10&
Private Const STD_OUTPUT_HANDLE = -11&
Private Const STD_ERROR_HANDLE = -12&
Private Const ENABLE_LINE_INPUT = &H2
Private Const ENABLE_ECHO_INPUT = &H4
Private Const ENABLE_MOUSE_INPUT = &H10
Private Const ENABLE_PROCESSED_INPUT = &H1
Private Const ENABLE_WINDOW_INPUT = &H8
Private Const ENABLE_PROCESSED_OUTPUT = &H1
Private Const ENABLE_WRAP_AT_EOL_OUTPUT = &H2
Private lHInput As Long, lHOutPut As Long, lHErr As Long

Private Sub Class_Initialize()
    'AddDebug ">Initializing DOS..."
    AllocConsole
    lHInput = GetStdHandle(STD_INPUT_HANDLE)
    lHOutPut = GetStdHandle(STD_OUTPUT_HANDLE)
    lHErr = GetStdHandle(STD_ERROR_HANDLE)
    SetConsoleMode lHInput, ENABLE_ECHO_INPUT
End Sub

Public Sub WriteToCon(WhatToWrite As String)
    WriteConsole lHOutPut, WhatToWrite, Len(WhatToWrite), vbNull, vbNull
End Sub

Public Function ReadFromCon() As String
    On Error Resume Next
    Dim sUserInput As String * 256
    ReadConsole lHInput, sUserInput, Len(sUserInput), vbNull, vbNull
    ReadFromCon = Left$(sUserInput, InStr(sUserInput, Chr$(0)) - 3)
End Function

Public Sub Init()
Class_Initialize
End Sub

Public Sub SetTitle(Title)
    SetConsoleTitle Title
End Sub

Public Sub Terminate()
Class_Terminate
End Sub

Private Sub Class_Terminate()
    'AddDebug ">Teminating DOS..."
    FreeConsole
    'AddDebug ">D++ Application Finished"
End Sub

