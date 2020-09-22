VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.Form frmDownload 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Download Latest D++ DLL"
   ClientHeight    =   5040
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5055
   ControlBox      =   0   'False
   Icon            =   "frmDownload.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5040
   ScaleWidth      =   5055
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel Download"
      Enabled         =   0   'False
      Height          =   375
      Left            =   1800
      TabIndex        =   18
      Top             =   4560
      Width           =   1455
   End
   Begin VB.Frame Frame3 
      Caption         =   "Total File Status"
      Height          =   615
      Left            =   120
      TabIndex        =   15
      Top             =   3840
      Width           =   4815
      Begin VB.Label lblRemaining 
         Caption         =   "2 File(s) Remaining"
         Height          =   255
         Left            =   3240
         TabIndex        =   17
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label lblTotal 
         Caption         =   "0 File(s) Downloaded"
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   240
         Width           =   1575
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "File Download Information"
      Height          =   2295
      Left            =   120
      TabIndex        =   5
      Top             =   1440
      Width           =   4815
      Begin MSComctlLib.ProgressBar ProgressBar 
         Height          =   225
         Left            =   120
         TabIndex        =   6
         Top             =   810
         Width           =   4515
         _ExtentX        =   7964
         _ExtentY        =   397
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.Label StatusLabel 
         Caption         =   "StatusLabel"
         Height          =   195
         Left            =   120
         TabIndex        =   14
         Top             =   240
         Width           =   4455
      End
      Begin VB.Label EstimatedTimeLeft 
         Caption         =   "Estimated time left:"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   1125
         Width           =   1455
      End
      Begin VB.Label SourceLabel 
         Caption         =   "SourceLabel"
         Height          =   195
         Left            =   120
         TabIndex        =   12
         Top             =   525
         Width           =   4530
      End
      Begin VB.Label TimeLabel 
         Height          =   255
         Left            =   1560
         TabIndex        =   11
         Top             =   1125
         Width           =   3045
      End
      Begin VB.Label DownloadTo 
         Caption         =   "Download to:"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   1440
         Width           =   1335
      End
      Begin VB.Label ToLabel 
         Height          =   195
         Left            =   1560
         TabIndex        =   9
         Top             =   1440
         Width           =   3075
      End
      Begin VB.Label TransferRate 
         Caption         =   "Transfer rate:"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   1800
         Width           =   1335
      End
      Begin VB.Label RateLabel 
         Height          =   255
         Left            =   1560
         TabIndex        =   7
         Top             =   1800
         Width           =   3045
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Download Information"
      Height          =   1215
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   4815
      Begin VB.Label Label1 
         Caption         =   $"frmDownload.frx":030A
         Height          =   855
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   4575
      End
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      Height          =   375
      Left            =   3480
      TabIndex        =   2
      Top             =   4560
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3480
      TabIndex        =   1
      Top             =   4560
      Width           =   1455
   End
   Begin VB.CommandButton cmdDownload 
      Caption         =   "Download"
      Default         =   -1  'True
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   4560
      Width           =   1455
   End
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   4320
      Top             =   1080
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      AccessType      =   1
      Protocol        =   2
      RemotePort      =   21
      URL             =   "ftp://"
   End
End
Attribute VB_Name = "frmDownload"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private CancelSearch As Boolean

Public Function DownloadFile(strURL As String, _
                             strDestination As String, _
                             Optional UserName As String = Empty, _
                             Optional Password As String = Empty) _
                             As Boolean

'THIS FUNCTION IS NOT MINE
' Funtion DownloadFile: Download a file via HTTP
'
' Inputs:   strURL String; the source URL of the file
'           strDestination; valid Win95/NT path to where you want it
'           (i.e. "C:\Program Files\My Stuff\Purina.pdf")
'
' Returns:  Boolean; Was the download successful?

Const CHUNK_SIZE As Long = 1024 ' Download chunk size
Const ROLLBACK As Long = 4096   ' Bytes to roll back on resume
                                ' You can be less conservative,
                                ' and roll back less, but I
                                ' don't recommend it.
Dim bData() As Byte             ' Data var
Dim blnResume As Boolean        ' True if resuming download
Dim intFile As Integer          ' FreeFile var
Dim lngBytesReceived As Long    ' Bytes received so far
Dim lngFileLength As Long       ' Total length of file in bytes
Dim lngX                        ' Temp long var
Dim sglLastTime As Single          ' Time last chunk received
Dim sglRate As Single           ' Var to hold transfer rate
Dim sglTime As Single           ' Var to hold time remaining
Dim strFile As String           ' Temp filename var
Dim strHeader As String         ' HTTP header store
Dim strHost As String           ' HTTP Host

On Local Error GoTo InternetErrorHandler

' Start with Cancel flag = False
CancelSearch = False

' Get just filename (without dirs) for display
strFile = ReturnFileOrFolder(strDestination, True)
strHost = ReturnFileOrFolder(strURL, True, True)
              
SourceLabel = Empty
TimeLabel = Empty
ToLabel = Empty
RateLabel = Empty


StartDownload:

If blnResume Then
    StatusLabel = "Resuming download..."
    lngBytesReceived = lngBytesReceived - ROLLBACK
    If lngBytesReceived < 0 Then lngBytesReceived = 0
Else
    StatusLabel = "Getting file information..."
End If
' Give the system time to update the form gracefully
DoEvents

' Download file
With Inet1
    .URL = strURL
    .UserName = UserName
    .Password = Password
    ' GET file, sending the magic resume input header...
    .Execute , "GET", , "Range: bytes=" & CStr(lngBytesReceived) & "-" & vbCrLf
    
    ' While initiating connection, yield CPU to Windows
    While .StillExecuting
        DoEvents
        ' If user pressed Cancel button on StatusForm
        ' then fail, cancel, and exit this download
        If CancelSearch Then GoTo ExitDownload
    Wend

    StatusLabel = "Saving:"
    SourceLabel = FitText(SourceLabel, strHost & " from " & .RemoteHost)
    ToLabel = FitText(ToLabel, strDestination)

    ' Get first header ("HTTP/X.X XXX ...")
    strHeader = .GetHeader
End With

' Trap common HTTP response codes
Select Case Mid(strHeader, 10, 3)
    Case "200"  ' OK
        ' If resuming, however, this is a failure
        If blnResume Then
            ' Delete partially downloaded file
            Kill strDestination
            ' Prompt
            If MsgBox("The server is unable to resume this download." & _
                      vbCr & vbCr & _
                      "Do you want to continue anyway?", _
                      vbExclamation + vbYesNo, _
                      "Unable to Resume Download") = vbYes Then
                    ' Yes - continue anyway:
                    ' Set resume flag to False
                    blnResume = False
                Else
                    ' No - cancel
                    CancelSearch = True
                    GoTo ExitDownload
                End If
            End If
            
    Case "206"  ' 206=Partial Content, which is GREAT when resuming!
    
    Case "204"  ' No content
        MsgBox "Nothing to download!", _
               vbInformation, _
               "No Content"
        CancelSearch = True
        GoTo ExitDownload
        
    Case "401"  ' Not authorized
        MsgBox "Authorization failed!", _
               vbCritical, _
               "Unauthorized"
        CancelSearch = True
        GoTo ExitDownload
    
    Case "404"  ' File Not Found
        MsgBox "The file, " & _
               """" & Inet1.URL & """" & _
               " was not found!", _
               vbCritical, _
               "File Not Found"
        CancelSearch = True
        GoTo ExitDownload
        
    Case vbCrLf ' Empty header
        MsgBox "Cannot establish connection." & vbCr & vbCr & _
               "Check your Internet connection and try again.", _
               vbExclamation, _
               "Cannot Establish Connection"
        CancelSearch = True
        GoTo ExitDownload
        
    Case Else
        ' Miscellaneous unexpected errors
        strHeader = Left(strHeader, InStr(strHeader, vbCr))
        If strHeader = Empty Then strHeader = "<nothing>"
        MsgBox "The server returned the following response:" & _
               vbCr & vbCr & _
               strHeader, _
               vbCritical, _
               "Error Downloading File"
        CancelSearch = True
        GoTo ExitDownload
End Select

' Get file length with "Content-Length" header request
If blnResume = False Then
    ' Set timer for gauging download speed
    sglLastTime = Timer - 1
    strHeader = Inet1.GetHeader("Content-Length")
    lngFileLength = Val(strHeader)
    If lngFileLength = 0 Then
        GoTo ExitDownload
    End If
End If

' Check for available disk space first...
' If on a physical or mapped drive. Can't with a UNC path.
If Mid(strDestination, 2, 2) = ":\" Then
    If DiskFreeSpace(Left(strDestination, _
                          InStr(strDestination, "\"))) < lngFileLength Then
        ' Not enough free space to download file
        MsgBox "There is not enough free space on disk for this file." _
               & vbCr & vbCr & "Please free up some disk space and try again.", _
               vbCritical, _
               "Insufficient Disk Space"
        GoTo ExitDownload
    End If
End If

' Prepare display
'
' Progress Bar
With ProgressBar
    .Value = 0
    .Max = lngFileLength
End With

' Give system a chance to show AVI
DoEvents

' Reset bytes received counter if not resuming
If blnResume = False Then lngBytesReceived = 0


On Local Error GoTo FileErrorHandler

' Create destination directory, if necessary
strHeader = ReturnFileOrFolder(strDestination, False)
If Dir(strHeader, vbDirectory) = Empty Then
    MkDir strHeader
End If

' If no errors occurred, then spank the file to disk
intFile = FreeFile        ' Set intFile to an unused file.
' Open a file to write to.
Open strDestination For Binary Access Write As #intFile
' If resuming, then seek byte position in downloaded file
' where we last left off...
If blnResume Then Seek #intFile, lngBytesReceived + 1
Do
    ' Get chunks...
    bData = Inet1.GetChunk(CHUNK_SIZE, icByteArray)
    Put #intFile, , bData   ' Put it into our destination file
    If CancelSearch Then Exit Do
    lngBytesReceived = lngBytesReceived + UBound(bData, 1) + 1
    sglRate = lngBytesReceived / (Timer - sglLastTime)
    sglTime = (lngFileLength - lngBytesReceived) / sglRate
    TimeLabel = FormatTime(sglTime) & _
                   " (" & _
                   FormatFileSize(lngBytesReceived) & _
                   " of " & _
                   FormatFileSize(lngFileLength) & _
                   " copied)"
    RateLabel = FormatFileSize(sglRate, "###.0") & "/Sec"
    ProgressBar.Value = lngBytesReceived
    Me.Caption = Format((lngBytesReceived / lngFileLength), "##0%") & _
                 " of " & strFile & " Completed"
Loop While UBound(bData, 1) > 0       ' Loop while there's still data...
Close #intFile

ExitDownload:
' Success if the # of bytes transferred = content length
If lngBytesReceived = lngFileLength Then
    StatusLabel = "Download completed!"
    DownloadFile = True
Else
    If Dir(strDestination) = Empty Then
        CancelSearch = True
    Else
        ' Resume? (If not cancelled)
        If CancelSearch = False Then
            If MsgBox("The connection with the server was reset." & _
                      vbCr & vbCr & _
                      "Click ""Retry"" to resume downloading the file." & _
                      vbCr & "(Approximate time remaining: " & FormatTime(sglTime) & ")" & _
                      vbCr & vbCr & _
                      "Click ""Cancel"" to cancel downloading the file.", _
                      vbExclamation + vbRetryCancel, _
                      "Download Incomplete") = vbRetry Then
                    ' Yes
                    blnResume = True
                    GoTo StartDownload
            End If
        End If
    End If
    ' No or unresumable failure:
    ' Delete partially downloaded file
    If Not Dir(strDestination) = Empty Then Kill strDestination
    DownloadFile = False
End If

CleanUp:

' Make sure that the Internet connection is closed...
Inet1.Cancel
' ...and exit this function

Exit Function

InternetErrorHandler:
    ' Err# 9 occurs when UBound(bData,1) < 0
    If Err.Number = 9 Then Resume Next
    ' Other errors...
    MsgBox "Error: " & Err.Description & " occurred.", _
           vbCritical, _
           "Error Downloading File"
    Err.Clear
    GoTo ExitDownload
    
FileErrorHandler:
    MsgBox "Cannot write file to disk." & _
           vbCr & vbCr & _
           "Error " & Err.Number & ": " & Err.Description, _
           vbCritical, _
           "Error Downloading File"
    CancelSearch = True
    Err.Clear
    GoTo ExitDownload
    
End Function

Private Sub cmdCancel_Click()
StatusLabel = "Cancelling..."
CancelSearch = True
End Sub

Private Sub cmdDownload_Click()
On Error Resume Next
Dim OldVersion As Integer

cmdQuit.Visible = False
cmdCancel.Enabled = True
cmdDownload.Enabled = False

StatusLabel.Caption = "Backing up DLL..."
FileCopy GetSystemDirectory & "\DPPAPP.dll", GetSystemDirectory & "\DPPAPP2.dll"
OldVersion = DLLVersion

DownloadFile "http://server39.hypermart.net/squeakmac/DPPAPP.dll", GetSystemDirectory & "\DPPAPP.dll"

lblRemaining.Caption = "1 File(s) Remaining"
lblTotal.Caption = "1 File(s) Downloaded"
Kill GetSystemDirectory & "\DLLINF.txt"

DownloadFile "http://server39.hypermart.net/squeakmac/DLLINF.txt", GetSystemDirectory & "\DLLINF.txt"

lblRemaining.Caption = "0 File(s) Remaining"
lblTotal.Caption = "2 File(s) Downloaded"

StatusLabel.Caption = "Confirming new DLL..."
If OldVersion > GetDLLVersion Then
    If GetSetting("D++", "Options", "DownloadDLL") = 0 Then GoTo finish
    Kill GetSystemDirectory & "\DPPAPP.dll"
    FileCopy GetSystemDirectory & "\DPPAPP2.dll", GetSystemDirectory & "\DPPAPP.dll"
    
    MsgBox "The DLL that was downloaded was not compatible with this compiler, or was older then the pervious DLL." & vbCrLf & vbCrLf & "The original, best compatible DLL was restored.", vbExclamation, "DLL Version"
    Unload Me
    Exit Sub
End If

finish:
cmdCancel.Enabled = False
Kill GetSystemDirectory & "\DPPAPP2.dll"

ShowInfo = GetSetting("D++", "Options", "Download")
Select Case ShowInfo
    Case -1
        ShowInformation ReadFile(GetSystemDirectory & "\DLLINF.txt"), "DLL Information"
    Case 0
        Unload Me
        Exit Sub
End Select

Unload Me
End Sub

Private Sub cmdClose_Click()
On Error Resume Next
inetDownload.Cancel
Unload Me
End Sub

Private Sub cmdQuit_Click()
On Error Resume Next
inetDownload.Cancel
End
End Sub

Private Sub Form_Load()
cmdClose.Caption = "Cancel"
cmdDownload.Enabled = True
End Sub
