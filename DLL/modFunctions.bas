Attribute VB_Name = "modFunctions"
'D++ Function Module
'Contains all D++ API functions

Private Declare Function GetSysDirectory Lib "kernel32" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Private Declare Function IcmpCreateFile Lib "icmp.dll" () As Long
Private Declare Function IcmpCloseHandle Lib "icmp.dll" (ByVal IcmpHandle As Long) As Long
Private Declare Function IcmpSendEcho Lib "icmp.dll" (ByVal IcmpHandle As Long, ByVal DestinationAddress As Long, ByVal RequestData As String, ByVal RequestSize As Long, ByVal RequestOptions As Long, ReplyBuffer As ICMP_ECHO_REPLY, ByVal ReplySize As Long, ByVal TimeOut As Long) As Long
Private Declare Function WSAGetLastError Lib "wsock32.dll" () As Long
Private Declare Function WSAStartup Lib "wsock32.dll" (ByVal wVersionRequired As Long, lpWSADATA As WSADATA) As Long
Private Declare Function WSACleanup Lib "wsock32.dll" () As Long
Private Declare Function gethostname Lib "wsock32.dll" (ByVal szHost As String, ByVal dwHostLen As Long) As Long
Private Declare Function gethostbyname Lib "wsock32.dll" (ByVal szHost As String) As Long
Private Declare Function gethostbyaddr Lib "wsock32.dll" (ByVal szHost As String, ByVal hLen As Integer, ByVal aType As Integer) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (hpvDest As Any, ByVal hpvSource As Long, ByVal cbCopy As Long)
Private Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long
Private Declare Function SystemParametersInfo Lib "User32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, lpvParam As Any, ByVal fuWinIni As Long) As Long

Private Declare Function LockWindowUpdate Lib "User32" (ByVal hWndLock As Long) As Long

'FILE API
Public Declare Function GetTickCount Lib "kernel32" () As Long
Public Declare Function CreateFile Lib "kernel32" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, ByVal lpSecurityAttributes As Long, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Public Declare Function ReadFileAPI Lib "kernel32" Alias "ReadFile" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToRead As Long, lpNumberOfBytesRead As Long, ByVal lpOverlapped As Long) As Long
Public Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Public Const GENERIC_READ = &H80000000
Public Const OPEN_EXISTING = 3
Public Const FILE_ATTRIBUTE_NORMAL = &H80
Public Const FILE_FLAG_SEQUENTIAL_SCAN = &H8000000
'/FILE API

Private Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
Private Declare Function waveOutGetNumDevs Lib "winmm.dll" () As Long

Private Const SND_ASYNC = &H1
Private Const SPI_SCREENSAVERRUNNING = 97

Private Enum IP_STATUS
    IP_STATUS_BASE = 11000
    IP_SUCCESS = 0
    IP_BUF_TOO_SMALL = (11000 + 1)
    IP_DEST_NET_UNREACHABLE = (11000 + 2)
    IP_DEST_HOST_UNREACHABLE = (11000 + 3)
    IP_DEST_PROT_UNREACHABLE = (11000 + 4)
    IP_DEST_PORT_UNREACHABLE = (11000 + 5)
    IP_NO_RESOURCES = (11000 + 6)
    IP_BAD_OPTION = (11000 + 7)
    IP_HW_ERROR = (11000 + 8)
    P_PACKET_TOO_BIG = (11000 + 9)
    IP_REQ_TIMED_OUT = (11000 + 10)
    IP_BAD_REQ = (11000 + 11)
    IP_BAD_ROUTE = (11000 + 12)
    IP_TTL_EXPIRED_TRANSIT = (11000 + 13)
    IP_TTL_EXPIRED_REASSEM = (11000 + 14)
    IP_PARAM_PROBLEM = (11000 + 15)
    IP_SOURCE_QUENCH = (11000 + 16)
    IP_OPTION_TOO_BIG = (11000 + 17)
    IP_BAD_DESTINATION = (11000 + 18)
    IP_ADDR_DELETED = (11000 + 19)
    IP_SPEC_MTU_CHANGE = (11000 + 20)
    IP_MTU_CHANGE = (11000 + 21)
    IP_UNLOAD = (11000 + 22)
    IP_ADDR_ADDED = (11000 + 23)
    IP_GENERAL_FAILURE = (11000 + 50)
    MAX_IP_STATUS = 11000 + 50
    IP_PENDING = (11000 + 255)
    PING_TIMEOUT = 200
End Enum
Private Const MAX_WSADescription = 256
Private Const MAX_WSASYSStatus = 128
Private Const ERROR_SUCCESS       As Long = 0
Private Const WS_VERSION_REQD     As Long = &H101
Private Const WS_VERSION_MAJOR    As Long = WS_VERSION_REQD \ &H100 And &HFF&
Private Const WS_VERSION_MINOR    As Long = WS_VERSION_REQD And &HFF&
Private Const MIN_SOCKETS_REQD    As Long = 1
Private Const SOCKET_ERROR        As Long = -1

Private Type ICMP_OPTIONS
    Ttl             As Byte
    Tos             As Byte
    Flags           As Byte
    OptionsSize     As Byte
    OptionsData     As Long
End Type

Dim ICMPOPT As ICMP_OPTIONS

Private Type ICMP_ECHO_REPLY
    Address         As Long
    Status          As Long
    RoundTripTime   As Long
    DataSize        As Long
  '  Reserved        As Integer
    DataPointer     As Long
    Options         As ICMP_OPTIONS
    Data            As String * 250
End Type

Private Type HOSTENT
    hName      As Long
    hAliases   As Long
    hAddrType  As Integer
    hLen       As Integer
    hAddrList  As Long
End Type

Private Type WSADATA
    wVersion      As Integer
    wHighVersion  As Integer
    szDescription(0 To MAX_WSADescription)   As Byte
    szSystemStatus(0 To MAX_WSASYSStatus)    As Byte
    wMaxSockets   As Integer
    wMaxUDPDG     As Integer
    dwVendorInfo  As Long
End Type

Public Function SoundSupported() As Boolean
If (waveOutGetNumDevs > 0) Then
    SoundSupported = True
Else
    SoundSupported = False
End If
End Function

Public Sub PlaySound(sPath As String)
On Error Resume Next
If SoundSupported = True Then
    Call sndPlaySound(sPath, SND_ASYNC)
End If
End Sub

Public Function GetIPAddress(Optional sHost As String, Optional sErrMsg As String) As String
'Resolves the host-name (or current machine if balnk) to an IP address
Dim sHostName   As String * 256
Dim lpHost      As Long
Dim Host        As HOSTENT
Dim dwIPAddr    As Long
Dim tmpIPAddr() As Byte
Dim i           As Integer
Dim sIPAddr     As String
Dim werr        As Long

    If Not SocketsInitialize() Then
        GetIPAddress = ""
        Exit Function
    End If
    
    If sHost = "" Then
        If gethostname(sHostName, 256) = SOCKET_ERROR Then
            werr = WSAGetLastError()
            GetIPAddress = "Unknown"
            sErrMsg = "Windows Sockets error " & Str$(werr) & " has occurred. Unable to successfully get Host Name." & vbCrLf
            GetIPAddress = "Unknown"
            SocketsCleanup
            Exit Function
        End If

        sHostName = Trim$(sHostName)
    Else
        sHostName = Trim$(sHost) & Chr$(0)
    End If
    
    lpHost = gethostbyname(sHostName)

    If lpHost = 0 Then
        werr = WSAGetLastError()
        GetIPAddress = "Unknown"
        sErrMsg = "Windows Sockets error " & Str$(werr) & _
                " has occurred. Unable to successfully get Host Name." & vbCrLf
        GetIPAddress = "Unknown"
        
        SocketsCleanup
        Exit Function
    End If

    CopyMemory Host, lpHost, Len(Host)
    CopyMemory dwIPAddr, Host.hAddrList, 4

    ReDim tmpIPAddr(1 To Host.hLen)
    CopyMemory tmpIPAddr(1), dwIPAddr, Host.hLen

    For i = 1 To Host.hLen
        sIPAddr = sIPAddr & tmpIPAddr(i) & "."
    Next

    GetIPAddress = Mid$(sIPAddr, 1, Len(sIPAddr) - 1)

    SocketsCleanup
End Function

Public Function GetIPHostName() As String
'Returns the current machine's name
Dim sHostName As String * 256

    If Not SocketsInitialize() Then
        GetIPHostName = ""
        Exit Function
    End If

    If gethostname(sHostName, 256) = SOCKET_ERROR Then
        GetIPHostName = "Unknown"
        MsgBox "Windows Sockets error " & Str$(WSAGetLastError()) & _
                " has occurred.  Unable to successfully get Host Name."
        SocketsCleanup
        Exit Function
    End If

    GetIPHostName = Left$(sHostName, InStr(sHostName, Chr(0)) - 1)
    SocketsCleanup
End Function

Private Function HiByte(ByVal wParam As Integer)
    HiByte = wParam \ &H1 And &HFF&
End Function

Private Function LoByte(ByVal wParam As Integer)
    LoByte = wParam And &HFF&
End Function

Private Sub SocketsCleanup()
    If WSACleanup() <> ERROR_SUCCESS Then
        App.LogEvent "Socket error occurred in Cleanup.", vbLogEventTypeError
    End If
End Sub

Private Function SocketsInitialize(Optional sErr As String) As Boolean
Dim WSAD As WSADATA, sLoByte As String, sHiByte As String
    If WSAStartup(WS_VERSION_REQD, WSAD) <> ERROR_SUCCESS Then
        sErr = "The 32-bit Windows Socket is not responding."
        SocketsInitialize = False
        Exit Function
    End If

    If WSAD.wMaxSockets < MIN_SOCKETS_REQD Then
        sErr = "This application requires a minimum of " & CStr(MIN_SOCKETS_REQD) & " supported sockets."

        SocketsInitialize = False
        Exit Function
    End If

    If LoByte(WSAD.wVersion) < WS_VERSION_MAJOR Or (LoByte(WSAD.wVersion) = WS_VERSION_MAJOR And HiByte(WSAD.wVersion) < WS_VERSION_MINOR) Then

        sHiByte = CStr(HiByte(WSAD.wVersion))
        sLoByte = CStr(LoByte(WSAD.wVersion))

        sErr = "Sockets version " & sLoByte & "." & sHiByte & " is not supported by 32-bit Windows Sockets."

        SocketsInitialize = False
        Exit Function
    End If
    SocketsInitialize = True
End Function

Public Function GetSystemDirectory() As String
    Dim strBuffer As String, lngReturn As String
    strBuffer = Space(255)
    lngReturn = GetSysDirectory(strBuffer, Len(strBuffer))
    GetSystemDirectory = Left(strBuffer, lngReturn)
End Function

Public Sub CloseCD()
mciSendString "set CDAudio door closed", vbNullString, 0, 0
End Sub

Public Sub OpenCD()
mciSendString "set CDAudio door open", vbNullString, 0, 0
End Sub

Public Sub DisableCAD()
'Disables the Crtl+Alt+Del
Dim Ret As Integer
Dim pOld As Boolean
Ret = SystemParametersInfo(SPI_SCREENSAVERRUNNING, True, pOld, 0)
End Sub

Public Sub EnableCAD()
'Enables the Crtl+Alt+Del
Dim Ret As Integer
Dim pOld As Boolean
Ret = SystemParametersInfo(SPI_SCREENSAVERRUNNING, False, pOld, 0)
End Sub

Public Function ReadFile(PathName As String) As String
Dim hFile As Long
Dim RawData() As Byte
Dim ReadString As String
Dim FileLength  As Long
Dim ActualBytes As Long
Dim Ret As Long
    
FileLength = FileLen(PathName)
ReDim RawData(FileLength)
    
hFile = CreateFile(PathName & vbNullChar, GENERIC_READ, 0, 0, OPEN_EXISTING, FILE_FLAG_SEQUENTIAL_SCAN, 0)

If hFile = 0 Then
    MsgBox "Open Error!"
    Exit Function
End If
    
Ret = ReadFileAPI(hFile, RawData(0), FileLength, ActualBytes, 0)

If Ret = 0 Or ActualBytes <> FileLength Then
    MsgBox "Read Error!"
    CloseHandle hFile
    Exit Function
End If

CloseHandle hFile
ReadFile = StrConv(RawData, vbUnicode)
End Function


