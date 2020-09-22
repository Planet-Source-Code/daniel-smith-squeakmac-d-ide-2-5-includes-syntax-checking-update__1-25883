Attribute VB_Name = "modFunctions"
'D++ Function Module

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

'for file downloading...
Private Declare Function GetDiskFreeSpace Lib "kernel32" Alias "GetDiskFreeSpaceA" (ByVal lpRootPathName As String, lpSectorsPerCluster As Long, lpBytesPerSector As Long, lpNumberOfFreeClusters As Long, lpTotalNumberOfClusters As Long) As Long

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



'-----------------------------------------------------------------------
'THE FOLLOWING IS NOT MINE
'This is from PSC, Jeff Cockayne
'-----------------------------------------------------------------------


Public Function FitText(ByRef Ctl As Control, _
                        ByVal strCtlCaption) As String

' Function FitText
' Author:   Jeff Cockayne
'
' Fit the caption text passed in strCtlCaption
' to the width of the passed Control, Ctl.
' There are a few ways to blow this function, like
' passing a control without a Caption Property, but
' this Function is for internal use, so...
'
' Example:
' If "C:\Program Files\Test.TXT" was too wide, the
' returned string might be: "C:\Pro...\Test.TXT"

Dim lngCtlLeft As Long
Dim lngMaxWidth As Long
Dim lngTextWidth As Long
Dim lngX As Long

' Store frequently referenced values to increase
' performance (saves some OLE lookup)
lngCtlLeft = Ctl.Left
lngMaxWidth = Ctl.Width
lngTextWidth = Ctl.Parent.TextWidth(strCtlCaption)


lngX = (Len(strCtlCaption) \ 2) - 2
While lngTextWidth > lngMaxWidth And lngX > 3
    ' Text is too wide for Ctl's width;
    ' shrink the caption from the middle,
    ' replacing the 3 middlemost characters
    ' with ellipses (...)
    strCtlCaption = Left(strCtlCaption, lngX) & "..." & _
                    Right(strCtlCaption, lngX)
    lngTextWidth = Ctl.Parent.TextWidth(strCtlCaption)
    lngX = lngX - 1
Wend

FitText = strCtlCaption

End Function

Public Function FormatFileSize(ByVal dblFileSize As Double, _
                               Optional ByVal strFormatMask As String) _
                               As String

' FormatFileSize:   Formats dblFileSize in bytes into
'                   X GB or X MB or X KB or X bytes depending
'                   on size (a la Win9x Properties tab)

Select Case dblFileSize
    Case 0 To 1023              ' Bytes
        FormatFileSize = Format(dblFileSize) & " bytes"
    Case 1024 To 1048575        ' KB
        If strFormatMask = Empty Then strFormatMask = "###0"
        FormatFileSize = Format(dblFileSize / 1024#, strFormatMask) & " KB"
    Case 1024# ^ 2 To 1073741823 ' MB
        If strFormatMask = Empty Then strFormatMask = "###0.0"
        FormatFileSize = Format(dblFileSize / (1024# ^ 2), strFormatMask) & " MB"
    Case Is > 1073741823#       ' GB
        If strFormatMask = Empty Then strFormatMask = "###0.0"
        FormatFileSize = Format(dblFileSize / (1024# ^ 3), strFormatMask) & " GB"
End Select

End Function

Public Function FormatTime(ByVal sglTime As Single) As String
                           
' FormatTime:   Formats time in seconds to time in
'               Hours and/or Minutes and/or Seconds

' Determine how to display the time
Select Case sglTime
    Case 0 To 59    ' Seconds
        FormatTime = Format(sglTime, "0") & " sec"
    Case 60 To 3599 ' Minutes Seconds
        FormatTime = Format(Int(sglTime / 60), "#0") & _
                     " min " & _
                     Format(sglTime Mod 60, "0") & " sec"
    Case Else       ' Hours Minutes
        FormatTime = Format(Int(sglTime / 3600), "#0") & _
                     " hr " & _
                     Format(sglTime / 60 Mod 60, "0") & " min"
End Select

End Function

Public Function DiskFreeSpace(strDrive As String) As Double

' DiskFreeSpace:    returns the amount of free space on a drive
'                   in Windows9x/2000/NT4+

Dim SectorsPerCluster As Long
Dim BytesPerSector As Long
Dim NumberOfFreeClusters As Long
Dim TotalNumberOfClusters As Long
Dim FreeBytes As Long
Dim spaceInt As Integer

strDrive = QualifyPath(strDrive)

' Call the API function
GetDiskFreeSpace strDrive, _
                 SectorsPerCluster, _
                 BytesPerSector, _
                 NumberOFreeClusters, _
                 TotalNumberOfClusters

' Calculate the number of free bytes
DiskFreeSpace = NumberOFreeClusters * SectorsPerCluster * BytesPerSector

End Function


Public Function QualifyPath(strPath As String) As String

' Make sure the path ends in "\"
QualifyPath = IIf(Right(strPath, 1) = "\", strPath, strPath & "\")

End Function


Public Function ReturnFileOrFolder(FullPath As String, _
                                   ReturnFile As Boolean, _
                                   Optional IsURL As Boolean = False) _
                                   As String

' ReturnFileOrFolder:   Returns the filename or path of an
'                       MS-DOS file or URL.
'
' Author:   Jeff Cockayne 4.30.99
'
' Inputs:   FullPath:   String; the full path
'           ReturnFile: Boolean; return filename or path?
'                       (True=filename, False=path)
'           IsURL:      Boolean; Pass True if path is a URL.
'
' Returns:  String:     the filename or path
'

Dim intDelimiterIndex As Integer

intDelimiterIndex = InStrRev(FullPath, IIf(IsURL, "/", "\"))
If intDelimiterIndex = 0 Then
    ReturnFileOrFolder = FullPath
Else
    ReturnFileOrFolder = IIf(ReturnFile, _
                         Right(FullPath, Len(FullPath) - intDelimiterIndex), _
                         Left(FullPath, intDelimiterIndex))
End If

End Function


