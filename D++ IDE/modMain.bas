Attribute VB_Name = "modMain"
'D++ IDE Module
'You might be thinking "Wow, thats a lot of junk".  Well, good news
'If your looking for the stuff on compiling, the only thing you need
'is 'CompileFile'.  No API or anything else.  Most of the functions
'here are for the IDE.
'D++ has also been modified for speed.  The DLL is now twice as fast
'and the compiling has also greatly increased in speed.  To see how
'fast it compiles, run a program and look at the Debug.  The first
'time is usually a little slower then the rest.


'Get System Directory
Private Declare Function GetSysDirectory Lib "kernel32" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long

'Get Desktop Directory
Private Declare Function SHGetSpecialFolderPath Lib "shell32" Alias "SHGetSpecialFolderPathA" (ByVal hwndOwner As Long, ByVal pszPath As String, ByVal nFolder As Long, ByVal fCreate As Boolean) As Long
Private Declare Function SHGetSpecialFolderLocation Lib "shell32.dll" (ByVal hwndOwner As Long, ByVal nFolder As Long, pidl As Long) As Long
Private Declare Function SHGetPathFromIDList Lib "shell32.dll" Alias "SHGetPathFromIDListA" (ByVal pidl As Long, ByVal pszPath As String) As Long
Private Declare Sub CoTaskMemFree Lib "ole32.dll" (ByVal pv As Long)
Private Const NOERROR = 0
Private Const MAX_PATH = 260
Private Const CSIDL_DESKTOPDIRECTORY = &H10

'For File Association
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Private Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal hKey As Long, ByVal lpSubKey As String) As Long
Private Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal hKey As Long, ByVal lpValueName As String) As Long
Private Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
Private Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long
Private Const REG_SZ = 1
Private Const REG_DWORD = 4
Private Const HKEY_CLASSES_ROOT = &H80000000
Private Const HKEY_CURRENT_USER = &H80000001
Private Const HKEY_LOCAL_MACHINE = &H80000002
Private Const HKEY_USERS = &H80000003
Private Const HKEY_PERFORMANCE_DATA = &H80000004

'D++ stuff
Private Declare Function GetTickCount Lib "kernel32" () As Long
Public DLLFILE As String
Public StopDebug As Boolean
Public DLLVersion As Integer

Public Errors As Integer

Public Sub SetDllLocation()
'This sub sets the DLL location (DPPAPP.dll)
If FileExist(GetSystemDirectory & "\DPPAPP.DLL") Then 'Check system folder
    DLLFILE = GetSystemDirectory & "\DPPAPP.DLL"
ElseIf FileExist(AppPath & "DPPAPP.DLL") Then 'Check current directory
    DLLFILE = AppPath & "DPPAPP.DLL"
Else 'Prompt for download
    response = MsgBox("A required file to run D++ IDE, DPPAPP.DLL, was not found.  Would you like to download this file?", vbYesNo + vbExclamation, "File Not Found")
    If response = vbYes Then
        frmDownload.cmdQuit.Visible = True
        frmDownload.cmdClose.Enabled = False
        frmDownload.Show 1 'Download new file
    Else
        End
    End If
End If
DLLVersion = GetDLLVersion
If DLLVersion < 193 Then
    response = MsgBox("The DLL you have is currently have is incompatible with this compiler." & vbCrLf & "Would you like to download the newer version?", vbYesNo + vbExclamation, "File Not Found")
    If response = vbYes Then
        frmDownload.cmdQuit.Visible = True
        frmDownload.cmdClose.Enabled = False
        frmDownload.Show 1 'Download new file
    Else
        End
    End If
ElseIf DLLVersion = 999 Then
    MsgBox "An error occured while attempting to detect the DLL version.  It is possible " & vbCrLf & "that the DLL is either incompatible or corrupt.  If the compiler does not work,  " & vbCrLf & "please download the newest DLL by pressing F4 in the IDE." & vbCrLf & vbCrLf & "If you continue getting this error, and / or the compiler still does not work, " & vbCrLf & "please visit PageMac Programming for technical support: " & vbCrLf & vbCrLf & "http://squeakmac.tripod.com", vbExclamation, "Error Deteting DLL"
End If
End Sub

Public Sub Pause(interval)
'This sub pauses the program
Current = Timer
Do While Timer - Current < Val(interval)
    DoEvents
Loop
End Sub

Public Function GetDesktopDirectory(hWnd As Long)
'This function gets the desktop directory
On Error GoTo NotExported
Dim pidl As Long, sPath As String * MAX_PATH, nFolder As Long

nFolder = CSIDL_DESKTOPDIRECTORY
Call SHGetSpecialFolderPath(hWnd, sPath, nFolder, 0)
If InStr(sPath, vbNullChar) > 1 Then
    GetDesktopDirectory = Left$(sPath, InStr(sPath, vbNullChar) - 1)
    Exit Function
End If

NotExported:
If SHGetSpecialFolderLocation(hWnd, nFolder, pidl) = NOERROR Then
    If pidl Then
        If SHGetPathFromIDList(pidl, sPath) Then
            GetDesktopDirectory = Left$(sPath, InStr(sPath, vbNullChar) - 1)
        End If
        Call CoTaskMemFree(pidl)
    End If
End If
End Function

Public Function GetSystemDirectory() As String
'This function gets the system directory
Dim strBuffer As String, lngReturn As String
strBuffer = Space(255)
lngReturn = GetSysDirectory(strBuffer, Len(strBuffer))
GetSystemDirectory = Left(strBuffer, lngReturn)
End Function

Public Sub AddDebug(TextToAdd As String)
On Error Resume Next
'Adds text to the Debug window
frmMain.txtDebug.Text = frmMain.txtDebug.Text & TextToAdd & vbCrLf
frmMain.lstPos.AddItem 0
If Mid(TextToAdd, 1, 6) = ">Error" Then Errors = Errors + 1
DoEvents
End Sub

Public Sub AddError(ErrorText As String, ErrorPos As Long)
On Error Resume Next
If CheckIDE = True Then Exit Sub
Dim x As Long
'Adds text to the Debug window, without returns

x = InStr(1, ErrorText, vbCrLf)
If x <> 0 Then
    ErrorText = Mid(ErrorText, 1, x - 1) & " ... (Error truncated)"
End If

frmMain.txtDebug.Text = frmMain.txtDebug.Text & ErrorText & vbCrLf
frmMain.lstPos.AddItem ErrorPos
Errors = Errors + 1
DoEvents
End Sub

Public Function FileExist(ByVal FileName As String) As Boolean
'Determines if a file exists
On Error Resume Next
If Dir(FileName, vbSystem + vbHidden) = "" Then
    FileExist = False
Else
    FileExist = True
End If
End Function

Public Sub ShowInformation(Information As String, Title As String, Optional ButtonCaption As String = "Ok")
'This shows information
frmInformation.txtText.Text = Information
frmInformation.Caption = Title
frmInformation.cmdOK.Caption = ButtonCaption
Beep
frmInformation.Show 1
End Sub

Public Function GetFileName(FullPath As String) As String
On Error Resume Next
'This parses a full path to get the file name
For i = Len(FullPath) To 1 Step -1
    sChar = Mid(FullPath, i, 1)
    If sChar = "\" Then
        GetFileName = StrReverse(GetFileName)
        Exit For
    End If
    GetFileName = GetFileName & sChar
Next i
End Function

Public Function GetExtension(FullPath As String) As String
'This parses a full path to get the file name
On Error Resume Next
For i = Len(FullPath) To 1 Step -1
    sChar = Mid(FullPath, i, 1)
    If sChar = "." Then
        GetExtension = StrReverse(GetExtension)
        Exit For
    End If
    GetExtension = GetExtension & sChar
Next i
End Function

Public Function GetLineNum(Source As String, SelPos As Integer) As Integer
'this gets the line number where the cursor is at
Dim Start As Integer
Dim Lines As Integer

Lines = 1
Start = InStr(1, Source, vbCrLf)
Do Until Start = 0 Or Start > SelPos
    Lines = Lines + 1
    Start = InStr(Start + 2, Source, vbCrLf)
Loop

GetLineNum = Lines
End Function

Public Sub MakeEXE(UseDOS As Boolean)
On Error GoTo Errorh
  
frmMain.CommonDialog1.Filter = "EXE Files (*.exe)|*.exe|All Files (*.*)|*.*"
frmMain.CommonDialog1.DialogTitle = "Save D++ File"
frmMain.CommonDialog1.CancelError = True
frmMain.CommonDialog1.ShowSave

APPFILE = frmMain.CommonDialog1.FileName

If FileExist(APPFILE) Then
    Select Case GetSetting("D++", "Options", "Compile")
        Case -1
            AddDebug ">EXE overwritten."
        Case 0
            AddDebug ">File exists! Aborting..."
            AddDebug ">Application Terminated."
            Exit Sub
        Case 1
            Overwrite = MsgBox("Overwrite Existing .EXE File?", 276, "File Found!")
            If Overwrite = 6 Then
                AddDebug ">EXE overwritten."
            Else
                AddDebug ">Canceled by user."
                AddDebug ">Application Terminated."
                Exit Sub
            End If
    End Select
End If

If UseDOS Then
    CompileFile APPFILE, False, True
Else
    CompileFile APPFILE, False, False
End If
SaveSetting "D++", "Memory", "Compile", APPFILE

Exit Sub

Errorh:
If Err.Number = 0 Then Exit Sub
If Err.Number = 32755 Then
    AddDebug ">Canceled compile."
    Exit Sub
End If
AddDebug ">Compile Error " & Err.Number & ": " & Err.Description
AddDebug ">D++ Application Terminated."
End Sub

Public Sub Run(UseDOS As Boolean)
On Error GoTo Errorh

APPFILE = GetSetting("D++", "Options", "RunAt") & "\D++APP1.EXE"

If FileExist(APPFILE) Then
    Select Case GetSetting("D++", "Options", "Run")
        Case 0
            AddDebug ">File exists! Aborting..."
            AddDebug ">Application Terminated."
            Exit Sub
        Case 1
            Overwrite = MsgBox("Overwrite Existing .EXE File?", 276, "File Found!")
            If Overwrite <> 6 Then
                AddDebug ">Canceled by user."
                AddDebug ">Application Terminated."
                Exit Sub
            End If
    End Select
End If

If UseDOS Then
    CompileFile APPFILE, True, True
Else
    CompileFile APPFILE, True, False
End If

SaveSetting "D++", "Memory", "Run", APPFILE
    
Exit Sub
    
Errorh:
If Err.Number = 0 Then Exit Sub
AddDebug ">Compile Error " & Err.Number & ": " & Err.Description
AddDebug ">D++ Application Terminated."
End Sub

Public Sub CompileFile(FilePath, Link As Boolean, UseDOS As Boolean)
'Compiles the DLL file (blank app with linker), a header, and source
'Thats how simple D++ works
Dim FileData As String
Dim StartTime As Single
Dim Header As String, SourceCode As String
Dim Encrypt As Boolean

StartTime = GetTickCount 'Set start time
If Header = "" Then Header = "DPP:" 'make sure there's a header
SourceCode = frmMain.txtText.Text 'Get Source Code
SetDllLocation 'Check for DLL

Errors = 0
If GetSetting("D++", "Options", "Debugging") = "0" Then             'check if debugging is enabled
    AddDebug ">Debugging syntax..."                                 'syntax..
    CheckCode (SourceCode)                                          'check syntax
End If

If Errors = 0 Then
    
    AddDebug ">Assembling Header..."
    If UseDOS Then                                                  'check if using dos...
        Header = "dppapp:type>dos>"                                 'set the header (dos)
    Else
        Header = "dppapp:type>dpp>"                                 'set the header (normal)
    End If

    If GetSetting("D++", "Options", "Encrypt") = "DPP:$D2>:" Then   'check encryption
        SourceCode = Crypt(SourceCode)                              'encrypt
        Header = Header & "crypt>t>"                                'set the header (crypt)
    Else
        Header = Header & "crypt>f>"                                'set the header (normal)
    End If

    If Not Link Then                                                'If we're not linking
        If GetSetting("D++", "Options", "Decompile") = "-1" Then    'if setting is true
            Header = Header & "read>f"                              'make it so it can't decompile
        Else
            Header = Header & "read>t"
        End If
    Else
        Header = Header & "read>t"                                  'make it so it can decompile
    End If
        
    Header = Header & ">dpp:"                                       'finish header
        
    AddDebug ">Compiling Project..."                                'debug..
    FileNum = FreeFile 'get free file number

    'Read the DLL file and get the data
    Open DLLFILE For Binary As #FileNum
        FileData = Space$(LOF(FileNum)) 'set file data length
        Get #FileNum, , FileData 'get file data
    Close #FileNum

    FileNum = FreeFile 'get free file number

    'Write the new file to the location
    Open FilePath For Output As #FileNum
        'The new D++ file will be the DLL file, a header, and the code
        Print #FileNum, FileData & Header & SourceCode 'write file
    Close #FileNum


    If Link = True Then 'check if linking...
        AddDebug ">Linking " & FilePath & "..." 'debug..
        Shell FilePath, vbNormalFocus 'run the file
    End If
    
    If GetSetting("D++", "Options", "Debugging") = "0" Then 'if it checked errors
        AddDebug ">Finished. Found " & Errors & " error(s). Compile Time: " & (GetTickCount - StartTime) & " milliseconds."
    Else 'nope
        AddDebug ">Finished. Skipped Error Check. Compile Time: " & (GetTickCount - StartTime) & " milliseconds."
    End If
Else 'found some errors in the program
    AddDebug ">Finished. Found " & Errors & " error(s). "
    'frmMain.lstDebug.ListIndex = frmMain.lstDebug.ListCount - 1
    Beep
End If
frmMain.txtDebug.SelStart = Len(frmMain.txtDebug.Text)
End Sub

Public Function ReadEXE(sPath As String) As String
'Reads an EXE file and gets D++ source
'On Error GoTo Errorh
Dim FileData As String
Dim Header As String, EXE_Type As String
Dim Encrypt_Source As String, Decompile As String
    
Open sPath For Binary As #1 'open file
    FileData = Space$(LOF(1) - 2) 'set file data length
    Get #1, , FileData 'get file data
Close #1 'close file

If InStr(1, FileData, "dppapp:") = 0 Or InStr(1, FileData, "dpp:") = 0 Then
    If InStr(1, FileData, "DPP:$D2>:") = 0 Then
        ReadEXE = "'Note: No source code found."
        Exit Function
    Else
        ReadEXE = "'Note: This application was compiled in a " & vbCrLf & "'previous version of the D++ compiler." & vbCrLf & vbCrLf
        ReadEXE = ReadEXE & Crypt(Mid(FileData, InStr(1, FileData, "DPP:$D2>:") + 9))
        Exit Function
    End If
End If

Header = Mid(FileData, InStr(1, FileData, "dppapp:"), 35)
EXE_Type = Mid(Header, 13, 3)
Encrypt_Source = Mid(Header, 23, 1)
Decompile = Mid(Header, 30, 1)

Select Case EXE_Type
    Case "dos", "dpp"
        'don't do anything, it's good
    Case "reg"
        ReadEXE = "'Note: This is a temporary EXE.  It does not contain source."
        Exit Function
    Case Else
        MsgBox "Fatal Read Error: Incorrect application type.  Please recompile." & vbCrLf & vbCrLf & "If this message continues, please download the latest D++ compiler at " & vbCrLf & vbCrLf & "http://squeakmac.tripod.com", vbCritical, "Invalid App Type"
End Select

If Encrypt_Source = "t" Then 'Parse file to get source
    ReadEXE = Crypt(Mid(FileData, InStr(1, FileData, "dpp:") + 4))
ElseIf Encrypt_Source = "f" Then
    ReadEXE = Mid(FileData, InStr(1, FileData, "dpp:") + 4)
Else
    MsgBox "Fatal Read Error: Incorrect application type.  Please recompile." & vbCrLf & vbCrLf & "If this message continues, please download the latest D++ compiler at " & vbCrLf & vbCrLf & "http://squeakmac.tripod.com", vbCritical, "Invalid App Type"
End If

If Decompile = "f" Then
    MsgBox "This EXE is protected.  Unable to decompile EXE.", vbCritical, "EXE Protected"
    ReadEXE = "'Note: Pretected EXE.  Cannot Decompile"
ElseIf Decompile <> "t" Then
    MsgBox "Fatal Read Error: Unable to determine security level.  Please recompile." & vbCrLf & "If this message continues, please download the latest D++ compiler at " & vbCrLf & vbCrLf & "http://squeakmac.tripod.com" & vbCrLf & vbCrLf & "The file will still be read.  No security enabled.", vbCritical, "Invalid App Type"
End If
    
If ReadEXE = "" Then ReadEXE = "Note: No source code found."

Exit Function

Errorh:
If Err.Number <> 0 Then MsgBox "Link Error #" & Err.Number & " has occured: " & Err.Description, vbCritical, "Link Error"
End
End Function

Public Function AppPath() As String
'This gets the app's path
If Right(App.Path, 1) = "\" Then AppPath = App.Path Else AppPath = App.Path & "\"
End Function

Function Crypt(Text) As String
'This is a simple crypter
For i = 1 To Len(Text)
    DoEvents
    Crypt = Crypt & Chr(255 - Asc(Mid(Text, i, 1)))
Next i
End Function

Public Function GetDLLVersion() As Double
Dim FileNum As Integer
Dim FileData As String
Dim Header As String

    On Error GoTo Err
    
    Header = "dppapp:type>reg>crypt>f>source>dpp:"
    
    FileNum = FreeFile 'get free file number

    'Read the DLL file and get the data
    Open DLLFILE For Binary As #FileNum
        FileData = Space$(LOF(FileNum)) 'set file data length
        Get #FileNum, , FileData 'get file data
    Close #FileNum

    FileNum = FreeFile 'get free file number

    'Write the new file to the location
    Open "C:\tempdpp.exe" For Output As #FileNum
        'The new D++ file will be the DLL file, a header, and the code
        Print #FileNum, FileData & Header & "end;" 'write file
    Close #FileNum
    
    Shell "C:\tempdpp.exe", vbNormalFocus
    GetDLLVersion = GetSetting("D++", "Version", "DLL_Version")
    If GetDLLVersion = 0 Then GoTo Err
    
Exit Function
Err:
GetDLLVersion = "999"
End Function

'This next section is regestry
Public Sub SaveKey(hKey As Long, strPath As String)
    Dim keyhand&
    r = RegCreateKey(hKey, strPath, keyhand&)
    r = RegCloseKey(keyhand&)
End Sub

Public Function GetString(hKey As Long, strPath As String, strValue As String)
    Dim keyhand As Long
    Dim datatype As Long
    Dim lResult As Long
    Dim strBuf As String
    Dim lDataBufSize As Long
    Dim intZeroPos As Integer
    r = RegOpenKey(hKey, strPath, keyhand)
    lResult = RegQueryValueEx(keyhand, strValue, 0&, lValueType, ByVal 0&, lDataBufSize)


    If lValueType = REG_SZ Then
        strBuf = String(lDataBufSize, " ")
        lResult = RegQueryValueEx(keyhand, strValue, 0&, 0&, ByVal strBuf, lDataBufSize)


        If lResult = ERROR_SUCCESS Then
            intZeroPos = InStr(strBuf, Chr$(0))


            If intZeroPos > 0 Then
                GetString = Left$(strBuf, intZeroPos - 1)
            Else
                GetString = strBuf
            End If
        End If
    End If
End Function

Public Sub SaveString(hKey As Long, strPath As String, strValue As String, strdata As String)
    Dim keyhand As Long
    Dim r As Long
    r = RegCreateKey(hKey, strPath, keyhand)
    r = RegSetValueEx(keyhand, strValue, 0, REG_SZ, ByVal strdata, Len(strdata))
    r = RegCloseKey(keyhand)
End Sub

Function SaveDword(ByVal hKey As Long, ByVal strPath As String, ByVal strValueName As String, ByVal lData As Long)
    Dim lResult As Long
    Dim keyhand As Long
    Dim r As Long
    r = RegCreateKey(hKey, strPath, keyhand)
    lResult = RegSetValueEx(keyhand, strValueName, 0&, REG_DWORD, lData, 4)
    r = RegCloseKey(keyhand)
End Function

Public Sub AssociateFiles()
'You MUST compile the EXE before this will work!
Dim IconPath As String
IconPath = GetSystemDirectory & "\dppicon.ico"
SavePicture frmMain.Icon, IconPath
Call SaveString(HKEY_CLASSES_ROOT, "\.dpp", "", "dppfile")
Call SaveString(HKEY_CLASSES_ROOT, "\.dpp", "Content Type", "application/dpp")
Call SaveString(HKEY_CLASSES_ROOT, "\dppfile", "", "D++ Files")
Call SaveDword(HKEY_CLASSES_ROOT, "\dppfile", "EditFlags", "0000")
Call SaveString(HKEY_CLASSES_ROOT, "\dppfile\DefaultIcon", "", IconPath)
Call SaveString(HKEY_CLASSES_ROOT, "\dppfile\Shell", "", "")
Call SaveString(HKEY_CLASSES_ROOT, "\dppfile\Shell\Open", "", "")
Call SaveString(HKEY_CLASSES_ROOT, "\dppfile\Shell\Open\command", "", Chr(34) & App.Path + "\" + App.EXEName + ".EXE" & Chr(34) & " %1")
End Sub
