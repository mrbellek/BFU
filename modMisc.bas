Attribute VB_Name = "modMisc"
 Option Explicit
'misc functions/subs

Private Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long
Private Declare Function GetSaveFileName Lib "comdlg32.dll" Alias "GetSaveFileNameA" (pOpenfilename As OPENFILENAME) As Long
Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As Any) As Long
Private Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Private Declare Function GetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Private Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long

Private Declare Function ExpandEnvironmentStrings Lib "kernel32" Alias "ExpandEnvironmentStringsA" (ByVal lpSrc As String, ByVal lpDst As String, ByVal nSize As Long) As Long
Private Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Private Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long

Private Declare Function GetShortPathName Lib "kernel32" Alias "GetShortPathNameA" (ByVal lpszLongPath As String, ByVal lpszShortPath As String, ByVal cchBuffer As Long) As Long

Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Public Declare Function GetLastError Lib "kernel32" () As Long

Private Declare Function CryptAcquireContext Lib "advapi32.dll" Alias "CryptAcquireContextA" (ByRef phProv As Long, ByVal pszContainer As String, ByVal pszProvider As String, ByVal dwProvType As Long, ByVal dwFlags As Long) As Long
Private Declare Function CryptCreateHash Lib "advapi32.dll" (ByVal hProv As Long, ByVal Algid As Long, ByVal hKey As Long, ByVal dwFlags As Long, ByRef phHash As Long) As Long
Private Declare Function CryptDestroyHash Lib "advapi32.dll" (ByVal hHash As Long) As Long
Private Declare Function CryptGetHashParam Lib "advapi32.dll" (ByVal pCryptHash As Long, ByVal dwParam As Long, ByRef pbData As Any, ByRef pcbData As Long, ByVal dwFlags As Long) As Long
Private Declare Function CryptHashData Lib "advapi32.dll" (ByVal hHash As Long, ByVal pbData As String, ByVal dwDataLen As Long, ByVal dwFlags As Long) As Long
Private Declare Function CryptReleaseContext Lib "advapi32.dll" (ByVal hProv As Long, ByVal dwFlags As Long) As Long

Public Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function ReadFile Lib "kernel32" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToRead As Long, lpNumberOfBytesRead As Long, lpOverlapped As Any) As Long
Private Declare Function WriteFile Lib "kernel32" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToWrite As Long, lpNumberOfBytesWritten As Long, lpOverlapped As Any) As Long
Private Declare Function CreateFile Lib "kernel32" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, lpSecurityAttributes As Any, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Private Declare Function GetFileSize Lib "kernel32" (ByVal hFile As Long, lpFileSizeHigh As Long) As Long
Private Declare Function SetFilePointer Lib "kernel32" (ByVal hFile As Long, ByVal lDistanceToMove As Long, lpDistanceToMoveHigh As Long, ByVal dwMoveMethod As Long) As Long

Private Declare Function InternetOpen Lib "wininet.dll" Alias "InternetOpenA" (ByVal sAgent As String, ByVal lAccessType As Long, ByVal sProxyName As String, ByVal sProxyBypass As String, ByVal lFlags As Long) As Long
Private Declare Function InternetOpenUrl Lib "wininet.dll" Alias "InternetOpenUrlA" (ByVal hInternetSession As Long, ByVal sURL As String, ByVal sHeaders As String, ByVal lHeadersLength As Long, ByVal lFlags As Long, ByVal lContext As Long) As Long
Private Declare Function InternetReadFile Lib "wininet.dll" (ByVal hFile As Long, ByVal sBuffer As String, ByVal lNumBytesToRead As Long, lNumberOfBytesRead As Long) As Long
Private Declare Function InternetCloseHandle Lib "wininet.dll" (ByVal hInet As Long) As Long

Private Type OSVERSIONINFO
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformId As Long
    szCSDVersion As String * 128
End Type

 Private Type OSVERSIONINFOEX
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformId As Long
    szCSDVersion As String * 128
    wServicePackMajor As Integer 'win2000 only
    wServicePackMinor As Integer 'win2000 only
    wSuiteMask As Integer 'win2000 only
    wProductType As Byte 'win2000 only
    wReserved As Byte
End Type

Private Type OPENFILENAME
    lStructSize As Long
    hwndOwner As Long
    hInstance As Long
    lpstrFilter As String
    lpstrCustomFilter As String
    nMaxCustFilter As Long
    nFilterIndex As Long
    lpstrFile As String
    nMaxFile As Long
    lpstrFileTitle As String
    nMaxFileTitle As Long
    lpstrInitialDir As String
    lpstrTitle As String
    flags As Long
    nFileOffset As Integer
    nFileExtension As Integer
    lpstrDefExt As String
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type

Private Const INTERNET_FLAG_RELOAD = &H80000000
Private Const INTERNET_OPEN_TYPE_DIRECT = 1

Private Const GENERIC_WRITE = &H40000000
Private Const GENERIC_READ = &H80000000
Private Const FILE_SHARE_READ = &H1
Private Const FILE_SHARE_WRITE = &H2
Private Const CREATE_ALWAYS = 2
Private Const OPEN_EXISTING = 3
Private Const FILE_BEGIN = 0

Private Const ALG_TYPE_ANY As Long = 0
Private Const ALG_SID_MD2 = 1
Private Const ALG_SID_MD4 = 2
Private Const ALG_SID_MD5 As Long = 3
Private Const ALG_SID_SHA1 = 4
Private Const ALG_CLASS_HASH As Long = 32768

Private Const HP_HASHVAL As Long = 2
Private Const HP_HASHSIZE As Long = 4

Private Const CRYPT_VERIFYCONTEXT = &HF0000000

Private Const PROV_RSA_FULL As Long = 1
Private Const MS_ENHANCED_PROV As String = "Microsoft Enhanced Cryptographic Provider v1.0"

Private Const VER_PLATFORM_WIN32_WINDOWS = 1
Private Const VER_PLATFORM_WIN32_NT = 2
Private Const VER_NT_WORKSTATION = 1
 
Private Const OFN_OVERWRITEPROMPT = &H2
Private Const OFN_HIDEREADONLY = &H4
Private Const OFN_PATHMUSTEXIST = &H800
Private Const OFN_FILEMUSTEXIST = &H1000

Public bIsWinNT As Boolean, bIsWinME As Boolean, bIsWinNT4 As Boolean
Public sWinDir$, sSysDir$, sTempDir$, sWinVer$
Public sHostsFile$

Public Sub GetWindowsInfo()
    Dim uOVI As OSVERSIONINFO
    With uOVI
        .dwOSVersionInfoSize = Len(uOVI)
        GetVersionEx uOVI
        If .dwPlatformId = VER_PLATFORM_WIN32_NT Then
            bIsWinNT = True
            If .dwMajorVersion = 4 Then bIsWinNT4 = True
            sWinVer = "WinNT "
            MAX_REG_VALUE_NAME = 16400
        Else
            If .dwMajorVersion = 4 And .dwMinorVersion = 90 Then bIsWinME = True
            sWinVer = "Win9x "
            MAX_REG_VALUE_NAME = 260
        End If

        .szCSDVersion = Replace(.szCSDVersion, "Service Pack ", "SP")
        .szCSDVersion = Replace(.szCSDVersion, "Service Pack", "SP")
        
        sWinVer = sWinVer & .dwMajorVersion & "." & _
                  Format(.dwMinorVersion, "00") & "." & _
                  Format(.dwBuildNumber And &HFFF, "0000") & " " & _
                  Trim(TrimNull(.szCSDVersion))
        sWinVer = GetFriendlyWinVer(bIsWinNT, .dwMajorVersion, .dwMinorVersion, .dwBuildNumber, Trim(TrimNull(.szCSDVersion))) & " (" & sWinVer & ")"
    End With
    
    sWinDir = String(260, 0)
    sWinDir = Left(sWinDir, GetWindowsDirectory(sWinDir, Len(sWinDir)))
    sSysDir = String(260, 0)
    sSysDir = Left(sSysDir, GetSystemDirectory(sSysDir, Len(sSysDir)))
    sTempDir = String(260, 0)
    sTempDir = Left(sTempDir, GetTempPath(Len(sTempDir), sTempDir) - 1)
    
    If bIsWinNT Then
        sHostsFile = sSysDir & "\drivers\etc\hosts"
    Else
        sHostsFile = sWinDir & "\hosts"
    End If
End Sub

Public Function CmnDialogGetFilename$(sFilter$, sTitle$)
    Dim uOFN As OPENFILENAME
    With uOFN
        .lStructSize = Len(uOFN)
        .hwndOwner = frmMain.hWnd
        '.lpstrDefExt = "bfu"
        .lpstrFilter = Replace(sFilter, "|", Chr(0)) & Chr(0) & Chr(0)
        .lpstrInitialDir = App.Path
        .lpstrTitle = sTitle
        .lpstrFile = String(260, 0)
        .nMaxFile = Len(.lpstrFile)
        .flags = OFN_FILEMUSTEXIST Or OFN_HIDEREADONLY Or OFN_PATHMUSTEXIST
        If GetOpenFileName(uOFN) <> 0 Then
            CmnDialogGetFilename = TrimNull(.lpstrFile)
        End If
    End With
End Function

Public Function CmnDialogSaveFilename$(sFilter$, sTitle$)
    Dim uOFN As OPENFILENAME
    With uOFN
        .lStructSize = Len(uOFN)
        .hwndOwner = frmMain.hWnd
        '.lpstrDefExt = "bfu"
        .lpstrFilter = Replace(sFilter, "|", Chr(0)) & Chr(0) & Chr(0)
        .lpstrInitialDir = App.Path
        .lpstrTitle = sTitle
        .lpstrFile = String(260, 0)
        .nMaxFile = Len(.lpstrFile)
        .flags = OFN_OVERWRITEPROMPT Or OFN_HIDEREADONLY Or OFN_PATHMUSTEXIST
        If GetSaveFileName(uOFN) <> 0 Then
            CmnDialogSaveFilename = TrimNull(.lpstrFile)
        End If
    End With
End Function

Public Function TrimNull$(s$)
    If InStr(s, Chr(0)) > 0 Then
        TrimNull = Left(s, InStr(s, Chr(0)) - 1)
    Else
        TrimNull = s
    End If
End Function

Public Function ExpandEnvironmentVars$(s$)
    Dim sDummy$, lLen&
    If InStr(s, "%") = 0 Then
        ExpandEnvironmentVars = s
        Exit Function
    End If
    lLen = ExpandEnvironmentStrings(s, ByVal 0, 0)
    If lLen > 0 Then
        sDummy = String(lLen, 0)
        ExpandEnvironmentStrings s, sDummy, Len(sDummy)
        sDummy = TrimNull(sDummy)
        
        If InStr(sDummy, "%") = 0 Then
            ExpandEnvironmentVars = sDummy
            Exit Function
        End If
    Else
        sDummy = s
    End If
    
    Dim sProgramFiles$, sDesktop$, sMyDocuments$, sStartMenu$
    Dim sStartup$, sAllUsersDesktop$, sAllUsersStartMenu$, sQuickLaunch$
    Dim sAppData$, sAllUsersAppData$, sAllUsersStartup$, sShFolders$
    Dim sUsername$, sComputerName$, sPrograms$, sAllUsersPrograms$
    Dim sFavorites$, sAllUsersFavorites$
    sShFolders = "Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders"
    sUsername = String(260, 0)
    GetUserName sUsername, Len(sUsername)
    sUsername = TrimNull(sUsername)
    sComputerName = String(260, 0)
    GetComputerName sComputerName, Len(sComputerName)
    sComputerName = TrimNull(sComputerName)
    
    sProgramFiles = RegGetString(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion", "ProgramFilesDir")
    sDesktop = RegGetString(HKEY_CURRENT_USER, sShFolders, "Desktop")
    sMyDocuments = RegGetString(HKEY_CURRENT_USER, sShFolders, "Personal")
    sStartMenu = RegGetString(HKEY_CURRENT_USER, sShFolders, "Start Menu")
    sStartup = RegGetString(HKEY_CURRENT_USER, sShFolders, "Startup")
    sPrograms = RegGetString(HKEY_CURRENT_USER, sShFolders, "Programs")
    sAppData = RegGetString(HKEY_CURRENT_USER, sShFolders, "AppData")
    sFavorites = RegGetString(HKEY_CURRENT_USER, sShFolders, "Favorites")
    sQuickLaunch = sAppData & "\Microsoft\Internet Explorer\QuickLaunch"
    
    sAllUsersStartMenu = RegGetString(HKEY_LOCAL_MACHINE, sShFolders, "Common Start Menu")
    sAllUsersPrograms = RegGetString(HKEY_LOCAL_MACHINE, sShFolders, "Common Programs")
    sAllUsersAppData = RegGetString(HKEY_LOCAL_MACHINE, sShFolders, "Common AppData")
    sAllUsersFavorites = RegGetString(HKEY_LOCAL_MACHINE, sShFolders, "Common Favorites")
    sAllUsersDesktop = RegGetString(HKEY_LOCAL_MACHINE, sShFolders, "Common Desktop")
    If sAllUsersDesktop = vbNullString Then sAllUsersDesktop = "%ALLUSERSDESKTOP"
    sAllUsersStartup = RegGetString(HKEY_LOCAL_MACHINE, sShFolders, "Common Startup")
    If sAllUsersStartup = vbNullString Then sAllUsersStartup = "%ALLUSERSSTARTUP%"
    
    If sAllUsersStartMenu = vbNullString And sAllUsersPrograms <> vbNullString Then
        'windows 9x may not have 'Common Start Menu' set in Registry
        sAllUsersStartMenu = Left(sAllUsersPrograms, InStrRev(sAllUsersPrograms, "\") - 1)
    End If
    
    sDummy = Replace(sDummy, "%systemdrive%", Left(sWinDir, 2), , , vbTextCompare)
    sDummy = Replace(sDummy, "%windir%", sWinDir, , , vbTextCompare)
    sDummy = Replace(sDummy, "%sysdir%", sSysDir, , , vbTextCompare)
    sDummy = Replace(sDummy, "%tempdir%", sTempDir, , , vbTextCompare)
    sDummy = Replace(sDummy, "%programfiles%", sProgramFiles, , , vbTextCompare)
    sDummy = Replace(sDummy, "%desktop%", sDesktop, , , vbTextCompare)
    sDummy = Replace(sDummy, "%mydocuments%", sMyDocuments, , , vbTextCompare)
    sDummy = Replace(sDummy, "%startmenu%", sStartMenu, , , vbTextCompare)
    sDummy = Replace(sDummy, "%startup%", sStartup, , , vbTextCompare)
    sDummy = Replace(sDummy, "%programs%", sPrograms, , , vbTextCompare)
    sDummy = Replace(sDummy, "%appdata%", sAppData, , , vbTextCompare)
    sDummy = Replace(sDummy, "%favorites%", sFavorites, , , vbTextCompare)
    sDummy = Replace(sDummy, "%allusersdesktop%", sAllUsersDesktop, , , vbTextCompare)
    sDummy = Replace(sDummy, "%allusersstartmenu%", sAllUsersStartMenu, , , vbTextCompare)
    sDummy = Replace(sDummy, "%allusersstartup%", sAllUsersStartup, , , vbTextCompare)
    sDummy = Replace(sDummy, "%allusersprograms%", sAllUsersPrograms, , , vbTextCompare)
    sDummy = Replace(sDummy, "%allusersappdata%", sAllUsersAppData, , , vbTextCompare)
    sDummy = Replace(sDummy, "%allusersfavorites%", sAllUsersFavorites, , , vbTextCompare)
    sDummy = Replace(sDummy, "%quicklaunch%", sQuickLaunch, , , vbTextCompare)
    sDummy = Replace(sDummy, "%username%", sUsername, , , vbTextCompare)
    sDummy = Replace(sDummy, "%computername%", sComputerName, , , vbTextCompare)
    
    If InStr(sDummy, "%") > 0 And InStr(sDummy, "%%") = 0 Then
        Logg "Warning: The following line has unexpanded aliases and will be skipped: " & sDummy
    End If
    ExpandEnvironmentVars = sDummy
End Function

Public Function GetScriptCRC32$(sFile$)
    modCRC32.Init
    GetScriptCRC32 = modCRC32.GetFileCRC32(sFile)
End Function

Public Function GetDOSFilename$(sFile$)
    Dim sBuf$
    sBuf = String(260, 0)
    GetShortPathName sFile, sBuf, Len(sBuf)
    GetDOSFilename = TrimNull(sBuf)
End Function

Public Function GetFileMD5$(sFile$)
    'note: this needs at least Win95 /w OSR2 (IE3)
    If Not FileExists(sFile) Then Exit Function
    Dim lFileLen&, sFileContents$
    On Error Resume Next
    lFileLen = FileLen(sFile)
    If lFileLen = 0 Then
        GetFileMD5 = "D41D8CD98F00B204E9800998ECF8427E"
        Exit Function
    End If
    sFileContents = InputFile(sFile)
    If Err Then Exit Function
    
    Dim hCrypt&, hHash&, uMD5(255) As Byte, lMD5Len&, i%, sMD5$
    If CryptAcquireContext(hCrypt, vbNullString, MS_ENHANCED_PROV, PROV_RSA_FULL, CRYPT_VERIFYCONTEXT) <> 0 Then
        If CryptCreateHash(hCrypt, ALG_TYPE_ANY Or ALG_CLASS_HASH Or ALG_SID_MD5, 0, 0, hHash) <> 0 Then
            If CryptHashData(hHash, sFileContents, Len(sFileContents), 0) <> 0 Then
                If CryptGetHashParam(hHash, HP_HASHSIZE, uMD5(0), UBound(uMD5) + 1, 0) <> 0 Then
                    lMD5Len = uMD5(0)
                    If CryptGetHashParam(hHash, HP_HASHVAL, uMD5(0), UBound(uMD5) + 1, 0) <> 0 Then
                        For i = 0 To lMD5Len - 1
                            sMD5 = sMD5 & Right("0" & Hex(uMD5(i)), 2)
                        Next i
                    End If
                End If
            End If
            CryptDestroyHash hHash
        End If
        CryptReleaseContext hCrypt, 0
    Else
        Logg "Failed: GetFileMD5 " & sFile & " (not supported)"
        Exit Function
    End If
    
    If sMD5 = vbNullString Then
        Logg "Failed: GetFileMD5 " & sFile
        Exit Function
    Else
        GetFileMD5 = sMD5
    End If
End Function

Public Function GetFileSHA1$(sFile$)
    'note: this needs at least Win95 /w OSR2 (IE3)
    If Not FileExists(sFile) Then Exit Function
    Dim lFileLen&, sFileContents$
    On Error Resume Next
    lFileLen = FileLen(sFile)
    If lFileLen = 0 Then
        GetFileSHA1 = "DA39A3EE5E6B4B0D3255BFEF95601890AFD80709"
        Exit Function
    End If
    sFileContents = InputFile(sFile)
    If Err Then Exit Function
    
    Dim hCrypt&, hHash&, uSHA1(255) As Byte, lSHA1Len&, i%, sSHA1$
    If CryptAcquireContext(hCrypt, vbNullString, MS_ENHANCED_PROV, PROV_RSA_FULL, CRYPT_VERIFYCONTEXT) <> 0 Then
        If CryptCreateHash(hCrypt, ALG_TYPE_ANY Or ALG_CLASS_HASH Or ALG_SID_SHA1, 0, 0, hHash) <> 0 Then
            If CryptHashData(hHash, sFileContents, Len(sFileContents), 0) <> 0 Then
                If CryptGetHashParam(hHash, HP_HASHSIZE, uSHA1(0), UBound(uSHA1) + 1, 0) <> 0 Then
                    lSHA1Len = uSHA1(0)
                    If CryptGetHashParam(hHash, HP_HASHVAL, uSHA1(0), UBound(uSHA1) + 1, 0) <> 0 Then
                        For i = 0 To lSHA1Len - 1
                            sSHA1 = sSHA1 & Right("0" & Hex(uSHA1(i)), 2)
                        Next i
                    End If
                End If
            End If
            CryptDestroyHash hHash
        End If
        CryptReleaseContext hCrypt, 0
    Else
        Logg "Failed: GetFileSHA1 " & sFile & " (not supported)"
        Exit Function
    End If
    
    If sSHA1 = vbNullString Then
        Logg "Failed: GetFileSHA1 " & sFile
        Exit Function
    Else
        GetFileSHA1 = sSHA1
    End If
End Function

Public Function GetFileMD2$(sFile$)
    'note: this needs at least Win95 /w OSR2 (IE3)
    If Not FileExists(sFile) Then Exit Function
    Dim lFileLen&, sFileContents$
    On Error Resume Next
    lFileLen = FileLen(sFile)
    If lFileLen = 0 Then
        GetFileMD2 = "8350E5A3E24C153DF2275C9F80692773"
        Exit Function
    End If
    sFileContents = InputFile(sFile)
    If Err Then Exit Function
    
    Dim hCrypt&, hHash&, uMD2(255) As Byte, lMD2Len&, i%, sMD2$
    If CryptAcquireContext(hCrypt, vbNullString, MS_ENHANCED_PROV, PROV_RSA_FULL, CRYPT_VERIFYCONTEXT) <> 0 Then
        If CryptCreateHash(hCrypt, ALG_TYPE_ANY Or ALG_CLASS_HASH Or ALG_SID_MD2, 0, 0, hHash) <> 0 Then
            If CryptHashData(hHash, sFileContents, Len(sFileContents), 0) <> 0 Then
                If CryptGetHashParam(hHash, HP_HASHSIZE, uMD2(0), UBound(uMD2) + 1, 0) <> 0 Then
                    lMD2Len = uMD2(0)
                    If CryptGetHashParam(hHash, HP_HASHVAL, uMD2(0), UBound(uMD2) + 1, 0) <> 0 Then
                        For i = 0 To lMD2Len - 1
                            sMD2 = sMD2 & Right("0" & Hex(uMD2(i)), 2)
                        Next i
                    End If
                End If
            End If
            CryptDestroyHash hHash
        End If
        CryptReleaseContext hCrypt, 0
    Else
        Logg "Failed: GetFileMD2 " & sFile & " (not supported)"
        Exit Function
    End If
    
    If sMD2 = vbNullString Then
        Logg "Failed: GetFileMD2 " & sFile
        Exit Function
    Else
        GetFileMD2 = sMD2
    End If
End Function

Public Function GetFileMD4$(sFile$)
    'note: this needs at least Win95 /w OSR2 (IE3)
    If Not FileExists(sFile) Then Exit Function
    Dim lFileLen&, sFileContents$
    On Error Resume Next
    lFileLen = FileLen(sFile)
    If lFileLen = 0 Then
        GetFileMD4 = "31D6CFE0D16AE931B73C59D7E0C089C0"
        Exit Function
    End If
    sFileContents = InputFile(sFile)
    If Err Then Exit Function
    
    Dim hCrypt&, hHash&, uMD4(255) As Byte, lMD4Len&, i%, sMD4$
    If CryptAcquireContext(hCrypt, vbNullString, MS_ENHANCED_PROV, PROV_RSA_FULL, CRYPT_VERIFYCONTEXT) <> 0 Then
        If CryptCreateHash(hCrypt, ALG_TYPE_ANY Or ALG_CLASS_HASH Or ALG_SID_MD4, 0, 0, hHash) <> 0 Then
            If CryptHashData(hHash, sFileContents, Len(sFileContents), 0) <> 0 Then
                If CryptGetHashParam(hHash, HP_HASHSIZE, uMD4(0), UBound(uMD4) + 1, 0) <> 0 Then
                    lMD4Len = uMD4(0)
                    If CryptGetHashParam(hHash, HP_HASHVAL, uMD4(0), UBound(uMD4) + 1, 0) <> 0 Then
                        For i = 0 To lMD4Len - 1
                            sMD4 = sMD4 & Right("0" & Hex(uMD4(i)), 2)
                        Next i
                    End If
                End If
            End If
            CryptDestroyHash hHash
        End If
        CryptReleaseContext hCrypt, 0
    Else
        Logg "Failed: GetFileMD4 " & sFile & " (not supported)"
        Exit Function
    End If
    
    If sMD4 = vbNullString Then
        Logg "Failed: GetFileMD4 " & sFile
        Exit Function
    Else
        GetFileMD4 = sMD4
    End If
End Function

Public Function InputFile$(sFile$)
    'internal function (multiple modules)
    'this uses APIs instead of Input(), which is ~3x slower and doesn't cache :P
    Dim hFile&, uBuffer() As Byte, lFileSize&, lBytesRead&
    hFile = CreateFile(sFile, GENERIC_READ, FILE_SHARE_READ Or FILE_SHARE_WRITE, ByVal 0, OPEN_EXISTING, 0, 0)
    If hFile = -1 Then Exit Function
    
    'second parameter is dwSizeHigh, we ignore that
    lFileSize = GetFileSize(hFile, 0)
    If lFileSize = -1 Or lFileSize = 0 Then
        CloseHandle hFile
        Exit Function
    End If
    
    ReDim uBuffer(lFileSize - 1)
    If ReadFile(hFile, uBuffer(0), lFileSize, lBytesRead, ByVal 0) <> 0 Then
        If lBytesRead <> lFileSize Then
            'buffer was too large
            ReDim Preserve uBuffer(lBytesRead)
        End If
        InputFile = StrConv(uBuffer, vbUnicode)
    End If
    CloseHandle hFile
End Function

Public Sub OutputFile(sFile$, sContent$, Optional bAppend As Boolean = False)
    'internal function (multiple modules)
    Dim hFile&, lFileSize&, uBuffer() As Byte
    If Trim(sContent) = vbNullString Then Exit Sub
    If Not bAppend Then
        'overwrite the file and open it
        hFile = CreateFile(sFile, GENERIC_WRITE, FILE_SHARE_READ Or FILE_SHARE_WRITE, ByVal 0, CREATE_ALWAYS, 0, 0)
    Else
        'open the file
        hFile = CreateFile(sFile, GENERIC_READ Or GENERIC_WRITE, FILE_SHARE_READ Or FILE_SHARE_WRITE, ByVal 0, OPEN_EXISTING, 0, 0)
    End If
    If hFile = -1 Then Exit Sub
    
    uBuffer = StrConv(sContent, vbFromUnicode)
    
    If bAppend Then
        lFileSize = GetFileSize(hFile, 0)
        If lFileSize = -1 Then
            CloseHandle hFile
            Exit Sub
        End If
        
        If SetFilePointer(hFile, lFileSize, 0, FILE_BEGIN) = -1 Then
            CloseHandle hFile
            Exit Sub
        End If
    End If
    
    WriteFile hFile, uBuffer(0), UBound(uBuffer) + 1, 0, ByVal 0
    DoEvents
    CloseHandle hFile
End Sub

Public Function InputURL$(sURL$)
    Dim hInternet&, hFile&, sFile$, sBuffer$, lBytesRead&
    Dim sUserAgent$
    sUserAgent = "BFU v" & App.Major & "." & Format(App.Minor, "00") & "." & App.Revision
    hInternet = InternetOpen(sUserAgent, INTERNET_OPEN_TYPE_DIRECT, vbNullString, vbNullString, 0)
    If hInternet <> 0 Then
        hFile = InternetOpenUrl(hInternet, sURL, vbNullString, 0, INTERNET_FLAG_RELOAD, 0)
        If hFile <> 0 Then
            Do
                sBuffer = Space(32768)
                lBytesRead = 0
                InternetReadFile hFile, sBuffer, Len(sBuffer), lBytesRead
                sFile = sFile & Left(sBuffer, lBytesRead)
            Loop Until lBytesRead = 0
            InputURL = sFile
            InternetCloseHandle hFile
        End If
        InternetCloseHandle hInternet
    End If
End Function

Private Function GetFriendlyWinVer$(bNT As Boolean, lMajor&, lMinor&, lBuild&, sCSD$)
    Dim sFriendly$ ', l9xBuild&
    If bNT Then
        Select Case lMajor
            Case 3 'WinNT 3
                sFriendly = "Windows NT 3." & CStr(lMinor) & " " & sCSD
            Case 4 'WinNT 4
                sFriendly = "Windows NT4 " & sCSD
            Case 5 'Win2000/XP/2003
                Select Case lMinor
                    Case 0 'Win2000
                        sFriendly = "Windows 2000 " & sCSD
                    Case 1 'WinXP
                        sFriendly = "Windows XP " & sCSD
                    Case 2 'Win2003
                        If IsWorkStation Then
                            'sFriendly = "Windows XP 64bit " & sCS
                            sFriendly = "Windows 2003 SBS " & sCSD
                        Else
                            sFriendly = "Windows 2003 " & sCSD
                        End If
                End Select
            Case 6 'WinVista
                sFriendly = "Windows Vista " & sCSD
        End Select
    Else
        'l9xBuild = (lBuild And &HFFF)
        Select Case lMajor
            Case 4 'Win95/98/ME
                Select Case lMinor
                    Case 0 'Win95/OSR2
                        If sCSD = "B" Or sCSD = "C" Then
                            'Win95 OSR2
                            sFriendly = "Windows 95 OSR2"
                        Else
                            'Win95
                            sFriendly = "Windows 95"
                        End If
                    Case 10 'Win98/98SE
                        If sCSD = "A" Then
                            'Win98SE
                            sFriendly = "Windows 98 SE"
                        Else
                            'Win98
                            sFriendly = "Windows 98"
                        End If
                    Case 90 'WinME
                        sFriendly = "Windows ME"
                End Select
        End Select
    End If
    
    If sFriendly = vbNullString Then
        GetFriendlyWinVer = "Unknown Windows version"
    Else
        GetFriendlyWinVer = RTrim(sFriendly)
    End If
End Function

Private Function IsWorkStation() As Boolean
    Dim uOVI2 As OSVERSIONINFOEX
    On Error Resume Next
    With uOVI2
        .dwOSVersionInfoSize = Len(uOVI2)
        GetVersionEx uOVI2
        If (.wProductType And VER_NT_WORKSTATION) Then
            IsWorkStation = True
        End If
    End With
End Function

Public Function BuildPath$(sPath$, sFile$)
    If Right(sPath, 1) = "\" Then
        BuildPath = sPath & sFile
    Else
        BuildPath = sPath & "\" & sFile
    End If
End Function

Public Sub AddToArray(ByRef uArray As Variant, sItem$)
    On Error Resume Next
    ReDim Preserve uArray(UBound(uArray) + 1)
    uArray(UBound(uArray)) = sItem
End Sub
