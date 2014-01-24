Attribute VB_Name = "modSystem"
Option Explicit
'system functions: run, restart, msgbox, dll reg/unreg

Private Declare Function SHRestartSystemMB Lib "shell32" Alias "#59" (ByVal hOwner As Long, ByVal sExtraPrompt As String, ByVal uFlags As Long) As Long
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Declare Function ShellExecuteEx Lib "shell32.dll" (SEI As SHELLEXECUTEINFO) As Long
Private Declare Function SHEmptyRecycleBin Lib "shell32.dll" Alias "SHEmptyRecycleBinA" (ByVal hWnd As Long, ByVal pszRootPath As String, ByVal dwFlags As Long) As Long
Private Declare Function SHUpdateRecycleBinIcon Lib "shell32.dll" () As Long

Private Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function FindNextFile Lib "kernel32" Alias "FindNextFileA" (ByVal hFindFile As Long, lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long

Private Type FILETIME
    dwLowDateTime As Long
    dwHighDateTime As Long
End Type

Private Type WIN32_FIND_DATA
    dwFileAttributes As Long
    ftCreationTime As FILETIME
    ftLastAccessTime As FILETIME
    ftLastWriteTime As FILETIME
    nFileSizeHigh As Long
    nFileSizeLow As Long
    dwReserved0 As Long
    dwReserved1 As Long
    cFileName As String * 260
    cAlternate As String * 14
End Type

Private Type SHELLEXECUTEINFO
    cbSize As Long
    fMask As Long
    hWnd As Long
    lpVerb As String
    lpFile As String
    lpParameters As String
    lpDirectory As String
    nShow As Long
    hInstApp As Long
    lpIDList As Long
    lpClass As String
    hkeyClass As Long
    dwHotKey As Long
    hIcon As Long
    hProcess As Long
End Type

Private Const SEE_MASK_INVOKEIDLIST = &HC
Private Const SEE_MASK_NOCLOSEPROCESS = &H40
Private Const SEE_MASK_FLAG_NO_UI = &H400

Private Const SW_HIDE = 0
Private Const SW_SHOWNORMAL = 1

Private Const EWX_REBOOT = 2
Private Const EWX_FORCE = 4

Private Const SHERB_NOCONFIRMATION = &H1
Private Const SHERB_NOPROGRESSUI = &H2
Private Const SHERB_NOSOUND = &H4

Private Const FILE_ATTRIBUTE_DIRECTORY = &H10

Public Sub SystemRestart(sCmd$)
    'SystemRestart [extra msg]|[force(0|1)]
    Dim sMsg$, bForce As Boolean
    sMsg = sCmd
    If sMsg <> vbNullString Then
        If InStr(sMsg, "|") > 0 Then
            bForce = CBool(Val(Right(sMsg, 1)))
            sMsg = Left(sMsg, InStr(sMsg, "|") - 1)
        End If
        sMsg = Replace(sMsg, "\n", vbCrLf)
    End If
    
    If bIsWinNT Then
        SHRestartSystemMB frmMain.hWnd, StrConv(sMsg & IIf(sMsg <> vbNullString, vbCrLf & vbCrLf, vbNullString), vbUnicode), EWX_REBOOT + IIf(bForce, EWX_FORCE, 0)
    Else
        SHRestartSystemMB frmMain.hWnd, sMsg & IIf(sMsg <> vbNullString, vbCrLf & vbCrLf, vbNullString), EWX_REBOOT + IIf(bForce, EWX_FORCE, 0)
    End If
End Sub

Public Sub SystemRestartIfNeeded(sCmd$)
    'same as SystemRestart but only if a reboot is actually needed
    If bRebootNeeded Then SystemRestart sCmd
End Sub

Public Sub SystemMsgBox(sCmd$)
    'SystemMsgBox <text>
    Dim sMsg$
    If sCmd = vbNullString Then Exit Sub
    sMsg = Replace(sCmd, "\n", vbCrLf)
    If sMsg <> vbNullString And Not bRunSilent Then MsgBox sMsg, vbInformation
End Sub

Public Sub DllRegister(sCmd$)
    'DllRegister <dll file>|[silent(0|1)]
    Dim sDll$, bSilent As Boolean
    sDll = sCmd
    If sDll <> vbNullString Then
        If InStr(sDll, "|") > 0 Then
            bSilent = CBool(Val(Mid(sDll, InStr(sDll, "|") + 1)))
            sDll = Left(sDll, InStr(sDll, "|") - 1)
        Else
            bSilent = True
        End If
        
        If FileExists(sDll) Then
            If ShellExecute(frmMain.hWnd, "open", sSysDir & "\regsvr32.exe " & IIf(bSilent, "/s ", vbNullString) & sDll, vbNullString, vbNullString, SW_SHOWNORMAL) <= 32 Then
                Logg "Failed: DllRegister " & sCmd & " (operation failed)"
            Else
                Logg "Success: DllRegister " & sCmd
            End If
        Else
            Logg "Failed: DllRegister " & sCmd & " (file not found)"
        End If
    End If
End Sub

Public Sub DllUnregister(sCmd$)
    'DllUnregister <dll file>|[silent(0|1)]
    Dim sDll$, bSilent As Boolean
    sDll = sCmd
    bSilent = True
    If sDll <> vbNullString Then
        If InStr(sDll, "|") > 0 Then
            sDll = Left(sDll, InStr(sDll, "|") - 1)
            bSilent = CBool(Val(Right(sDll, 1)))
        Else
            bSilent = True
        End If
        
        If FileExists(sDll) Then
            If ShellExecute(frmMain.hWnd, "open", sSysDir & "\regsvr32.exe /u " & IIf(bSilent, "/s ", vbNullString) & sDll, vbNullString, vbNullString, SW_SHOWNORMAL) <= 32 Then
                Logg "Failed: DllUnregister " & sCmd & " (operation failed)"
            Else
                Logg "Success: DllUnregister " & sCmd
            End If
        Else
            Logg "Failed: DllUnregister " & sCmd & " (file not found)"
        End If
    End If
End Sub

Public Sub SystemRun(sCmd$)
    'SystemRun <file>|[parameters]|[show(0|1)]
    Dim sFile$, sParams$, bShow As Boolean
    sFile = sCmd
    If sFile = vbNullString Then Exit Sub
    bShow = True
    If InStr(sFile, "|") > 0 Then
        sParams = Mid(sFile, InStr(sFile, "|") + 1)
        sFile = Left(sFile, InStr(sFile, "|") - 1)
        If InStr(sParams, "|") > 0 Then
            bShow = CBool(Val(Mid(sParams, InStr(sParams, "|") + 1)))
            sParams = Left(sParams, InStr(sParams, "|") - 1)
        End If
    End If
    If ShellExecute(frmMain.hWnd, "open", sFile, sParams, vbNullString, IIf(bShow, SW_SHOWNORMAL, SW_HIDE)) <= 32 Then
        Logg "Failed: SystemRun " & sCmd & " (operation failed)"
    Else
        Logg "Success: SystemRun " & sCmd
    End If
End Sub

Public Sub SystemResetWebSettings()
    'SystemResetWebSettings
    'copied from IERESET.INF
    Dim sIE$, sIS$
    sIE = "Software\Microsoft\Internet Explorer"
    sIS = "Software\Microsoft\Windows\CurrentVersion\Internet Settings"
    RegSetString HKEY_CURRENT_USER, sIE & "\Main", "Start Page", "http://www.microsoft.com/isapi/redir.dll?prd=ie&pver=6&ar=msnhome"
    RegSetString HKEY_LOCAL_MACHINE, sIE & "\Main", "Default_Page_URL", "http://www.microsoft.com/isapi/redir.dll?prd=ie&pver=6&ar=msnhome"
    RegSetString HKEY_LOCAL_MACHINE, sIE & "\Main", "Default_Search_URL", "http://www.microsoft.com/isapi/redir.dll?prd=ie&ar=iesearch"
    RegSetString HKEY_LOCAL_MACHINE, sIE & "\Main", "Search Page", "http://www.microsoft.com/isapi/redir.dll?prd=ie&ar=iesearch"
    RegSetString HKEY_CURRENT_USER, sIE & "\Main", "Search Page", "http://www.microsoft.com/isapi/redir.dll?prd=ie&ar=iesearch"
    
    RegSetString HKEY_LOCAL_MACHINE, sIE & "\Main\UrlTemplate", "1", "www.%s.com"
    RegSetString HKEY_LOCAL_MACHINE, sIE & "\Main\UrlTemplate", "2", "www.%s.org"
    RegSetString HKEY_LOCAL_MACHINE, sIE & "\Main\UrlTemplate", "3", "www.%s.net"
    RegSetString HKEY_LOCAL_MACHINE, sIE & "\Main\UrlTemplate", "4", "www.%s.edu"
    RegDelValue "HKLM\" & sIE & "\Main\UrlTemplate|5"
    RegDelValue "HKLM\" & sIE & "\Main\UrlTemplate|6"
    RegDelValue "HKLM\" & sIE & "\Main\UrlTemplate|7"
    RegDelValue "HKLM\" & sIE & "\Main\UrlTemplate|8"
    RegDelValue "HKLM\" & sIE & "\Main\UrlTemplate|9"
    
    RegSetString HKEY_CURRENT_USER, sIE & "\SearchUrl", "Provider", ""
    
    RegSetString HKEY_LOCAL_MACHINE, sIE & "\Search", "SearchAssistant", "http://ie.search.msn.com/{SUB_RFC1766}/srchasst/srchasst.htm"
    RegSetString HKEY_LOCAL_MACHINE, sIE & "\Search", "CustomizeSearch", "http://ie.search.msn.com/{SUB_RFC1766}/srchasst/srchcust.htm"
    
    RegSetString HKEY_LOCAL_MACHINE, sIS & "\SafeSites", "ie.search.msn.com", "http://ie.search.msn.com/*"
End Sub

Public Sub SystemEmptyRecycleBin()
    'empties the Recycle Bin (on all drives)
    'might not work - needs WinNT4+IE4 and up, Win95+IE5 and up
    On Error Resume Next
    If SHEmptyRecycleBin(0, 0, SHERB_NOCONFIRMATION Or SHERB_NOPROGRESSUI Or SHERB_NOSOUND) = 0 Then
        Logg "Failed: SystemEmptyRecycleBin (operation failed)"
    End If
    SHUpdateRecycleBinIcon
    If Err Then
        If Err.Number = 453 Then
            Logg "Failed: SystemEmptyRecycleBin (not supported)"
        Else
            Logg "Failed: SystemEmptyRecycleBin (operation failed)"
        End If
    Else
        Logg "Success: SystemEmptyRecycleBin"
    End If
End Sub

Public Sub SystemEmptyInternetCache()
    Dim sCacheFolder$, hFind&, uWFD As WIN32_FIND_DATA
    Dim sFolder$, sList$, i%, sSubfolders$()
    sCacheFolder = RegGetString(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\explorer\Shell Folders", "Cache")
    If sCacheFolder = vbNullString Then Exit Sub
    'this probably doesn't work for IE4 and lower
    hFind = FindFirstFile(sCacheFolder & "\Content.IE5\*.", uWFD)
    If hFind = 0 Then
        Logg "Failed: SystemEmptyInternetCache (not supported)"
        Exit Sub
    End If
    Do
        If (FILE_ATTRIBUTE_DIRECTORY And uWFD.dwFileAttributes) Then
            sFolder = TrimNull(uWFD.cFileName)
            If sFolder <> "." And sFolder <> ".." Then sList = sList & sFolder & "|"
        End If
    Loop Until FindNextFile(hFind, uWFD) = 0
    FindClose hFind
    If sList = vbNullString Then Exit Sub
    sSubfolders = Split(Left(sList, Len(sList) - 1), "|")
    For i = 0 To UBound(sSubfolders)
        FolderDelete sCacheFolder & "\Content.IE5\" & sSubfolders(i)
    Next i
End Sub

Public Sub SystemEmptyTempFolder()
    'empties the current user's temp folder
    'as well as $WINDIR\temp (if different)
    Dim hFind&, uWFD As WIN32_FIND_DATA, sFile$
    hFind = FindFirstFile(sTempDir & "\*.*", uWFD)
    If hFind <> 0 Then 'found something, so continue
        Do
            sFile = TrimNull(uWFD.cFileName)
            If (FILE_ATTRIBUTE_DIRECTORY And uWFD.dwFileAttributes) Then
                If sFile <> "." And sFile <> ".." Then
                    FolderDelete sTempDir & "\" & sFile
                End If
            Else
                FileDelete sTempDir & "\" & sFile
                DoEvents
            End If
        Loop Until FindNextFile(hFind, uWFD) = 0
        FindClose hFind
    End If
    
    If Not bIsWinNT Then Exit Sub
    hFind = FindFirstFile(sWinDir & "\Temp\*.*", uWFD)
    If hFind = 0 Then Exit Sub 'nothing found
    Do
        sFile = TrimNull(uWFD.cFileName)
        If (FILE_ATTRIBUTE_DIRECTORY And uWFD.dwFileAttributes) Then
            If sFile <> "." And sFile <> ".." Then
                FolderDelete sWinDir & "\Temp\" & sFile
            End If
        Else
            FileDelete sWinDir & "\Temp\" & sFile
        End If
    Loop Until FindNextFile(hFind, uWFD) = 0
    FindClose hFind
End Sub

