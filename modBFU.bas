Attribute VB_Name = "modBFU"
Option Explicit
'wrapper for uninstall functions

Private bUnloadShell As Boolean
Public bUseRecycleBin As Boolean

Public bUseDeleteOnReboot As Boolean
Public bAbortScript As Boolean
Public bRebootNeeded As Boolean
Private bStatusMsgs As Boolean

Public bRunSilent As Boolean
Public sSaveLogFile As String

Private lPause&
Public sLog$
Private sCommandsList$()

Public Sub LoadCommandsList()
    ReDim sCommandsList(0)
    AddToArray sCommandsList, "OptionUnloadShell"
    AddToArray sCommandsList, "OptionUseRecycleBin"
    AddToArray sCommandsList, "OptionBFUMinVersion"
    AddToArray sCommandsList, "OptionPauseBetweenCmds"
    AddToArray sCommandsList, "OptionPauseNow"
    AddToArray sCommandsList, "OptionCalcScriptCRC32"
    AddToArray sCommandsList, "OptionSetStatus"
    AddToArray sCommandsList, "OptionOnDeleteFailUseReboot"
    AddToArray sCommandsList, "OptionSetBFURunOnReboot"
    AddToArray sCommandsList, "OptionBFUExit"
    AddToArray sCommandsList, "OptionShowLog"
    AddToArray sCommandsList, "OptionSaveLog"

    AddToArray sCommandsList, "FileCreate"
    AddToArray sCommandsList, "FileDelete "
    AddToArray sCommandsList, "FileDeleteOnReboot"
    AddToArray sCommandsList, "FileRename"
    AddToArray sCommandsList, "FileMove"
    AddToArray sCommandsList, "FileClear"
    AddToArray sCommandsList, "FileSetAttributes"
    AddToArray sCommandsList, "FolderCreate"
    AddToArray sCommandsList, "FolderRename"
    AddToArray sCommandsList, "FolderMove"
    AddToArray sCommandsList, "FolderDelete"
    AddToArray sCommandsList, "FolderSetAttributes"
    AddToArray sCommandsList, "FileDeleteIfMD5Match"
    AddToArray sCommandsList, "FileDeleteIfCRC32Match"
    AddToArray sCommandsList, "FileDeleteIfSHA1Match"
    AddToArray sCommandsList, "FileDeleteIfMD2Match"
    AddToArray sCommandsList, "FileDeleteIfMD4Match"
    AddToArray sCommandsList, "FileDeleteIfContainsText"
    AddToArray sCommandsList, "FileDeleteIfContainsHex"
    AddToArray sCommandsList, "FileMoveIfMD5Match"
    AddToArray sCommandsList, "FileMoveIfCRC32Match"
    AddToArray sCommandsList, "FileMoveIfSHA1Match"
    AddToArray sCommandsList, "FileMoveIfMD2Match"
    AddToArray sCommandsList, "FileMoveIfMD4Match"
    AddToArray sCommandsList, "FileMoveIfContainsText"
    AddToArray sCommandsList, "FileMoveIfContainsHex"
    AddToArray sCommandsList, "FolderClear"
    AddToArray sCommandsList, "FileWrite"

    AddToArray sCommandsList, "IniSetValue"
    AddToArray sCommandsList, "IniDeleteValue"
    AddToArray sCommandsList, "IniDeleteFromValue"
    AddToArray sCommandsList, "IniClearValue"
    AddToArray sCommandsList, "IniCreateSection"

    AddToArray sCommandsList, "HostsFileReset"
    AddToArray sCommandsList, "HostsFileAddLine"
    AddToArray sCommandsList, "HostsFileDelLine"
    AddToArray sCommandsList, "HostsFileDisableLine"
    AddToArray sCommandsList, "HostsFileEnableLine"

    AddToArray sCommandsList, "RegCreateKey"
    AddToArray sCommandsList, "RegDeleteKey"
    AddToArray sCommandsList, "RegDeleteKeyIfNameContainsText"
    AddToArray sCommandsList, "RegDeleteKeyIfNameContainsHex"
    AddToArray sCommandsList, "RegSetStringValue"
    AddToArray sCommandsList, "RegSetDwordValue"
    AddToArray sCommandsList, "RegSetBinaryValue"
    AddToArray sCommandsList, "RegSetExpandValue"
    AddToArray sCommandsList, "RegDelValue"
    AddToArray sCommandsList, "RegDelFromValue"
    AddToArray sCommandsList, "RegRenameValue"
    AddToArray sCommandsList, "RegDelValueIfDataContainsText"
    AddToArray sCommandsList, "RegDelValueIfDataContainsHex"
    AddToArray sCommandsList, "RegDelValueIfNameContainsText"
    AddToArray sCommandsList, "RegDelValueIfNameContainsHex"
    AddToArray sCommandsList, "RegResetPermissions"
    AddToArray sCommandsList, "RegSetMultiValue"

    AddToArray sCommandsList, "ProcessKill"
    AddToArray sCommandsList, "ProcessKillIfMD5Match"
    AddToArray sCommandsList, "ProcessKillIfCRC32Match"
    AddToArray sCommandsList, "ProcessKillIfSHA1Match"
    AddToArray sCommandsList, "ProcessKillIfMD2Match"
    AddToArray sCommandsList, "ProcessKillIfMD4Match"
    AddToArray sCommandsList, "ProcessKillIfContainsText"
    AddToArray sCommandsList, "ProcessKillIfContainsHex"
    AddToArray sCommandsList, "ProcessSuspend"
    AddToArray sCommandsList, "ProcessSuspendIfMD5Match"
    AddToArray sCommandsList, "ProcessSuspendIfCRC32Match"
    AddToArray sCommandsList, "ProcessSuspendIfSHA1Match"
    AddToArray sCommandsList, "ProcessSuspendIfMD2Match"
    AddToArray sCommandsList, "ProcessSuspendIfMD4Match"
    AddToArray sCommandsList, "ProcessSuspendIfContainsText"
    AddToArray sCommandsList, "ProcessSuspendIfContainsHex"
    AddToArray sCommandsList, "ProcessResume"
    AddToArray sCommandsList, "ProcessResumeIfMD5Match"
    AddToArray sCommandsList, "ProcessResumeIfCRC32Match"
    AddToArray sCommandsList, "ProcessResumeIfSHA1Match"
    AddToArray sCommandsList, "ProcessResumeIfMD2Match"
    AddToArray sCommandsList, "ProcessResumeIfMD4Match"
    AddToArray sCommandsList, "ProcessResumeIfContainsText"
    AddToArray sCommandsList, "ProcessResumeIfContainsHex"
    
    AddToArray sCommandsList, "ServiceStart"
    AddToArray sCommandsList, "ServiceStop"
    AddToArray sCommandsList, "ServiceDisable"
    AddToArray sCommandsList, "ServiceEnable"
    AddToArray sCommandsList, "ServiceDelete"

    AddToArray sCommandsList, "WinsockKillProtocol"
    AddToArray sCommandsList, "WinsockKillNameSpace"

    AddToArray sCommandsList, "DllRegister"
    AddToArray sCommandsList, "DllUnregister"

    AddToArray sCommandsList, "SystemRestart"
    AddToArray sCommandsList, "SystemRestartIfNeeded"
    AddToArray sCommandsList, "SystemRun"
    AddToArray sCommandsList, "SystemMsgBox"
    AddToArray sCommandsList, "SystemResetWebSettings"
    AddToArray sCommandsList, "SystemEmptyRecycleBin"
    AddToArray sCommandsList, "SystemEmptyInternetCache"
    AddToArray sCommandsList, "SystemEmptyTempFolder"

    AddToArray sCommandsList, "LogIfFileExists"
    AddToArray sCommandsList, "LogIfFileContainsText"
    AddToArray sCommandsList, "LogIfFileContainsHex"
    AddToArray sCommandsList, "LogIfFolderExists"
    AddToArray sCommandsList, "LogIfRegKeyExists"
    AddToArray sCommandsList, "LogIfRegValExists"
    AddToArray sCommandsList, "LogIfRegValContainsText"
    AddToArray sCommandsList, "LogIfRegValContainsHex"
End Sub

Public Sub GetScriptOptions(sFullScript$)
    Dim i&, sScript$()
    'get only script options that user can change
    
    sScript = Split(sFullScript, vbCrLf)
    frmMain.chkUnloadShell.Value = 0
    bUnloadShell = False
    frmMain.chkUseRecycleBin.Value = 0
    bUseRecycleBin = False
    For i = 0 To UBound(sScript)
        'If InStr(sScript(i), "Option") <> 1 Then Exit For
        
        If InStr(sScript(i), "OptionUnloadShell") = 1 Then
            frmMain.chkUnloadShell.Value = 1
            bUnloadShell = True
        End If
        If InStr(sScript(i), "OptionUseRecycleBin") = 1 Then
            frmMain.chkUseRecycleBin.Value = 1
            bUseRecycleBin = True
        End If
    Next i
    
    'reset other options
    bUseDeleteOnReboot = False
    bStatusMsgs = False
    'bRebootNeeded = False
    bRunSilent = False
End Sub

Private Sub SetScriptOption(sCmd$)
    'anything else I can think of
    Dim i%
    
    If InStr(1, sCmd, "OptionBFUMinVersion", vbTextCompare) = 1 Then
        Dim lMinVersion&, lThisVersion&
        lMinVersion = CLng(LTrim(Mid(sCmd, Len("OptionBFUMinVersion") + 1)))
        Logg "Option BFU minimum version: " & lMinVersion
        lThisVersion = CLng(CStr(App.Major) & Format(App.Minor, "00") & Format(App.Revision, "0000"))
        If lMinVersion > lThisVersion Then
            MsgBox "Your version of BFU is too old to execute this script. " & _
                   "You need at least version " & Mid(CStr(lMinVersion), 1, 1) & "." & _
                   Mid(CStr(lMinVersion), 2, 2) & "." & _
                   Val(Mid(CStr(lMinVersion), 4, 4)) & ".", vbCritical
            Logg "BFU version too old, aborting."
            bAbortScript = True
        End If
    End If
    
    If InStr(1, sCmd, "OptionPauseBetweenCmds", vbTextCompare) = 1 Then
        lPause = CLng(LTrim(Mid(sCmd, Len("OptionPauseBetweenCmds") + 1)))
        Logg "Option pause between commands: " & lPause & " ms"
    End If
    
    If InStr(1, sCmd, "OptionPauseNow", vbTextCompare) = 1 Then
        Sleep CLng(LTrim(Mid(sCmd, Len("OptionPauseNow") + 1)))
        DoEvents
    End If
    
    If InStr(1, sCmd, "OptionCalcScriptCRC32", vbTextCompare) = 1 And Not bRunSilent Then
        Dim sFilename$, sCRC32$
        sFilename = frmMain.txtScript.Text
        sCRC32 = GetScriptCRC32(sFilename)
        sFilename = Mid(sFilename, InStrRev(sFilename, "\") + 1)
        If MsgBox("The CRC32 checksum of this script (" & _
                  sFilename & ") is: " & sCRC32 & vbCrLf & vbCrLf & _
                  "If this checksum is not the same as the one provided " & _
                  "with this script, the script may not have downloaded " & _
                  "correctly and could be damaged. Running a damaged " & _
                  "script could be harmful to your system." & vbCrLf & _
                  "Continue script execution?", vbExclamation + vbYesNo + vbDefaultButton2) = vbNo Then
            bAbortScript = True
        End If
    End If
    
    'If InStr(1, sCmd, "OptionStatusOn", vbTextCompare) = 1 Then bStatusMsgs = True
    If InStr(1, sCmd, "OptionSetStatus", vbTextCompare) = 1 Then
        bStatusMsgs = True
        Status Mid(sCmd, Len("OptionSetStatus") + 2)
    End If
    
    If InStr(1, sCmd, "OptionOnDeleteFailUseReboot", vbTextCompare) = 1 Then
        bUseDeleteOnReboot = True
    End If
    
    If InStr(1, sCmd, "OptionRunSilent", vbTextCompare) = 1 Then
        bRunSilent = True
        frmMain.Visible = False
    End If
    
    If InStr(1, sCmd, "OptionSetBFURunOnReboot", vbTextCompare) = 1 Then
        RegSetBFURunOnReboot sCmd
    End If
    
    If InStr(1, sCmd, "OptionBFUExit", vbTextCompare) = 1 Then
        Close
        End
    End If
    
    If InStr(1, sCmd, "OptionShowLog", vbTextCompare) = 1 Then
        frmMain.chkShowLog.Value = 1
    End If
    
    If InStr(1, sCmd, "OptionSaveLog", vbTextCompare) = 1 Then
        sSaveLogFile = Mid(sCmd, Len("OptionSaveLog") + 2)
    End If
End Sub

Public Sub ExecuteScript(sFullScript$)
    Dim i&, j&, sScript$(), bUnknownCmd As Boolean
    sLog = vbNullString
    Logg "BFU v" & App.Major & "." & Format(App.Minor, "00") & "." & App.Revision
    Logg sWinVer
    Logg "Script started at " & Time & ", on " & Format(Date, "Short Date") & vbCrLf
    If frmMain.chkUnloadShell.Value = 1 Then Logg "Option Unload Explorer: Yes"
    If frmMain.chkUseRecycleBin.Value = 1 Then Logg "Option Delete files to Recycle Bin: Yes"
    
    'show progress frame
    SetProgress 0&
    Status "Initializing..."
    
    lPause = 0
        
    'kill explorer when needed
    If bUnloadShell Then ProcessKill sWinDir & "\explorer.exe|1"
    
    sFullScript = ExpandEnvironmentVars(Replace(sFullScript, "%%", "#"))
    sScript = Split(sFullScript, vbCrLf)
    For i = 0 To UBound(sScript)
        sScript(i) = Trim(sScript(i))
        If Not Trim(sScript(i)) = vbNullString Then
            
            'check command against list
            If Left(LTrim(sScript(i)), 1) <> "#" Then
                bUnknownCmd = True
                For j = 0 To UBound(sCommandsList)
                    If sCommandsList(j) <> vbNullString Then
                        If InStr(1, sScript(i), sCommandsList(j), vbTextCompare) > 0 Then
                            bUnknownCmd = False
                        End If
                    End If
                Next j
                If bUnknownCmd Then
                    Logg "Warning: unknown command '" & Trim(sScript(i)) & "' on line #" & i + 1
                End If
            End If
            
            If Not bStatusMsgs Then Status sScript(i)
            
            If InStr(1, sScript(i), "Option", vbTextCompare) = 1 Then SetScriptOption sScript(i)
            If bAbortScript Then Exit For
            
            If InStr(sScript(i), "%") > 0 Then GoTo Skip
            
            'files & folders
            If InStr(1, sScript(i), "FileCreate ", vbTextCompare) = 1 Then FileCreate LTrim(Mid(sScript(i), Len("FileCreate") + 1))
            If InStr(1, sScript(i), "FileDelete ", vbTextCompare) = 1 Then FileDelete LTrim(Mid(sScript(i), Len("FileDelete") + 1))
            If InStr(1, sScript(i), "FileDeleteOnReboot ", vbTextCompare) = 1 Then FileDeleteOnReboot LTrim(Mid(sScript(i), Len("FileDeleteOnReboot") + 1))
            If InStr(1, sScript(i), "FileRename ", vbTextCompare) = 1 Then FileRename LTrim(Mid(sScript(i), Len("FileRename") + 1))
            If InStr(1, sScript(i), "FileMove ", vbTextCompare) = 1 Then FileMove LTrim(Mid(sScript(i), Len("FileMove") + 1))
            If InStr(1, sScript(i), "FileClear ", vbTextCompare) = 1 Then FileClear LTrim(Mid(sScript(i), Len("FileClear") + 1))
            If InStr(1, sScript(i), "FileSetAttributes ", vbTextCompare) = 1 Then FileSetAttributes LTrim(Mid(sScript(i), Len("FileSetAttributes") + 1))
            If InStr(1, sScript(i), "FolderCreate ", vbTextCompare) = 1 Then FolderCreate LTrim(Mid(sScript(i), Len("FolderCreate") + 1))
            If InStr(1, sScript(i), "FolderRename ", vbTextCompare) = 1 Then FolderRename LTrim(Mid(sScript(i), Len("FolderRename") + 1))
            If InStr(1, sScript(i), "FolderMove ", vbTextCompare) = 1 Then FolderMove LTrim(Mid(sScript(i), Len("FolderMove") + 1))
            If InStr(1, sScript(i), "FolderDelete ", vbTextCompare) = 1 Then FolderDelete LTrim(Mid(sScript(i), Len("FolderDelete") + 1))
            If InStr(1, sScript(i), "FolderSetAttributes ", vbTextCompare) = 1 Then FolderSetAttributes LTrim(Mid(sScript(i), Len("FolderSetAttributes") + 1))
            If InStr(1, sScript(i), "FolderClear ", vbTextCompare) = 1 Then FolderClear LTrim(Mid(sScript(i), Len("FolderClear") + 1))
            If InStr(1, sScript(i), "FileWrite ", vbTextCompare) = 1 Then FileWrite LTrim(Mid(sScript(i), Len("FileWrite") + 1))
            
            If InStr(1, sScript(i), "FileDeleteIfContainsText ", vbTextCompare) = 1 Then FileDeleteIfContainsText LTrim(Mid(sScript(i), Len("FileDeleteIfContainsText") + 1))
            If InStr(1, sScript(i), "FileDeleteIfContainsHex ", vbTextCompare) = 1 Then FileDeleteIfContainsHex LTrim(Mid(sScript(i), Len("FileDeleteIfContainsHex") + 1))
            If InStr(1, sScript(i), "FileDeleteIfMD5Match ", vbTextCompare) = 1 Then FileDeleteIfMD5Match LTrim(Mid(sScript(i), Len("FileDeleteIfMD5Match") + 1))
            If InStr(1, sScript(i), "FileDeleteIfCRC32Match ", vbTextCompare) = 1 Then FileDeleteIfCRC32Match LTrim(Mid(sScript(i), Len("FileDeleteIfCRC32Match") + 1))
            If InStr(1, sScript(i), "FileDeleteIfSHA1Match ", vbTextCompare) = 1 Then FileDeleteIfSHA1Match LTrim(Mid(sScript(i), Len("FileDeleteIfSHA1Match") + 1))
            If InStr(1, sScript(i), "FileDeleteIfMD2Match ", vbTextCompare) = 1 Then FileDeleteIfMD2Match LTrim(Mid(sScript(i), Len("FileDeleteIfMD2Match") + 1))
            If InStr(1, sScript(i), "FileDeleteIfMD4Match ", vbTextCompare) = 1 Then FileDeleteIfMD4Match LTrim(Mid(sScript(i), Len("FileDeleteIfMD4Match") + 1))
            
            If InStr(1, sScript(i), "FileMoveIfContainsText ", vbTextCompare) = 1 Then FileMoveIfContainsText LTrim(Mid(sScript(i), Len("FileMoveIfContainsText") + 1))
            If InStr(1, sScript(i), "FileMoveIfContainsHex ", vbTextCompare) = 1 Then FileMoveIfContainsHex LTrim(Mid(sScript(i), Len("FileMoveIfContainsHex") + 1))
            If InStr(1, sScript(i), "FileMoveIfMD5Match ", vbTextCompare) = 1 Then FileMoveIfMD5Match LTrim(Mid(sScript(i), Len("FileMoveIfMD5Match") + 1))
            If InStr(1, sScript(i), "FileMoveIfCRC32Match ", vbTextCompare) = 1 Then FileMoveIfCRC32Match LTrim(Mid(sScript(i), Len("FileMoveIfCRC32Match") + 1))
            If InStr(1, sScript(i), "FileMoveIfSHA1Match ", vbTextCompare) = 1 Then FileMoveIfSHA1Match LTrim(Mid(sScript(i), Len("FileMoveIfSHA1Match") + 1))
            If InStr(1, sScript(i), "FileMoveIfMD2Match ", vbTextCompare) = 1 Then FileMoveIfMD2Match LTrim(Mid(sScript(i), Len("FileMoveIfMD2Match") + 1))
            If InStr(1, sScript(i), "FileMoveIfMD4Match ", vbTextCompare) = 1 Then FileMoveIfMD4Match LTrim(Mid(sScript(i), Len("FileMoveIfMD4Match") + 1))
            
            'ini files
            If InStr(1, sScript(i), "IniSetValue ", vbTextCompare) = 1 Then IniSetValue LTrim(Mid(sScript(i), Len("IniSetValue") + 1))
            If InStr(1, sScript(i), "IniDeleteValue ", vbTextCompare) = 1 Then IniDeleteValue LTrim(Mid(sScript(i), Len("IniDeleteValue") + 1))
            If InStr(1, sScript(i), "IniDeleteFromValue ", vbTextCompare) = 1 Then IniDeleteFromValue LTrim(Mid(sScript(i), Len("IniDeleteFromValue") + 1))
            If InStr(1, sScript(i), "IniClearValue ", vbTextCompare) = 1 Then IniClearValue LTrim(Mid(sScript(i), Len("IniClearValue") + 1))
            If InStr(1, sScript(i), "IniCreateSection ", vbTextCompare) = 1 Then IniCreateSection LTrim(Mid(sScript(i), Len("IniCreateSection") + 1))
            
            'hosts file
            If InStr(1, sScript(i), "HostsFileReset ", vbTextCompare) = 1 Then HostsFileReset
            If InStr(1, sScript(i), "HostsFileAddLine ", vbTextCompare) = 1 Then HostsFileAddLine LTrim(Mid(sScript(i), Len("HostsFileAddLine") + 1))
            If InStr(1, sScript(i), "HostsFileDelLine ", vbTextCompare) = 1 Then HostsFileDelLine LTrim(Mid(sScript(i), Len("HostsFileDelLine") + 1))
            If InStr(1, sScript(i), "HostsFileDisableLine ", vbTextCompare) = 1 Then HostsFileDisableLine LTrim(Mid(sScript(i), Len("HostsFileDisableLine") + 1))
            If InStr(1, sScript(i), "HostsFileEnableLine ", vbTextCompare) = 1 Then HostsFileEnableLine LTrim(Mid(sScript(i), Len("HostsFileEnableLine") + 1))
            
            'regkeys & regvalues
            If InStr(1, sScript(i), "RegCreateKey ", vbTextCompare) = 1 Then RegCreateKey LTrim(Mid(sScript(i), Len("RegCreateKey") + 1))
            If InStr(1, sScript(i), "RegDeleteKey ", vbTextCompare) = 1 Then RegDeleteKey LTrim(Mid(sScript(i), Len("RegDeleteKey") + 1))
            If InStr(1, sScript(i), "RegDeleteKeyIfNameContainsText ", vbTextCompare) = 1 Then RegDeleteKeyIfNameContainsText LTrim(Mid(sScript(i), Len("RegDeleteKeyIfNameContainsText") + 1))
            If InStr(1, sScript(i), "RegDeleteKeyIfNameContainsHex ", vbTextCompare) = 1 Then RegDeleteKeyIfNameContainsHex LTrim(Mid(sScript(i), Len("RegDeleteKeyIfNameContainsHex") + 1))
            If InStr(1, sScript(i), "RegSetStringValue ", vbTextCompare) = 1 Then RegSetStringValue LTrim(Mid(sScript(i), Len("RegSetStringValue") + 1))
            If InStr(1, sScript(i), "RegSetDwordValue ", vbTextCompare) = 1 Then RegSetDwordValue LTrim(Mid(sScript(i), Len("RegSetDwordValue") + 1))
            If InStr(1, sScript(i), "RegSetBinaryValue ", vbTextCompare) = 1 Then RegSetBinaryValue LTrim(Mid(sScript(i), Len("RegSetBinaryValue") + 1))
            If InStr(1, sScript(i), "RegSetMultiValue ", vbTextCompare) = 1 Then RegSetMultiValue LTrim(Mid(sScript(i), Len("RegSetMultiValue") + 1))
            If InStr(1, sScript(i), "RegSetExpandValue ", vbTextCompare) = 1 Then RegSetExpandValue LTrim(Mid(sScript(i), Len("RegSetExpandValue") + 1))
            If InStr(1, sScript(i), "RegDelValue ", vbTextCompare) = 1 Then RegDelValue LTrim(Mid(sScript(i), Len("RegDelValue") + 1))
            If InStr(1, sScript(i), "RegDelFromValue ", vbTextCompare) = 1 Then RegDelFromValue LTrim(Mid(sScript(i), Len("RegDelFromValue") + 1))
            If InStr(1, sScript(i), "RegRenameValue ", vbTextCompare) = 1 Then RegRenameValue LTrim(Mid(sScript(i), Len("RegRenameValue") + 1))
            If InStr(1, sScript(i), "RegDelValueIfDataContainsText ", vbTextCompare) = 1 Then RegDelValueIfDataContainsText LTrim(Mid(sScript(i), Len("RegDelValueIfDataContainsText") + 1))
            If InStr(1, sScript(i), "RegDelValueIfDataContainsHex ", vbTextCompare) = 1 Then RegDelValueIfDataContainsHex LTrim(Mid(sScript(i), Len("RegDelValueIfDataContainsHex") + 1))
            If InStr(1, sScript(i), "RegDelValueIfNameContainsText ", vbTextCompare) = 1 Then RegDelValueIfNameContainsText LTrim(Mid(sScript(i), Len("RegDelValueIfNameContainsText") + 1))
            If InStr(1, sScript(i), "RegDelValueIfNameContainsHex ", vbTextCompare) = 1 Then RegDelValueIfNameContainsHex LTrim(Mid(sScript(i), Len("RegDelValueIfNameContainsHex") + 1))
            If InStr(1, sScript(i), "RegResetPermissions ", vbTextCompare) = 1 Then RegResetPermissions LTrim(Mid(sScript(i), Len("RegResetPermissions") + 1))
            
            'processes
            If InStr(1, sScript(i), "ProcessKill ", vbTextCompare) = 1 Then ProcessKill LTrim(Mid(sScript(i), Len("ProcessKill") + 1))
            If InStr(1, sScript(i), "ProcessKillIfMD5Match ", vbTextCompare) = 1 Then ProcessKillIfMD5Match LTrim(Mid(sScript(i), Len("ProcessKillIfMD5Match") + 1))
            If InStr(1, sScript(i), "ProcessKillIfCRC32Match ", vbTextCompare) = 1 Then ProcessKillIfCRC32Match LTrim(Mid(sScript(i), Len("ProcessKillIfCRC32Match") + 1))
            If InStr(1, sScript(i), "ProcessKillIfSHA1Match ", vbTextCompare) = 1 Then ProcessKillIfSHA1Match LTrim(Mid(sScript(i), Len("ProcessKillIfSHA1Match") + 1))
            If InStr(1, sScript(i), "ProcessKillIfMD2Match ", vbTextCompare) = 1 Then ProcessKillIfMD2Match LTrim(Mid(sScript(i), Len("ProcessKillIfMD2Match") + 1))
            If InStr(1, sScript(i), "ProcessKillIfMD4Match ", vbTextCompare) = 1 Then ProcessKillIfMD4Match LTrim(Mid(sScript(i), Len("ProcessKillIfMD4Match") + 1))
            If InStr(1, sScript(i), "ProcessKillIfContainsText ", vbTextCompare) = 1 Then ProcessKillIfContainsText LTrim(Mid(sScript(i), Len("ProcessKillIfContainsText") + 1))
            If InStr(1, sScript(i), "ProcessKillIfContainsHex ", vbTextCompare) = 1 Then ProcessKillIfContainsHex LTrim(Mid(sScript(i), Len("ProcessKillIfContainsHex") + 1))
            
            If InStr(1, sScript(i), "ProcessSuspend ", vbTextCompare) = 1 Then ProcessSuspend LTrim(Mid(sScript(i), Len("ProcessSuspend") + 1))
            If InStr(1, sScript(i), "ProcessSuspendIfMD5Match ", vbTextCompare) = 1 Then ProcessSuspendIfMD5Match LTrim(Mid(sScript(i), Len("ProcessSuspendIfMD5Match") + 1))
            If InStr(1, sScript(i), "ProcessSuspendIfCRC32Match ", vbTextCompare) = 1 Then ProcessSuspendIfCRC32Match LTrim(Mid(sScript(i), Len("ProcessSuspendIfCRC32Match") + 1))
            If InStr(1, sScript(i), "ProcessSuspendIfSHA1Match ", vbTextCompare) = 1 Then ProcessSuspendIfSHA1Match LTrim(Mid(sScript(i), Len("ProcessSuspendIfSHA1Match") + 1))
            If InStr(1, sScript(i), "ProcessSuspendIfMD2Match ", vbTextCompare) = 1 Then ProcessSuspendIfMD2Match LTrim(Mid(sScript(i), Len("ProcessSuspendIfMD2Match") + 1))
            If InStr(1, sScript(i), "ProcessSuspendIfMD4Match ", vbTextCompare) = 1 Then ProcessSuspendIfMD4Match LTrim(Mid(sScript(i), Len("ProcessSuspendIfMD4Match") + 1))
            If InStr(1, sScript(i), "ProcessSuspendIfContainsText ", vbTextCompare) = 1 Then ProcessSuspendIfContainsText LTrim(Mid(sScript(i), Len("ProcessSuspendIfContainsText") + 1))
            If InStr(1, sScript(i), "ProcessSuspendIfContainsHex ", vbTextCompare) = 1 Then ProcessSuspendIfContainsHex LTrim(Mid(sScript(i), Len("ProcessSuspendIfContainsHex") + 1))
            
            If InStr(1, sScript(i), "ProcessResume ", vbTextCompare) = 1 Then ProcessResume LTrim(Mid(sScript(i), Len("ProcessResume") + 1))
            If InStr(1, sScript(i), "ProcessResumeIfMD5Match ", vbTextCompare) = 1 Then ProcessResumeIfMD5Match LTrim(Mid(sScript(i), Len("ProcessResumeIfMD5Match") + 1))
            If InStr(1, sScript(i), "ProcessResumeIfCRC32Match ", vbTextCompare) = 1 Then ProcessResumeIfCRC32Match LTrim(Mid(sScript(i), Len("ProcessResumeIfCRC32Match") + 1))
            If InStr(1, sScript(i), "ProcessResumeIfSHA1Match ", vbTextCompare) = 1 Then ProcessResumeIfSHA1Match LTrim(Mid(sScript(i), Len("ProcessResumeIfSHA1Match") + 1))
            If InStr(1, sScript(i), "ProcessResumeIfMD2Match ", vbTextCompare) = 1 Then ProcessResumeIfMD2Match LTrim(Mid(sScript(i), Len("ProcessResumeIfMD2Match") + 1))
            If InStr(1, sScript(i), "ProcessResumeIfMD4Match ", vbTextCompare) = 1 Then ProcessResumeIfMD4Match LTrim(Mid(sScript(i), Len("ProcessResumeIfMD4Match") + 1))
            If InStr(1, sScript(i), "ProcessResumeIfContainsText ", vbTextCompare) = 1 Then ProcessResumeIfContainsText LTrim(Mid(sScript(i), Len("ProcessResumeIfContainsText") + 1))
            If InStr(1, sScript(i), "ProcessResumeIfContainsHex ", vbTextCompare) = 1 Then ProcessResumeIfContainsHex LTrim(Mid(sScript(i), Len("ProcessResumeIfContainsHex") + 1))
            
            'services
            If InStr(1, sScript(i), "ServiceStart ", vbTextCompare) = 1 Then ServiceStart LTrim(Mid(sScript(i), Len("ServiceStart") + 1))
            If InStr(1, sScript(i), "ServiceStop ", vbTextCompare) = 1 Then ServiceStop LTrim(Mid(sScript(i), Len("ServiceStop") + 1))
            If InStr(1, sScript(i), "ServiceDisable ", vbTextCompare) = 1 Then ServiceDisable LTrim(Mid(sScript(i), Len("ServiceDisable") + 1))
            If InStr(1, sScript(i), "ServiceEnable ", vbTextCompare) = 1 Then ServiceEnable LTrim(Mid(sScript(i), Len("ServiceEnable") + 1))
            If InStr(1, sScript(i), "ServiceDelete ", vbTextCompare) = 1 Then ServiceDelete LTrim(Mid(sScript(i), Len("ServiceDelete") + 1))
            
            'winsock
            If InStr(1, sScript(i), "WinsockKillProtocol ", vbTextCompare) = 1 Then WinsockKillProtocol LTrim(Mid(sScript(i), Len("WinsockKillProtocol") + 1))
            If InStr(1, sScript(i), "WinsockKillNameSpace ", vbTextCompare) = 1 Then WinsockKillNameSpace LTrim(Mid(sScript(i), Len("WinsockKillNameSpace") + 1))
            
            'system
            If InStr(1, sScript(i), "SystemRestart", vbTextCompare) = 1 Then
                If InStr(1, sScript(i), "SystemRestartIfNeeded", vbTextCompare) = 0 Then
                    SystemRestart LTrim(Mid(sScript(i), Len("SystemRestart") + 1))
                End If
            End If
            If InStr(1, sScript(i), "SystemRestartIfNeeded", vbTextCompare) = 1 Then SystemRestartIfNeeded LTrim(Mid(sScript(i), Len("SystemRestartIfNeeded") + 1))
            If InStr(1, sScript(i), "SystemRun ", vbTextCompare) = 1 Then SystemRun LTrim(Mid(sScript(i), Len("SystemRun") + 1))
            If InStr(1, sScript(i), "SystemMsgBox ", vbTextCompare) = 1 Then SystemMsgBox LTrim(Mid(sScript(i), Len("SystemMsgBox") + 1))
            If InStr(1, sScript(i), "SystemResetWebSettings", vbTextCompare) = 1 Then SystemResetWebSettings
            If InStr(1, sScript(i), "SystemEmptyRecycleBin", vbTextCompare) = 1 Then SystemEmptyRecycleBin
            If InStr(1, sScript(i), "SystemEmptyInternetCache", vbTextCompare) = 1 Then SystemEmptyInternetCache
            If InStr(1, sScript(i), "SystemEmptyTempFolder", vbTextCompare) = 1 Then SystemEmptyTempFolder
            
            'dlls
            If InStr(1, sScript(i), "DllRegister ", vbTextCompare) = 1 Then DllRegister LTrim(Mid(sScript(i), Len("DllRegister") + 1))
            If InStr(1, sScript(i), "DllUnregister ", vbTextCompare) = 1 Then DllUnregister LTrim(Mid(sScript(i), Len("DllUnregister") + 1))
        
            If InStr(1, sScript(i), "LogIfFileExists ", vbTextCompare) = 1 Then LogIfFileExists LTrim(Mid(sScript(i), Len("LogIfFileExists") + 1))
            If InStr(1, sScript(i), "LogIfFileContainsText ", vbTextCompare) = 1 Then LogIfFileContainsText LTrim(Mid(sScript(i), Len("LogIfFileContainsText") + 1))
            If InStr(1, sScript(i), "LogIfFileContainsHex ", vbTextCompare) = 1 Then LogIfFileContainsHex LTrim(Mid(sScript(i), Len("LogIfFileContainsHex") + 1))
            If InStr(1, sScript(i), "LogIfFolderExists ", vbTextCompare) = 1 Then LogIfFolderExists LTrim(Mid(sScript(i), Len("LogIfFolderExists") + 1))
            If InStr(1, sScript(i), "LogIfRegKeyExists ", vbTextCompare) = 1 Then LogIfRegKeyExists LTrim(Mid(sScript(i), Len("LogIfRegKeyExists") + 1))
            If InStr(1, sScript(i), "LogIfRegValExists ", vbTextCompare) = 1 Then LogIfRegValExists LTrim(Mid(sScript(i), Len("LogIfRegValExists") + 1))
            If InStr(1, sScript(i), "LogIfRegValContainsText ", vbTextCompare) = 1 Then LogIfRegValContainsText LTrim(Mid(sScript(i), Len("LogIfRegValContainsText") + 1))
            If InStr(1, sScript(i), "LogIfRegValContainsHex ", vbTextCompare) = 1 Then LogIfRegValContainsHex LTrim(Mid(sScript(i), Len("LogIfRegValContainsHex") + 1))
        End If
                
Skip:
        'progress dialog
        If UBound(sScript) > 0 Then SetProgress CDbl(i) / UBound(sScript)
        DoEvents
        If lPause > 0 Then
            Sleep lPause
        End If
    Next i
    Status "Completed!"
    
    'reload explorer
    If bUnloadShell Then SystemRun sWinDir & "\explorer.exe||1"
    
    'hide progress frame
    SetProgress -1
    
    If bAbortScript Then
        If Not bRunSilent Then MsgBox "Script aborted!", vbExclamation
    Else
        If Not bRunSilent Then MsgBox "Completed script execution.", vbInformation
    End If
    Logg "Script completed at " & Time & "."
    If sSaveLogFile <> vbNullString Then SaveLogFile
End Sub

Private Sub SetProgress(nPercentage#)
    With frmMain
        If nPercentage = 0 Then
            .fraOptions.Visible = False
            .fraProgress.Visible = True
            .shpProgress.Width = 15
            .lblProgress.Caption = "0 %"
            Exit Sub
        End If
        If nPercentage = -1 Then
            .fraProgress.Visible = False
            .fraOptions.Visible = True
            Exit Sub
        End If
        .shpProgress.Width = nPercentage * .shpProgressBackgrond.Width
        .lblProgress.Caption = CStr(Int(nPercentage * 100)) & " %"
        .lblProgress.Left = .shpProgress.Left + .shpProgress.Width - .lblProgress.Width - 60
        If .lblProgress.Left < 360 Then .lblProgress.Left = 360
        DoEvents
    End With
End Sub

Public Sub Logg(s$)
    sLog = sLog & s & vbCrLf
End Sub

Public Sub Status(s$)
    frmMain.lblCurrentAction.Caption = s
    DoEvents
End Sub

Private Sub SaveLogFile()
    If InStr(sSaveLogFile, "\") = 0 Then
        sSaveLogFile = App.Path & IIf(Right(App.Path, 1) = "\", "", "\") & sSaveLogFile
    End If
    
    If FileExists(sSaveLogFile) Then
        On Error Resume Next
        Kill sSaveLogFile
        If Err Then Logg "Unable to save log to " & sSaveLogFile & " (access denied)"
        On Error GoTo 0:
        Exit Sub
    End If
    
    sSaveLogFile = ExpandEnvironmentVars(sSaveLogFile)
    OutputFile sSaveLogFile, sLog
    If FileExists(sSaveLogFile) Then
        Logg "Saved log output to " & sSaveLogFile
    Else
        Logg "Unable to save log to " & sSaveLogFile & " (write error)"
    End If
End Sub
