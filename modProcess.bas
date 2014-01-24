Attribute VB_Name = "modProcess"
Option Explicit
'killing processes

Private Declare Function CreateToolhelpSnapshot Lib "kernel32" Alias "CreateToolhelp32Snapshot" (ByVal lFlags As Long, ByVal lProcessID As Long) As Long
Private Declare Function ProcessFirst Lib "kernel32" Alias "Process32First" (ByVal hSnapshot As Long, uProcess As PROCESSENTRY32) As Long
Private Declare Function ProcessNext Lib "kernel32" Alias "Process32Next" (ByVal hSnapshot As Long, uProcess As PROCESSENTRY32) As Long
'Private Declare Function Module32First Lib "kernel32" (ByVal hSnapshot As Long, uProcess As MODULEENTRY32) As Long
'Private Declare Function Module32Next Lib "kernel32" (ByVal hSnapshot As Long, uProcess As MODULEENTRY32) As Long
Private Declare Function Thread32First Lib "kernel32" (ByVal hSnapshot As Long, uThread As THREADENTRY32) As Long
Private Declare Function Thread32Next Lib "kernel32" (ByVal hSnapshot As Long, ByRef ThreadEntry As THREADENTRY32) As Long
Private Declare Function TerminateProcess Lib "kernel32" (ByVal hProcess As Long, ByVal uExitCode As Long) As Long

Private Declare Function SuspendThread Lib "kernel32" (ByVal hThread As Long) As Long
Private Declare Function ResumeThread Lib "kernel32" (ByVal hThread As Long) As Long
Private Declare Function OpenThread Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Boolean, ByVal dwThreadId As Long) As Long
Private Declare Function GetCurrentProcessId Lib "kernel32" () As Long

Private Declare Function EnumProcesses Lib "PSAPI.DLL" (ByRef lpidProcess As Long, ByVal cb As Long, ByRef cbNeeded As Long) As Long
Private Declare Function GetModuleFileNameExA Lib "PSAPI.DLL" (ByVal hProcess As Long, ByVal hModule As Long, ByVal ModuleName As String, ByVal nSize As Long) As Long
Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Private Declare Function EnumProcessModules Lib "PSAPI.DLL" (ByVal hProcess As Long, ByRef lphModule As Long, ByVal cb As Long, ByRef cbNeeded As Long) As Long

Private Type PROCESSENTRY32
    dwSize As Long
    cntUsage As Long
    th32ProcessID As Long
    th32DefaultHeapID As Long
    th32ModuleID As Long
    cntThreads As Long
    th32ParentProcessID As Long
    pcPriClassBase As Long
    dwFlags As Long
    szExeFile As String * 260
End Type

Private Type MODULEENTRY32
    dwSize As Long
    th32ModuleID As Long
    th32ProcessID As Long
    GlblcntUsage As Long
    ProccntUsage As Long
    modBaseAddr As Long
    modBaseSize As Long
    hModule As Long
    szModule As String * 256
    szExePath As String * 260
End Type

Private Type THREADENTRY32
    dwSize As Long
    dwRefCount As Long
    th32ThreadID As Long
    th32ProcessID As Long
    dwBasePriority As Long
    dwCurrentPriority As Long
    dwFlags As Long
End Type

Private Const TH32CS_SNAPPROCESS = &H2
Private Const TH32CS_SNAPMODULE = &H8
Private Const TH32CS_SNAPTHREAD = &H4
Private Const THREAD_SUSPEND_RESUME = &H2
Private Const PROCESS_TERMINATE = &H1
Private Const PROCESS_QUERY_INFORMATION = 1024
Private Const PROCESS_VM_READ = 16

Public Sub ProcessKill(sCmd$)
    'ProcessKill <process/mask>|[all matches(0|1)]
    Dim sProcess$, bKillAll As Boolean
    sProcess = sCmd
    If sProcess = vbNullString Then Exit Sub
    If InStr(sProcess, "|") > 0 Then
        bKillAll = CBool(Val(Mid(sProcess, InStr(sProcess, "|") + 1)))
        sProcess = Left(sProcess, InStr(sProcess, "|") - 1)
    End If
    
    Dim sMatches$(), i&, lPID&, sExe$
    sMatches = Split(GetMatchingProcesses(sProcess), "|")
    For i = 0 To UBound(sMatches)
        lPID = CLng(Left(sMatches(i), InStr(sMatches(i), ",") - 1))
        sExe = Mid(sMatches(i), InStr(sMatches(i), ",") + 1)
        If Not PauseProcess(lPID) Then
            Logg "Warning: ProcessKill " & sCmd & " (suspend failed)"
        End If
        If ProcessKillByPID(lPID) = False Then
            Logg "Failed: ProcessKill " & sCmd & " (operation failed)"
            PauseProcess lPID, False
        Else
            Logg "Success: ProcessKill " & sCmd
        End If
        If Not bKillAll Then Exit For
    Next i
End Sub

Public Sub ProcessSuspend(sCmd$)
    'ProcessSuspend <process/mask>|[all matches(0|1)]
    Dim sProcess$, bSuspendAll As Boolean
    sProcess = sCmd
    If sProcess = vbNullString Then Exit Sub
    If InStr(sProcess, "|") > 0 Then
        bSuspendAll = CBool(Val(Mid(sProcess, InStr(sProcess, "|") + 1)))
        sProcess = Left(sProcess, InStr(sProcess, "|") - 1)
    End If
    
    Dim sMatches$(), i&, lPID&, sExe$
    sMatches = Split(GetMatchingProcesses(sProcess), "|")
    For i = 0 To UBound(sMatches)
        lPID = CLng(Left(sMatches(i), InStr(sMatches(i), ",") - 1))
        sExe = Mid(sMatches(i), InStr(sMatches(i), ",") + 1)
        If Not PauseProcess(lPID) Then
            Logg "Failed: ProcessSuspend " & sCmd & " (operation failed)"
        Else
            Logg "Success: ProcessSuspend " & sCmd
        End If
        If Not bSuspendAll Then Exit For
    Next i
End Sub

Public Sub ProcessResume(sCmd$)
    'ProcessResume <process/mask>|[all matches(0|1)]
    Dim sProcess$, bResumeAll As Boolean
    sProcess = sCmd
    If sProcess = vbNullString Then Exit Sub
    If InStr(sProcess, "|") > 0 Then
        bResumeAll = CBool(Val(Mid(sProcess, InStr(sProcess, "|") + 1)))
        sProcess = Left(sProcess, InStr(sProcess, "|") - 1)
    End If
    
    Dim sMatches$(), i&, lPID&, sExe$
    sMatches = Split(GetMatchingProcesses(sProcess), "|")
    For i = 0 To UBound(sMatches)
        lPID = CLng(Left(sMatches(i), InStr(sMatches(i), ",") - 1))
        sExe = Mid(sMatches(i), InStr(sMatches(i), ",") + 1)
        If Not PauseProcess(lPID, False) Then
            Logg "Failed: ProcessResume " & sCmd & " (operation failed)"
        Else
            Logg "Success: ProcessResume " & sCmd
        End If
        If Not bResumeAll Then Exit For
    Next i
End Sub

Private Function GetMatchingProcesses$(sProcMask$)
    'internal functions
    'gets PIDs of processes matching the mask
    Dim sList$, i&, hProc&
    Dim hSnap&, uPE32 As PROCESSENTRY32, sExeFile$
    Dim lProcesses&(1 To 1024), lNeeded&, lNumProcesses&, sProcessName$, lModules&(1 To 1024)
    If Not bIsWinNT Then
        'windows 9x/me method
        hSnap = CreateToolhelpSnapshot(TH32CS_SNAPPROCESS, 0)
        If hSnap > 0 Then
            uPE32.dwSize = Len(uPE32)
            If ProcessFirst(hSnap, uPE32) = 0 Then
                CloseHandle hSnap
                Exit Function
            End If
            
            Do
                sExeFile = TrimNull(uPE32.szExeFile)
                If LCase(sExeFile) Like "*" & LCase(sProcMask) & "*" Then
                'If InStr(1, sExeFile, sProcMask, vbTextCompare) > 0 Then
                    sList = sList & "|" & CStr(uPE32.th32ProcessID) & "," & sExeFile
                End If
            Loop Until ProcessNext(hSnap, uPE32) = 0
            CloseHandle hSnap
        End If
    Else
        'windows nt/2k/xp/2003/etc method
        On Error Resume Next
        If EnumProcesses(lProcesses(1), CLng(1024) * 4, lNeeded) = 0 Then
            Exit Function
        End If
        lNumProcesses = lNeeded / 4
        For i = 1 To lNumProcesses
            hProc = OpenProcess(PROCESS_QUERY_INFORMATION Or PROCESS_VM_READ Or PROCESS_TERMINATE, 0, lProcesses(i))
            If hProc <> 0 Then
                lNeeded = 0
                sProcessName = String(260, 0)
                If EnumProcessModules(hProc, lModules(i), CLng(1024) * 4, lNeeded) <> 0 Then
                    GetModuleFileNameExA hProc, lModules(1), sProcessName, Len(sProcessName)
                    sProcessName = TrimNull(sProcessName)
                    If sProcessName <> vbNullString Then
                        If Left(sProcessName, 1) = "\" Then sProcessName = Mid(sProcessName, 2)
                        If Left(sProcessName, 3) = "??\" Then sProcessName = Mid(sProcessName, 4)
                        If InStr(1, sProcessName, "%SystemRoot%", vbTextCompare) > 0 Then sProcessName = Replace(sProcessName, "%SystemRoot%", sWinDir, , , vbTextCompare)
                        If InStr(1, sProcessName, "SystemRoot", vbTextCompare) > 0 Then sProcessName = Replace(sProcessName, "SystemRoot", sWinDir, , , vbTextCompare)
                        
                        If LCase(sProcessName) Like "*" & LCase(sProcMask) & "*" Then
                        'If InStr(1, sProcessName, sProcMask, vbTextCompare) > 0 Then
                            sList = sList & "|" & CStr(lProcesses(i)) & "," & sProcessName
                        End If
                    End If
                End If
                CloseHandle hProc
            End If
        Next i
    End If
    If sList <> vbNullString Then GetMatchingProcesses = Mid(sList, 2)
End Function

Private Function ProcessKillByPID(lPID) As Boolean
    'internal functions
    Dim hProc&
    'this is platform independant :P
    hProc = OpenProcess(PROCESS_TERMINATE, 0, lPID)
    If hProc <> 0 Then
        If TerminateProcess(hProc, 0) <> 0 Then
            Logg "Success: ProcessKillByPID " & lPID
            ProcessKillByPID = True
        Else
            Logg "Failed: ProcessKillByPID " & lPID & " (operation failed)"
            ProcessKillByPID = False
        End If
        CloseHandle hProc
    End If
End Function

Private Function PauseProcess(lPID&, Optional bPauseOrResume As Boolean = True) As Boolean
    'internal function
    Dim hSnap&, uTE32 As THREADENTRY32, hThread&
    If Not bIsWinNT And Not bIsWinME Then Exit Function
    If bIsWinNT4 Then Exit Function
    If lPID = GetCurrentProcessId Then Exit Function
    
    hSnap = CreateToolhelpSnapshot(TH32CS_SNAPTHREAD, lPID)
    If hSnap = -1 Then Exit Function
    
    uTE32.dwSize = Len(uTE32)
    If Thread32First(hSnap, uTE32) = 0 Then
        CloseHandle hSnap
        Exit Function
    End If
    
    PauseProcess = True
    Do
        If uTE32.th32ProcessID = lPID Then
            hThread = OpenThread(THREAD_SUSPEND_RESUME, False, uTE32.th32ThreadID)
            If bPauseOrResume Then
                If SuspendThread(hThread) = -1 Then PauseProcess = False
            Else
                If ResumeThread(hThread) = -1 Then PauseProcess = False
            End If
            CloseHandle hThread
        End If
    Loop Until Thread32Next(hSnap, uTE32) = 0
    CloseHandle hSnap
End Function

Public Sub ProcessKillIfMD5Match(sCmd$)
    'ProcessKillIfMD5Match <file/mask>|<md5>|[0|1]
    Dim sProcess$, sMD5$, bKillAll As Boolean, sArgs$()
    sArgs = Split(sCmd, "|")
    Select Case UBound(sArgs)
        Case 1
            sProcess = sArgs(0)
            sMD5 = sArgs(1)
        Case 2
            sProcess = sArgs(0)
            sMD5 = sArgs(1)
            bKillAll = CBool(Val(sArgs(2)))
        Case Else: Exit Sub
    End Select
    
    Dim sMatches$(), lPID&, sExe$, i&
    sMatches = Split(GetMatchingProcesses(sProcess), "|")
    For i = 0 To UBound(sMatches)
        lPID = CLng(Left(sMatches(i), InStr(sMatches(i), ",") - 1))
        sExe = Mid(sMatches(i), InStr(sMatches(i), ",") + 1)
        If UCase(sMD5) = GetFileMD5(sExe) Then
            If Not PauseProcess(lPID) Then
                Logg "Warning: ProcessKillIfMD5Match " & sCmd & " (suspend failed)"
            End If
            If Not ProcessKillByPID(lPID) Then
                Logg "Failed: ProcessKillIfMD5Match " & sCmd & " (operation failed)"
                PauseProcess lPID, False
            Else
                Logg "Success: ProcessKillIfMD5Match " & sCmd & " (matched process " & sExe & ")"
            End If
            If Not bKillAll Then Exit For
        End If
    Next i
End Sub

Public Sub ProcessKillIfSHA1Match(sCmd$)
    'ProcessKillIfSHA1Match <file/mask>|<SHA1>|[0|1]
    Dim sProcess$, sSHA1$, bKillAll As Boolean, sArgs$()
    sArgs = Split(sCmd, "|")
    Select Case UBound(sArgs)
        Case 1
            sProcess = sArgs(0)
            sSHA1 = sArgs(1)
        Case 2
            sProcess = sArgs(0)
            sSHA1 = sArgs(1)
            bKillAll = CBool(Val(sArgs(2)))
        Case Else: Exit Sub
    End Select
    
    Dim sMatches$(), lPID&, sExe$, i&
    sMatches = Split(GetMatchingProcesses(sProcess), "|")
    For i = 0 To UBound(sMatches)
        lPID = CLng(Left(sMatches(i), InStr(sMatches(i), ",") - 1))
        sExe = Mid(sMatches(i), InStr(sMatches(i), ",") + 1)
        If UCase(sSHA1) = GetFileSHA1(sExe) Then
            If Not PauseProcess(lPID) Then
                Logg "Warning: ProcessKillIfSHA1Match " & sCmd & " (suspend failed)"
            End If
            If Not ProcessKillByPID(lPID) Then
                Logg "Failed: ProcessKillIfSHA1Match " & sCmd & " (operation failed)"
                PauseProcess lPID, False
            Else
                Logg "Success: ProcessKillIfSHA1Match " & sCmd & " (matched process " & sExe & ")"
            End If
            If Not bKillAll Then Exit For
        End If
    Next i
End Sub

Public Sub ProcessKillIfMD2Match(sCmd$)
    'ProcessKillIfMD2Match <file/mask>|<MD2>|[0|1]
    Dim sProcess$, sMD2$, bKillAll As Boolean, sArgs$()
    sArgs = Split(sCmd, "|")
    Select Case UBound(sArgs)
        Case 1
            sProcess = sArgs(0)
            sMD2 = sArgs(1)
        Case 2
            sProcess = sArgs(0)
            sMD2 = sArgs(1)
            bKillAll = CBool(Val(sArgs(2)))
        Case Else: Exit Sub
    End Select
    
    Dim sMatches$(), lPID&, sExe$, i&
    sMatches = Split(GetMatchingProcesses(sProcess), "|")
    For i = 0 To UBound(sMatches)
        lPID = CLng(Left(sMatches(i), InStr(sMatches(i), ",") - 1))
        sExe = Mid(sMatches(i), InStr(sMatches(i), ",") + 1)
        If UCase(sMD2) = GetFileMD2(sExe) Then
            If Not PauseProcess(lPID) Then
                Logg "Warning: ProcessKillIfMD2Match " & sCmd & " (suspend failed)"
            End If
            If Not ProcessKillByPID(lPID) Then
                Logg "Failed: ProcessKillIfMD2Match " & sCmd & " (operation failed)"
                PauseProcess lPID, False
            Else
                Logg "Success: ProcessKillIfMD2Match " & sCmd & " (matched process " & sExe & ")"
            End If
            If Not bKillAll Then Exit For
        End If
    Next i
End Sub

Public Sub ProcessKillIfMD4Match(sCmd$)
    'ProcessKillIfMD4Match <file/mask>|<MD4>|[0|1]
    Dim sProcess$, sMD4$, bKillAll As Boolean, sArgs$()
    sArgs = Split(sCmd, "|")
    Select Case UBound(sArgs)
        Case 1
            sProcess = sArgs(0)
            sMD4 = sArgs(1)
        Case 2
            sProcess = sArgs(0)
            sMD4 = sArgs(1)
            bKillAll = CBool(Val(sArgs(2)))
        Case Else: Exit Sub
    End Select
    
    Dim sMatches$(), lPID&, sExe$, i&
    sMatches = Split(GetMatchingProcesses(sProcess), "|")
    For i = 0 To UBound(sMatches)
        lPID = CLng(Left(sMatches(i), InStr(sMatches(i), ",") - 1))
        sExe = Mid(sMatches(i), InStr(sMatches(i), ",") + 1)
        If UCase(sMD4) = GetFileMD4(sExe) Then
            If Not PauseProcess(lPID) Then
                Logg "Warning: ProcessKillIfMD4Match " & sCmd & " (suspend failed)"
            End If
            If Not ProcessKillByPID(lPID) Then
                Logg "Failed: ProcessKillIfMD4Match " & sCmd & " (operation failed)"
                PauseProcess lPID, False
            Else
                Logg "Success: ProcessKillIfMD4Match " & sCmd & " (matched process " & sExe & ")"
            End If
            If Not bKillAll Then Exit For
        End If
    Next i
End Sub

Public Sub ProcessKillIfCRC32Match(sCmd$)
    'ProcessKillIfCRC32Match <process/mask>|<crc32>|[0|1]
    Dim sProcess$, sCRC32$, bKillAll As Boolean, sArgs$()
    sArgs = Split(sCmd, "|")
    Select Case UBound(sArgs)
        Case 1
            sProcess = sArgs(0)
            sCRC32 = sArgs(1)
        Case 2
            sProcess = sArgs(0)
            sCRC32 = sArgs(1)
            bKillAll = CBool(Val(sArgs(2)))
        Case Else: Exit Sub
    End Select
    
    Dim sMatches$(), lPID&, sExe$, i&
    sMatches = Split(GetMatchingProcesses(sProcess), "|")
    For i = 0 To UBound(sMatches)
        lPID = CLng(Left(sMatches(i), InStr(sMatches(i), ",") - 1))
        sExe = Mid(sMatches(i), InStr(sMatches(i), ",") + 1)
        If UCase(sCRC32) = modCRC32.GetFileCRC32(sExe) Then
            If Not PauseProcess(lPID) Then
                Logg "Warning: ProcessKillIfCRC32Match " & sCmd & " (suspend failed)"
            End If
            If Not ProcessKillByPID(lPID) Then
                Logg "Failed: ProcessKillIfCRC32Match " & sCmd & " (operation failed)"
                PauseProcess lPID, False
            Else
                Logg "Success: ProcessKillIfCRC32Match " & sCmd & " (matched process " & sExe & ")"
            End If
            If Not bKillAll Then Exit For
        End If
    Next i
End Sub

Public Sub ProcessKillIfContainsText(sCmd$)
    'ProcessKillIfContainsText <process/mask>|<text>|[0|1]
    Dim sProcess$, sText$, sFileContents$, bKillAll As Boolean, sArgs$()
    sArgs = Split(sCmd, "|")
    Select Case UBound(sArgs)
        Case 1
            sProcess = sArgs(0)
            sText = sArgs(1)
        Case 2
            sProcess = sArgs(0)
            sText = sArgs(1)
            bKillAll = CBool(Val(sArgs(2)))
        Case Else: Exit Sub
    End Select
    
    Dim sMatches$(), i&, lPID&, sExe$
    sMatches = Split(GetMatchingProcesses(sProcess), "|")
    For i = 0 To UBound(sMatches)
        lPID = CLng(Left(sMatches(i), InStr(sMatches(i), ",") - 1))
        sExe = Mid(sMatches(i), InStr(sMatches(i), ",") + 1)
        sFileContents = InputFile(sExe)
        If InStr(sFileContents, sText) > 0 Then
            If Not PauseProcess(lPID) Then
                Logg "Warning: ProcessKillIfContainsText " & sCmd & " (suspend failed)"
            End If
            If Not ProcessKillByPID(lPID) Then
                Logg "Failed: ProcessKillIfContainsText " & sCmd & " (operation failed)"
                PauseProcess lPID, False
            Else
                Logg "Success: ProcessKillIfContainsText " & sCmd & " (matched process " & sExe & ")"
            End If
            If Not bKillAll Then Exit For
        End If
    Next i
End Sub

Public Sub ProcessKillIfContainsHex(sCmd$)
    'ProcessKillIfContainsHex <process/mask>|<csv hex>|[0|1]
    Dim sProcess$, sHex$, i&, sHexArray$(), sFileContents$, bKillAll As Boolean, sArgs$()
    sArgs = Split(sCmd, "|")
    Select Case UBound(sArgs)
        Case 1
            sProcess = sArgs(0)
            sHex = sArgs(1)
        Case 2
            sProcess = sArgs(0)
            sHex = sArgs(1)
            bKillAll = CBool(Val(sArgs(2)))
        Case Else: Exit Sub
    End Select
    
    sHexArray = Split(sHex, ",")
    sHex = vbNullString
    For i = 0 To UBound(sHexArray)
        sHex = sHex & Chr(Val("&H" & sHexArray(i)))
    Next i
    Dim sMatches$(), lPID&, sExe$
    sMatches = Split(GetMatchingProcesses(sProcess), "|")
    For i = 0 To UBound(sMatches)
        lPID = CLng(Left(sMatches(i), InStr(sMatches(i), ",") - 1))
        sExe = Mid(sMatches(i), InStr(sMatches(i), ",") + 1)
        sFileContents = InputFile(sExe)
        If InStr(sFileContents, sHex) > 0 Then
            If Not PauseProcess(lPID) Then
                Logg "Warning: ProcessKillIfContainsHex " & sCmd & " (suspend failed)"
            End If
            If Not ProcessKillByPID(lPID) Then
                Logg "Failed: ProcessKillIfContainsHex " & sCmd & " (operation failed)"
                PauseProcess lPID, False
            Else
                Logg "Success: ProcessKillIfContainsHex " & sCmd & " (matched process " & sExe & ")"
            End If
            If Not bKillAll Then Exit For
        End If
    Next i
End Sub

Public Sub ProcessSuspendIfMD5Match(sCmd$)
    'ProcessSuspendIfMD5Match <file/mask>|<md5>|[0|1]
    Dim sProcess$, sMD5$, bSuspendAll As Boolean, sArgs$()
    sArgs = Split(sCmd, "|")
    Select Case UBound(sArgs)
        Case 1
            sProcess = sArgs(0)
            sMD5 = sArgs(1)
        Case 2
            sProcess = sArgs(0)
            sMD5 = sArgs(1)
            bSuspendAll = CBool(Val(sArgs(2)))
        Case Else: Exit Sub
    End Select
    
    Dim sMatches$(), lPID&, sExe$, i&
    sMatches = Split(GetMatchingProcesses(sProcess), "|")
    For i = 0 To UBound(sMatches)
        lPID = CLng(Left(sMatches(i), InStr(sMatches(i), ",") - 1))
        sExe = Mid(sMatches(i), InStr(sMatches(i), ",") + 1)
        If UCase(sMD5) = GetFileMD5(sExe) Then
            If Not PauseProcess(lPID) Then
                Logg "Failed: ProcessSuspendIfMD5Match " & sCmd & " (operation failed)"
            Else
                Logg "Success: ProcessSuspendIfMD5Match " & sCmd & " (matched process " & sExe & ")"
            End If
            If Not bSuspendAll Then Exit For
        End If
    Next i
End Sub

Public Sub ProcessSuspendIfSHA1Match(sCmd$)
    'ProcessSuspendIfSHA1Match <file/mask>|<SHA1>|[0|1]
    Dim sProcess$, sSHA1$, bSuspendAll As Boolean, sArgs$()
    sArgs = Split(sCmd, "|")
    Select Case UBound(sArgs)
        Case 1
            sProcess = sArgs(0)
            sSHA1 = sArgs(1)
        Case 2
            sProcess = sArgs(0)
            sSHA1 = sArgs(1)
            bSuspendAll = CBool(Val(sArgs(2)))
        Case Else: Exit Sub
    End Select
    
    Dim sMatches$(), lPID&, sExe$, i&
    sMatches = Split(GetMatchingProcesses(sProcess), "|")
    For i = 0 To UBound(sMatches)
        lPID = CLng(Left(sMatches(i), InStr(sMatches(i), ",") - 1))
        sExe = Mid(sMatches(i), InStr(sMatches(i), ",") + 1)
        If UCase(sSHA1) = GetFileSHA1(sExe) Then
            If Not PauseProcess(lPID) Then
                Logg "Failed: ProcessSuspendIfSHA1Match " & sCmd & " (operation failed)"
            Else
                Logg "Success: ProcessSuspendIfSHA1Match " & sCmd & " (matched process " & sExe & ")"
            End If
            If Not bSuspendAll Then Exit For
        End If
    Next i
End Sub

Public Sub ProcessSuspendIfMD2Match(sCmd$)
    'ProcessSuspendIfMD2Match <file/mask>|<MD2>|[0|1]
    Dim sProcess$, sMD2$, bSuspendAll As Boolean, sArgs$()
    sArgs = Split(sCmd, "|")
    Select Case UBound(sArgs)
        Case 1
            sProcess = sArgs(0)
            sMD2 = sArgs(1)
        Case 2
            sProcess = sArgs(0)
            sMD2 = sArgs(1)
            bSuspendAll = CBool(Val(sArgs(2)))
        Case Else: Exit Sub
    End Select
    
    Dim sMatches$(), lPID&, sExe$, i&
    sMatches = Split(GetMatchingProcesses(sProcess), "|")
    For i = 0 To UBound(sMatches)
        lPID = CLng(Left(sMatches(i), InStr(sMatches(i), ",") - 1))
        sExe = Mid(sMatches(i), InStr(sMatches(i), ",") + 1)
        If UCase(sMD2) = GetFileMD2(sExe) Then
            If Not PauseProcess(lPID) Then
                Logg "Failed: ProcessSuspendIfMD2Match " & sCmd & " (operation failed)"
            Else
                Logg "Success: ProcessSuspendIfMD2Match " & sCmd & " (matched process " & sExe & ")"
            End If
            If Not bSuspendAll Then Exit For
        End If
    Next i
End Sub

Public Sub ProcessSuspendIfMD4Match(sCmd$)
    'ProcessSuspendIfMD4Match <file/mask>|<MD4>|[0|1]
    Dim sProcess$, sMD4$, bSuspendAll As Boolean, sArgs$()
    sArgs = Split(sCmd, "|")
    Select Case UBound(sArgs)
        Case 1
            sProcess = sArgs(0)
            sMD4 = sArgs(1)
        Case 2
            sProcess = sArgs(0)
            sMD4 = sArgs(1)
            bSuspendAll = CBool(Val(sArgs(2)))
        Case Else: Exit Sub
    End Select
    
    Dim sMatches$(), lPID&, sExe$, i&
    sMatches = Split(GetMatchingProcesses(sProcess), "|")
    For i = 0 To UBound(sMatches)
        lPID = CLng(Left(sMatches(i), InStr(sMatches(i), ",") - 1))
        sExe = Mid(sMatches(i), InStr(sMatches(i), ",") + 1)
        If UCase(sMD4) = GetFileMD4(sExe) Then
            If Not PauseProcess(lPID) Then
                Logg "Failed: ProcessSuspendIfMD4Match " & sCmd & " (operation failed)"
            Else
                Logg "Success: ProcessSuspendIfMD4Match " & sCmd & " (matched process " & sExe & ")"
            End If
            If Not bSuspendAll Then Exit For
        End If
    Next i
End Sub

Public Sub ProcessSuspendIfCRC32Match(sCmd$)
    'ProcessSuspendIfCRC32Match <process/mask>|<crc32>|[0|1]
    Dim sProcess$, sCRC32$, bSuspendAll As Boolean, sArgs$()
    sArgs = Split(sCmd, "|")
    Select Case UBound(sArgs)
        Case 1
            sProcess = sArgs(0)
            sCRC32 = sArgs(1)
        Case 2
            sProcess = sArgs(0)
            sCRC32 = sArgs(1)
            bSuspendAll = CBool(Val(sArgs(2)))
        Case Else: Exit Sub
    End Select
    
    Dim sMatches$(), lPID&, sExe$, i&
    sMatches = Split(GetMatchingProcesses(sProcess), "|")
    For i = 0 To UBound(sMatches)
        lPID = CLng(Left(sMatches(i), InStr(sMatches(i), ",") - 1))
        sExe = Mid(sMatches(i), InStr(sMatches(i), ",") + 1)
        If UCase(sCRC32) = modCRC32.GetFileCRC32(sExe) Then
            If Not PauseProcess(lPID) Then
                Logg "Failed: ProcessSuspendIfCRC32Match " & sCmd & " (operation failed)"
            Else
                Logg "Success: ProcessSuspendIfCRC32Match " & sCmd & " (matched process " & sExe & ")"
            End If
            If Not bSuspendAll Then Exit For
        End If
    Next i
End Sub

Public Sub ProcessSuspendIfContainsText(sCmd$)
    'ProcessSuspendIfContainsText <process/mask>|<text>|[0|1]
    Dim sProcess$, sText$, sFileContents$, bSuspendAll As Boolean, sArgs$()
    sArgs = Split(sCmd, "|")
    Select Case UBound(sArgs)
        Case 1
            sProcess = sArgs(0)
            sText = sArgs(1)
        Case 2
            sProcess = sArgs(0)
            sText = sArgs(1)
            bSuspendAll = CBool(Val(sArgs(2)))
        Case Else: Exit Sub
    End Select
    
    Dim sMatches$(), i&, lPID&, sExe$
    sMatches = Split(GetMatchingProcesses(sProcess), "|")
    For i = 0 To UBound(sMatches)
        lPID = CLng(Left(sMatches(i), InStr(sMatches(i), ",") - 1))
        sExe = Mid(sMatches(i), InStr(sMatches(i), ",") + 1)
        sFileContents = InputFile(sExe)
        If InStr(sFileContents, sText) > 0 Then
            If Not PauseProcess(lPID) Then
                Logg "Failed: ProcessSuspendIfContainsText " & sCmd & " (operation failed)"
            Else
                Logg "Success: ProcessSuspendIfContainsText " & sCmd & " (matched process " & sExe & ")"
            End If
            If Not bSuspendAll Then Exit For
        End If
    Next i
End Sub

Public Sub ProcessSuspendIfContainsHex(sCmd$)
    'ProcessSuspendIfContainsHex <process/mask>|<csv hex>|[0|1]
    Dim sProcess$, sHex$, i&, sHexArray$(), sFileContents$, bSuspendAll As Boolean, sArgs$()
    sArgs = Split(sCmd, "|")
    Select Case UBound(sArgs)
        Case 1
            sProcess = sArgs(0)
            sHex = sArgs(1)
        Case 2
            sProcess = sArgs(0)
            sHex = sArgs(1)
            bSuspendAll = CBool(Val(sArgs(2)))
        Case Else: Exit Sub
    End Select
    
    sHexArray = Split(sHex, ",")
    sHex = vbNullString
    For i = 0 To UBound(sHexArray)
        sHex = sHex & Chr(Val("&H" & sHexArray(i)))
    Next i
    Dim sMatches$(), lPID&, sExe$
    sMatches = Split(GetMatchingProcesses(sProcess), "|")
    For i = 0 To UBound(sMatches)
        lPID = CLng(Left(sMatches(i), InStr(sMatches(i), ",") - 1))
        sExe = Mid(sMatches(i), InStr(sMatches(i), ",") + 1)
        sFileContents = InputFile(sExe)
        If InStr(sFileContents, sHex) > 0 Then
            If Not PauseProcess(lPID) Then
                Logg "Failed: ProcessSuspendIfContainsHex " & sCmd & " (operation failed)"
            Else
                Logg "Success: ProcessSuspendIfContainsHex " & sCmd & " (matched process " & sExe & ")"
            End If
            If Not bSuspendAll Then Exit For
        End If
    Next i
End Sub

Public Sub ProcessResumeIfMD5Match(sCmd$)
    'ProcessResumeIfMD5Match <file/mask>|<md5>|[0|1]
    Dim sProcess$, sMD5$, bResumeAll As Boolean, sArgs$()
    sArgs = Split(sCmd, "|")
    Select Case UBound(sArgs)
        Case 1
            sProcess = sArgs(0)
            sMD5 = sArgs(1)
        Case 2
            sProcess = sArgs(0)
            sMD5 = sArgs(1)
            bResumeAll = CBool(Val(sArgs(2)))
        Case Else: Exit Sub
    End Select
    
    Dim sMatches$(), lPID&, sExe$, i&
    sMatches = Split(GetMatchingProcesses(sProcess), "|")
    For i = 0 To UBound(sMatches)
        lPID = CLng(Left(sMatches(i), InStr(sMatches(i), ",") - 1))
        sExe = Mid(sMatches(i), InStr(sMatches(i), ",") + 1)
        If UCase(sMD5) = GetFileMD5(sExe) Then
            If Not PauseProcess(lPID, False) Then
                Logg "Failed: ProcessResumeIfMD5Match " & sCmd & " (operation failed)"
            Else
                Logg "Success: ProcessResumeIfMD5Match " & sCmd & " (matched process " & sExe & ")"
            End If
            If Not bResumeAll Then Exit For
        End If
    Next i
End Sub

Public Sub ProcessResumeIfSHA1Match(sCmd$)
    'ProcessResumeIfSHA1Match <file/mask>|<SHA1>|[0|1]
    Dim sProcess$, sSHA1$, bResumeAll As Boolean, sArgs$()
    sArgs = Split(sCmd, "|")
    Select Case UBound(sArgs)
        Case 1
            sProcess = sArgs(0)
            sSHA1 = sArgs(1)
        Case 2
            sProcess = sArgs(0)
            sSHA1 = sArgs(1)
            bResumeAll = CBool(Val(sArgs(2)))
        Case Else: Exit Sub
    End Select
    
    Dim sMatches$(), lPID&, sExe$, i&
    sMatches = Split(GetMatchingProcesses(sProcess), "|")
    For i = 0 To UBound(sMatches)
        lPID = CLng(Left(sMatches(i), InStr(sMatches(i), ",") - 1))
        sExe = Mid(sMatches(i), InStr(sMatches(i), ",") + 1)
        If UCase(sSHA1) = GetFileSHA1(sExe) Then
            If Not PauseProcess(lPID, False) Then
                Logg "Failed: ProcessResumeIfSHA1Match " & sCmd & " (operation failed)"
            Else
                Logg "Success: ProcessResumeIfSHA1Match " & sCmd & " (matched process " & sExe & ")"
            End If
            If Not bResumeAll Then Exit For
        End If
    Next i
End Sub

Public Sub ProcessResumeIfMD2Match(sCmd$)
    'ProcessResumeIfMD2Match <file/mask>|<MD2>|[0|1]
    Dim sProcess$, sMD2$, bResumeAll As Boolean, sArgs$()
    sArgs = Split(sCmd, "|")
    Select Case UBound(sArgs)
        Case 1
            sProcess = sArgs(0)
            sMD2 = sArgs(1)
        Case 2
            sProcess = sArgs(0)
            sMD2 = sArgs(1)
            bResumeAll = CBool(Val(sArgs(2)))
        Case Else: Exit Sub
    End Select
    
    Dim sMatches$(), lPID&, sExe$, i&
    sMatches = Split(GetMatchingProcesses(sProcess), "|")
    For i = 0 To UBound(sMatches)
        lPID = CLng(Left(sMatches(i), InStr(sMatches(i), ",") - 1))
        sExe = Mid(sMatches(i), InStr(sMatches(i), ",") + 1)
        If UCase(sMD2) = GetFileMD2(sExe) Then
            If Not PauseProcess(lPID, False) Then
                Logg "Failed: ProcessResumeIfMD2Match " & sCmd & " (operation failed)"
            Else
                Logg "Success: ProcessResumeIfMD2Match " & sCmd & " (matched process " & sExe & ")"
            End If
            If Not bResumeAll Then Exit For
        End If
    Next i
End Sub

Public Sub ProcessResumeIfMD4Match(sCmd$)
    'ProcessResumeIfMD4Match <file/mask>|<MD4>|[0|1]
    Dim sProcess$, sMD4$, bResumeAll As Boolean, sArgs$()
    sArgs = Split(sCmd, "|")
    Select Case UBound(sArgs)
        Case 1
            sProcess = sArgs(0)
            sMD4 = sArgs(1)
        Case 2
            sProcess = sArgs(0)
            sMD4 = sArgs(1)
            bResumeAll = CBool(Val(sArgs(2)))
        Case Else: Exit Sub
    End Select
    
    Dim sMatches$(), lPID&, sExe$, i&
    sMatches = Split(GetMatchingProcesses(sProcess), "|")
    For i = 0 To UBound(sMatches)
        lPID = CLng(Left(sMatches(i), InStr(sMatches(i), ",") - 1))
        sExe = Mid(sMatches(i), InStr(sMatches(i), ",") + 1)
        If UCase(sMD4) = GetFileMD4(sExe) Then
            If Not PauseProcess(lPID, False) Then
                Logg "Failed: ProcessResumeIfMD4Match " & sCmd & " (operation failed)"
            Else
                Logg "Success: ProcessResumeIfMD4Match " & sCmd & " (matched process " & sExe & ")"
            End If
            If Not bResumeAll Then Exit For
        End If
    Next i
End Sub

Public Sub ProcessResumeIfCRC32Match(sCmd$)
    'ProcessResumeIfCRC32Match <process/mask>|<crc32>|[0|1]
    Dim sProcess$, sCRC32$, bResumeAll As Boolean, sArgs$()
    sArgs = Split(sCmd, "|")
    Select Case UBound(sArgs)
        Case 1
            sProcess = sArgs(0)
            sCRC32 = sArgs(1)
        Case 2
            sProcess = sArgs(0)
            sCRC32 = sArgs(1)
            bResumeAll = CBool(Val(sArgs(2)))
        Case Else: Exit Sub
    End Select
    
    Dim sMatches$(), lPID&, sExe$, i&
    sMatches = Split(GetMatchingProcesses(sProcess), "|")
    For i = 0 To UBound(sMatches)
        lPID = CLng(Left(sMatches(i), InStr(sMatches(i), ",") - 1))
        sExe = Mid(sMatches(i), InStr(sMatches(i), ",") + 1)
        If UCase(sCRC32) = modCRC32.GetFileCRC32(sExe) Then
            If Not PauseProcess(lPID, False) Then
                Logg "Failed: ProcessResumeIfCRC32Match " & sCmd & " (operation failed)"
            Else
                Logg "Success: ProcessResumeIfCRC32Match " & sCmd & " (matched process " & sExe & ")"
            End If
            If Not bResumeAll Then Exit For
        End If
    Next i
End Sub

Public Sub ProcessResumeIfContainsText(sCmd$)
    'ProcessResumeIfContainsText <process/mask>|<text>|[0|1]
    Dim sProcess$, sText$, sFileContents$, bResumeAll As Boolean, sArgs$()
    sArgs = Split(sCmd, "|")
    Select Case UBound(sArgs)
        Case 1
            sProcess = sArgs(0)
            sText = sArgs(1)
        Case 2
            sProcess = sArgs(0)
            sText = sArgs(1)
            bResumeAll = CBool(Val(sArgs(2)))
        Case Else: Exit Sub
    End Select
    
    Dim sMatches$(), i&, lPID&, sExe$
    sMatches = Split(GetMatchingProcesses(sProcess), "|")
    For i = 0 To UBound(sMatches)
        lPID = CLng(Left(sMatches(i), InStr(sMatches(i), ",") - 1))
        sExe = Mid(sMatches(i), InStr(sMatches(i), ",") + 1)
        sFileContents = InputFile(sExe)
        If InStr(sFileContents, sText) > 0 Then
            If Not PauseProcess(lPID, False) Then
                Logg "Failed: ProcessResumeIfContainsText " & sCmd & " (operation failed)"
            Else
                Logg "Success: ProcessResumeIfContainsText " & sCmd & " (matched process " & sExe & ")"
            End If
            If Not bResumeAll Then Exit For
        End If
    Next i
End Sub

Public Sub ProcessResumeIfContainsHex(sCmd$)
    'ProcessResumeIfContainsHex <process/mask>|<csv hex>|[0|1]
    Dim sProcess$, sHex$, i&, sHexArray$(), sFileContents$, bResumeAll As Boolean, sArgs$()
    sArgs = Split(sCmd, "|")
    Select Case UBound(sArgs)
        Case 1
            sProcess = sArgs(0)
            sHex = sArgs(1)
        Case 2
            sProcess = sArgs(0)
            sHex = sArgs(1)
            bResumeAll = CBool(Val(sArgs(2)))
        Case Else: Exit Sub
    End Select
    
    sHexArray = Split(sHex, ",")
    sHex = vbNullString
    For i = 0 To UBound(sHexArray)
        sHex = sHex & Chr(Val("&H" & sHexArray(i)))
    Next i
    Dim sMatches$(), lPID&, sExe$
    sMatches = Split(GetMatchingProcesses(sProcess), "|")
    For i = 0 To UBound(sMatches)
        lPID = CLng(Left(sMatches(i), InStr(sMatches(i), ",") - 1))
        sExe = Mid(sMatches(i), InStr(sMatches(i), ",") + 1)
        sFileContents = InputFile(sExe)
        If InStr(sFileContents, sHex) > 0 Then
            If Not PauseProcess(lPID, False) Then
                Logg "Failed: ProcessResumeIfContainsHex " & sCmd & " (operation failed)"
            Else
                Logg "Success: ProcessResumeIfContainsHex " & sCmd & " (matched process " & sExe & ")"
            End If
            If Not bResumeAll Then Exit For
        End If
    Next i
End Sub
