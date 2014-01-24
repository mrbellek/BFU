Attribute VB_Name = "modFiles"
Option Explicit
'deleting, renaming, clearing, moving files/folders

Private Declare Function SHFileOperation Lib "shell32.dll" Alias "SHFileOperationA" (lpFileOp As SHFILEOPSTRUCT) As Long
Private Declare Function SHFileExists Lib "shell32" Alias "#45" (ByVal szPath As String) As Long
Private Declare Function MoveFileEx Lib "kernel32" Alias "MoveFileExA" (ByVal lpExistingFileName As String, ByVal lpNewFileName As String, ByVal dwFlags As Long) As Long
Private Declare Function GetFileAttributes Lib "kernel32" Alias "GetFileAttributesA" (ByVal lpFileName As String) As Long

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

Private Type SHFILEOPSTRUCT
    hWnd As Long
    wFunc As Long
    pFrom As String
    pTo As String
    fFlags As Integer
    fAnyOperationsAborted As Long
    hNameMappings As Long
    lpszProgressTitle As String
End Type

Private Const FO_MOVE = &H1
'Private Const FO_COPY = &H2
Private Const FO_DELETE = &H3
Private Const FO_RENAME = &H4
Private Const FOF_ALLOWUNDO = &H40
Private Const FOF_NOCONFIRMATION = &H10
Private Const FOF_NOERRORUI = &H400
Private Const FOF_SILENT = &H4

Private Const MOVEFILE_DELAY_UNTIL_REBOOT = &H4

Private Const FILE_ATTRIBUTE_READONLY = &H1
Private Const FILE_ATTRIBUTE_HIDDEN = &H2
Private Const FILE_ATTRIBUTE_SYSTEM = &H4
Private Const FILE_ATTRIBUTE_DIRECTORY = &H10
Private Const FILE_ATTRIBUTE_ARCHIVE = &H20
'Private Const FILE_ATTRIBUTE_NORMAL = &H80
'Private Const FILE_ATTRIBUTE_TEMPORARY = &H100
Private Const FILE_ATTRIBUTE_COMPRESSED = &H800

Public Function FileExists(sFile$) As Boolean
    'internal function
    If bIsWinNT Then
        FileExists = IIf(SHFileExists(StrConv(sFile, vbUnicode)) = 1, True, False)
    Else
        FileExists = IIf(SHFileExists(sFile) = 1, True, False)
    End If
End Function

Public Function FolderExists(sFolder$) As Boolean
    'internal function
    FolderExists = False
    If bIsWinNT Then
        If SHFileExists(StrConv(sFolder, vbUnicode)) Then
            If GetFileAttributes(sFolder) And FILE_ATTRIBUTE_DIRECTORY Then FolderExists = True
        End If
    Else
        If SHFileExists(sFolder) Then
            If GetFileAttributes(sFolder) And FILE_ATTRIBUTE_DIRECTORY Then FolderExists = True
        End If
    End If
End Function

Public Sub FileDelete(sCmd$)
    'FileDelete <file/filemask>
    Dim uSFOS As SHFILEOPSTRUCT, sFiles$
    sFiles = sCmd
    If sFiles = vbNullString Then Exit Sub
    If InStr(sFiles, "*") = 0 And InStr(sFiles, "?") = 0 Then
        'to prevent FileDeleteOnReboot from trying to delete all
        'non-existent files
        If Not FileExists(sFiles) Then Exit Sub
    End If
    
    With uSFOS
        .hWnd = frmMain.hWnd
        .wFunc = FO_DELETE
        .pFrom = sFiles
        .fFlags = FOF_NOCONFIRMATION Or FOF_SILENT Or FOF_NOERRORUI
        If bUseRecycleBin Then .fFlags = .fFlags Or FOF_ALLOWUNDO
    End With
    If SHFileOperation(uSFOS) <> 0 Then
        Logg "Failed: FileDelete " & sCmd & " (operation failed)"
        If bUseDeleteOnReboot Then FileDeleteOnReboot sCmd
    Else
        Logg "Success: FileDelete " & sCmd
    End If
End Sub

Public Sub FileRename(sCmd$)
    'FileRename <file>|<newfile>
    Dim uSFOS As SHFILEOPSTRUCT, sArgs$(), sFile$, sNewFile$
    sArgs = Split(sCmd, "|")
    If UBound(sArgs) <> 1 Then Exit Sub
    sFile = sArgs(0)
    sNewFile = sArgs(1)
        
    If Not FileExists(sFile) Then
        Logg "Failed: FileRename " & sCmd & " (source file not found)"
        Exit Sub
    End If
    If FileExists(sNewFile) Then
        Logg "Failed: FileRename " & sCmd & " (target file already exists)"
        Exit Sub
    End If
    
    With uSFOS
        .hWnd = frmMain.hWnd
        .wFunc = FO_RENAME
        .pFrom = sFile
        .pTo = sNewFile
        .fFlags = FOF_NOCONFIRMATION Or FOF_SILENT Or FOF_NOERRORUI
    End With
    If SHFileOperation(uSFOS) <> 0 Then
        Logg "Failed: FileRename " & sCmd & " (operation failed)"
    Else
        Logg "Success: FileRename " & sCmd
    End If
End Sub

Public Sub FileMove(sCmd$)
    'FileMove <file>|<folder>
    Dim uSFOS As SHFILEOPSTRUCT, sArgs$(), sFile$, sFolder$
    sArgs = Split(sCmd, "|")
    If UBound(sArgs) <> 1 Then Exit Sub
    sFile = sArgs(0)
    sFolder = sArgs(1)
        
    If Not FileExists(sFile) Then
        Logg "Failed: FileMove " & sCmd & " (source file not found)"
        Exit Sub
    End If
    If Not FileExists(sFolder) Then
        Logg "Failed: FileMove " & sCmd & " (target folder not found)"
        Exit Sub
    End If
    
    With uSFOS
        .hWnd = frmMain.hWnd
        .wFunc = FO_MOVE
        .pFrom = sFile
        .pTo = sFolder
        .fFlags = FOF_NOCONFIRMATION Or FOF_SILENT Or FOF_NOERRORUI
    End With
    If SHFileOperation(uSFOS) <> 0 Then
        Logg "Failed: FileMove " & sCmd & " (operation failed)"
    Else
        Logg "Success: FileMove " & sCmd
    End If
End Sub

Public Sub FileCreate(sCmd$)
    'FileCreate <file>
    Dim sFile$
    sFile = sCmd
    If Not FileExists(sFile) Then
        On Error Resume Next
        Open sFile For Output As #1
        Close #1
    Else
        Logg "Failed: FileCreate " & sCmd & " (file already exists)"
    End If
    If Not FileExists(sFile) Then
        Logg "Failed: FileCreate " & sCmd & " (file write failed)"
    Else
        Logg "Success: FileCreate " & sCmd
    End If
End Sub

Public Sub FileClear(sCmd$)
    'FileClear <file>
    Dim sFile$
    sFile = sCmd
    If FileExists(sFile) Then
        On Error Resume Next
        Open sFile For Output As #1
        Close #1
        If Err Then
            Logg "Failed: FileClear " & sCmd & " (file write failed)"
        Else
            Logg "Success: FileClear " & sCmd
        End If
    Else
        Logg "Failed: FileClear " & sCmd & " (file not found)"
    End If
End Sub

Public Sub FolderCreate(sCmd$)
    'FolderCreate <folder>
    Dim sFolder$
    sFolder = sCmd
    If Not FileExists(sFolder) Then
        On Error Resume Next
        MkDir sFolder
        If Err Then
            Logg "Failed: FolderCreate " & sCmd & " (file write failed)"
        Else
            Logg "Success: FolderCreate " & sCmd
        End If
    Else
        Logg "Failed: FolderCreate " & sCmd & " (folder already exists)"
    End If
End Sub

Public Sub FolderRename(sCmd$)
    'FolderRename <folder>|<newfolder>
    Dim uSFOS As SHFILEOPSTRUCT, sArgs$(), sFolder$, sNewFolder$
    sArgs = Split(sCmd, "|")
    If UBound(sArgs) <> 1 Then Exit Sub
    sFolder = sArgs(0)
    sNewFolder = sArgs(1)
        
    If Not FileExists(sFolder) Then
        Logg "Failed: FolderRename " & sCmd & " (source folder not found)"
        Exit Sub
    End If
    If FileExists(sNewFolder) Then
        Logg "Failed: FolderRename " & sCmd & " (target folder already exists)"
        Exit Sub
    End If
    
    With uSFOS
        .hWnd = frmMain.hWnd
        .wFunc = FO_RENAME
        .pFrom = sFolder
        .pTo = sNewFolder
        .fFlags = FOF_NOCONFIRMATION Or FOF_SILENT Or FOF_NOERRORUI
    End With
    If SHFileOperation(uSFOS) <> 0 Then
        Logg "Failed: FolderRename " & sCmd & " (operation failed)"
    Else
        Logg "Success: FolderRename " & sCmd
    End If
End Sub

Public Sub FolderDelete(sCmd$)
    'FolderDelete <folder>
    Dim uSFOS As SHFILEOPSTRUCT, sFolder$
    sFolder = sCmd
    If Not FileExists(sFolder) Then
        Logg "Failed: FolderDelete " & sCmd & " (folder not found)"
        Exit Sub
    End If
    With uSFOS
        .hWnd = frmMain.hWnd
        .wFunc = FO_DELETE
        .pFrom = sFolder
        .fFlags = FOF_NOCONFIRMATION Or FOF_SILENT Or FOF_NOERRORUI
        If bUseRecycleBin Then .fFlags = .fFlags Or FOF_ALLOWUNDO
    End With
    If SHFileOperation(uSFOS) <> 0 Then
        If bUseDeleteOnReboot Then
            Logg "Failed: FolderDelete " & sCmd & " (clearing folder and deleting on reboot)"
            FolderClear sCmd
            FileDeleteOnReboot sCmd
        Else
            Logg "Failed: FolderDelete " & sCmd & " (operation failed)"
        End If
    Else
        Logg "Success: FolderDelete " & sCmd
    End If
End Sub

Public Sub FolderMove(sCmd$)
    'FolderMove <folder>|<newfolder>
    Dim uSFOS As SHFILEOPSTRUCT, sArgs$(), sFolder$, sNewFolder$
    sArgs = Split(sCmd, "|")
    If UBound(sArgs) <> 1 Then Exit Sub
    sFolder = sArgs(0)
    sNewFolder = sArgs(1)
        
    If Not FileExists(sFolder) Then
        Logg "Failed: FolderMove " & sCmd & " (source folder not found)"
        Exit Sub
    End If
    If Not FileExists(sNewFolder) Then
        Logg "Failed: FolderMove " & sCmd & " (target folder not found)"
        Exit Sub
    End If
    
    With uSFOS
        .hWnd = frmMain.hWnd
        .wFunc = FO_MOVE
        .pFrom = sFolder
        .pTo = sNewFolder
        .fFlags = FOF_NOCONFIRMATION Or FOF_SILENT Or FOF_NOERRORUI
    End With
    If SHFileOperation(uSFOS) <> 0 Then
        Logg "Failed: FolderMove " & sCmd & " (operation failed)"
    Else
        Logg "Success: FolderMove " & sCmd
    End If
End Sub

Public Sub FileSetAttributes(sCmd$)
    'FileSetAttributes <file>|<A/R/H/S>
    'compress flag not supported by VB!
    Dim sArgs$(), sFile$, iAttr%, sAttr$
    sArgs = Split(sCmd, "|")
    If UBound(sArgs) <> 1 Then Exit Sub
    sFile = sArgs(0)
    sAttr = sArgs(1)
    
    If Not FileExists(sFile) Then
        Logg "Failed: FileSetAttributes " & sCmd & " (file not found)"
        Exit Sub
    End If
    
    If InStr(1, sAttr, "A", vbTextCompare) > 0 Then iAttr = iAttr + FILE_ATTRIBUTE_ARCHIVE
    If InStr(1, sAttr, "R", vbTextCompare) > 0 Then iAttr = iAttr + FILE_ATTRIBUTE_READONLY
    If InStr(1, sAttr, "H", vbTextCompare) > 0 Then iAttr = iAttr + FILE_ATTRIBUTE_HIDDEN
    If InStr(1, sAttr, "S", vbTextCompare) > 0 Then iAttr = iAttr + FILE_ATTRIBUTE_SYSTEM
    On Error Resume Next
    SetAttr sFile, iAttr
    DoEvents
    If GetAttr(sFile) <> iAttr Then
        Logg "Failed: FileSetAttributes " & sCmd & " (operation failed)"
    Else
        Logg "Success: FileSetAttributes " & sCmd
    End If
End Sub

Public Sub FolderSetAttributes(sCmd$)
    'FolderSetAttributes <folder>|<A/R/H/S>
    'compress flag not supported by VB!
    Dim sArgs$(), sFolder$, iAttr%, sAttr$
    sArgs = Split(sCmd, "|")
    If UBound(sArgs) <> 1 Then Exit Sub
    sFolder = sArgs(0)
    sAttr = sArgs(1)
    
    If Not FileExists(sFolder) Then
        Logg "Failed: FolderSetAttributes " & sCmd & " (folder not found)"
        Exit Sub
    End If
    
    If InStr(1, sAttr, "A", vbTextCompare) > 0 Then iAttr = iAttr + FILE_ATTRIBUTE_ARCHIVE
    If InStr(1, sAttr, "R", vbTextCompare) > 0 Then iAttr = iAttr + FILE_ATTRIBUTE_READONLY
    If InStr(1, sAttr, "H", vbTextCompare) > 0 Then iAttr = iAttr + FILE_ATTRIBUTE_HIDDEN
    If InStr(1, sAttr, "S", vbTextCompare) > 0 Then iAttr = iAttr + FILE_ATTRIBUTE_SYSTEM
    On Error Resume Next
    SetAttr sFolder, iAttr
    DoEvents
    If GetAttr(sFolder) <> iAttr + FILE_ATTRIBUTE_DIRECTORY Then
        Logg "Failed: FolderSetAttributes " & sCmd & " (operation failed)"
    Else
        Logg "Success: FolderSetAttributes " & sCmd
    End If
End Sub

Public Sub FileDeleteOnReboot(sCmd$)
    'FileDeleteOnReboot <file/filemask>
    Dim sFile$, sWininit$(), i%, sFileMatch$()
    On Error Resume Next
    sFile = sCmd
    If sFile = vbNullString Then Exit Sub
    If InStr(sFile, "?") > 0 Or InStr(sFile, "*") > 0 Then
        'file is a wildcard match - get matching files
        'and perform this cmd on all of them
        sFileMatch = Split(GetMatchingFiles(sFile), "|")
        For i = 0 To UBound(sFileMatch)
            FileDeleteOnReboot sFileMatch(i)
        Next i
        Exit Sub
    End If
    
    If bIsWinNT Then
        If MoveFileEx(sFile, vbNullString, MOVEFILE_DELAY_UNTIL_REBOOT) <> 0 Then
            Logg "Success: FileDeleteOnReboot " & sCmd
            bRebootNeeded = True
        Else
            Logg "Failed: FileDeleteOnReboot " & sCmd & " (operation failed)"
        End If
    Else
        If FileExists(sWinDir & "\wininit.ini") Then
            'file exists, to read it and append our filename
            sWininit = Split(InputFile(sWinDir & "\wininit.ini"), vbCrLf)
            SetAttr sWinDir & "\wininit.ini", vbArchive
            Open sWinDir & "\wininit.ini" For Output As #1
                For i = 0 To UBound(sWininit)
                    If InStr(1, sWininit(i), "[rename]", vbTextCompare) = 1 Then
                        Print #1, sWininit(i)
                        Print #1, "NUL=" & GetDOSFilename(sFile)
                    Else
                        Print #1, sWininit(i)
                    End If
                Next i
            Close #1
            If Err Then
                Logg "Failed: FileDeleteOnReboot " & sCmd & " (file write error)"
            Else
                Logg "Success: FileDeleteOnReboot " & sCmd
            End If
        Else
            'does not exist, create new
            Open sWinDir & "\wininit.ini" For Output As #1
                Print #1, "[rename]"
                Print #1, "NUL=" & GetDOSFilename(sFile)
            Close #1
            If Err Then
                Logg "Failed: FileDeleteOnReboot " & sCmd & " (file write error)"
            Else
                Logg "Success: FileDeleteOnReboot " & sCmd
            End If
        End If
        If Not Err Then bRebootNeeded = True
    End If
End Sub

Public Sub FileWrite(sCmd$)
    'FileWrite <file>|<data>|<overwrite[0|1]>
    Dim sArgs$(), sFile$, sData$, bOverwrite As Boolean, iAttr%
    sArgs = Split(sCmd, "|")
    If UBound(sArgs) <> 1 And UBound(sArgs) <> 2 Then Exit Sub
    sFile = sArgs(0)
    sData = sArgs(1)
    If UBound(sArgs) = 2 Then bOverwrite = CBool(sArgs(2))
    
    sData = Replace(sData, "\n", vbCrLf)
    sData = Replace(sData, "\t", vbTab)
    
    On Error Resume Next
    iAttr = GetAttr(sFile)
    If (iAttr And FILE_ATTRIBUTE_COMPRESSED) Then iAttr = iAttr - FILE_ATTRIBUTE_COMPRESSED
    SetAttr sFile, vbArchive
    If bOverwrite Then
        OutputFile sFile, sData
    Else
        If Not FileExists(sFile) Then
            Logg "Failed: FileWrite " & sCmd & " (target file does not exist)"
            Exit Sub
        End If
        OutputFile sFile, sData, True
    End If
    If Not FileExists(sFile) Then
        Logg "Failed: FileWrite " & sCmd & " (file write error)"
        Close
        Exit Sub
    End If
    SetAttr sFile, iAttr
    Logg "Success: FileWrite " & sCmd
End Sub

Public Sub FileDeleteIfMD5Match(sCmd$)
    'FileDeleteIfMD5Match <file/filemask>|<md5>
    Dim sArgs$(), sFile$, sMD5$
    sArgs = Split(sCmd, "|")
    If UBound(sArgs) <> 1 Then Exit Sub
    sFile = sArgs(0)
    sMD5 = sArgs(1)
    
    If InStr(sFile, "?") > 0 Or InStr(sFile, "*") > 0 Then
        Dim sMatches$(), i&
        sMatches = Split(GetMatchingFiles(sFile), "|")
        For i = 0 To UBound(sMatches)
            FileDeleteIfMD5Match sMatches(i) & "|" & sMD5
        Next i
        Exit Sub
    End If
    If UCase(sMD5) = GetFileMD5(sFile) Then
        'match!
        Logg "Success: FileDeleteIfMD5Match " & sCmd & " (matched " & sFile & ")"
        FileDelete sFile
    End If
End Sub

Public Sub FileDeleteIfSHA1Match(sCmd$)
    'FileDeleteIfSHA1Match <file/filemask>|<sha1>
    Dim sArgs$(), sFile$, sSHA1$
    sArgs = Split(sCmd, "|")
    If UBound(sArgs) <> 1 Then Exit Sub
    sFile = sArgs(0)
    sSHA1 = sArgs(1)
    
    If InStr(sFile, "?") > 0 Or InStr(sFile, "*") > 0 Then
        Dim sMatches$(), i&
        sMatches = Split(GetMatchingFiles(sFile), "|")
        For i = 0 To UBound(sMatches)
            FileDeleteIfSHA1Match sMatches(i) & "|" & sSHA1
        Next i
        Exit Sub
    End If
    If UCase(sSHA1) = GetFileSHA1(sFile) Then
        'match!
        Logg "Success: FileDeleteIfSHA1Match " & sCmd & " (matched " & sFile & ")"
        FileDelete sFile
    End If
End Sub

Public Sub FileDeleteIfMD2Match(sCmd$)
    'FileDeleteIfMD2Match <file/filemask>|<md2>
    Dim sArgs$(), sFile$, sMD2$
    sArgs = Split(sCmd, "|")
    If UBound(sArgs) <> 1 Then Exit Sub
    sFile = sArgs(0)
    sMD2 = sArgs(1)
    
    If InStr(sFile, "?") > 0 Or InStr(sFile, "*") > 0 Then
        Dim sMatches$(), i&
        sMatches = Split(GetMatchingFiles(sFile), "|")
        For i = 0 To UBound(sMatches)
            FileDeleteIfMD2Match sMatches(i) & "|" & sMD2
        Next i
        Exit Sub
    End If
    If UCase(sMD2) = GetFileMD2(sFile) Then
        'match!
        Logg "Success: FileDeleteIfMD2Match " & sCmd & " (matched " & sFile & ")"
        FileDelete sFile
    End If
End Sub

Public Sub FileDeleteIfMD4Match(sCmd$)
    'FileDeleteIfMD4Match <file/filemask>|<md4>
    Dim sArgs$(), sFile$, sMD4$
    sArgs = Split(sCmd, "|")
    If UBound(sArgs) <> 1 Then Exit Sub
    sFile = sArgs(0)
    sMD4 = sArgs(1)
    
    If InStr(sFile, "?") > 0 Or InStr(sFile, "*") > 0 Then
        Dim sMatches$(), i&
        sMatches = Split(GetMatchingFiles(sFile), "|")
        For i = 0 To UBound(sMatches)
            FileDeleteIfMD4Match sMatches(i) & "|" & sMD4
        Next i
        Exit Sub
    End If
    If UCase(sMD4) = GetFileMD4(sFile) Then
        'match!
        Logg "Success: FileDeleteIfMD4Match " & sCmd & " (matched " & sFile & ")"
        FileDelete sFile
    End If
End Sub

Public Sub FileDeleteIfCRC32Match(sCmd$)
    'FileDeleteIfCRC32Match <file/filemask>|<crc32>
    Dim sArgs$(), sFile$, sCRC32$
    sArgs = Split(sCmd, "|")
    If UBound(sArgs) <> 1 Then Exit Sub
    sFile = sArgs(0)
    sCRC32 = sArgs(1)
    
    If InStr(sFile, "*") > 0 Or InStr(sFile, "?") > 0 Then
        Dim sMatches$(), i&
        sMatches = Split(GetMatchingFiles(sFile), "|")
        For i = 0 To UBound(sMatches)
            FileDeleteIfCRC32Match sMatches(i) & "|" & sCRC32
        Next i
        Exit Sub
    End If
    If UCase(sCRC32) = modCRC32.GetFileCRC32(sFile) Then
        'match!
        Logg "Success: FileDeleteIfCRC32Match " & sCmd & " (matched " & sFile & ")"
        FileDelete sFile
    End If
End Sub

Public Sub FileDeleteIfContainsText(sCmd$)
    'FileDeleteIfContainsText <file/filemask>|<text>
    Dim sArgs$(), sFile$, sText$, sFileContents$
    sArgs = Split(sCmd, "|")
    If UBound(sArgs) <> 1 Then Exit Sub
    sFile = sArgs(0)
    sText = sArgs(1)
    
    
    If InStr(sFile, "*") > 0 Or InStr(sFile, "?") > 0 Then
        Dim sMatches$(), i&
        sMatches = Split(GetMatchingFiles(sFile), "|")
        For i = 0 To UBound(sMatches)
            If sMatches(i) <> vbNullString Then
                FileDeleteIfContainsText sMatches(i) & "|" & sText
            End If
        Next i
        Exit Sub
    End If
    If Not FileExists(sFile) Then Exit Sub
    sFileContents = InputFile(sFile)
    If InStr(sFileContents, sText) > 0 Then
        'match!
        Logg "Success: FileDeleteIfContainsText " & sCmd & " (matched " & sFile & ")"
        FileDelete sFile
    End If
End Sub

Public Sub FileDeleteIfContainsHex(sCmd$)
    'FileDeleteIfContainsHex <file/filemask>|<hex, comma-seperated>
    Dim sArgs$(), sFile$, sHex$, i&, sHexArray$(), sFileContents$
    sArgs = Split(sCmd, "|")
    If UBound(sArgs) <> 1 Then Exit Sub
    sFile = sArgs(0)
    sHex = sArgs(1)
    
    If InStr(sFile, "*") > 0 Or InStr(sFile, "?") > 0 Then
        Dim sMatches$()
        sMatches = Split(GetMatchingFiles(sFile), "|")
        For i = 0 To UBound(sMatches)
            If sMatches(i) <> vbNullString Then
                FileDeleteIfContainsHex sMatches(i) & "|" & sHex
            End If
        Next i
        Exit Sub
    End If
    If Not FileExists(sFile) Then Exit Sub
    sHexArray = Split(sHex, ",")
    sHex = vbNullString
    For i = 0 To UBound(sHexArray)
        sHex = sHex & Chr(Val("&H" & sHexArray(i)))
    Next i
    sFileContents = InputFile(sFile)
    If InStr(sFileContents, sHex) > 0 Then
        'match!
        Logg "Success: FileDeleteIfContainsHex " & sCmd & " (matched " & sFile & ")"
        FileDelete sFile
    End If
End Sub

Public Sub FileMoveIfMD5Match(sCmd$)
    'FileMoveIfMD5Match <file/filemask>|<folder>|<md5>
    Dim sArgs$(), sFile$, sFolder$, sMD5$
    sArgs = Split(sCmd, "|")
    If UBound(sArgs) <> 2 Then Exit Sub
    sFile = sArgs(0)
    sFolder = sArgs(1)
    sMD5 = sArgs(2)
    
    If InStr(sFile, "?") > 0 Or InStr(sFile, "*") > 0 Then
        Dim sMatches$(), i&
        sMatches = Split(GetMatchingFiles(sFile), "|")
        For i = 0 To UBound(sMatches)
            FileMoveIfMD5Match sMatches(i) & "|" & sFolder & "|" & sMD5
        Next i
        Exit Sub
    End If
    If UCase(sMD5) = GetFileMD5(sFile) Then
        'match!
        Logg "Success: FileMoveIfMD5Match " & sCmd & " (matched " & sFile & ")"
        FileMove sFile & "|" & sFolder
    End If
End Sub

Public Sub FileMoveIfSHA1Match(sCmd$)
    'FileMoveIfSHA1Match <file/filemask>|<folder>|<sha1>
    Dim sFile$, sFolder$, sSHA1$, sArgs$()
    sArgs = Split(sCmd, "|")
    If UBound(sArgs) <> 2 Then Exit Sub
    sFile = sArgs(0)
    sFolder = sArgs(1)
    sSHA1 = sArgs(2)
    
    If InStr(sFile, "?") > 0 Or InStr(sFile, "*") > 0 Then
        Dim sMatches$(), i&
        sMatches = Split(GetMatchingFiles(sFile), "|")
        For i = 0 To UBound(sMatches)
            FileMoveIfSHA1Match sMatches(i) & "|" & sFolder & "|" & sSHA1
        Next i
        Exit Sub
    End If
    If UCase(sSHA1) = GetFileSHA1(sFile) Then
        'match!
        Logg "Success: FileMoveIfSHA1Match " & sCmd & " (matched " & sFile & ")"
        FileMove sFile & "|" & sFolder
    End If
End Sub

Public Sub FileMoveIfMD2Match(sCmd$)
    'FileMoveIfMD2Match <file/filemask>|<folder>|<md2>
    Dim sFile$, sFolder$, sMD2$, sArgs$()
    sArgs = Split(sCmd, "|")
    If UBound(sArgs) <> 2 Then Exit Sub
    sFile = sArgs(0)
    sFolder = sArgs(1)
    sMD2 = sArgs(2)
    
    If InStr(sFile, "?") > 0 Or InStr(sFile, "*") > 0 Then
        Dim sMatches$(), i&
        sMatches = Split(GetMatchingFiles(sFile), "|")
        For i = 0 To UBound(sMatches)
            FileMoveIfMD2Match sMatches(i) & "|" & sFolder & "|" & sMD2
        Next i
        Exit Sub
    End If
    If UCase(sMD2) = GetFileMD2(sFile) Then
        'match!
        Logg "Success: FileMoveIfMD2Match " & sCmd & " (matched " & sFile & ")"
        FileMove sFile & "|" & sFolder
    End If
End Sub

Public Sub FileMoveIfMD4Match(sCmd$)
    'FileMoveIfMD4Match <file/filemask>|<folder>|<md4>
    Dim sFile$, sFolder$, sMD4$, sArgs$()
    sArgs = Split(sCmd, "|")
    If UBound(sArgs) <> 2 Then Exit Sub
    sFile = sArgs(0)
    sFolder = sArgs(1)
    sMD4 = sArgs(2)
    
    If InStr(sFile, "?") > 0 Or InStr(sFile, "*") > 0 Then
        Dim sMatches$(), i&
        sMatches = Split(GetMatchingFiles(sFile), "|")
        For i = 0 To UBound(sMatches)
            FileMoveIfMD4Match sMatches(i) & "|" & sFolder & "|" & sMD4
        Next i
        Exit Sub
    End If
    If UCase(sMD4) = GetFileMD4(sFile) Then
        'match!
        Logg "Success: FileMoveIfMD4Match " & sCmd & " (matched " & sFile & ")"
        FileMove sFile & "|" & sFolder
    End If
End Sub

Public Sub FileMoveIfCRC32Match(sCmd$)
    'FileMoveIfCRC32Match <file/filemask>|<crc32>
    Dim sFile$, sFolder$, sCRC32$, sArgs$()
    sArgs = Split(sCmd, "|")
    If UBound(sArgs) <> 2 Then Exit Sub
    sFile = sArgs(0)
    sFolder = sArgs(1)
    sCRC32 = sArgs(2)
    
    If InStr(sFile, "*") > 0 Or InStr(sFile, "?") > 0 Then
        Dim sMatches$(), i&
        sMatches = Split(GetMatchingFiles(sFile), "|")
        For i = 0 To UBound(sMatches)
            FileMoveIfCRC32Match sMatches(i) & "|" & sFolder & "|" & sCRC32
        Next i
        Exit Sub
    End If
    If UCase(sCRC32) = modCRC32.GetFileCRC32(sFile) Then
        'match!
        Logg "Success: FileMoveIfCRC32Match " & sCmd & " (matched " & sFile & ")"
        FileMove sFile & "|" & sFolder
    End If
End Sub

Public Sub FileMoveIfContainsText(sCmd$)
    'FileMoveIfContainsText <file/filemask>|<folder>|<text>
    Dim sFile$, sFolder$, sText$, sFileContents$, sArgs$()
    sArgs = Split(sCmd, "|")
    If UBound(sArgs) <> 2 Then Exit Sub
    sFile = sArgs(0)
    sFolder = sArgs(1)
    sText = sArgs(2)
    
    If InStr(sFile, "*") > 0 Or InStr(sFile, "?") > 0 Then
        Dim sMatches$(), i&
        sMatches = Split(GetMatchingFiles(sFile), "|")
        For i = 0 To UBound(sMatches)
            If sMatches(i) <> vbNullString Then
                FileMoveIfContainsText sMatches(i) & "|" & sFolder & "|" & sText
            End If
        Next i
        Exit Sub
    End If
    If Not FileExists(sFile) Then Exit Sub
    sFileContents = InputFile(sFile)
    If InStr(sFileContents, sText) > 0 Then
        'match!
        Logg "Success: FileMoveIfContainsText " & sCmd & " (matched " & sFile & ")"
        FileMove sFile & "|" & sFolder
    End If
End Sub

Public Sub FileMoveIfContainsHex(sCmd$)
    'FileMoveIfContainsHex <file/filemask>|<folder>|<hex, comma-seperated>
    Dim sFile$, sFolder$, sHex$, i&, sHexArray$(), sFileContents$, sArgs$()
    sArgs = Split(sCmd, "|")
    If UBound(sArgs) <> 2 Then Exit Sub
    sFile = sArgs(0)
    sFolder = sArgs(1)
    sHex = sArgs(2)
    
    If InStr(sFile, "*") > 0 Or InStr(sFile, "?") > 0 Then
        Dim sMatches$()
        sMatches = Split(GetMatchingFiles(sFile), "|")
        For i = 0 To UBound(sMatches)
            If sMatches(i) <> vbNullString Then
                FileMoveIfContainsHex sMatches(i) & "|" & sFolder & "|" & sHex
            End If
        Next i
        Exit Sub
    End If
    If Not FileExists(sFile) Then Exit Sub
    sHexArray = Split(sHex, ",")
    sHex = vbNullString
    For i = 0 To UBound(sHexArray)
        sHex = sHex & Chr(Val("&H" & sHexArray(i)))
    Next i
    sFileContents = InputFile(sFile)
    If InStr(sFileContents, sHex) > 0 Then
        'match!
        Logg "Success: FileMoveIfContainsHex " & sCmd & " (matched " & sFile & ")"
        FileMove sFile & "|" & sFolder
    End If
End Sub

Public Sub FolderClear(sCmd$)
    'FolderClear <folder>
    Dim sFolder$
    If Not FolderExists(sCmd) Then Exit Sub
    sFolder = sCmd
    
    Dim hFind&, uWFD As WIN32_FIND_DATA, sFile$
    hFind = FindFirstFile(sFolder & "\*.*", uWFD)
    If hFind < 0 Then
        Logg "Failed: FolderClear " & sFolder & " (folder is empty or doesn't exist)"
        Exit Sub
    End If
    
    Do
        sFile = TrimNull(uWFD.cFileName)
        If sFile <> "." And sFile <> ".." Then
            If (uWFD.dwFileAttributes And FILE_ATTRIBUTE_DIRECTORY) Then
                FolderClear sFolder & "\" & sFile
                FolderDelete sFolder & "\" & sFile
            Else
                FileDelete sFolder & "\" & sFile
            End If
        End If
    Loop Until FindNextFile(hFind, uWFD) = 0
    FindClose hFind
End Sub

Public Sub LogIfFileExists(sCmd$)
    'LogIfFileExist <file/filemask>
    Dim sFile$
    sFile = sCmd
    If InStr(sFile, "*") > 0 Or InStr(sFile, "?") > 0 Then
        Dim sMatches$(), i&
        sMatches = Split(GetMatchingFiles(sFile), "|")
        For i = 0 To UBound(sMatches)
            If FileExists(sMatches(i)) Then
                Logg "File exists: " & sMatches(i)
            End If
        Next i
        Exit Sub
    End If
    If FileExists(sFile) Then
        Logg "File exists: " & sFile
    End If
End Sub

Public Sub LogIfFolderExists(sCmd$)
    'LogIfFolderExist <folder>
    Dim sFolder$
    sFolder = sCmd
    If FolderExists(sFolder) Then Logg "Folder exists: " & sFolder
End Sub

Public Sub LogIfFileContainsText(sCmd$)
    'LogIfFileContainsText <file/filemask>|<text>
    Dim sFile$, sText$, sContent$, sArgs$()
    sArgs = Split(sCmd, "|")
    If UBound(sArgs) <> 1 Then Exit Sub
    sFile = sArgs(0)
    sText = sArgs(1)
    
    If InStr(sFile, "*") > 0 Or InStr(sFile, "?") > 0 Then
        Dim sMatches$(), i&
        sMatches = Split(GetMatchingFiles(sFile), "|")
        For i = 0 To UBound(sMatches)
            LogIfFileContainsText sMatches(i) & "|" & sText
        Next i
        Exit Sub
    End If
    
    sContent = InputFile(sFile)
    If InStr(sContent, sText) > 0 Then
        'match!
        Logg "File contains text '" & sText & "': " & sFile
    End If
End Sub

Public Sub LogIfFileContainsHex(sCmd$)
    'LogIfFileContainsHex <file/filemask>|<hex>
    Dim sFile$, sHex$, sHexArray$(), sHex2$, sContent$, i&, sArgs$()
    sArgs = Split(sCmd, "|")
    If UBound(sArgs) <> 1 Then Exit Sub
    sFile = sArgs(0)
    sHex = sArgs(1)
    
    If InStr(sFile, "*") > 0 Or InStr(sFile, "?") > 0 Then
        Dim sMatches$()
        sMatches = Split(GetMatchingFiles(sFile), "|")
        For i = 0 To UBound(sMatches)
            LogIfFileContainsHex sMatches(i) & "|" & sHex
        Next i
        Exit Sub
    End If
    
    sHexArray = Split(sHex, ",")
    For i = 0 To UBound(sHexArray)
        sHex2 = sHex2 & Chr(Val("&H" & sHexArray(i)))
    Next i
    sContent = InputFile(sFile)
    If InStr(sContent, sHex2) > 0 Then
        'match!
        Logg "File contains hexadecimal string '" & sHex & "': " & sFile
    End If
End Sub

Public Sub LogIfFileMD5Match(sCmd$)
    'LogIfFileMD5Match <file/filemask>|<md5>
    Dim sFile$, sMD5$, sArgs$()
    sArgs = Split(sCmd, "|")
    If UBound(sArgs) <> 1 Then Exit Sub
    sFile = sArgs(0)
    sMD5 = sArgs(1)
    
    If InStr(sFile, "?") > 0 Or InStr(sFile, "*") > 0 Then
        Dim sMatches$(), i&
        sMatches = Split(GetMatchingFiles(sFile), "|")
        For i = 0 To UBound(sMatches)
            LogIfFileMD5Match sMatches(i) & "|" & sMD5
        Next i
        Exit Sub
    End If
    If UCase(sMD5) = GetFileMD5(sFile) Then
        'match!
        Logg "File matches MD5 " & sMD5 & ": " & sFile
    End If
End Sub

Public Sub LogIfFileSHA1Match(sCmd$)
    'LogIfFileSHA1Match <file/filemask>|<sha1>
    Dim sFile$, sSHA1$, sArgs$()
    sArgs = Split(sCmd, "|")
    If UBound(sArgs) <> 1 Then Exit Sub
    sFile = sArgs(0)
    sSHA1 = sArgs(1)
    
    If InStr(sFile, "?") > 0 Or InStr(sFile, "*") > 0 Then
        Dim sMatches$(), i&
        sMatches = Split(GetMatchingFiles(sFile), "|")
        For i = 0 To UBound(sMatches)
            LogIfFileSHA1Match sMatches(i) & "|" & sSHA1
        Next i
        Exit Sub
    End If
    If UCase(sSHA1) = GetFileSHA1(sFile) Then
        'match!
        Logg "File matches SHA1 " & sSHA1 & ": " & sFile
    End If
End Sub

Public Sub LogIfFileMD2Match(sCmd$)
    'LogIfFileMD2Match <file/filemask>|<md2>
    Dim sFile$, sMD2$, sArgs$()
    sArgs = Split(sCmd, "|")
    If UBound(sArgs) <> 1 Then Exit Sub
    sFile = sArgs(0)
    sMD2 = sArgs(1)
    
    If InStr(sFile, "?") > 0 Or InStr(sFile, "*") > 0 Then
        Dim sMatches$(), i&
        sMatches = Split(GetMatchingFiles(sFile), "|")
        For i = 0 To UBound(sMatches)
            LogIfFileMD2Match sMatches(i) & "|" & sMD2
        Next i
        Exit Sub
    End If
    If UCase(sMD2) = GetFileMD2(sFile) Then
        'match!
        Logg "File matches MD2 " & sMD2 & ": " & sFile
    End If
End Sub

Public Sub LogIfFileMD4Match(sCmd$)
    'LogIfFileMD4Match <file/filemask>|<md4>
    Dim sFile$, sMD4$, sArgs$()
    sArgs = Split(sCmd, "|")
    If UBound(sArgs) <> 1 Then Exit Sub
    sFile = sArgs(0)
    sMD4 = sArgs(1)
    
    If InStr(sFile, "?") > 0 Or InStr(sFile, "*") > 0 Then
        Dim sMatches$(), i&
        sMatches = Split(GetMatchingFiles(sFile), "|")
        For i = 0 To UBound(sMatches)
            LogIfFileMD4Match sMatches(i) & "|" & sMD4
        Next i
        Exit Sub
    End If
    If UCase(sMD4) = GetFileMD4(sFile) Then
        'match!
        Logg "File matches MD4 " & sMD4 & ": " & sFile
    End If
End Sub

Public Sub LogIfFileCRC32Match(sCmd$)
    'LogIfFileCRC32Match <file/filemask>|<crc32>
    Dim sFile$, sCRC32$, sArgs$()
    sArgs = Split(sCmd, "|")
    If UBound(sArgs) <> 1 Then Exit Sub
    sFile = sArgs(0)
    sCRC32 = sArgs(1)
    
    If InStr(sFile, "?") > 0 Or InStr(sFile, "*") > 0 Then
        Dim sMatches$(), i&
        sMatches = Split(GetMatchingFiles(sFile), "|")
        For i = 0 To UBound(sMatches)
            LogIfFileCRC32Match sMatches(i) & "|" & sCRC32
        Next i
        Exit Sub
    End If
    If UCase(sCRC32) = GetFileCRC32(sFile) Then
        'match!
        Logg "File matches CRC32 " & sCRC32 & ": " & sFile
    End If
End Sub

Public Function GetMatchingFiles$(sFileMask$, Optional bSubfolders As Boolean = True)
    'internal function - used by FileDeleteOnReboot and
    'anything with files that uses wildcards
    '= this searches subfolders as well since v1.10 =
    
    Dim hFind&, uWFD As WIN32_FIND_DATA, sPath$, sMatch$, sFile$, sMatches$
    If InStr(sFileMask, "\") = 0 Then Exit Function
    sPath = Left(sFileMask, InStrRev(sFileMask, "\"))
    sMatch = Mid(sFileMask, InStrRev(sFileMask, "\") + 1)
    
    hFind = FindFirstFile(sPath & "*", uWFD)
    If hFind <= 0 Then Exit Function
    Do
        sFile = TrimNull(uWFD.cFileName)
        If sFile <> "." And sFile <> ".." And sFile <> vbNullString Then
            If (uWFD.dwFileAttributes And FILE_ATTRIBUTE_DIRECTORY) And bSubfolders Then
                sMatches = sMatches & "|" & GetMatchingFiles(sPath & sFile & "\" & sMatch)
                sMatches = Replace(sMatches, "||", "|")
            Else
                If LCase(sFile) Like LCase(sMatch) Then
                    sMatches = sMatches & "|" & sPath & sFile
                End If
            End If
        End If
    Loop Until FindNextFile(hFind, uWFD) = 0
    FindClose hFind
    If sMatches <> vbNullString Then
        GetMatchingFiles = Mid(sMatches, 2)
    End If
End Function

