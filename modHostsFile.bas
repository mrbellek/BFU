Attribute VB_Name = "modHostsFile"
Option Explicit
'hosts file: delete/add/disable/enable lines, or reset hosts file

Private Const FILE_ATTRIBUTE_COMPRESSED = &H800

Public Sub HostsFileReset()
    On Error Resume Next
    SetAttr sHostsFile, vbArchive
    Kill sHostsFile
    Open sHostsFile For Output As #1
        Print #1, "# Copyright (c) 1993-1999 Microsoft Corp."
        Print #1, "#"
        Print #1, "# This is a sample HOSTS file used by Microsoft TCP/IP for Windows."
        Print #1, "#"
        Print #1, "# This file contains the mappings of IP addresses to host names. Each"
        Print #1, "# entry should be kept on an individual line. The IP address should"
        Print #1, "# be placed in the first column followed by the corresponding host name."
        Print #1, "# The IP address and the host name should be separated by at least one"
        Print #1, "# space."
        Print #1, "#"
        Print #1, "# Additionally, comments (such as these) may be inserted on individual"
        Print #1, "# lines or following the machine name denoted by a '#' symbol."
        Print #1, "#"
        Print #1, "# For example:"
        Print #1, "#"
        Print #1, "#      102.54.94.97     rhino.acme.com          # source server"
        Print #1, "#       38.25.63.10     x.acme.com              # x client host"
        Print #1,
        Print #1, "127.0.0.1       localhost"
    Close #1
    SetAttr sHostsFile, vbArchive
    If bIsWinNT Then HostsFileResetDatabasePath
    If Err Then
        Logg "Failed: HostsFileReset (write error)"
    Else
        Logg "Success: HostsFileReset"
    End If
End Sub

Public Sub HostsFileAddLine(sCmd$)
    'HostsFileAddLine <line>
    On Error Resume Next
    Dim sLine$, iAttr%
    sLine = sCmd
    If sLine = vbNullString Then Exit Sub
    iAttr = GetAttr(sHostsFile)
    If (iAttr And FILE_ATTRIBUTE_COMPRESSED) Then iAttr = iAttr - FILE_ATTRIBUTE_COMPRESSED
    SetAttr sHostsFile, vbArchive
    Open sHostsFile For Append As #1
        Print #1, sLine
    Close #1
    SetAttr sHostsFile, iAttr
    If Err Then
        Logg "Failed: HostsFileAddLine " & sCmd & " (write error)"
    Else
        Logg "Success: HostsFileAddLine " & sCmd
    End If
End Sub

Public Sub HostsFileDelLine(sCmd$)
    'HostsFileDelLine <line>
    On Error Resume Next
    Dim sLine$, sContent$(), i%, iAttr%
    sLine = sCmd
    If sLine = vbNullString Then Exit Sub
    iAttr = GetAttr(sHostsFile)
    If (iAttr And FILE_ATTRIBUTE_COMPRESSED) Then iAttr = iAttr - FILE_ATTRIBUTE_COMPRESSED
    SetAttr sHostsFile, vbArchive
    sContent = Split(InputFile(sHostsFile), vbCrLf)
    'Open sHostsFile For Binary As #1
    '    sContent = Split(Input(FileLen(sHostsFile), #1), vbCrLf)
    'Close #1
    If UBound(sContent) = -1 Then Exit Sub
    If InStr(sContent(0), Chr(10)) > 0 Then
        sContent = Split(Join(sContent, vbCrLf), Chr(10))
    End If
    For i = 0 To UBound(sContent)
        sContent(i) = Replace(sContent(i), vbTab, " ")
        Do
            sContent(i) = Replace(sContent(i), "  ", " ")
        Loop Until InStr(sContent(i), "  ") = 0
        
        If InStr(1, sContent(i), sLine, vbTextCompare) = 1 Then
            sContent(i) = "<line deleted>"
        End If
    Next i
    Open sHostsFile For Output As #1
        For i = 0 To UBound(sContent)
            If sContent(i) <> "<line deleted>" Then Print #1, sContent(i)
        Next i
    Close #1
    SetAttr sHostsFile, iAttr
    If Err Then
        Logg "Failed: HostsFileDelLine " & sCmd & " (write error)"
    Else
        Logg "Success: HostsFileDelLine " & sCmd
    End If
End Sub

Public Sub HostsFileDisableLine(sCmd$)
    'HostsFileDisableLine <line>
    On Error Resume Next
    Dim sLine$, sContent$(), i%, iAttr%
    sLine = sCmd
    If sLine = vbNullString Then Exit Sub
    iAttr = GetAttr(sHostsFile)
    If (iAttr And FILE_ATTRIBUTE_COMPRESSED) Then iAttr = iAttr - FILE_ATTRIBUTE_COMPRESSED
    SetAttr sHostsFile, vbArchive
    
    sContent = Split(InputFile(sHostsFile), vbCrLf)
    If UBound(sContent) = -1 Then Exit Sub
    If InStr(sContent(0), Chr(10)) > 0 Then
        sContent = Split(Join(sContent, vbCrLf), Chr(10))
    End If
    For i = 0 To UBound(sContent)
        sContent(i) = Replace(sContent(i), vbTab, " ")
        Do
            sContent(i) = Replace(sContent(i), "  ", " ")
        Loop Until InStr(sContent(i), "  ") = 0
        
        If InStr(1, sContent(i), sLine, vbTextCompare) > 0 Then
            sContent(i) = "#" & sContent(i)
        End If
    Next i
    Open sHostsFile For Output As #1
        For i = 0 To UBound(sContent)
            Print #1, sContent(i)
        Next i
    Close #1
    SetAttr sHostsFile, iAttr
    If Err Then
        Logg "Failed: HostsFileDisableLine " & sCmd & " (write error)"
    Else
        Logg "Success: HostsFileDisableLine " & sCmd
    End If
End Sub

Public Sub HostsFileEnableLine(sCmd$)
    'HostsFileEnableLine <line>
    On Error Resume Next
    Dim sLine$, sContent$(), i%, iAttr%
    sLine = sCmd
    If sLine = vbNullString Then Exit Sub
    If Left(sLine, 1) <> "#" Then Exit Sub
    iAttr = GetAttr(sHostsFile)
    If (iAttr And FILE_ATTRIBUTE_COMPRESSED) Then iAttr = iAttr - FILE_ATTRIBUTE_COMPRESSED
    SetAttr sHostsFile, vbArchive
    sContent = Split(InputFile(sHostsFile), vbCrLf)
    'Open sHostsFile For Binary As #1
    '    sContent = Split(Input(FileLen(sHostsFile), #1), vbCrLf)
    'Close #1
    If UBound(sContent) = -1 Then Exit Sub
    If InStr(sContent(0), Chr(10)) > 0 Then
        sContent = Split(Join(sContent, vbCrLf), Chr(10))
    End If
    For i = 0 To UBound(sContent)
        sContent(i) = Replace(sContent(i), vbTab, " ")
        Do
            sContent(i) = Replace(sContent(i), "  ", " ")
        Loop Until InStr(sContent(i), "  ") = 0
        
        If InStr(1, sContent(i), sLine, vbTextCompare) > 0 Then
            sContent(i) = Mid(sContent(i), 2)
        End If
    Next i
    Open sHostsFile For Output As #1
        For i = 0 To UBound(sContent)
            Print #1, sContent(i)
        Next i
    Close #1
    SetAttr sHostsFile, iAttr
    If Err Then
        Logg "Failed: HostsFileEnableLine " & sCmd & " (write error)"
    Else
        Logg "Success: HostsFileEnableLine " & sCmd
    End If
End Sub
