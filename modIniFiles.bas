Attribute VB_Name = "modIniFiles"
Option Explicit
'ini file (not inimapping) settings handling

Public Sub IniSetValue(sCmd$)
    'IniSetValue <file>|<section>|<value>|<data>
    On Error Resume Next
    Dim sFile$, sSection$, sValue$, sData$, sArgs$()
    Dim sLine$, bSectionFound As Boolean
    
    sArgs = Split(sCmd, "|")
    If UBound(sArgs) <> 3 Then Exit Sub
    sFile = sArgs(0)
    sSection = sArgs(1)
    sValue = sArgs(2)
    sData = sArgs(3)
    
    If Not FileExists(sFile) Then
        Open sFile For Output As #1
            Print #1, "[" & sSection & "]"
            Print #1, sValue & "=" & sData
        Close #1
    Else
        Kill sFile & ".new"
        Open sFile For Input As #1
        Open sFile & ".new" For Output As #2
            Do
                Line Input #1, sLine
                If bSectionFound And InStr(sLine, "[") = 1 Then bSectionFound = False
                If InStr(1, sLine, "[" & sSection & "]", vbTextCompare) = 1 Then bSectionFound = True
                
                If InStr(1, sLine, sValue, vbTextCompare) = 1 And bSectionFound Then
                    Print #2, sValue & "=" & sData
                    bSectionFound = False
                Else
                    Print #2, sLine
                End If
            Loop Until EOF(1)
        Close #2
        Close #1
        
        Kill sFile
        Name sFile & ".new" As sFile
    End If
    If Err Then
        Logg "Failed: IniSetValue " & sCmd & " (write error)"
    Else
        Logg "Success: IniSetValue " & sCmd
    End If
End Sub

Public Sub IniDeleteValue(sCmd$)
    'IniDeleteValue <file>|<section>|<value>
    On Error Resume Next
    Dim sFile$, sSection$, sValue$, sData$, sArgs$()
    Dim sLine$, bSectionFound As Boolean
    
    sArgs = Split(sCmd, "|")
    If UBound(sArgs) <> 2 Then Exit Sub
    sFile = sArgs(0)
    sSection = sArgs(1)
    sValue = sArgs(2)
    
    If Not FileExists(sFile) Then
        Logg "Failed: IniDeleteValue " & sCmd & " (ini file not found)"
        Exit Sub
    Else
        Kill sFile & ".new"
        Open sFile For Input As #1
        Open sFile & ".new" For Output As #2
            Do
                Line Input #1, sLine
                If bSectionFound And InStr(sLine, "[") = 1 Then bSectionFound = False
                If InStr(1, sLine, "[" & sSection & "]", vbTextCompare) = 1 Then bSectionFound = True
                
                If InStr(1, sLine, sValue, vbTextCompare) = 1 And bSectionFound Then
                    Print #2, sValue & "=" & sData
                    bSectionFound = False
                Else
                    Print #2, sLine
                End If
            Loop Until EOF(1)
        Close #2
        Close #1
        
        Kill sFile
        Name sFile & ".new" As sFile
    End If
    If Err Then
        Logg "Failed: IniDeleteValue " & sCmd & " (write error)"
    Else
        Logg "Success: IniDeleteValue " & sCmd
    End If
End Sub

Public Sub IniDeleteFromValue(sCmd$)
    'IniDeleteFromValue <file>|<section>|<value>|<data>
    On Error Resume Next
    Dim sFile$, sSection$, sValue$, sData$, sData2$, sArgs$()
    Dim sLine$, bSectionFound As Boolean
    
    sArgs = Split(sCmd, "|")
    If UBound(sArgs) <> 3 Then Exit Sub
    sFile = sArgs(0)
    sSection = sArgs(1)
    sValue = sArgs(2)
    sData2 = sArgs(3)
    
    If Not FileExists(sFile) Then
        Logg "Failed: IniDeleteFromValue " & sCmd & " (ini file not found)"
        Exit Sub
    Else
        Kill sFile & ".new"
        Open sFile For Input As #1
        Open sFile & ".new" For Output As #2
            Do
                Line Input #1, sLine
                If bSectionFound And InStr(sLine, "[") = 1 Then bSectionFound = False
                If InStr(1, sLine, "[" & sSection & "]", vbTextCompare) = 1 Then bSectionFound = True
                
                If InStr(1, sLine, sValue, vbTextCompare) = 1 And bSectionFound Then
                    sData = Mid(sLine, InStr(sLine, "=") + 1)
                    sData = Replace(sData, sData2, vbNullString, , , vbTextCompare)
                    Print #2, sValue & "=" & sData
                    bSectionFound = False
                Else
                    Print #2, sLine
                End If
            Loop Until EOF(1)
        Close #2
        Close #1
        
        Kill sFile
        Name sFile & ".new" As sFile
    End If
    If Err Then
        Logg "Failed: IniDeleteFromValue " & sCmd & " (write error)"
    Else
        Logg "Success: IniDeleteFromValue " & sCmd
    End If
End Sub

Public Sub IniClearValue(sCmd$)
    'IniClearValue <file>|<section>|<value>
    On Error Resume Next
    Dim sFile$, sSection$, sValue$, sArgs$()
    Dim sLine$, bSectionFound As Boolean
    
    sArgs = Split(sCmd, "|")
    If UBound(sArgs) <> 2 Then Exit Sub
    sFile = sArgs(0)
    sSection = sArgs(1)
    sValue = sArgs(2)
    
    If Not FileExists(sFile) Then
        Logg "Failed: IniClearValue " & sCmd & " (ini file not found)"
        Exit Sub
    Else
        Kill sFile & ".new"
        Open sFile For Input As #1
        Open sFile & ".new" For Output As #2
            Do
                Line Input #1, sLine
                If bSectionFound And InStr(sLine, "[") = 1 Then bSectionFound = False
                If InStr(1, sLine, "[" & sSection & "]", vbTextCompare) = 1 Then bSectionFound = True
                
                If InStr(1, sLine, sValue, vbTextCompare) = 1 And bSectionFound Then
                    Print #2, sValue & "="
                    bSectionFound = False
                Else
                    Print #2, sLine
                End If
            Loop Until EOF(1)
        Close #2
        Close #1
        
        Kill sFile
        Name sFile & ".new" As sFile
    End If
    If Err Then
        Logg "Failed: IniClearValue " & sCmd & " (write error)"
    Else
        Logg "Success: IniClearValue " & sCmd
    End If
End Sub

Public Sub IniCreateSection(sCmd$)
    'IniCreateSection <file>|<section>
    On Error Resume Next
    Dim sFile$, sSection$, sArgs$()
    Dim sLines$

    sArgs = Split(sCmd, "|")
    If UBound(sArgs) <> 1 Then Exit Sub
    sFile = sArgs(0)
    sSection = sArgs(1)
    
    If Not FileExists(sFile) Then
        Open sFile For Output As #1
            Print #1, "[" & sSection & "]"
        Close #1
        Exit Sub
    Else
        sLines = InputFile(sFile)
        If InStr(sLines, vbCrLf & "[" & sSection & "]" & vbCrLf) = 0 Then
            sLines = sLines & vbCrLf & "[" & sSection & "]" & vbCrLf
            Open sFile & ".new" For Output As #1
                Print #1, sLines
            Close #1
            Kill sFile
            Name sFile & ".new" As sFile
        End If
    End If
    If Err Then
        Logg "Failed: IniCreateSection " & sCmd & " (write error)"
    Else
        Logg "Success: IniCreateSection " & sCmd
    End If
End Sub
