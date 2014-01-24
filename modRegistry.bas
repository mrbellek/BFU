Attribute VB_Name = "modRegistry"
Option Explicit
'setting, deleting, creating registry values/keys

Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Function RegCreateKeyEx Lib "advapi32.dll" Alias "RegCreateKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal Reserved As Long, ByVal lpClass As String, ByVal dwOptions As Long, ByVal samDesired As Long, lpSecurityAttributes As Any, phkResult As Long, lpdwDisposition As Long) As Long
Private Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal hKey As Long, ByVal lpValueName As String) As Long
Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
Private Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long
Private Declare Function RegEnumKeyEx Lib "advapi32.dll" Alias "RegEnumKeyExA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpName As String, lpcbName As Long, ByVal lpReserved As Long, ByVal lpClass As String, lpcbClass As Long, lpftLastWriteTime As Any) As Long
Private Declare Function RegEnumValue Lib "advapi32.dll" Alias "RegEnumValueA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpValueName As String, lpcbValueName As Long, ByVal lpReserved As Long, lpType As Long, lpData As Byte, lpcbData As Long) As Long
Private Declare Function RegSaveKey Lib "advapi32.dll" Alias "RegSaveKeyA" (ByVal hKey As Long, ByVal lpFile As String, lpSecurityAttributes As Any) As Long
Private Declare Function RegRestoreKey Lib "advapi32.dll" Alias "RegRestoreKeyA" (ByVal hKey As Long, ByVal lpFile As String, ByVal dwFlags As Long) As Long
Private Declare Function SHDeleteKey Lib "shlwapi.dll" Alias "SHDeleteKeyA" (ByVal lRootKey As Long, ByVal szKeyToDelete As String) As Long
'note: MSDN says SHDeleteKey will only delete regkey on WinNT if it
'has no subkeys - this is not true, it will always work.

Private Declare Function GetCurrentProcess Lib "kernel32" () As Long
Private Declare Function OpenProcessToken Lib "advapi32.dll" (ByVal ProcessHandle As Long, ByVal DesiredAccess As Long, TokenHandle As Long) As Long
Private Declare Function LookupPrivilegeValue Lib "advapi32.dll" Alias "LookupPrivilegeValueA" (ByVal lpSystemName As String, ByVal lpName As String, lpLuid As LUID) As Long
Private Declare Function AdjustTokenPrivileges Lib "advapi32.dll" (ByVal TokenHandle As Long, ByVal DisableAllPrivileges As Long, NewState As TOKEN_PRIVILEGES, ByVal BufferLength As Long, PreviousState As TOKEN_PRIVILEGES, ReturnLength As Long) As Long

Private Type LUID
    LowPart As Long
    HighPart As Long
End Type

Private Type LUID_AND_ATTRIBUTES
    pLuid As LUID
    Attributes As Long
End Type

Private Type TOKEN_PRIVILEGES
    PrivilegeCount As Long
    Privileges As LUID_AND_ATTRIBUTES
End Type

Private Const SE_PRIVILEGE_ENABLED = &H2
Private Const TOKEN_QUERY As Long = &H8&
Private Const TOKEN_ADJUST_PRIVILEGE As Long = &H20&
Private Const REG_FORCE_RESTORE As Long = 8&

Public Const HKEY_CLASSES_ROOT = &H80000000
Public Const HKEY_CURRENT_USER = &H80000001
Public Const HKEY_LOCAL_MACHINE = &H80000002
Public Const HKEY_USERS = &H80000003
Private Const HKEY_PERFORMANCE_DATA = &H80000004
Private Const HKEY_CURRENT_CONFIG = &H80000005
Private Const HKEY_DYN_DATA = &H80000006

Private Const REG_SZ = 1
Private Const REG_EXPAND_SZ = 2
Private Const REG_BINARY = 3
Private Const REG_DWORD = 4
Private Const REG_MULTI_SZ = 7

Private Const KEY_CREATE_SUB_KEY = &H4
Private Const KEY_QUERY_VALUE = &H1
Private Const KEY_SET_VALUE = &H2
Private Const KEY_ENUMERATE_SUB_KEYS = &H8
Private Const READ_CONTROL = &H20000
Private Const SYNCHRONIZE = &H100000
Private Const STANDARD_RIGHTS_WRITE = (READ_CONTROL)
Private Const KEY_WRITE = ((STANDARD_RIGHTS_WRITE Or KEY_SET_VALUE Or KEY_CREATE_SUB_KEY) And (Not SYNCHRONIZE))

Public MAX_REG_VALUE_NAME As Long

Public Sub RegCreateKey(sCmd$)
    'RegCreateKey <hive\key>
    Dim hKey&, lHive&, sKey$
    If Len(sCmd) < 7 Then Exit Sub
    Select Case Left(sCmd, 4)
        Case "HKCR": lHive = HKEY_CLASSES_ROOT
        Case "HKCU": lHive = HKEY_CURRENT_USER
        Case "HKLM": lHive = HKEY_LOCAL_MACHINE
        Case "HKUS": lHive = HKEY_USERS
        Case "HKPD": lHive = HKEY_PERFORMANCE_DATA
        Case "HKCC": lHive = HKEY_CURRENT_CONFIG
        Case "HKDD": lHive = HKEY_DYN_DATA
        Case Else: Exit Sub
    End Select
    sKey = Mid(sCmd, 6)
    If RegKeyExists(lHive, sKey) Then
        Logg "Failed RegCreateKey " & sCmd & " (key alredy exists)"
        Exit Sub
    End If
    'this won't error out if key exists - it will just open the key
    If RegCreateKeyEx(lHive, sKey, 0, vbNullString, 0, KEY_CREATE_SUB_KEY, ByVal 0, hKey, 0) <> 0 Then
        Logg "Failed: RegCreateKey " & sCmd & " (operation failed)"
    Else
        Logg "Success: RegCreateKey " & sCmd
    End If
    RegCloseKey hKey
End Sub

Public Sub RegDeleteKey(sCmd$)
    'RegDeleteKey <hive\key>
    Dim lHive&, sKey$
    If Len(sCmd) < 7 Then Exit Sub
    Select Case Left(sCmd, 4)
        Case "HKCR": lHive = HKEY_CLASSES_ROOT
        Case "HKCU": lHive = HKEY_CURRENT_USER
        Case "HKLM": lHive = HKEY_LOCAL_MACHINE
        Case "HKUS": lHive = HKEY_USERS
        Case "HKPD": lHive = HKEY_PERFORMANCE_DATA
        Case "HKCC": lHive = HKEY_CURRENT_CONFIG
        Case "HKDD": lHive = HKEY_DYN_DATA
        Case Else: Exit Sub
    End Select
    sKey = Mid(sCmd, 6)
    If InStr(sKey, "*") = 0 And InStr(sKey, "?") = 0 Then
        If Not RegKeyExists(lHive, sKey) Then
            Logg "Failed: RegDeleteKey " & sCmd & " (key does not exist)"
            Exit Sub
        End If
        If SHDeleteKey(lHive, sKey) <> 0 Then
            Logg "Failed: RegDeleteKey " & sCmd & " (operation failed)"
        Else
            Logg "Success: RegDeleteKey " & sCmd
        End If
    Else
        Dim sKeys$(), i&, sParent$
        sParent = Left(sKey, InStrRev(sKey, "\") - 1)
        sKeys = Split(RegEnumSubKeys(lHive, sParent), "|")
        For i = 0 To UBound(sKeys)
            If LCase(sParent & "\" & sKeys(i)) Like LCase(sKey) Then
                If SHDeleteKey(lHive, sParent & "\" & sKeys(i)) <> 0 Then
                    Logg "Failed: RegDeleteKey " & sCmd & " (operation failed)"
                Else
                    Logg "Success: RegDeleteKey " & sCmd
                End If
            End If
        Next i
    End If
End Sub

Public Sub RegDeleteKeyIfNameContainsText(sCmd$)
    'RegDeleteKeyIfNameContainsText <hive\key>|<mask>|<text>
    Dim lHive&, sKey$, sMask$, sText$, sSubKeys$(), i&, sArgs$()
    sArgs = Split(sCmd, "|")
    If UBound(sArgs) <> 2 Then Exit Sub
    Select Case Left(sArgs(0), 4)
        Case "HKCR": lHive = HKEY_CLASSES_ROOT
        Case "HKCU": lHive = HKEY_CURRENT_USER
        Case "HKLM": lHive = HKEY_LOCAL_MACHINE
        Case "HKUS": lHive = HKEY_USERS
        Case "HKPD": lHive = HKEY_PERFORMANCE_DATA
        Case "HKCC": lHive = HKEY_CURRENT_CONFIG
        Case "HKDD": lHive = HKEY_DYN_DATA
        Case Else: Exit Sub
    End Select
    sKey = Mid(sArgs(0), 6)
    sMask = sArgs(1)
    sText = sArgs(2)
    
    If Not RegKeyExists(lHive, sKey) Then
        Logg "Failed: RegDeleteKeyIfNameContainsText " & sCmd & " (key not found)"
        Exit Sub
    End If
    
    sSubKeys = Split(RegEnumSubKeys(lHive, sKey), "|")
    For i = 0 To UBound(sSubKeys)
        If LCase(sSubKeys(i)) Like "*" & LCase(sMask) & "*" And _
           InStr(1, sSubKeys(i), sText, vbTextCompare) > 0 Then
            RegDeleteKey sArgs(0) & "\" & sSubKeys(i)
        End If
    Next i
End Sub

Public Sub RegDeleteKeyIfNameContainsHex(sCmd$)
    'RegDeleteKeyIfNameContainsHex <hive\key>|<mask>|<hex>
    Dim lHive&, sKey$, sMask$, sSubKeys$(), i&, sArgs$()
    Dim sHex$, sHex2$, sHexArray$()
    
    sArgs = Split(sCmd, "|")
    If UBound(sArgs) <> 2 Then Exit Sub
    Select Case Left(sArgs(0), 4)
        Case "HKCR": lHive = HKEY_CLASSES_ROOT
        Case "HKCU": lHive = HKEY_CURRENT_USER
        Case "HKLM": lHive = HKEY_LOCAL_MACHINE
        Case "HKUS": lHive = HKEY_USERS
        Case "HKPD": lHive = HKEY_PERFORMANCE_DATA
        Case "HKCC": lHive = HKEY_CURRENT_CONFIG
        Case "HKDD": lHive = HKEY_DYN_DATA
        Case Else: Exit Sub
    End Select
    sKey = Mid(sArgs(0), 6)
    sMask = sArgs(1)
    sHex = sArgs(2)
    
    sHexArray = Split(sHex, ",")
    For i = 0 To UBound(sHexArray)
        sHex2 = sHex2 & Chr(Val("&H" & sHexArray(i)))
    Next i
    
    If Not RegKeyExists(lHive, sKey) Then
        Logg "Failed: RegDeleteKeyIfNameContainsHex " & sCmd & " (key not found)"
        Exit Sub
    End If
    
    sSubKeys = Split(RegEnumSubKeys(lHive, sKey), "|")
    For i = 0 To UBound(sSubKeys)
        If LCase(sSubKeys(i)) Like "*" & LCase(sMask) & "*" And _
           InStr(sSubKeys(i), sHex2) > 0 Then
            RegDeleteKey sArgs(0) & "\" & sSubKeys(i)
        End If
    Next i
End Sub

Public Sub RegSetStringValue(sCmd$)
    'RegSetStringValue <hive\key>|<value>|<data>
    Dim hKey&, lHive&, sKey$, sVal$, sData$, sArgs$()
    sArgs = Split(sCmd, "|")
    If UBound(sArgs) <> 2 Then Exit Sub
    Select Case Left(sArgs(0), 4)
        Case "HKCR": lHive = HKEY_CLASSES_ROOT
        Case "HKCU": lHive = HKEY_CURRENT_USER
        Case "HKLM": lHive = HKEY_LOCAL_MACHINE
        Case "HKUS": lHive = HKEY_USERS
        Case "HKPD": lHive = HKEY_PERFORMANCE_DATA
        Case "HKCC": lHive = HKEY_CURRENT_CONFIG
        Case "HKDD": lHive = HKEY_DYN_DATA
        Case Else: Exit Sub
    End Select
    sKey = Mid(sArgs(0), 6)
    sVal = sArgs(1)
    sData = sArgs(2)
    
    If RegCreateKeyEx(lHive, sKey, 0, vbNullString, 0, KEY_CREATE_SUB_KEY Or KEY_SET_VALUE, ByVal 0, hKey, 0) = 0 Then
        If RegSetValueEx(hKey, sVal, 0, REG_SZ, ByVal sData, Len(sData)) <> 0 Then
            Logg "Failed: RegSetStringValue " & sCmd & " (unable to write to Registry)"
        Else
            Logg "Success: RegSetStringValue " & sCmd
        End If
        RegCloseKey hKey
    Else
        Logg "Failed: RegSetStringValue " & sCmd & " (unable to open/create Registry key)"
    End If
End Sub

Public Sub RegSetDwordValue(sCmd$)
    'RegSetDwordValue <hive\key>|<value>|<data>
    Dim hKey&, lHive&, sKey$, sVal$, lData&, sArgs$()
    sArgs = Split(sCmd, "|")
    If UBound(sArgs) <> 2 Then Exit Sub
    Select Case Left(sArgs(0), 4)
        Case "HKCR": lHive = HKEY_CLASSES_ROOT
        Case "HKCU": lHive = HKEY_CURRENT_USER
        Case "HKLM": lHive = HKEY_LOCAL_MACHINE
        Case "HKUS": lHive = HKEY_USERS
        Case "HKPD": lHive = HKEY_PERFORMANCE_DATA
        Case "HKCC": lHive = HKEY_CURRENT_CONFIG
        Case "HKDD": lHive = HKEY_DYN_DATA
        Case Else: Exit Sub
    End Select
    sKey = Mid(sArgs(0), 6)
    sVal = sArgs(1)
    lData = CLng(Val(sArgs(2)))
    
    If RegCreateKeyEx(lHive, sKey, 0, vbNullString, 0, KEY_CREATE_SUB_KEY Or KEY_SET_VALUE, ByVal 0, hKey, 0) = 0 Then
        If RegSetValueEx(hKey, sVal, 0, REG_DWORD, lData, 4) <> 0 Then
            Logg "Failed: RegSetDwordValue " & sCmd & " (unable to write to Registry)"
        Else
            Logg "Success: RegSetDwordValue " & sCmd
        End If
        RegCloseKey hKey
    Else
        Logg "Failed: RegSetDwordValue " & sCmd & " (unable to open/create Registry key)"
    End If
End Sub

Public Sub RegSetBinaryValue(sCmd$)
    'RegSetDwordValue <hive\key>|<value>|<data>
    'data should be in hex, comma delimited
    Dim hKey&, lHive&, sKey$, sVal$, sData$, sTemp$(), i&, uData() As Byte, sArgs$()
    sArgs = Split(sCmd, "|")
    If UBound(sArgs) <> 2 Then Exit Sub
    Select Case Left(sArgs(0), 4)
        Case "HKCR": lHive = HKEY_CLASSES_ROOT
        Case "HKCU": lHive = HKEY_CURRENT_USER
        Case "HKLM": lHive = HKEY_LOCAL_MACHINE
        Case "HKUS": lHive = HKEY_USERS
        Case "HKPD": lHive = HKEY_PERFORMANCE_DATA
        Case "HKCC": lHive = HKEY_CURRENT_CONFIG
        Case "HKDD": lHive = HKEY_DYN_DATA
        Case Else: Exit Sub
    End Select
    sKey = Mid(sArgs(0), 6)
    sVal = sArgs(1)
    sData = sArgs(2)
    
    sTemp = Split(sData, ",")
    ReDim uData(UBound(sTemp))
    For i = 0 To UBound(sTemp)
        uData(i) = Val("&H" & sTemp(i))
    Next i
    
    If RegCreateKeyEx(lHive, sKey, 0, vbNullString, 0, KEY_CREATE_SUB_KEY Or KEY_SET_VALUE, ByVal 0, hKey, 0) = 0 Then
        If RegSetValueEx(hKey, sVal, 0, REG_BINARY, uData(0), UBound(uData) + 1) <> 0 Then
            Logg "Failed: RegSetBinaryValue " & sCmd & " (unable to write to Registry)"
        Else
            Logg "Success: RegSetBinaryValue " & sCmd
        End If
        RegCloseKey hKey
    Else
        Logg "Failed: RegSetBinaryValue " & sCmd & " (unable to open/create Registry key)"
    End If
End Sub

Public Sub RegSetMultiValue(sCmd$)
    'RegSetMultiValue <hive\key>|<value>|<data>
    Dim hKey&, lHive&, sKey$, sVal$, sData$, sArgs$()
    sArgs = Split(sCmd, "|")
    If UBound(sArgs) <> 2 Then Exit Sub
    Select Case Left(sArgs(0), 4)
        Case "HKCR": lHive = HKEY_CLASSES_ROOT
        Case "HKCU": lHive = HKEY_CURRENT_USER
        Case "HKLM": lHive = HKEY_LOCAL_MACHINE
        Case "HKUS": lHive = HKEY_USERS
        Case "HKPD": lHive = HKEY_PERFORMANCE_DATA
        Case "HKCC": lHive = HKEY_CURRENT_CONFIG
        Case "HKDD": lHive = HKEY_DYN_DATA
        Case Else: Exit Sub
    End Select
    sKey = Mid(sArgs(0), 6)
    sVal = sArgs(1)
    sData = Replace(sArgs(2), "\0", Chr(0))
    
    If RegCreateKeyEx(lHive, sKey, 0, vbNullString, 0, KEY_CREATE_SUB_KEY Or KEY_SET_VALUE, ByVal 0, hKey, 0) = 0 Then
        If RegSetValueEx(hKey, sVal, 0, REG_MULTI_SZ, ByVal sData, Len(sData)) <> 0 Then
            Logg "Failed: RegSetMultiValue " & sCmd & " (unable to write to Registry)"
        Else
            Logg "Success: RegSetMultiValue " & sCmd
        End If
        RegCloseKey hKey
    Else
        Logg "Failed: RegSetMultiValue " & sCmd & " (unable to open/create Registry key)"
    End If
End Sub

Public Sub RegSetExpandValue(sCmd$)
    'RegSetExpandValue <hive\key>|<value>|<data>
    Dim hKey&, lHive&, sKey$, sVal$, sData$, sArgs$()
    sArgs = Split(sCmd, "|")
    If UBound(sArgs) <> 2 Then Exit Sub
    Select Case Left(sArgs(0), 4)
        Case "HKCR": lHive = HKEY_CLASSES_ROOT
        Case "HKCU": lHive = HKEY_CURRENT_USER
        Case "HKLM": lHive = HKEY_LOCAL_MACHINE
        Case "HKUS": lHive = HKEY_USERS
        Case "HKPD": lHive = HKEY_PERFORMANCE_DATA
        Case "HKCC": lHive = HKEY_CURRENT_CONFIG
        Case "HKDD": lHive = HKEY_DYN_DATA
        Case Else: Exit Sub
    End Select
    sKey = Mid(sArgs(0), 6)
    sVal = sArgs(1)
    sData = Replace(sArgs(2), "#", "%")
    
    If RegCreateKeyEx(lHive, sKey, 0, vbNullString, 0, KEY_CREATE_SUB_KEY Or KEY_SET_VALUE, ByVal 0, hKey, 0) = 0 Then
        If RegSetValueEx(hKey, sVal, 0, REG_EXPAND_SZ, ByVal sData, Len(sData)) <> 0 Then
            Logg "Failed: RegSetExpandValue " & sCmd & " (unable to write to Registry)"
        Else
            Logg "Success: RegSetExpandValue " & sCmd
        End If
        RegCloseKey hKey
    Else
        Logg "Failed: RegSetExpandValue " & sCmd & " (unable to open/create Registry key)"
    End If
End Sub

Public Sub RegDelValue(sCmd$)
    'RegDelValue <hive\key>|<value>
    Dim hKey&, lHive&, sKey$, sVal$, sArgs$()
    sArgs = Split(sCmd, "|")
    If UBound(sArgs) <> 1 Then Exit Sub
    Select Case Left(sArgs(0), 4)
        Case "HKCR": lHive = HKEY_CLASSES_ROOT
        Case "HKCU": lHive = HKEY_CURRENT_USER
        Case "HKLM": lHive = HKEY_LOCAL_MACHINE
        Case "HKUS": lHive = HKEY_USERS
        Case "HKPD": lHive = HKEY_PERFORMANCE_DATA
        Case "HKCC": lHive = HKEY_CURRENT_CONFIG
        Case "HKDD": lHive = HKEY_DYN_DATA
        Case Else: Exit Sub
    End Select
    sKey = Mid(sArgs(0), 6)
    sVal = sArgs(1)
    
    If Not RegKeyExists(lHive, sKey) Then
        Logg "Failed: RegDelValue " & sCmd & " (key not found)"
        Exit Sub
    End If
    If RegOpenKeyEx(lHive, sKey, 0, KEY_SET_VALUE, hKey) = 0 Then
        If InStr(sVal, "*") = 0 And InStr(sVal, "?") = 0 Then
            If Not RegValueExists(lHive, sKey, sVal) Then Exit Sub
            If RegDeleteValue(hKey, sVal) <> 0 Then
                Logg "Failed: RegDelValue " & sCmd & " (unable to write to Registry)"
            Else
                Logg "Success: RegDelValue " & sCmd
            End If
            RegCloseKey hKey
        Else
            Dim sVals$(), i&
            sVals = Split(RegEnumValues(lHive, sKey), Chr(0))
            For i = 0 To UBound(sVals) Step 2
                If LCase(sVals(i)) Like LCase(sVal) Then
                    If RegDeleteValue(hKey, sVals(i)) <> 0 Then
                        Logg "Failed: RegDelValue " & sCmd & ", value " & sVals(i) & " (unable to write to Registry)"
                    Else
                        Logg "Success: RegDelValue " & sCmd
                    End If
                End If
            Next i
            RegCloseKey hKey
        End If
    Else
        Logg "Failed: RegDelValue " & sCmd & " (unable to open Registry key)"
    End If
End Sub

Public Sub RegDelFromValue(sCmd$)
    'RegDelFromValue <hive\key>|<value>|<data>
    'only possible for string values ofcourse
    Dim hKey&, lHive&, sKey$, sVal$, sData$, lType&, lDataLen&, sData2$, sArgs$()
    sArgs = Split(sCmd, "|")
    If UBound(sArgs) <> 2 Then Exit Sub
    Select Case Left(sArgs(0), 4)
        Case "HKCR": lHive = HKEY_CLASSES_ROOT
        Case "HKCU": lHive = HKEY_CURRENT_USER
        Case "HKLM": lHive = HKEY_LOCAL_MACHINE
        Case "HKUS": lHive = HKEY_USERS
        Case "HKPD": lHive = HKEY_PERFORMANCE_DATA
        Case "HKCC": lHive = HKEY_CURRENT_CONFIG
        Case "HKDD": lHive = HKEY_DYN_DATA
        Case Else: Exit Sub
    End Select
    sKey = Mid(sArgs(0), 6)
    sVal = sArgs(1)
    sData2 = sArgs(2)
    
    If Not RegKeyExists(lHive, sKey) Then
        Logg "Failed: RegDelFromValue " & sCmd & " (key not found)"
        Exit Sub
    End If
    
    If InStr(sVal, "*") = 0 And InStr(sVal, "?") = 0 Then
        If Not RegValueExists(lHive, sKey, sVal) Then
            Logg "Failed: RegDelFromValue " & sCmd & " (value not found)"
            Exit Sub
        End If
        sData = RegGetString(lHive, sKey, sVal)
        If sData <> vbNullString And InStr(sData, sData2) > 0 Then
            sData = Replace(TrimNull(sData), sData2, vbNullString, , , vbTextCompare)
            RegSetString lHive, sKey, sVal, sData
            If RegGetString(lHive, sKey, sVal) <> sData Then
                Logg "Failed: RegDelFromValue " & sCmd & " (unable to write to Registry)"
            Else
                Logg "Success: RegDelFromValue " & sCmd
            End If
        Else
            'empty or failed
            Logg "Failed: RegDelFromValue " & sCmd & " (value is missing or does not contain target data)"
        End If
    Else
        Dim sVals$(), i&
        sVals = Split(RegEnumValues(lHive, sKey), Chr(0))
        For i = 0 To UBound(sVals) Step 2
            If LCase(sVals(i)) Like LCase(sVal) Then
                RegDelFromValue Left(sCmd, 5) & sKey & "|" & sVals(i) & "|" & sData2
            End If
        Next i
    End If
End Sub

Public Sub RegRenameValue(sCmd$)
    'RegRenameValue <hive\key>|<value>|<newvalue>
    'possible for data types string/dword/binary
    Dim hKey&, lHive&, sKey$, sVal$, sVal2$, sData$
    Dim lType&, lDataLen&, lData&, uData() As Byte, lRet&, sArgs$()
    sArgs = Split(sCmd, "|")
    If UBound(sArgs) <> 2 Then Exit Sub
    Select Case Left(sArgs(0), 4)
        Case "HKCR": lHive = HKEY_CLASSES_ROOT
        Case "HKCU": lHive = HKEY_CURRENT_USER
        Case "HKLM": lHive = HKEY_LOCAL_MACHINE
        Case "HKUS": lHive = HKEY_USERS
        Case "HKPD": lHive = HKEY_PERFORMANCE_DATA
        Case "HKCC": lHive = HKEY_CURRENT_CONFIG
        Case "HKDD": lHive = HKEY_DYN_DATA
        Case Else: Exit Sub
    End Select
    sKey = Mid(sArgs(0), 6)
    sVal = sArgs(1)
    sVal2 = sArgs(2)
    
    If Not RegKeyExists(lHive, sKey) Then
        Logg "Failed: RegRenameValue " & sCmd & " (key not found)"
        Exit Sub
    End If
    If Not RegValueExists(lHive, sKey, sVal) Then
        Logg "Failed: RegRenameValue " & sCmd & " (source value not found)"
        Exit Sub
    End If
    If RegValueExists(lHive, sKey, sVal2) Then
        Logg "Failed: RegRenameValue " & sCmd & " (target value already exists)"
        Exit Sub
    End If
    
    If RegOpenKeyEx(lHive, sKey, 0, KEY_QUERY_VALUE, hKey) = 0 Then
        ReDim uData(0)
        RegQueryValueEx hKey, sVal, 0, lType, uData(0), lDataLen
        Select Case lType
            Case REG_SZ
                sData = String(lDataLen, 0)
                lRet = RegQueryValueEx(hKey, sVal, 0, lType, ByVal sData, lDataLen)
                sData = TrimNull(sData)
            Case REG_DWORD
                lRet = RegQueryValueEx(hKey, sVal, 0, lType, lData, 4)
            Case REG_BINARY
                ReDim uData(lDataLen)
                lRet = RegQueryValueEx(hKey, sVal, 0, lType, uData(0), lDataLen)
            'Case 0
            '    RegCloseKey hKey
            '    Exit Sub
            Case Else
                Logg "Failed: RegRenameValue " & sCmd & " (unsupported data type)"
                RegCloseKey hKey
                Exit Sub
        End Select
        RegCloseKey hKey
        If lRet <> 0 Then
            Logg "Failed: RegRenameValue " & sCmd & " (unable to read from Registry)"
        End If
        
        If RegOpenKeyEx(lHive, sKey, 0, KEY_SET_VALUE, hKey) = 0 Then
            If RegDeleteValue(hKey, sVal) = 0 Then
                Select Case lType
                    Case REG_SZ:     lRet = RegSetValueEx(hKey, sVal2, 0, REG_SZ, ByVal sData, lDataLen)
                    Case REG_DWORD:  lRet = RegSetValueEx(hKey, sVal2, 0, REG_DWORD, lData, 4)
                    Case REG_BINARY: lRet = RegSetValueEx(hKey, sVal2, 0, REG_BINARY, uData(0), lDataLen)
                End Select
                If lRet <> 0 Then
                    Logg "Failed: RegRenameValue " & sCmd & " (unable to write to Registry)"
                Else
                    Logg "Success: RegRenameValue " & sCmd
                End If
            Else
                Logg "Failed: RegRenameValue " & sCmd & " (unable to write to Registry)"
            End If
            RegCloseKey hKey
        Else
            Logg "Failed: RegRenameValue " & sCmd & " (unable to open Registry key)"
        End If
    Else
        Logg "Failed: RegRenameValue " & sCmd & " (unable to open Registry key)"
    End If
End Sub

Public Sub RegDelValueIfDataContainsText(sCmd$)
    'RegDelValueIfDataContainsText <hive\key>|<value>|<text>|[case]
    Dim lHive&, sKey$, sVal$, sVal2$, sText$, sData$, iCaseSens%, sArgs$()
    sArgs = Split(sCmd, "|")
    If UBound(sArgs) <> 2 And UBound(sArgs) <> 3 Then Exit Sub
    Select Case Left(sArgs(0), 4)
        Case "HKCR": lHive = HKEY_CLASSES_ROOT
        Case "HKCU": lHive = HKEY_CURRENT_USER
        Case "HKLM": lHive = HKEY_LOCAL_MACHINE
        Case "HKUS": lHive = HKEY_USERS
        Case "HKPD": lHive = HKEY_PERFORMANCE_DATA
        Case "HKCC": lHive = HKEY_CURRENT_CONFIG
        Case "HKDD": lHive = HKEY_DYN_DATA
        Case Else: Exit Sub
    End Select
    sKey = Mid(sArgs(0), 6)
    sVal = sArgs(1)
    sText = sArgs(2)
    If UBound(sArgs) = 3 Then iCaseSens = IIf(CBool(sArgs(3)), vbBinaryCompare, vbTextCompare)
    
    If Not RegKeyExists(lHive, sKey) Then
        Logg "Failed: RegDelValueIfDataContainsText " & sCmd & " (key not found)"
        Exit Sub
    End If
    
    Dim sValues$(), i&
    sValues = Split(RegEnumValues(lHive, sKey), Chr(0))
    If UBound(sValues) > -1 Then
        For i = 0 To UBound(sValues) - 1 Step 2
            sVal2 = sValues(i)
            sData = sValues(i + 1)
            If InStr(sVal, "*") = 0 And InStr(sVal, "?") = 0 Then
                If LCase(sVal2) = LCase(sVal) And InStr(1, sData, sText, iCaseSens) > 0 Then
                    RegDelValue Left(sCmd, 5) & sKey & "|" & sVal2
                End If
            Else
                If LCase(sVal2) Like "*" & LCase(sVal) & "*" And InStr(1, sData, sText, iCaseSens) > 0 Then
                    RegDelValue Left(sCmd, 5) & sKey & "|" & sVal2
                End If
            End If
        Next i
    End If
End Sub

Public Sub RegDelValueIfDataContainsHex(sCmd$)
    'RegDelValueIfDataContainsHex <hive\key>|<value>|<hex>
    Dim lHive&, sKey$, sVal$, sData$
    Dim sHex$, sHexArray$(), sHex2$, i&, sValues$(), sArgs$()
    sArgs = Split(sCmd, "|")
    If UBound(sArgs) <> 2 Then Exit Sub
    Select Case Left(sArgs(0), 4)
        Case "HKCR": lHive = HKEY_CLASSES_ROOT
        Case "HKCU": lHive = HKEY_CURRENT_USER
        Case "HKLM": lHive = HKEY_LOCAL_MACHINE
        Case "HKUS": lHive = HKEY_USERS
        Case "HKPD": lHive = HKEY_PERFORMANCE_DATA
        Case "HKCC": lHive = HKEY_CURRENT_CONFIG
        Case "HKDD": lHive = HKEY_DYN_DATA
        Case Else: Exit Sub
    End Select
    sKey = Mid(sArgs(0), 6)
    sVal = sArgs(1)
    sHex = sArgs(2)
    
    If Not RegKeyExists(lHive, sKey) Then
        Logg "Failed: RegDelValueIfNameContainsHex " & sCmd & " (key not found)"
        Exit Sub
    End If
    
    sHexArray = Split(sHex, ",")
    For i = 0 To UBound(sHexArray)
        sHex2 = sHex2 & Chr(Val("&H" & sHexArray(i)))
    Next i
    
    sValues = Split(RegEnumValues(lHive, sKey), Chr(0))
    If UBound(sValues) > -1 Then
        If InStr(sVal, "*") > 0 Or InStr(sVal, "?") > 0 Then
            For i = 0 To UBound(sValues) - 1 Step 2
                If LCase(sValues(i)) Like LCase(sVal) And _
                   InStr(sValues(i + 1), sHex2) > 0 Then
                    RegDelValue Left(sCmd, 5) & sKey & "|" & sValues(i)
                End If
            Next i
        Else
            For i = 0 To UBound(sValues) - 1 Step 2
                If LCase(sValues(i)) = LCase(sVal) And _
                   InStr(sValues(i + 1), sHex2) > 0 Then
                    RegDelValue Left(sCmd, 5) & sKey & "|" & sValues(i)
                End If
            Next i
        End If
    End If
End Sub

Public Sub RegDelValueIfNameContainsText(sCmd$)
    'RegDelValueIfNameContainsText <hive\key>|<value>|<text>
    Dim lHive&, sKey$, sText$
    Dim sValues$(), i&, sVal$, sVal2$, sData$, sArgs$()
    sArgs = Split(sCmd, "|")
    If UBound(sArgs) <> 2 And UBound(sArgs) <> 3 Then Exit Sub
    Select Case Left(sArgs(0), 4)
        Case "HKCR": lHive = HKEY_CLASSES_ROOT
        Case "HKCU": lHive = HKEY_CURRENT_USER
        Case "HKLM": lHive = HKEY_LOCAL_MACHINE
        Case "HKUS": lHive = HKEY_USERS
        Case "HKPD": lHive = HKEY_PERFORMANCE_DATA
        Case "HKCC": lHive = HKEY_CURRENT_CONFIG
        Case "HKDD": lHive = HKEY_DYN_DATA
        Case Else: Exit Sub
    End Select
    sKey = Mid(sArgs(0), 6)
    sVal = sArgs(1)
    sText = sArgs(2)
    If InStr(sText, "|") > 0 Then sText = Left(sText, InStr(sText, "|") - 1)
    
    If Not RegKeyExists(lHive, sKey) Then
        Logg "Failed: RegDelValueIfNameContainsText " & sCmd & " (key not found)"
        Exit Sub
    End If
    
    sValues = Split(RegEnumValues(lHive, sKey), Chr(0))
    If UBound(sValues) > -1 Then
        For i = 0 To UBound(sValues) - 1 Step 2
            sVal2 = sValues(i)
            'sData = sValues(i + 1)
            If LCase(sVal2) Like "*" & LCase(sVal) & "*" And InStr(1, sVal2, sText, vbTextCompare) > 0 Then
                RegDelValue Left(sCmd, 5) & sKey & "|" & sVal2
            End If
        Next i
    End If
End Sub

Public Sub RegDelValueIfNameContainsHex(sCmd$)
    'RegDelValueIfNameContainsHex <hive\key>|<value>|<hex>
    Dim lHive&, sKey$, sVal$, sVal2$, sData$
    Dim sHex$, sHexArray$(), sHex2$, i&, sValues$(), sArgs$()
    sArgs = Split(sCmd, "|")
    If UBound(sArgs) <> 2 And UBound(sArgs) <> 3 Then Exit Sub
    Select Case Left(sArgs(0), 4)
        Case "HKCR": lHive = HKEY_CLASSES_ROOT
        Case "HKCU": lHive = HKEY_CURRENT_USER
        Case "HKLM": lHive = HKEY_LOCAL_MACHINE
        Case "HKUS": lHive = HKEY_USERS
        Case "HKPD": lHive = HKEY_PERFORMANCE_DATA
        Case "HKCC": lHive = HKEY_CURRENT_CONFIG
        Case "HKDD": lHive = HKEY_DYN_DATA
        Case Else: Exit Sub
    End Select
    sKey = Mid(sArgs(0), 6)
    sVal = sArgs(1)
    sHex = sArgs(2)
    If InStr(sHex, "|") > 0 Then sHex = Left(sHex, InStr(sHex, "|") - 1)
    
    If Not RegKeyExists(lHive, sKey) Then
        Logg "Failed: RegDelValueIfNameContainsHex " & sCmd & " (key not found)"
        Exit Sub
    End If
    
    sHexArray = Split(sHex, ",")
    For i = 0 To UBound(sHexArray)
        sHex2 = sHex2 & Chr(Val("&H" & sHexArray(i)))
    Next i
    
    sValues = Split(RegEnumValues(lHive, sKey), Chr(0))
    If UBound(sValues) > -1 Then
        For i = 0 To UBound(sValues) - 1 Step 2
            If LCase(sValues(i)) Like LCase(sVal) And _
               InStr(sValues(i), sHex2) > 0 Then
                RegDelValue Left(sCmd, 5) & sKey & "|" & sValues(i)
            End If
        Next i
    End If
End Sub

Private Function RegKeyExists(lHive&, sKey$) As Boolean
    'internal function
    Dim hKey&
    If RegOpenKeyEx(lHive, sKey, 0, KEY_QUERY_VALUE, hKey) = 0 Then
        RegKeyExists = True
        RegCloseKey hKey
    End If
End Function

Public Function RegValueExists(lHive&, sKey$, sVal$) As Boolean
    'internal function
    If Not RegKeyExists(lHive, sKey) Then
        RegValueExists = False
        Exit Function
    End If
    
    Dim hKey&, uData() As Byte
    If RegOpenKeyEx(lHive, sKey, 0, KEY_QUERY_VALUE, hKey) = 0 Then
        ReDim uData(260)
        If RegQueryValueEx(hKey, sVal, 0, ByVal 0, uData(0), UBound(uData)) = 0 Then
            RegValueExists = True
        Else
            RegValueExists = False
        End If
        RegCloseKey hKey
    End If
End Function

Public Function RegEnumSubKeys$(lHive&, sKey$)
    'internal function
    Dim hKey&, i&, sName$, sSubKeys$
    If RegOpenKeyEx(lHive, sKey, 0, KEY_ENUMERATE_SUB_KEYS, hKey) <> 0 Then
        Exit Function
    End If
    
    sName = String(260, 0)
    If RegEnumKeyEx(hKey, 0, sName, Len(sName), 0, vbNullString, 0, ByVal 0) <> 0 Then
        RegCloseKey hKey
        Exit Function
    End If
    
    Do
        sName = TrimNull(sName)
        sSubKeys = sSubKeys & sName & "|"
        
        sName = String(260, 0)
        i = i + 1
    Loop Until RegEnumKeyEx(hKey, i, sName, Len(sName), 0, vbNullString, 0, ByVal 0) <> 0
    RegCloseKey hKey
    RegEnumSubKeys = Left(sSubKeys, Len(sSubKeys) - 1)
End Function

Public Function RegEnumValues(lHive&, sKey$)
    'internal function
    Dim hKey&, i&, sList$, sName$, sData$, lType&, uData() As Byte, lDataLen&
    If RegOpenKeyEx(lHive, sKey, 0, KEY_QUERY_VALUE, hKey) <> 0 Then
        Exit Function
    End If
    
    sName = String(MAX_REG_VALUE_NAME, 0)
    ReDim uData(32768)
    lDataLen = UBound(uData)
    If RegEnumValue(hKey, i, sName, Len(sName), 0, lType, uData(0), lDataLen) <> 0 Then
       RegCloseKey hKey
       Exit Function
    End If
    
    Do
        If lType = REG_SZ Then
            sName = TrimNull(sName)
            ReDim Preserve uData(lDataLen)
            sData = TrimNull(StrConv(uData, vbUnicode))
            
            sList = sList & sName & Chr(0) & sData & Chr(0)
        End If
        
        i = i + 1
        sName = String(MAX_REG_VALUE_NAME, 0)
        ReDim uData(32768)
        lDataLen = UBound(uData)
    Loop Until RegEnumValue(hKey, i, sName, Len(sName), 0, lType, uData(0), lDataLen) <> 0
    RegCloseKey hKey
    
    If sList <> vbNullString Then RegEnumValues = Left(sList, Len(sList) - 1)
End Function

Public Function RegGetString$(lHive&, sKey$, sVal$)
    'internal function
    Dim hKey&, sData$, lType&, lDataLen&
    If RegOpenKeyEx(lHive, sKey, 0, KEY_QUERY_VALUE, hKey) = 0 Then
        RegQueryValueEx hKey, sVal, 0, lType, ByVal 0, lDataLen
        If lType = REG_SZ Then
            sData = String(lDataLen, 0)
            RegQueryValueEx hKey, sVal, 0, ByVal 0, ByVal sData, lDataLen
        End If
        RegCloseKey hKey
        RegGetString = TrimNull(sData)
    End If
End Function

Public Sub RegSetString(lHive&, sKey$, sVal$, sData$)
    'internal function
    Dim hKey&
    If RegOpenKeyEx(lHive, sKey, 0, KEY_SET_VALUE, hKey) = 0 Then
        RegSetValueEx hKey, sVal, 0, REG_SZ, ByVal sData, Len(sData) + 1
        RegCloseKey hKey
    End If
End Sub

Public Sub HostsFileResetDatabasePath()
    Dim hKey&
    'internal function for modHostsFile
    Const sDefPath$ = "%SystemRoot%\System32\drivers\etc"
    If RegOpenKeyEx(HKEY_LOCAL_MACHINE, "System\CurrentControlSet\Services\Tcpip\Parameters", 0, KEY_SET_VALUE, hKey) = 0 Then
        RegSetValueEx hKey, "DataBasePath", 0, REG_EXPAND_SZ, ByVal sDefPath, Len(sDefPath)
        RegCloseKey hKey
    End If
End Sub

Public Sub RegSetBFURunOnReboot(sCmd$)
    Dim i%, sKey$, sKey2$, sFile$
    sFile = Mid(sCmd, InStr(sCmd, " ") + 1)
    sKey = "Software\Microsoft\Windows\CurrentVersion\RunOnceEx"
    Do
        i = i + 1
    Loop Until Not RegKeyExists(HKEY_LOCAL_MACHINE, sKey & "\" & Format(i, "000"))
    sKey2 = sKey & "\" & Format(i, "000")
    
    'set runonce command
    RegCreateKey "HKLM\" & sKey2
    RegSetString HKEY_LOCAL_MACHINE, sKey2, "bfurunonceex", "||""" & BuildPath(App.Path, "BFU.exe") & """ " & sFile
    
    If Not RegValueExists(HKEY_LOCAL_MACHINE, sKey2, "bfurunonceex") Then
        Logg "Failed: RegSetBFURunOnReboot " & sCmd & " (unable to write to Registry)"
    Else
        Logg "Success: RegSetBFURunOnReboot " & sCmd
    End If
    
    'BFU.exe depends on the VB runtime dll - but don't seem to actually
    'need to load it (?)
'    RegCreateKey "HKLM\" & sKey & "\Depends"
'    RegSetString HKEY_LOCAL_MACHINE, sKey & "\Depends", Format(i, "000"), "msvbvm60.dll"
'    bRebootNeeded = True
End Sub

Public Sub RegResetPermissions(sCmd$)
    'RegResetPermissions <hive\key>
    Dim lHive&, sKey$
    If Len(sCmd) < 7 Then Exit Sub
    Select Case Left(sCmd, 4)
        Case "HKCR": lHive = HKEY_CLASSES_ROOT
        Case "HKCU": lHive = HKEY_CURRENT_USER
        Case "HKLM": lHive = HKEY_LOCAL_MACHINE
        Case "HKUS": lHive = HKEY_USERS
        Case "HKPD": lHive = HKEY_PERFORMANCE_DATA
        Case "HKCC": lHive = HKEY_CURRENT_CONFIG
        Case "HKDD": lHive = HKEY_DYN_DATA
        Case Else: Exit Sub
    End Select
    sKey = Mid(sCmd, 6)
    
    Dim hKey&
    If InStr(sKey, "*") = 0 And InStr(sKey, "?") = 0 Then
        If Not RegKeyExists(lHive, sKey) Then
            Logg "Failed: RegResetPermissions " & sCmd & " (key not found)"
            Exit Sub
        End If
        EnablePrivilege "SeBackupPrivilege"
        EnablePrivilege "SeRestorePrivilege"
        If RegCreateKeyEx(lHive, sKey & "dummy", 0, vbNullString, 0, KEY_WRITE, ByVal 0, hKey, 0) = 0 Then
            FileDelete sTempDir & "\~bfu.hiv"
            If RegSaveKey(hKey, sTempDir & "\~bfu.hiv", ByVal 0) <> 0 Then
                Logg "Failed: RegResetPermissions " & sCmd & ": unable to write dummy hive"
            End If
            RegCloseKey hKey
            SHDeleteKey lHive, sKey & "dummy"
        End If
        If FileExists(sTempDir & "\~bfu.hiv") Then
            If RegOpenKeyEx(lHive, sKey, 0, KEY_WRITE, hKey) = 0 Then
                If RegRestoreKey(hKey, sTempDir & "\~bfu.hiv", REG_FORCE_RESTORE) <> 0 Then
                    Logg "Failed: RegResetPermissions " & sCmd & ": unable to restore dummy hive"
                Else
                    Logg "Success: RegResetPermissions " & sCmd
                End If
                RegCloseKey hKey
            End If
            FileDelete sTempDir & "\~bfu.hiv"
        End If
    Else
        Dim sKeys$(), i&, sParent$
        sParent = Left(sKey, InStrRev(sKey, "\") - 1)
        sKeys = Split(RegEnumSubKeys(lHive, sParent), "|")
        For i = 0 To UBound(sKeys)
            If LCase(sParent & "\" & sKeys(i)) Like LCase(sKey) Then
                RegResetPermissions Left(sCmd, 5) & sParent & "\" & sKeys(i)
            End If
        Next i
    End If
End Sub

Private Sub EnablePrivilege(sPrivilege$)
    'internal function
    Dim lToken&, liLUID As LUID
    Dim uPriv As TOKEN_PRIVILEGES, uPrivOld As TOKEN_PRIVILEGES
    If OpenProcessToken(GetCurrentProcess, TOKEN_ADJUST_PRIVILEGE Or TOKEN_QUERY, lToken) > 0 Then
        If LookupPrivilegeValue(vbNullString, sPrivilege, liLUID) > 0 Then
            uPriv.PrivilegeCount = 1
            uPriv.Privileges.Attributes = SE_PRIVILEGE_ENABLED
            uPriv.Privileges.pLuid = liLUID
            AdjustTokenPrivileges lToken, 0, uPriv, Len(uPriv), uPrivOld, Len(uPrivOld)
        End If
    End If
End Sub

Public Sub LogIfRegKeyExists(sCmd$)
    'LogIfRegKeyExists <hive\key>
    Dim lHive&, sKey$
    If Len(sCmd) < 7 Then Exit Sub
    Select Case Left(sCmd, 4)
        Case "HKCR": lHive = HKEY_CLASSES_ROOT
        Case "HKCU": lHive = HKEY_CURRENT_USER
        Case "HKLM": lHive = HKEY_LOCAL_MACHINE
        Case "HKUS": lHive = HKEY_USERS
        Case "HKPD": lHive = HKEY_PERFORMANCE_DATA
        Case "HKCC": lHive = HKEY_CURRENT_CONFIG
        Case "HKDD": lHive = HKEY_DYN_DATA
        Case Else: Exit Sub
    End Select
    sKey = Mid(sCmd, 6)
    If InStr(sKey, "?") = 0 And InStr(sKey, "*") = 0 Then
        If RegKeyExists(lHive, sKey) Then
            Logg "Registry key exists: " & sCmd
        End If
    Else
        Dim sKeys$(), i&, sParent$
        sParent = Left(sKey, InStrRev(sKey, "\") - 1)
        sKeys = Split(RegEnumSubKeys(lHive, sParent), "|")
        For i = 0 To UBound(sKeys)
            If LCase(sParent & "\" & sKeys(i)) Like LCase(sKey) Then
                Logg "Registry key exists: " & Left(sCmd, 5) & sParent & "\" & sKeys(i) & " (matches " & sKey & ")"
            End If
        Next i
    End If
End Sub

Public Sub LogIfRegValExists(sCmd$)
    'LogIfRegValExists <hive\key>|<value>
    Dim lHive&, sKey$, sVal$, sArgs$()
    sArgs = Split(sCmd, "|")
    If UBound(sArgs) <> 1 Then Exit Sub
    Select Case Left(sArgs(0), 4)
        Case "HKCR": lHive = HKEY_CLASSES_ROOT
        Case "HKCU": lHive = HKEY_CURRENT_USER
        Case "HKLM": lHive = HKEY_LOCAL_MACHINE
        Case "HKUS": lHive = HKEY_USERS
        Case "HKPD": lHive = HKEY_PERFORMANCE_DATA
        Case "HKCC": lHive = HKEY_CURRENT_CONFIG
        Case "HKDD": lHive = HKEY_DYN_DATA
        Case Else: Exit Sub
    End Select
    sKey = Mid(sArgs(0), 6)
    sVal = sArgs(1)
    
    If Not RegKeyExists(lHive, sKey) Then Exit Sub
    If InStr(sVal, "?") = 0 And InStr(sVal, "*") = 0 Then
        If RegValueExists(lHive, sKey, sVal) Then
            Logg "Registry value found: " & Replace(sCmd, "|", ": ")
        End If
    Else
        Dim sVals$(), i&
        sVals = Split(RegEnumValues(lHive, sKey), Chr(0))
        For i = 0 To UBound(sVals) Step 2
            If LCase(sVals(i)) Like LCase(sVal) Then
                Logg "Registry value found: " & Left(sCmd, 5) & sKey & ": " & sVals(i) & " (matches " & sVal & ")"
            End If
        Next i
    End If
End Sub

Public Sub LogIfRegValContainsText(sCmd$)
    'LogIfRegValContainsText <hive\key>|<value>|<string>
    Dim lHive&, sKey$, sVal$, sText$, sData$, sValues$(), i&, sArgs$()
    sArgs = Split(sCmd, "|")
    If UBound(sArgs) <> 2 Then Exit Sub
    Select Case Left(sArgs(0), 4)
        Case "HKCR": lHive = HKEY_CLASSES_ROOT
        Case "HKCU": lHive = HKEY_CURRENT_USER
        Case "HKLM": lHive = HKEY_LOCAL_MACHINE
        Case "HKUS": lHive = HKEY_USERS
        Case "HKPD": lHive = HKEY_PERFORMANCE_DATA
        Case "HKCC": lHive = HKEY_CURRENT_CONFIG
        Case "HKDD": lHive = HKEY_DYN_DATA
        Case Else: Exit Sub
    End Select
    sKey = Mid(sArgs(0), 6)
    sVal = sArgs(1)
    sText = sArgs(2)
    
    If InStr(sVal, "*") = 0 And InStr(sVal, "?") = 0 Then
        sData = RegGetString(lHive, sKey, sVal)
        If InStr(sData, sText) > 0 Then
            Logg "Registry value contains '" & sText & "': " & Left(sCmd, 5) & sKey & ": " & sVal
        End If
    Else
        sValues = Split(RegEnumValues(lHive, sKey), Chr(0))
        Dim sVal2$
        For i = 0 To UBound(sValues) - 1 Step 2
            sVal2 = sValues(i)
            sData = sValues(i + 1)
            If sVal2 Like "*" & sVal & "*" And InStr(sData, sText) > 0 Then
                Logg "Registry value contains '" & sText & "': " & Left(sCmd, 5) & sKey & ": " & sVal
            End If
        Next i
    End If
End Sub

Public Sub LogIfRegValContainsHex(sCmd$)
    'LogIfRegValContainsHex <hive\key>|<value>|<hex>
    Dim lHive&, sKey$, sVal$, sData$, sHex$, sHexArray$(), i&, sHex2$, sArgs$()
    sArgs = Split(sCmd, "|")
    If UBound(sArgs) <> 2 Then Exit Sub
    Select Case Left(sArgs(0), 4)
        Case "HKCR": lHive = HKEY_CLASSES_ROOT
        Case "HKCU": lHive = HKEY_CURRENT_USER
        Case "HKLM": lHive = HKEY_LOCAL_MACHINE
        Case "HKUS": lHive = HKEY_USERS
        Case "HKPD": lHive = HKEY_PERFORMANCE_DATA
        Case "HKCC": lHive = HKEY_CURRENT_CONFIG
        Case "HKDD": lHive = HKEY_DYN_DATA
        Case Else: Exit Sub
    End Select
    sKey = Mid(sArgs(0), 6)
    sVal = sArgs(1)
    sHex = sArgs(2)
    
    sHexArray = Split(sHex, ",")
    For i = 0 To UBound(sHexArray)
        sHex2 = sHex2 & Chr(Val("&H" & sHexArray(i)))
    Next i
    
    If InStr(sVal, "*") = 0 And InStr(sVal, "?") = 0 Then
        sData = RegGetString(lHive, sKey, sVal)
        If InStr(sData, sHex2) > 0 Then
            Logg "Registry value contains '" & sHex & "': " & Left(sCmd, 5) & sKey & ": " & sVal
        End If
    Else
        Dim sValues$(), sVal2$
        sValues = Split(RegEnumValues(lHive, sKey), Chr(0))
        For i = 0 To UBound(sValues) - 1 Step 2
            sVal2 = sValues(i)
            sData = sValues(i + 1)
            If sVal2 Like "*" & sVal & "*" And InStr(sData, sHex2) > 0 Then
                Logg "Registry value contains '" & sHex & "': " & Left(sCmd, 5) & sKey & ": " & sVal
            End If
        Next i
    End If
End Sub

