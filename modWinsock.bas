Attribute VB_Name = "modWinsock"
Option Explicit
'winsock control, uninstall procotols/namespaces

Private Declare Function WSAStartup Lib "ws2_32.dll" (ByVal wVR As Long, lpWSAD As WSAData) As Long
Private Declare Function WSACleanup Lib "ws2_32.dll" () As Long

Private Declare Function WSAEnumProtocols Lib "ws2_32.dll" Alias "WSAEnumProtocolsA" (ByVal lpiProtocols As Long, lpProtocolBuffer As Any, lpdwBufferLength As Long) As Long
Private Declare Function WSAEnumNameSpaceProviders Lib "ws2_32.dll" Alias "WSAEnumNameSpaceProvidersA" (lpdwBufferLength As Long, lpnspBuffer As Any) As Long
Private Declare Function WSCDeinstallProvider Lib "ws2_32.dll" (ByVal lpProviderId As Long, ByRef lpErrno As Long) As Long
Private Declare Function WSCUnInstallNameSpace Lib "ws2_32.dll" (ByVal lpProviderId As Long) As Long

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef pDest As Any, ByRef pSource As Any, ByVal Length As Long)
Private Declare Function lstrcpy Lib "kernel32.dll" Alias "lstrcpyA" (ByVal lpString1 As String, lpString2 As Any) As String
Private Declare Function StringFromGUID2 Lib "ole32.dll" (rguid As Any, ByVal lpsz As String, ByVal cchMax As Long) As Long

Private Type WSAData
    wVersion As Integer
    wHighVersion As Integer
    szDescription As String * 257
    szSystemStatus As String * 129
    iMaxSockets As Integer
    iMaxUdpDg As Integer
    lpVendorInfo As Long
End Type

Private Type GUID
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(0 To 7) As Byte
End Type

Private Type WSANAMESPACE_INFO
    NSProviderId   As GUID
    dwNameSpace    As Long
    fActive        As Long
    dwVersion      As Long
    lpszIdentifier As Long
End Type

Private Type WSAPROTOCOLCHAIN
    ChainLen As Long
    ChainEntries(6) As Long
End Type

Private Type WSAPROTOCOL_INFO
    dwServiceFlags1 As Long
    dwServiceFlags2 As Long
    dwServiceFlags3 As Long
    dwServiceFlags4 As Long
    dwProviderFlags As Long
    ProviderId As GUID
    dwCatalogEntryId As Long
    ProtocolChain As WSAPROTOCOLCHAIN
    iVersion As Long
    iAddressFamily As Long
    iMaxSockAddr As Long
    iMinSockAddr As Long
    iSocketType As Long
    iProtocol As Long
    iProtocolMaxOffset As Long
    iNetworkByteOrder As Long
    iSecurityScheme As Long
    dwMessageSize As Long
    dwProviderReserved As Long
    szProtocol As String * 256
End Type

Public Sub WinsockKillProtocol(sCmd$)
    'WinsockKillProtocol <mask>
    On Error Resume Next
    Dim sMask$
    Dim uFoundGuid(99) As GUID, i%, j%
    Dim uWSAData As WSAData
    Dim uWSAProtInfo As WSAPROTOCOL_INFO
    Dim uBuffer() As Byte, lBufferSize&
    Dim lNumProtocols&, sLSPName$, lDummy&
    
    sMask = sCmd
    If sMask = vbNullString Then Exit Sub
    
    If WSAStartup(&H202, uWSAData) <> 0 Then
        Logg "Failed: WinsockKillProtocol " & sCmd & " (unable to load Winsock)"
        Exit Sub
    End If
    
    ReDim uBuffer(1)
    WSAEnumProtocols 0, uBuffer(0), lBufferSize
    ReDim uBuffer(lBufferSize - 1)
    
    lNumProtocols = WSAEnumProtocols(0, uBuffer(0), lBufferSize)
    If lNumProtocols <> -1 Then
        For i = 0 To lNumProtocols - 1
            CopyMemory uWSAProtInfo, uBuffer(i * Len(uWSAProtInfo)), Len(uWSAProtInfo)
            sLSPName = TrimNull(uWSAProtInfo.szProtocol)
            
            If InStr(1, sLSPName, sMask, vbTextCompare) > 0 Then
                'match!
                uFoundGuid(j) = uWSAProtInfo.ProviderId
                j = j + 1
            End If
        Next i
    End If
    
    If j > 0 Then
        On Error Resume Next
        For i = 0 To j - 1
            If uFoundGuid(i).Data1 = 0 Then Exit For
            If WSCDeinstallProvider(VarPtr(uFoundGuid(i)), lDummy) <> 0 Then
                Logg "Failed: WinsockKillProtocol " & sCmd & " (operation failed)"
            Else
                Logg "Success: WinsockKillProtocol " & sCmd
            End If
            bRebootNeeded = True
            DoEvents
        Next i
    Else
        Logg "Failed: WinsockKillProtocol " & sCmd & " (protocol not found)"
    End If
    
    Do
    Loop Until WSACleanup() = -1
End Sub

Public Sub WinsockKillNameSpace(sCmd$)
    'WinsockKillNameSpace <mask>
    On Error Resume Next
    Dim sMask$
    sMask = sCmd
    If sMask = vbNullString Then Exit Sub
    
    Dim lNumNameSpace&, sLSPName$, uFoundGuid(99) As GUID, j%
    Dim uWSANameSpaceInfo As WSANAMESPACE_INFO
    Dim uWSAData As WSAData, i%
    Dim uBuffer() As Byte, lBufferSize&
    
    If WSAStartup(&H202, uWSAData) <> 0 Then
        Logg "Failed: WinsockKillNameSpace " & sCmd & " (unable to load Winsock)"
        Exit Sub
    End If

    ReDim uBuffer(1)
    lBufferSize = 0
    WSAEnumNameSpaceProviders lBufferSize, ByVal 0
    ReDim uBuffer(lBufferSize - 1)
    
    lNumNameSpace = WSAEnumNameSpaceProviders(lBufferSize, uBuffer(0))
    If lNumNameSpace <> -1 Then
        For i = 0 To lNumNameSpace - 1
            CopyMemory uWSANameSpaceInfo, uBuffer(i * Len(uWSANameSpaceInfo)), Len(uWSANameSpaceInfo)
            sLSPName = String(255, 0)
            lstrcpy sLSPName, ByVal uWSANameSpaceInfo.lpszIdentifier
            sLSPName = TrimNull(sLSPName)
            
            If InStr(1, sLSPName, sMask, vbTextCompare) > 0 Then
                'match!
                uFoundGuid(j) = uWSANameSpaceInfo.NSProviderId
                j = j + 1
            End If
        Next i
    End If

    If j > 0 Then
        On Error Resume Next
        For i = 0 To j - 1
            If uFoundGuid(i).Data1 = 0 Then Exit For
            If WSCUnInstallNameSpace(VarPtr(uFoundGuid(i))) <> 0 Then
                Logg "Failed: WinsockKillNameSpace " & sCmd & " (operation failed)"
            Else
                Logg "Success: WinsockKillNameSpace " & sCmd
                bRebootNeeded = True
                DoEvents
            End If
        Next i
    Else
        Logg "Failed: WinsockKillNameSpace " & sCmd & " (namespace not found)"
    End If

    Do
    Loop Until WSACleanup() = -1
End Sub

Private Function GuidToString$(uGuid As GUID)
    'internal function
    Dim sGUID$
    sGUID = String(80, 0)
    If StringFromGUID2(uGuid, sGUID, Len(sGUID)) > 0 Then
        GuidToString = StrConv(sGUID, vbFromUnicode)
        GuidToString = TrimNull(GuidToString)
    End If
End Function

