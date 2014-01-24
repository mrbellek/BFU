Attribute VB_Name = "modServices"
Option Explicit
'starting, stopping, disabling, deleting of NT services
'TESTED - EVERYTHING WORKS ^_^

Private Declare Function OpenSCManager Lib "advapi32.dll" Alias "OpenSCManagerA" (ByVal lpMachineName As String, ByVal lpDatabaseName As String, ByVal dwDesiredAccess As Long) As Long
Private Declare Function StartService Lib "advapi32.dll" Alias "StartServiceA" (ByVal hService As Long, ByVal lNumServiceArgs As Long, ByVal strArgs As String) As Long
Private Declare Function OpenService Lib "advapi32.dll" Alias "OpenServiceA" (ByVal hSCManager As Long, ByVal lpServiceName As String, ByVal dwDesiredAccess As Long) As Long
Private Declare Function ControlService Lib "advapi32.dll" (ByVal hService As Long, ByVal lControlCode As Long, lpServiceStatus As SERVICE_STATUS) As Long
Private Declare Function ChangeServiceConfig Lib "advapi32.dll" Alias "ChangeServiceConfigA" (ByVal hService As Long, ByVal dwServiceType As Long, ByVal dwStartType As Long, ByVal dwErrorControl As Long, ByVal lpBinaryPathName As String, ByVal lpLoadOrderGroup As String, ByVal lpdwTagID As Long, ByVal lpDependencies As String, ByVal lpServiceStartName As String, ByVal lpPassword As String, ByVal lpDisplayName As String) As Boolean
Private Declare Function DeleteService Lib "advapi32.dll" (ByVal hService As Long) As Long
Private Declare Function CloseServiceHandle Lib "advapi32.dll" (ByVal hSCObject As Long) As Long

Private Type SERVICE_STATUS
    dwServiceType As Long
    dwCurrentState As Long
    dwControlsAccepted As Long
    dwWin32ExitCode As Long
    dwServiceSpecificExitCode As Long
    dwCheckPoint As Long
    dwWaitHint As Long
End Type

Private Const SC_MANAGER_CREATE_SERVICE = &H2
Private Const SC_MANAGER_ENUMERATE_SERVICE = &H4
Private Const SERVICE_QUERY_CONFIG = &H1
Private Const SERVICE_CHANGE_CONFIG = &H2
Private Const SERVICE_QUERY_STATUS = &H4
Private Const SERVICE_ENUMERATE_DEPENDENTS = &H8
Private Const SERVICE_START = &H10
Private Const SERVICE_STOP = &H20
Private Const SERVICE_PAUSE_CONTINUE = &H40
Private Const SERVICE_INTERROGATE = &H80
Private Const SERVICE_USER_DEFINED_CONTROL = &H100
Private Const STANDARD_RIGHTS_REQUIRED = &HF0000
Private Const SERVICE_ALL_ACCESS = (STANDARD_RIGHTS_REQUIRED Or SERVICE_QUERY_CONFIG Or SERVICE_CHANGE_CONFIG Or SERVICE_QUERY_STATUS Or SERVICE_ENUMERATE_DEPENDENTS Or SERVICE_START Or SERVICE_STOP Or SERVICE_PAUSE_CONTINUE Or SERVICE_INTERROGATE Or SERVICE_USER_DEFINED_CONTROL)
Private Const SERVICE_CONTROL_STOP = &H1
Private Const SERVICE_AUTO_START = &H2
Private Const SERVICE_DISABLED = &H4
Private Const SERVICE_NO_CHANGE = &HFFFFFFFF

Public Sub ServiceStart(sCmd$)
    If Not bIsWinNT Then Exit Sub
    'ServiceStart <full/short name>
    Dim hSCManager&, hService&, sServiceName$
    sServiceName = sCmd
    If Not ServiceExists(sServiceName) Then sServiceName = GetServiceShortName(sServiceName)
    If Not ServiceExists(sServiceName) Then
        Logg "Failed: ServiceStart " & sCmd & " (service not found)"
        Exit Sub
    End If
    hSCManager = OpenSCManager(vbNullString, vbNullString, SC_MANAGER_ENUMERATE_SERVICE)
    If hSCManager = 0 Then
        Logg "Failed: ServiceStart " & sCmd & " (unable to get handle from service manager)"
        Exit Sub
    End If
    hService = OpenService(hSCManager, sServiceName, SERVICE_ALL_ACCESS)
    If hService > 0 Then
        If StartService(hService, 0, ByVal 0) = 0 Then
            Logg "Failed: ServiceStart " & sCmd & " (operation failed)"
        Else
            Logg "Success: ServiceStart " & sCmd
        End If
        CloseServiceHandle hService
    Else
        Logg "Failed: ServiceStart " & sCmd & " (unable to open service)"
    End If
    CloseServiceHandle hSCManager
End Sub

Public Sub ServiceStop(sCmd$)
    If Not bIsWinNT Then Exit Sub
    'ServiceStop <full/short name>
    Dim sServiceName$, hSCManager&, hService&, uSS As SERVICE_STATUS
    sServiceName = sCmd
    If Not ServiceExists(sServiceName) Then sServiceName = GetServiceShortName(sServiceName)
    If Not ServiceExists(sServiceName) Then
        Logg "Failed: ServiceStop " & sCmd & " (service not found)"
        Exit Sub
    End If
    hSCManager = OpenSCManager(vbNullString, vbNullString, SC_MANAGER_ENUMERATE_SERVICE)
    If hSCManager = 0 Then
        Logg "Failed: ServiceStop " & sCmd & " (unable to get handle from services manager)"
        Exit Sub
    End If
    hService = OpenService(hSCManager, sServiceName, SERVICE_ALL_ACCESS)
    If hService > 0 Then
        If ControlService(hService, SERVICE_CONTROL_STOP, uSS) = 0 Then
            Logg "Failed: ServiceStop " & sCmd & " (operation failed)"
        Else
            Logg "Success: ServiceStop " & sCmd
        End If
        CloseServiceHandle hService
    Else
        Logg "Failed: ServiceStop " & sCmd & " (unable to open service)"
    End If
    CloseServiceHandle hSCManager
End Sub

Public Sub ServiceDisable(sCmd$)
    If Not bIsWinNT Then Exit Sub
    'ServiceDisable <full/short name>
    Dim sServiceName$, hSCManager&, hService&, uSS As SERVICE_STATUS
    sServiceName = sCmd
    If Not ServiceExists(sServiceName) Then sServiceName = GetServiceShortName(sServiceName)
    If Not ServiceExists(sServiceName) Then
        Logg "Failed: ServiceDisable " & sCmd & " (service not found)"
        Exit Sub
    End If
    hSCManager = OpenSCManager(vbNullString, vbNullString, SC_MANAGER_ENUMERATE_SERVICE)
    If hSCManager = 0 Then
        Logg "Failed: ServiceDisable " & sCmd & " (unable to get handle from services manager)"
        Exit Sub
    End If
    hService = OpenService(hSCManager, sServiceName, SERVICE_ALL_ACCESS)
    If hService > 0 Then
        If ChangeServiceConfig(hService, SERVICE_NO_CHANGE, SERVICE_DISABLED, SERVICE_NO_CHANGE, vbNullString, vbNullString, 0, vbNullString, vbNullString, vbNullString, vbNullString) = 0 Then
            Logg "Failed: ServiceDisable " & sCmd & " (operation failed)"
        Else
            Logg "Success: ServiceDisable " & sCmd
        End If
        CloseServiceHandle hService
    Else
        Logg "Failed: ServiceDisable " & sCmd & " (unable to open service)"
    End If
    CloseServiceHandle hSCManager
End Sub

Public Sub ServiceEnable(sCmd$)
    If Not bIsWinNT Then Exit Sub
    'ServiceEnable <full/short name>
    Dim sServiceName$, hSCManager&, hService&, uSS As SERVICE_STATUS
    sServiceName = sCmd
    If Not ServiceExists(sServiceName) Then sServiceName = GetServiceShortName(sServiceName)
    If Not ServiceExists(sServiceName) Then
        Logg "Failed: ServiceEnable " & sCmd & " (service not found)"
        Exit Sub
    End If
    hSCManager = OpenSCManager(vbNullString, vbNullString, SC_MANAGER_ENUMERATE_SERVICE)
    If hSCManager = 0 Then
        Logg "Failed: ServiceEnable " & sCmd & " (unable to get handle from services manager)"
        Exit Sub
    End If
    hService = OpenService(hSCManager, sServiceName, SERVICE_ALL_ACCESS)
    If hService > 0 Then
        If ChangeServiceConfig(hService, SERVICE_NO_CHANGE, SERVICE_AUTO_START, SERVICE_NO_CHANGE, vbNullString, vbNullString, 0, vbNullString, vbNullString, vbNullString, vbNullString) = 0 Then
            Logg "Failed: ServiceEnable " & sCmd & " (operation failed)"
        Else
            Logg "Success: ServiceEnable " & sCmd
        End If
        CloseServiceHandle hService
    Else
        Logg "Failed: ServiceEnable " & sCmd & " (unable to open service)"
    End If
    CloseServiceHandle hSCManager
End Sub

Public Sub ServiceDelete(sCmd$)
    If Not bIsWinNT Then Exit Sub
    'ServiceDelete <full/short name>
    Dim sServiceName$, hSCManager&, hService&
    sServiceName = sCmd
    If sServiceName = vbNullString Then Exit Sub
    If Not ServiceExists(sServiceName) Then sServiceName = GetServiceShortName(sServiceName)
    If Not ServiceExists(sServiceName) Then
        Logg "Failed: ServiceDelete " & sCmd & " (service not found)"
        Exit Sub
    End If
    hSCManager = OpenSCManager(vbNullString, vbNullString, SC_MANAGER_CREATE_SERVICE)
    If hSCManager > 0 Then
        hService = OpenService(hSCManager, sServiceName, SERVICE_ALL_ACCESS)
        If hService > 0 Then
            If DeleteService(hService) = 0 Then
                Logg "Failed: ServiceDelete " & sCmd & " (operation failed)"
            Else
                Logg "Success: ServiceDelete " & sCmd
                bRebootNeeded = True
            End If
            CloseServiceHandle hService
        Else
            Logg "Failed: ServiceDelete " & sCmd & " (unable to open service)"
        End If
        CloseServiceHandle hSCManager
    Else
        Logg "Failed: ServiceDelete " & sCmd & " (unable to get handle from services manager)"
    End If
End Sub

Private Function GetServiceShortName$(sDisplayName$)
    'internal function
    Dim sServices$(), i&, sName$
    sServices = Split(RegEnumSubKeys(HKEY_LOCAL_MACHINE, "System\CurrentControlSet\Services"), "|")
    If UBound(sServices) = -1 Then Exit Function
    
    For i = 0 To UBound(sServices)
        sName = RegGetString(HKEY_LOCAL_MACHINE, "System\CurrentControlSet\Services\" & sServices(i), "DisplayName")
        If sName = sDisplayName Then
            GetServiceShortName = sServices(i)
            Exit For
        End If
    Next i
End Function

Private Function ServiceExists(sServiceName$) As Boolean
    If Not bIsWinNT Then Exit Function
    'internal function
    Dim hSCManager&, hService&
    If sServiceName = vbNullString Then Exit Function
    hSCManager = OpenSCManager(vbNullString, vbNullString, SC_MANAGER_ENUMERATE_SERVICE)
    hService = OpenService(hSCManager, sServiceName, SERVICE_ALL_ACCESS)
    If hService > 0 Then
        ServiceExists = True
        CloseServiceHandle hService
    Else
        ServiceExists = False
    End If
    CloseServiceHandle hSCManager
End Function
