VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Seguranca"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Type STARTUPINFO
    cb As Long
    lpReserved As Long ' !!! must be Long for Unicode string
    lpDesktop As Long  ' !!! must be Long for Unicode string
    lpTitle As Long    ' !!! must be Long for Unicode string
    dwX As Long
    dwY As Long
    dwXSize As Long
    dwYSize As Long
    dwXCountChars As Long
    dwYCountChars As Long
    dwFillAttribute As Long
    dwFlags As Long
    wShowWindow As Integer
    cbReserved2 As Integer
    lpReserved2 As Long
    hStdInput As Long
    hStdOutput As Long
    hStdError As Long
End Type

Private Type PROCESS_INFORMATION
    hProcess As Long
    hThread As Long
    dwProcessId As Long
    dwThreadId As Long
End Type

' CreateProcessWithLogonW API is available only on Windows 2000 and later.
Private Declare Function CreateProcessWithLogonWAPI Lib "advapi32.dll" Alias "CreateProcessWithLogonW" _
        (ByVal lpUsername As String, _
        ByVal lpDomain As String, _
        ByVal lpPassword As String, _
        ByVal dwLogonFlags As Long, _
        ByVal lpApplicationName As Long, _
        ByVal lpCommandLine As String, _
        ByVal dwCreationFlags As Long, _
        ByVal lpEnvironment As Long, _
        ByVal lpCurrentDirectory As String, _
        ByRef lpStartupInfo As STARTUPINFO, _
        ByRef lpProcessInformation As PROCESS_INFORMATION) As Long
Private Declare Function CloseHandleAPI Lib "kernel32.dll" Alias "CloseHandle" (ByVal hObject As Long) As Long
Private Declare Function WaitForSingleObjectAPI Lib "kernel32" Alias "WaitForSingleObject" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
Private Declare Function GetExitCodeProcessAPI Lib "kernel32" Alias "GetExitCodeProcess" (ByVal hProcess As Long, lpExitCode As Long) As Long

' Fonte: https://support.microsoft.com/pt-br/kb/285879
Public Function RunAsUser(ByVal UserName As String, ByVal Password As String, ByVal DomainName As String, ByVal CommandLine As String, ByVal CurrentDirectory As String) As Long

    Const CREATE_DEFAULT_ERROR_MODE = &H4000000
    'Const LOGON_WITH_PROFILE = &H1
    'Const LOGON_NETCREDENTIALS_ONLY = &H2
    Const LOGON_WITH_PORRANENHUMA = 0
    Dim si As STARTUPINFO
    Dim pi As PROCESS_INFORMATION
    Dim wUser As String
    Dim wDomain As String
    Dim wPassword As String
    Dim wCommandLine As String
    Dim wCurrentDir As String
    Dim lResult As Long

    si.cb = Len(si)

    wUser = StrConv(UserName + Chr$(0), vbUnicode)
    wDomain = StrConv(DomainName + Chr$(0), vbUnicode)
    wPassword = StrConv(Password + Chr$(0), vbUnicode)
    wCommandLine = StrConv(CommandLine + Chr$(0), vbUnicode)
    wCurrentDir = StrConv(CurrentDirectory + Chr$(0), vbUnicode)

    lResult = CreateProcessWithLogonWAPI(wUser, wDomain, wPassword, _
          LOGON_WITH_PORRANENHUMA, 0&, wCommandLine, _
          CREATE_DEFAULT_ERROR_MODE, 0&, wCurrentDir, si, pi)
    If lResult <> 0 Then
        CloseHandleAPI pi.hThread
        CloseHandleAPI pi.hProcess
        RunAsUser = 0
    Else
        RunAsUser = Err.LastDllError
    End If

End Function

' Fonte: https://support.microsoft.com/pt-br/kb/285879
' https://support.microsoft.com/en-us/kb/129796
Public Function RunAndWaitAsUser(ByVal UserName As String, ByVal Password As String, ByVal DomainName As String, ByVal CommandLine As String, ByVal CurrentDirectory As String, ByVal TimeOutInMilliseconds As Long, ByRef OUT_ExitCode As Long) As Long

    Const INFINITE = &HFFFFFFFF
    Const CREATE_DEFAULT_ERROR_MODE = &H4000000
    'Const LOGON_WITH_PROFILE = &H1
    'Const LOGON_NETCREDENTIALS_ONLY = &H2
    Const LOGON_WITH_PORRANENHUMA = 0
    Dim si As STARTUPINFO
    Dim pi As PROCESS_INFORMATION
    Dim wUser As String
    Dim wDomain As String
    Dim wPassword As String
    Dim wCommandLine As String
    Dim wCurrentDir As String
    Dim lResult As Long

    OUT_ExitCode = 0
    If TimeOutInMilliseconds = 0 Then TimeOutInMilliseconds = INFINITE
    si.cb = Len(si)

    wUser = StrConv(UserName + Chr$(0), vbUnicode)
    wDomain = StrConv(DomainName + Chr$(0), vbUnicode)
    wPassword = StrConv(Password + Chr$(0), vbUnicode)
    wCommandLine = StrConv(CommandLine + Chr$(0), vbUnicode)
    wCurrentDir = StrConv(CurrentDirectory + Chr$(0), vbUnicode)

    lResult = CreateProcessWithLogonWAPI(wUser, wDomain, wPassword, _
          LOGON_WITH_PORRANENHUMA, 0&, wCommandLine, _
          CREATE_DEFAULT_ERROR_MODE, 0&, wCurrentDir, si, pi)
    If lResult <> 0 Then
        OUT_ExitCode = WaitForSingleObjectAPI(pi.hProcess, TimeOutInMilliseconds)
        GetExitCodeProcessAPI pi.hProcess, OUT_ExitCode
        CloseHandleAPI pi.hThread
        CloseHandleAPI pi.hProcess
        RunAndWaitAsUser = 0
    Else
        RunAndWaitAsUser = Err.LastDllError
    End If

End Function
