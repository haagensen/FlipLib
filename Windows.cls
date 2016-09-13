VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Windows"
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

Private Declare Function CreateProcessAPI Lib "kernel32" Alias "CreateProcessA" _
      (ByVal lpApplicationName As String, ByVal lpCommandLine As String, _
      ByVal lpProcessAttributes As Long, ByVal lpThreadAttributes As Long, _
      ByVal bInheritHandles As Long, ByVal dwCreationFlags As Long, _
      ByVal lpEnvironment As Long, ByVal lpCurrentDirectory As String, _
      lpStartupInfo As STARTUPINFO, _
      lpProcessInformation As PROCESS_INFORMATION) As Long

Private Declare Function CloseHandleAPI Lib "kernel32.dll" Alias "CloseHandle" (ByVal hObject As Long) As Long
Private Declare Function WaitForSingleObjectAPI Lib "kernel32" Alias "WaitForSingleObject" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
Private Declare Function GetExitCodeProcessAPI Lib "kernel32" Alias "GetExitCodeProcess" (ByVal hProcess As Long, lpExitCode As Long) As Long

Private Declare Function FindExecutableAPI Lib "shell32" Alias "FindExecutableA" _
  (ByVal lpFile As String, _
   ByVal lpDirectory As String, _
   ByVal sResult As String) As Long

Private Declare Function SendMessageAPI Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function GetDesktopWindowAPI Lib "user32" Alias "GetDesktopWindow" () As Long
Private Declare Function GetWindowAPI Lib "user32" Alias "GetWindow" (ByVal hWnd As Long, ByVal wCmd As Long) As Long
Private Declare Function GetWindowTextAPI Lib "user32" Alias "GetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Private Declare Function GetClassNameAPI Lib "user32" Alias "GetClassNameA" (ByVal hWnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long

Public Sub CloseWindow(ByVal hWnd As Long)
    Const WM_CLOSE = &H10
    Call SendMessageAPI(hWnd, WM_CLOSE, 0, ByVal 0&)
End Sub

Public Function SetAsTopMost(ByVal hWnd As Long, ByVal bTopmost As Boolean) As Long

    On Error Resume Next

    Const SWP_NOMOVE = 2
    Const SWP_NOSIZE = 1
    Const HWND_TOPMOST = -1
    Const HWND_NOTOPMOST = -2

    If bTopmost Then
        SetAsTopMost = SetWindowPos(hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE)
    Else
        SetAsTopMost = SetWindowPos(hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE)
    End If

End Function

''
' Pesquisa por uma janela cujo texto (caption) contenha o texto informado.
'
' @param sTextoProcurado Texto da janela procurada (pode ser parte do texto)
' @param OUT_sWindowText Parâmetro somente de saída. Conterá o texto da janela encontrada.
' @return Zero se a janela não for encontrada, ou, se for, o handle da mesma.
'
Public Function FindWindowLike(ByVal sTextoProcurado As String, ByRef OUT_sWindowText As String) As Long

    Const GW_CHILD = 5
    Const GW_HWNDNEXT = 2
    Dim hWndInicial As Long
    Dim hWnd As Long, sCaption As String, r As Long

    On Error GoTo Sai

    OUT_sWindowText = ""
    hWndInicial = GetDesktopWindowAPI()

    ' Obtém primeira janela filha
    hWnd = GetWindowAPI(hWndInicial, GW_CHILD)

    Do Until hWnd = 0

        ' Obtém nome da janela
        sCaption = Space$(255)
        r = GetWindowTextAPI(hWnd, sCaption, 255)
        sCaption = Left$(sCaption, r)

        If sCaption <> "" Then
            If InStr(1, sCaption, sTextoProcurado, vbTextCompare) > 0 Then
                OUT_sWindowText = sCaption
                FindWindowLike = hWnd
                Exit Do
            End If
        End If

        hWnd = GetWindowAPI(hWnd, GW_HWNDNEXT)
    Loop

Sai:
End Function

Public Function GetWindows(ByVal sTextoProcurado As String) As CInfosJanelas

    Const GW_CHILD = 5
    Const GW_HWNDNEXT = 2
    Dim hWndInicial As Long
    Dim obj As CInfosJanela, col As CInfosJanelas
    Dim hWnd As Long, sClasse As String, sCaption As String, r As Long

    On Error GoTo Sai

    hWndInicial = GetDesktopWindowAPI()
    hWnd = GetWindowAPI(hWndInicial, GW_CHILD)

    Do Until hWnd = 0

        sCaption = Space$(255)
        r = GetWindowTextAPI(hWnd, sCaption, 255)
        sCaption = Left$(sCaption, r)

        If sCaption <> "" Then
            If InStr(1, sCaption, sTextoProcurado, vbTextCompare) > 0 Then

                sClasse = Space(255)
                r = GetClassNameAPI(hWnd, sClasse, 255)
                sClasse = Left$(sClasse, r)

                Set obj = New CInfosJanela
                obj.Caption = sCaption
                obj.Classe = sClasse
                obj.Handle = hWnd

                If col Is Nothing Then Set col = New CInfosJanelas
                col.Add obj

            End If
        End If

        hWnd = GetWindowAPI(hWnd, GW_HWNDNEXT)
    Loop

    Set GetWindows = col

Sai:
End Function

Public Function ObtemAppAssociado(ByVal sDocumento As String) As String

    Const NOMEROTINA As String = "FlipLib.Windows::ObtemAppAssociado"
    Const MAX_PATH = 260
    Dim sResult As String
    Dim ret As Long
    Dim sMsg As String
    Dim utl As FnUteis

    sResult = String$(MAX_PATH, vbNull)

    ret = FindExecutableAPI(sDocumento, "", sResult)
    Select Case ret ' https://msdn.microsoft.com/en-us/library/windows/desktop/bb776419%28v=vs.85%29.aspx
        Case 2:  sMsg = "o arquivo não foi encontrado."
        Case 3:  sMsg = "o caminho informado é inválido."
        Case 5:  sMsg = "o arquivo especificado não pode ser acessado."
        Case 8:  sMsg = "sistema sem memória ou recursos."
        Case 31: sMsg = "não existe associação entre este tipo de arquivo e um arquivo " & _
                        "executável: o Windows não sabe como abrir arquivos deste tipo."
        Case Is < 32: sMsg = "FindExecutable retornou o código " & ret & "."
    End Select
    If ret < 32 Then
        Err.Raise vbObjectError + 1, NOMEROTINA, sMsg
    End If
    Set utl = New FnUteis
    ObtemAppAssociado = utl.StripNulls(sResult)

End Function

' De <http://forums.codeguru.com/showthread.php?457148-Advantages-of-ShellExecuteEx-over-CreateProcess>:
'
' CreateProcess is a Windows kernel-API implemented in kernel32.dll while ShellExecuteEx is a
' shell-API implemented in shell32.dll. In other words, ShellExecuteEx will eventually use CreateProcess.
' In doing so, ShellExecuteEx will make some things simple for the user. For example: you can ShellExecuteEx
' a TXT file, and that will open up in Notepad or the user's preferred Text-File Viewer, automatically. If you
' wish to use CreateProcess for opening a non-executable file, you need to first lookup the registry to find
' the registered application for that file-extension and then use CreateProcess with the correct command-line
' parameters. ShellExecuteEx did all that for you. That apart ShellExecuteEx can display properties, print,
' etc (depending on SHELLEXECUTEINFO::lpVerb parameter).
'
' For its advantages, ShellExecuteEx however has the slight disadvantage that its usage is not very intuitive.
' The handle to the process started by it is not always stored in the SHELLEXECUTEINFO::hProcess member. To get
' the process handle from ShellExecuteEx, you need to explicitly supply the mask SEE_MASK_NOCLOSEPROCESS (which
' given its interesting name adds the the lack of intuitiveness of the APIs usage).
'
' Purists may say that it is possible to replace the Windows Shell and hence a reason to not rely on it via
' shell APIs like ShellExecuteEx. However, IMO, this theoretical possibility is anything but prevalent (I' m
' still to see one), if not also difficult to engineer given that the Windows Shell is installed by Microsoft
' automatically via the installation of the OS, is a protected file and used by Microsoft applications aplenty.
'
' All told, ShellExecuteEx makes programming simpler than CreateProcess (as it develops on the latter) and
' hence should be used.

Public Function Run(ByVal CommandLine As String) As Long

    Const NORMAL_PRIORITY_CLASS = &H20&
    Dim si As STARTUPINFO
    Dim pi As PROCESS_INFORMATION
    Dim lResult As Long

    si.cb = Len(si)

    ' https://msdn.microsoft.com/en-us/library/windows/desktop/ms682425(v=vs.85).aspx
    lResult = CreateProcessAPI(vbNullString, CommandLine, 0&, 0&, 1&, NORMAL_PRIORITY_CLASS, 0&, vbNullString, si, pi)
    If lResult <> 0 Then
        ' Sucesso
        CloseHandleAPI pi.hThread
        CloseHandleAPI pi.hProcess
        Run = 0
    Else
        Run = Err.LastDllError
    End If

End Function

Public Function RunAndWait(ByVal CommandLine As String, ByVal TimeOutInMilliseconds As Long, ByRef OUT_ExitCode As Long) As Long

    ' Const WAIT_FAILED = &HFFFFFFFF 'Our WaitForSingleObject failed to wait and returned -1
    ' Const WAIT_OBJECT_0 = &H0& 'The waitable object got signaled
    ' Const WAIT_ABANDONED = &H80& 'We got out of the waitable object
    ' Const WAIT_TIMEOUT = &H102& 'the interval we used, timed out.
    Const INFINITE = &HFFFFFFFF
    Const NORMAL_PRIORITY_CLASS = &H20&
    Dim si As STARTUPINFO
    Dim pi As PROCESS_INFORMATION
    Dim lResult As Long

    OUT_ExitCode = 0
    If TimeOutInMilliseconds = 0 Then TimeOutInMilliseconds = INFINITE
    si.cb = Len(si)

    ' https://msdn.microsoft.com/en-us/library/windows/desktop/ms682425(v=vs.85).aspx
    lResult = CreateProcessAPI(vbNullString, CommandLine, 0&, 0&, 1&, NORMAL_PRIORITY_CLASS, 0&, vbNullString, si, pi)
    If lResult <> 0 Then
        ' Sucesso ao chamar CreateProcess
        OUT_ExitCode = WaitForSingleObjectAPI(pi.hProcess, TimeOutInMilliseconds)
        ' OUT_ExitCode = 0 significa sucesso
        GetExitCodeProcessAPI pi.hProcess, OUT_ExitCode
        CloseHandleAPI pi.hThread
        CloseHandleAPI pi.hProcess
        RunAndWait = 0
    Else
        ' lResult=0 significa ERRO ao chamar CreateProcess!
        ' A função retorna o código do erro (portanto, um número maior que zero)
        RunAndWait = Err.LastDllError
    End If

End Function
