VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FileInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Declare Function FormatMessageAPI Lib "kernel32" Alias "FormatMessageA" (ByVal dwFlags As Long, lpSource As Any, ByVal dwMessageId As Long, ByVal dwLanguageId As Long, ByVal lpBuffer As String, ByVal nSize As Long, Arguments As Long) As Long

' Usado na função "FileVersion" --------------------------------------
Private Declare Sub CopyMemAPI Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal length As Long)
Private Declare Function VerQueryValueAPI Lib "version.dll" Alias "VerQueryValueA" (pBlock As Any, ByVal lpSubBlock As String, lplpBuffer As Any, puLen As Long) As Long
Private Declare Function GetFileVersionInfoSizeAPI Lib "version.dll" Alias "GetFileVersionInfoSizeA" (ByVal lptstrFilename As String, lpdwHandle As Long) As Long
Private Declare Function GetFileVersionInfoAPI Lib "version.dll" Alias "GetFileVersionInfoA" (ByVal lptstrFilename As String, ByVal dwHandle As Long, ByVal dwLen As Long, lpData As Any) As Long
Private Type VS_FIXEDFILEINFO
   Signature As Long
   StrucVersionl As Integer     '  e.g. = &h0000 = 0
   StrucVersionh As Integer     '  e.g. = &h0042 = .42
   FileVersionMSl As Integer    '  e.g. = &h0003 = 3
   FileVersionMSh As Integer    '  e.g. = &h0075 = .75
   FileVersionLSl As Integer    '  e.g. = &h0000 = 0
   FileVersionLSh As Integer    '  e.g. = &h0031 = .31
   ProductVersionMSl As Integer '  e.g. = &h0003 = 3
   ProductVersionMSh As Integer '  e.g. = &h0010 = .1
   ProductVersionLSl As Integer '  e.g. = &h0000 = 0
   ProductVersionLSh As Integer '  e.g. = &h0031 = .31
   FileFlagsMask As Long        '  = &h3F for version "0.42"
   FileFlags As Long            '  e.g. VFF_DEBUG Or VFF_PRERELEASE
   FileOS As Long               '  e.g. VOS_DOS_WINDOWS16
   FileType As Long             '  e.g. VFT_DRIVER
   FileSubtype As Long          '  e.g. VFT2_DRV_KEYBOARD
   FileDateMS As Long           '  e.g. 0
   FileDateLS As Long           '  e.g. 0
End Type
' fim da parte "FileVersion" --------------------------------------


Private Const INVALID_HANDLE_VALUE As Long = -1
Private Const MAX_PATH As Long = 260

Private Type SYSTEMTIME
    wYear As Integer
    wMonth As Integer
    wDayOfWeek As Integer
    wDay As Integer
    wHour As Integer
    wMinute As Integer
    wSecond As Integer
    wMilliseconds As Integer
End Type

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
   cFileName As String * MAX_PATH
   cAlternate As String * 14
End Type

Private Declare Function CreateFileAPI Lib "kernel32" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, ByVal lpSecurityAttributes As Long, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Private Declare Function SetFileTimeAPI Lib "kernel32" Alias "SetFileTime" (ByVal hFile As Long, lpCreationTime As FILETIME, lpLastAccessTime As FILETIME, lpLastWriteTime As FILETIME) As Long
Private Declare Function SystemTimeToFileTimeAPI Lib "kernel32" Alias "SystemTimeToFileTime" (lpSystemTime As SYSTEMTIME, lpFileTime As FILETIME) As Long
Private Declare Function CloseHandleAPI Lib "kernel32" Alias "CloseHandle" (ByVal hObject As Long) As Long
Private Declare Function LocalFileTimeToFileTimeAPI Lib "kernel32" Alias "LocalFileTimeToFileTime" (lpLocalFileTime As FILETIME, lpFileTime As FILETIME) As Long

Private Declare Function FindFirstFileAPI Lib "kernel32" Alias "FindFirstFileA" (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function FindCloseAPI Lib "kernel32" Alias "FindClose" (ByVal hFindFile As Long) As Long

Private Declare Function PathMatchSpecAPI Lib "shlwapi" Alias "PathMatchSpecW" (ByVal pszFileParam As Long, ByVal pszSpec As Long) As Long
Private Declare Sub SleepAPI Lib "kernel32.dll" Alias "Sleep" (ByVal dwMilliseconds As Long)

Private msUltimoErro As String
Private miNumeroTentativas As Integer

'
'##########################################################################################
' Eventos
'##########################################################################################
'

Public Event ErroCopia(ByVal lCodErro As Long, ByVal sDescricaoErro As String, ByVal iNumeroDaTentativa As Integer)

'
'##########################################################################################
' Propriedades
'##########################################################################################
'

Public Property Get UltimoErro() As String
    UltimoErro = msUltimoErro
End Property
Public Property Let UltimoErro(ByVal sUltimoErro As String)
    msUltimoErro = sUltimoErro
End Property

Public Property Get NumeroTentativas() As Integer
    NumeroTentativas = miNumeroTentativas
End Property
Public Property Let NumeroTentativas(ByVal iNumeroTentativas As Integer)
    If iNumeroTentativas < 1 Then iNumeroTentativas = 1
    miNumeroTentativas = iNumeroTentativas
End Property

'
'##########################################################################################
' Métodos
'##########################################################################################
'

Public Sub FileCopyWithRetries(ByVal sOrigem As String, ByVal sDestino As String)

    Dim i As Long
    Dim bExiste As Boolean
    Dim file As FileInfo

    On Error Resume Next

    msUltimoErro = ""

    Set file = New FileInfo
    bExiste = file.Exists(sOrigem)
    Set file = Nothing

    If Not bExiste Then
        Err.Raise vbObjectError, "FileCopyWithRetries", "Arquivo origem não existe."
    End If

    For i = 1 To miNumeroTentativas

        On Error Resume Next

        ' só pra garantir...
        SetAttr sOrigem, vbNormal
        SetAttr sDestino, vbNormal
        Err.Clear

        If i = miNumeroTentativas Then
            On Error GoTo 0
        End If

        FileCopy sOrigem, sDestino

        Select Case Err.Number

            Case 0 ' Sucesso
                Exit For

            Case Else
                ' Erro 70, permission denied: arquivo está aberto ou não tem permissão de gravação.
                '   Pode tentar gravar um arquivo qq no destino. Se não conseguir, mensagem ao usuário pedindo que informe o erro
                '   ao admin da rede para providenciar as permissões necessárias. Se conseguir, informar que o arquivo está aberto.
                ' Erro 53, path not found: possível erro de rede.
                '   Se não atingiu o total de tentativas, tentar novamente após um segundo. Caso contrário avisar ao usuário,
                '   pedindo para que verifique conectividade com a rede. Se não for path de rede, tentar criar a pasta destino.
                ' Outros erros: pausa de um segundo e tentar novamente.
                RaiseEvent ErroCopia(Err.Number, Err.Description, i)

        End Select

    Next

End Sub

Public Sub KillWithRetries(ByVal sArquivo As String)

    Dim i As Long
    Dim bExiste As Boolean
    Dim file As FileInfo

    On Error Resume Next

    msUltimoErro = ""

    Set file = New FileInfo
    bExiste = file.Exists(sArquivo)
    Set file = Nothing

    If Not bExiste Then
        Err.Raise vbObjectError, "KillWithRetries", "Arquivo informado não existe."
    End If

    For i = 1 To miNumeroTentativas

        On Error Resume Next

        ' só pra garantir...
        SetAttr sArquivo, vbNormal
        Err.Clear

        If i = miNumeroTentativas Then
            On Error GoTo 0
        End If

        Kill sArquivo

        If Err.Number = 0 Then
            Exit For
        Else
            SleepAPI 1000
            DoEvents
        End If

    Next

End Sub

''
' MatchSpec(sOrigem, "*.TIF;*.TIFF;*.PDF")
' If MatchSpec(WFD.cFileName, sMask) Then
Public Function MatchSpec(ByVal sFile As String, ByVal sSpec As String) As Boolean
    msUltimoErro = ""
    MatchSpec = PathMatchSpecAPI(StrPtr(sFile), StrPtr(sSpec))
End Function

Public Function Tipo(ByVal sExtensao As String) As String

    Dim sNome As String
    Dim reg As cRegistry

    On Error GoTo Sai

    msUltimoErro = ""
    Set reg = New cRegistry

    reg.ClassKey = HKEY_CLASSES_ROOT
    If Left$(sExtensao, 1) <> "." Then sExtensao = "." & sExtensao
    reg.SectionKey = sExtensao
    If Not reg.KeyExists() Then GoTo Sai
    sNome = reg.Value

    reg.SectionKey = sNome
    If Not reg.KeyExists() Then GoTo Sai
    Tipo = reg.Value

Sai:
    Set reg = Nothing

End Function

Public Function Exists(ByVal sArquivo As String) As Boolean

    Const ERROR_FILE_NOT_FOUND = 2&
    'Const ERROR_PATH_NOT_FOUND = 3&
    'Const ERROR_ACCESS_DENIED = 5&
    Dim WFD As WIN32_FIND_DATA
    Dim hFile As Long, lErro As Long
    Dim utl As FnUteis: Set utl = New FnUteis

    msUltimoErro = ""

    sArquivo = utl.UnBackSlash(sArquivo)
    hFile = FindFirstFileAPI(sArquivo, WFD)
    lErro = Err.LastDllError
    Exists = (hFile <> INVALID_HANDLE_VALUE)
    Call FindCloseAPI(hFile)

    Select Case lErro
        Case 0
            ' sucesso
        Case ERROR_FILE_NOT_FOUND
            ' não é erro, já que o objetivo da função é checar a existência...
        Case Else
            msUltimoErro = lErro & " " & ObtemMensagemErro(lErro) ' acesso negado (5), etc. ISSO é erro!
    End Select

End Function

''
' Dado o arquivo ARQ, retorna sua extensão, SEM o ponto. Se o arquivo não possuir uma extensão,
' retorna vazio.
'
' @remarks Ex: ObtemExtensao("notas.txt") retorna "txt"
'     ObtemExtensao("arquivo.jpeg") retorna "jpeg"
Public Function Extension(ByVal sArq As String) As String

    ' . Sob Win32 (Win95/98/NT/2000) uma extensão não  precisa ter necessariamente apenas 3
    '   caracteres (por exemplo, arquivos JPG de extensão ".JPEG")
    '
    ' . Um nome de arquivo pode ter pontos entre ele, e  ainda assim ter uma extensão
    '   (por exemplo, "nome.com.pontos.txt")

    Dim i As Integer, pos As Integer

    On Error GoTo Erro

    msUltimoErro = ""

    ' há uma extensão?
    If InStr(sArq, ".") = 0 Then Exit Function
    
    ' obtém posição onde está o ponto delimitador
    For i = Len(sArq) To 1 Step -1
        If Mid$(sArq, i, 1) = "." Then
            pos = i
            Exit For
        End If
    Next
    
    ' retorna extensão
    Extension = Mid$(sArq, pos + 1)

Sai:
Exit Function

Erro:
    Extension = ""
GoTo Sai

End Function

''
' Dado um arquivo ARQ que possua o path, retira este path e retorna apenas o nome do arquivo + extensão.
'
' @remarks Se o nome não possui o path, retorna o próprio nome.
' Se o nome possuir apenas o path, retorna vazio.
' Ex.  C:\WINDOWS\WIN.COM -> WIN.COM
Public Function Name(ByVal sArq As String) As String

    On Error Resume Next

    Dim i As Integer

    msUltimoErro = ""
    Name = sArq
    
    For i = Len(sArq) To 1 Step -1
        If Mid(sArq, i, 1) = "\" Then
            Name = Mid(sArq, i + 1)
            Exit Function
        End If
    Next

End Function

Public Function NameOnly(ByVal sFilename As String) As String

    Dim ext As String
    
    On Error GoTo Erro

    msUltimoErro = ""
    If LenB(sFilename) = 0 Then Exit Function
    ext = Extension(sFilename)
    
    If ext <> "" Then
        NameOnly = Mid$(sFilename, 1, Len(sFilename) - Len(ext) - 1)
    Else
        NameOnly = IIf(Right$(sFilename, 1) = ".", Left$(sFilename, Len(sFilename) - 1), sFilename)
    End If
    
Sai:
Exit Function
Erro:
    NameOnly = ""
GoTo Sai
End Function

Public Function DirectoryName(ByVal sArq As String) As String

    Dim sNSP As String
    Dim cVB As VB6: Set cVB = New VB6

    On Error Resume Next

    msUltimoErro = ""
    sNSP = Me.Name(sArq)
    DirectoryName = cVB.Replace(sArq, sNSP, "")

End Function

Public Sub SetFileDate(ByVal sFilename As String, ByVal dtDate As Date)

    Const GENERIC_WRITE = &H40000000
    Const OPEN_EXISTING = 3
    Const FILE_SHARE_READ = &H1
    Const FILE_SHARE_WRITE = &H2
    Dim udtFileTime As FILETIME
    Dim udtLocalTime As FILETIME
    Dim udtSystemTime As SYSTEMTIME
    Dim lFileHandle As Long

    msUltimoErro = ""

    udtSystemTime.wYear = Year(dtDate)
    udtSystemTime.wMonth = Month(dtDate)
    udtSystemTime.wDay = Day(dtDate)
    udtSystemTime.wDayOfWeek = WeekDay(dtDate) - 1
    udtSystemTime.wHour = Hour(dtDate)
    udtSystemTime.wMinute = Minute(dtDate)
    udtSystemTime.wSecond = Second(dtDate)
    udtSystemTime.wMilliseconds = 0

    SystemTimeToFileTimeAPI udtSystemTime, udtLocalTime
    LocalFileTimeToFileTimeAPI udtLocalTime, udtFileTime

    lFileHandle = CreateFileAPI(sFilename, GENERIC_WRITE, FILE_SHARE_READ Or FILE_SHARE_WRITE, ByVal 0&, OPEN_EXISTING, 0, 0)
    If lFileHandle >= 0 Then
        SetFileTimeAPI lFileHandle, udtFileTime, udtFileTime, udtFileTime
    Else
        Err.Raise vbObjectError, "SetFileDate", sFilename & " não existe."
    End If

    CloseHandleAPI lFileHandle

End Sub

''
' Retorna uma string (formatada) com o número de versão do arquivo sFilename.
' Por exemplo, Version("c:\abc\xpto.exe") retornaria "1.0.0"
'
Public Function Version(ByVal sFilename As String) As String

    Dim lFileHwnd As Long, lBufferLen As Long, lpBuffer As Long, lpuLen As Long
    Dim abytBuffer() As Byte
    Dim tVerInfo As VS_FIXEDFILEINFO
    Dim sStrucVer As String

    msUltimoErro = ""

    'Get the size File version info structure
    lBufferLen = GetFileVersionInfoSizeAPI(sFilename, lFileHwnd)
    If lBufferLen = 0 Then
       Exit Function
    End If

    'Create byte array buffer, then copy memory into structure
    ReDim abytBuffer(lBufferLen)
    Call GetFileVersionInfoAPI(sFilename, 0&, lBufferLen, abytBuffer(0))
    Call VerQueryValueAPI(abytBuffer(0), "\", lpBuffer, lpuLen)
    Call CopyMemAPI(tVerInfo, ByVal lpBuffer, Len(tVerInfo))

    'Determine structure version number (For info only)
    sStrucVer = Format$(tVerInfo.StrucVersionh) & "." & Format$(tVerInfo.StrucVersionl)

    'Concatenate file version number details into a result string
    Version = Format$(tVerInfo.FileVersionMSh) & "." & Format$(tVerInfo.FileVersionMSl, "00") & "."
    If tVerInfo.FileVersionLSh > 0 Then
        Version = Version & Format$(tVerInfo.FileVersionLSh, "0000") & "." & Format$(tVerInfo.FileVersionLSl, "00")
    Else
        Version = Version & Format$(tVerInfo.FileVersionLSl, "0000")
    End If

End Function

Private Function ObtemMensagemErro(ByVal lErro As Long) As String

    Const FORMAT_MESSAGE_FROM_SYSTEM = &H1000
    Dim strMessage As String
    Dim lRet As Long
    strMessage = Space$(4096)
    lRet = FormatMessageAPI(FORMAT_MESSAGE_FROM_SYSTEM, ByVal 0, lErro, 0, strMessage, Len(strMessage), 0)
    ObtemMensagemErro = Left$(strMessage, lRet)

End Function

Private Sub Class_Initialize()
    miNumeroTentativas = 3
End Sub
