VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DirectoryInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Type TBROWSEINFO
    hOwner As Long
    pidlRoot As Long
    pszDisplayName As String
    lpszTitle As String
    ulFlags As Long
    lpfn As Long
    lParam As Long
    iImage As Long
End Type
Private Declare Function SHBrowseForFolderAPI Lib "shell32.dll" Alias "SHBrowseForFolderA" (lpBrowseInfo As TBROWSEINFO) As Long
Private Declare Function LocalAllocAPI Lib "kernel32" Alias "LocalAlloc" (ByVal uFlags As Long, ByVal uBytes As Long) As Long
Private Declare Function LocalFreeAPI Lib "kernel32" Alias "LocalFree" (ByVal hMem As Long) As Long
Private Declare Sub CopyMemoryAPI Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, pSource As Any, ByVal dwLength As Long)
Private Declare Function SHGetPathFromIDListAPI Lib "shell32.dll" Alias "SHGetPathFromIDListA" (ByVal pidl As Long, ByVal pszPath As String) As Long
Private Declare Sub CoTaskMemFreeAPI Lib "ole32.dll" Alias "CoTaskMemFree" (ByVal hMem As Long)

Private Declare Function FormatMessageAPI Lib "kernel32" Alias "FormatMessageA" (ByVal dwFlags As Long, lpSource As Any, ByVal dwMessageId As Long, ByVal dwLanguageId As Long, ByVal lpBuffer As String, ByVal nSize As Long, Arguments As Long) As Long

Private Const INVALID_HANDLE_VALUE As Long = -1
Private Const MAX_PATH As Long = 260

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

Private Declare Function FindFirstFileAPI Lib "kernel32" Alias "FindFirstFileA" (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function FindNextFileAPI Lib "kernel32" Alias "FindNextFileA" (ByVal hFindFile As Long, lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function FindCloseAPI Lib "kernel32" Alias "FindClose" (ByVal hFindFile As Long) As Long

Private msUltimoErro As String

Public Property Get UltimoErro() As String
    UltimoErro = msUltimoErro
End Property
Public Property Let UltimoErro(ByVal sUltimoErro As String)
    msUltimoErro = sUltimoErro
End Property

Public Function Exists(ByVal sFolder As String) As Boolean

    Const ERROR_FILE_NOT_FOUND = 2&
    Const ERROR_PATH_NOT_FOUND = 3&
    Const FILE_ATTRIBUTE_DIRECTORY As Long = &H10
    Dim WFD As WIN32_FIND_DATA
    Dim hFile As Long, lErro As Long
    Dim utl As FnUteis: Set utl = New FnUteis

    msUltimoErro = ""
    sFolder = utl.UnBackSlash(sFolder)

    hFile = FindFirstFileAPI(sFolder, WFD)
    lErro = Err.LastDllError
    Exists = ((hFile <> INVALID_HANDLE_VALUE) And (WFD.dwFileAttributes And FILE_ATTRIBUTE_DIRECTORY))
    Call FindCloseAPI(hFile)

    Select Case lErro
        Case 0
            ' Sucesso
        Case ERROR_FILE_NOT_FOUND, ERROR_PATH_NOT_FOUND
            ' Ignoramos -> a "n�o exist�ncia" n�o � erro, mas retorno poss�vel
        Case Else
            msUltimoErro = lErro & " " & ObtemMensagemErro(lErro) ' acesso negado (5), etc. ISSO � erro!
    End Select

    Set utl = Nothing

End Function

''
' Dado um path CAMINHO, retorna o drive.
'
' Um drive, para paths mapeados, � (por exemplo) "d:\", mas para
' drives n�o-mapeados (UNC), � preciso incluir o nome do drive da rede.
'
' Exemplo para paths UNC:
' "\\master\disys\testes"   retornaria   "\\master\"
'
' Exemplo para paths mapeados:
' "g:\master\disys\testes"   retornaria   "g:\"
'
Public Function Root(ByVal sFolder As String) As String

    Dim pos As Integer, caracter As String
    Dim utl As FnUteis: Set utl = New FnUteis

    On Error Resume Next

    msUltimoErro = ""

    If Left$(sFolder, 2) = "\\" Then 'UNC
        ' procura pela primeira ocorr�ncia de "\" ap�s a "\\" inicial
        For pos = 3 To Len(sFolder)
            caracter = Mid$(sFolder, pos, 1)
            If caracter = "\" Then Exit For
        Next
        Root = Left$(sFolder, pos)
    ElseIf Mid$(sFolder, 2, 2) = ":\" Then
        ' drive "normal"
        Root = Left$(sFolder, 3)
    ElseIf Right$(sFolder, 1) = ":" Then
        ' "c:" por exemplo
        Root = utl.BackSlash(sFolder)
    Else
        ' n�o inclui drive
        Root = ""
    End If

    Set utl = Nothing

End Function

Public Function Create(ByVal sPath As String) As Boolean

    Dim nPos As Long
    Dim i As Long
    Dim bUNC As Boolean
    Dim sSubpath As String
    Dim cVB As VB6: Set cVB = New VB6
    Dim utl As FnUteis: Set utl = New FnUteis
    'Dim dir As DirectoryInfo: Set dir = New DirectoryInfo

    On Error GoTo Erro

    msUltimoErro = ""
    If sPath = "" Then GoTo Sai

    ' Ignoramos caracteres inv�lidos no path passado
    If Left$(sPath, 2) = "\\" Then
        sPath = Mid$(sPath, 3)
        bUNC = True
    End If
    If Mid$(sPath, 2, 2) = ":\" Then
        sPath = cVB.Replace(sPath, ":\", Chr$(1), , 1) ' uma troca tempor�ria...
    End If
    sPath = utl.StripNulls(sPath)
    sPath = cVB.Replace(sPath, "/", "")
    sPath = cVB.Replace(sPath, "?", "")
    sPath = cVB.Replace(sPath, "*", "")
    sPath = cVB.Replace(sPath, ":", "")
    sPath = cVB.Replace(sPath, """", "")
    sPath = cVB.Replace(sPath, "<", "")
    sPath = cVB.Replace(sPath, ">", "")
    sPath = cVB.Replace(sPath, "|", "")
    Do While True
        If InStr(sPath, "\\") > 0 Then
            sPath = cVB.Replace(sPath, "\\", "\")
        Else
            Exit Do
        End If
    Loop
    If bUNC Then sPath = "\\" & sPath
    sPath = cVB.Replace(sPath, Chr$(1), ":\")
    sPath = utl.BackSlash(sPath)
    If sPath = "\" Then GoTo Sai

    nPos = Posicao_da_primeira_letra_do_dirPF(sPath)
    On Error Resume Next
        For i = nPos To Len(sPath)
            If Mid$(sPath, i, 1) = "\" Then
                sSubpath = Left$(sPath, i - 1)
                If Not Exists(sSubpath) Then
                    MkDir sSubpath ' FolderExists("\\master2\trabalho\") retornando FALSE, a� MkDir causa erro!
                End If
            End If
        Next
    On Error GoTo Erro

    If Exists(sSubpath) Then
        Create = True
    End If

Sai:
    Set utl = Nothing
    Set cVB = Nothing
Exit Function

Erro:
    Err.Raise Err.Number, Err.Source, Err.Description

End Function

Private Function Posicao_da_primeira_letra_do_dirPF(ByVal caminho As String) As Integer

    ' Retorna a posi��o da primeira letra do path (fora o drive).
    ' Um drive, para paths mapeados, � (por exemplo) "d:\", mas para drives n�o-mapeados (UNC), � preciso incluir o nome do drive da rede.
    '
    ' Exemplo para paths UNC:
    ' "\\master\disys\testes"
    ' retornaria a posi��o 10 (o "drive" seria "\\master\")
    '
    ' Exemplo para paths mapeados:
    ' "g:\master\disys\testes"
    ' retornaria a posi��o 4 (o drive seria "g:\")
    
    Dim pos As Integer, caracter As String
    
    On Error Resume Next

    msUltimoErro = ""

    If Left$(caminho, 2) = "\\" Then
        ' UNC
        
        ' procura pela primeira ocorr�ncia de "\" ap�s a "\\" inicial
        For pos = 3 To Len(caminho)
            caracter = Mid$(caminho, pos, 1)
            If caracter = "\" Then Exit For
        Next
        
        ' verifica se ap�s a posi��o encontrada ainda h� caracteres
        If pos + 1 < Len(caminho) Then
            ' sim, ok
            Posicao_da_primeira_letra_do_dirPF = pos + 1
        Else
            ' n�o, houve algum erro. Retorna a primeira posi��o
            Posicao_da_primeira_letra_do_dirPF = 1
        End If
        
    ElseIf Mid$(caminho, 2, 2) = ":\" Then
        ' mapeado
        Posicao_da_primeira_letra_do_dirPF = 4
        
    Else
        ' n�o inclui drive
        Posicao_da_primeira_letra_do_dirPF = 1
    End If

End Function

''
' Apaga um diret�rio, e seus subdiret�rios/arquivos se houver.
'
Public Sub Delete(ByVal sPath As String)

    Dim WFD As WIN32_FIND_DATA
    Dim hFile As Long, sArq As String
    Dim utl As FnUteis: Set utl = New FnUteis
    'Dim dir As DirectoryInfo: Set dir = New DirectoryInfo

    On Error Resume Next

    msUltimoErro = ""

    If Not Exists(sPath) Then
        Exit Sub
    End If

    sPath = utl.BackSlash(sPath)

    hFile = FindFirstFileAPI(sPath & "*.*", WFD)
    If hFile <> INVALID_HANDLE_VALUE Then
        Do
            sArq = utl.StripNulls(WFD.cFileName)
            If (WFD.dwFileAttributes And vbDirectory) = 0 Then
                SetAttr sPath & sArq, vbNormal
                Kill sPath & sArq
            Else
                If sArq <> "." And sArq <> ".." Then
                    ' recursivo
                    Delete sPath & sArq
                End If
            End If
        Loop While FindNextFileAPI(hFile, WFD)
        hFile = FindCloseAPI(hFile)
    End If

    RmDir sPath

    Set utl = Nothing

End Sub

Public Function MontaPath(ParamArray sPartesDoPath()) As String

    Dim i As Long, j As Long
    Dim bUNC As Boolean
    Dim utl As FnUteis: Set utl = New FnUteis

    On Error Resume Next

    msUltimoErro = ""

    ' ParamArray � passado por refer�ncia, se mudarmos algum par�metro, atrapalhamos o retorno
    ReDim sVet(UBound(sPartesDoPath)) As String
    If Err.Number <> 0 Then Exit Function
    Err.Clear
    For i = 0 To UBound(sPartesDoPath)
        sVet(i) = sPartesDoPath(i)
        If Err.Number <> 0 Then Exit Function
    Next
    Err.Clear

    If Left$(sVet(0), 2) = "\\" Then
        bUNC = True
    End If

    For i = 0 To UBound(sVet)
        Do While True
            If Left$(sVet(i), 1) = "\" Then
                sVet(i) = Mid$(sVet(i), 2)
            Else
                Exit Do
            End If
        Loop
        Do While True
            If Right$(sVet(i), 1) = "\" Then
                sVet(i) = Mid$(sVet(i), 1, Len(sVet(i)) - 1)
            Else
                Exit Do
            End If
        Loop
        If bUNC And i = 0 Then sVet(i) = "\\" & sVet(i)
        If Len(sVet(i)) > 0 Then sVet(i) = utl.BackSlash(sVet(i))
        MontaPath = MontaPath & sVet(i)
    Next

    Set utl = Nothing

End Function

Public Function AsUNC(ByVal p As String, ByRef houveErro As Boolean) As String

    '===============================================================================================
    ' Dada a string P contendo um caminho no sistema de arquivos (ex: c:\winnt\xpto), retorna-a
    ' como um path UNC (ex: \\master\disco1\winnt\xpto). Em caso de erro, o par�metro HOUVEERRO �
    ' retornado com TRUE
    '===============================================================================================

    Dim pathUNC As String, bPrecisaDePrivilegios As Boolean
    
    msUltimoErro = ""
    houveErro = False

    ' por hora, assume que a sa�da � a mesma da entrada
    AsUNC = p

    If Len(p) > 3 Then
        If UCase$(Left$(p, 3)) Like "[A-Z]:\" Then  'formato <drive>:\
            pathUNC = vbNullString
            On Error Resume Next
                pathUNC = GetUncNameGF(p, bPrecisaDePrivilegios)
            On Error GoTo 0
            If pathUNC <> "" Then
                If StrComp(pathUNC, p, vbTextCompare) = 0 Then
                    If bPrecisaDePrivilegios Then
                        msUltimoErro = "Eleva��o necess�ria para obter o path."
                    Else
                        msUltimoErro = "Path indicado n�o est� em um dispositivo compartilhado."
                    End If
                    houveErro = True
                Else
                    ' ok, transformou em UNC!
                    AsUNC = pathUNC
                End If
            End If
        End If
    End If

End Function

''
' Exibe a janela "BrowseForFolders" (provida pelo sistema operacional) para que o usu�rio escolha um diret�rio.
'
' @param sTextoJanela Texto que ser� exibido na caixa de di�logo.
' @param sPathPreselecionado Um path (existente) que ser� pr�-selecionado na janela.
' @param OUT_bCancelado Par�metro apenas de retorno. Indica que o bot�o "cancelar" da janela foi pressionado.
'
' @return String com o diret�rio escolhido pelo usu�rio (incluindo uma "\" no fim).
'   Se o usu�rio cancelar, o path retornado � igual ao par�metro sPathPreselecionado, e OUT_bCancelado � True.
'
Public Function BrowseForFolders(ByVal sTextoJanela As String, _
                                 ByVal sPathPreselecionado As String, _
                                 ByRef OUT_bCancelado As Boolean) As String

    Const BIF_RETURNONLYFSDIRS = &H1
    Const LMEM_FIXED = &H0
    Const LMEM_ZEROINIT = &H40
    Const lPtr = (LMEM_FIXED Or LMEM_ZEROINIT)
    Dim bInf As TBROWSEINFO
    Dim PathID As Long, lpSelPath As Long
    Dim aux As String, sPathEscolhido As String, sRetPath As String
    Dim utl As FnUteis: Set utl = New FnUteis

    msUltimoErro = ""
    BrowseForFolders = ""
    OUT_bCancelado = False

    ' Necess�rio ou a chamada � API n�o funcionar�
    If LenB(sPathPreselecionado) = 0 Then
        sPathPreselecionado = "C:" ' Se for vazio, teremos um GPF
    Else
        sPathPreselecionado = utl.UnBackSlash(sPathPreselecionado)
    End If

    ' A fun��o usa um callback, que permite que pr�-selecionemos, via c�digo, um diret�rio como ponto inicial.
    With bInf
        .hOwner = 0     ' Desktop
        .pidlRoot = 0   ' Raiz = desktop
        .ulFlags = BIF_RETURNONLYFSDIRS
        .lpszTitle = sTextoJanela
        .lpfn = FARPROC(AddressOf BrowseForFoldersCallbackProc)

        ' Alocamos alguma mem�ria para o path pr�-selecionado, jogamos a string nesta mem�ria alocada, e
        ' ajustamos o valor do ponteiro retornado para lParam (poder�amos, mas n�o vou, checar o sucesso de LocalAlloc).
        '
        ' Notas:
        ' 1) "StrPtr" do VB n�o funciona aqui, porque o endere�o de mem�ria de uma vari�vel sai do escopo quando passado para SHBrowseForFolder
        ' 2) Win2000 requer que o bloco de mem�ria inclua algum espa�o extra para o NULL no fim da string.
        lpSelPath = LocalAllocAPI(lPtr, Len(sPathPreselecionado)) + 1
        CopyMemoryAPI ByVal lpSelPath, ByVal sPathPreselecionado, Len(sPathPreselecionado) + 1
        .lParam = lpSelPath
    End With

    ' Exibe a janela e s� retorna quando ela for fechada.
    ' A fun��o callback vai receber as mensagens da janela BrowseForFolders, enquanto a mesma estiver aberta.
    ' PathID vai conter o pidl do path selecionado se a caixa de di�logo n�o for cancelada.
    PathID = SHBrowseForFolderAPI(bInf)

    If PathID <> 0 Then

        ' sRetPath *deve* ser pr�-alocado ou teremos GPFs
        sRetPath = String$(260, 0)

        ' Obt�m o path a partir do pidl do path selecionado via SHBrowseForFolder. Retorna TRUE em caso de sucesso.
        ' Se o usu�rio selecionar um path, o mesmo � armazenado em AUX. Se clicar "cancelar", aux continuar� vazio.
        If SHGetPathFromIDListAPI(PathID, sRetPath) Then
            sPathEscolhido = utl.BackSlash(Left$(sRetPath, InStr(sRetPath, vbNullChar) - 1))
        End If

    Else
        ' PathID=zero => o usu�rio cancelou
        OUT_bCancelado = True
    End If

    ' *Devemos* liberar a mem�ria que o shell alocou para o pidl
    Call CoTaskMemFreeAPI(PathID)

    ' Libera o pidl usado para identificar a pasta origem
    Call CoTaskMemFreeAPI(bInf.pidlRoot)

    ' Libera a mem�ria que o shell alocou para o path pr�-selecionado
    Call LocalFreeAPI(lpSelPath)

    Set utl = Nothing

    BrowseForFolders = sPathEscolhido

End Function

Private Function FARPROC(ByVal pfn As Long) As Long
    FARPROC = pfn
End Function

Private Function ObtemMensagemErro(ByVal lErro As Long) As String

    Const FORMAT_MESSAGE_FROM_SYSTEM = &H1000
    Dim strMessage As String
    Dim lRet As Long
    strMessage = Space$(4096)
    lRet = FormatMessageAPI(FORMAT_MESSAGE_FROM_SYSTEM, ByVal 0, lErro, 0, strMessage, Len(strMessage), 0)
    ObtemMensagemErro = Left$(strMessage, lRet)

End Function
