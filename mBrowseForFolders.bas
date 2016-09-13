Attribute VB_Name = "mBrowseForFolders"
Option Explicit

Private Declare Function SendMessageAPI Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

''
' Callback usado pela função "BrowseForFolders" (vide).
' Como todo callback, DEVE estar em um módulo .bas
'
' Na inicialização, ajusta o diretório pré-selecionado da caixa de diálogo a partir do ponteiro ao path alocado em bInf.lParam,
' passado de volta ao callback como parâmetro lpData.
'
Public Function BrowseForFoldersCallbackProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal lParam As Long, ByVal lpData As Long) As Long

    Const BFFM_SETSELECTION = &H466
    Const BFFM_INITIALIZED = 1

    If uMsg = BFFM_INITIALIZED Then
        Call SendMessageAPI(hWnd, BFFM_SETSELECTION, ByVal CLng(1), ByVal lpData)
    End If

    ' A função deve sempre retornar zero
    BrowseForFoldersCallbackProc = 0

End Function
