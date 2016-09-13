Attribute VB_Name = "mBrowseForFolders"
Option Explicit

Private Declare Function SendMessageAPI Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

''
' Callback usado pela fun��o "BrowseForFolders" (vide).
' Como todo callback, DEVE estar em um m�dulo .bas
'
' Na inicializa��o, ajusta o diret�rio pr�-selecionado da caixa de di�logo a partir do ponteiro ao path alocado em bInf.lParam,
' passado de volta ao callback como par�metro lpData.
'
Public Function BrowseForFoldersCallbackProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal lParam As Long, ByVal lpData As Long) As Long

    Const BFFM_SETSELECTION = &H466
    Const BFFM_INITIALIZED = 1

    If uMsg = BFFM_INITIALIZED Then
        Call SendMessageAPI(hWnd, BFFM_SETSELECTION, ByVal CLng(1), ByVal lpData)
    End If

    ' A fun��o deve sempre retornar zero
    BrowseForFoldersCallbackProc = 0

End Function
