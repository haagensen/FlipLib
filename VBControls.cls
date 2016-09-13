VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "VBControls"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Declare Function SendMessageAnyAPI Lib "user32" Alias "SendMessageA" _
    (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) _
    As Long

Public Function ComboFindStr(ByVal hWnd As Long, ByVal sTexto As String, _
                             Optional ByVal bBuscaExata As Boolean = True) As Long

    Const CB_FINDSTRINGEXACT = &H158
    Const CB_FINDSTRING = &H14C

    ComboFindStr = -1
    If sTexto = "" Then Exit Function

    If bBuscaExata Then
        ComboFindStr = SendMessageAnyAPI(hWnd, CB_FINDSTRINGEXACT, -1, ByVal (sTexto))
    Else
        ComboFindStr = SendMessageAnyAPI(hWnd, CB_FINDSTRING, -1, ByVal (sTexto))
    End If

End Function

Public Function ListFindStr(ByVal hWnd As Long, ByVal sTexto As String, _
                            Optional ByVal bBuscaExata As Boolean = True) As Long

    Const LB_FINDSTRINGEXACT = &H1A2
    Const LB_FINDSTRING = &H18F

    ListFindStr = -1
    If sTexto = "" Then Exit Function

    If bBuscaExata Then
        ListFindStr = SendMessageAnyAPI(hWnd, LB_FINDSTRINGEXACT, -1, ByVal (sTexto))
    Else
        ListFindStr = SendMessageAnyAPI(hWnd, LB_FINDSTRING, -1, ByVal (sTexto))
    End If

End Function

'
' cbo foi declarado como Variant para suportar não apenas o ComboBox padrão, como
' também o "OwnerDrawnComboList" do Edoardo Morcillo
'
Public Function ComboFindByItemData(ByVal cbo As Variant, ByVal lValor As Long) As Long

    On Error Resume Next

    Dim i As Long

    ComboFindByItemData = -1

    For i = 0 To cbo.ListCount - 1
        If cbo.ItemData(i) = lValor Then
            ComboFindByItemData = i
            Exit Function
        End If
    Next

End Function

Public Sub SetProgressBarColour(ByVal hWnd As Long, ByVal lCor As Long)
    Const WM_USER = &H400
    Const PBM_SETBARCOLOR As Long = (WM_USER + 9) ' IE4+
    Call SendMessageAnyAPI(hWnd, PBM_SETBARCOLOR, 0&, ByVal lCor)
End Sub

Public Sub SetProgressBarBackColour(ByVal hWnd As Long, ByVal lCor As Long)
    Const CCM_FIRST As Long = &H2000&
    Const CCM_SETBKCOLOR As Long = (CCM_FIRST + 1)
    Const PBM_SETBKCOLOR As Long = CCM_SETBKCOLOR ' IE3+
    Call SendMessageAnyAPI(hWnd, PBM_SETBKCOLOR, 0&, ByVal lCor)
End Sub
