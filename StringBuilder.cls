VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "StringBuilder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
''
' Classe para cria��o em alto desempenho de strings.
'
' @author Christian Haagensen Gontijo
'
Option Explicit


''
' Armazena strings.
Private mvsStringArray() As String

''
' Armazena n�mero de elementos no vetor.
Private mlArrayItems As Long

''
' Retorna uma cole��o "Items".
' @return Cole��o.
'
Public Property Get Items() As Collection

    Dim itemsCol As Collection
    Dim i As Integer
    
    Set itemsCol = New Collection
    For i = 0 To mlArrayItems - 1
        Call itemsCol.Add(mvsStringArray(i))
    Next
    Set Items = itemsCol

End Property

''
' Adiciona a string na string sendo constru�da.
' @param newStr String a adicionar.
'
Public Sub Add(ByVal newStr As String)
    ReDim Preserve mvsStringArray(mlArrayItems) As String
    mvsStringArray(mlArrayItems) = newStr
    mlArrayItems = mlArrayItems + 1
End Sub

''
' Atualiza elemento especificado no vetor.
'
' @param elementId N�mero do elemento.
' @param newElementStr Nova string a ser colocada na posi��o especificada.
' @return True se o elemento foi atualizado com sucesso, False se n�o.
'
Public Function UpdateElement(elementId As Long, newElementStr As String) As Boolean
    On Error Resume Next
        If elementId < mlArrayItems And elementId >= 0 Then
            mvsStringArray(elementId) = newElementStr
            UpdateElement = True
        End If
    On Error GoTo 0
End Function

''
' Obt�m determinado elemento do vetor.
'
' @param elementID elemento a obter.
' @return String com o elemento especificado.
'
Public Function GetElement(elementId As Long) As String
    On Error Resume Next
        If elementId < mlArrayItems And elementId >= 0 Then
            GetElement = mvsStringArray(elementId)
        End If
    On Error GoTo 0
End Function

''
' Remove �ltimo elemento do vetor.
' @return True se a opera��o teve sucesso, False se n�o.
'
Public Function RemoveLastElement() As Boolean
    If mlArrayItems > 0 Then
        mlArrayItems = mlArrayItems - 1
        ReDim Preserve mvsStringArray(mlArrayItems) As String
        RemoveLastElement = True
    End If
End Function

''
' Obt�m o n�mero de elementos existentes no vetor.
' @return N�mero de elementos.
'
Public Property Get Elements() As Long
    Elements = mlArrayItems
End Property

''
' Retorna a frase montada como uma string.
' @return Frase SQL montada.
'
Public Property Get ToString() As String
    If mlArrayItems > 0 Then ToString = JoinPF(mvsStringArray, "")
End Property

''
' Retorna a frase montada como uma string delimitada.
'
' @param Delimiter Caractere delimitador a ser utilizado.
' @return Frase SQL montada.
'
Public Function ToDelimitedString(Delimiter As String) As String
    If mlArrayItems > 0 Then ToDelimitedString = JoinPF(mvsStringArray, Delimiter)
End Function

''
' Reinicia o vetor que armazena a frase SQL.
'
Public Sub Reset()
    mlArrayItems = 0
    Erase mvsStringArray
End Sub

''
' Inicializa��o da classe.
'
Private Sub Class_Initialize()
    If mlArrayItems > 0 Then Reset
End Sub

Private Function JoinPF(arr() As String, ByVal Delimiter As String) As String
    Dim i As Long
    For i = LBound(arr) To UBound(arr) - 1
        JoinPF = JoinPF & arr(i) & Delimiter
    Next
    JoinPF = JoinPF & arr(UBound(arr))
End Function
