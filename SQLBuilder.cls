VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SQLBuilder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
''
' Classe para criação de frases SQL.
'
' @author Christian Haagensen Gontijo
'

Option Explicit

''
' Tipos de banco de dados
'
Public Enum AcessoBancoGerenciadores
    Gerenciador_Nenhum = 0
    Gerenciador_Oracle = 1
    Gerenciador_SQLServer = 2
    Gerenciador_MySql = 3
    Gerenciador_Access = 4
    Gerenciador_TCPServer = 5
End Enum

Private meBanco As AcessoBancoGerenciadores

''
' Tipos de join.
'
' @param InnerJoin O Join será um "Inner Join".
' @param LeftJoin O Join será um "Left Outer Join".
' @param RightJoin O Join será um "Right Outer Join".
'
Public Enum EJoin
    InnerJoin
    LeftJoin
    RightJoin
End Enum

''
' Tipos de ordenação.
'
' @param Ascendente Ordenação ascendente.
' @param Descendente Ordenação descendente.
'
Public Enum EOrderBy
    Ascendente
    Descendente
End Enum

''
' Armazena se usaremos a "nova" sintaxe para JOINs (from x inner join y on x.a=y.b...)
' ou a velha (from x,y where x.a=y.b(+)...).
' O Oracle 8, por exemplo, usa a velha, enquanto o SQLServer 2005 usa a nova.
Private mbUsarNovaSintaxeJoin As Boolean

''
' Armazena status de adicionar ou não um "newline" ao final da linha.
Private mbAddNL As Boolean
''
' Armazena cláusula SELECT da frase SQL.
Private msSelect As String
''
' Armazena se teremos um SELECT DISTINCT.
Private mbSelectDistinct As Boolean
''
' Armazena cláusula FROM da frase SQL.
Private msFrom As String
''
' Armazena cláusula WHERE da frase SQL.
Private msWhere As String
''
' Armazena cláusula ORDER BY da frase SQL.
Private msOrderBy As String
''
' Armazena cláusula GROUP BY da frase SQL.
Private msGroupBy As String
''
' Armazena cláusula HAVING da frase SQL.
Private msHaving As String
''
' Armazena cláusula LIMIT da frase SQL.
Private msLimit As String
''
' Armazena cláusula JOIN da frase SQL. Até 11 tabelas podem ser relacionadas.
Private mvsJoin(0 To 10) As String
''
' Armazena tipo de JOIN especificado da frase SQL.
Private mveJoin(0 To 10) As EJoin
''
' Armazena ordenação especificada na frase SQL.
Private meOrder As EOrderBy
''
' Armazena cláusula ON do JOIN especificado da frase SQL.
Private mvsJoinOn(0 To 10) As String
''
' Armazena frase SQL a ser construída via classe StringBuilder.
Private msbFrase As StringBuilder

'
'##############################################################################
' MÉTODOS E PROPRIEDADES
'##############################################################################
'

Public Property Let Limit(ByVal v As String)
    msLimit = v
End Property
Public Property Get Limit() As String
    Limit = msLimit
End Property

Public Property Let Banco(ByVal v As AcessoBancoGerenciadores)
    meBanco = v
End Property
Public Property Get Banco() As AcessoBancoGerenciadores
    Banco = meBanco
End Property

Public Property Let UsarNovaSintaxeJoin(ByVal v As Boolean)
    mbUsarNovaSintaxeJoin = v
End Property
Public Property Get UsarNovaSintaxeJoin() As Boolean
    UsarNovaSintaxeJoin = mbUsarNovaSintaxeJoin
End Property
    

''
' Se definido em True, uma marca de "nova linha" será adicionada ao final de cada cláusula da frase SQL.
'
' @param valor True para adicionar uma nova linha ao final de cada cláusula, False para não adicionar.
'
Public Property Let AdicionarNewLine(ByVal valor As Boolean)
    mbAddNL = valor
End Property

''
' Permite definir a cláusula SELECT da frase SQL.
'
' @param valor Campos que irão compor a cláusula SELECT da frase SQL. A palavra-chave "SELECT" não precisa ser especificada.
'
Public Property Let SQLSelect(ByVal valor As String)

    If UCase$(Left$(valor, 7)) = "SELECT " Then
        msSelect = Mid$(valor, 8)
    Else
        msSelect = valor
    End If

End Property

''
' Retorna a cláusula SELECT da frase SQL.
'
' @return Cláusula SELECT da frase.
'
Public Property Get SQLSelect() As String
    SQLSelect = msSelect
End Property

''
' Permite definir que a cláusula SELECT da frase SQL conterá a propriedade "DISTINCT".
'
' @param valor True para usar Distinct, False para não.
'
Public Property Let SelectDistinct(ByVal valor As Boolean)
    mbSelectDistinct = valor
End Property

''
' Retorna o uso do DISTINCT na cláusula SELECT da frase SQL.
'
' @return True se DISTINCT foi especificado, False se não.
'
Public Property Get SelectDistinct() As Boolean
    SelectDistinct = mbSelectDistinct
End Property

''
' Permite definir a cláusula FROM da frase SQL.
'
' @param valor Campos que irão compor a cláusula FROM da frase SQL. A palavra-chave "FROM" não precisa ser especificada.
'
Public Property Let From(ByVal valor As String)

    If UCase$(Left$(valor, 5)) = "FROM " Then
        msFrom = Mid$(valor, 6)
    Else
        msFrom = valor
    End If

End Property

''
' Retorna a cláusula FROM da frase SQL.
'
' @return Cláusula FROM da frase.
'
Public Property Get From() As String
    From = msFrom
End Property

''
' Permite definir a cláusula WHERE da frase SQL.
'
' @param valor Campos que irão compor a cláusula WHERE da frase SQL. A palavra-chave "WHERE" não precisa ser especificada.
'
Public Property Let Where(ByVal valor As String)
    
    If UCase$(Left$(valor, 6)) = "WHERE " Then
        msWhere = Mid$(valor, 7)
    Else
        msWhere = valor
    End If
    
End Property

''
' Retorna a cláusula WHERE da frase SQL.
'
' @return Cláusula WHERE da frase.
'
Public Property Get Where() As String
    Where = msWhere
End Property

Public Sub JoinClause(ByVal Index As Integer, ByVal TIPO As EJoin, ByVal JoinTable As String, ByVal JoinOn As String)
    mveJoin(Index) = TIPO
    mvsJoin(Index) = JoinTable
    mvsJoinOn(Index) = JoinOn
End Sub

''
' Permite definir o tipo de Join a ser feito para determinada cláusula JOIN.
'
' @param Index Índice que aponta a cláusula Join a ser definida (base zero).
' @param Tipo Tipo de Join a utilizar. Valores válidos estão no enumerador EJoin.
'
' @remarks Podem existir várias cláusulas "Join" em uma única frase SQL, dependendo
'   do número de tabelas a relacionar; daí o uso do índice.
'
Public Property Let JoinType(ByVal Index As Integer, ByVal TIPO As EJoin)
    mveJoin(Index) = TIPO
End Property

''
' Obtém o tipo de Join da cláusula JOIN especificada.
'
' @param Index Índice da cláusula Join a obter.
' @return Tipo de join. Valores válidos estão no enumerador EJoin.
'
Public Property Get JoinType(ByVal Index As Integer) As EJoin
    JoinType = mveJoin(Index)
End Property

''
' Permite definir uma cláusula JOIN na frase SQL.
'
' @param Index Índice da cláusula Join a montar.
' @param valor Nome da tabela que será adicionada ao relacionamento. A palavra-chave "JOIN" não precisa ser especificada.
'
Public Property Let Join(ByVal Index As Integer, ByVal valor As String)

    Dim t As Long
    
    t = InStr(1, "JOIN", valor, vbTextCompare)
    
    If t <> 0 Then
        mvsJoin(Index) = Mid$(valor, t + 1)
    Else
        mvsJoin(Index) = valor
    End If

End Property

''
' Retorna uma cláusula JOIN da frase SQL.
'
' @param Index Índice da cláusula Join a obter.
'
' @return Uma cláusula JOIN da frase.
'
Public Property Get Join(ByVal Index As Integer) As String
    Join(Index) = mvsJoin(Index)
End Property

''
' Permite definir os campos que ligam as tabelas especificadas numa cláusula JOIN na frase SQL.
'
' @param Index Índice da cláusula Join relacionada.
' @param valor Nome dos campos que serão utilizados para o relacionamento. A palavra-chave "ON" não precisa ser especificada.
'
Public Property Let JoinOn(ByVal Index As Integer, ByVal valor As String)

    If UCase$(Left$(valor, 3)) = "ON " Then
        mvsJoinOn(Index) = Mid$(valor, 4)
    Else
        mvsJoinOn(Index) = valor
    End If

End Property

''
' Retorna os campos que ligam as tabelas em uma cláusula JOIN da frase SQL.
'
' @param Index Índice da cláusula JOIN relacionada.
'
' @return Campos de ligação da cláusula JOIN.
'
Public Property Get JoinOn(ByVal Index As Integer) As String
    JoinOn(Index) = mvsJoinOn(Index)
End Property

Public Property Let GroupBy(ByVal valor As String)
    If UCase$(Left$(valor, 9)) = "GROUP BY " Then
        msGroupBy = Mid$(valor, 10)
    Else
        msGroupBy = valor
    End If
End Property
 
Public Property Get GroupBy() As String
    GroupBy = msGroupBy
End Property

Public Property Let Having(ByVal valor As String)
    If UCase$(Left$(valor, 7)) = "HAVING " Then
        msHaving = Mid$(valor, 8)
    Else
        msHaving = valor
    End If
End Property

Public Property Get Having() As String
    Having = msHaving
End Property


''
' Permite definir uma cláusula ORDER BY na frase SQL.
'
' @param valor Nome do campo ou campos que serã usados na ordenação. A palavra-chave "ORDER BY" não precisa ser especificada.
'
Public Property Let OrderBy(ByVal valor As String)

    If UCase$(Left$(valor, 9)) = "ORDER BY " Then
        msOrderBy = Mid$(valor, 10)
    Else
        msOrderBy = valor
    End If

End Property

''
' Retorna a cláusula ORDER BY da frase SQL.
'
' @return Cláusula ORDER BY da frase.
'
Public Property Get OrderBy() As String
    OrderBy = msOrderBy
End Property

''
' Permite definir a ordenação de um campo.
'
' @param valor Tipo de ordenação.
Public Property Let DirecaoOrderBy(ByVal valor As EOrderBy)
    meOrder = valor
End Property

''
' Retorna a ordenação da cláusula ORDER BY da frase SQL, conforme os valores definidos no enum "EOrderBy"
'
' @return Ordenação.
'
Public Property Get DirecaoOrderBy() As EOrderBy
    DirecaoOrderBy = meOrder
End Property

''
' Retorna a frase SQL devidamente montada.
'
Public Property Get FraseSQL() As String

    On Error Resume Next

    Dim sb As StringBuilder: Set sb = New StringBuilder
    Dim lim As Variant
    If Len(msLimit) > 0 Then
        lim = zSplit(msLimit, ",")
    End If

    msbFrase.Reset

    ' A ordem é importante...
    zMontaSelect
    zMontaFrom
    zMontaJoin
    zMontaWhere
    zMontaGroupBy
    zMontaHaving
    zMontaOrderBy

    ' Usa LIMIT?
    If msLimit <> "" Then
        Select Case meBanco
            Case Gerenciador_MySql
                msbFrase.Add (" LIMIT " & msLimit)
            Case Gerenciador_SQLServer
                sb.Add ("SELECT z2.*")
                sb.Add ("FROM (")
                sb.Add ("   SELECT ROW_NUMBER OVER(ORDER BY " & msOrderBy & ") AS rn, z1.*")
                sb.Add ("   FROM (" & msbFrase.ToString() & ") z1")
                sb.Add (") z2")
                sb.Add ("WHERE z2.rn BETWEEN " & lim(0) & " AND " & lim(1))
                Set msbFrase = sb
            Case Gerenciador_Oracle
                sb.Add ("SELECT *")
                sb.Add ("FROM (")
                sb.Add ("   SELECT ROWNUM as ID, a.*")
                sb.Add ("   FROM (" & msbFrase.ToString() & ") a")
                sb.Add ("   WHERE ROWNUM <= " & lim(0))
                sb.Add (") WHERE rnum >= " & lim(1))
                Set msbFrase = sb
        End Select
    End If

    FraseSQL = msbFrase.ToString()

End Property

''
' Reinicia o objeto.
'
Public Sub Reset()

    On Error Resume Next

    mbAddNL = True
    msSelect = ""
    mbSelectDistinct = False
    msFrom = ""
    msWhere = ""
    msOrderBy = ""
    msGroupBy = ""
    msHaving = ""
    Erase mvsJoin
    Erase mveJoin
    Erase mvsJoinOn
    meOrder = Ascendente

End Sub

'
'##############################################################################
' FUNÇÕES INTERNAS DA CLASSE
'##############################################################################
'

''
' Monta cláusula SELECT da frase SQL.
'
Private Sub zMontaSelect()

    On Error Resume Next

    Dim sFrase As String

' comentado -- causou problema com "CASE WHEN" no select
'    If Len(msFrom) <> 0 Then
'        sPalavra = Split(msSelect, ",")
'        sAlias = Split(Trim$(msFrom), " ")
'        For i = 0 To UBound(sPalavra)
'            If InStr(1, sPalavra(i), ".") < 1 Then
'                If UBound(sAlias) = 1 Then
'                    sPalavra(i) = sAlias(1) & "." & Trim$(sPalavra(i))
'                Else
'                    sPalavra(i) = msFrom & "." & Trim$(sPalavra(i))
'                End If
'            End If
'            sFrase = sFrase & Trim$(sPalavra(i)) & ", "
'        Next
'        If Len(sFrase) > 2 Then sFrase = Mid$(sFrase, 1, Len(sFrase) - 2)
'    Else
        sFrase = msSelect
'    End If

    If Len(msSelect) <> 0 Then
        msbFrase.Add "SELECT "
        If mbSelectDistinct Then
            msbFrase.Add "DISTINCT "
        End If
        msbFrase.Add sFrase
        msbFrase.Add " "
        If mbAddNL Then msbFrase.Add vbNewLine
    End If

End Sub

''
' Retorna um separador, conforme o tipo de banco de dados.
'
' @param bSeparadorInicial Ajuste em True para retornar o separador inicial, ou False para
'   retornar o separador final.
'
' @return Separador.
'
' @remarks Diferentes SGBDs possuem diferentes separadores. Os mesmos são geralmente usados
'   com campos contendo espaços no nome, ou com palavras reservadas do BD em questão.
'
Private Function zSeparador(ByVal bSeparadorInicial As Boolean) As String

    On Error Resume Next

    If bSeparadorInicial Then

        Select Case meBanco
            Case Gerenciador_MySql
                zSeparador = "`"
            Case Gerenciador_SQLServer
                zSeparador = "["
            Case Gerenciador_Oracle
                zSeparador = """"
        End Select

    Else

        Select Case meBanco
            Case Gerenciador_MySql
                zSeparador = "`"
            Case Gerenciador_SQLServer
                zSeparador = "]"
            Case Gerenciador_Oracle
                zSeparador = """"
        End Select

    End If

End Function

''
' Monta cláusula FROM da frase SQL.
'
Private Sub zMontaFrom()

    On Error Resume Next

    Dim sFrase As String

    sFrase = msFrom

    If Len(msFrom) <> 0 Then

        msbFrase.Add "FROM "

        'msbFrase.Add Separador(True)
        If Not mbUsarNovaSintaxeJoin Then
            sFrase = sFrase & zAdicionaJoinsAoFrom()
        End If
        msbFrase.Add sFrase
        'msbFrase.Add Separador(False)
        
        msbFrase.Add " "
        If mbAddNL Then msbFrase.Add vbNewLine

    End If

End Sub

''
' Monta cláusula JOIN da frase SQL.
'
Private Sub zMontaJoin()

    On Error Resume Next

    Dim i As Long
    Dim aux As StringBuilder
    Dim sAux As String

    ' Sintaxe antiga não possui cláusulas JOIN...
    If Not mbUsarNovaSintaxeJoin Then
        Exit Sub
    End If

    For i = 0 To 10
        If Len(mvsJoin(i)) = 0 Then Exit For

            ' Valida o join
            If Len(mvsJoinOn(i)) = 0 Then
                Set aux = New StringBuilder
                With aux
                    .Add "Erro: cláusula ""ON"" da sentença ""JOIN"" número "
                    .Add CStr(i)
                    .Add " não especificada!"
                End With
                sAux = aux.ToString
                Set aux = Nothing
                Err.Raise vbObjectError + 1, "SQLFrase", sAux
                Exit Sub
            End If

            ' Monta o join
            Select Case mveJoin(i)
                Case InnerJoin
                    msbFrase.Add "INNER JOIN "
                Case LeftJoin
                    msbFrase.Add "LEFT OUTER JOIN "
                Case RightJoin
                    msbFrase.Add "RIGHT OUTER JOIN "
                Case Else
                    msbFrase.Add "JOIN "
            End Select
            'If InStr(1, mvsJoin(i), " AS ", vbTextCompare) > 0 Then
                msbFrase.Add mvsJoin(i)
'            Else
'                msbFrase.Add Separador(True)
'                msbFrase.Add mvsJoin(i)
'                msbFrase.Add Separador(False)
'            End If
            msbFrase.Add " ON "
            msbFrase.Add mvsJoinOn(i)

            msbFrase.Add " "
            If mbAddNL Then msbFrase.Add vbNewLine

    Next

End Sub

''
' Monta cláusula WHERE da frase SQL.
'
Private Sub zMontaWhere()

    On Error Resume Next

    Dim sFrase As String
    Dim sJoin As String

    ' Se é sintaxe nova, e não tem nada na cláusula, sai
    If mbUsarNovaSintaxeJoin And (Len(msWhere) = 0) Then Exit Sub

    sFrase = msWhere

    ' Se é sintaxe velha, as ligações feitas nas cláusulas JOIN vêm pra cá
    If (Not mbUsarNovaSintaxeJoin) Then
        sJoin = zAdicionaJoinsAoWhere()
        If sFrase <> "" And sJoin <> "" Then sJoin = " AND " & vbNewLine & sJoin
        If sJoin <> "" Then sFrase = sFrase & sJoin
    End If
    If sFrase = "" Then Exit Sub

    msbFrase.Add "WHERE ("
    msbFrase.Add sFrase
    msbFrase.Add ") "
    If mbAddNL Then msbFrase.Add vbNewLine

End Sub

Private Sub zMontaGroupBy()
    On Error Resume Next
    If Len(msGroupBy) <> 0 Then
        msbFrase.Add "GROUP BY " & msGroupBy
        If mbAddNL Then msbFrase.Add vbNewLine
    End If
End Sub

Private Sub zMontaHaving()
    On Error Resume Next
    If Len(msHaving) <> 0 Then
        msbFrase.Add "HAVING " & msHaving
        If mbAddNL Then msbFrase.Add vbNewLine
    End If
End Sub

''
' Monta cláusula ORDER BY da frase SQL.
'
Private Sub zMontaOrderBy()
    On Error Resume Next
    If Len(msOrderBy) <> 0 Then
        msbFrase.Add "ORDER BY " & msOrderBy
        If meOrder = Ascendente Then
            msbFrase.Add " ASC "
        Else
            msbFrase.Add " DESC "
        End If
        If mbAddNL Then msbFrase.Add vbNewLine
    End If
End Sub

''
' Inicialização da classe.
'
Private Sub Class_Initialize()
    On Error Resume Next
    Set msbFrase = New StringBuilder
    mbAddNL = True
    mbUsarNovaSintaxeJoin = True
End Sub

''
' Finalização da classe.
'
Private Sub Class_Terminate()
    On Error Resume Next
    Set msbFrase = Nothing
End Sub

Private Function zAdicionaJoinsAoFrom() As String

    Dim sAux As String
    Dim i As Long

    For i = 0 To 10
        If mvsJoin(i) <> "" Then
            sAux = sAux & ", " & mvsJoin(i)
        End If
    Next

    zAdicionaJoinsAoFrom = sAux

End Function

Private Function zAdicionaJoinsAoWhere() As String

    Dim sAux As String
    Dim i As Long
    Dim condicoes As Variant
    Dim partes As Variant
    Dim sOperador As String
    Dim sSinal As String
    Dim j As Long

    For i = 0 To 10
        If mvsJoinOn(i) <> "" Then

            If Len(sAux) > 0 Then sAux = sAux & " AND "
            If Len(sAux) > 0 And mbAddNL Then sAux = sAux & vbNewLine

            ' Obtém as condições
            Call zObtemCondicoes(mvsJoinOn(i), condicoes, sOperador)

            ' Faz para cada condição...
            For j = 0 To UBound(condicoes)

                condicoes(j) = Trim$(condicoes(j))
                If j > 0 And Len(condicoes(j)) > 0 Then sAux = sAux & sOperador

                If Len(condicoes(j)) > 0 Then

                    Call zObtemClausulas(condicoes(j), partes, sSinal)

                    Select Case mveJoin(i)
                        Case InnerJoin
                            sAux = sAux & Trim$(partes(0)) & sSinal & Trim$(partes(1))
                        Case LeftJoin
                            If meBanco = Gerenciador_Oracle Then
                                sAux = sAux & Trim$(partes(0)) & _
                                       sSinal & Trim$(partes(1)) & "(+)"
                            ElseIf meBanco = Gerenciador_SQLServer Then
                                sAux = sAux & Trim$(partes(0)) & " =* " & Trim$(partes(1))
                            End If
                        Case RightJoin
                            If meBanco = Gerenciador_Oracle Then
                                sAux = sAux & Trim$(partes(0)) & "(+)" & _
                                       sSinal & Trim$(partes(1))
                            ElseIf meBanco = Gerenciador_SQLServer Then
                                sAux = sAux & Trim$(partes(0)) & " *= " & Trim$(partes(1))
                            End If
                    End Select

                End If

            Next

        End If
    Next

    zAdicionaJoinsAoWhere = sAux

End Function

Private Function zSplit(ByVal Text As String, _
                        Optional ByVal Delimiter As String = " ", _
                        Optional ByVal Limit As Long = -1, _
                        Optional CompareMethod As VbCompareMethod = vbBinaryCompare) As Variant

    ReDim res(0 To 100) As String
    Dim resCount As Long
    Dim length As Long
    Dim startIndex As Long
    Dim endIndex As Long
    
    length = Len(Text)
    startIndex = 1
    
    Do While startIndex <= length And resCount <> Limit
        ' get the next delimiter
        endIndex = InStr(startIndex, Text, Delimiter, CompareMethod)
        If endIndex = 0 Then endIndex = length + 1
        
        ' make room in the array, if necessary
        If resCount > UBound(res) Then
            ReDim Preserve res(0 To resCount + 99) As String
        End If
        ' store the new element
        res(resCount) = Mid$(Text, startIndex, endIndex - startIndex)
        resCount = resCount + 1
        
        startIndex = endIndex + Len(Delimiter)
    Loop
    
    ' trim unused values
    ReDim Preserve res(0 To resCount - 1) As String

    ' return the array inside a Variant
    zSplit = res()

End Function

''
' Obtém as condições da cláusula "join on".
'
' Ex: na frase
'   LEFT OUTER JOIN dbo.STATE s ON s.COUNTRYCODE = a.COUNTRYCODE AND s.STATE = a.STATE
' a primeira condição seria "s.COUNTRYCODE = a.COUNTRYCODE",
' a segunda seria "s.STATE = a.STATE", e a condição seria "AND"
'
Private Sub zObtemCondicoes(ByVal clausula As String, _
                            ByRef condicoes As Variant, _
                            ByRef sOperador As String)

    If InStr(1, clausula, " AND ", vbTextCompare) > 0 Then
        condicoes = zSplit(clausula, " AND ")
        sOperador = " AND "
    ElseIf InStr(1, clausula, " OR ", vbTextCompare) > 0 Then
        condicoes = zSplit(clausula, " OR ")
        sOperador = " OR "
    Else
        ReDim condicoes(0 To 1)
        condicoes(0) = clausula
        condicoes(1) = ""
        sOperador = ""
    End If

End Sub

''
' Obtém as partes de determinada condição da cláusula "join on".
'
' Ex: na frase
'   LEFT OUTER JOIN dbo.STATE s ON s.COUNTRYCODE = a.COUNTRYCODE AND s.STATE = a.STATE
' a primeira condição seria "s.COUNTRYCODE = a.COUNTRYCODE",
' as partes seriam "s.COUNTRYCODE" e "a.COUNTRYCODE" e o sinal "=".
'
Private Sub zObtemClausulas(ByVal sFrase As String, _
                            ByRef partes As Variant, _
                            ByRef sSinal As String)

    ' Obtém as partes da cláusula "join on"
    If InStr(1, sFrase, "=") > 0 Then
        partes = zSplit(sFrase, "=")
        sSinal = " = "
    ElseIf InStr(1, sFrase, "<>") > 0 Then
        partes = zSplit(sFrase, "<>")
        sSinal = " <> "
    ElseIf InStr(1, sFrase, "<=") > 0 Then
        partes = zSplit(sFrase, "<=")
        sSinal = " <= "
    ElseIf InStr(1, sFrase, ">=") > 0 Then
        partes = zSplit(sFrase, ">=")
        sSinal = " >= "
    ElseIf InStr(1, sFrase, "<") > 0 Then
        partes = zSplit(sFrase, "<")
        sSinal = " < "
    ElseIf InStr(1, sFrase, ">") > 0 Then
        partes = zSplit(sFrase, ">")
        sSinal = " > "
    Else
        ReDim partes(0 To 1)
        partes(0) = ""
        partes(1) = ""
        sSinal = ""
    End If

End Sub
