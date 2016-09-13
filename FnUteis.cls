VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FnUteis"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Declare Function FoldStringAPI Lib "kernel32.dll" Alias "FoldStringA" (ByVal dwMapFlags As Long, ByVal lpSrcStr As Long, ByVal cchSrc As Long, ByVal lpDestStr As Long, ByVal cchDest As Long) As Long
Private Declare Function FormatMessageAPI Lib "kernel32" Alias "FormatMessageA" (ByVal dwFlags As Long, lpSource As Any, ByVal dwMessageId As Long, ByVal dwLanguageId As Long, ByVal lpBuffer As String, ByVal nSize As Long, Arguments As Long) As Long

'
'##########################################################################################
' Métodos
'##########################################################################################
'

Public Function GetMySQLODBCDriverName() As String

    Dim obj As cRegistry
    Dim ub As Long, i As Long
    Dim sKeys() As String, lKeyCount As Long

    On Error GoTo Sai

    Set obj = New cRegistry
    With obj
        .ClassKey = HKEY_LOCAL_MACHINE
        .SectionKey = "SOFTWARE\ODBC\ODBCINST.INI\ODBC Drivers"
        .EnumerateValues sKeys(), lKeyCount 'base 1
    End With

    ub = UBound(sKeys)
    If ub < 1 Then GoTo Sai
    
    For i = 1 To ub
        obj.ValueKey = sKeys(i)
        obj.ValueType = REG_SZ
        If obj.Value = "Installed" Then
            If InStr(1, sKeys(i), "MySQL", vbTextCompare) > 0 Then
                GetMySQLODBCDriverName = sKeys(i)
                GoTo Sai
            End If
        End If
    Next

Sai:
    On Error Resume Next
    Set obj = Nothing

End Function

''
' Retorna a string T sem nenhum caracter NULL que porventura exista em seu final.
'
Public Function StripNulls(ByVal t As String) As String

    If (InStr(t, Chr$(0)) > 0) Then
        t = Left$(t, InStr(t, Chr$(0)) - 1)
    End If
    StripNulls = t

End Function

''
' **** Se a classe for compilada, a função abaixo NÃO irá funcionar! *****
'
'Public Function IsRunningIDE() As Boolean
'    On Error Resume Next
'    Debug.Assert 1 / 0
'    IsRunningIDE = (Err <> 0)
'    On Error GoTo 0
'End Function

Public Function IsTrue(ByVal vValor As Variant) As Boolean
    On Error Resume Next
    If Trim$(vValor & "") = "1" Then
        IsTrue = True
    ElseIf Left$(Trim$(UCase$(vValor & "")), 1) = "T" Then
        IsTrue = True
    ElseIf Left$(Trim$(UCase$(vValor & "")), 1) = "V" Then
        IsTrue = True
    ElseIf Left$(Trim$(UCase$(vValor & "")), 1) = "Y" Then
        IsTrue = True
    ElseIf Left$(Trim$(UCase$(vValor & "")), 1) = "S" Then
        IsTrue = True
    Else
        Dim bTeste As Boolean
        Err.Clear
        bTeste = CBool(Trim$(UCase$(vValor & "")))
        If Err.Number = 0 Then
            IsTrue = bTeste
        End If
    End If
End Function

Public Function Descriptografa(ByVal s As String, ByVal Chave As Integer) As String

    Dim i As Integer, out As String
    Dim Ant As Integer, aux As String
    
    On Error Resume Next
    
    out = ""
    For i = 2 To Len(Trim$(s))
        If Int((i - 1) / 2) = (i - 1) / 2 Then
            out = out + Chr(Asc(Mid$(s, i, 1)) - Chave * (i - 1))
        Else
            out = out + Chr(Asc(Mid$(s, i, 1)) - Chave * (i - 1) * 2)
        End If
    Next

    Ant = (Asc(Left$(s, 1)) - 40)
    aux = ""
    For i = 1 To Len(Trim$(out))
        aux = aux + Chr(Asc(Mid$(out, i, 1)) - Ant)
    Next

    Descriptografa = aux

End Function

Public Function BackSlash(ByVal sStr As String) As String
    On Error Resume Next
    If Right$(sStr, 1) <> "\" Then
        sStr = sStr & "\"
    End If
    BackSlash = sStr
End Function

Public Function UnBackSlash(ByVal sStr As String) As String
    On Error Resume Next
    If Len(sStr) > 0 Then
        If Right$(sStr, 1) = "\" Then
            UnBackSlash = Left$(sStr, Len(sStr) - 1)
            Exit Function
        End If
    End If
    UnBackSlash = sStr
End Function

Public Function ObtemTextoSemAcentos(ByVal psTexto As String) As String

    On Error Resume Next

    Const MAP_COMPOSITE = &H40
    Dim sBuffer As String
    Dim i As Long

    sBuffer = String$(Len(psTexto), vbNullChar)
    For i = 0 To (Len(psTexto) - 1) * 2 Step 2
        FoldStringAPI MAP_COMPOSITE, StrPtr(psTexto) + i, 1, StrPtr(sBuffer) + i, 1
    Next
    ObtemTextoSemAcentos = sBuffer

End Function

''
' Verifica se determinado endereço de email é um email válido.
'
' @param psAddressString Endereço de email a ser verificado.
'
' @return True se o email é válido, False se não.
'
Public Function IsValidEmailAddress(ByVal psAddressString As String) As Boolean

    Dim sHost As String
    Dim iPos As Integer
    Dim sInvalidChars As String

    On Error Resume Next

    If Len(Trim$(psAddressString)) = 0 Then
        Exit Function
    End If

    sInvalidChars = "!#$%^&*()=+{}[]|\;:'/?>,< áéíóúçâêîôûüãõàèìòù"
    For iPos = 1 To Len(psAddressString)
        If InStr(sInvalidChars, Mid$(LCase$(psAddressString), iPos, 1)) > 0 Then
            Exit Function
        End If
    Next
  
    If InStr(psAddressString, ".@") > 0 Then
        Exit Function
    End If
  
    If InStr(psAddressString, "@.") > 0 Then
        Exit Function
    End If

    iPos = InStr(psAddressString, "@")
    If iPos = 0 Or Left$(psAddressString, 1) = "@" Then
        Exit Function
    End If

    sHost = Mid$(psAddressString, iPos + 1)
    If InStr(sHost, "@") > 0 Then
        Exit Function
    End If

    IsValidEmailAddress = zIsValidIPHostPF(UCase$(sHost))

End Function

''
' Verifica se o IP passado é um número de IP v.4 válido.
'
' @param psHostString String contendo o endereço IP a verificar.
' @return True se o IP é válido, False se não.
'
Private Function zIsValidIPHostPF(ByVal psHostString As String) As Boolean

    Dim sHost As String
    Dim bDottedQuad As Boolean
    Dim sSplit '() As String
    Dim iCtr As Integer
    Dim bAns As Boolean

    On Error Resume Next

    sHost = psHostString

    If InStr(sHost, ".") = 0 Then
        Exit Function
    End If

    Dim vbSix As VB6: Set vbSix = New VB6
    sSplit = vbSix.Split(sHost, ".")

    If UBound(sSplit) = 3 Then

        bDottedQuad = True
        For iCtr = 0 To 3
            If Not IsNumeric(sSplit(iCtr)) Then
                bDottedQuad = False
                Exit For
            End If
        Next

        If bDottedQuad Then
            bAns = True
            For iCtr = 0 To 3
                If iCtr = 0 Then
                    bAns = Val(sSplit(iCtr)) <= 239
                    If bAns = False Then Exit For
                Else
                    bAns = Val(sSplit(iCtr)) <= 255
                    If bAns = False Then Exit For
                End If
            Next
            zIsValidIPHostPF = bAns
            Exit Function
        End If

    End If

    zIsValidIPHostPF = zIsTopLevelDomainPF(sSplit(UBound(sSplit)))

End Function

''
' Verifica se um endereço internet pertence a um "top level domain".
'
' @param psDomainString Endereço a verificar.
' @return True se é um domínio superior, False se não.
'
Private Function zIsTopLevelDomainPF(ByVal psDomainString As String) As Boolean

    Dim asTopLevels() As String
    Dim iCtr As Integer
    Dim iNumDomains As Integer
    Dim bAns As Boolean

    On Error Resume Next

    iNumDomains = 251
    ReDim asTopLevels(iNumDomains - 1) As String

    asTopLevels(0) = "COM":  asTopLevels(1) = "ORG":  asTopLevels(2) = "NET"
    asTopLevels(3) = "EDU":  asTopLevels(4) = "GOV":  asTopLevels(5) = "MIL"
    asTopLevels(6) = "INT":  asTopLevels(7) = "AF":   asTopLevels(8) = "AL"
    asTopLevels(9) = "DZ":   asTopLevels(10) = "AS":  asTopLevels(11) = "AD"
    asTopLevels(12) = "AO":  asTopLevels(13) = "AI":  asTopLevels(14) = "AQ"
    asTopLevels(15) = "AG":  asTopLevels(16) = "AR":  asTopLevels(17) = "AM"
    asTopLevels(18) = "AW":  asTopLevels(19) = "AC":  asTopLevels(20) = "AU"
    asTopLevels(21) = "AT":  asTopLevels(22) = "AZ":  asTopLevels(23) = "BS"
    asTopLevels(24) = "BH":  asTopLevels(25) = "BD":  asTopLevels(26) = "BB"
    asTopLevels(27) = "BY":  asTopLevels(28) = "BZ":  asTopLevels(29) = "BT"
    asTopLevels(30) = "BJ":  asTopLevels(31) = "BE":  asTopLevels(32) = "BM"
    asTopLevels(33) = "BO":  asTopLevels(34) = "BA":  asTopLevels(35) = "BW"
    asTopLevels(36) = "BV":  asTopLevels(37) = "BR":  asTopLevels(38) = "IO"
    asTopLevels(39) = "BN":  asTopLevels(40) = "BG":  asTopLevels(41) = "BF"
    asTopLevels(42) = "BI":  asTopLevels(43) = "KH":  asTopLevels(44) = "CM"
    asTopLevels(45) = "CA":  asTopLevels(46) = "CV":  asTopLevels(47) = "KY"
    asTopLevels(48) = "CF":  asTopLevels(49) = "TD":  asTopLevels(50) = "CL"
    asTopLevels(51) = "CN":  asTopLevels(52) = "CX":  asTopLevels(53) = "CC"
    asTopLevels(54) = "CO":  asTopLevels(55) = "KM":  asTopLevels(56) = "CD"
    asTopLevels(57) = "CG":  asTopLevels(58) = "CK":  asTopLevels(59) = "CR"
    asTopLevels(60) = "CI":  asTopLevels(61) = "HR":  asTopLevels(62) = "CU"
    asTopLevels(63) = "CY":  asTopLevels(64) = "CZ":  asTopLevels(65) = "DK"
    asTopLevels(66) = "DJ":  asTopLevels(67) = "DM":  asTopLevels(68) = "DO"
    asTopLevels(69) = "TP":  asTopLevels(70) = "EC":  asTopLevels(71) = "EG"
    asTopLevels(72) = "SV":  asTopLevels(73) = "GQ":  asTopLevels(74) = "ER"
    asTopLevels(75) = "EE":  asTopLevels(76) = "ET":  asTopLevels(77) = "FK"
    asTopLevels(78) = "FO":  asTopLevels(79) = "FJ":  asTopLevels(80) = "FI"
    asTopLevels(81) = "FR":  asTopLevels(82) = "GF":  asTopLevels(83) = "PF"
    asTopLevels(84) = "TF":  asTopLevels(85) = "GA":  asTopLevels(86) = "GM"
    asTopLevels(87) = "GE":  asTopLevels(88) = "DE":  asTopLevels(89) = "GH"
    asTopLevels(90) = "GI":  asTopLevels(91) = "GR":  asTopLevels(92) = "GL"
    asTopLevels(93) = "GD":  asTopLevels(94) = "GP":  asTopLevels(95) = "GU"
    asTopLevels(96) = "GT":  asTopLevels(97) = "GG":  asTopLevels(98) = "GN"
    asTopLevels(99) = "GW":  asTopLevels(100) = "GY": asTopLevels(101) = "HT"
    asTopLevels(102) = "HM": asTopLevels(103) = "VA": asTopLevels(104) = "HN"
    asTopLevels(105) = "HK": asTopLevels(106) = "HU": asTopLevels(107) = "IS"
    asTopLevels(108) = "IN": asTopLevels(109) = "ID": asTopLevels(110) = "IR"
    asTopLevels(111) = "IQ": asTopLevels(112) = "IE": asTopLevels(113) = "IM"
    asTopLevels(114) = "IL": asTopLevels(115) = "IT": asTopLevels(116) = "JM"
    asTopLevels(117) = "JP": asTopLevels(118) = "JE": asTopLevels(119) = "JO"
    asTopLevels(120) = "KZ": asTopLevels(121) = "KE": asTopLevels(122) = "KI"
    asTopLevels(123) = "KP": asTopLevels(124) = "KR": asTopLevels(125) = "KW"
    asTopLevels(126) = "KG": asTopLevels(127) = "LA": asTopLevels(128) = "LV"
    asTopLevels(129) = "LB": asTopLevels(130) = "LS": asTopLevels(131) = "LR"
    asTopLevels(132) = "LY": asTopLevels(133) = "LI": asTopLevels(134) = "LT"
    asTopLevels(135) = "LU": asTopLevels(136) = "MO": asTopLevels(137) = "MK"
    asTopLevels(138) = "MG": asTopLevels(139) = "MW": asTopLevels(140) = "MY"
    asTopLevels(141) = "MV": asTopLevels(142) = "ML": asTopLevels(143) = "MT"
    asTopLevels(144) = "MH": asTopLevels(145) = "MQ": asTopLevels(146) = "MR"
    asTopLevels(147) = "MU": asTopLevels(148) = "YT": asTopLevels(149) = "MX"
    asTopLevels(150) = "FM": asTopLevels(151) = "MD": asTopLevels(152) = "MC"
    asTopLevels(153) = "MN": asTopLevels(154) = "MS": asTopLevels(155) = "MA"
    asTopLevels(156) = "MZ": asTopLevels(157) = "MM": asTopLevels(158) = "NA"
    asTopLevels(159) = "NR": asTopLevels(160) = "NP": asTopLevels(161) = "NL"
    asTopLevels(162) = "AN": asTopLevels(163) = "NC": asTopLevels(164) = "NZ"
    asTopLevels(165) = "NI": asTopLevels(166) = "NE": asTopLevels(167) = "NG"
    asTopLevels(168) = "NU": asTopLevels(169) = "NF": asTopLevels(170) = "MP"
    asTopLevels(171) = "NO": asTopLevels(172) = "OM": asTopLevels(173) = "PK"
    asTopLevels(174) = "PW": asTopLevels(175) = "PA": asTopLevels(176) = "PG"
    asTopLevels(177) = "PY": asTopLevels(178) = "PE": asTopLevels(179) = "PH"
    asTopLevels(180) = "PN": asTopLevels(181) = "PL": asTopLevels(182) = "PT"
    asTopLevels(183) = "PR": asTopLevels(184) = "QA": asTopLevels(185) = "RE"
    asTopLevels(186) = "RO": asTopLevels(187) = "RU": asTopLevels(188) = "RW"
    asTopLevels(189) = "KN": asTopLevels(190) = "LC": asTopLevels(191) = "VC"
    asTopLevels(192) = "WS": asTopLevels(193) = "SM": asTopLevels(194) = "ST"
    asTopLevels(195) = "SA": asTopLevels(196) = "SN": asTopLevels(197) = "SC"
    asTopLevels(198) = "SL": asTopLevels(199) = "SG": asTopLevels(200) = "SK"
    asTopLevels(201) = "SI": asTopLevels(202) = "SB": asTopLevels(203) = "SO"
    asTopLevels(204) = "ZA": asTopLevels(205) = "GS": asTopLevels(206) = "ES"
    asTopLevels(207) = "LK": asTopLevels(208) = "SH": asTopLevels(209) = "PM"
    asTopLevels(210) = "SD": asTopLevels(211) = "SR": asTopLevels(212) = "SJ"
    asTopLevels(213) = "SZ": asTopLevels(214) = "SE": asTopLevels(215) = "CH"
    asTopLevels(216) = "SY": asTopLevels(217) = "TW": asTopLevels(218) = "TJ"
    asTopLevels(219) = "TZ": asTopLevels(220) = "TH": asTopLevels(221) = "TG"
    asTopLevels(222) = "TK": asTopLevels(223) = "TO": asTopLevels(224) = "TT"
    asTopLevels(225) = "TN": asTopLevels(226) = "TR": asTopLevels(227) = "TM"
    asTopLevels(228) = "TC": asTopLevels(229) = "TV": asTopLevels(230) = "UG"
    asTopLevels(231) = "UA": asTopLevels(232) = "AE": asTopLevels(233) = "GB"
    asTopLevels(234) = "US": asTopLevels(235) = "UM": asTopLevels(236) = "UY"
    asTopLevels(237) = "UZ": asTopLevels(238) = "VU": asTopLevels(239) = "VE"
    asTopLevels(240) = "VN": asTopLevels(241) = "VG": asTopLevels(242) = "VI"
    asTopLevels(243) = "WF": asTopLevels(244) = "EH": asTopLevels(245) = "YE"
    asTopLevels(246) = "YU": asTopLevels(247) = "ZR": asTopLevels(248) = "ZM"
    asTopLevels(249) = "ZW": asTopLevels(250) = "UK"

    For iCtr = 0 To iNumDomains - 1
        If asTopLevels(iCtr) = psDomainString Then
            bAns = True
            Exit For
        End If
    Next

    zIsTopLevelDomainPF = bAns

End Function

Public Function ObtemMensagemLastDllError() As String

    Dim lLDE As Long
    lLDE = Err.LastDllError

    Const FORMAT_MESSAGE_FROM_SYSTEM = &H1000
    Dim strMessage As String
    Dim lRet As Long
    strMessage = Space$(4096)
    lRet = FormatMessageAPI(FORMAT_MESSAGE_FROM_SYSTEM, ByVal 0, lLDE, 0, strMessage, Len(strMessage), 0)
    ObtemMensagemLastDllError = Left$(strMessage, lRet)

End Function

Public Function ThisSunday() As Date
    On Error Resume Next
    ThisSunday = DateAdd("d", -WeekDay(Date) + vbSunday, Date)
End Function

