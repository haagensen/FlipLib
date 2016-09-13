VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Conversoes"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Declare Sub CopyMemoryAPI Lib "kernel32" Alias "RtlMoveMemory" (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)

' WinApi function that maps a UTF-16 (wide character) string to a new character string
Private Declare Function WideCharToMultiByteAPI Lib "kernel32" Alias "WideCharToMultiByte" ( _
    ByVal CodePage As Long, _
    ByVal dwFlags As Long, _
    ByVal lpWideCharStr As Long, _
    ByVal cchWideChar As Long, _
    ByVal lpMultiByteStr As Long, _
    ByVal cbMultiByte As Long, _
    ByVal lpDefaultChar As Long, _
    ByVal lpUsedDefaultChar As Long) As Long

Public Function StringToUTF8(ByVal strInput As String) As Variant ' As Byte()
    ' fonte: <http://www.di-mgt.com.au/howto-convert-vba-unicode-to-utf8.html>
    Const CP_UTF8 = 65001
    Dim nBytes As Long
    Dim abBuffer() As Byte
    ' Get length in bytes *including* terminating null
    nBytes = WideCharToMultiByteAPI(CP_UTF8, 0&, ByVal StrPtr(strInput), -1, vbNull, 0&, 0&, 0&)
    ' We don't want the terminating null in our byte array, so ask for `nBytes-1` bytes
    ReDim abBuffer(nBytes - 2)  ' NB ReDim with one less byte than you need
    nBytes = WideCharToMultiByteAPI(CP_UTF8, 0&, ByVal StrPtr(strInput), -1, ByVal VarPtr(abBuffer(0)), nBytes - 1, 0&, 0&)
    StringToUTF8 = abBuffer
'
'    Exemplo de uso:
'
'    Dim f As Integer, vetorBytes() As Byte
'    f = FreeFile
'    Open sLogFile For Binary Access Write Lock Read Write As #f
'    vetorBytes = StringToUTF8(sFrase)
'    Put #f, , vetorBytes
'    Close #f
'
End Function

Public Function ByteArrayToString(vetorBytes As Variant) As String

    Dim iUnicode As Long, i As Long, j As Long

    On Error Resume Next

    i = UBound(vetorBytes)

    If (i < 1) Then ' ANSI?
        ByteArrayToString = StrConv(vetorBytes, vbUnicode)
        Exit Function
    End If
    i = i + 1

    ' Verifica primeiros dois bytes
    CopyMemoryAPI iUnicode, vetorBytes(0), 2

    If iUnicode = vetorBytes(0) Then ' Unicode
        If (i Mod 2) Then i = i - 1 ' tira NULL final
        ByteArrayToString = String$(i / 2, 0) ' Buffer
        CopyMemoryAPI ByVal StrPtr(ByteArrayToString), vetorBytes(0), i
    Else 'ANSI
        ByteArrayToString = StrConv(vetorBytes, vbUnicode)
    End If

End Function

Public Function StringToByteArray(ByVal strInput As String, _
                                  Optional ByVal bReturnAsUnicode As Boolean = True, _
                                  Optional ByVal bAddNullTerminator As Boolean = False) As Variant ' Byte()

    Const CP_ACP = 0 'ANSI
    Dim lRet As Long
    Dim bytBuffer() As Byte
    Dim lLenB As Long

    If bReturnAsUnicode Then
        lLenB = LenB(strInput)
        If bAddNullTerminator Then
            ReDim bytBuffer(lLenB)
        Else
            ReDim bytBuffer(lLenB - 1)
        End If
        CopyMemoryAPI bytBuffer(0), ByVal StrPtr(strInput), lLenB
    Else
        lLenB = Len(strInput)
        If bAddNullTerminator Then
            ReDim bytBuffer(lLenB)
        Else
            ReDim bytBuffer(lLenB - 1)
        End If
        lRet = WideCharToMultiByteAPI(CP_ACP, 0&, ByVal StrPtr(strInput), -1, ByVal VarPtr(bytBuffer(0)), lLenB, 0&, 0&)
    End If

    StringToByteArray = bytBuffer

End Function

Public Function ToDate(ByVal vData As Variant) As Date

    On Error Resume Next
    If IsDate(vData) Then ToDate = CDate(vData) Else ToDate = CDate(0)

End Function

Public Function ToCur(ByVal pvValor As Variant) As Currency

    On Error Resume Next

    Dim i As Long, sLetra As String
    Dim cVB As VB6: Set cVB = New VB6

    If IsNull(pvValor) Then
        ToCur = 0
    Else

        ' Retira separador de milhares
        If (InStr(1, pvValor, "US$", vbTextCompare) > 0) Or InStr(1, pvValor, "£") > 0 Then
            pvValor = cVB.Replace(pvValor, ",", "")
            pvValor = cVB.Replace(pvValor, ".", ",")
        Else
            pvValor = cVB.Replace(pvValor, ".", "")
        End If
        pvValor = cVB.Replace(pvValor, " ", "") ' em alguns países o separador de milhar é o espaço

        ' transforma valores entre parênteses para negativo
        If Left$(pvValor, 1) = "(" And Right$(pvValor, 1) = ")" Then
            pvValor = "-" & Mid$(pvValor, 2, Len(pvValor) - 2)
        End If

        ' retira demais sinais
        For i = 1 To Len(pvValor)
            sLetra = Mid$(pvValor, i, 1)
            If (Asc(sLetra) > 47 And Asc(sLetra) < 58) Or sLetra = "," Or sLetra = "." Then '0123456789,.
            Else
                Mid$(pvValor, i, 1) = " "
            End If
        Next
        pvValor = cVB.Replace(pvValor, " ", "")

        ToCur = cVB.Round(CCur(pvValor), 2)

    End If

End Function
