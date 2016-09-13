VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "VB6"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Function Join(arr() As String, ByVal Delimiter As String) As String
    Dim i As Long
    For i = LBound(arr) To UBound(arr) - 1
        Join = Join & arr(i) & Delimiter
    Next
    Join = Join & arr(UBound(arr))
End Function

Public Function Replace(ByVal sIn As String, ByVal sFind As String, ByVal sReplace As String, _
                        Optional nStart As Long = 1, Optional nCount As Long = -1, _
                        Optional bCompare As VbCompareMethod = vbBinaryCompare) As String
    
    Dim nC As Long, nPos As Long
    Dim nFindLen As Long, nReplaceLen As Long
    
    nFindLen = Len(sFind)
    nReplaceLen = Len(sReplace)
    
    If (sFind <> "") And (sFind <> sReplace) Then
        nPos = InStr(nStart, sIn, sFind, bCompare)
        Do While nPos
            nC = nC + 1
            sIn = Left(sIn, nPos - 1) & sReplace & Mid(sIn, nPos + nFindLen)
            If nCount <> -1 And nC >= nCount Then Exit Do
            nPos = InStr(nPos + nReplaceLen, sIn, sFind, bCompare)
        Loop
    End If

    Replace = sIn

End Function

' Tirei de http://www.xbeat.net/vbspeed/c_Round.htm
Public Function Round(ByVal v As Double, Optional ByVal lngDecimals As Long = 0) As Double

  Dim xint As Double, yint As Double, xrest As Double
  Static PreviousValue    As Double
  Static PreviousDecimals As Long
  Static PreviousOutput   As Double
  Static M                As Double
      
  If PreviousValue = v And PreviousDecimals = lngDecimals Then Round = PreviousOutput: Exit Function
  If v = 0 Then Exit Function

  If PreviousDecimals = lngDecimals Then
      If M = 0 Then M = 1
      Else
        PreviousDecimals = lngDecimals
        M = 10 ^ lngDecimals
      End If
  
  If M = 1 Then xint = v Else xint = v * CDec(M)
  Round = Fix(xint)

  If Abs(Fix(10 * (xint - Round))) > 4 Then
    If xint < 0 Then
      Round = Round - 1
    Else
      Round = Round + 1
    End If
  End If

  If M = 1 Then Else Round = Round / M

  PreviousOutput = Round
  PreviousValue = v

End Function

Public Function Split(ByVal Text As String, _
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
    Split = res()

End Function

Public Function InStrRev(ByVal StringCheck As String, ByVal StringMatch As String, _
                         Optional ByVal Start As Long = -1, _
                         Optional ByVal Compare As VbCompareMethod = vbBinaryCompare) As Long

    '===============================================================================================
    ' A clone of VB6's InstrRev function (including its quirks) that works under VB4 and VB5
    '===============================================================================================
    
    Dim Index As Long
    Dim lastIndex As Long
    
    If Start > Len(StringCheck) Then
        ' this is probably a quirk in VB6's InstrRev function: when
        ' start is higher than the source string length, the function returns zero
        Exit Function
    ElseIf Start < 0 Then
        ' if Start is omitted, last valid index is the end of string
        lastIndex = Len(StringCheck)
    Else
        ' else, we must account for the length of the searched string
        ' (this is the way VB6's InstrRev function works)
        lastIndex = Start + 1 - Len(StringMatch)
    End If
    
    Do
        Index = InStr(Index + 1, StringCheck, StringMatch, Compare)
        ' exit if not found, or if it's past the start index
        If Index = 0 Or Index > lastIndex Then Exit Do
        ' remember the index we've just found
        InStrRev = Index
    Loop
    
End Function
