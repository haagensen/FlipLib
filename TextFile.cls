VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TextFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
''
' Origem: <http://stackoverflow.com/questions/15809081/encoding-of-text-files-in-vb-6-0>
'

Option Explicit

Private Declare Function MultiByteToWideChar Lib "kernel32.dll" ( _
    ByVal CodePage As Long, _
    ByVal dwFlags As Long, _
    ByVal lpMultiByteStr As Long, _
    ByVal cbMultiByte As Long, _
    ByVal lpWideCharStr As Long, _
    ByVal cchWideChar As Long _
) As Long

Private Declare Function WideCharToMultiByte Lib "kernel32.dll" ( _
    ByVal CodePage As Long, _
    ByVal dwFlags As Long, _
    ByVal lpWideCharStr As Long, _
    ByVal cchWideChar As Long, _
    ByVal lpMultiByteStr As Long, _
    ByVal cbMultiByte As Long, _
    ByVal lpDefaultChar As Long, _
    ByVal lpUsedDefaultChar As Long _
) As Long

Public Enum ECodePages
    CP_ANSI = 0
    CP_UTF8 = 65001
    CP_UTF16_LE = 1200
    CP_UTF16_BE = 1201
    CP_UTF32_LE = 12000
    CP_UTF32_BE = 12001
End Enum

'Private Const CP_ACP        As Long = 0          ' Default ANSI code page.
'Private Const CP_UTF8       As Long = 65001      ' UTF8.
'Private Const CP_UTF16_LE   As Long = 1200       ' UTF16 - little endian.
'Private Const CP_UTF16_BE   As Long = 1201       ' UTF16 - big endian.
'Private Const CP_UTF32_LE   As Long = 12000      ' UTF32 - little endian.
'Private Const CP_UTF32_BE   As Long = 12001      ' UTF32 - big endian.

' Purpose:  Heuristic to determine whether bytes in a file are UTF-8.
Private Function pFileBytesAreUTF8(ByVal iFileNo As Integer) As Boolean

    Const knSampleByteSize   As Long = 2048
    Dim nLof                 As Long
    Dim nByteCount           As Long
    Dim nByteIndex           As Long
    Dim nCharExtraByteCount  As Long
    Dim bytValue             As Byte

    ' We look at the first <knSampleByteSize> bytes of the file. However, if the file is smaller, we will have to
    ' use the smaller size.
    nLof = LOF(iFileNo)
    If nLof < knSampleByteSize Then
        nByteCount = nLof
    Else
        nByteCount = knSampleByteSize
    End If

    ' Go to the start of the file.
    Seek #iFileNo, 1

    For nByteIndex = 1 To nByteCount

        Get #iFileNo, , bytValue

        ' If the character we are processing has bytes beyond 1, then we are onto the next character.
        If nCharExtraByteCount = 0 Then
            '
            ' The UTF-8 specification says that the first byte of a character has masking bits which indicate how many bytes follow.
            '
            ' See: http://en.wikipedia.org/wiki/UTF-8#Description
            '
            ' Bytes in
            ' sequence   Byte 1   Byte 2   Byte 3   Byte 4
            ' 1          0xxxxxxx
            ' 2          110xxxxx 10xxxxxx
            ' 3          1110xxxx 10xxxxxx 10xxxxxx
            ' 4          11110xxx 10xxxxxx 10xxxxxx 10xxxxxx
            '
            If (bytValue And &H80) = &H0 Then
                nCharExtraByteCount = 0
            ElseIf (bytValue And &HE0) = &HC0 Then
                nCharExtraByteCount = 1
            ElseIf (bytValue And &HF0) = &HE0 Then
                nCharExtraByteCount = 2
            ElseIf (bytValue And &HF8) = &HF0 Then
                nCharExtraByteCount = 3
            Else
                ' If none of these masks were matched, then this can't be a UTF-8 character.
                Exit Function
            End If
        Else
            ' All following bytes must be masked as in the table above.
            If (bytValue And &HC0) = &H80 Then
                nCharExtraByteCount = nCharExtraByteCount - 1
                If nCharExtraByteCount = 0 Then
                    pFileBytesAreUTF8 = True
                End If
            Else
                ' Not a UTF8 character.
                Exit Function
            End If
        End If

    Next nByteIndex

End Function

' Purpose: Take a string whose bytes are in the byte array <barrStr>, with code page <lCodePage>, convert to a VB string.
Private Function pFromCPString(ByRef barrStr() As Byte, ByVal lCodePage As ECodePages) As String

    Dim sOutput         As String
    Dim nValueLen       As Long
    Dim nOutputCharLen  As Long

    ' If the code page says this is already compatible with the VB string, then just copy it into the string. No messing.
    If lCodePage = CP_UTF16_LE Then
        pFromCPString = barrStr()
    Else

        ' Cache the input length.
        nValueLen = UBound(barrStr) - LBound(barrStr) + 1

        ' See how big the output buffer will be.
        nOutputCharLen = MultiByteToWideChar(lCodePage, 0&, VarPtr(barrStr(LBound(barrStr))), nValueLen, 0&, 0&)

        ' Resize output byte array to the size of the UTF-8 string.
        sOutput = Space$(nOutputCharLen)

        ' Make this API call again, this time giving a pointer to the output byte array.
        MultiByteToWideChar lCodePage, 0&, VarPtr(barrStr(LBound(barrStr))), nValueLen, StrPtr(sOutput), nOutputCharLen

        ' Return the array.
        pFromCPString = sOutput

    End If

End Function

Public Function GetContents(ByVal sTextFile As String, ByRef OUT_lCodePage As ECodePages, Optional ByVal lDesiredCodePage As ECodePages = -1, Optional ByRef OUT_bContainedBOM As Boolean) As String

    Dim iFileNo            As Integer
    Dim abytFileContents() As Byte
    Dim nDataSize          As Long

    iFileNo = FreeFile

    OpenForInput sTextFile, iFileNo, OUT_lCodePage, lDesiredCodePage, OUT_bContainedBOM

    ' We want to read the entire contents of the file (not including any BOM value).
    ' After calling OpenForInput(), the file pointer should be positioned after any BOM.
    ' So size file contents buffer to <file size> - <current position> + 1.
    nDataSize = LOF(iFileNo) - Seek(iFileNo) + 1
    ReDim abytFileContents(1 To nDataSize)
    Get #iFileNo, , abytFileContents()

    Close iFileNo

    ' Now we must convert this to UTF-8. But we have to first convert to the Windows NT standard UTF-16 LE.
    GetContents = pFromCPString(abytFileContents(), OUT_lCodePage)

End Function

' Purpose:  Reads up to the end of the current line of the file, repositions to the beginning of the next line, if any, and
'           outputs all characters found.
' Inputs:   lFileNo     The number of the file.
' Outputs:  OUT_sLine   The line from the current position in the file.
' Return:   True if there is more data.
Public Function LineInputUTF8(ByVal lFileNo As Integer, ByRef OUT_sLine As String) As Boolean

    Dim bytValue         As Byte
    Dim abytLine()       As Byte
    Dim nStartOfLinePos  As Long
    Dim nEndOfLinePos    As Long
    Dim nStartOfNextLine As Long
    Dim nLineLen         As Long

    ' Save the current file position as the beginning of the line, and cache this value.
    nStartOfLinePos = Seek(lFileNo)

    ' Retrieves the first byte from the current position.
    Get #lFileNo, , bytValue

    ' Loop until the end of file is encountered.
    Do Until EOF(lFileNo)

        ' Check whether this byte represents a carriage return or line feed character (indicating new line).
        If bytValue = 13 Or bytValue = 10 Then
            ' By this point, the current position is *after* the CR or LF character, so to get the position of the
            ' last byte in the line, we must go back two bytes.
            nEndOfLinePos = Seek(lFileNo) - 2

            ' If this is a carriage return, then we must check the next character.
            If bytValue = 13 Then
                Get #lFileNo, , bytValue
                ' Is this a line feed?
                If bytValue = 10 Then
                ' Yes. Assume that CR-LF counts as a single NewLine. So the start of the next line should skip over the line feed.
                    nStartOfNextLine = nEndOfLinePos + 3
                Else
                ' No. The start of the next line is the current position.
                    nStartOfNextLine = nEndOfLinePos + 2
                End If
            ElseIf bytValue = 10 Then
            ' If this is a line feed, then the start of the next line is the current position.
                nStartOfNextLine = nEndOfLinePos + 2
            End If

            ' Since we have processed all the bytes in the line, exit the loop.
            Exit Do
        End If

        ' Get the next byte.
        Get #lFileNo, , bytValue
    Loop

    ' Check to see if there was an end of line.
    If nEndOfLinePos = 0 Then
    ' No, this is the end of the file - so use all the remaining characters.
        nLineLen = Seek(lFileNo) - nStartOfLinePos - 1
    Else
    ' Yes - so use all the characters up to the end of line position.
        nLineLen = nEndOfLinePos - nStartOfLinePos + 1
    End If

    ' Is this line empty?
    If nLineLen = 0 Then
    ' Yes - just return an empty string.
        OUT_sLine = vbNullString
    Else
    ' No - pull all the bytes from the beginning to the end of the line into a byte array, and then convert that from UTF-8 to a VB string.
        ReDim abytLine(1 To nLineLen)
        Get #lFileNo, nStartOfLinePos, abytLine()
        OUT_sLine = pFromCPString(abytLine(), CP_UTF8)
    End If

    ' If there is a line afterwards, then move to the beginning of the line, and return True.
    If nStartOfNextLine > 0 Then
        Seek #lFileNo, nStartOfNextLine
        LineInputUTF8 = True
    End If

End Function

' Purpose:  Analogue of 'Open "fileName" For Input As #fileNo' - but also return what type of text this is via a Code Page value.
' Inputs:   sFilename
'           iFileNo
'           (lDesiredCodePage)  The code page that you want to use with this file.
'                               If this value is set to the default, -1, this indicates that the code page will be ascertained from the file.
' Outputs:  OUT_lCodePage       There are only six valid values that are returned if <lDesiredCodePage> was set to -1.
'               CP_ANSI       ANSI code page
'               CP_UTF8       UTF-8
'               CP_UTF16LE    UTF-16 Little Endian (VB and NT default string encoding)
'               CP_UTF16BE    UTF-16 Big Endian
'               CP_UTF32LE    UTF-32 Little Endian
'               CP_UTF32BE    UTF-32 Big Endian
'           (out_bContainedBOM) If this was set to True, then the file started with a BOM (Byte Order Marker).
Public Sub OpenForInput(ByRef sFilename As String, ByVal iFileNo As Integer, ByRef OUT_lCodePage As ECodePages, Optional ByVal lDesiredCodePage As ECodePages = -1, Optional ByRef OUT_bContainedBOM As Boolean)

    ' Note if we want to take account of every case, we should read in the first four bytes,
    ' and check for UTF-32 low and high endian BOMs, check the first three bytes for the UTF-8 BOM,
    ' and finally check the first two bytes for UTF-16 low and hight endian BOMs.
    Dim abytBOM(1 To 4) As Byte
    Dim nCodePage       As Long

    ' By default, there is no BOM.
    OUT_bContainedBOM = False

    Open sFilename For Binary Access Read As #iFileNo

    ' We are interested in -1 (ascertain code page), and then various UTF encodings.
    Select Case lDesiredCodePage
    Case -1, CP_UTF8, CP_UTF16_BE, CP_UTF16_LE, CP_UTF32_BE, CP_UTF32_LE

        ' Default code page.
        nCodePage = CP_ANSI

        ' Pull in the first four bytes to determine the BOM (byte order marker).
        Get #iFileNo, , abytBOM()

        ' The following are the BOMs for text files:
        '
        ' FF FE         UTF-16, little endian
        ' FE FF         UTF-16, big endian
        ' EF BB BF      UTF-8
        ' FF FE 00 00   UTF-32, little endian
        ' 00 00 FE FF   UTF-32, big-endian
        '
        ' Work out the code page from this information.

        Select Case abytBOM(1)
        Case &HFF
            If abytBOM(2) = &HFE Then
                If abytBOM(3) = 0 And abytBOM(4) = 0 Then
                    nCodePage = CP_UTF32_LE
                Else
                    nCodePage = CP_UTF16_LE
                End If
            End If
        Case &HFE
            If abytBOM(2) = &HFF Then
                nCodePage = CP_UTF16_BE
            End If
        Case &HEF
            If abytBOM(2) = &HBB And abytBOM(3) = &HBF Then
                nCodePage = CP_UTF8
            End If
        Case &H0
            If abytBOM(2) = &H0 And abytBOM(3) = &HFE And abytBOM(4) = &HFF Then
                nCodePage = CP_UTF32_BE
            End If
        End Select

        ' Did we match any BOMs?
        If nCodePage = CP_ANSI Then
        ' No - we are still defaulting to the ANSI code page.
            ' Special check for UTF-8. The BOM is not specified in the standard for UTF-8,
            ' but according to Wikipedia (which is always right :-) ),
            ' only Microsoft includes this marker at the beginning of files.
            If pFileBytesAreUTF8(iFileNo) Then
                OUT_lCodePage = CP_UTF8
            Else
                OUT_lCodePage = CP_ANSI
            End If
        Else
        ' Yes - we have worked out the code page from the BOM.
            ' If no code page was suggested, we now return the code page we found.
            If lDesiredCodePage = -1 Then
                OUT_lCodePage = nCodePage
            End If

            ' Inform the caller that a BOM was found.
            OUT_bContainedBOM = True
        End If

        ' Reset the file pointer to the beginning of the file data.
        If OUT_bContainedBOM Then
            ' Note that if the code page found was one of the two UTF-32 values, then we are already in the correct position.
            ' Otherwise, we have to move to just after the end of the BOM.
            Select Case nCodePage
            Case CP_UTF16_BE, CP_UTF16_LE
                Seek #iFileNo, 3
            Case CP_UTF8
                Seek #iFileNo, 4
            End Select
        Else
            ' There is no BOM, so simply go the beginning of the file.
            Seek #iFileNo, 1
        End If

    Case Else
        OUT_lCodePage = lDesiredCodePage
    End Select

End Sub

' Purpose:  Analogue of 'Open "fileName" For Append As #fileNo'
Public Sub OpenForAppend(ByRef sFilename As String, ByVal iFileNo As Integer, Optional ByVal lCodePage As ECodePages = CP_ANSI, Optional ByVal bPrefixWithBOM As Boolean = True)

    ' Open the file and move to the end of the file.
    Open sFilename For Binary Access Write As #iFileNo
    Seek iFileNo, LOF(iFileNo) + 1

    If bPrefixWithBOM Then
        pWriteBOM iFileNo, lCodePage
    End If

End Sub

' Purpose:  Analogue of 'Open "fileName" For Output As #fileNo'
Public Sub OpenForOutput(ByRef sFilename As String, ByVal iFileNo As Integer, Optional ByVal lCodePage As ECodePages = CP_ANSI, Optional ByVal bPrefixWithBOM As Boolean = True)

    ' Ensure we overwrite the file by deleting it ...
    On Error Resume Next
    Kill sFilename
    On Error GoTo 0

    ' ... before creating it.
    Open sFilename For Binary Access Write As #iFileNo

    If bPrefixWithBOM Then
        pWriteBOM iFileNo, lCodePage
    End If

End Sub

Public Function EOF_(ByVal iFileNo As Integer) As Boolean
    EOF_ = EOF(iFileNo)
End Function

Public Sub Close_(Optional ByVal iFileNo As Integer)
    If IsMissing(iFileNo) Then
        Close
    Else
        Close #iFileNo
    End If
End Sub

' Purpose:  Analogue of the 'Print #fileNo, value' statement. But only one value allowed.
'           Setting <bAppendNewLine> = False is analogous to 'Print #fileNo, value;'.
Public Sub Print_(ByVal iFileNo As Integer, ByRef sValue As String, Optional ByVal lCodePage As ECodePages = CP_ANSI, Optional ByVal bAppendNewLine As Boolean = True)

    Const kbytNull           As Byte = 0
    Const kbytCarriageReturn As Byte = 13
    Const kbytNewLine        As Byte = 10

    Put #iFileNo, , pToCPString(sValue, lCodePage)

    If bAppendNewLine Then
        Select Case lCodePage
        Case CP_UTF16_BE
            Put #iFileNo, , kbytNull
            Put #iFileNo, , kbytCarriageReturn
            Put #iFileNo, , kbytNull
            Put #iFileNo, , kbytNewLine
        Case CP_UTF16_LE
            Put #iFileNo, , kbytCarriageReturn
            Put #iFileNo, , kbytNull
            Put #iFileNo, , kbytNewLine
            Put #iFileNo, , kbytNull
        Case CP_UTF32_BE
            Put #iFileNo, , kbytNull
            Put #iFileNo, , kbytNull
            Put #iFileNo, , kbytNull
            Put #iFileNo, , kbytCarriageReturn
            Put #iFileNo, , kbytNull
            Put #iFileNo, , kbytNull
            Put #iFileNo, , kbytNull
            Put #iFileNo, , kbytNewLine
        Case CP_UTF32_LE
            Put #iFileNo, , kbytCarriageReturn
            Put #iFileNo, , kbytNull
            Put #iFileNo, , kbytNull
            Put #iFileNo, , kbytNull
            Put #iFileNo, , kbytNewLine
            Put #iFileNo, , kbytNull
            Put #iFileNo, , kbytNull
            Put #iFileNo, , kbytNull
        Case Else
            Put #iFileNo, , kbytCarriageReturn
            Put #iFileNo, , kbytNewLine
        End Select
    End If

End Sub

Public Sub PutContents(ByRef sFilename As String, ByRef sFileContents As String, Optional ByVal lCodePage As ECodePages = CP_ANSI, Optional bPrefixWithBOM As Boolean)

    Dim iFileNo As Integer

    iFileNo = FreeFile
    OpenForOutput sFilename, iFileNo, lCodePage, bPrefixWithBOM
    Print_ iFileNo, sFileContents, lCodePage, False
    Close iFileNo

End Sub

' Purpose:  Converts a VB string (UTF-16) to UTF8 - as a binary array.
Private Function pToCPString(ByRef sValue As String, ByVal lCodePage As ECodePages) As Variant 'Byte()

    Dim abytOutput()   As Byte
    Dim nValueLen      As Long
    Dim nOutputByteLen As Long

    If lCodePage = CP_UTF16_LE Then
        pToCPString = sValue
    Else

        ' Cache the input length.
        nValueLen = Len(sValue)

        ' See how big the output buffer will be.
        nOutputByteLen = WideCharToMultiByte(lCodePage, 0&, StrPtr(sValue), nValueLen, 0&, 0&, 0&, 0&)

        If nOutputByteLen > 0 Then
            ' Resize output byte array to the size of the UTF-8 string.
            ReDim abytOutput(1 To nOutputByteLen)

            ' Make this API call again, this time giving a pointer to the output byte array.
            WideCharToMultiByte lCodePage, 0&, StrPtr(sValue), nValueLen, VarPtr(abytOutput(1)), nOutputByteLen, 0&, 0&
        End If

        ' Return the array.
        pToCPString = abytOutput()

    End If

End Function

Private Sub pWriteBOM(ByVal iFileNo As Integer, ByVal lCodePage As ECodePages)

    ' FF FE         UTF-16, little endian
    ' FE FF         UTF-16, big endian
    ' EF BB BF      UTF-8
    ' FF FE 00 00   UTF-32, little endian
    ' 00 00 FE FF   UTF-32, big-endian

    Select Case lCodePage
        Case CP_UTF8
            Put #iFileNo, , CByte(&HEF)
            Put #iFileNo, , CByte(&HBB)
            Put #iFileNo, , CByte(&HBF)
        Case CP_UTF16_LE
            Put #iFileNo, , CByte(&HFF)
            Put #iFileNo, , CByte(&HFE)
        Case CP_UTF16_BE
            Put #iFileNo, , CByte(&HFE)
            Put #iFileNo, , CByte(&HFF)
        Case CP_UTF32_LE
            Put #iFileNo, , CByte(&HFF)
            Put #iFileNo, , CByte(&HFE)
            Put #iFileNo, , CByte(&H0)
            Put #iFileNo, , CByte(&H0)
        Case CP_UTF32_BE
            Put #iFileNo, , CByte(&H0)
            Put #iFileNo, , CByte(&H0)
            Put #iFileNo, , CByte(&HFE)
            Put #iFileNo, , CByte(&HFF)
    End Select

End Sub

' Exemplos de uso
'
'Purpose:  Converts *any* pure text file (UTF16, ASCII, ANSI) to UTF8.
'Private Sub Example_ConvertFileToUTF8()
'
'    Dim nCodePage     As Long
'    Dim bContainedBOM As Boolean
'    Dim sFileContents As String
'
'    ' Read in contents.
'    sFileContents = GetContents("C:\MysteryEncoding.txt", nCodePage, , bContainedBOM)
'
'    ' And then convert to UTF8.
'    PutContents "C:\output.txt", sFileContents, CP_UTF8, bContainedBOM
'
'End Sub
'
' Purpose:  Iterates through each line of a UTF-8 text file, and adds it to a control which can display
'           VB strings containing non-ANSI characters. In this case, I am adding items to a FlexGrid
'           with Font = "Lucida Console".
'Private Sub Example_IterateUTF8Lines()
'
'    Dim iFileNo             As Integer
'    Dim lCodePage           As Long
'    Dim sLine               As String
'
'    iFileNo = FreeFile
'
'    TextFile.OpenForInput "C:\UTF8.txt", iFileNo, lCodePage
'
'    If lCodePage = CP_UTF8 Then
'        Do While TextFile.LineInputUTF8(iFileNo, sLine)
'            VSFlexGrid.AddItem sLine
'        Loop
'        VSFlexGrid.AddItem sLine
'    Else
'        MsgBox "This is not a UTF8 file."
'    End If
'
'    Close #iFileNo
'
'End Sub
'
'Private Sub Example_ReadWriteUTF8Lines()
'
'    Dim iFileNoInput        As Integer
'    Dim iFileNoOutput       As Integer
'    Dim lCodePage           As Long
'    Dim bBOM                As Boolean
'    Dim sLine               As String
'
'    iFileNoInput = FreeFile
'    TextFile.OpenForInput "C:\UTF8.txt", iFileNoInput, lCodePage, , bBOM
'
'    If lCodePage = CP_UTF8 Then
'
'        iFileNoOutput = FreeFile
'        TextFile.OpenForOutput "C:\output.txt", iFileNoOutput, lCodePage, bBOM
'
'        Do While TextFile.LineInputUTF8(iFileNoInput, sLine)
'            TextFile.Print_ iFileNoOutput, sLine, lCodePage
'        Loop
'        TextFile.Print_ iFileNoOutput, sLine, lCodePage, False
'
'        Close #iFileNoOutput
'
'    Else
'        MsgBox "This is not a UTF8 file."
'    End If
'
'    Close #iFileNoInput
'
'End Sub
