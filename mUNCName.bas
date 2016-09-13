Attribute VB_Name = "mUNCName"
' *********************************************************************
'  Copyright ©1999 Karl E. Peterson, All Rights Reserved
'  http://www.mvps.org/vb
' *********************************************************************
'  You are free to use this code within your own applications, but you
'  are expressly forbidden from selling or otherwise distributing this
'  source code without prior written consent.
' *********************************************************************
Option Explicit

Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long

Private Type OSVERSIONINFO
   dwOSVersionInfoSize As Long
   dwMajorVersion As Long
   dwMinorVersion As Long
   dwBuildNumber As Long
   dwPlatformId As Long
   szCSDVersion As String * 128
End Type

'Private Const VER_PLATFORM_WIN32s = 0
Private Const VER_PLATFORM_WIN32_WINDOWS = 1
Private Const VER_PLATFORM_WIN32_NT = 2

Private Declare Function WNetGetUniversalName Lib "mpr" Alias "WNetGetUniversalNameA" (ByVal lpLocalPath As String, ByVal dwInfoLevel As Long, lpBuffer As Any, lpBufferSize As Long) As Long
Private Declare Function WNetGetConnection Lib "mpr" Alias "WNetGetConnectionA" (ByVal lpLocalName As String, lpRemoteName As Any, lpnLength As Long) As Long
Private Declare Function NetShareEnum Lib "netapi32" (ByVal lpServerName As Long, ByVal dwLevel As Long, lpBuffer As Any, ByVal dwPrefMaxLen As Long, EntriesRead As Long, TotalEntries As Long, hResume As Long) As Long
Private Declare Function NetShareEnum95 Lib "svrapi" Alias "NetShareEnum" (ByVal lpServerName As String, ByVal dwLevel As Long, lpBuffer As Any, ByVal cbBuffer As Long, EntriesRead As Long, TotalEntries As Long) As Long
Private Declare Function NetApiBufferFree Lib "netapi32" (ByVal lpBuffer As Long) As Long
Private Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pTo As Any, uFrom As Any, ByVal lSize As Long)
Private Declare Function lstrlenA Lib "kernel32" (ByVal PointerToString As Long) As Long
Private Declare Function lstrlenW Lib "kernel32" (ByVal PointerToString As Long) As Long

'Private Const UNIVERSAL_NAME_INFO_LEVEL = &H1
Private Const REMOTE_NAME_INFO_LEVEL = &H2

'Private Type UNIVERSAL_NAME_INFO
'   lpUniversalName As Long
'End Type
'
Private Type REMOTE_NAME_INFO
   lpUniversalName As Long
   lpConnectionName As Long
   lpRemainingPath As Long
End Type

Private Const MAX_PREFERRED_LENGTH As Long = -1
Private Const MAX_PATH = 260

Private Const NO_ERROR = 0
Private Const ERROR_ACCESS_DENIED = 5&

'Private Const ERROR_BAD_DEVICE = 1200&
Private Const ERROR_NOT_CONNECTED = 2250&
Private Const ERROR_MORE_DATA = 234
'Private Const ERROR_CONNECTION_UNAVAIL = 1201&
'Private Const ERROR_NO_NETWORK = 1222&
'Private Const ERROR_EXTENDED_ERROR = 1208&
'Private Const ERROR_NO_NET_OR_BAD_PATH = 1203&

Private Const STYPE_DISKTREE = 0      ' /* disk share */
'Private Const STYPE_PRINTQ = 1        ' /* printer share */
'Private Const STYPE_DEVICE = 2
'Private Const STYPE_IPC = 3
'Private Const STYPE_SPECIAL = &H80000000

Private Type SHARE_INFO_2
   Netname As String
   ShareType As Long
   Remark As String
   Permissions As Long
   MaxUsers As Long
   CurrentUsers As Long
   path As String
   Password As String
End Type

Private Const LM20_NNLEN = 12         ' // LM 2.0 Net name length
Private Const SHPWLEN = 8             ' // Share password length (bytes)

Private Type SHARE_INFO_50            'struct _share_info_50 {
   Netname(0 To LM20_NNLEN) As Byte   '   char            shi50_netname[LM20_NNLEN+1];
   ShareType As Byte                  '   unsigned char   shi50_type;
   Flags As Integer                   '   unsigned short  shi50_flags;
   lpRemark As Long                   '   char FAR *      shi50_remark;
   lpPath As Long                     '   char FAR *      shi50_path;
   PasswordRW(0 To SHPWLEN) As Byte   '   char            shi50_rw_password[SHPWLEN+1];
   PasswordRO(0 To SHPWLEN) As Byte   '   char            shi50_ro_password[SHPWLEN+1];
End Type                              '};

Public uncNeedAdminPrivs As Boolean

Public Function GetUncNameGF(ByVal FileSpec As String, ByRef UNCPrecisaDePrivilegiosDeAdministrador As Boolean) As String
   Dim Buffer() As Byte
   Dim nRet As Long
   Dim BufferLen As Long
   Dim rni As REMOTE_NAME_INFO
   Dim shi() As SHARE_INFO_2
   Dim i As Long
   
   ' make sure we actually have what looks like a drive-based spec
   On Error Resume Next
   If Asc(UCase$(Left$(FileSpec, 1))) < Asc("A") Then
      Exit Function
   ElseIf Asc(UCase$(Left$(FileSpec, 1))) > Asc("Z") Then
      Exit Function
   ElseIf Mid$(FileSpec, 2, 1) <> ":" Then
      Exit Function
   End If
   If Err.Number Then Exit Function
   On Error GoTo 0
   
   
   If IsWin95 Then
      ' ***************************************************************
      ' ** Important note: Q131416 states that WNetGetUniversalName  **
      ' **                 always fails under Win95.  :-(            **
      ' ***************************************************************
      ReDim Buffer(1 To MAX_PATH) As Byte
      nRet = WNetGetConnection(Left$(FileSpec, 2), Buffer(1), UBound(Buffer))
      If nRet = NO_ERROR Then
         ' Success! Obtained the universal name for the share
         GetUncNameGF = TrimNull(StrConv(Buffer, vbUnicode)) & Mid$(FileSpec, 3)
         Exit Function
      End If
   Else
      ' in NT/98, call it once to get required size of structure
      nRet = WNetGetUniversalName(FileSpec, REMOTE_NAME_INFO_LEVEL, vbNullString, BufferLen)
   End If
   
   
   Select Case nRet
   
      Case ERROR_MORE_DATA
         ' resize buffer and call again
         ReDim Buffer(0 To BufferLen - 1) As Byte
         nRet = WNetGetUniversalName(FileSpec, REMOTE_NAME_INFO_LEVEL, Buffer(0), BufferLen)
         ' extract UNC name from buffer
         If nRet = NO_ERROR Then
            ' retrieve pointers to each of the returned strings
            rni.lpUniversalName = PointerToDWord(VarPtr(Buffer(0)))
            rni.lpConnectionName = PointerToDWord(VarPtr(Buffer(4)))
            rni.lpRemainingPath = PointerToDWord(VarPtr(Buffer(8)))
            ' extract and return lpUniversalName
            GetUncNameGF = PointerToStringA(rni.lpUniversalName)
            ' show lpConnectionName and lpRemainingPath just for kicks?
            ' Debug.Print PointerToStringA(rni.lpConnectionName)
            ' Debug.Print PointerToStringA(rni.lpRemainingPath)
         End If
      
      Case ERROR_NOT_CONNECTED
         ' user passed a local filename, default to passed spec
         GetUncNameGF = FileSpec
         ' get list of local shares
         nRet = EnumShares(shi)
         If nRet > 0 Then
            ' loop through shares, looking for a potential match
            ' ambiguous: any path can be on more than one share
            For i = 0 To nRet - 1
               If shi(i).ShareType = STYPE_DISKTREE Then
                  If InStr(1, FileSpec, shi(i).path, vbTextCompare) = 1 Then
                     ' this element starts with the same path
                     ' have to accept first match
                     GetUncNameGF = "\\" & CurrentMachineName & _
                                    "\" & shi(i).Netname & _
                                    Mid(FileSpec, Len(shi(i).path))
                     Exit For
                  End If
               End If
            Next
         End If
      
      Case Else
         ' bad path or network down
         GetUncNameGF = vbNullString
   End Select
   
   UNCPrecisaDePrivilegiosDeAdministrador = uncNeedAdminPrivs
   
End Function

Private Function EnumShares(shi() As SHARE_INFO_2, Optional ByVal Server As String = "") As Long
   Dim nRet As Long
   ' delegate to OS-appropriate routine
   If IsWinNT() Then
      nRet = EnumSharesNT(shi(), Server)
      uncNeedAdminPrivs = (nRet < 0)
   Else
      nRet = EnumShares9x(shi(), Server)
   End If
   EnumShares = nRet
End Function

Private Function EnumSharesNT(shi() As SHARE_INFO_2, Optional ByVal Server As String = "") As Long
   Dim Level As Long
   Dim lpBuffer As Long
   Dim EntriesRead As Long
   Dim TotalEntries As Long
   Dim hResume As Long
   Dim Offset As Long
   Dim nRet As Long
   Dim i As Long
   
   ' convert Server to null pointer if none requested.
   ' this has the effect of asking for the local machine.
   If Len(Server) = 0 Then Server = vbNullString

   ' ask for all available shares; try level 2 first
   Level = 2
   nRet = NetShareEnum(StrPtr(Server), Level, lpBuffer, MAX_PREFERRED_LENGTH, EntriesRead, TotalEntries, hResume)
   
   If nRet = ERROR_ACCESS_DENIED Then
      ' bummer -- need admin privs for level 2, drop to level 1
      Level = 1
      nRet = NetShareEnum(StrPtr(Server), Level, lpBuffer, MAX_PREFERRED_LENGTH, EntriesRead, TotalEntries, hResume)
   End If
   
   If nRet = NO_ERROR Then
      ' make sure there are shares to decipher
      If EntriesRead > 0 Then
         ' prepare UDT buffer to hold all share info
         ReDim shi(0 To EntriesRead - 1)
         ' loop through API buffer, extracting each element
         For i = 0 To EntriesRead - 1
            With shi(i)
               .Netname = PointerToStringW(PointerToDWord(lpBuffer + Offset))
               .ShareType = PointerToDWord(lpBuffer + Offset + 4)
               .Remark = PointerToStringW(PointerToDWord(lpBuffer + Offset + 8))
               If Level = 2 Then
                  .Permissions = PointerToDWord(lpBuffer + Offset + 12)
                  .MaxUsers = PointerToDWord(lpBuffer + Offset + 16)
                  .CurrentUsers = PointerToDWord(lpBuffer + Offset + 20)
                  .path = PointerToStringW(PointerToDWord(lpBuffer + Offset + 24))
                  .Password = PointerToStringW(PointerToDWord(lpBuffer + Offset + 28))
                  Offset = Offset + Len(shi(i))
               Else
                  Offset = Offset + 12  ' Len(SHARE_INFO_1)
               End If
            End With
         Next
      End If
      
      ' return number of entries found
      If Level = 1 Then
         ' negative if we don't have admin privs
         EnumSharesNT = -EntriesRead
      ElseIf Level = 2 Then
         EnumSharesNT = EntriesRead
      End If
   End If
   
   ' clean up
   If lpBuffer Then
      Call NetApiBufferFree(lpBuffer)
   End If
End Function

Private Function EnumShares9x(shi() As SHARE_INFO_2, Optional ByVal Server As String = "") As Long
   Dim Buffer() As Byte
   Dim EntriesRead As Long
   Dim TotalEntries As Long
   Dim Offset As Long
   Dim shi95 As SHARE_INFO_50
   Dim nRet As Long
   Dim i As Long
   Const BufferSize = &H4000
   
   ' convert Server to null pointer if none requested.
   ' this has the effect of asking for the local machine.
   If Len(Server) = 0 Then Server = vbNullString

   ' ask for all available shares, using really large buffer
   ReDim Buffer(0 To BufferSize - 1) As Byte
   nRet = NetShareEnum95(Server, 50, Buffer(0), BufferSize, EntriesRead, TotalEntries)
   
   If nRet = NO_ERROR Then
      ' make sure there are shares to decipher
      If EntriesRead > 0 Then
         ' prepare UDT buffer to hold all share info
         ReDim shi(0 To EntriesRead - 1)
         ' loop through API buffer, extracting each element
         For i = 0 To EntriesRead - 1
            With shi(i)
               Call CopyMemory(shi95, Buffer(Offset), Len(shi95))
               .Netname = TrimNull(StrConv(shi95.Netname, vbUnicode))
               .ShareType = shi95.ShareType
               .path = PointerToStringA(shi95.lpPath)
               .Remark = PointerToStringA(shi95.lpRemark)
               If shi95.PasswordRW(0) = 0 Then
                  .Password = TrimNull(StrConv(shi95.PasswordRO, vbUnicode))
               Else
                  .Password = TrimNull(StrConv(shi95.PasswordRW, vbUnicode))
               End If
               Offset = Offset + Len(shi95)
            End With
         Next
         
         ' return number of entries found
         EnumShares9x = EntriesRead
      End If
   End If
End Function

Private Function IsWinNT() As Boolean
   Static os As OSVERSIONINFO
   Static bRet As Boolean
   ' just do this once, for optimization
   If os.dwPlatformId = 0 Then
      os.dwOSVersionInfoSize = Len(os)
      Call GetVersionEx(os)
      bRet = (os.dwPlatformId = VER_PLATFORM_WIN32_NT)
   End If
   IsWinNT = bRet
End Function

Private Function IsWin95() As Boolean
   Static os As OSVERSIONINFO
   Static bRet As Boolean
   ' just do this once, for optimization
   If os.dwPlatformId = 0 Then
      os.dwOSVersionInfoSize = Len(os)
      Call GetVersionEx(os)
      bRet = (os.dwMinorVersion < 10) And _
             (os.dwPlatformId = VER_PLATFORM_WIN32_WINDOWS)
   End If
   IsWin95 = bRet
End Function

Private Function CurrentMachineName() As String
   Dim Buffer As String
   Dim nLen As Long
   Const CNLEN = 15          ' Maximum computer name length
   
   Buffer = Space$(CNLEN + 1)
   nLen = Len(Buffer)
   If GetComputerName(Buffer, nLen) Then
      CurrentMachineName = Left$(Buffer, nLen)
   End If
End Function

Private Function PointerToDWord(ByVal lpDWord As Long) As Long
   Call CopyMemory(PointerToDWord, ByVal lpDWord, 4)
End Function

Private Function PointerToStringA(lpStringA As Long) As String
   Dim Buffer() As Byte
   Dim nLen As Long
   
   If lpStringA Then
      nLen = lstrlenA(ByVal lpStringA)
      If nLen Then
         ReDim Buffer(0 To (nLen - 1)) As Byte
         CopyMemory Buffer(0), ByVal lpStringA, nLen
         PointerToStringA = StrConv(Buffer, vbUnicode)
      End If
   End If
End Function

Private Function PointerToStringW(ByVal lpStringW As Long) As String
   Dim Buffer() As Byte
   Dim nLen As Long
   
   If lpStringW Then
      nLen = lstrlenW(lpStringW) * 2
      If nLen Then
         ReDim Buffer(0 To (nLen - 1)) As Byte
         CopyMemory Buffer(0), ByVal lpStringW, nLen
         PointerToStringW = Buffer
      End If
   End If
End Function

Private Function TrimNull(ByVal StrIn As String) As String
   Dim nul As Long
   '
   ' Truncate input string at first null.
   ' If no nulls, perform ordinary Trim.
   '
   nul = InStr(StrIn, vbNullChar)
   Select Case nul
      Case Is > 1
         TrimNull = Left$(StrIn, nul - 1)
      Case 1
         TrimNull = vbNullString
      Case 0
         TrimNull = Trim$(StrIn)
   End Select
End Function


