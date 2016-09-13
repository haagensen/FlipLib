VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WinDirs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'------------------------------------------------------------
' Usado por CheckFileAccess
'------------------------------------------------------------
Public Enum EAcessos
    MAXIMUM_ALLOWED = &H2000000
    ACCESS_DELETE = &H10000 ' "Delete"
    ACCESS_READ_CONTROL = &H20000 ' "Read Permissions"
    ACCESS_WRITE_DAC = &H40000 ' "Change Permissions"
    ACCESS_WRITE_OWNER = &H80000 ' "Take ownership"
    ACCESS_SYNCHRONIZE = &H100000 ' "Synchronize"
    STANDARD_RIGHTS_READ = ACCESS_READ_CONTROL
    STANDARD_RIGHTS_WRITE = ACCESS_READ_CONTROL
    STANDARD_RIGHTS_EXECUTE = ACCESS_READ_CONTROL
    STANDARD_RIGHTS_REQUIRED = &HF0000
    FILE_READ_DATA = &H1              '  file & pipe - "Read Data/List Folder"
    FILE_LIST_DIRECTORY = &H1         '  directory
    FILE_ADD_FILE = &H2               '  directory
    FILE_WRITE_DATA = &H2             '  file & pipe - "Write Data/Create Files"
    FILE_CREATE_PIPE_INSTANCE = &H4   '  named pipe
    FILE_ADD_SUBDIRECTORY = &H4       '  directory
    FILE_APPEND_DATA = &H4            '  file - "Append Data/Create Folders"
    FILE_READ_EA = &H8                '  file & directory - "Read Extended Attributes"
    FILE_READ_PROPERTIES = FILE_READ_EA
    FILE_WRITE_EA = &H10              '  file & directory - "Write Extended Attributes"
    FILE_WRITE_PROPERTIES = FILE_WRITE_EA
    FILE_EXECUTE = &H20               '  file - "Execute File/Traverse Folder"
    FILE_TRAVERSE = &H20              '  directory
    FILE_DELETE_CHILD = &H40          '  directory - "Delete Subfolders and Files"
    FILE_READ_ATTRIBUTES = &H80       '  all - "Read Attributes"
    FILE_WRITE_ATTRIBUTES = &H100     '  all - "Write Attributes"
    FILE_GENERIC_READ = (STANDARD_RIGHTS_READ Or FILE_READ_DATA Or FILE_READ_ATTRIBUTES Or FILE_READ_EA Or ACCESS_SYNCHRONIZE)
    FILE_GENERIC_WRITE = (STANDARD_RIGHTS_WRITE Or FILE_WRITE_DATA Or FILE_WRITE_ATTRIBUTES Or FILE_WRITE_EA Or FILE_APPEND_DATA Or ACCESS_SYNCHRONIZE)
    FILE_GENERIC_EXECUTE = (STANDARD_RIGHTS_EXECUTE Or FILE_READ_ATTRIBUTES Or FILE_EXECUTE Or ACCESS_SYNCHRONIZE)
    FILE_ALL_ACCESS = (STANDARD_RIGHTS_REQUIRED Or ACCESS_SYNCHRONIZE Or &H1FF&)
    GENERIC_READ = &H80000000
    GENERIC_WRITE = &H40000000
    GENERIC_EXECUTE = &H20000000
    GENERIC_ALL = &H10000000
End Enum

' Types, constants and functions to work with access rights
Const OWNER_SECURITY_INFORMATION As Long = &H1
Const GROUP_SECURITY_INFORMATION As Long = &H2
Const DACL_SECURITY_INFORMATION  As Long = &H4
Const TOKEN_QUERY                As Long = 8
Const SECURITY_IMPERSONATION As Integer = 3
Const ANYSIZE_ARRAY = 1
Private Type GENERIC_MAPPING
    GenericRead As Long
    GenericWrite As Long
    GenericExecute As Long
    GenericAll As Long
End Type
Private Type LUID
    LowPart As Long
    HighPart As Long
End Type
Private Type LUID_AND_ATTRIBUTES
    pLuid As LUID
    Attributes As Long
End Type
Private Type PRIVILEGE_SET
    PrivilegeCount As Long
    Control As Long
    Privilege(ANYSIZE_ARRAY) As LUID_AND_ATTRIBUTES
End Type
Private Declare Function GetFileSecurity Lib "advapi32.dll" Alias "GetFileSecurityA" (ByVal lpFileName As String, ByVal RequestedInformation As Long, pSecurityDescriptor As Byte, ByVal nLength As Long, lpnLengthNeeded As Long) As Long
Private Declare Function AccessCheck Lib "advapi32.dll" (pSecurityDescriptor As Byte, ByVal ClientToken As Long, ByVal DesiredAccess As Long, GenericMapping As GENERIC_MAPPING, PrivilegeSet As PRIVILEGE_SET, PrivilegeSetLength As Long, GrantedAccess As Long, Status As Long) As Long
Private Declare Function ImpersonateSelf Lib "advapi32.dll" (ByVal ImpersonationLevel As Integer) As Long
Private Declare Function RevertToSelf Lib "advapi32.dll" () As Long
Private Declare Function GetCurrentThread Lib "kernel32" () As Long
Private Declare Function OpenThreadToken Lib "advapi32.dll" (ByVal ThreadHandle As Long, ByVal DesiredAccess As Long, ByVal OpenAsSelf As Long, TokenHandle As Long) As Long
Private Declare Sub MapGenericMask Lib "advapi32.dll" (AccessMask As Long, GenericMapping As GENERIC_MAPPING)
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long

Const FS_PERSISTENT_ACLS As Long = &H8
Private Declare Function GetVolumeInformation Lib "kernel32" Alias "GetVolumeInformationA" (ByVal lpRootPathName As String, ByVal lpVolumeNameBuffer As String, ByVal nVolumeNameSize As Long, lpVolumeSerialNumber As Long, lpMaximumComponentLength As Long, lpFileSystemFlags As Long, ByVal lpFileSystemNameBuffer As String, ByVal nFileSystemNameSize As Long) As Long

Private Type OSVERSIONINFO
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformId As Long
    szCSDVersion As String * 128      '  Maintenance string for PSS usage
End Type
Private Const VER_PLATFORM_WIN32_NT As Long = 2
Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long


'------------------------------------------------------------
' Usado p/ obter diretório "All Users"
'------------------------------------------------------------
Private Declare Function SHGetSpecialFolderLocationAPI _
    Lib "shell32.dll" _
    Alias "SHGetSpecialFolderLocation" (ByVal hwndOwner As Long, ByVal nFolder As Long, pidl As Long) As Long
Private Declare Function SHGetPathFromIDListAPI _
    Lib "shell32.dll" _
    Alias "SHGetPathFromIDListA" (ByVal pidl As Long, ByVal pszPath As String) As Long
Private Declare Sub CoTaskMemFreeAPI _
    Lib "ole32.dll" _
    Alias "CoTaskMemFree" (ByVal pv As Long)
Private Declare Function GetWindowsDirectory _
    Lib "kernel32" _
    Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long

Private Declare Function GetComputerNameAPI Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Private Declare Function GetDriveTypeAPI Lib "kernel32" Alias "GetDriveTypeA" (ByVal nDrive As String) As Long
Private Declare Function GetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Private Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long

Public Function GetComputerName() As String

   Dim sBuffer As String
   Dim nLen As Long
   Const MAX_COMPUTERNAME_LENGTH = 15
   
   sBuffer = Space$(MAX_COMPUTERNAME_LENGTH + 1)
   nLen = Len(sBuffer)
   If GetComputerNameAPI(sBuffer, nLen) Then
      GetComputerName = Left$(sBuffer, nLen)
   End If

End Function

Public Function GetTempDir() As String

    Dim sFolder As String
    Dim lRet As Long

    sFolder = String(255, 0)
    lRet = GetTempPath(255, sFolder)

    If lRet <> 0 Then
        GetTempDir = Left(sFolder, InStr(sFolder, vbNullChar) - 1)
        If Right$(GetTempDir, 1) <> "\" Then
            GetTempDir = GetTempDir & "\"
        End If
    End If

End Function

Public Function GetMyDocumentsDir() As String

    Const CSIDL_PERSONAL As Long = &H5 'My Documents
    Const NOERROR = 0
    Const MAX_PATH = 260
    Dim pidl As Long
    Dim specialFDLRPath As String
    Dim sPath As String
    Dim lRetVal As Long

    On Error Resume Next

    lRetVal = SHGetSpecialFolderLocationAPI(0&, CSIDL_PERSONAL, pidl)

    If lRetVal = NOERROR Then
        sPath = Space$(MAX_PATH)
        lRetVal = SHGetPathFromIDListAPI(pidl, sPath)
        If lRetVal = 1 Then
            GetMyDocumentsDir = Left$(sPath, InStr(sPath, Chr$(0)) - 1)
            If Right$(GetMyDocumentsDir, 1) <> "\" Then
                GetMyDocumentsDir = GetMyDocumentsDir & "\"
            End If
        End If
        Call CoTaskMemFreeAPI(pidl)
    End If

End Function

Public Function GetAllUsersAppData() As String

    Const CSIDL_COMMON_APPDATA = &H23 '(all users)\Application Data
    Const NOERROR = 0
    Const MAX_PATH = 260
    Dim pidl As Long
    Dim specialFDLRPath As String
    Dim sPath As String
    Dim lRetVal As Long

    On Error Resume Next

    lRetVal = SHGetSpecialFolderLocationAPI(0&, CSIDL_COMMON_APPDATA, pidl)

    If lRetVal = NOERROR Then
        sPath = Space$(MAX_PATH)
        lRetVal = SHGetPathFromIDListAPI(pidl, sPath)
        If lRetVal = 1 Then
            GetAllUsersAppData = Left$(sPath, InStr(sPath, Chr$(0)) - 1)
            If Right$(GetAllUsersAppData, 1) <> "\" Then
                GetAllUsersAppData = GetAllUsersAppData & "\"
            End If
        End If
        Call CoTaskMemFreeAPI(pidl)
    End If

End Function

Public Function GetWindowsDir() As String

    Const MAX_PATH = 260
    Dim nSize As Long
    Dim buff As String
   
    buff = Space$(MAX_PATH)
    nSize = Len(buff)
    Call GetWindowsDirectory(buff, nSize)
    GetWindowsDir = zTrimNull(buff)
    If Right$(GetWindowsDir, 1) <> "\" Then
        GetWindowsDir = GetWindowsDir & "\"
    End If

End Function

Public Function GetSystemDir() As String

    Dim buffer As String
    Dim X As Long
    buffer = Space$(255)
    X = GetSystemDirectory(buffer, Len(buffer))
    If X Then
        GetSystemDir = Left$(buffer, InStr(1, buffer, vbNullChar) - 1)
        If Right$(GetSystemDir, 1) <> "\" Then
            GetSystemDir = GetSystemDir & "\"
        End If
    End If

End Function

Private Function zTrimNull(ByVal Item As String) As String

   Dim pos As Integer

   pos = InStr(Item, Chr$(0))

   If pos Then
      zTrimNull = Left$(Item, pos - 1)
   Else
      zTrimNull = Item
   End If

End Function

'Private Function pfObtemDriveWindows() As String
'
'    '-----------------------------------------------------------------
'    ' Retorna drive onde o Windows está instalado.
'    ' Em caso de erro, retorna "C:\"
'    '-----------------------------------------------------------------
'
'    Dim windir As String, pos As Integer, caracter As String
'
'    On Error Resume Next
'
'    windir = GetWindowsDirGF
'
'    If Mid$(windir, 2, 2) = ":\" Then
'        ' Path está no formato "x:\windowsdir". Retorna "x:\"
'        pfObtemDriveWindows = Left$(windir, 3)
'    ElseIf Left$(windir, 2) = "\\" Then
'        ' Path está no formato UNC ("\\servidor\diretorio"), então
'        ' procura pela primeira ocorrência de "\" após a "\\" inicial.
'        For pos = 3 To Len(windir)
'            caracter = Mid$(windir, pos, 1)
'            If caracter = "\" Then Exit For
'        Next
'        pfObtemDriveWindows = Left$(windir, pos)
'    Else
'        ' ops!
'        pfObtemDriveWindows = "C:\"
'    End If
'
'End Function

Public Function IsNetworkDrive(ByVal sDiretorioRaiz As String) As Boolean
    Const DRIVE_REMOTE = 4
    IsNetworkDrive = (GetDriveTypeAPI(sDiretorioRaiz) = DRIVE_REMOTE)
End Function

Public Function IsCDROMDrive(ByVal sDiretorioRaiz As String) As Boolean
    Const DRIVE_CDROM = 5
    IsCDROMDrive = (GetDriveTypeAPI(sDiretorioRaiz) = DRIVE_CDROM)
End Function

Public Function IsRemovableDrive(ByVal sDiretorioRaiz As String) As Boolean
    Const DRIVE_REMOVABLE = 2
    IsRemovableDrive = (GetDriveTypeAPI(sDiretorioRaiz) = DRIVE_REMOVABLE)
End Function

Public Function IsFixedDrive(ByVal sDiretorioRaiz As String) As Boolean
    Const DRIVE_FIXED = 3
    IsFixedDrive = (GetDriveTypeAPI(sDiretorioRaiz) = DRIVE_FIXED)
End Function

''
' CheckFileAccess function checks access rights to given file.
'
' Filename - file or directory full path. Directory path must not end on "\" character.
' DesiredAccess - desired access rights bit mask.
'
' The function returns a bit mask which consists of those bits of desired bit mask, which correspond with allowed access rights.
' In case of access rights to given file or directory not supported, the function returns -1 value.
'
' As desired access mask you may use any combination with OR operator of constants from the beginning of CheckFileAccess function
' listing. The most popular of them are:
'
' FILE_GENERIC_READ - read access,
' FILE_GENERIC_WRITE - write access,
' FILE_GENERIC_EXECUTE - execute access,
' DELETE - delete access,
' WRITE_DAC - change access rights access,
' WRITE_OWNER - change owner access,
' FILE_ALL_ACCESS - full access,
' MAXIMUM_ALLOWED - maximal allowed access.
'
' It is also possible to use constants, applicable to any secure OS objects:
'
' GENERIC_READ - read access,
' GENERIC_WRITE - write access,
' GENERIC_EXECUTE - execute access,
' GENERIC_ALL - full access,
' but in this case the function returns correspondingly values FILE_GENERIC_READ, FILE_GENERIC_WRITE, FILE_GENERIC_EXECUTE, FILE_ALL_ACCESS
' (of course, if correspondent rights exist).
'
' http://www.freevbcode.com/ShowCode.asp?ID=4318
' http://www.smsoft.ru/en/fileperm.htm
'
Public Function CheckFileAccess(ByVal sFilename As String, ByVal lDesiredAccess As EAcessos) As EAcessos

    Dim vbf As VB6: Set vbf = New VB6
    Dim r As Long, SecDesc() As Byte, SDSize As Long, hToken As Long
    Dim PrivSet As PRIVILEGE_SET, GenMap As GENERIC_MAPPING
    Dim sVolume As String, FSFlags As Long

    If Not zIsNT() Then
        CheckFileAccess = -1
        Exit Function
    End If

    If Left$(sFilename, 2) = "\\" Then 'UNC
        'r = vbf.InStr(3, sFilename, "\")
        r = vbf.InStrRev(sFilename, "\")
        If r = 0 Then
            sVolume = sFilename & "\"
        Else
            sVolume = Left$(sFilename, r)
        End If
    ElseIf Mid$(sFilename, 2, 2) = ":\" Then
        sVolume = Left$(sFilename, 3)
    'Else
    ' If path not set, we are leaving Volume blank.
    ' It returns information about current drive.
    End If

    ' FSFlags é parâmetro de saída, recebe os flags associados com o sistema de arquivos especificado.
    ' FILE_PERSISTENT_ACLS: "The specified volume preserves and enforces access control lists (ACL).
    ' For example, the NTFS file system preserves and enforces ACLs, and the FAT file system does not".
    GetVolumeInformation sVolume, vbNullString, 0, ByVal 0&, ByVal 0&, FSFlags, vbNullString, 0
    If (FSFlags And FS_PERSISTENT_ACLS) = 0 Then
        CheckFileAccess = -2
        Exit Function
    End If

    ' Determination of buffer size
    GetFileSecurity sFilename, OWNER_SECURITY_INFORMATION Or GROUP_SECURITY_INFORMATION Or DACL_SECURITY_INFORMATION, 0, 0, SDSize
    If Err.LastDllError <> 122 Then
        CheckFileAccess = -3
        Exit Function
    End If
    If SDSize = 0 Then
        CheckFileAccess = -4
        Exit Function
    End If

    ReDim SecDesc(1 To SDSize)
    ' One more call of function to obtain Security Descriptor
    If GetFileSecurity(sFilename, OWNER_SECURITY_INFORMATION Or GROUP_SECURITY_INFORMATION Or DACL_SECURITY_INFORMATION, SecDesc(1), SDSize, SDSize) = 0 Then
        CheckFileAccess = -5
        Exit Function
    End If

    ' Adding Impersonation Token for thread
    ImpersonateSelf SECURITY_IMPERSONATION
    OpenThreadToken GetCurrentThread(), TOKEN_QUERY, 0, hToken
    If hToken <> 0 Then
        ' Filling GenericMask type
        GenMap.GenericRead = EAcessos.FILE_GENERIC_READ
        GenMap.GenericWrite = EAcessos.FILE_GENERIC_WRITE
        GenMap.GenericExecute = EAcessos.FILE_GENERIC_EXECUTE
        GenMap.GenericAll = EAcessos.FILE_ALL_ACCESS
        ' Conversion of generic rights to specific file access rights
        MapGenericMask lDesiredAccess, GenMap
        ' Checking access
        AccessCheck SecDesc(1), hToken, lDesiredAccess, GenMap, PrivSet, Len(PrivSet), CheckFileAccess, r
        CloseHandle hToken
    End If
    ' Deleting Impersonation Token
    RevertToSelf

End Function

Private Function zIsNT() As Boolean
    Dim OSVer As OSVERSIONINFO
    OSVer.dwOSVersionInfoSize = Len(OSVer)
    GetVersionEx OSVer
    zIsNT = (OSVer.dwPlatformId = VER_PLATFORM_WIN32_NT)
End Function

