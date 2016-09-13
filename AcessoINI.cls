VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FnINI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Declare Function WritePrivateProfileStringAPI Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Private Declare Function GetPrivateProfileStringAPI Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long

Public Function ReadINI(ByVal sINI As String, _
                        ByVal sSecao As String, _
                        ByVal sKey As String) As String

    Dim sRetVal As String, sRet As String, iWorked As Integer

    sRetVal = Space$(255)
    iWorked = GetPrivateProfileStringAPI(sSecao, sKey, "", sRetVal, Len(sRetVal), sINI)
    If iWorked <> 0 Then
        sRet = Trim$(Left$(sRetVal, iWorked))
    End If

    ReadINI = sRet

End Function

Public Sub WriteINI(ByVal sINI As String, _
                    ByVal sSecao As String, _
                    ByVal sKey As String, _
                    ByVal sValor As String)

    Call WritePrivateProfileStringAPI(sSecao, sKey, sValor, sINI)

End Sub
