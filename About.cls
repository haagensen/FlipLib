VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "About"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Function Executavel() As String

    Dim sNome As String
    Dim utl As FnUteis

    Set utl = New FnUteis
    sNome = utl.BackSlash(App.Path) & App.EXEName
    Set utl = Nothing

    If Dir$(sNome & ".EXE") <> "" Then
        sNome = sNome & ".EXE"
    ElseIf Dir$(sNome & ".DLL") <> "" Then
        sNome = sNome & ".DLL"
    ElseIf Dir$(sNome & ".OCX") <> "" Then
        sNome = sNome & ".OCX"
    End If
    Executavel = sNome

End Function

Public Function Versao() As String
    Versao = App.Major & "." & App.Minor & "." & App.Revision
End Function
