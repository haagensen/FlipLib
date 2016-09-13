VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CInfosJanela"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private msCaption As String
Private msClasse As String
Private mlHandle As Long

Public Property Get Caption() As String
    Caption = msCaption
End Property
Public Property Let Caption(ByVal sCaption As String)
    msCaption = sCaption
End Property

Public Property Get Classe() As String
    Classe = msClasse
End Property
Public Property Let Classe(ByVal sClasse As String)
    msClasse = sClasse
End Property

Public Property Get Handle() As Long
    Handle = mlHandle
End Property
Public Property Let Handle(ByVal lHandle As Long)
    mlHandle = lHandle
End Property
