VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cObteTradLitRpt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Function ObteTraduccioLiteralsReport(oComuns As Object, ByVal sTexteOriginal As String, ByRef sTraduccio As String, ByVal sIdioma As String, ByVal sArxiuTraduccio As String, Optional ByVal sReport As String) As Boolean
 
     
    On Error GoTo Error
    
    With oComuns
    
        ' substitueix els retorns de carro per cadenes en blanc
        
        sTraduccio = sTexteOriginal
        sTexteOriginal = .StrTran(sTexteOriginal, Chr(13), "")
        
        sTraduccio = .RecuperaIni(sTexteOriginal, sIdioma, sArxiuTraduccio, sTexteOriginal)
        
        
        ' a estudiar de si és millor que hi hagi un caràcter que 'emuli' el retorn de carro, com ara tenim
        ' en unes formules d'enviament de mail que permetem fer CRLF2 i són 2 retorns de carro
        ObteTraduccioLiteralsReport = True
    End With
    Exit Function
Error:
    Const NOM_PROCEDIMENT   As String = "ObteTraduccioLiteralsReport"
    'Err.Raise -28122812, MODUL, MODUL & "::" & NOM_PROCEDIMENT & vbCrLf & Err.Description

End Function

