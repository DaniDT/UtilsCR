VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Procesa"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Const MODUL As String = "Procesa"

Dim m_iIdioma           As Integer
Dim m_bBuscaIdiomaEnIni As Boolean
Dim m_oComuns           As oComuns
Dim m_dbds              As Object

Function IniciaLLibreria(ByVal sDatabase As String) As Boolean
    On Error GoTo Error
    Dim sArxiuConfig    As String
    
    ' copia local
    Set m_oComuns = New oComuns

    With m_oComuns
    
        sArxiuConfig = .UbicacioCompartida() & "\Gestium.ini"
        ' paràmetre per si no està la traducció en la base de dades, ho buscariem en un .ini
        m_bBuscaIdiomaEnIni = Val(.RecuperaIni("BuscaIdiomaEnIni", "Internacional", sArxiuConfig, "0")) = 1
        
        IniciaLLibreria = Not (m_dbds Is Nothing)
    End With
    Exit Function
Error:
End Function

Public Function Modul_Versio(sModul As String, Optional oApp As Object) As String
'    On Error GoTo Error -->> Se quita el control de errores xq limpia los errores.
    If oApp Is Nothing Then
    
        Set oApp = App
    End If
    With oApp
    
        Modul_Versio = "(" & .EXEName & ") " & sModul & String(1, "-") & .Major & "." & .Minor & "." & .Revision
    End With
    Exit Function
Error:
End Function

' per poder-la utilitzar des dels llistats, ha de ser per referència ?
Public Function ObteTraduccio(sTexteOriginal As String, sTraduccio As String, iIdioma As Integer, Optional sReport As String) As Boolean
 
    On Error GoTo Error
    Dim oIdioma As New cIdioma
    
    With m_oComuns
    
        ' substitueix els retorns de carro per cadenes en blanc, ja que en la traducció no ho posem
        sTexteOriginal = .StrTran(sTexteOriginal, Chr(13), "")
        
    
        If Not ObteTraduccio Then
        
            Select Case Trim(sTexteOriginal)
                Case Else
                    sTraduccio = "#" & sTexteOriginal
                    
            End Select
            
            ' ho posarà en un fitxer .ini per poder-ho traduir a mà
            Dim sFitxer As String

            sFitxer = .UbicacioCompartida() & "\Actualitzacions\Dades\Translat-" & iIdioma & ".cfg"
            Call .GravaTexteINI(sTexteOriginal, "Translat", sFitxer, ".")
            
        Else
        
            ' si hi ha traducció la posa al diccionari
            'AfegirParaules sTexteOriginal, sTraduccio
        End If
    End With
    Exit Function
Error:
    Const NOM_PROCEDIMENT   As String = "ObteTraduccio"
    Err.Raise -28122812, MODUL, MODUL & "::" & NOM_PROCEDIMENT & vbCrLf & Err.Description

End Function
