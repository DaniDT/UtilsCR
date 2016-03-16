VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Gestium"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Const MODUL                 As String = "CRUFLDTTxt::Gestium"

Const dbReadOnly            As Integer = 4

Dim m_oComuns               As oComuns
Dim m_sArxiuTraduccio       As String
Dim m_iComptador            As Long
Dim m_oUtils                As New cUtils
Dim m_oParaules             As Collection
Dim m_iVersioLocal          As Integer

' per a Crystal Reports és necessari que es digui la DLL CRUFLxx, i ho mostrarà com
' DT + nom de la classe +'.' + nom de la funció pública.
' en aquest cas DTGestium.Principal

' <REVISAR> si els valors han de ser passats per referència

Public Function Principal(sParametre As String, sIdioma As String, Optional sReport As String) As String
    On Error GoTo Error
    Dim sTexte  As String
    Dim sFitxerIni  As String
    
    Set m_oComuns = New oComuns
    Principal = sParametre
    
    m_iVersioLocal = 0
    
    'Exit Function
    ' no puc instanciar la comuns des de Gestium, ja que sino perdo la connexió de l'actual comuns que està a memoria
    
    ' 1.1.13 16/07/2012 comprovació de si és de Gestium
    If m_iVersioLocal = 0 Then
    
        If InStr(UCase(sReport), UCase("Gestium® Software")) > 0 Then
        
            m_iVersioLocal = 1
        ElseIf InStr(UCase(sReport), UCase("nom del llistat.rpt")) = 0 Then

            m_iVersioLocal = 1
        End If
    End If
    
    With m_oComuns
    
        sFitxerIni = .UbicacioCompartida() & "\Gestium.ini"
        sFitxerIni = .RecuperaIni("Fitxer", "InternacionalImpresos", sFitxerIni, "Dades\IdiomaImpresos.txt")
        m_sArxiuTraduccio = .UbicacioCompartida() & "\" & sFitxerIni
        
        m_iComptador = m_iComptador + 1
        If sParametre = "@#@TEST@#@" Then Principal = "Test - " & m_iComptador & " " & Now
        
        If ObteTraduccioLiteralsReport(m_oComuns, sParametre, sTexte, sIdioma, m_sArxiuTraduccio, sReport) Then
            
            If m_oComuns.LenTrim(sTexte) = 0 Then
            
                sTexte = sParametre
            End If
        End If
    End With
    Principal = sTexte
    Exit Function
Error:
    If m_iComptador < 1000 Then
    VBA.MsgBox Err.Description & " " & m_iComptador
    End If
    
    'Principal = "#Error# " & Err.Number & " " & Err.Source
End Function

Private Sub Class_Initialize()
    On Error Resume Next
    m_oUtils.EscriuLogFitxer MODUL & "-Class_Initialize"
    Set m_oParaules = Nothing
    Set m_oParaules = New Collection

End Sub

Private Sub Class_Terminate()
    On Error Resume Next
    m_oUtils.EscriuLogFitxer MODUL & "-Class_Terminate"
    Set m_oComuns = Nothing
    
End Sub

Private Function ObteTraduccioLiteralsReport(oComuns As Object, ByVal sTexteOriginal As String, ByRef sTraduccio As String, ByVal sIdioma As String, ByVal sArxiuTraduccio As String, Optional ByVal sReport As String) As Boolean
 
    On Error GoTo Error
    Dim c As New cObteTradLitRpt
    
    With c
    
        ObteTraduccioLiteralsReport = .ObteTraduccioLiteralsReport(oComuns, sTexteOriginal, sTraduccio, sIdioma, sArxiuTraduccio, sReport)
    End With
    Exit Function
Error:

End Function

' per tenir en una col.lecció i que vagi més ràpid
Private Function AfegirParaules(sTexteOriginal As String, sTraduccio As String)
    On Error GoTo Error:
    
    m_oParaules.Add sTexteOriginal & "@#@" & sTraduccio, sTraduccio
    Exit Function
Error:
End Function

Public Function ParaulaOriginal(sTraduccio As String) As String
    On Error GoTo Error
    Dim sTexte  As String
    Dim aValors As Variant
    
    ParaulaOriginal = sTraduccio
    sTexte = m_oParaules.Item(sTraduccio)
    
    ' s'ha de fer un split per separar el texte de la traducció
    aValors = m_oComuns.Split(sTexte, "@#@")
    ParaulaOriginal = CStr(aValors(0))
    Exit Function
Error:
    Const NOM_PROCEDIMENT   As String = "ParaulaOriginal"
    Err.Raise -28122812, MODUL, MODUL & "::" & NOM_PROCEDIMENT & vbCrLf & Err.Description

End Function
