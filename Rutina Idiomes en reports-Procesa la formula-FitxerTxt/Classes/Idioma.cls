VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cIdioma"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Dim m_dbds              As Object
Dim m_iIdioma           As Integer
Dim m_bBuscaIdiomaEnIni As Boolean
Dim m_oComuns           As Object
'utils
Dim m_oUtils                    As New oUtilsIni

Public Property Let oComuns(oDLL As Object)
    Set m_oComuns = oDLL
End Property

Public Property Let BaseDades(oBd As Object)
    Set m_dbds = oBd
End Property

Public Function ObteTraduccioReport(ByVal sTexte As String, ByRef sTraduccio As String, ByVal sReport As String) As Boolean
Const NOM_PROCEDIMENT As String = "ObteTraduccioReport"
On Error GoTo Error

Dim sSql            As String
Dim rs              As Object
Dim bPrimer         As Boolean
Dim oUtilsIdioma    As New cUtilsIdioma

    oUtilsIdioma.oComuns = m_oComuns

    bPrimer = True

    sSql = " Select Traduccio "
    sSql = sSql & " From DiccionariReports"
    sSql = sSql & " Where "
    sSql = sSql & " Texte"
    sSql = sSql & " = "
    sSql = sSql & "'"
    sSql = sSql & sTexte
    sSql = sSql & "'"
    sSql = sSql & " And "
    sSql = sSql & " ("
    sSql = sSql & " Report "
    sSql = sSql & " =  "
    sSql = sSql & "'"
    sSql = sSql & sReport
    sSql = sSql & "'"
    sSql = sSql & " OR "
    sSql = sSql & " Report "
    sSql = sSql & " = "
    sSql = sSql & "'"
    sSql = sSql & "'"
    sSql = sSql & " )"
    
    sSql = sSql & " And "
    sSql = sSql & " CodiIdioma = "
    
    sSql = sSql & m_iIdioma
    
    ' per si busquem el nom de l'informe
    sSql = sSql & " Order by "
    sSql = sSql & " Report DESC "
    
    Set rs = m_dbds.OpenRecordset(sSql, , dbReadOnly)
    With rs
    
        If Not (.EOF And .BOF) Then
        
            sTraduccio = !Traduccio & ""
            ObteTraduccioReport = True
        Else
        
            ' el crea com en blanc i com a report ?
            oUtilsIdioma.CreaDiccionariReports m_dbds, sTexte, sReport
            ' ara el crearà en la taula de la traducció
            oUtilsIdioma.CreaDiccionariReports m_dbds, sTexte, sReport, m_iIdioma

        End If
        
        .Close
    End With
    If Not ObteTraduccioReport Then
    
        If m_bBuscaIdiomaEnIni Then
        
            m_oUtils.oComuns = m_oComuns
        
            sTraduccio = m_oUtils.BuscaIdiomaInformeEnIni(88888, , , , sTexte)
            If Len(Trim(sTraduccio)) > 0 Then
            
                ObteTraduccioReport = True
            End If
        End If
    End If
    Set rs = Nothing
    Exit Function
Error:
    If Err.Number = 3078 And bPrimer Then
    
        bPrimer = False
        m_dbds.Execute "Create table DiccionariReports ( CodiIdioma integer, Texte varchar(150), Traduccio varchar(150) , Report varchar(255), EstatTraspas integer, DataUltimaModificacio datetime, Usuari varchar(15) null, Maquina varchar(15) null )"
        m_dbds.Execute "Create index Principal on DiccionariReports ( CodiIdioma, Texte , Report ) with primary"
        Resume
    End If

    ObteTraduccioReport = False
End Function


