VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cObteTraduccioReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' és una copia del que hi ha a la comuns
Public Function ObteTraduccioReport(oComuns As Object, oIdioma As Object, iIdioma As Integer, bPreparaInternacional As Boolean, sTexte As String, sTraduccio As String, sReport As String) As Boolean
    On Error GoTo Error
    
    ' anirà a buscar en la llibreria de l'idioma
    On Error GoTo Error
    Const NOM_PROCEDIMENT As String = "ObteTraduccioReport"
    Dim bCorrecte   As Boolean
    
    bCorrecte = False
    ' com a mínim que tingui el mateix valor
    sTraduccio = sTexte
    If bPreparaInternacional Then
    
        bCorrecte = bPreparaInternacional
        ' si la llibreria no està inicialitzada s'inicialitza en aquest moment
        If oIdioma Is Nothing Then
                
            bCorrecte = InicialitzaIdioma(oComuns, iIdioma)
        End If
        If bCorrecte Then
        
            With oIdioma
                
                ' control de versio
                If .Versio() >= "1.9.52" Then
                    
                    bCorrecte = .ObteTraduccioReport(sTexte, sTraduccio, sReport)
                End If
            End With
        End If
    End If
    ObteTraduccioReport = bCorrecte
    Exit Function
Error:
'    g_oSelf.ShowError MODUL, NOM_PROCEDIMENT, Err.Number & " - " & Err.Description, "NI"
    
End Function


' ------------------------------------------------
' Funcio        : InicialitzaIdioma
' Observacions  : Pendent finalitzacio
' ------------------------------------------------
Public Function InicialitzaIdioma(oComuns As Object, Optional iIdioma As Integer = -999) As Boolean
    On Error GoTo Error
    Const NOM_PROCEDIMENT As String = "InicialitzaIdioma"
    
    Dim bCorrecte                               As Boolean
    Dim sLlibreria                              As String
    Dim dbInternacional                         As Object
    Dim sBaseDadesInternacionalitzacio          As String
    Dim sBaseDadesInternacionalitzacioGestium   As String
    
    sLlibreria = "dtInternacional.Idioma"
    
    Set m_oIdioma = CreateObject(sLlibreria)
    With oComuns
    
        ' obté la base de dades d'internacionalitzacio
        ' en el cas de no trobar-la l'hauria d'anar a buscar a les DADES_COMPARTIDES\Actualitzacions\Dades
        ' copiar-la en DADES_COMPARTIDES\Dades
        ' i registrar-la en el registre de Windows
        
        ' la del client
        sBaseDadesInternacionalitzacio = .IdiomaBaseDadesInternacionalitzacio()
        If LenTrim(sBaseDadesInternacionalitzacio) = 0 Then
        
            sBaseDadesInternacionalitzacio = .UbicacioCompartida() & "\Diccionari\Internacionalitzacio Gestium.mdb"
        End If
        If Not .ExisteixFitxer(sBaseDadesInternacionalitzacio) Then
            
            ' la de defecte de Gestium, que donem nosaltres
            sBaseDadesInternacionalitzacioGestium = .UbicacioCompartida() & "\Actualitzacions\Dades\Internacionalitzacio Gestium.mdb"
            If Not .ExisteixFitxer(sBaseDadesInternacionalitzacioGestium) Then
            
                .ShowError MODUL, NOM_PROCEDIMENT, "No existeix :" & sBaseDadesInternacionalitzacioGestium
                Exit Function
            Else
                ' si existeix la copia des d'actualitzacions a les dades del client ...
                .FileCopy sBaseDadesInternacionalitzacioGestium, sBaseDadesInternacionalitzacio
            End If
        End If
        ' per defecte serà el 0, en cas de no trobar-ne
'        iIdioma = 0 'Val(.RecuperaTexteINI("Idioma", "Parametres", gsArxiuConfig))
    
        If .Obrir(sBaseDadesInternacionalitzacio, dbInternacional) Then
            
            With m_oIdioma
            
                If .IniciaLlibreria(dbInternacional, oComuns) Then
                    
                    .Idioma = iIdioma
                    Debug.Print "Base de dades d'idioma inicialitzada " & sBaseDadesInternacionalitzacio
                    bCorrecte = True
                End If
            End With
        End If
    End With
    
    InicialitzaIdioma = bCorrecte
    Exit Function
Error:
    oComuns.ShowError MODUL, NOM_PROCEDIMENT, sLlibreria & " " & Err.Description
End Function

