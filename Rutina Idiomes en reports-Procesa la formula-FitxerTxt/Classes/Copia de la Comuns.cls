VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "oComuns"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Const MODUL As String = "oComuns"

Dim m_sFitxerConfigAlternatiu   As String   ' fitxer alternatiu de configuració, en comptes d'anar pel registre, ja que en el citrix, pot donar-se el cas
                                            ' de que no pugui obtenir accés al registre


' funcions idèntiques a la comuns

' ---------------------------------------------------------------------------
' Funció        : RecuperaIni
' Autor         : xavi
' Data          : 9/05/2001 19:59:27
' Observacions  :
' Revissions    : Dani 26/07/2001
'               Afegit en la comuns
' Parametres    :
' Finalitat     : Recupera un texte i si no el trobi hi guarda el de defecte d'aquesta manera és + fàcil
'               saber quin valor hi hem de posar
' ---------------------------------------------------------------------------
Public Function RecuperaIni(ByVal sCamp As String, ByVal sCategoria As String, ByVal sArxiu As String, Optional ByVal sDefecte As String) As String

    On Error GoTo Error
    Dim vValor As String
    
    With Me
    
        vValor = .RecuperaTexteINI(sCamp, sCategoria, sArxiu)
    
        ' si no existeix el valor el permetem gravar
        If vValor = vbNullString Then
            
            vValor = sDefecte
            If vValor <> vbNullString Then
            
                ' en aquest punt gravariem el parametre
                Call .GravaTexteINI(sCamp, sCategoria, sArxiu, vValor)
            End If
        End If
    End With
    
    RecuperaIni = vValor
    Exit Function
Error:
    
    Const NOM_PROCEDIMENT As String = "RecuperaIni"
    ShowError MODUL, NOM_PROCEDIMENT, Err.Number & " - " & Err.Description

End Function

Public Function RecuperaTexteINI(lpKeyName As String, lpApplicationName As String, lpFilename As String) As String
Dim lpReturnedString As String, nSize As Long, Resp As Long
Dim lpDefault   As String
    
    On Error GoTo Error
    lpReturnedString = Space(255)
    nSize = 255
    Resp = GetPrivateProfileString(lpApplicationName, lpKeyName, lpDefault, lpReturnedString, nSize, lpFilename)
    If Resp <> 0 Then
    
        RecuperaTexteINI = Left(lpReturnedString, Resp)
        
        ' 5.5.886 05/11/2005
        If RecuperaTexteINI = "." Then
        
            RecuperaTexteINI = ""
        End If
    Else
    
        RecuperaTexteINI = ""
    End If
    Exit Function
Error:
    ShowError MODUL, "RecuperaTexteIni", Err.Number & " - " & Err.Description
    
End Function

Public Function GravaTexteINI(lpKeyName As String, lpApplicationName As String, lpFilename As String, sValor As String) As Long
Dim lResp   As Long
    
    On Error GoTo Error
    lResp = WritePrivateProfileString(lpApplicationName, lpKeyName, sValor, lpFilename)
    GravaTexteINI = lResp
    Exit Function
Error:
    ShowError MODUL, "GravaTexteIni", Err.Number & " - " & Err.Description
    
End Function


Public Function UbicacioCompartida(Optional sAplicacio As String = "Gestium") As String
Const MODUL              As String = "SOFTWARE\DT Software\"

Const Aplicacio          As String = "Gestium"
Const SUB_CLAU           As String = "Ubicació Arxius"
Const DADES              As String = "Dades Compartides"
    
    On Error GoTo Error
Dim sUbicacio   As String

    ' 5.6.1150
    sUbicacio = RecuperaTexteINI(DADES, SUB_CLAU, FitxerConfigAlternatiu())
    If LenTrim(sUbicacio) = 0 Then
        
        sUbicacio = GetSetting(MODUL & sAplicacio & "\" & SUB_CLAU, DADES)
    End If

    UbicacioCompartida = sUbicacio
    Exit Function
Error:
    ShowError MODUL, "UbicacioCompartida", Err.Number & " - " & Err.Description
End Function

' Està també en el DT Inst
Public Function StrTranFirst(ByVal sCadena As String, ByVal sCaracterBuscar As String, ByVal sCaracterSubstituir As String) As String
Dim iPos    As Integer

    'iPos = InStr(sCadena, sCaracterBuscar)
    iPos = InStr(1, sCadena, sCaracterBuscar, vbTextCompare)
    If iPos > 0 Then
        
        StrTranFirst = Left(sCadena, iPos - 1) & _
                sCaracterSubstituir & _
                Mid(sCadena, iPos + Len(sCaracterBuscar))
    Else
        
        StrTranFirst = sCadena
    End If
End Function


' Transforma una cadena,
Public Function StrTran(sCadena As String, ByVal sCarBuscar As String, Optional ByVal sCarCanvi As String, Optional bTotal As Boolean = True, Optional bSensibleAMajuscula As Boolean = False) As String
Dim sCadenaAuxiliar     As String
Dim sValor              As String
    
    ' 5.6.1005 17/02/2007 Dani, ja que el lentrim treia si buscàvem per un espai
    If Len(sCarBuscar) > 0 Then
' 5.5.992 01/02/2007 Dani
'    If LenTrim(sCarBuscar) > 0 Then
        
        ' sCadena
        If Not bTotal Then
        
            ' només ho fa un sol cop, quan troba el caràcter ja para.
            StrTran = StrTranFirst(sCadena, sCarBuscar, sCarCanvi)
        Else
        
            Dim iPos As Integer
            sCadenaAuxiliar = sCadena
            iPos = -1
            While iPos <> 0
                
                iPos = InStr(1, sCadenaAuxiliar, sCarBuscar, vbTextCompare) ', Abs(Not bSensibleAMajuscula))
                If iPos <> 0 Then
                
                    sValor = sValor & Left(sCadenaAuxiliar, iPos - 1)
                    ' afegeixo el texte de canvi
                    sValor = sValor & sCarCanvi
                    
                    ' ara per la cadena que queda
                    sCadenaAuxiliar = Mid(sCadenaAuxiliar, iPos + Len(sCarBuscar))
                Else
                
                    StrTran = sValor & sCadenaAuxiliar
                End If
            Wend
        End If
    Else
        StrTran = sCadena
        
    End If
    Exit Function
End Function


Public Property Let FitxerConfigAlternatiu(sFitxerConfigAlternatiu As String)
    m_sFitxerConfigAlternatiu = sFitxerConfigAlternatiu
End Property

Public Property Get FitxerConfigAlternatiu() As String

    If LenTrim(m_sFitxerConfigAlternatiu) = 0 Then
    
        m_sFitxerConfigAlternatiu = "C:\Gestium.cfg\Gestium.ini"
    End If
    FitxerConfigAlternatiu = m_sFitxerConfigAlternatiu
End Property



' ------------------------------------------------------------------------------------------
' estan també a Globals.Bas de Gestium
' ------------------------------------------------------------------------------------------
Public Function LenTrim(sCadena As String) As Long
    On Error GoTo Error
    LenTrim = Len(sCadena)
    LenTrim = Len(Trim(sCadena))
    Exit Function
Error:
    
End Function

Public Function Ucasetrim(sCadena As String) As String
    On Error GoTo Error
    Ucasetrim = sCadena
    Ucasetrim = UCase(Trim(sCadena))
    Exit Function
Error:
    
End Function

Public Function LcaseTrim(sCadena As String) As String
    On Error GoTo Error
    LcaseTrim = sCadena
    LcaseTrim = LCase(Trim(sCadena))
    Exit Function
Error:
    
End Function


Public Function GetSetting(ByVal Section As String, ByVal Key As String, Optional ByVal Default As String = "") As String
Dim cObj    As New cRegistreWin

    On Error GoTo Error
    With cObj
    
        GetSetting = .GetSetting(Section, Key, Default)
    End With
    Exit Function
Error:
    ShowError MODUL, "GetSetting", Err.Number & " - " & Err.Description
End Function



Public Function Split(ByVal sCadena As String, ByVal sCaracter As String, Optional bNoUtilitzable As Boolean, Optional cColeccio As Collection) As Variant
On Error GoTo Error
Dim iPos        As Integer
Dim bSortir     As Boolean
Dim aValors()   As String
Dim I           As Integer
Dim sValor      As String
    
    ' 5.7.1442 13/11/2012
    Set cColeccio = New Collection
    I = 0
    Do While Not bSortir
    
        ' creem dimensió de l'arrai ...
        ReDim Preserve aValors(I)
        ' trobem la 1ª posició del caracter introduit com a clau
        iPos = InStr(sCadena, sCaracter)
        ' si hi ha un caracter com la clau
        If iPos > 0 Then
            
            sValor = Mid(sCadena, 1, iPos - 1)
            aValors(I) = sValor
            cColeccio.Add sValor
            ' refà la cadena de búsqueda ...
            sCadena = Mid(sCadena, iPos + Len(sCaracter))
            I = I + 1
        Else
        
            ' perque només hi ha un camp ...
            aValors(I) = sCadena
            cColeccio.Add sCadena
            bSortir = True
        End If
    Loop
    Split = aValors
Error:
I = 0
End Function

