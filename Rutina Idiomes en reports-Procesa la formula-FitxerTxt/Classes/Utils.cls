VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cUtils"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Sub EscriuLogFitxer(sDescripcio As String, Optional sFitxer As String, Optional ByVal sPath As String = "TEMP", Optional bDataHora As Boolean = True)
    On Error GoTo Error
    Dim iHandle As Integer
    Dim sCadena As String
        
    If sFitxer = "" Then
        
        sFitxer = sFitxer & "Error " & Format(Date, "yyyymmdd") & ".log"
    End If
    
    If UCase(sPath) = "TEMP" Then
    
        sPath = Environ("TEMP")
    End If
    ' comprova la barra
    If sPath <> "" Then
    
        If Right(sPath, 1) <> "\" Then
        
            sPath = sPath & "\"
        End If
        ' hi afegeixo el path
        sFitxer = sPath & sFitxer
    End If
    iHandle = FreeFile
    
    sCadena = sCadena & sDescripcio
    Open sFitxer For Append As iHandle
    If bDataHora Then
    
        sCadena = Now & " - " & sCadena
    End If
    Print #iHandle, sCadena
    Close #iHandle
    Exit Sub
Error:
End Sub

