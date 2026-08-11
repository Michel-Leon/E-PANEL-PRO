Attribute VB_Name = "M�dulo1"

'============== vARIABLES GLOBALES ===================
Public Linea As String
Public Celda As String
Public wsActual As Worksheet
Public ResultadoENV As String
Public ResultadoENVP As String
Public ResultadoENVL As String
Public ResultadoENVPS As String
Public ResultadoDPS As String
public ResultadoSERV As String
Public ResultadoBRR As String
Public ResultadoUF As String
Public ResultadoCT As String
Public ResultadoUC As String
'============== Mover Bloque ===================
Public Sub MoverBloque(ByVal ref As String, ByVal BD As String, _
                       ByVal wsDestino As Worksheet, ByVal celdaDestino As String)

    Dim origen As Range, destino As Range, wsDatos As Worksheet

    On Error GoTo errHandler

    Set wsDatos = ThisWorkbook.Sheets(BD)
    Set origen = wsDatos.Range(ref)
    Set destino = wsDestino.Range(celdaDestino)

    origen.Copy Destination:=destino
    Exit Sub

errHandler:
    MsgBox "Fallo al mover el bloque." & vbCrLf & _
           "Ref: " & ref & vbCrLf & _
           "BD: " & BD & vbCrLf & _
           "Destino: " & celdaDestino & vbCrLf & vbCrLf & _
           "Error " & Err.Number & ": " & Err.Description, vbCritical
End Sub
Private Function ResolverBloque(ByVal nombre As String) As Range
    Dim n As Name
    On Error Resume Next

    ' 1) Nombre global
    Set ResolverBloque = ThisWorkbook.Names(nombre).RefersToRange
    If Not ResolverBloque Is Nothing Then Exit Function
    Err.Clear

    ' 2) Nombre local en cualquier hoja
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        Set ResolverBloque = ws.Names(nombre).RefersToRange
        If Not ResolverBloque Is Nothing Then Exit Function
        Err.Clear
    Next ws

    On Error GoTo 0
End Function
'======================================================================================================

Sub Vista_puerta(ByVal resEnvP As String)

    WITH ThisWorkbook.Sheets("Vista Puerta").Range("BX44:IX144")
        .Clear
        .UnMerge
    End With
    Set wsActual = ThisWorkbook.Sheets("Vista Puerta")
    Select case thisWorkbook.Sheets("Vista FM").Range("A1").Value  
        Case "600":Celda = "BX97"
        Case "700":Celda = "BX89"
        Case "900":Celda ="BX73"
        Case "800":Celda = "BX81"
        Case "110":Celda = "BX57"
    End select 
    'MsgBox resEnvP
    MoverBloque resEnvP, "ENVOLVENTES_PUERTA", wsActual, celda
End Sub
sub Vista_lateral(ByVal resEnvL As String)
    With ThisWorkbook.Sheets("Vista Lateral").Range("CG49:IR145")
        .Clear
        .UnMerge
    End With
    Set wsActual = ThisWorkbook.Sheets("Vista Lateral")
    Select case thisWorkbook.Sheets("Vista FM").Range("A1").Value  
        Case "600":Celda = "CH98"
        Case "700":Celda = "CH89"
        Case "900":Celda ="CH74"
        Case "800":Celda = "CH82"
        Case "110":Celda = "CH58"
    End select    
    'MsgBox resEnvL
    MoverBloque resEnvL, "ENVOLVENTES L-Post", wsActual, celda
End Sub
Sub Vista_Posterior(ByVal resEnvPS As String)
    With ThisWorkbook.Sheets("Vista Posterior").Range("CH49:IX144")
        .Clear
        .UnMerge
    End With
    Set wsActual = ThisWorkbook.Sheets("Vista Posterior")
    Select case thisWorkbook.Sheets("Vista FM").Range("A1").Value  
        Case "600":Celda = "BX97"
        Case "700":Celda = "BX89"
        Case "900":Celda ="BX73"
        Case "800":Celda = "BX81"
        Case "110":Celda = "BX57"
    End select    
    'MsgBox resEnvPS
    MoverBloque resEnvPS, "ENVOLVENTES L-Post", wsActual, celda
End Sub

