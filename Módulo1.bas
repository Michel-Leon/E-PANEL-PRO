Attribute VB_Name = "M�dulo1"
'============== vARIABLES GLOBALES ===================
Public Linea As String
Public Celda As String
Public wsActual As Worksheet
Public ResultadoENV As String
'Public ResultadoENVP As String
'Public ResultadoENVL As String
'Public ResultadoENVPS As String
Public ResultadoDPS As String
public ResultadoSERV As String
Public ResultadoBRR As String
Public ResultadoUF As String
Public ResultadoCT As String
Public ResultadoEQ As String
Public ResultadoACC As String
Public ResultadoPC As String
Public ResultadoCOM As String
'============== Mover Bloque ===================
Public Sub MoverBloque(ByVal ref As String, ByVal BD As String, _
                       ByVal wsDestino As Worksheet, ByVal celdaDestino As String)

    Dim origen As Range, destino As Range, wsDatos As Worksheet

    On Error GoTo errHandler

    Set wsDatos = ThisWorkbook.Sheets(BD)
    Set origen = wsDatos.Range(ref)
    Set destino = wsDestino.Range(celdaDestino)

    origen.Copy Destination:=destino
    wsActual.Range(destino, destino.Offset(origen.Rows.Count - 1, origen.Columns.Count - 1)).Select
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