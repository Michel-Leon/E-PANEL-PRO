Attribute VB_Name = "M�dulo1"
'============== Mover Bloque ===================
Sub MoverBloque(ref As String, BD As String)
    Dim origen As Range
    Dim destino As Range
    Dim wsDatos As Worksheet
    Dim wsActual As Worksheet

    Set wsDatos = ThisWorkbook.Sheets(BD)
    Set wsActual = ActiveSheet

    On Error GoTo errHandler
    Set origen = wsDatos.Range(ref)
    Set destino = wsActual.Range("JV94")

    origen.Copy Destination:=destino

    wsActual.Range(destino, destino.Offset(origen.Rows.Count - 1, origen.Columns.Count - 1)).Select
    Exit Sub

    errHandler:
            MsgBox "Referencia no encontrada ", vbCritical
End Sub
'============== vARIABLES GLOBALES ===================
Public Linea As String
Public ResultadoENV As String
