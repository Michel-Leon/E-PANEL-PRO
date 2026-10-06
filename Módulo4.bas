Attribute VB_Name = "M�dulo4"
Public Funcion As String, Celda_Cubiculo As String

Public Sub ObtenerHojaYFilaDestino(ByRef hojaDestino As Worksheet, ByRef filaDestino As Long)

    Const FILA_INICIO As Long = 56
    Const FILA_LIMITE As Long = 146
   
    Dim hojas(1 To 1) As Worksheet
    Set hojas(1) = ThisWorkbook.Sheets("FrontalFM")

    Dim i As Integer
    Dim filaLibre As Long

       For i = 1 To 1
        filaLibre = BuscarFilaLibre(hojas(i), FILA_INICIO)

        ' ¿Cabe en esta hoja?
        If filaLibre <= FILA_LIMITE Then
            Set hojaDestino = hojas(i)
            filaDestino = filaLibre
            Exit Sub
        End If
    Next i

    ' Si llegamos aquí, las 3 hojas están llenas
    MsgBox "La tabla de informe esta llena.", _
           vbCritical, "Espacio agotado"
    Set hojaDestino = Nothing
    filaDestino = 0

End Sub
' Función auxiliar: busca la primera fila libre en una hoja BOM
Private Function BuscarFilaLibre(hoja As Worksheet, filaInicio As Long) As Long
    Dim fila As Long
    fila = filaInicio

    Do While hoja.Range("FO" & fila).Text <> "" Or _
             hoja.Range("FT" & fila).Text <> "" Or _
             hoja.Range("FX" & fila).Text <> "" Or _
             hoja.Range("GC" & fila).Text <> "" Or _
             hoja.Range("GM" & fila).Text <> "" Or _
             hoja.Range("GW" & fila).Text <> "" Or _
             hoja.Range("HF" & fila).Text <> "" Or _
             hoja.Range("HO" & fila).Text <> "" Or _
             hoja.Range("HX" & fila).Text <> "" Or _
             hoja.Range("II" & fila).Text <> "" 
        fila = fila + 3
    Loop

    BuscarFilaLibre = fila
End Function

Sub Informe_Unidades(ByVal codigoBuscado As String, Optional ByVal escribirFrontal As Boolean = True)
    Dim hojaDestino As Worksheet, hojaDatos As Worksheet
    Dim celdaCodigo As Range
    Dim filaCodigo As Long, filaDestino As Long
    Dim valorN As String, valorC As String 

    If codigoBuscado = "" Then
        MsgBox "No se ha generado ningún código de Unidad.", vbExclamation
        Exit Sub
    End If

    Set hojaDatos = ThisWorkbook.Sheets("UNIDADES_FUNCIONALES")
    Set celdaCodigo = hojaDatos.Range("B:B").Find(What:=codigoBuscado, _
                        LookIn:=xlValues, LookAt:=xlWhole)

    If celdaCodigo Is Nothing Then
        MsgBox "Código '" & codigoBuscado & "' no encontrado en las hojas de datos.", vbExclamation
        Exit Sub
    End If
    filaCodigo = celdaCodigo.Row

    valorN = InputBox("Ingrese el Tag de la Columna" & vbCrLf & "(" & codigoBuscado & ")", "Tag columna")
    valorC = InputBox("Ingrese el Tag del Cubículo" & vbCrLf & "(" & codigoBuscado & ")", "Tag cubículo")

    ' Obtener hoja y fila destino
    Call ObtenerHojaYFilaDestino(hojaDestino, filaDestino)
    If hojaDestino Is Nothing Then Exit Sub

    With hojaDestino
        .Range("FO" & filaDestino).Value = valorN
        .Range("FT" & filaDestino).Value = valorC
        .Range("FX" & filaDestino).Value = hojaDatos.Cells(filaCodigo, "C").Value  ' Tipo
        .Range("GC" & filaDestino).Value = hojaDatos.Cells(filaCodigo, "D").Value  ' Breaker
        .Range("GM" & filaDestino).Value = hojaDatos.Cells(filaCodigo, "E").Value  ' Complemento
        .Range("GW" & filaDestino).Value = hojaDatos.Cells(filaCodigo, "F").Value  ' Posición
        .Range("HF" & filaDestino).Value = hojaDatos.Cells(filaCodigo, "G").Value  ' Instalación
        .Range("HO" & filaDestino).Value = hojaDatos.Cells(filaCodigo, "H").Value  ' Mando
        .Range("HX" & filaDestino).Value = Funcion
        .Range("II" & filaDestino).Value = hojaDatos.Cells(filaCodigo, "J").Value  ' Código APM
    End With

    ' Solo la unidad/transferencia principal escribe el +C en FrontalFM
    If escribirFrontal Then
        ThisWorkbook.Worksheets("FrontalFM").Range(Celda_Cubiculo).Value = valorC
    End If
End Sub
