Attribute VB_Name = "M�dulo3"
Public Funcion As String, Celda_Cubiculo As String
Public Sub ObtenerHojaYFilaDestino(ByRef hojaDestino As Worksheet, ByRef filaDestino As Long)

    Const FILA_INICIO As Long = 60
    Const FILA_LIMITE As Long = 104
   
    Dim hojas(1 To 1) As Worksheet
    Set hojas(1) = ThisWorkbook.Sheets("VistaFrontal")

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

    Do While hoja.Range("GG" & fila).Text <> "" Or _
             hoja.Range("GL" & fila).Text <> "" Or _
             hoja.Range("GP" & fila).Text <> "" Or _
             hoja.Range("GU" & fila).Text <> "" Or _
             hoja.Range("HE" & fila).Text <> "" Or _
             hoja.Range("HO" & fila).Text <> "" Or _
             hoja.Range("HX" & fila).Text <> "" Or _
             hoja.Range("IG" & fila).Text <> "" Or _
             hoja.Range("IP" & fila).Text <> "" 
        fila = fila + 4
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

    Set hojaDatos = ThisWorkbook.Sheets("Code_UF")
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
        .Range("GG" & filaDestino).Value = valorN
        .Range("GL" & filaDestino).Value = valorC
        .Range("GP" & filaDestino).Value = hojaDatos.Cells(filaCodigo, "C").Value  ' Tipo
        .Range("GU" & filaDestino).Value = hojaDatos.Cells(filaCodigo, "D").Value  ' Breaker
        .Range("HE" & filaDestino).Value = hojaDatos.Cells(filaCodigo, "E").Value  ' Complemento
        .Range("HO" & filaDestino).Value = hojaDatos.Cells(filaCodigo, "F").Value  ' Posición
        .Range("HX" & filaDestino).Value = hojaDatos.Cells(filaCodigo, "G").Value  ' Instalación
        .Range("IG" & filaDestino).Value = hojaDatos.Cells(filaCodigo, "H").Value  ' Mando
        .Range("IP" & filaDestino).Value = hojaDatos.Cells(filaCodigo, "J").Value  ' Código APM
    End With

    ' Solo la unidad/transferencia principal escribe el +C en VistaFrontal
    If escribirFrontal Then
        ThisWorkbook.Worksheets("VistaFrontal").Range(Celda_Cubiculo).Value = valorC
    End If
End Sub