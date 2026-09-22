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

Sub Informe_Unidades(codigoBuscado As String)
    Dim hojaDestino As Worksheet
    Dim hojaDatos As Worksheet
    Dim hojasDatos As Variant
    Dim nombreHoja As Variant
    Dim celdaCodigo As Range
    Dim filaCodigo As Long
    Dim filaDestino As Long
    'VARIABLES QUE SE EXTRAEN DE LA TABLA----------------------------------------
    Dim Tipo As String, Breaker As String, Posicion As String, Instalacion As String
    Dim Mando As String , CodigoAPM As String, Complemento As String
    Dim encontrado As Boolean
    '-------------------------------------------------------------------
    'VARIABLES QUE SE Preguntan
    Dim valorN As String, valorC As String
    hojasDatos = Array("UNIDADES_FUNCIONALES")
    encontrado = False
    If codigoBuscado = "" Then
        MsgBox "No se ha generado ningún código de Unidad.", vbExclamation
        Exit Sub
    End If
    ' Buscar en cada hoja
    For Each nombreHoja In hojasDatos
        Set hojaDatos = ThisWorkbook.Sheets(nombreHoja)
        Set celdaCodigo = hojaDatos.Range("B:B").Find(What:=codigoBuscado, LookIn:=xlValues, LookAt:=xlWhole)
        If Not celdaCodigo Is Nothing Then
            filaCodigo = celdaCodigo.Row
            Tipo = hojaDatos.Cells(filaCodigo, "C").Value
            Breaker = hojaDatos.Cells(filaCodigo, "D").Value
            Complemento = hojaDatos.Cells(filaCodigo, "E").Value
            Posicion = hojaDatos.Cells(filaCodigo, "F").Value
            Instalacion = hojaDatos.Cells(filaCodigo, "G").Value
            Mando = hojaDatos.Cells(filaCodigo, "H").Value
            CodigoAPM = hojaDatos.Cells(filaCodigo, "J").Value
            encontrado = True
            Exit For
        End If
    Next nombreHoja
    If Not encontrado Then
        MsgBox "Código no encontrado en las hojas de datos.", vbExclamation
        Exit Sub
    End If
    valorN = InputBox("Ingrese el Tag de la Columna", "Tag columna")
    ValorC = InputBox("Ingrese el Tag de la Cubiculo", "Tag cubiculo")
    ' Obtener hoja y fila destino 
    Call ObtenerHojaYFilaDestino(hojaDestino, filaDestino)
    If hojaDestino Is Nothing Then Exit Sub

    With hojaDestino
        .Range("FO" & filaDestino).Value = ValorN
        .Range("FT" & filaDestino).Value = ValorC
        .Range("FX" & filaDestino).Value = Tipo
        .Range("GC" & filaDestino).Value = Breaker
        .Range("GM" & filaDestino).Value = Complemento
        .Range("GW" & filaDestino).Value = Posicion
        .Range("HF" & filaDestino).Value = Instalacion
        .Range("HO" & filaDestino).Value = Mando
        .Range("HX" & filaDestino).Value = Funcion
        .Range("II" & filaDestino).Value = CodigoAPM
    End With
    ThisWorkbook.Worksheets("FrontalFM").Range(Celda_Cubiculo).Value = ValorC  
    MsgBox "Referencia agregada al BOM correctamente en hoja '" & hojaDestino.Name & "'", vbInformation
End Sub

