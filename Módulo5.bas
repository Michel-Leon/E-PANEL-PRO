Attribute VB_Name = "M�dulo5"
public N_INICIO AS LONG
Public N_Fin as long
Public Sub ObtenerHojaYFilaDestino2(ByRef hojaDestino As Worksheet, ByRef filaDestino As Long)

    Dim FILA_INICIO As Long
    Dim FILA_LIMITE As Long
    FILA_INICIO = N_INICIO
    FILA_LIMITE = N_Fin
   
    Dim hojas(1 To 1) As Worksheet
    Set hojas(1) = ThisWorkbook.Sheets("B_Conf")

    Dim i As Integer
    Dim filaLibre As Long

       For i = 1 To 1
        filaLibre = BuscarFilaLibre2(hojas(i), FILA_INICIO)

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
Private Function BuscarFilaLibre2(hoja As Worksheet, filaInicio As Long) As Long
    Dim fila As Long
    fila = filaInicio

    Do While hoja.Range("GB" & fila).Text <> "" Or _
             hoja.Range("GG" & fila).Text <> "" Or _
             hoja.Range("GK" & fila).Text <> "" Or _
             hoja.Range("GV" & fila).Text <> "" Or _
             hoja.Range("IP" & fila).Text <> "" 
             'hoja.Range("GW" & fila).Text <> "" Or _
             hoja.Range("HF" & fila).Text <> "" Or _
             hoja.Range("HO" & fila).Text <> "" Or _
             hoja.Range("HX" & fila).Text <> "" Or _
             hoja.Range("II" & fila).Text <> "" 
        fila = fila + 3
    Loop

    BuscarFilaLibre2 = fila
End Function

Sub Informe_Barras(codigoBuscado As String)
    Dim hojaDestino As Worksheet
    Dim hojaDatos As Worksheet
    Dim hojasDatos As Variant
    Dim nombreHoja As Variant
    Dim celdaCodigo As Range
    Dim filaCodigo As Long
    Dim filaDestino As Long
    'VARIABLES QUE SE EXTRAEN DE LA TABLA----------------------------------------
    Dim Corriente as String , Descripcion as string , Porcentaje as string 
    Dim CodigoAPM
    Dim encontrado As Boolean
    '-------------------------------------------------------------------
    'VARIABLES QUE SE Preguntan
    Dim valorN As String, valorC As String
    hojasDatos = Array("Codigos_Barras")
    encontrado = False
    If codigoBuscado = "" Then
        MsgBox "No se ha generado ningún código de Barras.", vbExclamation
        Exit Sub
    End If
    ' Buscar en cada hoja
    For Each nombreHoja In hojasDatos
        Set hojaDatos = ThisWorkbook.Sheets(nombreHoja)
        Set celdaCodigo = hojaDatos.Range("B:B").Find(What:=codigoBuscado, LookIn:=xlValues, LookAt:=xlWhole)
        If Not celdaCodigo Is Nothing Then
            filaCodigo = celdaCodigo.Row
            Corriente = hojaDatos.Cells(filaCodigo, "G").Value
            Descripcion = hojaDatos.Cells(filaCodigo, "M").Value
            Porcentaje = hojaDatos.Cells(filaCodigo, "H").Value
            CodigoAPM = hojaDatos.Cells(filaCodigo, "L").Value
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
    Call ObtenerHojaYFilaDestino2(hojaDestino, filaDestino)
    If hojaDestino Is Nothing Then Exit Sub

    With hojaDestino
        .Range("GB" & filaDestino).Value = ValorN
        .Range("GG" & filaDestino).Value = ValorC
        .Range("GK" & filaDestino).Value = Corriente
        .Range("GV" & filaDestino).Value = Descripcion
        .Range("IP" & filaDestino).Value = CodigoAPM
    End With

    MsgBox "Referencia agregada al BOM correctamente en hoja '" & hojaDestino.Name & "'", vbInformation
End Sub
Sub Informe_Barras_GBB(codigoBuscado As String)
    Dim hojaDestino As Worksheet
    Dim hojaDatos As Worksheet
    Dim hojasDatos As Variant
    Dim nombreHoja As Variant
    Dim celdaCodigo As Range
    Dim filaCodigo As Long
    Dim filaDestino As Long
    'VARIABLES QUE SE EXTRAEN DE LA TABLA----------------------------------------
    Dim Corriente as String , Descripcion as string , Porcentaje as string 
    Dim CodigoAPM
    Dim encontrado As Boolean
    '-------------------------------------------------------------------
    'VARIABLES QUE SE Preguntan
    Dim valorN As String, valorC As String
    hojasDatos = Array("Codigos_Barras")
    encontrado = False
    If codigoBuscado = "" Then
        MsgBox "No se ha generado ningún código de Barras.", vbExclamation
        Exit Sub
    End If
    ' Buscar en cada hoja
    For Each nombreHoja In hojasDatos
        Set hojaDatos = ThisWorkbook.Sheets(nombreHoja)
        Set celdaCodigo = hojaDatos.Range("B:B").Find(What:=codigoBuscado, LookIn:=xlValues, LookAt:=xlWhole)
        If Not celdaCodigo Is Nothing Then
            filaCodigo = celdaCodigo.Row
            Corriente = hojaDatos.Cells(filaCodigo, "G").Value
            Descripcion = hojaDatos.Cells(filaCodigo, "M").Value
            Porcentaje = hojaDatos.Cells(filaCodigo, "H").Value
            CodigoAPM = hojaDatos.Cells(filaCodigo, "L").Value
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
    Call ObtenerHojaYFilaDestino2(hojaDestino, filaDestino)
    If hojaDestino Is Nothing Then Exit Sub

    With hojaDestino
        .Range("GB" & filaDestino).Value = ValorN
        .Range("GG" & filaDestino).Value = ValorC
        .Range("GK" & filaDestino).Value = Corriente
        .Range("GV" & filaDestino).Value = Porcentaje
        .Range("HJ" & filaDestino).Value = Descripcion
        .Range("IP" & filaDestino).Value = CodigoAPM
    End With

    MsgBox "Referencia agregada al BOM correctamente en hoja '" & hojaDestino.Name & "'", vbInformation
End Sub

