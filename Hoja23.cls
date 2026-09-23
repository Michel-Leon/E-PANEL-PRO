'============== INTERFAZ BARRAJE ==========================
Public Altura As String
Public Ancho As String, ICC As String
' Prepara la hoja destino y limpia la zona (celdas + shapes)
    Private Sub PrepararHoja(ByRef ws As Worksheet, ByVal sZona As String)
        Set ws = ThisWorkbook.Sheets("B_Conf")
        Application.ScreenUpdating = False
        Application.CutCopyMode = False
        LimpiarZona ws, sZona
    End Sub
    ' Devuelve True si alguno de los valores esta vacio
    Private Function Vacio(ParamArray v() As Variant) As Boolean
        Dim i As Long
        For i = LBound(v) To UBound(v)
            If Trim$(CStr(v(i) & "")) = "" Then
                Vacio = True
                Exit Function
            End If
        Next i
    End Function
    Private Sub AvisoFaltan(ByVal sQue As String)
        Application.ScreenUpdating = True
        MsgBox "Faltan selecciones en: " & sQue, vbExclamation
    End Sub
   '--------------------------------------------------------------- UTILIDADES
    Private Function cbo(ByVal sNombre As String) As Object
        Set cbo = Me.OLEObjects(sNombre).Object
    End Function
    ' El "& """"" evita el error 94 (uso invalido de Null) con el combo vacio.
    Private Function Txt(ByVal sNombre As String) As String
        Txt = Trim$(cbo(sNombre).Value & "")
    End Function
' UNIDAD DE BARRAJE VERTICAL
    Private Sub CorrienteVBB_Change()
        Dim dato As String
        dato = Me.OLEObjects("CorrienteVBB").Object.Value
         With Me.OLEObjects("ICCVBB").Object
            .Clear
            Select Case dato
                Case "630"
                    If ICC = "Basic" Then
                        .AddItem "25"
                    Else
                        .AddItem "25"
                        .AddItem "65"
                    End If
                Case "800", "1250"
                    .AddItem "65"
            End Select
        End With
    End Sub
'UNIDAD DE BARRAJE HORIZONTAL
    Private Sub CorrienteHBB_Change()
        Dim dato As String
        dato = Me.OLEObjects("CorrienteHBB").Object.Value
         With Me.OLEObjects("ICCHBB").Object
            .Clear
            Select Case dato
                Case "400"
                   .AddItem "25"
                Case "630"
                    If ICC = "Basic" Then
                        .AddItem "25"
                    Else
                        .AddItem "40"
                        .AddItem "65"
                    End If
                Case "1250", "NS1250"
                    .AddItem "40"
                    .AddItem "65"
            End Select
        End With
    End Sub
'UNIDAD DE BARRAJE NEUTRO
    Private Sub CorrienteNBB_Change()
        Dim dato As String
        dato = Me.OLEObjects("CorrienteNBB").Object.Value
        With Me.OLEObjects("PorcentajeNBB").Object
            .Clear
            Select Case dato
                Case "630"
                    .AddItem "100"
                Case "1000", "1250"
                    .AddItem "50"
                    .AddItem "100"
            End Select
        End With
    End Sub
Public Sub Worksheet_Activate()
    Extraer_datos
    'Ocultar menus
        ActiveSheet.Shapes("Menu Vertical").Visible = False
        ActiveSheet.Shapes("Menu Horizontal").Visible = False
        ActiveSheet.Shapes("Menu Tierra").Visible = False
        ActiveSheet.Shapes("Menu Neutro").Visible = False
    '---------------------------- HBB
        With Me.OLEObjects("CorrienteHBB").Object
            .Clear
            Select Case ICC
                Case "Basic"
                    .AddItem "400"
                    .AddItem "630"
                Case "Compact"
                    .AddItem "630"
                    '.AddItem "800"
                    '.AddItem "1000"
                    .AddItem "1250"
                    .AddItem "NS1250"
            End Select
        End With
        With Me.OLEObjects("ICCHBB").Object
            .Clear
            Select Case ICC
                Case "Basic"
                    .AddItem "25"
                Case "Compact"
                    .AddItem "25"
                    .AddItem "40"
            End Select
        End With
    '---------------------------- VBB
        With Me.OLEObjects("FraccionVBB").Object
            .Clear
            Select Case ICC
                Case "Basic"
                    .AddItem "1/2"
                    .AddItem "3/4"
                Case "Compact"
                    .AddItem "1/2"
                    .AddItem "3/4"
                    .AddItem "1"
            End Select
        End With
        With Me.OLEObjects("CorrienteVBB").Object
            .Clear
            Select Case ICC
                Case "Basic"
                    .AddItem "400"
                    .AddItem "630"
                Case "Compact"
                    .AddItem "630"
                    .AddItem "800"
                    '.AddItem "1000"
                    .AddItem "1250"
            End Select
        End With
    '---------------------------- GBB
         With Me.OLEObjects("ICCGBB").Object
            .Clear
            Select Case ICC
                Case "Basic"
                    .AddItem "25"
                Case "Compact"
                    .AddItem "25"
                    .AddItem "40"
            End Select
        End With
    '---------------------------- NBB
        With Me.OLEObjects("CorrienteNBB").Object
            .Clear
            Select Case ICC
                Case "Basic"
                    '.AddItem "400"
                    .AddItem "630"
                Case "Compact"
                    .AddItem "630"
                    '.AddItem "800"
                    .AddItem "1000"
                    .AddItem "1250"
            End Select
        End With
        
End Sub
'--- Llena un ComboBox ActiveX a partir de una cadena separada por comas
Private Sub LlenarCombo(cbo As Object, ByVal datos As String)
    Dim partes() As String
    Dim i As Long
    Dim item As String

    cbo.Clear

    If Len(Trim$(datos)) = 0 Then Exit Sub

    partes = Split(datos, ",")
    For i = LBound(partes) To UBound(partes)
        item = Trim$(partes(i))
        If Len(item) > 0 Then cbo.AddItem item
    Next i

    If cbo.ListCount = 1 Then cbo.ListIndex = 0   ' si hay uno solo, lo selecciona
End Sub
Public Sub Extraer_datos()
    ' EXTRAER DATOS DE DIMENSIONES
    Altura = CStr(Me.Range("A1").Value)
    Ancho = CStr(Me.Range("A4").Value)
    ICC = CStr(Me.Range("A3").Value)
    'msgbox MenuUnidades
    LlenarCombo Me.OLEObjects("AnchoHBB").Object, Ancho
    LlenarCombo Me.OLEObjects("AnchoGBB").Object, Ancho
    LlenarCombo Me.OLEObjects("AnchoNBB").Object, Ancho
    LlenarCombo Me.OLEObjects("AlturaVBB").Object, Altura
End Sub
Private Sub LimpiarZona(ws As Worksheet, ByVal sRango As String)
    Dim rngZona As Range, shp As Shape, i As Long

    Set rngZona = ws.Range(sRango)

    Application.CutCopyMode = False          'clave: sale del modo copiar

    'borrar de atras hacia adelante porque la coleccion se reindexa
    For i = ws.Shapes.Count To 1 Step -1
        Set shp = ws.Shapes(i)
        If Not Intersect(shp.TopLeftCell, rngZona) Is Nothing Then
            'no borrar los combos ActiveX si estuvieran en la zona
            If shp.Type <> msoOLEControlObject Then shp.Delete
        End If
    Next i

    rngZona.Clear
End Sub
'---------------------------------------------------------------
Sub Traercolumna()
    ' Copia el rango A1:C10 y lo pega en E1
    Sheets("FrontalFM").Range("BV48:FJ155").Copy
    Sheets("B_Conf").Range("BV48").PasteSpecial xlPasteAll
    
    ' Limpia el portapapeles
    Application.CutCopyMode = False
End Sub
'========================================= HBB
    Public Sub Barraje_Horizontal()
        Dim wsActual As Worksheet, Celda As String
        Dim dato1 As String, dato2 As String, dato3 As String

        PrepararHoja wsActual, "KO125:NJ185"
        Celda = "KS105"

        dato1 = Txt("AnchoHBB")
        dato2 = Txt("CorrienteHBB")
        dato3 = Txt("ICCHBB")
       

        If Vacio(dato1, dato2, dato3) Then AvisoFaltan "Corriente, Ancho o ICC": Exit Sub

        ResultadoHBB = "HBB" & "X" & dato1 & "_" & dato2 & "X" & dato3
        MoverBloque ResultadoHBB, "TC-TB", wsActual, Celda
        Application.ScreenUpdating = True
    End Sub
'========================================= VBB
    Public Sub Barraje_Vertical()
        Dim wsActual As Worksheet, Celda As String
        Dim dato1 As String, dato2 As String, dato3 As String
        Dim dato4 As String
        PrepararHoja wsActual, "KO125:NJ185"
        Celda = "KS105"

        dato1 = Txt("AlturaVBB")
        dato2 = Txt("CorrienteVBB")
        dato3 = Txt("ICCVBB")
        dato4 = Txt("FraccionVBB")
        Select Case dato4
            Case "1/2": dato4 = "1of2"
            Case "3/4": dato4 = "3of4"
            Case "1":   dato4 = "1"
        End Select
       
        If Vacio(dato1, dato2, dato3, dato4) Then AvisoFaltan "Corriente, Altura, Fraccion o ICC": Exit Sub
        ResultadoVBB = "VBB" & "X" & dato4 & "_" & dato2 & "X" & dato3 & "X" & dato1
        MoverBloque ResultadoVBB, "TC-TB", wsActual, Celda
        Application.ScreenUpdating = True
    End Sub
'========================================= GBB
    Public Sub Barraje_Tierra()
        Dim wsActual As Worksheet, Celda As String
        Dim dato1 As String, dato2 As String

        PrepararHoja wsActual, "KO125:NJ185"
        Celda = "KS105"

        dato1 = Txt("AnchoGBB")
        dato2 = Txt("ICCGBB")

        If Vacio(dato1, dato2) Then AvisoFaltan "Ancho o ICC": Exit Sub

        ResultadoGBB = "GBB" & "X" & dato1 & "X" & dato2
        MoverBloque ResultadoGBB, "TC-TB", wsActual, Celda
        Application.ScreenUpdating = True
    End Sub
'========================================= NBB
    Public Sub Barraje_Neutro()
        Dim wsActual As Worksheet, Celda As String
        Dim dato1 As String, dato2 As String, dato3 As String

        PrepararHoja wsActual, "KO125:NJ185"
        Celda = "KS105"

        dato1 = Txt("AnchoNBB")
        dato2 = Txt("CorrienteNBB")
        dato3 = Txt("PorcentajeNBB")

        If Vacio(dato1, dato2) Then AvisoFaltan "Ancho o ICC": Exit Sub

        ResultadoNBB = "NBB" & "X" & dato1 & "_" & dato3 & "X" & dato2
        MoverBloque ResultadoNBB, "TC-TB", wsActual, Celda
        Application.ScreenUpdating = True
    End Sub
'======================================================= BUSCAR DATOS
    Public Sub Bucar_datos_HBB()
        N_INICIO = 68
        N_Fin = 71
        Celda_cubiculo = "KS105"
        Informe_Barras ResultadoHBB
    End Sub
    Public Sub Buscar_datosVBB()
        N_INICIO = 56
        N_Fin = 59
        Celda_cubiculo = "KS105"
        Informe_Barras ResultadoVBB
    End Sub
    Public Sub Buscar_datosGBB()
        N_INICIO = 80
        N_Fin = 86
        Celda_cubiculo = "KS105"
        Informe_Barras_GBB ResultadoGBB
    End Sub
    Public Sub Buscar_datosNBB()
        N_INICIO = 80
        N_Fin = 86
        Celda_cubiculo = ""
        Informe_Barras_GBB ResultadoNBB
    End Sub


