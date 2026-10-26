'DISENO POSTERIOR

Public Altura As String,  tipo As String
Public Ancho As String, AnchoLISAS As String
Public Profundidad As String
Public IP As String, Color As String

'ACCESORIOS MTM
    Private Sub TipoACC_Change()
        Dim dato As String
        dato = Me.OLEObjects("TipoACC").Object.Value
        With Me.OLEObjects("MedidaACC").Object
            .Clear
            Select Case dato
                Case "Fijacion_muro"
                    '.AddItem "Frontal"
                    '.AddItem "Lateral"
                    .AddItem "Posterior"
                Case "Fijacion_poste"
                    '.AddItem "Lateral"
                    .AddItem "Posterior"
                Case "Ventilador"
                    .AddItem "Frontal"
                    '.AddItem  "Lateral"
            End Select
        End With
    End Sub
'ETIQUETAS
    Private Sub ClaseETQ_Change()
        Clase_etiqueta
        Dim dato As String
        dato = Me.OLEObjects("ClaseETQ").Object.Value
        With Me.OLEObjects("TipoETQ").Object
            .Clear
            Select Case dato
                Case "Etiqueta"
                    .AddItem "Codigo_QR"
                    .AddItem "Placa_GR"
                    .AddItem "Riesgo_electrico1"
                    .AddItem "Ventilación"
                    .AddItem "Descarga"
                    .AddItem "Riesgo_electrico2"
                    .AddItem "Advertencia_choque_electrico"
                    .AddItem "Riesgo_electricoV1"
                    .AddItem "Riesgo_electricoV2"
                    .AddItem "Riesgo_electricoV3"
                    .AddItem "Choque_electrico_1"
                    .AddItem "Choque_electrico_2"
                    .AddItem "Prohibido_paso"
                    .AddItem "Prohibido_acceso"
                    .AddItem "Respaldo"
                    .AddItem "Peligro_advertencia"
                    .AddItem "Voltaje_peligroso"
                    .AddItem "Voltaje_208"
                    .AddItem "Voltaje_220"
                    .AddItem "Voltaje_440"
                    .AddItem "Voltaje_460"
                    .AddItem "Voltaje_480"
                    .AddItem "Aviso_acometidas"
                    .AddItem "Programador_horario_MedidaACC"
                    .AddItem "Equipo_medida"
                    .AddItem "Ajuste_higrostato"
                    .AddItem "Inspeccion_DPS"
                    .AddItem "No_manipular_Guaya"
                    .AddItem "Ajuste_corriente"
                    .AddItem "Ajuste_tension"
                    .AddItem "Mantenimiento_filtros"
                    .AddItem "Senales_extrernas_transferencia"
                    .AddItem "Parametrizacion_medidor"
                    .AddItem "Control_produccion"
                    .AddItem "Incumplimiento_instrucciones"
                    .AddItem "Mantenimiento"
                    .AddItem "Cableado"
                    .AddItem "Superficie_caliente"
                    .AddItem "Linea"
                    .AddItem "Generador"
                    .AddItem "Red"
                    .AddItem "Transformador"
                    .AddItem "UPS"
                    .AddItem "Blindobarra"
                    .AddItem "Blindobarra_Transformador"
                    .AddItem "Blindobarra_Generador"
                Case "Linea de Producto"
                    .AddItem "B_Rack"
                    .AddItem "C_Rack"
                    .AddItem "C_Box"
                    .AddItem "D_Critical"
                    .AddItem "E_Center"
                    .AddItem "G_Flex"
                    .AddItem "I_MCC"
                    .AddItem "L_Meter"
                    .AddItem "M_Board"
                    .AddItem "P_Ats"
                    .AddItem "Q_Bank"
                    .AddItem "S_Energy"
                    .AddItem "T_Meter"
                    .AddItem "T_Power"
                    .AddItem "V_Drive"
                    .AddItem "X_Proof"
                Case "Logo de clientes"
                    .AddItem "Schneider_Electric"
                    .AddItem "Ecopetrol"
                    .AddItem "SierraCol"
                    .AddItem "Geopark"
                    .AddItem "Tryenergy"
                    .AddItem "Copower"
                    .AddItem "Terpel"
                    .AddItem "La_Fazenda"
                    .AddItem "Canaveral"
                    .AddItem "Parque_Arauco"
                    .AddItem "Parque_Caracoli"
                    .AddItem "Babaria"
                    .AddItem "Higuera_Escalante"
                    .AddItem "ESSA"
                    .AddItem "Postobon"
                    .AddItem "Lycon"
                    .AddItem "Alfrio"
                    .AddItem "ANI"
                    .AddItem "Cool"
                    .AddItem "Delta"
                    .AddItem "Devimar"
                    .AddItem "Energymed"
                    .AddItem "Gecolsa"
                    .AddItem "JE_Jaimes"
                    .AddItem "PQS_Energy"
                    .AddItem "Ruta_Cacao"
                    .AddItem "Telefonica"
                    .AddItem "UTS"
                    .AddItem "UIS"
                    .AddItem "Friocol"
                    .AddItem "2Kinse"
                    .AddItem "EnerteK"
                    .AddItem "Hacer_Ingenieria"
                    .AddItem "MSS_Ingenieria"
                    .AddItem "Kluger"
                    .AddItem "Bimbo"
            End Select
        End With
    End Sub
    Public Sub Clase_etiqueta()
        Dim dato As String
        dato = Me.OLEObjects("ClaseETQ").Object.Value
        Select Case dato
            Case "Etiqueta"
                ActiveSheet.Shapes("Dato3titulo").TextFrame2.TextRange.Text = "Tama�o:"
                Me.OLEObjects("Dato3ETQ").Object.Enabled = True
            Case "Linea de Producto"
                ActiveSheet.Shapes("Dato3titulo").TextFrame2.TextRange.Text = "Color:"
                Me.OLEObjects("Dato3ETQ").Object.Enabled = True
            Case "Logo de clientes"
                ActiveSheet.Shapes("Dato3titulo").TextFrame2.TextRange.Text = "......"
                Me.OLEObjects("Dato3ETQ").Object.Enabled = False
        End Select
    End Sub
    Private Sub TipoETQ_Change()
        Dim dato As String, dato2 As String
        dato2 = Me.OLEObjects("ClaseETQ").Object.Value
        dato = Me.OLEObjects("TipoETQ").Object.Value
        With Me.OLEObjects("Dato3ETQ").Object
            .Clear
            Select Case dato
                Case "Codigo_QR", "Riesgo_electrico1", "Ventilación", "Descarga", "Riesgo_electrico2", _
                     "Voltaje_peligroso", "Voltaje_208", "Voltaje_220", "Voltaje_440", _
                     "Voltaje_460", "Voltaje_480", "Aviso_acometidas", "Programador_horario", "Equipo_medida", _
                     "Ajuste_higrostato", "Inspeccion_DPS", "No_manipular_Guaya", "Ajuste_corriente", "Mantenimiento_filtros", _
                     "Senales_extrernas_transferencia", "Parametrizacion_medidor", "Control_produccion", "Incumplimiento_instrucciones", _
                     "Mantenimiento", "Cableado", "Superficie caliente"
                    .AddItem "Pequeno"
                Case "Ventilación", "Descarga", "Riesgo_electrico2"
                    .AddItem "Pequeno"
                    .AddItem "Mediano"
                    .AddItem "Grande"
                Case "Placa_GR", "Ventilación", "Descarga", "Riesgo_electrico2", "Riesgo_electricoV1", "Riesgo_electricoV2", _
                     "Riesgo_electricoV3", "Choque_electrico_1", "Choque_electrico_2", "Prohibido_paso", "Prohibido_acceso", "Respaldo", _
                     "Generador", "Red", "Transformador", "UPS"
                    .AddItem "Mediano"
                Case "Advertencia_choque_electrico"
                    .AddItem "Grande"
                Case "Voltaje_peligroso"
                    .AddItem "Pequeno"
                    .AddItem "Mediano"
                Case "Blindobarra", "Blindobarra_Generador", "Blindobarra_Transformador"
                    .AddItem "NA"
            End Select
            Select Case dato2
                Case "Linea de Producto"
                    .AddItem "Gris"
                    .AddItem "Azul"
            End Select
        End With
    End Sub
Public Sub Worksheet_Activate()
    Extraer_datos
    '------- Ocultar menus ---------
        ActiveSheet.Shapes("Menu EnvolventesGFL").Visible = False
        ActiveSheet.Shapes("Menu ColorGFL").Visible = False
        ActiveSheet.Shapes("Menu AcotarGFL").Visible = False
        ActiveSheet.Shapes("Menu EtiquetaGFL").Visible = False
        ActiveSheet.Shapes("Menu AccesoriosMTMGFL").Visible = False
    '============== ACC MTM ============================
        With Me.OLEObjects("TipoACC").Object
            .Clear
            .AddItem "Fijacion_muro"
            .AddItem "Fijacion_poste"
            .AddItem "Ventilador"
        End With
    '============== COTAS  =============================
        With Me.OLEObjects("CotasDIBUJO").Object
            .Clear
            .AddItem "Ancho: 520mm"
            .AddItem "Ancho: 620mm"
            .AddItem "Ancho: 720mm"
            .AddItem "Ancho: 1120mm"
            '.AddItem "Profundidad: 300mm" 
            '.AddItem "Profundidad: 400mm"
            .AddItem "Altura: 900mm"
            .AddItem "Altura: 800mm"
            .AddItem "Altura: 700mm" 
            .AddItem "Division Tabelero"
            .AddItem "Corte Tren"
            .AddItem "Continuacion Tren"
        End With
    '============== COLORES  =============================
        With Me.OLEObjects("ColorConf").Object
            .Clear
            .AddItem ""
            .AddItem "RAL 7035"
            .AddItem "RAL 7032"
            .AddItem "RAL 9003"
            .AddItem "RAL 9004"
            .AddItem "Inoxidable"
            .AddItem "Naranja"
            .AddItem "Amarillo"
        End With
    '============== ETIQUETA  =============================
        With Me.OLEObjects("ClaseETQ").Object
            .Clear
            .AddItem "Etiqueta"
            .AddItem "Linea de Producto"
            .AddItem "Logo de clientes"
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
    Ancho = CStr(Me.Range("A2").Value)
    IP = CStr(Me.Range("A4").Value)
    tipo = CStr(Me.Range("A6").Value)

    LlenarCombo Me.OLEObjects("AlturaENV").Object, Altura
    LlenarCombo Me.OLEObjects("AnchoENV").Object, Ancho
    'LlenarCombo Me.OLEObjects("AnchoLSS").Object, AnchoLISAS
    'msgbox MenuUnidades
End Sub
Public Sub Envolventes()
    '========================
        celda = "KP94"
        ThisWorkbook.Sheets("Posterior").Range("KO91:OJ185").Clear
        Set wsActual = ThisWorkbook.Sheets("Posterior")
    '========================
    Dim dato1 As String, dato2 As String, dato3 As String
    dato1 = Me.OLEObjects("AlturaENV").Object.Value
    dato2 = Me.OLEObjects("AnchoENV").Object.Value

    ResultadoENV = "EPT_" & tipo & "X" & dato1 & "X" & dato2
    MoverBloque ResultadoENV, "ENVOLVENTES L-Post", wsActual, celda
End Sub
Public Sub ACESORIOS_MTM()
    '========================
        celda = "KR113"
        ThisWorkbook.Sheets("Posterior").Range("KO91:OJ185").Clear
        Set wsActual = ThisWorkbook.Sheets("Posterior")
    '========================
    Dim dato1 As String, dato2 As String
    dato1 = Me.OLEObjects("TipoACC").Object.Value
    dato2 = Me.OLEObjects("MedidaACC").Object.Value
    ResultadoACM = "ACC_" & dato1 & "_" & dato2
    MoverBloque ResultadoACM, "accesorios MTM", wsActual, celda
End Sub
Public Sub COTAS()
    '========================
        celda = "KR116"
        ThisWorkbook.Sheets("Posterior").Range("KO91:OJ185").Clear
        Set wsActual = ThisWorkbook.Sheets("Posterior")
    '========================
    Dim dato1 As String
    dato1 = Me.OLEObjects("CotasDIBUJO").Object.Value
    Select Case dato1
        Case "Ancho: 520mm": dato1 = "CG_520"
        Case "Ancho: 620mm": dato1 = "CG_620"
        Case "Ancho: 720mm": dato1 = "CG_720"
        Case "Ancho: 1120mm": dato1 = "CG_1120"
        case "Profundidad: 300mm": dato1 = "CG_300"
        case "Profundidad: 400mm": dato1 = "CG_400"
        Case "Altura: 900mm": dato1 = "CG_900"
        Case "Altura: 800mm": dato1 = "CG_800"
        Case "Altura: 700mm": dato1 = "CG_700"
        Case "Division Tabelero": dato1 = "Division"
        Case "Corte Tren": dato1 = "Corte_Tren"
        Case "Continuacion Tren": dato1 = "Continuidad_Tren"
    End Select
    ResultadoCOT = dato1
    MoverBloque ResultadoCOT, "COTAS", wsActual, celda
End Sub
Public Sub Cambiodecolor() 'Vista frontal
    Dim ws As Worksheet
    Dim celda As Range
    Dim valorColor As String
    Dim colorDestino As Long
    
    ' Obtener el valor del color desde la portada
    valorColor = Me.OLEObjects("ColorConf").Object.Value
    Set ws = ThisWorkbook.Sheets("Posterior")
    ' Definir color destino seg�n el valor
    Select Case valorColor
        Case "RAL 7035": colorDestino = RGB(197, 199, 196)
        Case "RAL 7032": colorDestino = RGB(181, 176, 161)
        Case "RAL 9003": colorDestino = RGB(236, 236, 231)
        Case "RAL 9004": colorDestino = RGB(43, 43, 44)
        case "Inoxidable": colorDestino = RGB(143;143;143)
        Case "Naranja": colorDestino = RGB(201, 133, 12)
        Case "Amarillo": colorDestino = RGB(235, 198, 21)
    End Select
    
    ' Recorrer el rango y cambiar solo los colores espec�ficos
    For Each celda In ws.Range("BZ57:KI155")
        Select Case valorColor
            Case "RAL 7035"
                If celda.Interior.Color = RGB(181, 176, 161) Or _
                   celda.Interior.Color = RGB(43, 43, 44) Or _
                   celda.Interior.Color = RGB(201, 133, 12) Or _
                   celda.Interior.Color = RGB(235, 198, 21) Or _
                   celda.Interior.Color = RGB(181, 184, 177) Or _
                   celda.Interior.Color = RGB(244, 244, 244) Or _
                   celda.Interios.ColOr = RGB(143; 143; 143) Or _
                   celda.Interior.Color = RGB(236, 236, 231) Then
                    celda.Interior.Color = RGB(197, 199, 196)
                End If
            Case "RAL 7032"
                If celda.Interior.Color = RGB(197, 199, 196) Or _
                   celda.Interior.Color = RGB(43, 43, 44) Or _
                   celda.Interior.Color = RGB(201, 133, 12) Or _
                   celda.Interior.Color = RGB(235, 198, 21) Or _
                   celda.Interior.Color = RGB(244, 244, 244) Or _
                   celda.Interios.ColOr = RGB(143; 143; 143) Or _
                   celda.Interior.Color = RGB(236, 236, 231) Then
                    celda.Interior.Color = RGB(181, 176, 161)
                End If
            Case "RAL 9003"
                If celda.Interior.Color = RGB(197, 199, 196) Or _
                   celda.Interior.Color = RGB(181, 176, 161) Or _
                   celda.Interior.Color = RGB(43, 43, 44) Or _
                   celda.Interior.Color = RGB(201, 133, 12) Or _
                   celda.Interior.Color = RGB(235, 198, 21) Or _
                   celda.Interios.ColOr = RGB(143; 143; 143) Or _
                   celda.Interior.Color = RGB(244, 244, 244) Then
                    celda.Interior.Color = RGB(236, 236, 231)
                End If
            Case "Inoxidable"
                If celda.Interior.Color = RGB(197, 199, 196) Or _
                   celda.Interior.Color = RGB(181, 176, 161) Or _
                   celda.Interior.Color = RGB(43, 43, 44) Or _
                   celda.Interior.Color = RGB(201, 133, 12) Or _
                   celda.Interior.Color = RGB(235, 198, 21) Or _
                   celda.Interior.Color = RGB(244, 244, 244) Or _
                   celda.Interior.Color = RGB(236, 236, 231) Then
                    celda.Interior.Color = RGB(143; 143; 143)
                End If
            Case "RAL 9004"
                If celda.Interior.Color = RGB(197, 199, 196) Or _
                   celda.Interior.Color = RGB(181, 176, 161) Or _
                   celda.Interior.Color = RGB(201, 133, 12) Or _
                   celda.Interior.Color = RGB(235, 198, 21) Or _
                   celda.Interior.Color = RGB(244, 244, 244) Or _
                   celda.Interios.ColOr = RGB(143; 143; 143) Or _
                   celda.Interior.Color = RGB(236, 236, 231) Then
                    celda.Interior.Color = RGB(43, 43, 44)
                End If
            Case "Naranja"
                If celda.Interior.Color = RGB(197, 199, 196) Or _
                   celda.Interior.Color = RGB(181, 176, 161) Or _
                   celda.Interior.Color = RGB(43, 43, 44) Or _
                   celda.Interior.Color = RGB(235, 198, 21) Or _
                   celda.Interior.Color = RGB(244, 244, 244) Or _
                   celda.Interios.ColOr = RGB(143; 143; 143) Or _
                   celda.Interior.Color = RGB(236, 236, 231) Then
                    celda.Interior.Color = RGB(201, 133, 12)
                End If
            Case "Amarillo"
                If celda.Interior.Color = RGB(197, 199, 196) Or _
                   celda.Interior.Color = RGB(181, 176, 161) Or _
                   celda.Interior.Color = RGB(43, 43, 44) Or _
                   celda.Interior.Color = RGB(201, 133, 12) Or _
                   celda.Interior.Color = RGB(244, 244, 244) Or _
                   celda.Interios.ColOr = RGB(143; 143; 143) Or _
                   celda.Interior.Color = RGB(236, 236, 231) Then
                    celda.Interior.Color = RGB(235, 198, 21)
                End If
        End Select
    Next celda
    
    MsgBox "Proceso Finalizado con Exito", vbInformation, "Cambio de Color"
End Sub
Public Sub ETIQUETAS()
    '========================
        celda = "KR116"
        ThisWorkbook.Sheets("Posterior").Range("KO91:OJ185").Clear
        Set wsActual = ThisWorkbook.Sheets("Posterior")
    '========================
    Dim dato1 As String, dato2 As String, dato3 As String
    dato2 = Me.OLEObjects("Dato3ETQ").Object.Value
    Select Case dato2
        Case "Pequeno": dato2 = "1"
        Case "Mediano": dato2 = "2"
        Case "Grande": dato2 = "3"
    End Select
    dato1 = Me.OLEObjects("ClaseETQ").Object.Value
    dato3 = Me.OLEObjects("TipoETQ").Object.Value
    Select Case dato1
        Case "Etiqueta"
            dato1 = "GFE_"
            ResultadoETQ = dato1 & dato3 & "X" & dato2

        Case "Linea de Producto"
            dato1 = "GFL_"
            ResultadoETQ = dato1 & dato3 & "X" & dato2

        Case "Logo de clientes": dato1 = "GFLC_"
            ResultadoETQ = dato1 & dato3
    End Select
    MoverBloque ResultadoETQ, "ETIQUETAS", wsActual, celda
End Sub