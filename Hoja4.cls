'=================== VISTA PUERTA =================
Public Altura As String
Public Ancho As String, AnchoLISAS As String
Public Profundidad As String
Public IP As String, Color As String

'Accesorio MTM
    Private Sub TipoACC_Change()
        Dim dato As String
        dato = Me.OLEObjects("TipoACC").Object.Value
        With Me.OLEObjects("MedidaACC").Object
            .Clear
            Select Case dato
                Case "Techo_Frontal", "Hablador_Frontal"
                    .AddItem "200"
                    .AddItem "400"
                    .AddItem "720"
                Case "Techo Lateral"
                    .AddItem "400"
                    .AddItem "600"
                Case "Cancamo_Frontal", "Cancamo_Lateral", "Celosidas_Grande", "Celosidas_Pequeña", _
                        "Chapa_Antivandalica", "Hablador_Lateral", "Angulo_Frontal", "Angulo Lateral", _
                        "Cubierta_Frontal","Techo_Frontal", "Ventilador_Frontal","Ventilador_Superior"
                    .AddItem "NA"
                Case Else
                    .AddItem "No se reconoce" 
                End Select
        End With
    End Sub
'Placas de Control
    Private Sub VistaPC_Change()
        Dim dato1 As String, dato2 As String
        dato1 = Me.OLEObjects("VistaPC").Object.Value
        dato2 = Me.OLEObjects("TamanoPC").Object.Value
        With Me.OLEObjects("EquipoPC").Object
            .Clear
            Select Case dato2
                Case "Pequeña"
                    If dato1 = "Lateral" Then
                        .AddItem "Liso"
                    ElseIf dato1 = "Frontal" Then
                        .AddItem "Liso"
                        .AddItem "PM"
                        '.AddItem "ATS PREMIUM"
                        '.AddItem "ATS ESTANDAR"
                        '.AddItem "ATS BASICA"
                    End If
                Case "Grande"
                    If dato1 = "Lateral" Then
                        .AddItem "Liso"
                    ElseIf dato1 = "Frontal" Then
                        .AddItem "Liso"
                        .AddItem "ATS PREMIUM"
                        .AddItem "ATS ESTANDAR"
                        .AddItem "ATS BASICA"
                    End If
                Case "Mediana"
                    If dato1 = "Lateral" Then
                        .AddItem "Liso"
                    ElseIf dato1 = "Frontal" Then
                        .AddItem "Liso"
                        .AddItem "ATL600"
                        .AddItem "DCRL5"
                        .AddItem "Rapidus"
                    End If
            End Select
        End With
    End Sub
    Private Sub EquipoPC_Change()
        Dim dato As String, dato2 As String
        dato = Me.OLEObjects("EquipoPC").Object.Value
        With Me.OLEObjects("DisponibilidadPC").Object
            .Clear
            Select Case dato
                Case "Liso", "PM"
                    .AddItem "NA"
                Case "ATS PREMIUM", "ATS ESTANDAR", "ATS BASICA"
                    'IF dato2 = "Peque�a" then
                       ' .AddItem "Interna"
                   ' elseif dato2 = "Grande" then
                        .AddItem "Externa"
                   ' end if
                Case "DCRL5"
                    .AddItem "5 Pasos"
                Case "ATL600", "Rapidus"
                    .AddItem "6 Pasos"
                    .AddItem "7 Pasos"
                    .AddItem "8 Pasos"
                    .AddItem "9 Pasos"
                    .AddItem "10 Pasos"
                    .AddItem "11 Pasos"
                    .AddItem "12 Pasos"
                    .AddItem "13 Pasos"
                    .AddItem "14 Pasos"
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
                Case "Codigo_QR", "Riesgo_electrico1", "Ventilaci�n", "Descarga", "Riesgo_electrico2", _
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
    '============== PLACAS DE CONTROL ===================
        With Me.OLEObjects("TamanoPC").Object
            .Clear
            .AddItem "Pequeña"
            '.AddItem "Mediana"
            .AddItem "Grande"
        End With
        With Me.OLEObjects("VistaPC").Object
            .Clear
            .AddItem "Frontal"
            .AddItem "Lateral"
        End With
    '============== Accesorios MTM  =============================
        With Me.OLEObjects("TipoACC").Object
            .Clear
            .AddItem "Techo_Frontal"
            '.AddItem "Techo_Laterial"
            .AddItem "Cancamo_Frontal"
            '.AddItem "Cancamo_Lateral"
            .AddItem "Hablador_Frontal"
            '.AddItem "Hablador_Lateral"
            .AddItem "Angulo_Frontal"
            '.AddItem "Angulo_Lateral"
            .AddItem "Cubierta_Frontal"
            '.AddItem "Cubierta_Lateral"
            .AddItem "Ventilador_Frontal"
            .AddItem "Ventilador_Superior"
            '.AddItem "Ventilador_Lateral"
            .AddItem "Celosidas_Pequeña"
            .AddItem "Celosidas_Grande"
            .AddItem "Chapa_Antivandalica"
        End With   
    '============== COTAS  =============================
        With Me.OLEObjects("CotasDIBUJO").Object
            .Clear
            .AddItem "Ancho: 200mm"
            .AddItem "Ancho: 230mm"
            .AddItem "Ancho: 400mm"
            .AddItem "Ancho: 430mm"
            .AddItem "Ancho: 720mm"
            .AddItem "Ancho: 750mm"
            .AddItem "Ancho: 980mm"
            .AddItem "Ancho: 1180mm"
            .AddItem "Ancho: 1380mm"
            .AddItem "Altura: 1600mm"
            .AddItem "Altura: 1800mm"
            .AddItem "Altura: 2000mm"
            .AddItem "Division Tabelero"
            .AddItem "Corte Tren"
            .AddItem "Continuacion Tren"
        End With 
    '============== COLORES  =============================
        'Me.oleObjects("ColorPTD").Object.Value = Color
        With Me.OLEObjects("ColorConf").Object
            .Clear
            .AddItem ""
            .AddItem "RAL 7035"
            .AddItem "RAL 7032"
            .AddItem "RAL 9003"
            .AddItem "RAL 9004"
            .AddItem "Naranja"
            .AddItem "Amarillo"
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
    AnchoLISAS = CStr(Me.Range("A5").Value)

    LlenarCombo Me.OLEObjects("AlturaENV").Object, Altura
    LlenarCombo Me.OLEObjects("AnchoENV").Object, Ancho
    'msgbox MenuUnidades
    
End Sub
Public Sub Envolventes()
    '======================== 
        Celda = "KR83"
        ThisWorkbook.Sheets("FrontalPuerta").Range("KR83:NJ185").Clear
        Set wsActual = ThisWorkbook.Sheets("FrontalPuerta")
    '========================
    Dim dato1 As String, dato2 As String
    dato1 = Me.OLEObjects("AlturaENV").Object.Value
    dato2 = Me.OLEObjects("AnchoENV").Object.Value
    ResultadoENV = "EP" & dato1 & "X" & dato2
    MoverBloque ResultadoENV, "ENVOLVENTES", wsActual, Celda
End Sub
Public Sub ACESORIOS_MTM()
    '========================
        Celda = "KR104"
        ThisWorkbook.Sheets("FrontalPuerta").Range("KR83:NJ185").Clear
        Set wsActual = ThisWorkbook.Sheets("FrontalPuerta")
    '========================
    Dim dato1 As String, dato2 As String
    dato1 = Me.OLEObjects("TipoACC").Object.Value
    dato2 = Me.OLEObjects("MedidaACC").Object.Value
    ResultadoACM = "ACC_" & dato1 & "X" & dato2
    MoverBloque ResultadoACM, "ACCESOSRIOS MTM", wsActual, Celda
End Sub
Public Sub Placas_Control()
    '========================
        Celda = "KR132"
        ThisWorkbook.Sheets("FrontalPuerta").Range("KR83:NJ185").Clear
        Set wsActual = ThisWorkbook.Sheets("FrontalPuerta")
    '========================
    Dim dato1 As String, dato2 As String, dato3 As String, dato4 As String
    dato1 = Me.OLEObjects("TamanoPC").Object.Value
    Select Case dato1
        Case "Pequeña": dato1 = "PQ"
        Case "Grande": dato1 = "GD"
        Case "Mediana": dato1 = "MD"
    End Select
    dato2 = Me.OLEObjects("VistaPC").Object.Value
    Select Case dato2
        Case "Frontal": dato2 = "FR"
        Case "Lateral": dato2 = "LT"
    End Select
    dato3 = Me.OLEObjects("EquipoPC").Object.Value
    Select Case dato3
        Case "ATS PREMIUM": dato3 = "ATSPREMIUM"
        Case "ATS ESTANDAR": dato3 = "ATSESTANDAR"
        Case "ATS BASICA": dato3 = "ATSBASICA"
    End Select
    dato4 = Me.OLEObjects("DisponibilidadPC").Object.Value
    Select Case dato4
        Case "Externa":  dato4 = "EXT"
        Case "Interna":  dato4 = "INT"
        Case "5 Pasos":  dato4 = "5P"
        Case "6 Pasos":  dato4 = "6P"
        Case "7 Pasos":  dato4 = "7P"
        Case "8 Pasos":  dato4 = "8P"
        Case "9 Pasos":  dato4 = "9P"
        Case "10 Pasos": dato4 = "10P"
        Case "11 Pasos": dato4 = "11P"
        Case "12 Pasos": dato4 = "12P"
        Case "13 Pasos": dato4 = "13P"
        Case "14 Pasos": dato4 = "14P"
    End Select
    ResultadoPC = "PCTR_X" & dato1 & "X" & dato2 & "_" & dato3 & "X" & dato4
    MoverBloque ResultadoPC, "SERVICIOS", wsActual, Celda
End Sub
Public Sub COTAS()
    '========================
        Celda = "KR104"
        ThisWorkbook.Sheets("FrontalPuerta").Range("KR83:NJ185").Clear
        Set wsActual = ThisWorkbook.Sheets("FrontalPuerta")
    '========================
    Dim dato1 As String
    dato1 = Me.OLEObjects("CotasDIBUJO").Object.Value
    Select Case dato1
        Case "Ancho: 200mm": dato1 = "C_200"
        Case "Ancho: 230mm": dato1 = "C_230"
        Case "Ancho: 400mm": dato1 = "C_400"
        Case "Ancho: 430mm": dato1 = "C_430"
        Case "Ancho: 720mm": dato1 = "C_720"
        Case "Ancho: 750mm": dato1 = "C_750"
        Case "Ancho: 980mm": dato1 = "C_980"
        Case "Ancho: 1180mm": dato1 = "C_1180"
        Case "Ancho: 1380mm": dato1 = "C_1380"
        Case "Altura: 1600mm": dato1 = "C_1600"
        Case "Altura: 1800mm": dato1 = "C_1800"
        Case "Altura: 2000mm": dato1 = "C_2000"
        Case "Division Tabelero": dato1 = "Division"
        Case "Corte Tren": dato1 = "Corte_Tren"
        Case "Continuacion Tren": dato1 = "Continuidad_Tren"
    End Select
    ResultadoCOT = dato1
    MoverBloque ResultadoCOT, "COTAS", wsActual, Celda
End Sub
Public Sub Cambiodecolor() 'Vista frontal
    Dim ws As Worksheet
    Dim Celda As Range
    Dim valorColor As String
    Dim colorDestino As Long
    
    valorColor = Me.OLEObjects("ColorConf").Object.Value
    Set ws = ThisWorkbook.Sheets("FrontalPuerta")
    Select Case valorColor
        Case "RAL 7035": colorDestino = RGB(197, 199, 196)
        Case "RAL 7032": colorDestino = RGB(181, 176, 161)
        Case "RAL 9003": colorDestino = RGB(236, 236, 231)
        Case "RAL 9004": colorDestino = RGB(43, 43, 44)
        Case "Naranja": colorDestino = RGB(201, 133, 12)
        Case "Amarillo": colorDestino = RGB(235, 198, 21)
    End Select
    
    For Each Celda In ws.Range("CG50:FI146")
        Select Case valorColor
            Case "RAL 7035"
                If Celda.Interior.Color = RGB(181, 176, 161) Or _
                   Celda.Interior.Color = RGB(43, 43, 44) Or _
                   Celda.Interior.Color = RGB(201, 133, 12) Or _
                   Celda.Interior.Color = RGB(235, 198, 21) Or _
                   Celda.Interior.Color = RGB(181, 184, 177) Or _
                   Celda.Interior.Color = RGB(244, 244, 244) Or _
                   Celda.Interior.Color = RGB(236, 236, 231) Then
                    Celda.Interior.Color = RGB(197, 199, 196)
                End If
            Case "RAL 7032"
                If Celda.Interior.Color = RGB(197, 199, 196) Or _
                   Celda.Interior.Color = RGB(43, 43, 44) Or _
                   Celda.Interior.Color = RGB(201, 133, 12) Or _
                   Celda.Interior.Color = RGB(235, 198, 21) Or _
                   Celda.Interior.Color = RGB(244, 244, 244) Or _
                   Celda.Interior.Color = RGB(236, 236, 231) Then
                    Celda.Interior.Color = RGB(181, 176, 161)
                End If
            Case "RAL 9003"
                If Celda.Interior.Color = RGB(197, 199, 196) Or _
                   Celda.Interior.Color = RGB(181, 176, 161) Or _
                   Celda.Interior.Color = RGB(43, 43, 44) Or _
                   Celda.Interior.Color = RGB(201, 133, 12) Or _
                   Celda.Interior.Color = RGB(235, 198, 21) Or _
                   Celda.Interior.Color = RGB(244, 244, 244) Then
                    Celda.Interior.Color = RGB(236, 236, 231)
                End If
            Case "RAL 9004"
                If Celda.Interior.Color = RGB(197, 199, 196) Or _
                   Celda.Interior.Color = RGB(181, 176, 161) Or _
                   Celda.Interior.Color = RGB(201, 133, 12) Or _
                   Celda.Interior.Color = RGB(235, 198, 21) Or _
                   Celda.Interior.Color = RGB(244, 244, 244) Or _
                   Celda.Interior.Color = RGB(236, 236, 231) Then
                    Celda.Interior.Color = RGB(43, 43, 44)
                End If
            Case "Naranja"
                If Celda.Interior.Color = RGB(197, 199, 196) Or _
                   Celda.Interior.Color = RGB(181, 176, 161) Or _
                   Celda.Interior.Color = RGB(43, 43, 44) Or _
                   Celda.Interior.Color = RGB(235, 198, 21) Or _
                   Celda.Interior.Color = RGB(244, 244, 244) Or _
                   Celda.Interior.Color = RGB(236, 236, 231) Then
                    Celda.Interior.Color = RGB(201, 133, 12)
                End If
            Case "Amarillo"
                If Celda.Interior.Color = RGB(197, 199, 196) Or _
                   Celda.Interior.Color = RGB(181, 176, 161) Or _
                   Celda.Interior.Color = RGB(43, 43, 44) Or _
                   Celda.Interior.Color = RGB(201, 133, 12) Or _
                   Celda.Interior.Color = RGB(244, 244, 244) Or _
                   Celda.Interior.Color = RGB(236, 236, 231) Then
                    Celda.Interior.Color = RGB(235, 198, 21)
                End If
        End Select
    Next Celda
    
    MsgBox "Proceso Finalizado con Exito", vbInformation, "Cambio de Color"
End Sub
Public Sub ETIQUETAS()
    '========================
        Celda = "KP104"
        ThisWorkbook.Sheets("FrontalPuerta").Range("KR83:NJ185").Clear
        Set wsActual = ThisWorkbook.Sheets("FrontalPuerta")
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
            dato1 = "E_"
            ResultadoETQ = dato1 & dato3 & "X" & dato2
        Case "Linea de Producto"
            dato1 = "L_"
            ResultadoETQ = dato1 & dato3 & "X" & dato2
        Case "Logo de clientes": dato1 = "LC_"
            ResultadoETQ = dato1 & dato3
    End Select
    MoverBloque ResultadoETQ, "ETIQUETAS", wsActual, Celda
End Sub