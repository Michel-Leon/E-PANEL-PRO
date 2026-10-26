' DISENO G-FLEX
Public Altura As String
Public Ancho As String, AnchoLISAS As String
Public Profundidad As String
Public IP As String, Color As String
Private mbCargando As Boolean          'evita eventos en cascada
Private mBloq As Boolean
Private Const O720 As String = "Horizontal,H.Compuesta,V.Compacta"   
'ENVOLVENTES
    'Private Sub AlturaENV_change()
        'Dim dato1 As String
        'dato1 = Me.OLEObjects("AlturaENV").Object.Value
        'With Me.OLEObjects("AnchoENV").Object
         '   .Clear
         '   Select Case dato1
          '      Case "600": .AddItem "400" 'neo
          '      Case "700": .AddItem "520" 'basic
          '      Case "900": .AddItem "720" 'neo, basis
          '      Case "800": .AddItem "620" 'basic
          '      Case "1100" ' neo
          '      .AddItem "720"
          '      .AddItem "400+720"
          '  End Select
       ' End With
   ' End Sub
'ACCESORIOS MTM
    Private Sub TipoACC_Change()
        Dim dato As String
        dato = Me.OLEObjects("TipoACC").Object.Value
        With Me.OLEObjects("MedidaACC").Object
            .Clear
            Select Case dato
                Case "Fijacion_muro"
                    .AddItem "Frontal"
                    '.AddItem "Lateral"
                    '.AddItem "Posterior"
                Case "Fijacion_poste"
                    .AddItem "Lateral"
                    .AddItem "Posterior"
                Case "Ventilador"
                    .AddItem "Frontal"
                    '.AddItem "Lateral"
            End Select
        End With
    End Sub
'UNIDADES FUNCIONALES
   
    ' 1) FRAME ? llena Medida y Corriente
    Private Sub FrameUF_Change()
        If mBloq Then Exit Sub
        Dim r As String, p As Variant
        Select Case Me.OLEObjects("FrameUF").Object.Text
            Case "MiniBreaker", "MiniBreaker_ONKA":          r = "400,720|1-63A"
            Case "GoPact100":                                r = "400,720|100A,125A,150A"
            Case "GoPact150":                                r = "400|100A,125A,150A"
            Case "GoPact200":                                r = "400,720|200A"
            Case "GoPact250", "EZC250", "CVS250", "NSX250":  r = "400,720|250A"
            Case "GoPact400":                                r = "720|400A"
            Case "EZC100":                                   r = "400,720|100A"
            Case "CVS100", "NSX100":                         r = "400|100A"
            Case Else:                                       r = "|"
        End Select
        p = Split(r, "|")
        Llenar "MedidaUF", CStr(p(0))
        Llenar "CorienteUF", CStr(p(1))
        MedidaUF_Change
    End Sub
    ' 2) MEDIDA ? llena Orientaci�n (400 = Vertical, 720 = m�s opciones, salvo excepciones)
    Private Sub MedidaUF_Change()
        If mBloq Then Exit Sub
        Dim m As String, o As String
        m = Me.OLEObjects("MedidaUF").Object.Text
        Select Case Me.OLEObjects("FrameUF").Object.Text & "|" & m
            Case "MiniBreaker|720", "MiniBreaker|400", "MiniBreaker_ONKA|400", _
             "MiniBreaker_ONKA|720", "GoPact100|400", "EZC100|400", _
             "GoPact150|400", "GoPact200|400", "GoPact250|400", _
             "EZC250|400", "CVS250|400", "CVS100|400": o = "Vertical"
        '_________________________________________________________________________________________________________________________
            Case "GoPact100|720", "GoPact200|720", "GoPact250|720", _
             "GoPact400|720", "EZC100|720", "EZC250|720", "CVS250|720", _
              "NSX250|720": o = "Horizontal,H.Compuesta"
        '_________________________________________________________________________________________________________________________
            Case "GoPact100|720", "GoPact250|720", "EZC100|720", "EZC250|720", _
             "CVS250|720", "NSX100|400", "NSX250|720", "NSX250|400": o = "V.Compacta"

            Case Else: o = IIf(m = "400", "Vertical", IIf(m = "720", O720, ""))
        End Select
        Llenar "OrientacionUF", o
        OrientacionUF_Change
    End Sub
    ' 3) ORIENTACI�N ? llena Funci�n
    Private Sub OrientacionUF_Change()
        If mBloq Then Exit Sub
        Select Case Me.OLEObjects("OrientacionUF").Object.Text
            Case "":            Llenar "FuncionUF", ""
            Case "H.Compuesta": Llenar "FuncionUF", "Totalizador"
            Case Else:          Llenar "FuncionUF", "Carga"
        End Select
    End Sub
    ' Llena el combo, autoselecciona si hay 1 opci�n y conserva la anterior si sigue v�lida
    Private Sub Llenar(nom As String, lista As String)
        Dim cb As Object, v As Variant, ant As String
        mBloq = True
        Set cb = Me.OLEObjects(nom).Object
        ant = cb.Text
        cb.Clear
        If lista <> "" Then For Each v In Split(lista, ","): cb.AddItem v: Next
        On Error Resume Next
        If cb.ListCount = 1 Then cb.ListIndex = 0 Else cb.Value = ant
        If cb.ListIndex = -1 Then cb.Value = ""
        On Error GoTo 0
        mBloq = False
    End Sub
'UNIDADES LISAS
    Private Sub AnchoLSS_Change()
        Dim dato As String
        dato = Me.OLEObjects("AnchoLSS").Object.Value
        With Me.OLEObjects("UnidadLSS").Object
            .Clear
            Select Case dato
                Case "400"
                    If Me.OLEObjects("TipoLSS").Object.Value = "SFI - Frontal Interno" Then
                        .AddItem "1u"
                        .AddItem "2u"
                        .AddItem "3u"
                        .AddItem "4u"
                        .AddItem "5u"
                        .AddItem "6u"
                        .AddItem "7u"
                        .AddItem "8u"
                    Else
                        .AddItem "5u"
                        .AddItem "6u"
                        .AddItem "7u"
                        .AddItem "8u"
                    End If

                Case "720"
                     If Me.OLEObjects("TipoLSS").Object.Value = "SFI - Frontal Interno" Then
                        .AddItem "1u"
                        .AddItem "2u"
                        .AddItem "3u"
                        .AddItem "4u"
                        .AddItem "5u"
                        .AddItem "6u"
                        .AddItem "7u"
                        .AddItem "8u"
                        .AddItem "9u"
                        .AddItem "10u"
                    Else
                        .AddItem "5u"
                        .AddItem "6u"
                        .AddItem "7u"
                        .AddItem "8u"
                        .AddItem "9u"
                        .AddItem "10u"
                        .AddItem "11u"
                        .AddItem "12u"
                    End If
                End Select
            End With
    End Sub
'UNIDADES BARRAJE
    Private Sub OrientacionBRR_Change()
        Dim dato As String
        dato = Me.OLEObjects("OrientacionBRR").Object.Text
        With Me.OLEObjects("AmperajeBRR").Object
            .Clear
            Select Case dato
                Case "Horizontal Bus Bar"
                    .AddItem "400A"
                Case "Horizontal Tierra Neutro"
                    .AddItem "NA"
                Case "Vertical Bus Bar"
                    .AddItem "NSX400"
                    .AddItem "EZC400"
                    .AddItem "CVS400"
            End Select
        End With
        With Me.OLEObjects("AnchoBRR").Object
            .Clear
            Select Case dato
                Case "Horizontal Bus Bar"
                    .AddItem "720"
                Case "Horizontal Tierra Neutro"
                    .AddItem "720"
                    .AddItem "400"
                Case "Vertical Bus Bar"
                    .AddItem "400"
            End Select
        End With
    End Sub
'UNIDADES DPS
    Private Sub FabricanteDPS_Change()
        Dim dato As String
        dato = Me.OLEObjects("FabricanteDPS").Object.Value
        With Me.OLEObjects("EquipoDPS").Object
            .Clear
            Select Case dato
                Case "Ciprotec", "Citel", "Phoenix Contac", "Schneider Electric"
                    .AddItem "MiniBreaker"
                    .AddItem "EZC100"
                    .AddItem "CVS100"
                    .AddItem "NSX100"
                Case "HWA"
                    .AddItem "EZC100"
                    .AddItem "CVS100"
                    .AddItem "NSX100"
            End Select
        End With
        With Me.OLEObjects("AnchoDPS").Object
            .Clear
            Select Case dato
                Case "Ciprotec", "Citel", "Phoenix Contac", "Schneider Electric"
                    .AddItem "400"
                    .AddItem "720"
                Case "HWA"
                    .AddItem "720"
            End Select
        End With
    End Sub
'Accesorios de control
    Private Sub ElementoCTR_Change()
        Dim dato As String
        dato = Me.OLEObjects("ElementoCTR").Object.Value
        With Me.OLEObjects("vistaACCCTR").Object
            .Clear
            Select Case dato
                Case "ClavijaRoja", "ClavijaAzul"
                    .AddItem "Frontal"
                    .AddItem "Derecha_Abierta"
                    .AddItem "Derecha_Cerrada"
                    .AddItem "Izquierda_Abierta"
                    .AddItem "Izquierda_Cerrada"
                Case Else
                    .AddItem "NA"
            End Select
        End With
    End Sub
'TRANFERENCIAS
    Private Sub OrientacionTRF_Change()
        Dim dato As String
        dato = Me.OLEObjects("OrientacionTRF").Object.Value
        With Me.OLEObjects("EquipoTRF").Object
            .Clear
            Select Case dato
                Case "Vertical"
                    .AddItem "EZC250"
                    .AddItem "EZC630"
                    .AddItem "CVS250"
                    .AddItem "CVS630"
                    .AddItem "NSX250"
                    .AddItem "NSX630"
                    .AddItem "GoPact250"
                    .AddItem "GoPact400"
                    .AddItem "NS800"
                Case "Horizontal"
                    .AddItem "EZC250X2"
                    .AddItem "CVS250X2"
                    .AddItem "NSX250X2"
                    .AddItem "GoPact250X2"
            End Select
        End With
    End Sub
    Private Sub EquipoTRF_Change()
        Dim dato As String
        dato = Me.OLEObjects("EquipoTRF").Object.Value
        With Me.OLEObjects("MandoTFF").Object
            .Clear
            Select Case dato
                Case "CVS250", "NSX250", "NSX400", "CVS400"
                    .AddItem "Manual"
                    .AddItem "Motorizado"
                Case "EZC250", "EZC400", "GoPact250", "GoPact400"
                    .AddItem "Manual"
            End Select
        End With
    End Sub
    'Conmutadoras
    Private Sub CorrienteCOM_Change()
        Dim dato As String
        dato = Me.OLEObjects("CorrienteCOM").Object.Value
        With Me.OLEObjects("EquipoCOMT").Object
            .Clear
            Select Case dato
                Case "25A"
                    .AddItem "MiniBreaker"
                Case "63A"
                    .AddItem "MiniBreaker"
                    .AddItem "EZC100"
                Case "125A"
                    .AddItem "CVS100"
                    .AddItem "EZC100"
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
'ARRANQUE MOTOR
    Private Sub EquipoAM_Change()
        Dim dato As String
        dato = Me.OLEObjects("EquipoAM").Object.Value
        With Me.OLEObjects("TallaAM").Object
            .Clear
            Select Case dato
                Case "Contactores"
                    .AddItem "Talla 1-2"
                    .AddItem "Talla 3"
                Case "Guardamotor", "GuardamotorXContactores"
                    .AddItem "Vertical"
            End Select
        End With
    End Sub
Public Sub Worksheet_Activate()
    Extraer_datos
    '---OCULTAR MENUS ------
        ActiveSheet.Shapes("Menu EnvolventesGFL").Visible = False
        ActiveSheet.Shapes("Menu UnidadFuncionalGFL").Visible = False
        ActiveSheet.Shapes("Menu LisasGFL").Visible = False
        ActiveSheet.Shapes("Menu DPSGFL").Visible = False
        ActiveSheet.Shapes("Menu Equipo CTRGFL").Visible = False
        ActiveSheet.Shapes("Menu ACCESORIOS CTRGFL").Visible = False
        ActiveSheet.Shapes("Menu ETOGFL").Visible = False
        ActiveSheet.Shapes("Menu CONMUTADORESGFL").Visible = False
        ActiveSheet.Shapes("Menu AccesoriosMTMGFL").Visible = False
        ActiveSheet.Shapes("Menu BarrajeGFL").Visible = False
        ActiveSheet.Shapes("Menu TransferenciasGFL").Visible = False
        ActiveSheet.Shapes("Menu ServiciosGFL").Visible = False
        ActiveSheet.Shapes("Menu ColorGFL").Visible = False
        ActiveSheet.Shapes("Menu AcotarGFL").Visible = False
        ActiveSheet.Shapes("Menu EtiquetaGFL").Visible = False
        ActiveSheet.Shapes("Menu ArranqueM GF").Visible = False

    '============== ACC MTM ============================
        With Me.OLEObjects("TipoACC").Object
            .Clear
            .AddItem "Fijacion_muro"
            '.AddItem "Fijacion_poste"
            .AddItem "Ventilador"
        End With
    '=============== UNIDADES FUNCIONALES ===============
        Me.OLEObjects("TipoUF").Object.Value = "SFI - Frontal Interno"
        With Me.OLEObjects("FrameUF").Object
            .Clear
            .AddItem "MiniBreaker"
            .AddItem "MiniBreaker_ONKA"
            .AddItem "GoPact100"
            .AddItem "GoPact150"
            .AddItem "GoPact200"
            .AddItem "GoPact250"
            .AddItem "GoPact400"
            .AddItem "EZC100"
            .AddItem "EZC250"
            .AddItem "CVS100"
            .AddItem "CVS250"
            .AddItem "NSX100"
            .AddItem "NSX250"
        End With
        With Me.OLEObjects("MandoUF").Object
            .Clear
            .AddItem "Toggle"
        End With
    '=============== UNIDADES LISAS =====================
        With Me.OLEObjects("TipoLSS").Object
            .AddItem "SFI - Frontal Interno"
            .AddItem "SIC - Placa de Control"
        End With
        With Me.OLEObjects("AnchoLSS").Object
            .Clear
            .AddItem "400"
            .AddItem "720"
        End With
    '=============== UNIDAD DE BARRAJE =============
        With Me.OLEObjects("OrientacionBRR").Object
            .Clear
            .AddItem "Horizontal Bus Bar"
            .AddItem "Horizontal Tierra Neutro"
            .AddItem "Vertical Bus Bar"
        End With
    '=============== UNIDAD DPS ==================
        Me.OLEObjects("UnidadDPS").Object.Value = "SFI - Frontal Interno"
        With Me.OLEObjects("FabricanteDPS").Object
            .Clear
            .AddItem "Ciprotec"
            .AddItem "Schneider Electric"
            .AddItem "HWA"
            .AddItem "Phoenix Contac"
            .AddItem "Citel"
        End With
        With Me.OLEObjects("MandoDPS").Object
            .Clear
            .AddItem "Toggle"
        End With
    '=============== UNIDADES DE CONTROL ==================
        Me.OLEObjects("UnidadCTR").Object.Value = "SFI - Frontal Interno"
        With Me.OLEObjects("EquipoCTR").Object
            .Clear
            .AddItem "PM"
            .AddItem "HMI"
            .AddItem "InteliATS_STD"
            .AddItem "ATL600"
            .AddItem "AGC150"
            .AddItem "AGC4"
            .AddItem "DCRL5"
            .AddItem "EASYGEN"
            .AddItem "DSE8610"
            .AddItem "Rapidus"
            .AddItem "Sincronoscopio"
        End With
        With Me.OLEObjects("MedidasCTR").Object
            .Clear
            .AddItem "400"
            .AddItem "720"
        End With
    '=============== SERVICIOS AUXILIARES ==================
        Me.OLEObjects("MontajeSV").Object.Value = "A_Unit"
        With Me.OLEObjects("EquipoSV").Object
            .Clear
            .AddItem "Iluminaci�n"
            .AddItem "Cambio Tension"
            .AddItem "Batery.ST"
            .AddItem "Pont I/O"
            .AddItem "Switch"
            .AddItem "GateWay"
            .AddItem "Toma Corriente"
            .AddItem "A_UnitBB_400"
        End With
    '============== ACCESORIOS DE CONTROL ===================
        With Me.OLEObjects("ElementoCTR").Object
            .Clear
            .AddItem "ClavijaRoja"
            .AddItem "ClavijaAzul"
            .AddItem "Pilotos"
            .AddItem "Parada_Emergencia"
            .AddItem "Selector1"
            .AddItem "Selector2"
            .AddItem "Pulsador"
            .AddItem "Piloto_Verde"
            .AddItem "Piloto_Azul"
            .AddItem "Piloto_Rojo"
            .AddItem "Piloto_Amarillo"
            .AddItem "Piloto_gris"
            .AddItem "Piloto_grisX3"
        End With
    '============== TRANSFERENCIA =============================
        Me.OLEObjects("TipoUnidadTRF").Object.Value = "SFI - Frontal Interno"
        Me.OLEObjects("OrientacionTRF").Object.Value = "Vertical"
        Me.OLEObjects("AnchoTFR").Object.Value = "720"
        Me.OLEObjects("ConexionTRF").Object.Value = "Superior"
    '============== MENU ETOS =================
        With Me.OLEObjects("TipoU_ETO").Object
            .Clear
            .AddItem "SFI - Frontal Interno"
        End With
        With Me.OLEObjects("Ancho_ETO").Object
            .Clear
            .AddItem "400"
            .AddItem "720"
        End With
        With Me.OLEObjects("Unidades_ETO").Object
            .Clear
            .AddItem "4u"
            .AddItem "5u"
            .AddItem "6u"
            .AddItem "7u"
            .AddItem "8u"
            .AddItem "9u"
            .AddItem "10u"
            .AddItem "11u"
            .AddItem "12u"
        End With
    '============== CONMUTADORES =======================
        Me.OLEObjects("TipoUNCOM").Object.Value = "SFI - Frontal Interno"
        With Me.OLEObjects("CorrienteCOM").Object
            .Clear
            .AddItem "25A"
            .AddItem "63A"
            .AddItem "125A"
        End With
        With Me.OLEObjects("OrientacionCOM").Object
            .Clear
            .AddItem "Vertical"
        End With
    '============== COTAS  =============================
        With Me.OLEObjects("CotasDIBUJO").Object
            .AddItem "Ancho: 520mm"
            .AddItem "Ancho: 620mm"
            .AddItem "Ancho: 720mm"
            .AddItem "Ancho: 1120mm"
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
    '============== ARRANQUE MOTOR ===========================
        With Me.OLEObjects("TipoUnidadAM").Object
            .Clear
            .AddItem "SFI - Frontal Interno"
            .AddItem "SIC - Placa de Control"
        End With
        With Me.OLEObjects("EquipoAM").Object
            .Clear
            .AddItem "Contactores"
            .AddItem "Guardamotor"
            .AddItem "GuardamotorXContactores"
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
    Dim tipo As String
    tipo = CStr(Me.Range("A6").Value)

    LlenarCombo Me.OLEObjects("AlturaENV").Object, Altura
    LlenarCombo Me.OLEObjects("AnchoENV").Object, Ancho
    'LlenarCombo Me.OLEObjects("AnchoLSS").Object, AnchoLISAS
    'msgbox MenuUnidades
    Select Case tipo
        Case "Basic"
            ActiveSheet.Shapes("Distribucion").Visible = False
            ActiveSheet.Shapes("lisas").Visible = False
            ActiveSheet.Shapes("DPS").Visible = False
            ActiveSheet.Shapes("Arranque motor").Visible = False
            ActiveSheet.Shapes("Barraje").Visible = False
            ActiveSheet.Shapes("Control").Visible = False
            ActiveSheet.Shapes("A-UNIT").Visible = False
            ActiveSheet.Shapes("Transferencia").Visible = False
            ActiveSheet.Shapes("ETOS").Visible = False
            ActiveSheet.Shapes("Conmutadores").Visible = False
            ThisWorkbook.Sheets("VistaFrontal").Range("D61:D97").Font.Color = RGB(0, 101, 187)
            ThisWorkbook.Sheets("VistaFrontal").Range("D109:D121").Font.Color = RGB(0, 101, 187)
        Case "Neo"
            ActiveSheet.Shapes("Distribucion").Visible = True
            ActiveSheet.Shapes("lisas").Visible = True
            ActiveSheet.Shapes("DPS").Visible = True
            ActiveSheet.Shapes("Arranque motor").Visible = True
            ActiveSheet.Shapes("Barraje").Visible = True
            ActiveSheet.Shapes("Control").Visible = True
            ActiveSheet.Shapes("A-UNIT").Visible = True
            ActiveSheet.Shapes("Transferencia").Visible = True
            ActiveSheet.Shapes("ETOS").Visible = True
            ActiveSheet.Shapes("Conmutadores").Visible = True
            ThisWorkbook.Sheets("VistaFrontal").Range("D61:D97").Font.Color = RGB(255, 255, 255)
            ThisWorkbook.Sheets("VistaFrontal").Range("D109:D121").Font.Color = RGB(255, 255, 255)

    End Select
End Sub
Public Sub Envolventes()
    '========================
        celda = "KP94"
        ThisWorkbook.Sheets("VistaFrontal").Range("KO91:OJ185").Clear
        Set wsActual = ThisWorkbook.Sheets("VistaFrontal")
    '========================
    Dim dato1 As String, dato2 As String, dato3 As String, ResultadoENV As String
    dato1 = Me.OLEObjects("AlturaENV").Object.Value
    dato2 = ThisWorkbook.Sheets("VistaFrontal").Range("A6").Value
    dato3 = Me.OLEObjects("AnchoENV").Object.Value
    Select Case dato3
        Case "400+720": dato3 = "1120"
    End Select
    ResultadoENV = "EFM_" & dato2 & "X" & dato1 & "X" & dato3
    MoverBloque ResultadoENV, "ENVOLVENTES", wsActual, celda
End Sub
Public Sub ACESORIOS_MTM()
    '========================
        celda = "KR113"
        ThisWorkbook.Sheets("VistaFrontal").Range("KO91:OJ185").Clear
        Set wsActual = ThisWorkbook.Sheets("VistaFrontal")
    '========================
    Dim dato1 As String, dato2 As String
    dato1 = Me.OLEObjects("TipoACC").Object.Value
    dato2 = Me.OLEObjects("MedidaACC").Object.Value
    ResultadoACM = "ACC_" & dato1 & "_" & dato2
    MoverBloque ResultadoACM, "ACCESOSRIOS MTM", wsActual, celda
End Sub
Public Sub UNIDADES_FUNCIONALES()
    '========================
        celda = "KR113"
        ThisWorkbook.Sheets("VistaFrontal").Range("KO91:OJ185").Clear
        Set wsActual = ThisWorkbook.Sheets("VistaFrontal")
    '========================
        Dim dato1 As String, dato2 As String, dato3 As String, dato4 As String
    Dim dato5 As String, dato6 As String

    dato1 = Me.OLEObjects("TipoUF").Object.Value
    Select Case dato1
        Case "SFI - Frontal Interno": dato1 = "SFI_"
    End Select
    dato2 = Me.OLEObjects("FrameUF").Object.Value
    dato3 = Me.OLEObjects("MedidaUF").Object.Value
    dato4 = Me.OLEObjects("OrientacionUF").Object.Value
    Select Case dato4
        Case "Horizontal": dato4 = "H"
        Case "H.Compuesta": dato4 = "HC"
        Case "Vertical": dato4 = "V"
        Case "V.Compacta": dato4 = "VC"
    End Select
    dato6 = Me.OLEObjects("MandoUF").Object.Value
    Select Case dato6
        Case "Toggle": dato6 = "To"
    End Select

    ResultadoUF = dato1 & dato2 & dato6 & "X" & dato3 & dato4
    'msgbox ResultadoUF
    MoverBloque ResultadoUF, "U_FUNCIONALES", wsActual, celda
End Sub
Public Sub UNIDADES_LISAS()
    '========================
        celda = "KR113"
        ThisWorkbook.Sheets("VistaFrontal").Range("KO91:OJ185").Clear
        Set wsActual = ThisWorkbook.Sheets("VistaFrontal")
    '========================
        Dim dato1 As String, dato2 As String, dato3 As String
    dato1 = Me.OLEObjects("TipoLSS").Object.Value

    Select Case dato1
        Case "SFI - Frontal Interno": dato1 = "SFI"
        Case "SIC - Placa de Control": dato1 = "SIC"
    End Select

    dato2 = Me.OLEObjects("AnchoLSS").Object.Value
    dato3 = Me.OLEObjects("UnidadLSS").Object.Value
    ResultadoUL = "GF_" & dato1 & dato2 & "X" & dato3
    'msgbox ResultadoUL
    MoverBloque ResultadoUL, "LISAS", wsActual, celda
End Sub
Public Sub BARRAJE()
    '========================
        celda = "KR101"
        ThisWorkbook.Sheets("VistaFrontal").Range("KO91:OJ185").Clear
        Set wsActual = ThisWorkbook.Sheets("VistaFrontal")
    '========================
        Dim dato1 As String, dato2 As String, dato3 As String
    dato1 = Me.OLEObjects("OrientacionBRR").Object.Value
    Select Case dato1
        Case "Horizontal Bus Bar": dato1 = "HBB"
        Case "Horizontal Tierra Neutro": dato1 = "HTN"
        Case "Vertical Bus Bar": dato1 = "VBB"
    End Select
    dato2 = Me.OLEObjects("AmperajeBRR").Object.Value
    dato3 = Me.OLEObjects("AnchoBRR").Object.Value
    ResultadoBAR = "GF_" & dato1 & "_" & dato2 & "X" & dato3
    'msgbox ResultadoBAR
    MoverBloque ResultadoBAR, "GF_BARRAJE", wsActual, celda
End Sub
Public Sub DPS()
    '========================
        celda = "KR116"
        ThisWorkbook.Sheets("VistaFrontal").Range("KO91:OJ185").Clear
        Set wsActual = ThisWorkbook.Sheets("VistaFrontal")
    '========================
    Dim dato1 As String, dato2 As String, dato3 As String, dato4 As String, dato5 As String
    dato1 = Me.OLEObjects("UnidadDPS").Object.Value
    Select Case dato1
        Case "SFI - Frontal Interno": dato1 = "GFSFI_D"
    End Select
    dato2 = Me.OLEObjects("FabricanteDPS").Object.Value
    Select Case dato2
        Case "Ciprotec": dato2 = "CP"
        Case "Citel": dato2 = "CI"
        Case "HWA": dato2 = "HW"
        Case "Schneider Electric": dato2 = "SE"
        Case "Phoenix Contac": dato2 = "PC"
    End Select
    dato3 = Me.OLEObjects("EquipoDPS").Object.Value
    dato4 = Me.OLEObjects("MandoDPS").Object.Value
    Select Case dato4
        Case "Toggle": dato4 = "To"
    End Select
    dato5 = Me.OLEObjects("AnchoDPS").Object.Value

    ResultadoDPS = dato1 & dato2 & "X" & dato5 & "_" & dato3 & dato4
    'msgbox ResultadoDPS
    MoverBloque ResultadoDPS, "DPS", wsActual, celda
End Sub
Public Sub Equipos_Control()
    '========================
        celda = "KR116"
        ThisWorkbook.Sheets("VistaFrontal").Range("KO91:OJ185").Clear
        Set wsActual = ThisWorkbook.Sheets("VistaFrontal")
    '========================
    Dim dato1 As String, dato2 As String
    dato1 = Me.OLEObjects("EquipoCTR").Object.Value
    dato2 = Me.OLEObjects("MedidasCTR").Object.Value
    ResultadoCT = "GFSFI_" & dato1 & "X" & dato2
    MoverBloque ResultadoCT, "UCONTROL", wsActual, celda
End Sub
Public Sub Servicios_Auxiliares()
    '========================
        celda = "KR116"
        ThisWorkbook.Sheets("VistaFrontal").Range("KO91:OJ185").Clear
        Set wsActual = ThisWorkbook.Sheets("VistaFrontal")
    '========================
    Dim dato1 As String, dato2 As String, dato3 As String, dato4 As String
    dato1 = Me.OLEObjects("MontajeSV").Object.Value
    Select Case dato1
        Case "A_Unit": dato1 = "AU_"
    End Select
    dato2 = Me.OLEObjects("EquipoSV").Object.Value
    Select Case dato2
        Case "Cambio Tension": dato2 = "CambioTension"
        Case "Pont I/O": dato2 = "PontI_O"
        Case "Toma Corriente": dato2 = "TomaCorriente"
    End Select
    ResultadoSV = "GF" & dato1 & dato2
    MoverBloque ResultadoSV, "SERVICIOS", wsActual, celda
End Sub
Public Sub Accesorios_CTR()
    '========================
        celda = "KR116"
        ThisWorkbook.Sheets("VistaFrontal").Range("KO91:OJ185").Clear
        Set wsActual = ThisWorkbook.Sheets("VistaFrontal")
    '========================
    Dim dato1 As String, dato2 As String
    dato1 = Me.OLEObjects("ElementoCTR").Object.Value
    dato2 = Me.OLEObjects("vistaACCCTR").Object.Value
    Select Case dato2
        Case "Derecha_Abierta": dato2 = "DA"
        Case "Derecha_Cerrada": dato2 = "DC"
        Case "Izquierda_Abierta": dato2 = "IA"
        Case "Izquierda_Cerrada": dato2 = "IC"
    End Select

    ResultadoACC = "GFAC_" & dato1 & "_" & dato2
    MoverBloque ResultadoACC, "UCONTROL", wsActual, celda
End Sub
Public Sub Transferencia()
    '========================
        celda = "KR116"
        ThisWorkbook.Sheets("VistaFrontal").Range("KO91:OJ185").Clear
        Set wsActual = ThisWorkbook.Sheets("VistaFrontal")
    '========================
    Dim dato1 As String, dato2 As String, dato3 As String, dato4 As String, dato5 As String, dato6 As String
    dato1 = Me.OLEObjects("TipoUnidadTRF").Object.Value
    Select Case dato1
        Case "SFI - Frontal Interno": dato1 = "GFSFI"
    End Select
    dato2 = Me.OLEObjects("OrientacionTRF").Object.Value
    Select Case dato2
        Case "Vertical": dato2 = "V"
    End Select
    dato3 = Me.OLEObjects("MandoTFF").Object.Value
    Select Case dato3
        Case "Manual": dato3 = "MN"
        Case "Motorizado": dato3 = "M"
    End Select
    dato4 = Me.OLEObjects("AnchoTFR").Object.Value
    dato5 = Me.OLEObjects("ConexionTRF").Object.Value
    Select Case dato5
        Case "Superior": dato5 = "SUP"
    End Select
    dato6 = Me.OLEObjects("EquipoTRF").Object.Value
    ResultadoTRF = dato1 & "_UT_" & dato6 & dato3 & "X" & dato2 & "X" & dato4 & "X" & dato5
    MoverBloque ResultadoTRF, "ATS_SFI", wsActual, celda
End Sub
Public Sub Unidades_ETO()
    '========================
        celda = "KR116"
        ThisWorkbook.Sheets("VistaFrontal").Range("KO91:OJ185").Clear
        Set wsActual = ThisWorkbook.Sheets("VistaFrontal")
    '========================
    Dim dato1 As String, dato2 As String
    dato1 = Me.OLEObjects("Ancho_ETO").Object.Value
    dato2 = Me.OLEObjects("Unidades_ETO").Object.Value
    ResultadoETO = "ETO" & dato1 & "X" & dato2
    MoverBloque ResultadoETO, "ETO_SFC_SFI", wsActual, celda
End Sub
Public Sub Conmutadores()
    '========================
        celda = "KR116"
        ThisWorkbook.Sheets("VistaFrontal").Range("KO91:OJ185").Clear
        Set wsActual = ThisWorkbook.Sheets("VistaFrontal")
    '========================
    Dim dato1 As String, dato2 As String, dato3 As String, dato4 As String
    dato1 = Me.OLEObjects("TipoUNCOM").Object.Value
    Select Case dato1
        Case "SFI - Frontal Interno": dato1 = "GFSFI"
    End Select
    dato2 = Me.OLEObjects("CorrienteCOM").Object.Value
    Select Case dato2
        Case "25A": dato2 = "C25A"
        Case "63A": dato2 = "C63A"
        Case "125A": dato2 = "C125A"
    End Select
    dato3 = Me.OLEObjects("EquipoCOMT").Object.Value
    dato4 = Me.OLEObjects("OrientacionCOM").Object.Value
    Select Case dato4
        Case "Vertical": dato4 = "V"
    End Select
    ResultadoCOMT = dato1 & "_" & dato2 & "_" & dato3 & "ToX720" & dato4
    MoverBloque ResultadoCOMT, "Conmutadora", wsActual, celda
End Sub
Public Sub Arranque_Motor()
    '========================
        celda = "KR116"
        ThisWorkbook.Sheets("VistaFrontal").Range("KO91:OJ185").Clear
        Set wsActual = ThisWorkbook.Sheets("VistaFrontal")
    '========================
    Dim dato1 As String, dato2 As String, dato3 As String
    dato1 = Me.OLEObjects("TipoUnidadAM").Object.Value
    Select Case dato1
        Case "SFI - Frontal Interno": dato1 = "GFSFI"
        Case "SIC - Placa de Control": dato1 = "GFSIC"
    End Select
    dato2 = Me.OLEObjects("EquipoAM").Object.Value
    dato3 = Me.OLEObjects("TallaAM").Object.Value
    Select Case dato3
        Case "Talla 1-2", "Vertical": dato3 = "T1"
        Case "Talla 3": dato3 = "T3"
    End Select
    ResultadoAM = dato1 & "_" & dato2 & dato3 & "X720"
    MoverBloque ResultadoAM, "Arranque_Motor", wsActual, celda
End Sub
Public Sub COTAS()
    '========================
        celda = "KR116"
        ThisWorkbook.Sheets("VistaFrontal").Range("KO91:OJ185").Clear
        Set wsActual = ThisWorkbook.Sheets("VistaFrontal")
    '========================
    Dim dato1 As String
    dato1 = Me.OLEObjects("CotasDIBUJO").Object.Value
    Select Case dato1
        Case "Ancho: 520mm": dato1 = "CG_520"
        Case "Ancho: 620mm": dato1 = "CG_620"
        Case "Ancho: 720mm": dato1 = "CG_720"
        Case "Ancho: 1120mm": dato1 = "CG_1120"
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
    Set ws = ThisWorkbook.Sheets("VistaFrontal")
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
        ThisWorkbook.Sheets("VistaFrontal").Range("KO91:OJ185").Clear
        Set wsActual = ThisWorkbook.Sheets("VistaFrontal")
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



