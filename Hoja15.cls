' DISENO G-FLEX
Public Altura as string
Public Ancho as string, AnchoLISAS as string
Public Profundidad as string
Public IP as string, Color as string 
'ENVOLVENTES
    Private sub AlturaENV_change()
        dim dato1 as string
        dato1 = Me.OleObjects("AlturaENV").Object.Value
        with Me.OLEObjects("AnchoENV").Object
            .Clear
            select case dato1
                case "600": .AddItem "400" 'neo
                case "700": .AddItem "520" 'basic
                case "900": .AddItem "720" 'neo, basis
                case "800": .AddItem "620" 'basic
                case "1100" ' neo
                .AddItem "720"
                .AddItem "400+720"
            end select
        end with
    End sub
'ACCESORIOS MTM
    Private sub TipoACC_Change()
        Dim dato As String
        dato = Me.OleObjects("TipoACC").Object.Value
        with Me.OleObjects("MedidaACC").Object
            .Clear
            select case dato
                case "Fijacion_muro"
                    .AddItem "Frontal"
                    '.AddItem "Lateral"
                    '.AddItem "Posterior"
                case "Fijacion_poste"
                    .AddItem "Lateral"
                    .AddItem "Posterior"
                case "Ventilador"
                    .AddItem "Frontal"
                    '.AddItem "Lateral"
            end select
        end with
    end sub
'UNIDADES FUNCIONALES
    Private mBloq As Boolean
    Private Const O720 As String = "Horizontal,H.Compuesta,V.Compacta"
    ' 1) FRAME → llena Medida y Corriente
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
    ' 2) MEDIDA → llena Orientación (400 = Vertical, 720 = más opciones, salvo excepciones)
    Private Sub MedidaUF_Change()
        If mBloq Then Exit Sub
        Dim m As String, o As String
        m = Me.OLEObjects("MedidaUF").Object.Text
        Select Case Me.OLEObjects("FrameUF").Object.Text & "|" & m
            Case "MiniBreaker|720", "MiniBreaker|400", "MiniBreaker_ONKA|400", _
             "MiniBreaker_ONKA|720","GoPact100|400","EZC100|400", _
             "GoPact150|400","GoPact200|400", "GoPact250|400", _
             "EZC250|400", "CVS250|400","CVS100|400":                                                   o = "Vertical"
        '_________________________________________________________________________________________________________________________
            Case "GoPact100|720","GoPact200|720", "GoPact250|720", _
             "GoPact400|720", "EZC100|720","EZC250|720", "CVS250|720", _
              "NSX250|720":                                                                            o = "Horizontal,H.Compuesta"
        '_________________________________________________________________________________________________________________________
            Case  "GoPact100|720", "GoPact250|720", "EZC100|720","EZC250|720", _
             "CVS250|720", "NSX100|400","NSX250|720","NSX250|400":                                     o = "V.Compacta"

            Case Else: o = IIf(m = "400", "Vertical", IIf(m = "720", O720, ""))
        End Select
        Llenar "OrientacionUF", o
        OrientacionUF_Change
    End Sub
    ' 3) ORIENTACIÓN → llena Función
    Private Sub OrientacionUF_Change()
        If mBloq Then Exit Sub
        Select Case Me.OLEObjects("OrientacionUF").Object.Text
            Case "":            Llenar "FuncionUF", ""
            Case "H.Compuesta": Llenar "FuncionUF", "Totalizador"
            Case Else:          Llenar "FuncionUF", "Carga"
        End Select
    End Sub
    ' Llena el combo, autoselecciona si hay 1 opción y conserva la anterior si sigue válida
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
    Private sub AnchoLSS_Change()
        dim dato as string
        dato = Me.OleObjects("AnchoLSS").Object.Value
        with Me.OLEObjects("UnidadLSS").Object
            .Clear
            select case dato
                case "400"
                    IF Me.OleObjects("TipoLSS").Object.Value = "SFI - Frontal Interno" Then
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

                case "720"
                     IF Me.OleObjects("TipoLSS").Object.Value = "SFI - Frontal Interno" Then
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
                    ELSE
                        .AddItem "5u"
                        .AddItem "6u"
                        .AddItem "7u"
                        .AddItem "8u"
                        .AddItem "9u"
                        .AddItem "10u"
                        .AddItem "11u"
                        .AddItem "12u"
                    End If
                end select
            end with
    end sub   
'UNIDADES BARRAJE
    Private sub OrientacionBRR_Change()    
        dim dato as string
        dato = Me.OleObjects("OrientacionBRR").Object.Text
        with Me.OLEObjects("AmperajeBRR").Object  
            .Clear
            Select case dato
                case "Horizontal Bus Bar"
                    .AddItem "400A"
                Case "Horizontal Tierra Neutro"
                    .AddItem "NA" 
                Case "Vertical Bus Bar"
                    .AddItem "NSX400"
                    .AddItem "EZC400"
                    .AddItem "CVS400"
            End Select
        end with
        with Me.OLEObjects("AnchoBRR").Object
            .Clear
            Select case dato
                case "Horizontal Bus Bar"
                    .AddItem "720"
                Case "Horizontal Tierra Neutro"
                    .AddItem "720"
                    .AddItem "400"
                Case "Vertical Bus Bar"
                    .AddItem "400"
            End Select
        end with
    end sub
'UNIDADES DPS
    Private sub FabricanteDPS_Change()
        Dim dato As String
        dato = Me.OleObjects("FabricanteDPS").Object.Value
        With Me.OleObjects("EquipoDPS").Object
            .Clear
            select case dato
                case "Ciprotec","Citel","Phoenix Contac","Schneider Electric"
                    .AddItem "MiniBreaker"
                    .AddItem "EZC100"
                    .AddItem "CVS100"
                    .AddItem "NSX100"
                case "HWA"
                    .AddItem "EZC100"
                    .AddItem "CVS100"
                    .AddItem "NSX100"
            end select
        end with
        with Me.OleObjects("AnchoDPS").Object
            .Clear
            select case dato
                case "Ciprotec","Citel","Phoenix Contac","Schneider Electric"
                    .AddItem "400"
                    .AddItem "720"
                case "HWA"
                    .AddItem "720"    
            end select
        end with
    end sub 
'Accesorios de control  
    Private sub ElementoCTR_Change()
        Dim dato As String
        dato = Me.OleObjects("ElementoCTR").Object.Value
        With Me.OleObjects("vistaACCCTR").Object
            .Clear
            select case dato
                case "ClavijaRoja", "ClavijaAzul"
                    .AddItem "Frontal"
                    .AddItem "Derecha_Abierta"
                    .AddItem "Derecha_Cerrada"
                    .AddItem "Izquierda_Abierta"
                    .AddItem "Izquierda_Cerrada"
                case else
                    .AddItem "NA"
            end select
        End With
    End Sub 
'TRANFERENCIAS
    Private sub OrientacionTRF_Change()
        DIM dato As String
        dato = Me.OleObjects("OrientacionTRF").Object.Value
        With Me.OleObjects("EquipoTRF").Object
            .Clear
            select case dato 
                case "Vertical"
                    .additem "EZC250"
                    .additem "EZC630"
                    .additem "CVS250"
                    .additem "CVS630"
                    .additem "NSX250"
                    .additem "NSX630"
                    .additem "GoPact250"
                    .additem "GoPact400"
                    .additem "NS800"
                case "Horizontal"   
                    .additem "EZC250X2"
                    .additem "CVS250X2" 
                    .additem "NSX250X2"
                    .additem "GoPact250X2"
            end select
        end with
    end sub
    Private sub EquipoTRF_Change()
        Dim dato As String
        dato = Me.OleObjects("EquipoTRF").Object.Value
        With Me.OleObjects("MandoTFF").Object
            .Clear
            select case dato 
                case "CVS250","NSX250","NSX400","CVS400"
                    .AddItem "Manual"
                    .AddItem "Motorizado"
                CASE "EZC250","EZC400","GoPact250","GoPact400"
                    .AddItem "Manual" 
            end select
        end with
    end sub
    'Conmutadoras
    private sub CorrienteCOM_Change()
        Dim dato As String
        dato = Me.OleObjects("CorrienteCOM").Object.Value
        With Me.OleObjects("EquipoCOMT").Object
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
        dato = Me.OleObjects("ClaseETQ").Object.Value
        with Me.OleObjects("TipoETQ").Object  
            .Clear
            select case dato
                case "Etiqueta" 
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
                    .additem "Senales_extrernas_transferencia"
                    .additem "Parametrizacion_medidor"
                    .additem "Control_produccion"
                    .additem "Incumplimiento_instrucciones"
                    .additem "Mantenimiento"
                    .additem "Cableado"
                    .additem "Superficie_caliente"
                    .additem "Linea"
                    .additem "Generador"
                    .additem "Red"
                    .additem "Transformador"
                    .additem "UPS"
                    .additem "Blindobarra"
                    .additem "Blindobarra_Transformador"
                    .additem "Blindobarra_Generador"
                case "Linea de Producto"
                    .AddItem "B_Rack"
                    .AddItem "C_Rack"
                    .AddItem "C_Box"
                    .AddItem "D_Critical"
                    .additem "E_Center"
                    .additem "G_Flex"
                    .additem "I_MCC"
                    .additem "L_Meter"
                    .additem "M_Board"
                    .additem "P_Ats"
                    .additem "Q_Bank"
                    .additem "S_Energy"
                    .additem "T_Meter"
                    .additem "T_Power"
                    .additem "V_Drive"
                    .additem "X_Proof"
                case "Logo de clientes"    
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
            end select
        End With
    End Sub  
    public sub Clase_etiqueta()
        dim dato as string
        dato = Me.OleObjects("ClaseETQ").Object.Value
        select case dato
            case "Etiqueta"
                ActiveSheet.Shapes("Dato3titulo").TextFrame2.TextRange.Text = "Tamaño:"
                Me.OleObjects("Dato3ETQ").Object.Enabled = true
            case "Linea de Producto"  
                ActiveSheet.Shapes("Dato3titulo").TextFrame2.TextRange.Text = "Color:"
                Me.OleObjects("Dato3ETQ").Object.Enabled = true
            case "Logo de clientes" 
                ActiveSheet.Shapes("Dato3titulo").TextFrame2.TextRange.Text = "......"
                Me.OleObjects("Dato3ETQ").Object.Enabled = false
        end select
    end sub   
    Private sub TipoETQ_Change()
        Dim dato As String, dato2 As String
        dato2 = Me.OleObjects("ClaseETQ").Object.Value
        dato = Me.OleObjects("TipoETQ").Object.Value
        with Me.OleObjects("Dato3ETQ").Object
            .clear
            select case dato
                case "Codigo_QR", "Riesgo_electrico1", "Ventilación", "Descarga", "Riesgo_electrico2", _
                     "Voltaje_peligroso", "Voltaje_208", "Voltaje_220", "Voltaje_440", _
                     "Voltaje_460", "Voltaje_480", "Aviso_acometidas", "Programador_horario", "Equipo_medida", _
                     "Ajuste_higrostato", "Inspeccion_DPS", "No_manipular_Guaya", "Ajuste_corriente", "Mantenimiento_filtros", _
                     "Senales_extrernas_transferencia", "Parametrizacion_medidor", "Control_produccion", "Incumplimiento_instrucciones", _
                     "Mantenimiento", "Cableado", "Superficie caliente"
                    .AddItem "Pequeno" 
                case "Ventilación", "Descarga", "Riesgo_electrico2"
                    .AddItem "Pequeno"
                    .AddItem "Mediano"  
                    .AddItem "Grande"
                case "Placa_GR", "Ventilación", "Descarga", "Riesgo_electrico2", "Riesgo_electricoV1", "Riesgo_electricoV2", _
                     "Riesgo_electricoV3", "Choque_electrico_1", "Choque_electrico_2", "Prohibido_paso", "Prohibido_acceso", "Respaldo", _
                     "Generador", "Red", "Transformador", "UPS"  
                    .AddItem "Mediano" 
                case "Advertencia_choque_electrico"
                    .AddItem "Grande" 
                case "Voltaje_peligroso"   
                    .AddItem "Pequeno"
                    .AddItem "Mediano" 
                case "Blindobarra","Blindobarra_Generador","Blindobarra_Transformador"
                    .AddItem "NA"      
            end select 
            Select case dato2
                case "Linea de Producto"
                    .AddItem "Gris"
                    .AddItem "Azul"
            end select        
        end with
    end sub 
'ARRANQUE MOTOR
    Private sub  EquipoAM_Change()
        Dim dato As String
        dato = Me.OleObjects("EquipoAM").Object.Value
        With Me.OleObjects("TallaAM").Object          
            .Clear
            select case dato
                case "Contactores"  
                    .Additem "Talla 1-2"
                    .Additem "Talla 3"
                case "Guardamotor", "GuardamotorXContactores"
                    .Additem "Vertical"
            end select
        end with
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
        ActiveSheet.Shapes("Menu ServiciosGFL").Visible = false
        ActiveSheet.Shapes("Menu ColorGFL").Visible = False
        ActiveSheet.Shapes("Menu AcotarGFL").Visible = False
        ActiveSheet.Shapes("Menu EtiquetaGFL").Visible = False
        ActiveSheet.Shapes("Menu ArranqueM GF").Visible = False

    '============== ACC MTM ============================ 
        with Me.OleObjects("TipoACC").Object
            .Clear
            .AddItem "Fijacion_muro"
            '.AddItem "Fijacion_poste"
            .AddItem "Ventilador"
        end with 
    '=============== UNIDADES FUNCIONALES ===============
        Me.OleObjects("TipoUF").Object.Value = "SFI - Frontal Interno"
        with Me.OleObjects("FrameUF").Object
            .clear
            .AddItem "MiniBreaker"
            .Additem "MiniBreaker_ONKA"
            .AddItem "GoPact100"
            .AddItem "GoPact150"
            .AddItem "GoPact200"
            .AddItem "GoPact250"
            .AddItem "GoPact400"
            .Additem "EZC100"
            .Additem "EZC250"
            .Additem "CVS100"
            .Additem "CVS250"
            .Additem "NSX100"
            .Additem "NSX250"
        end with
        With Me.OleObjects("MandoUF").Object
            .Clear
            .AddItem "Toggle"   
        End With
    '=============== UNIDADES LISAS =====================    
        With Me.OleObjects("TipoLSS").Object
            .additem "SFI - Frontal Interno"
            .additem "SIC - Placa de Control"
        End With
        With  Me.OleObjects("AnchoLSS").Object
            .Clear
            .AddItem "400"
            .AddItem "720"
        End With 
    '=============== UNIDAD DE BARRAJE =============
        With  Me.OleObjects("OrientacionBRR").Object   
            .Clear
            .AddItem "Horizontal Bus Bar"
            .AddItem "Horizontal Tierra Neutro"
            .AddItem "Vertical Bus Bar"
        End With
    '=============== UNIDAD DPS ==================
        Me.OleObjects("UnidadDPS").Object.value = "SFI - Frontal Interno"
        With  Me.OleObjects("FabricanteDPS").Object
            .Clear
            .AddItem "Ciprotec"
            .AddItem "Schneider Electric"
            .AddItem "HWA"
            .AddItem "Phoenix Contac"
            .AddItem "Citel"
        End With
        with Me.OleObjects("MandoDPS").Object
            .Clear
            .AddItem "Toggle"
        end with
    '=============== UNIDADES DE CONTROL ==================
        Me.OleObjects("UnidadCTR").Object.value = "SFI - Frontal Interno"
        With Me.OleObjects("EquipoCTR").Object
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
        with Me.OleObjects("MedidasCTR").Object
            .Clear    
            .AddItem "400"
            .AddItem "720"
        End With
    '=============== SERVICIOS AUXILIARES ==================
        Me.OleObjects("MontajeSV").Object.Value = "A_Unit"    
        with Me.OleObjects("EquipoSV").Object
            .Clear
            .AddItem "Iluminación"
            .AddItem "Cambio Tension"
            .AddItem "Batery.ST"
            .AddItem "Pont I/O"
            .AddItem "Switch"
            .AddItem "GateWay"
            .AddItem "Toma Corriente"
            .AddItem "A_UnitBB_400"
        end with
    '============== ACCESORIOS DE CONTROL ===================
        With Me.OleObjects("ElementoCTR").Object 
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
        Me.OleObjects("TipoUnidadTRF").Object.value = "SFI - Frontal Interno"
        Me.OleObjects("OrientacionTRF").Object.value = "Vertical"
        Me.OleObjects("AnchoTFR").Object.value = "720"  
        Me.OleObjects("ConexionTRF").Object.Value = "Superior"
    '============== MENU ETOS =================
        With Me.OleObjects("TipoU_ETO").Object   
            .Clear
            .AddItem "SFI - Frontal Interno"
        End With
        With Me.OleObjects("Ancho_ETO").Object
            .Clear
            .AddItem "400"
            .AddItem "720"
        End With
        with Me.OleObjects("Unidades_ETO").Object
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
        Me.OLEObjects("TipoUNCOM").Object.value = "SFI - Frontal Interno"
        With Me.OLEObjects("CorrienteCOM").Object  
            .Clear
            .AddItem "25A"
            .AddItem "63A"
            .AddItem "125A"
        End With
        with Me.OLEObjects("OrientacionCOM").Object
            .Clear
            .AddItem "Vertical"
        end with  
    '============== COTAS  =============================
        With Me.OleObjects("CotasDIBUJO").Object  
            .Additem "Ancho: 520mm"
            .Additem "Ancho: 620mm"
            .Additem "Ancho: 720mm"
            .Additem "Ancho: 1120mm"
            .Additem "Altura: 900mm"
            .Additem "Altura: 800mm"
            .Additem "Altura: 700mm"
            .Additem "Division Tabelero"
            .Additem "Corte Tren"
            .Additem "Continuacion Tren"
        End With
    '============== COLORES  =============================
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
    '============== ETIQUETA  =============================    
        With Me.OleObjects("ClaseETQ").Object
            .Clear
            .AddItem "Etiqueta"
            .AddItem "Linea de Producto"
            .AddItem "Logo de clientes"
        End With  
    '============== ARRANQUE MOTOR ===========================
        With Me.OleObjects("TipoUnidadAM").Object  
            .Clear
            .AddItem "SFI - Frontal Interno"
            .AddItem "SIC - Placa de Control"
        End With 
        With Me.OleObjects("EquipoAM").Object
            .Clear
            .AddItem "Contactores"
            .AddItem "Guardamotor"
            .AddItem "GuardamotorXContactores"
        End With

end sub
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
Public sub Extraer_datos()
    ' EXTRAER DATOS DE DIMENSIONES
    Altura = CStr(Me.Range("A1").Value)
    Ancho = CStr(Me.Range("A2").Value)
    IP = CStr(Me.Range("A4").Value)
    AnchoLISAS= CStr(Me.Range("A5").Value)
    Dim tipo As String
    tipo = CStr(Me.Range("A6").Value)

    LlenarCombo Me.OLEObjects("AlturaENV").Object, Altura
    'LlenarCombo Me.OLEObjects("AnchoENV").Object, Ancho
    'LlenarCombo Me.OLEObjects("AnchoLSS").Object, AnchoLISAS
    'msgbox MenuUnidades
    select case tipo
        case "Basic"
            ActiveSheet.shapes("Distribucion").Visible = False
            ActiveSheet.shapes("lisas").Visible = False
            ActiveSheet.shapes("DPS").Visible = False
            ActiveSheet.shapes("Arranque motor").Visible = False
            ActiveSheet.shapes("Barraje").Visible = False
            ActiveSheet.shapes("Control").Visible = False
            ActiveSheet.shapes("A-UNIT").Visible = False
            ActiveSheet.shapes("Transferencia").Visible = False
            ActiveSheet.shapes("ETOS").Visible = False
            ActiveSheet.shapes("Conmutadores").Visible = False
            ThisWorkbook.Sheets("DISENO_GF").Range("D61:D97").Font.Color = RGB(0, 101, 187)
            ThisWorkbook.Sheets("DISENO_GF").Range("D109:D121").Font.Color = RGB(0, 101, 187)
        case "Neo"
            ActiveSheet.shapes("Distribucion").Visible = True
            ActiveSheet.shapes("lisas").Visible = True
            ActiveSheet.shapes("DPS").Visible = True
            ActiveSheet.shapes("Arranque motor").Visible = True
            ActiveSheet.shapes("Barraje").Visible = True
            ActiveSheet.shapes("Control").Visible = True
            ActiveSheet.shapes("A-UNIT").Visible = True
            ActiveSheet.shapes("Transferencia").Visible = True
            ActiveSheet.shapes("ETOS").Visible = True
            ActiveSheet.shapes("Conmutadores").Visible = True
            ThisWorkbook.Sheets("DISENO_GF").Range("D61:D97").Font.Color = RGB(255, 255, 255)
            ThisWorkbook.Sheets("DISENO_GF").Range("D109:D121").Font.Color = RGB(255, 255, 255)

    end select    

    
End Sub
Public sub Envolventes()
    '========================
        celda = "LM89"
        thisWorkbook.Sheets("DISENO_GF").Range("LJ89:QI182").Clear
        Set wsActual = thisWorkbook.Sheets("DISENO_GF")
    '========================
    dim dato1 as string, dato2 as string, dato3 as string, ResultadoENV as string
    dato1 = Me.OleObjects("AlturaENV").Object.Value
    dato2 = thisWorkbook.Sheets("DISENO_GF").Range("A6").Value
    dato3 = Me.OleObjects("AnchoENV").Object.Value
    select case dato3
        case "400+720": dato3 = "1120"
    end select     
    ResultadoENV= "EFM_" & dato2 & "X" & dato1 & "X" & dato3
    MoverBloque ResultadoENV,"ENVOLVENTES", wsActual, celda
End sub
Public sub ACESORIOS_MTM()
    '========================
        celda = "LN111"
        thisWorkbook.Sheets("DISENO_GF").Range("LJ89:QI182").Clear
        Set wsActual = thisWorkbook.Sheets("DISENO_GF")
    '========================
    Dim dato1 As String, dato2 As String
    dato1 = Me.OleObjects("TipoACC").Object.Value
    dato2 = Me.OleObjects("MedidaACC").Object.Value
    ResultadoACM = "ACC_" & dato1 & "_" & dato2
    MoverBloque ResultadoACM,"ACCESOSRIOS MTM", wsActual, celda
End sub
Public sub UNIDADES_FUNCIONALES()
    '========================
        celda = "LN111"
        thisWorkbook.Sheets("DISENO_GF").Range("LJ89:QI182").Clear
        Set wsActual = thisWorkbook.Sheets("DISENO_GF")
    '========================
        Dim dato1 As String, dato2 As String, dato3 As String, dato4 As String
    dim dato5 as string, dato6 as string

    dato1 = Me.OleObjects("TipoUF").Object.Value
    select case dato1
        case "SFI - Frontal Interno": dato1 = "SFI_"
    end select
    dato2 = Me.OleObjects("FrameUF").Object.Value   
    dato3 = Me.OleObjects("MedidaUF").Object.Value
    dato4 = Me.OleObjects("OrientacionUF").Object.Value
    select case dato4
        case "Horizontal": dato4 = "H"
        case "H.Compuesta": dato4 = "HC"
        case "Vertical": dato4 = "V"
        case "V.Compacta": dato4 = "VC"
    end select 
    dato6 = Me.OleObjects("MandoUF").Object.Value   
    Select case dato6
        case "Toggle": dato6 = "To" 
    End Select

    ResultadoUF= dato1 & dato2 & dato6 & "X" & dato3 & dato4
    'msgbox ResultadoUF
    MoverBloque ResultadoUF,"U_FUNCIONALES", wsActual, celda
end sub
Public sub UNIDADES_LISAS()
    '========================
        celda = "LN111"
        thisWorkbook.Sheets("DISENO_GF").Range("LJ89:QI182").Clear
        Set wsActual = thisWorkbook.Sheets("DISENO_GF")
    '========================
        Dim dato1 As String, dato2 As String, dato3 As String
    dato1 = Me.OleObjects("TipoLSS").Object.Value

    select case dato1
        case "SFI - Frontal Interno": dato1 = "SFI"
        Case "SIC - Placa de Control": dato1 = "SIC"
    end select

    dato2 = Me.OleObjects("AnchoLSS").Object.Value   
    dato3 = Me.OleObjects("UnidadLSS").Object.Value
    ResultadoUL= "GF_" & dato1 & dato2 & "X" & dato3
    'msgbox ResultadoUL
    MoverBloque ResultadoUL,"LISAS", wsActual, celda
end sub
Public sub BARRAJE()
    '========================
        celda = "LN101"
        thisWorkbook.Sheets("DISENO_GF").Range("LJ89:QI182").Clear
        Set wsActual = thisWorkbook.Sheets("DISENO_GF")
    '========================
        Dim dato1 As String, dato2 As String, dato3 As String
    dato1 = Me.OleObjects("OrientacionBRR").Object.Value
    select case dato1
        case "Horizontal Bus Bar": dato1 = "HBB"
        case "Horizontal Tierra Neutro": dato1 = "HTN"
        case "Vertical Bus Bar": dato1 = "VBB"
    end select
    dato2 = Me.OleObjects("AmperajeBRR").Object.Value
    dato3 = Me.OleObjects("AnchoBRR").Object.Value
    ResultadoBAR ="GF_" & dato1 & "_" & dato2 & "X" & dato3
    'msgbox ResultadoBAR
    MoverBloque ResultadoBAR,"GF_BARRAJE", wsActual, celda
end sub
Public sub DPS()
    '========================
        celda = "LK116"
        thisWorkbook.Sheets("DISENO_GF").Range("LJ89:QI182").Clear
        Set wsActual = thisWorkbook.Sheets("DISENO_GF")
    '========================
    Dim dato1 as string, dato2 as string, dato3 as string, dato4 as string, dato5 as string
    dato1 = Me.OleObjects("UnidadDPS").Object.Value
    select case dato1
        case "SFI - Frontal Interno": dato1 = "GFSFI_D"
    End Select
    dato2 = Me.OleObjects("FabricanteDPS").Object.Value
    select case dato2
        case "Ciprotec": dato2 = "CP"
        case "Citel": dato2 = "CI"
        Case "HWA": dato2 = "HW"
        case "Schneider Electric": dato2 = "SE"
        case "Phoenix Contac": dato2 = "PC"
    End Select
    dato3 = Me.OleObjects("EquipoDPS").Object.Value
    dato4 = Me.OleObjects("MandoDPS").Object.Value
    select case dato4
        case "Toggle": dato4 = "To"
    End Select
    dato5 = Me.OleObjects("AnchoDPS").Object.Value

    ResultadoDPS= dato1 & dato2 & "X" & dato5 & "_" & dato3 &  dato4
    'msgbox ResultadoDPS
    MoverBloque ResultadoDPS,"DPS", wsActual, celda
end sub
Public sub Equipos_Control()
    '========================
        celda = "LK116"
        thisWorkbook.Sheets("DISENO_GF").Range("LJ89:QI182").Clear
        Set wsActual = thisWorkbook.Sheets("DISENO_GF")
    '========================
    dim dato1 as string, dato2 as string
    dato1 = Me.OleObjects("EquipoCTR").Object.Value
    dato2 = Me.OleObjects("MedidasCTR").Object.Value
    ResultadoCT = "GFSFI_" & dato1 & "X" & dato2
    MoverBloque ResultadoCT,"UCONTROL", wsActual, celda
End sub
Public sub Servicios_Auxiliares()
    '========================
        celda = "LK116"
        thisWorkbook.Sheets("DISENO_GF").Range("LJ89:QI182").Clear
        Set wsActual = thisWorkbook.Sheets("DISENO_GF")
    '========================
    dim dato1 as string, dato2 as string, dato3 as string, dato4 as string
    dato1 = Me.OleObjects("MontajeSV").Object.Value
    select case dato1
        case "A_Unit": dato1 = "AU_"
    End select
    dato2 = Me.OleObjects("EquipoSV").Object.Value
    select case dato2
        case "Cambio Tension": dato2 = "CambioTension"
        case "Pont I/O": dato2 = "PontI_O"
        case "Toma Corriente": dato2 = "TomaCorriente"
    end select    
    ResultadoSV= "GF"& dato1 & dato2
    MoverBloque ResultadoSV,"SERVICIOS", wsActual, celda 
End sub
Public sub Accesorios_CTR()
    '========================
        celda = "LK116"
        thisWorkbook.Sheets("DISENO_GF").Range("LJ89:QI182").Clear
        Set wsActual = thisWorkbook.Sheets("DISENO_GF")
    '========================
    dim dato1 as string, dato2 as string
    dato1 = Me.OleObjects("ElementoCTR").Object.Value
    dato2 = Me.OleObjects("vistaACCCTR").Object.Value
    Select case dato2
        case "Derecha_Abierta": dato2 = "DA"
        case "Derecha_Cerrada": dato2 = "DC"
        case "Izquierda_Abierta": dato2 = "IA"
        case "Izquierda_Cerrada": dato2 = "IC"
    End Select

    ResultadoACC = "GFAC_" & dato1 & "_" & dato2
    MoverBloque ResultadoACC,"UCONTROL", wsActual, celda
End sub
Public sub Transferencia()
    '========================
        celda = "LK116"
        thisWorkbook.Sheets("DISENO_GF").Range("LJ89:QI182").Clear
        Set wsActual = thisWorkbook.Sheets("DISENO_GF")
    '========================
    Dim dato1 As String, dato2 As String, dato3 As String, dato4 As String, dato5 As String, dato6 As String
    dato1 = Me.OleObjects("TipoUnidadTRF").Object.Value
    Select case dato1
        case "SFI - Frontal Interno": dato1 = "GFSFI"
    End Select
    dato2 = Me.OleObjects("OrientacionTRF").Object.Value
    Select case dato2
        Case "Vertical": dato2 = "V"
    End Select
    dato3 = Me.OleObjects("MandoTFF").Object.Value
    select case dato3
        case "Manual": dato3 = "MN"
        case "Motorizado": dato3 = "M"
    end select
    dato4 = Me.OleObjects("AnchoTFR").Object.Value
    dato5 = Me.OleObjects("ConexionTRF").Object.Value
    select case dato5
        case "Superior": dato5 = "SUP"
    end select
    dato6 = Me.OleObjects("EquipoTRF").Object.Value
    ResultadoTRF = dato1 & "_UT_" & dato6 & dato3 & "X" & dato2 & "X" & dato4 & "X" & dato5
    MoverBloque ResultadoTRF,"ATS_SFI", wsActual, celda
End sub
Public sub Unidades_ETO()
    '========================
        celda = "LK116"
        thisWorkbook.Sheets("DISENO_GF").Range("LJ89:QI182").Clear
        Set wsActual = thisWorkbook.Sheets("DISENO_GF")
    '========================
    dim dato1 as string, dato2 as string
    dato1 = Me.OleObjects("Ancho_ETO").Object.Value
    dato2 = Me.OleObjects("Unidades_ETO").Object.Value
    ResultadoETO= "ETO" & dato1 & "X" & dato2 
    MoverBloque ResultadoETO,"ETO_SFC_SFI", wsActual, celda
End sub
Public sub Conmutadores()
    '========================
        celda = "LK116"
        thisWorkbook.Sheets("DISENO_GF").Range("LJ89:QI182").Clear
        Set wsActual = thisWorkbook.Sheets("DISENO_GF")
    '========================
    Dim dato1 As String, dato2 As String, dato3 As String, dato4 As String
    dato1 = Me.OleObjects("TipoUNCOM").Object.Value
    select case dato1
        case "SFI - Frontal Interno": dato1 = "GFSFI"
    end select
    dato2 = Me.OleObjects("CorrienteCOM").Object.Value
    select case dato2
        case "25A": dato2 = "C25A"
        case "63A": dato2 = "C63A"
        case "125A": dato2 = "C125A"
    End select
    dato3 = Me.OleObjects("EquipoCOMT").Object.Value 
    dato4 = Me.OleObjects("OrientacionCOM").Object.Value   
    select case dato4
        case "Vertical": dato4 = "V"
    end select 
    ResultadoCOMT= dato1 & "_" & dato2 & "_" & dato3 & "ToX720" & dato4
    MoverBloque ResultadoCOMT,"Conmutadora", wsActual, celda
End Sub
Public sub Arranque_Motor()
    '========================
        celda = "LK116"
        thisWorkbook.Sheets("DISENO_GF").Range("LJ89:QI182").Clear
        Set wsActual = thisWorkbook.Sheets("DISENO_GF")
    '========================
    Dim dato1 As String, dato2 As String, dato3 As String
    dato1 = Me.OleObjects("TipoUnidadAM").Object.Value
    select case dato1
        case "SFI - Frontal Interno": dato1 = "GFSFI"
        case "SIC - Placa de Control": dato1 = "GFSIC"
    end select
    dato2 = Me.OleObjects("EquipoAM").Object.Value
    dato3 = Me.OleObjects("TallaAM").Object.Value
    select case dato3
        case "Talla 1-2","Vertical": dato3 = "T1"
        Case "Talla 3": dato3 = "T3"
    end select
    ResultadoAM= dato1 & "_" & dato2 & dato3 & "X720" 
    MoverBloque ResultadoAM,"Arranque_Motor", wsActual, celda
End Sub
public sub COTAS()
    '========================
        celda = "LK116"
        thisWorkbook.Sheets("DISENO_GF").Range("LJ89:QI182").Clear
        Set wsActual = thisWorkbook.Sheets("DISENO_GF")
    '========================
    Dim dato1 As String
    dato1 = Me.OleObjects("CotasDIBUJO").Object.Value
    Select case dato1
        case "Ancho: 520mm": dato1 = "CG_520"
        case "Ancho: 620mm": dato1 = "CG_620"
        case "Ancho: 720mm": dato1 = "CG_720"
        case "Ancho: 1120mm": dato1 = "CG_1120"
        case "Altura: 900mm": dato1 = "CG_900"
        case "Altura: 800mm": dato1 = "CG_800"
        case "Altura: 700mm": dato1 = "CG_700"
        case "Division Tabelero": dato1 = "Division"
        case "Corte Tren": dato1 = "Corte_Tren"
        case "Continuacion Tren": dato1 = "Continuidad_Tren"
    end select
    ResultadoCOT = dato1
    MoverBloque ResultadoCOT,"COTAS", wsActual, celda
End sub
Public Sub Cambiodecolor() 'Vista frontal
    Dim ws As Worksheet
    Dim celda As Range
    Dim valorColor As String
    Dim colorDestino As Long
    
    ' Obtener el valor del color desde la portada
    valorColor = Me.OLEObjects("ColorConf").Object.Value
    Set ws = ThisWorkbook.Sheets("DISENO_GF")
    ' Definir color destino según el valor
    Select Case valorColor
        Case "RAL 7035": colorDestino = RGB(197, 199, 196)
        Case "RAL 7032": colorDestino = RGB(181, 176, 161)
        Case "RAL 9003": colorDestino = RGB(236, 236, 231)
        Case "RAL 9004": colorDestino = RGB(43, 43, 44)
        Case "Naranja" : colorDestino= RGB(201,133,12)
        Case "Amarillo": colorDestino =RGB(235,198,21)
    End Select
    
    ' Recorrer el rango y cambiar solo los colores específicos
    For Each celda In ws.Range("BZ57:KI155")
        Select Case valorColor
            Case "RAL 7035"
                If celda.Interior.Color = RGB(181, 176, 161) Or _
                   celda.Interior.Color = RGB(43, 43, 44) Or _
                   celda.Interior.Color = RGB(201, 133, 12) Or _
                   celda.Interior.Color = RGB(235, 198, 21) Or _
                   celda.Interior.Color = RGB(181, 184, 177) Or _
                   celda.Interior.Color = RGB(244, 244, 244) Or _
                   celda.Interior.Color = RGB(236, 236, 231) Then
                    celda.Interior.Color = RGB(197, 199, 196)
                End If
            Case "RAL 7032"
                If celda.Interior.Color = RGB(197, 199, 196) Or _
                   celda.Interior.Color = RGB(43, 43, 44) Or _
                   celda.Interior.Color = RGB(201, 133, 12) Or _
                   celda.Interior.Color = RGB(235, 198, 21) Or _
                   celda.Interior.Color = RGB(244, 244, 244) Or _
                   celda.Interior.Color = RGB(236, 236, 231) Then
                    celda.Interior.Color = RGB(181, 176, 161)
                End If
            Case "RAL 9003"
                If celda.Interior.Color = RGB(197, 199, 196) Or _
                   celda.Interior.Color = RGB(181, 176, 161) Or _
                   celda.Interior.Color = RGB(43, 43, 44) Or _
                   celda.Interior.Color = RGB(201, 133, 12) Or _
                   celda.Interior.Color = RGB(235, 198, 21) Or _
                   celda.Interior.Color = RGB(244, 244, 244) Then
                    celda.Interior.Color = RGB(236, 236, 231)
                End If
            Case "RAL 9004"
                If celda.Interior.Color = RGB(197, 199, 196) Or _
                   celda.Interior.Color = RGB(181, 176, 161) Or _
                   celda.Interior.Color = RGB(201, 133, 12) Or _
                   celda.Interior.Color = RGB(235, 198, 21) Or _
                   celda.Interior.Color = RGB(244, 244, 244) Or _
                   celda.Interior.Color = RGB(236, 236, 231) Then
                    celda.Interior.Color = RGB(43, 43, 44)
                End If
            Case "Naranja"
                If celda.Interior.Color = RGB(197, 199, 196) Or _
                   celda.Interior.Color = RGB(181, 176, 161) Or _
                   celda.Interior.Color = RGB(43, 43, 44) Or _
                   celda.Interior.Color = RGB(235, 198, 21) Or _
                   celda.Interior.Color = RGB(244, 244, 244) Or _
                   celda.Interior.Color = RGB(236, 236, 231) Then
                    celda.Interior.Color = RGB(201, 133, 12)
                End If
            Case "Amarillo"
                If celda.Interior.Color = RGB(197, 199, 196) Or _
                   celda.Interior.Color = RGB(181, 176, 161) Or _
                   celda.Interior.Color = RGB(43, 43, 44) Or _
                   celda.Interior.Color = RGB(201, 133, 12) Or _
                   celda.Interior.Color = RGB(244, 244, 244) Or _
                   celda.Interior.Color = RGB(236, 236, 231) Then
                    celda.Interior.Color = RGB(235, 198, 21)
                End If
        End Select
    Next celda
    
    MSGBox "Proceso Finalizado con Exito", vbInformation, "Cambio de Color"
End Sub
Public sub  ETIQUETAS()
    '========================
        celda = "LK116"
        thisWorkbook.Sheets("DISENO_GF").Range("LJ89:QI182").Clear
        Set wsActual = thisWorkbook.Sheets("DISENO_GF")
    '========================
    Dim dato1 As String, dato2 As String, dato3 As String
    dato2 = Me.OleObjects("Dato3ETQ").Object.Value
    select case dato2
        case "Pequeno": dato2 = "1"
        case "Mediano": dato2 = "2"
        case "Grande": dato2 = "3"
    end select
    dato1 = Me.OleObjects("ClaseETQ").Object.Value
    dato3 = Me.OleObjects("TipoETQ").Object.Value
    Select case dato1
        case "Etiqueta"
            dato1 = "GFE_"
            ResultadoETQ = dato1 & dato3 & "X" & dato2

        case "Linea de Producto"
            dato1 = "GFL_"
            ResultadoETQ = dato1 & dato3 & "X" & dato2 

        case "Logo de clientes": dato1 = "GFLC_"
            ResultadoETQ = dato1  & dato3
    End Select
    MoverBloque ResultadoETQ,"ETIQUETAS", wsActual, celda
End sub