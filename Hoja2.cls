'========================== DISENO TREN =======================
Public Altura as string
Public Ancho as string, AnchoLISAS as string
Public Profundidad as string
Public IP as string, Color as string 
'======================================================================
'  UNIDAD FUNCIONAL (UF) - Cascada de combos ActiveX de la hoja
'
'      TipoUF --> FrameUF --+--> MedidaUF --> OrientacionUF
'                           +--> CorrienteUF
'
'  La ORIENTACION depende de TipoUF + FrameUF + MEDIDA, por eso se
'  llena en MedidaUF_Change y ya no dentro de FrameUF_Change.
'
'  REGLA NUEVA: si MedidaUF = "400" -> OrientacionUF = "Vertical"
'======================================================================
    Private mbCargando As Boolean          'evita eventos en cascada
    '---------------------------------------------------------------- EVENTOS
    Private Sub TipoUF_Change()
        If mbCargando Then Exit Sub
        On Error GoTo Fin
        mbCargando = True
        LlenarFrame
        LlenarMando
    Fin:
        mbCargando = False
        RecargarDesdeFrame                 'reconstruye Medida, Corriente y Orientacion
    End Sub

    Private Sub FrameUF_Change()
        If mbCargando Then Exit Sub
        RecargarDesdeFrame
    End Sub
    Private Sub MedidaUF_Change()
        If mbCargando Then Exit Sub
        RecargarDesdeMedida
    End Sub
    '------------------------------------------------------------ ORQUESTADOR
    Private Sub RecargarDesdeFrame()
        On Error GoTo Fin
        mbCargando = True
        LlenarMedida
        LlenarCorriente
    Fin:
        mbCargando = False
        RecargarDesdeMedida                'antes llamaba solo a LlenarOrientacion
    End Sub
    Private Sub OrientacionUF_Change()
            If mbCargando Then Exit Sub
            On Error GoTo Fin
            mbCargando = True
            LlenarFuncionalidad
        Fin:
            mbCargando = False
            LlenarConexion
    End Sub
    Private Sub FuncionalidadUF_Change()   'CORREGIDO: tu stub se llamaba solo "FuncionalidadUF"
        If mbCargando Then Exit Sub
        LlenarConexion
    End Sub
    Private Sub RecargarDesdeMedida()
            On Error GoTo Fin
            mbCargando = True
            LlenarOrientacion
            LlenarFuncionalidad
        Fin:
            mbCargando = False
            LlenarConexion
    End Sub
    '--------------------------------------------------------------- UTILIDADES
    Private Function Cbo(ByVal sNombre As String) As Object
        Set Cbo = Me.OLEObjects(sNombre).Object
    End Function

    ' El "& """"" evita el error 94 (uso invalido de Null) con el combo vacio.
    Private Function Txt(ByVal sNombre As String) As String
        Txt = Trim$(Cbo(sNombre).Value & "")
    End Function

    ' AJUSTE: MenuUnidades no se puede leer suelto con Option Explicit.
    '   - Si es un nombre definido o celda -> deje esta version.
    '   - Si es una variable Public de un modulo -> use: MenuActivo = MenuUnidades
    '   - Si es otro combo ActiveX de la hoja  -> use: MenuActivo = Txt("MenuUnidades")
    Private Function MenuActivo() As String
        MenuActivo = Trim$(ThisWorkbook.Names("MenuUnidades").RefersToRange.Value & "")
    End Function
    '---------------------------------------------------------------- FRAME
    Private Sub LlenarFrame()
        Dim sTipo As String
        sTipo = Txt("TipoUF")

        With Cbo("FrameUF")
        .Clear
        Select Case sTipo
            Case "SFI - Frontal Interno", "SFC - Frontal Compacto"
                If MenuUnidades = "Opcion1" Then
                    .AddItem "MiniBreaker"
                    .AddItem "EZC100"
                    .AddItem "EZC250"
                    .AddItem "EZC630"
                    .AddItem "CVS100"
                    .AddItem "CVS250"
                    .AddItem "CVS630"
                    .AddItem "NSX100"
                    .AddItem "NSX250"
                    .AddItem "NSX630"
                    .AddItem "NS800_1250"
                    .AddItem "GoPact150"
                    .AddItem "GoPact200"
                    .AddItem "GoPact250"
                    .AddItem "GoPact400"
                    .AddItem "GoPact800"
                ElseIf MenuUnidades = "Opcion2" Then
                .AddItem "EZC250"
                .AddItem "EZC630"
                .AddItem "CVS100"
                .AddItem "CVS250"
                .AddItem "CVS630"
                .AddItem "NSX100"
                .AddItem "NSX250"
                .AddItem "NSX630"
                .AddItem "NS800_1250"
                .AddItem "NS800_1600"
                .AddItem "GoPact150"
                .AddItem "GoPact200"
                .AddItem "GoPact250"
                .AddItem "GoPact400"
                .AddItem "GoPact800"
                .AddItem "NWFIJO"
                .AddItem "NWEXTRAIBLE"
                .AddItem "MTZFIJO"
                .AddItem "MTZEXTRAIBLE"
                .AddItem "MVSFIJO"
                .AddItem "MVSEXTRAIBLE"
                End If
            Case "SIA - Integral Abatible"
            If MenuUnidades= "Opcion1" Then
            .AddItem "EZC100"
            .AddItem "EZC250"
            .AddItem "CVS100"
            .AddItem "CVS250"
            .AddItem "CVS630"
            .AddItem "NSX250"
            .AddItem "NSX630"
            ElseIf MenuUnidades = "Opcion2" Then
            .AddItem "EZC100"
            .AddItem "EZC250"
            .AddItem "CVS100"
            .AddItem "CVS250"
            .AddItem "CVS630"
            .AddItem "NSX250"
            .AddItem "NSX630"
            .AddItem "NS800_1600"
            .AddItem "NWFIJO"
            .AddItem "NWEXTRAIBLE"
            .AddItem "MTZFIJO"
            .AddItem "MTZEXTRAIBLE"
            .AddItem "MVSFIJO"
            .AddItem "MVSEXTRAIBLE"
            End If
            case "SFA - Fijo Atornillado"
        If MenuUnidades= "Opcion1" Then
        .AddItem "EZC100"
        .AddItem "EZC250"
        .AddItem "CVS100"
        .AddItem "CVS250"
        .AddItem "CVS630"
        .AddItem "NSX250"
        .AddItem "NSX630" 
        ElseIf MenuUnidades = "Opcion2" Then
        .AddItem "NS800_1600" 
        End If    
        End Select
        If .ListCount = 1 Then .ListIndex = 0
        End With
    End Sub
    '------------------------------------------------------------ MEDIDA (mm)
        Private Sub LlenarMedida()
            Dim sTipo As String, sFrame As String
            sTipo = Txt("TipoUF")
            sFrame = Txt("FrameUF")
            With Cbo("MedidaUF")
                .Clear
                If sFrame = "" Then Exit Sub
                Select Case sTipo
                    Case "SFI - Frontal Interno", "SFC - Frontal Compacto"
                        If Admite400(sFrame) Then
                            .AddItem "400"
                            .AddItem "720"
                        ElseIf sFrame = "NS800_1250" Then
                            .AddItem "400"
                        Else
                            .AddItem "720"
                        End If
                    Case "SIA - Integral Abatible", "SFA - Fijo Atornillado"
                        .AddItem "720"                 'siempre 720
                End Select
                If .ListCount = 1 Then .ListIndex = 0
            End With
        End Sub
    ' Unico lugar donde se define que frames caben en una UF de 400 mm.
    Private Function Admite400(ByVal sFrame As String) As Boolean
        Select Case sFrame
            Case "MiniBreaker", _
                "EZC100", "EZC250", "EZC630", _
                "CVS100", "CVS250", "CVS630", _
                "NSX100", "NSX250", "NSX630", _
                "GoPact150", "GoPact200", "GoPact250", "GoPact400"',"NS800_1250"
                Admite400 = True
            Case Else
                'GoPact800, NS800_1600, NW / MTZ / MVS -> solo 720
                Admite400 = False
        End Select
    End Function
    '------------------------------------------------------------- CORRIENTE
    Private Sub LlenarCorriente()
        Dim sFrame As String
        sFrame = Txt("FrameUF")

        With Cbo("CorrienteUF")
            .Clear
            Select Case sFrame

                Case "MiniBreaker"
                    .AddItem "1-63A"

                Case "EZC100", "CVS100", "NSX100"
                    .AddItem "100A"

                Case "GoPact150"
                    .AddItem "100A"
                    .AddItem "125A"
                    .AddItem "150A"

                Case "GoPact200"
                    .AddItem "200A"

                Case "EZC250", "CVS250", "NSX250", "GoPact250"
                    .AddItem "250A"

                Case "GoPact400"
                    .AddItem "400A"

                Case "EZC630", "CVS630", "NSX630"
                    .AddItem "400A"
                    .AddItem "630A"

                Case "GoPact800"
                    .AddItem "630A"
                    .AddItem "800A"

                Case "NS800_1250"
                    .AddItem "800A"
                    .AddItem "1250A"

                Case "NS800_1600"                  'CORREGIDO: antes compartia lista con GoPact150
                    .AddItem "800A"
                    .AddItem "1250A"
                    .AddItem "1600A"

                Case "NWFIJO", "NWEXTRAIBLE", "MTZFIJO", "MTZEXTRAIBLE", _
                    "MVSFIJO", "MVSEXTRAIBLE"
                    .AddItem "800A"
                    .AddItem "1250A"
                    .AddItem "1600A"
                    .AddItem "3200A"
                    .AddItem "4000A"
            End Select

            If .ListCount = 1 Then .ListIndex = 0
        End With
    End Sub
    '----------------------------------------------------------- ORIENTACION
    Private Sub LlenarOrientacion()
        Dim sTipo As String, sFrame As String, sMedida As String

        sTipo = Txt("TipoUF")
        sFrame = Txt("FrameUF")
        sMedida = Txt("MedidaUF")

        With Cbo("OrientacionUF")
            .Clear
            If sFrame = "" Or sMedida = "" Then Exit Sub

            '--- REGLA 1: manda la medida. En 400 mm solo hay montaje vertical.
            If sMedida = "400" Then

                .AddItem "Vertical"

                'Si en 400 mm algun frame debe conservar su nombre propio de vertical
                '(por ejemplo "V.Compacta" en CVS100 / NSX100), reemplace la linea
                'anterior por:
                '
                'Select Case sFrame
                '    Case "CVS100", "NSX100": .AddItem "V.Compacta"
                '    Case Else:               .AddItem "Vertical"
                'End Select

            Else
                '--- REGLA 2: en 720 mm depende del tipo de UF y del frame.
                Select Case sTipo
                    Case "SIA - Integral Abatible"
                        Select Case sFrame
                            Case "EZC100", "EZC250", "CVS100", "CVS250", "NSX250"
                                .AddItem "Horizontal"
                            Case "CVS630", "NSX630", "NS800_1600"
                                .AddItem "Horizontal"
                                .AddItem "V.Extendida"
                            Case "NWFIJO", "NWEXTRAIBLE", "MTZFIJO", "MTZEXTRAIBLE", _
                                "MVSFIJO", "MVSEXTRAIBLE"
                                .AddItem "Vertical"
                        End Select
                    case "SFA - Fijo Atornillado"
                        .AddItem "Horizontal"
                    Case Else       'SFI - Frontal Interno / SFC - Frontal Compacto
                        Select Case sFrame
                            Case "MiniBreaker", _
                                "NWFIJO", "NWEXTRAIBLE", "MTZFIJO", "MTZEXTRAIBLE", _
                                "MVSFIJO", "MVSEXTRAIBLE"
                                .AddItem "Vertical"
                            Case "CVS100", "NSX100"
                                .AddItem "V.Compacta"
                            Case "GoPact150"                'CORREGIDO: antes decia "GoPact100"
                                .AddItem "Horizontal"
                                .AddItem "H.Compuesta"
                                .AddItem "V.Compacta"
                            Case "EZC100", "EZC250", "CVS250", "NSX250", _
                                "GoPact200", "GoPact250"
                                .AddItem "Horizontal"
                                .AddItem "H.Compuesta"
                                .AddItem "V.Extendida"
                            Case "GoPact400"
                                .AddItem "Horizontal"
                                .AddItem "H.Compuesta"
                            Case "CVS630", "NSX630", "NS800_1250", "NS800_1600"
                                .AddItem "Horizontal"
                                .AddItem "V.Extendida"
                            Case "EZC630", "GoPact800"
                                .AddItem "Horizontal"
                        End Select
                End Select
            End If

            If .ListCount = 1 Then .ListIndex = 0
        End With
    End Sub 
    '-------------------------------------------------------------- MANDO
    ' SFI / SFC -> Toggle      SIA -> Prolongado
    ' (era tu "OPeracion_Mando"; va Private, no Public, y sin evento propio
    '  porque nada depende de MandoUF)

    Private Sub LlenarMando()
        Dim sTipo As String
        sTipo = Txt("TipoUF")

        With Cbo("MandoUF")
            .Clear
            Select Case sTipo
                Case "SFI - Frontal Interno", "SFC - Frontal Compacto"
                    .AddItem "Toggle"
                Case "SIA - Integral Abatible", "SFA - Fijo Atornillado"
                    .AddItem "Prolongado"
            End Select
            If .ListCount = 1 Then .ListIndex = 0
        End With
    End Sub
    '------------------------------------------------------- FUNCIONALIDAD
    ' 400 mm -> solo Carga        720 mm -> Carga y Totalizador
    ' H.Compuesta -> solo Totalizador
    ' Horizontal  -> Carga y Totalizador
    ' Verticales  -> segun la medida (400 mm no admite Totalizador)
    Private Sub LlenarFuncionalidad()
        Dim sMedida As String, sOrient As String
        sMedida = Txt("MedidaUF")
        sOrient = Txt("OrientacionUF")

        With Cbo("FuncionalidadUF")
            .Clear
            If sMedida = "" Then Exit Sub

            Select Case sOrient

                Case "H.Compuesta"
                    .AddItem "Totalizador"

                Case "Horizontal"
                    if Me.OleObjects("TipoUF").Object.Value = "SFA - Fijo Atornillado" Then
                        .AddItem "Carga"
                    elseif Me.OleObjects("TipoUF").Object.Value = "SIA - Integral Abatible" Then
                        .AddItem "Totalizador"
                    else 
                        .AddItem "Carga"
                        .AddItem "Totalizador"
                    end if    
                Case Else   'Vertical, V.Extendida, V.Compacta o aun sin elegir
                    Select Case sMedida
                        Case "400"
                            .AddItem "Carga"
                        Case "720"
                            .AddItem "Carga"
                            .AddItem "Totalizador"
                    End Select

            End Select

            If .ListCount = 1 Then .ListIndex = 0
        End With
    End Sub
    '------------------------------------------------------------ CONEXION
    ' Manda la orientacion:
    '   Vertical / V.Extendida / V.Compacta -> Superior  (igual en 400 y 720
    ' Verticales   -> Superior
    ' H.Compuesta  -> solo Derecha
    ' Horizontal   -> Carga: Derecha e Izquierda / Totalizador: solo Derecha
    Private Sub LlenarConexion()
        Dim sOrient As String, sFunc As String

        sOrient = Txt("OrientacionUF")
        sFunc = Txt("FuncionalidadUF")

        With Cbo("ConexionUF")
            .Clear
            If sOrient = "" Or sFunc = "" Then Exit Sub

            Select Case sOrient

                Case "Vertical", "V.Extendida", "V.Compacta"
                    .AddItem "Superior"

                Case "H.Compuesta"
                    .AddItem "Derecha"

                Case "Horizontal"
                    Select Case sFunc
                        Case "Carga"
                            if Me.OleObjects("TipoUF").Object.Value = "SFA - Fijo Atornillado" Then
                                .AddItem "Derecha"
                            else 
                                .AddItem "Derecha"    
                                .AddItem "Izquierda"
                            end if
                        Case "Totalizador"
                            .AddItem "Izquierda"
                    End Select

            End Select

            If .ListCount = 1 Then .ListIndex = 0
        End With
    End Sub
'UNIDADES FUNCIONALES LISAS
    Private sub TipoLSS_Change()
        dim dato As String
        dato = Me.OleObjects("TipoLSS").Object.Value
        with Me.OleObjects("UnidadesLSS").Object
            .Clear
            Select case dato
                case "SFI - Frontal Interno","SFC - Frontal Compacto"
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
                    .AddItem "11u"
                    .AddItem "12u"
                case "SIA - Integral Abatible" 
                    .AddItem "5u"  
                    .AddItem "6u"
                    .AddItem "7u"
                    .AddItem "8u"
                    .AddItem "9u"
                    .AddItem "10u"
                    .AddItem "11u"
                    .AddItem "12u"
                case "SFA - Fijo Atornillado"
                    .AddItem "2u" 
                    .AddItem "3u"
                    .AddItem "4u"  
                    .AddItem "5u"     
                    .AddItem "6u"
                    .AddItem "7u"
                    .AddItem "8u"
                    .AddItem "9u"
                    .AddItem "10u"
                    .AddItem "11u"
                    .AddItem "12u"
                case "SIC - Placa de Control"
                    .AddItem "5u"     
                    .AddItem "6u"
                    .AddItem "7u"
                    .AddItem "8u"
                    .AddItem "9u"
                    .AddItem "10u"
                    .AddItem "11u"
            end select
        end with
    end sub 
'CONTACTORES
    Private sub EquipoCT_Change()
        dim dato as string
        dato = Me.OleObjects("EquipoCT").Object.Value
        with Me.oleobjects("VistaCT").Object
            .clear
            select case dato
                case "A-Unit (Barra + Mini)"
                    .AddItem "Sin frente muerto"
                case "Contactores"
                    .AddItem "Con frente muerto"
                    .AddItem "Sin frente muerto"
            end select
        end with
        With Me.OleObjects("TamañoCT").Object
            .Clear
            Select Case dato
                Case "A-Unit (Barra + Mini)"
                    .AddItem "4u"
                Case "Contactores"
                    .AddItem "Talla 1-2"
                    .AddItem "Talla 3"
            End Select
        End With
    end sub
'DPS
    Private sub FabriUNDPS_Change() 
        Dim dato As String, dato2 As String
        dato = Me.OleObjects("FabriUNDPS").Object.Value
        dato2 = Me.OleObjects("TipoUNDPS").Object.Value
        With Me.OleObjects("EquipoDPS").Object   
        .Clear
        select case dato
            
            case "HWA"
           
                .AddItem "NSX100"
                .AddItem "EZC100"
                .AddItem "CVS100"
            case "Ciprotec" 
                .AddItem "MiniBreaker"
                .AddItem "NSX100"
                .AddItem "EZC100"
                .AddItem "CVS100"   
            case else   
                if dato2 = "SFI - Frontal Interno" then 
                    .AddItem "MiniBreaker"
                    .AddItem "NSX100"
                    .AddItem "EZC100"
                    .AddItem "CVS100"
                elseif dato2 = "SIA - Integral Abatible" then
                    .AddItem "NSX100"
                    .AddItem "EZC100"
                    .AddItem "CVS100"
                end if
            end select
        End With
        With Me.OleObjects("AnchoDPS").Object
            .Clear
            Select Case dato
                Case "HWA"
                    .AddItem "720"
                case else
                    .AddItem "400"   
                    .AddItem "720"
            End Select
        End With 
    End Sub  
    Private sub TipoUNDPS_Change()  
        Dim dato As String
        dato = Me.OleObjects("TipoUNDPS").Object.Value
        With Me.OleObjects("MandoDPS").Object
            .Clear
            select case dato
                case "SFI - Frontal Interno","SFC - Frontal Compacto"
                    .AddItem "Toggle"
                case "SIA - Integral Abatible"
                    .AddItem "Prolongado"
            end select
        End With
    End Sub
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
'Servicios Auxiliares
    Public sub  Habilitar_combos_Change()
        Dim dato As String
        dato = Me.OleObjects("MontajeSV").Object.Value  
        select case dato
            case "A-Unit"
                Me.OleObjects("AnchoSV").Object.Enabled = False
                Me.OleObjects("UnidadesSV").Object.Enabled = False
            case "Cubiculo"
                Me.OleObjects("AnchoSV").Object.Enabled = True
                Me.OleObjects("UnidadesSV").Object.Enabled = True
        end select
    End sub  
    Private sub MontajeSV_Change()
        Habilitar_combos_Change
        Dim dato As String
        dato = Me.OleObjects("MontajeSV").Object.Value
        With Me.OleObjects("EquipoSV").Object
            .Clear
            select case dato
                case "A-Unit"
                .AddItem "Iluminación"
                .AddItem "Cambio Tension"
                .AddItem "Batery.ST"
                .AddItem "Pont I/O"
                .AddItem "Switch"
                .AddItem "GateWay"
                .AddItem "Toma Corriente"
                .AddItem "A_UnitBB_400"
                .AddItem "ATSPREMIUM"
                .AddItem "ATSESTANDAR"
                .AddItem "ATSBASIC"

                case "Cubiculo"
                    .AddItem "BusBar"
                    '.AddItem "BusBar_MINI"
                    '.AddItem "Ventilacion"
                    .AddItem "Transformador"
            end select
        end with    
    End sub 
    Private sub EquipoSV_Change()
    dim dato as string
    dato = Me.OleObjects("EquipoSV").Object.Value
    with Me.oleobjects("UnidadesSV").Object
        .clear
        select case dato
            case "BusBar","BusBar_MINI"
                .AddItem "NA"
            case "Ventilacion"   
                .AddItem "5u" 
                .AddItem "6u"
                .AddItem "7u"
            case "Transformador" 
                .AddItem "8u"   
                .AddItem "12u"
                .AddItem "15u"
            end select
        end with
        with Me.OleObjects("AnchoSV").Object
            .Clear
            Select case dato
                case "BusBar"
                    .AddItem "400"
                case "BusBar_MINI"
                    .AddItem "720"
                case else 
                    .AddItem "400"   
                    .AddItem "720"
            end select 
        end with 
    end sub   
'Placas de Control
    Private sub VistaPC_Change()
        Dim dato1 As String, dato2 As String
        dato1 = Me.OleObjects("VistaPC").Object.Value
        dato2 = Me.OleObjects("TamanoPC").Object.Value
        With Me.OleObjects("EquipoPC").Object
            .Clear
            select case dato2
                case "Pequeña" 
                    if dato1 = "Lateral" then
                        .AddItem "Liso" 
                    elseif dato1 = "Frontal" then
                        .AddItem "Liso"
                        .AddItem "PM"
                        '.AddItem "ATS PREMIUM"
                        '.AddItem "ATS ESTANDAR"
                        '.AddItem "ATS BASICA"
                    end if 
                case "Grande"  
                    if dato1 = "Lateral" then
                        .AddItem "Liso" 
                    elseif dato1 = "Frontal" then  
                        .AddItem "Liso"
                        .AddItem "ATS PREMIUM"
                        .AddItem "ATS ESTANDAR"
                        .AddItem "ATS BASICA"
                    end if
                Case "Mediana"
                    if dato1 = "Lateral" then   
                        .AddItem "Liso" 
                    elseif dato1 = "Frontal" then  
                        .AddItem "Liso" 
                        .AddItem "ATL600"
                        .AddItem "DCRL5"
                        .AddItem "Rapidus"
                    end if 
            end select
        end with
    End sub 
    private sub  EquipoPC_Change() 
        Dim dato As String, dato2 As String
        dato = Me.OleObjects("EquipoPC").Object.Value 
        with Me.oleobjects("DisponibilidadPC").Object
            .clear
            select case dato
                case "Liso","PM"
                    .AddItem "NA"
                case "ATS PREMIUM","ATS ESTANDAR","ATS BASICA"
                    'IF dato2 = "Pequeña" then
                       ' .AddItem "Interna"
                   ' elseif dato2 = "Grande" then
                        .AddItem "Externa"
                   ' end if
                case "DCRL5"
                    .AddItem "5 Pasos" 
                case "ATL600","Rapidus"
                    .AddItem "6 Pasos"
                    .AddItem "7 Pasos"
                    .AddItem "8 Pasos"
                    .AddItem "9 Pasos"
                    .AddItem "10 Pasos"
                    .AddItem "11 Pasos"
                    .AddItem "12 Pasos"
                    .AddItem "13 Pasos"
                    .AddItem "14 Pasos"
            end select
        end with
    end sub    
'Compensacion
    Private sub PresentacionCOM_Change()
        dim dato as string
        dato = Me.OleObjects("PresentacionCOM").Object.Value
        with Me.OleObjects("EquipoCOM").Object
            .Clear
            select case dato
                case "Cubiculo"
                    .AddItem "EZC100_Contactor"
                    .AddItem "CVS250_Contactor"
                    .AddItem "NSX250_Contactor"
                    .AddItem "GoPact250_Contactor"
                    .AddItem "GoPack100_Contactor"
                    .additem "Botellas_Grandes"
                    .additem "Botellas_Medianas"
                    .additem "Botellas_Pequenas"
                case "Columna" 
                    .AddItem "EZC630"
                    .AddItem "CVS630"
                    .AddItem "NSX630"
                end select
        end with
    end sub
    Private sub EquipoCOM_Change()
        Dim dato As String
        dato = Me.OleObjects("EquipoCOM").Object.Value
        With Me.OleObjects("MedidaCOM").Object 
            .Clear
            select case dato   
                case "Botellas_Grandes","Botellas_Medianas","Botellas_Pequenas"
                    .AddItem "720" 
                    .AddItem "400"
                case "GoPact250_Contactor","GoPack100_Contactor","NSX250_Contactor","CVS250_Contactor","EZC100_Contactor"
                    .AddItem "720" 
                case "EZC630" ,"CVS630","NSX630"
                    .AddItem "1800"  
                    .AddItem "2000" 
            end select
        end with
    end sub
    private sub MedidaCOM_Change()   
        Dim dato As String, dato2 As String
        dato = Me.OleObjects("MedidaCOM").Object.Value
        dato2 = Me.OleObjects("EquipoCOM").Object.Value 
        with Me.OleObjects("UnidadCOM").Object
            .Clear
            select case dato
                case "720"
                    Select case dato2 
                        case"Botellas_Grandes" 
                            .AddItem "12"
                        case "Botellas_Medianas","CVS250_Contactor","EZC100_Contactor" ,"GoPact250_Contactor","GoPack100_Contactor","NSX250_Contactor"
                            .AddItem "10"   
                        case "Botellas_Pequenas" 
                            .AddItem "8"
                    end select 
                case "400"     
                    Select case dato2 
                        case"Botellas_Grandes" 
                            .AddItem "12"
                        case "Botellas_Medianas"
                            .AddItem "10"   
                        case "Botellas_Pequenas" 
                            .AddItem "8"
                    end select  
                Case "1800","2000"
                    .AddItem "400"  
            end select
        end with
    end sub  
'BARRAJE
    Private sub OrientacionBRR_Change()
        Dim dato As String
        dato = Me.OleObjects("OrientacionBRR").Object.Value
        With Me.OleObjects("AmperajeBRR").Object
            .Clear
            select case dato
                case "Horizontal Bus Bar"
                    .AddItem "400A"
                    .AddItem "630_1250A"
                    .AddItem "1600_2500A"
                    .AddItem "3200A"
                    .AddItem "4000A"
                case "Horizontal Tierra Neutro"   
                    .AddItem "NA" 
                case "Vertical Bus Bar"
                    .AddItem "1250A"   
                    .AddItem "1600A"
            end select
        end with
    end sub 
    Private sub AmperajeBRR_Change()
        Dim dato As String,dato2 As String
        dato = Me.OleObjects("OrientacionBRR").Object.Value
        dato2 = Me.OleObjects("AmperajeBRR").Object.Value

        With Me.OleObjects("MedidaBRR").Object
            .Clear
            select case dato
                case "Horizontal Bus Bar"
                    Select Case dato2
                        Case "400A"
                            .AddItem "400"
                        Case "630_1250A","1600_2500A"
                            .AddItem "400"
                            .AddItem "720"
                        Case "3200A","4000A"
                            .AddItem "720"
                    End Select
                case "Horizontal Tierra Neutro"   
                    .AddItem "720"
                Case "Vertical Bus Bar"
                    select case dato2
                        case "1250A"
                            .AddItem "1600"
                            .AddItem "1800"
                            .AddItem "2000"
                            .AddItem "2200"
                        case "1600A"
                            .AddItem "1800"
                            .AddItem "2000"
                            .AddItem "2200"
                    end select
            end select
        end with
    end sub
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
                case "CVS250","CVS630","NSX250","NSX630"
                    .AddItem "Manual"
                    .AddItem "Motorizado"
                CASE "EZC250","EZC630","GoPact250","GoPact400"
                    .AddItem "Manual"
                case "NS800"
                    .AddItem "Motorizado"
                Case "EZC250X2","CVS250X2","NSX250X2","GoPact250X2"
                    .AddItem "Manual"    
            end select
        end with
    end sub
   Private sub MandoTFF_Change()
        Dim dato As String
        dato = Me.OleObjects("OrientacionTRF").Object.Value
        With Me.OleObjects("ConexionTRF").Object
            .Clear
                SELECT CASE dato
                    CASE "Vertical"
                        .AddItem "Superior"
                    CASE "Horizontal"   
                        .AddItem "NA"
            end select
        end with
    end sub
    
'Accesorio MTM
    Private sub TipoACC_Change()
        Dim dato As String
        dato = Me.OleObjects("TipoACC").Object.Value
        with Me.OleObjects("MedidaACC").Object
            .Clear
            select case dato
                case "Techo_Frontal","Hablador_Frontal"
                    .AddItem "200"
                    .AddItem "400"
                    .AddItem "720"
                case "Techo Lateral"    
                    .AddItem "400"
                    .AddItem "600"
                case "Cancamo_Frontal", "Cancamo_Lateral","Celosidas_Grande","Celosidas_Pequeña", _
                        "Chapa_Antivandalica","Hablador_Lateral","Angulo_Frontal","Angulo Lateral", _
                        "Cubierta Frontal"
                    .AddItem "NA" 
                case else
                    .AddItem "No se reconoce"  
                end select
        end with
    end sub      
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
Public Sub Worksheet_Activate()
    Extraer_datos
    'ocultar menus emergentes
        ActiveSheet.Shapes("Menu Envolventes").Visible = False
        ActiveSheet.Shapes("Menu UnidadFuncional").Visible = False 
        ActiveSheet.Shapes("Menu Lisas").Visible = False
        ActiveSheet.Shapes("Menu Contactores").Visible = False
        ActiveSheet.Shapes("Menu DPS").Visible = False
        ActiveSheet.Shapes("Menu Equipo CTR").Visible = False
        ActiveSheet.Shapes("Menu ACCESORIOS CTR").Visible = False
        ActiveSheet.Shapes("Menu Equipos").Visible = False
        ActiveSheet.Shapes("Menu Equipos").Visible = False
        ActiveSheet.Shapes("Menu Servicios").Visible = False
        ActiveSheet.Shapes("Menu Placas de control").Visible = False
        ActiveSheet.Shapes("Menu Compensacion").Visible = False
        ActiveSheet.Shapes("Menu Barraje").Visible = False
        'ActiveSheet.Shapes("Menu Transferencias").Visible = False
        ActiveSheet.Shapes("Menu AccesoriosMTM").Visible = False
        ActiveSheet.Shapes("Menu Etiqueta").Visible = False
        activeSheet.Shapes("Menu Acotar").Visible = false
        ActiveSheet.Shapes("Menu Color").Visible = False
    
    '============== Envolventes ============================
    '============== Unidades Funcionales ===================
        With  Me.OleObjects("TipoUF").Object
            .Clear
            select case IP
                Case "30"
                .AddItem "SFI - Frontal Interno"
                .AddItem "SFC - Frontal Compacto"
                .AddItem "SIA - Integral Abatible"
                .AddItem "SFA - Fijo Atornillado"
            case else
                .AddItem "SFI - Frontal Interno"
                .AddItem "SIA - Integral Abatible" 
                .AddItem "SFA - Fijo Atornillado"
            end select  
        End With
        
    '============== Unidades Funcionales Lisas ===================
        With  Me.OleObjects("TipoLSS").Object
            .Clear
            if IP = "30" then
                .AddItem "SFI - Frontal Interno"
                .AddItem "SIA - Integral Abatible"
                .AddItem "SFC - Frontal Compacto"
                .AddItem "SIC - Placa de Control"
            else
                .AddItem "SFI - Frontal Interno"
                .AddItem "SIA - Integral Abatible"
                .AddItem "SFA - Fijo Atornillado"
                .AddItem "SIC - Placa de Control"
            end if
        End With
        'With  Me.OleObjects("AnchoLSS").Object
          ' .Clear
           '.AddItem "400"
           ' .AddItem "720"
        
        'End With
    '============== Contactores ===================
        with Me.OleObjects("EquipoCT").Object
            .Clear
            '.AddItem "A-Unit (Barra + Mini)"
            .AddItem "Contactores"
        End With
    '============== DPS ===================
        With Me.OleObjects("TipoUNDPS").Object
            .Clear
            if IP = "30" then
                .AddItem "SFI - Frontal Interno"
                .AddItem "SIA - Integral Abatible"
                .AddItem "SFC - Frontal Compacto"
            else
                .AddItem "SFI - Frontal Interno"
                .AddItem "SIA - Integral Abatible"
            end if
        End With
        With Me.OleObjects("FabriUNDPS").Object
            .Clear
            .AddItem "Citel"
            .AddItem "HWA"
            .AddItem "Ciprotec"
            .AddItem "Schneider Electric"
            .AddItem "Phoenix Contac"
        End With
    '============== Unidades Funcionales de Control ===================
        Me.OleObjects("TipoCTR").Object.Value = "SFI - Frontal Interno"
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
    '============== MENU ETOS =================
        With Me.OleObjects("TipoU_ETO").Object   
            .Clear
            .AddItem "SFI - Frontal Interno"
            .AddItem "SFC - Frontal Compacto"
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
    '============== EQUIPOS INDIVIDUALES ===================
        With Me.OleObjects("EquipoIND").Object
            .Clear
            .AddItem "Contactor"
            .AddItem "Medio_Contactor"
            .AddItem "Amperimetro"
            .AddItem "Voltimetro"
            .AddItem "TH110"
            .AddItem "Transformador"
        End With
    '============== SERVICIOS AUXILIARES ===================
        With Me.OleObjects("MontajeSV").Object
            .Clear
            .AddItem "A-Unit"
            '.AddItem "Cubiculo"
        End With
    '============== PLACAS DE CONTROL ===================  
        With Me.OleObjects("TamanoPC").Object
            .Clear
            .AddItem "Pequeña"
            '.AddItem "Mediana"
            .AddItem "Grande"
        End With
        With Me.OleObjects("VistaPC").Object  
            .Clear
            .AddItem "Frontal"
            .AddItem "Lateral"
        End With
    '============== COMPENSACION ==========================    
        Me.OleObjects("TipoCOM").Object.Value = "SFI - Frontal Interno"
        With Me.OleObjects("PresentacionCOM").Object
            .Clear
            .AddItem "Cubiculo"
            .AddItem "Columna"
        End With
    '============== BARRAJE =============================
        with Me.OleObjects("OrientacionBRR").Object
            .Clear
            .AddItem "Horizontal Bus Bar"
            .AddItem "Horizontal Tierra Neutro"
            .AddItem "Vertical Bus Bar"
        End With
    '============== TRANSFERENCIA =============================
        with Me.OleObjects("TipoUnidadTRF").Object
            .Clear
            .AddItem "SFI - Frontal Interno"
        End With
        with Me.OleObjects("OrientacionTRF").Object
            .Clear
            .AddItem "Horizontal"
            .AddItem "Vertical"
        End With
        With Me.OleObjects("AnchoTFR").Object
            .Clear
            .AddItem "720"
        End With
    '============== Accesorios MTM  =============================
        with Me.OleObjects("TipoACC").Object
            .Clear
            .AddItem "Techo_Frontal"
            .AddItem "Techo_Laterial"
            .AddItem "Cancamo_Frontal"
            .AddItem "Cancamo_Lateral"
            .AddItem "Hablador_Frontal"
            .AddItem "Hablador_Lateral"
            .AddItem "Angulo_Frontal"
            .AddItem "Angulo_Lateral"
            .AddItem "Cubierta_Frontal"
            .AddItem "Cubierta_Lateral"
            .AddItem "Ventilador_Frontal"
            .AddItem "Ventilador_Superior"
            .AddItem "Ventilador_Lateral"
            .AddItem "Celosidas_Pequeña"
            .AddItem "Celosidas_Grande"
            .AddItem "Chapa_Antivandalica"
        End With
    '============== ETIQUETA  =============================    
        With Me.OleObjects("ClaseETQ").Object
            .Clear
            .AddItem "Etiqueta"
            .AddItem "Linea de Producto"
            .AddItem "Logo de clientes"
        End With
    '============== COTAS  =============================
        With Me.OleObjects("CotasDIBUJO").Object
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
END Sub
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

    LlenarCombo Me.OLEObjects("AlturaENV").Object, Altura
    LlenarCombo Me.OLEObjects("AnchoENV").Object, Ancho
    LlenarCombo Me.OLEObjects("AnchoLSS").Object, AnchoLISAS
    'msgbox MenuUnidades
    
End Sub
Public sub Envolventes()
    '========================
        celda = "NS82"
        thisWorkbook.Sheets("DISENO").Range("NP82:QJ182").Clear
        Set wsActual = thisWorkbook.Sheets("DISENO")
    '========================
    dim dato1 as string, dato2 as string
    dato1 = Me.OleObjects("AlturaENV").Object.Value
    dato2 = Me.OleObjects("AnchoENV").Object.Value
    ResultadoENV= "EFM_" & dato1 & "X" & dato2
    MoverBloque ResultadoENV,"ENVOLVENTES", wsActual, celda
End sub
Public sub UnidadesFuncionales()
    '========================
        celda = "NS132"
        thisWorkbook.Sheets("DISENO").Range("NP82:QJ182").Clear
        Set wsActual = thisWorkbook.Sheets("DISENO")
    '========================
    Dim dato1 As String, dato2 As String, dato3 As String, dato4 As String
    dim dato5 as string, dato6 as string

    dato1 = Me.OleObjects("TipoUF").Object.Value
    select case dato1
        case "SFI - Frontal Interno": dato1 = "SFI_"
        Case "SFC - Frontal Compacto": dato1 = "SFC_"
        Case "SIA - Integral Abatible": dato1 = "SIA_"
        case "SFA - Fijo Atornillado": dato1 = "SFA_"
    end select
    dato2 = Me.OleObjects("FrameUF").Object.Value   
    dato3 = Me.OleObjects("MedidaUF").Object.Value
    dato4 = Me.OleObjects("OrientacionUF").Object.Value
    select case dato4
        case "Horizontal": dato4 = "H"
        case "H.Compuesta": dato4 = "HC"
        case "Vertical": dato4 = "V"
        case "V.Extendida": dato4 = "VE"
        case "V.Compuesta": dato4 = "VC"
    end select 
    dato5 = Me.OleObjects("ConexionUF").Object.Value
    select case dato5
        case "Superior": dato5 = "_NA"
        case "Izquierda": dato5 = "_IZ"
        case "Derecha": dato5 = "_DR"
    END SELECT
    dato6 = Me.OleObjects("MandoUF").Object.Value   
    Select case dato6
        case "Toggle": dato6 = "To" 
        case "Prolongado": dato6 = "P"
    End Select

    ResultadoUF= dato1 & dato2 & dato6 & "X" & dato3 & dato4 & dato5
    msgbox ResultadoUF
    MoverBloque ResultadoUF,"UF-SFI", wsActual, celda
end sub    
Public sub Unidades_funcionales_lisas()
    '========================
        celda = "NS132"
        thisWorkbook.Sheets("DISENO").Range("NP82:QJ182").Clear
        Set wsActual = thisWorkbook.Sheets("DISENO")
    '========================
    dim dato1 as string, dato2 as string, dato3 as string
    dato1 = Me.OleObjects("TipoLSS").Object.Value
    select case dato1
        case "SFI - Frontal Interno": dato1 = "SFI"
        Case "SFC - Frontal Compacto": dato1 = "SFC"
        Case "SIA - Integral Abatible", "SFA - Fijo Atornillado": dato1 = "SFA"
        case "SIC - Placa de Control": dato1 = "SIC_"
    End select
    dato3 = Me.OleObjects("UnidadesLSS").Object.Value
    dato2 = Me.OleObjects("AnchoLSS").Object.Value
    ResultadoLSS= dato1 & dato2 & "X" & dato3
    MoverBloque ResultadoLSS,"LISAS", wsActual, celda
End sub
Public Sub Contactores()
    '========================
        celda = "NS132"
        thisWorkbook.Sheets("DISENO").Range("NP82:QJ182").Clear
        Set wsActual = thisWorkbook.Sheets("DISENO")
    '========================

    Select Case CStr(Me.OLEObjects("EquipoCT").Object.Value & "")
        Case "A-Unit (Barra + Mini)": dato1 = "MBXAUNIT"
        Case "Contactores":           dato1 = "Cont"
    End Select

    Select Case CStr(Me.OLEObjects("TamañoCT").Object.Value & "")
        Case "Talla 1-2": dato2 = "T1_2"
        Case "Talla 3":   dato2 = "T1_3"
        Case "4u":        dato2 = "4u"      ' <-- faltaba
    End Select

    Select Case CStr(Me.OLEObjects("VistaCT").Object.Value & "")
        Case "Con frente muerto": dato3 = "FM"
        Case "Sin frente muerto": dato3 = "SFM"
    End Select

    If dato1 = "" Or dato2 = "" Or dato3 = "" Then
        MsgBox "Faltan selecciones en Equipo, Tamaño o Vista.", vbExclamation
        Exit Sub
    End If

    ResultadoCT = "UN_" & dato1 & "_" & dato3 & "X" & dato2
    MoverBloque ResultadoCT, "Ufuncionales", wsActual, celda
End Sub
Public sub DPS()
    '========================
        celda = "NS132"
        thisWorkbook.Sheets("DISENO").Range("NP82:QJ182").Clear
        Set wsActual = thisWorkbook.Sheets("DISENO")
    '========================
    dim dato1 as string, dato2 as string, dato3 as string, dato4 as string, dato5 as string
    dato1 = Me.OleObjects("TipoUNDPS").Object.Value
    Select case dato1
        case "SFI - Frontal Interno": dato1 = "SFC_D"
        case "SFC - Frontal Compacto": dato1 = "SFC_D"
        Case "SIA - Integral Abatible": dato1 = "SIA_D"
    End Select
    dato2 = Me.OleObjects("FabriUNDPS").Object.Value
    Select case dato2
        case "Citel": dato2 = "CI"
        case "HWA": dato2 = "HW"
        case "Ciprotec": dato2 = "CP"
        case "Schneider Electric": dato2 = "SE"
        case "Phoenix Contac": dato2 = "PC"
    End Select
    dato3 = Me.OleObjects("AnchoDPS").Object.Value
    dato4 = Me.OleObjects("EquipoDPS").Object.Value
    dato5 = Me.OleObjects("MandoDPS").Object.Value
    Select case dato5
        case "Toggle": dato5 = "To"
        case "Prolongado": dato5 = "P"
    End Select

    ResultadoDPS= dato1 & dato2 & "X" & dato3 & "_" & dato4 &  dato5
    MoverBloque ResultadoDPS,"DPS", wsActual, celda
End sub
Public sub Equipos_Control()
    '========================
        celda = "NS132"
        thisWorkbook.Sheets("DISENO").Range("NP82:QJ182").Clear
        Set wsActual = thisWorkbook.Sheets("DISENO")
    '========================
    dim dato1 as string, dato2 as string
    dato1 = Me.OleObjects("EquipoCTR").Object.Value
    dato2 = Me.OleObjects("MedidasCTR").Object.Value
    ResultadoCT = "SFI_" &dato1 & "_" & dato2
    MoverBloque ResultadoCT,"UCONTROL", wsActual, celda
End sub
Public sub Accesorios_CTR()
    '========================
        celda = "NS132"
        thisWorkbook.Sheets("DISENO").Range("NP82:QJ182").Clear
        Set wsActual = thisWorkbook.Sheets("DISENO")
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

    ResultadoACC = "AC_" & dato1 & "_" & dato2
    MoverBloque ResultadoACC,"UCONTROL", wsActual, celda
End sub
Public sub Unidades_ETO()
    '========================
        celda = "NS132"
        thisWorkbook.Sheets("DISENO").Range("NP82:QJ182").Clear
        Set wsActual = thisWorkbook.Sheets("DISENO")
    '========================
    dim dato1 as string, dato2 as string

    dato1 = Me.OleObjects("Ancho_ETO").Object.Value
    dato2 = Me.OleObjects("Unidades_ETO").Object.Value

    ResultadoETO= "ETO" & dato1 & "X" & dato2 
    MoverBloque ResultadoETO,"ETO_SFC_SFI", wsActual, celda
End sub
Public sub Equipo_Individual()
    '========================
        celda = "NS132" 
        thisWorkbook.Sheets("DISENO").Range("NP82:QJ182").Clear
        Set wsActual = thisWorkbook.Sheets("DISENO")
    '========================
    dim dato1 as string
    dato1 = Me.OleObjects("EquipoIND").Object.Value
    ResultadoEQ = "EQ_" & dato1
    MoverBloque ResultadoEQ,"UCONTROL", wsActual, celda
End sub
Public sub Servicios_Auxiliares()
    '========================
        celda = "NS132"
        thisWorkbook.Sheets("DISENO").Range("NP82:QJ182").Clear
        Set wsActual = thisWorkbook.Sheets("DISENO")
    '========================
    dim dato1 as string, dato2 as string, dato3 as string, dato4 as string
    dato1 = Me.OleObjects("MontajeSV").Object.Value
    select case dato1
        case "A-Unit": dato1 = "AU_"
        case "Cubiculo": dato1 = "CB_"
    End select
    dato2 = Me.OleObjects("EquipoSV").Object.Value
    select case dato2
        case "Cambio Tension": dato2 = "CambioTension"
        case "Pont I/O": dato2 = "PontI_O"
        case "Toma Corriente": dato2 = "TomaCorriente"
    end select    
    dato3 = Me.OleObjects("AnchoSV").Object.Value
    dato4 = Me.OleObjects("UnidadesSV").Object.Value
    IF dato1 ="AU_" THEN
        ResultadoSV= dato1 & dato2
    ELSE 
        ResultadoSV= dato1 & dato2 & "X" & dato3 & "_" & dato4
    END IF    
    MoverBloque ResultadoSV,"SERVICIOS", wsActual, celda 
End sub
Public sub Placas_Control()
    '========================
        celda = "NS132"
        thisWorkbook.Sheets("DISENO").Range("NP82:QJ182").Clear
        Set wsActual = thisWorkbook.Sheets("DISENO")
    '========================
    dim dato1 as string, dato2 as string, dato3 as string, dato4 as string
    dato1 = Me.OleObjects("TamanoPC").Object.Value
    select case dato1 
        case "Pequeña": dato1 = "PQ"
        case "Grande": dato1 = "GD"
        case "Mediana": dato1 = "MD"
    end select
    dato2 = Me.OleObjects("VistaPC").Object.Value  
    select case dato2
        case "Frontal": dato2 = "FR"  
        case "Lateral": dato2 = "LT"
    end select
    dato3 = Me.OleObjects("EquipoPC").Object.Value   
    Select case dato3
        case "ATS PREMIUM": dato3 = "ATSPREMIUM" 
        case "ATS ESTANDAR": dato3 = "ATSESTANDAR"
        case "ATS BASICA": dato3 = "ATSBASICA"
    End Select
    dato4 = Me.OleObjects("DisponibilidadPC").Object.Value
    select case dato4
        case "Externa":  dato4 = "EXT"
        case "Interna":  dato4 = "INT"
        case "5 Pasos":  dato4 = "5P"
        case "6 Pasos":  dato4 = "6P"
        case "7 Pasos":  dato4 = "7P"
        case "8 Pasos":  dato4 = "8P"
        case "9 Pasos":  dato4 = "9P"
        case "10 Pasos": dato4 = "10P"
        case "11 Pasos": dato4 = "11P"
        case "12 Pasos": dato4 = "12P"
        casE "13 Pasos": dato4 = "13P"
        case "14 Pasos": dato4 = "14P"
    End select    
    ResultadoPC= "PCTR_X" & dato1 & "X" & dato2 & "_" & dato3 & "X" & dato4    
    MoverBloque ResultadoPC,"SERVICIOS", wsActual, celda
End sub
Public sub Compensacion()
    '========================
        celda = "NU107"
        thisWorkbook.Sheets("DISENO").Range("NP82:QJ182").Clear
        Set wsActual = thisWorkbook.Sheets("DISENO")
    '========================
    Dim dato1 As String, dato2 As String, dato3 As String
    dato1 = Me.OleObjects("EquipoCOM").Object.Value
    dato2 = Me.OleObjects("MedidaCOM").Object.Value
    dato3 = Me.OleObjects("UnidadCOM").Object.Value
    ResultadoCOM = "UNC_" & dato1 & "X" & dato2 & "X" & dato3
    MoverBloque ResultadoCOM,"COMPENSACION", wsActual, celda
End sub
Public sub BARRAS()
    '========================
        celda = "NS104"
        thisWorkbook.Sheets("DISENO").Range("NP82:QJ182").Clear
        Set wsActual = thisWorkbook.Sheets("DISENO")
    '========================
    Dim dato1 As String, dato2 As String, dato3 As String
    dato1 = Me.OleObjects("OrientacionBRR").Object.Value
    Select case dato1
        case "Horizontal Bus Bar": dato1 = "HBB"
        Case "Horizontal Tierra Neutro": dato1 = "HTN"
        case "Vertical Bus Bar": dato1 = "VBB"
    End Select
    dato2 = Me.OleObjects("AmperajeBRR").Object.Value
    dato3 = Me.OleObjects("MedidaBRR").Object.Value
    ResultadoBAR = dato1 & "X" & dato3 & "X" & dato2
    MoverBloque ResultadoBAR,"BARRAJE", wsActual, celda
End sub
Public sub Transferencia()
    '========================
        celda = "NU127"
        thisWorkbook.Sheets("DISENO").Range("NP82:QJ182").Clear
        Set wsActual = thisWorkbook.Sheets("DISENO")
    '========================
    Dim dato1 As String, dato2 As String, dato3 As String, dato4 As String, dato5 As String, dato6 As String
    dato1 = Me.OleObjects("TipoUnidadTRF").Object.Value
    Select case dato1
        case "SFI - Frontal Interno": dato1 = "SFI"
    End Select
    dato2 = Me.OleObjects("OrientacionTRF").Object.Value
    Select case dato2
        case "Horizontal": dato2 = "H"
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
        case "NA": dato5 = "NA"
    end select
    dato6 = Me.OleObjects("EquipoTRF").Object.Value

    ResultadoTRF = dato1 & "_UT_" & dato6 & dato3 & "X" & dato2 & "X" & dato4 & "X" & dato5
    
    MoverBloque ResultadoTRF,"ATS_SFI", wsActual, celda

End sub
Public sub ACESORIOS_MTM()
    '========================
        celda = "NS104"
        thisWorkbook.Sheets("DISENO").Range("NP82:QJ182").Clear
        Set wsActual = thisWorkbook.Sheets("DISENO")
    '========================
    Dim dato1 As String, dato2 As String
    dato1 = Me.OleObjects("TipoACC").Object.Value
    dato2 = Me.OleObjects("MedidaACC").Object.Value
    ResultadoACM = "ACC_" & dato1 & "X" & dato2
    MoverBloque ResultadoACM,"ACCESOSRIOS MTM", wsActual, celda
End sub
Public sub  ETIQUETAS()
    '========================
        celda = "NS104"
        thisWorkbook.Sheets("DISENO").Range("NP82:QJ182").Clear
        Set wsActual = thisWorkbook.Sheets("DISENO")
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
            dato1 = "E_"
            ResultadoETQ = dato1 & dato3 & "X" & dato2

        case "Linea de Producto"
            dato1 = "L_"
            ResultadoETQ = dato1 & dato3 & "X" & dato2 

        case "Logo de clientes": dato1 = "LC_"
            ResultadoETQ = dato1  & dato3
    End Select
    MoverBloque ResultadoETQ,"ETIQUETAS", wsActual, celda
End sub
public sub COTAS()
    '========================
        celda = "NS104"
        thisWorkbook.Sheets("DISENO").Range("NP82:QJ182").Clear
        Set wsActual = thisWorkbook.Sheets("DISENO")
    '========================
    Dim dato1 As String
    dato1 = Me.OleObjects("CotasDIBUJO").Object.Value
    Select case dato1
        case "Ancho: 200mm": dato1 = "C_200"
        case "Ancho: 230mm": dato1 = "C_230"
        case "Ancho: 400mm": dato1 = "C_400"
        case "Ancho: 430mm": dato1 = "C_430"
        case "Ancho: 720mm": dato1 = "C_720"
        case "Ancho: 750mm": dato1 = "C_750"
        case "Ancho: 980mm": dato1 = "C_980"
        case "Ancho: 1180mm": dato1 = "C_1180"
        case "Ancho: 1380mm": dato1 = "C_1380"
        case "Altura: 1600mm": dato1 = "C_1600"
        case "Altura: 1800mm": dato1 = "C_1800"
        case "Altura: 2000mm": dato1 = "C_2000"
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
    valorColor = Me.OLEObjects("Colorselec").Object.Value
    Set ws = ThisWorkbook.Sheets("Vista frontal")
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
    For Each celda In ws.Range("BZ38:MX155")
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