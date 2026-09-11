'========================== DISENO TREN =======================
Public Altura As String
Public Ancho As String, AnchoLISAS As String
Public Profundidad As String
Public IP As String, Color As String
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
    Private Function cbo(ByVal sNombre As String) As Object
        Set cbo = Me.OLEObjects(sNombre).Object
    End Function

    ' El "& """"" evita el error 94 (uso invalido de Null) con el combo vacio.
    Private Function Txt(ByVal sNombre As String) As String
        Txt = Trim$(cbo(sNombre).Value & "")
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

        With cbo("FrameUF")
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
            If MenuUnidades = "Opcion1" Then
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
            Case "SFA - Fijo Atornillado"
        If MenuUnidades = "Opcion1" Then
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
            With cbo("MedidaUF")
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
                "GoPact150", "GoPact200", "GoPact250", "GoPact400" ',"NS800_1250"
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

        With cbo("CorrienteUF")
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

        With cbo("OrientacionUF")
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
                    Case "SFA - Fijo Atornillado"
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

        With cbo("MandoUF")
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

        With cbo("FuncionalidadUF")
            .Clear
            If sMedida = "" Then Exit Sub

            Select Case sOrient

                Case "H.Compuesta"
                    .AddItem "Totalizador"

                Case "Horizontal"
                    If Me.OLEObjects("TipoUF").Object.Value = "SFA - Fijo Atornillado" Then
                        .AddItem "Carga"
                    ElseIf Me.OLEObjects("TipoUF").Object.Value = "SIA - Integral Abatible" Then
                        .AddItem "Totalizador"
                    Else
                        .AddItem "Carga"
                        .AddItem "Totalizador"
                    End If
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

        With cbo("ConexionUF")
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
                            If Me.OLEObjects("TipoUF").Object.Value = "SFA - Fijo Atornillado" Then
                                .AddItem "Derecha"
                            Else
                                .AddItem "Derecha"
                                .AddItem "Izquierda"
                            End If
                        Case "Totalizador"
                            .AddItem "Izquierda"
                    End Select

            End Select

            If .ListCount = 1 Then .ListIndex = 0
        End With
    End Sub
'UNIDADES FUNCIONALES LISAS
    Private Sub TipoLSS_Change()
        Dim dato As String
        dato = Me.OLEObjects("TipoLSS").Object.Value
        With Me.OLEObjects("UnidadesLSS").Object
            .Clear
            Select Case dato
                Case "SFI - Frontal Interno", "SFC - Frontal Compacto"
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
                Case "SIA - Integral Abatible"
                    .AddItem "5u"
                    .AddItem "6u"
                    .AddItem "7u"
                    .AddItem "8u"
                    .AddItem "9u"
                    .AddItem "10u"
                    .AddItem "11u"
                    .AddItem "12u"
                Case "SFA - Fijo Atornillado"
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
                Case "SIC - Placa de Control"
                    .AddItem "5u"
                    .AddItem "6u"
                    .AddItem "7u"
                    .AddItem "8u"
                    .AddItem "9u"
                    .AddItem "10u"
                    .AddItem "11u"
            End Select
        End With
    End Sub
'CONTACTORES
    Private Sub EquipoCT_Change()
        Dim dato As String
        dato = Me.OLEObjects("EquipoCT").Object.Value
        With Me.OLEObjects("VistaCT").Object
            .Clear
            Select Case dato
                Case "A-Unit (Barra + Mini)"
                    .AddItem "Sin frente muerto"
                Case "Contactores"
                    .AddItem "Con frente muerto"
                    .AddItem "Sin frente muerto"
            End Select
        End With
        With Me.OLEObjects("Tama�oCT").Object
            .Clear
            Select Case dato
                Case "A-Unit (Barra + Mini)"
                    .AddItem "4u"
                Case "Contactores"
                    .AddItem "Talla 1-2"
                    .AddItem "Talla 3"
            End Select
        End With
    End Sub
'DPS
    Private Sub FabriUNDPS_Change()
        Dim dato As String, dato2 As String
        dato = Me.OLEObjects("FabriUNDPS").Object.Value
        dato2 = Me.OLEObjects("TipoUNDPS").Object.Value
        With Me.OLEObjects("EquipoDPS").Object
        .Clear
        Select Case dato
            
            Case "HWA"
           
                .AddItem "NSX100"
                .AddItem "EZC100"
                .AddItem "CVS100"
            Case "Ciprotec"
                .AddItem "MiniBreaker"
                .AddItem "NSX100"
                .AddItem "EZC100"
                .AddItem "CVS100"
            Case Else
                If dato2 = "SFI - Frontal Interno" Then
                    .AddItem "MiniBreaker"
                    .AddItem "NSX100"
                    .AddItem "EZC100"
                    .AddItem "CVS100"
                ElseIf dato2 = "SIA - Integral Abatible" Then
                    .AddItem "NSX100"
                    .AddItem "EZC100"
                    .AddItem "CVS100"
                End If
            End Select
        End With
        With Me.OLEObjects("AnchoDPS").Object
            .Clear
            Select Case dato
                Case "HWA"
                    .AddItem "720"
                Case Else
                    .AddItem "400"
                    .AddItem "720"
            End Select
        End With
    End Sub
    Private Sub TipoUNDPS_Change()
        Dim dato As String
        dato = Me.OLEObjects("TipoUNDPS").Object.Value
        With Me.OLEObjects("MandoDPS").Object
            .Clear
            Select Case dato
                Case "SFI - Frontal Interno", "SFC - Frontal Compacto"
                    .AddItem "Toggle"
                Case "SIA - Integral Abatible"
                    .AddItem "Prolongado"
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
'Servicios Auxiliares
    Public Sub Habilitar_combos_Change()
        Dim dato As String
        dato = Me.OLEObjects("MontajeSV").Object.Value
        Select Case dato
            Case "A-Unit"
                Me.OLEObjects("AnchoSV").Object.Enabled = False
                Me.OLEObjects("UnidadesSV").Object.Enabled = False
            Case "Cubiculo"
                Me.OLEObjects("AnchoSV").Object.Enabled = True
                Me.OLEObjects("UnidadesSV").Object.Enabled = True
        End Select
    End Sub
    Private Sub MontajeSV_Change()
        Habilitar_combos_Change
        Dim dato As String
        dato = Me.OLEObjects("MontajeSV").Object.Value
        With Me.OLEObjects("EquipoSV").Object
            .Clear
            Select Case dato
                Case "A-Unit"
                .AddItem "Iluminaci�n"
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

                Case "Cubiculo"
                    .AddItem "BusBar"
                    '.AddItem "BusBar_MINI"
                    '.AddItem "Ventilacion"
                    .AddItem "Transformador"
            End Select
        End With
    End Sub
    Private Sub EquipoSV_Change()
    Dim dato As String
    dato = Me.OLEObjects("EquipoSV").Object.Value
    With Me.OLEObjects("UnidadesSV").Object
        .Clear
        Select Case dato
            Case "BusBar", "BusBar_MINI"
                .AddItem "NA"
            Case "Ventilacion"
                .AddItem "5u"
                .AddItem "6u"
                .AddItem "7u"
            Case "Transformador"
                .AddItem "8u"
                .AddItem "12u"
                .AddItem "15u"
            End Select
        End With
        With Me.OLEObjects("AnchoSV").Object
            .Clear
            Select Case dato
                Case "BusBar"
                    .AddItem "400"
                Case "BusBar_MINI"
                    .AddItem "720"
                Case Else
                    .AddItem "400"
                    .AddItem "720"
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
                Case "Peque�a"
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
'Compensacion
    Private Sub PresentacionCOM_Change()
        Dim dato As String
        dato = Me.OLEObjects("PresentacionCOM").Object.Value
        With Me.OLEObjects("EquipoCOM").Object
            .Clear
            Select Case dato
                Case "Cubiculo"
                    .AddItem "EZC100_Contactor"
                    .AddItem "CVS250_Contactor"
                    .AddItem "NSX250_Contactor"
                    .AddItem "GoPact250_Contactor"
                    .AddItem "GoPack100_Contactor"
                    .AddItem "Botellas_Grandes"
                    .AddItem "Botellas_Medianas"
                    .AddItem "Botellas_Pequenas"
                Case "Columna"
                    .AddItem "EZC630"
                    .AddItem "CVS630"
                    .AddItem "NSX630"
                End Select
        End With
    End Sub
    Private Sub EquipoCOM_Change()
        Dim dato As String
        dato = Me.OLEObjects("EquipoCOM").Object.Value
        With Me.OLEObjects("MedidaCOM").Object
            .Clear
            Select Case dato
                Case "Botellas_Grandes", "Botellas_Medianas", "Botellas_Pequenas"
                    .AddItem "720"
                    .AddItem "400"
                Case "GoPact250_Contactor", "GoPack100_Contactor", "NSX250_Contactor", "CVS250_Contactor", "EZC100_Contactor"
                    .AddItem "720"
                Case "EZC630", "CVS630", "NSX630"
                    .AddItem "1800"
                    .AddItem "2000"
            End Select
        End With
    End Sub
    Private Sub MedidaCOM_Change()
        Dim dato As String, dato2 As String
        dato = Me.OLEObjects("MedidaCOM").Object.Value
        dato2 = Me.OLEObjects("EquipoCOM").Object.Value
        With Me.OLEObjects("UnidadCOM").Object
            .Clear
            Select Case dato
                Case "720"
                    Select Case dato2
                        Case "Botellas_Grandes"
                            .AddItem "12"
                        Case "Botellas_Medianas", "CVS250_Contactor", "EZC100_Contactor", "GoPact250_Contactor", "GoPack100_Contactor", "NSX250_Contactor"
                            .AddItem "10"
                        Case "Botellas_Pequenas"
                            .AddItem "8"
                    End Select
                Case "400"
                    Select Case dato2
                        Case "Botellas_Grandes"
                            .AddItem "12"
                        Case "Botellas_Medianas"
                            .AddItem "10"
                        Case "Botellas_Pequenas"
                            .AddItem "8"
                    End Select
                Case "1800", "2000"
                    .AddItem "400"
            End Select
        End With
    End Sub
'BARRAJE
    Private Sub OrientacionBRR_Change()
        Dim dato As String
        dato = Me.OLEObjects("OrientacionBRR").Object.Value
        With Me.OLEObjects("AmperajeBRR").Object
            .Clear
            Select Case dato
                Case "Horizontal Bus Bar"
                    .AddItem "400A"
                    .AddItem "630_1250A"
                    .AddItem "1600_2500A"
                    .AddItem "3200A"
                    .AddItem "4000A"
                Case "Horizontal Tierra Neutro"
                    .AddItem "NA"
                Case "Vertical Bus Bar"
                    .AddItem "1250A"
                    .AddItem "1600A"
            End Select
        End With
    End Sub
    Private Sub AmperajeBRR_Change()
        Dim dato As String, dato2 As String
        dato = Me.OLEObjects("OrientacionBRR").Object.Value
        dato2 = Me.OLEObjects("AmperajeBRR").Object.Value

        With Me.OLEObjects("MedidaBRR").Object
            .Clear
            Select Case dato
                Case "Horizontal Bus Bar"
                    Select Case dato2
                        Case "400A"
                            .AddItem "400"
                        Case "630_1250A", "1600_2500A"
                            .AddItem "400"
                            .AddItem "720"
                        Case "3200A", "4000A"
                            .AddItem "720"
                    End Select
                Case "Horizontal Tierra Neutro"
                    .AddItem "720"
                Case "Vertical Bus Bar"
                    Select Case dato2
                        Case "1250A"
                            .AddItem "1600"
                            .AddItem "1800"
                            .AddItem "2000"
                            .AddItem "2200"
                        Case "1600A"
                            .AddItem "1800"
                            .AddItem "2000"
                            .AddItem "2200"
                    End Select
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
                Case "CVS250", "CVS630", "NSX250", "NSX630"
                    .AddItem "Manual"
                    .AddItem "Motorizado"
                Case "EZC250", "EZC630", "GoPact250", "GoPact400"
                    .AddItem "Manual"
                Case "NS800"
                    .AddItem "Motorizado"
                Case "EZC250X2", "CVS250X2", "NSX250X2", "GoPact250X2"
                    .AddItem "Manual"
            End Select
        End With
    End Sub
   Private Sub MandoTFF_Change()
        Dim dato As String
        dato = Me.OLEObjects("OrientacionTRF").Object.Value
        With Me.OLEObjects("ConexionTRF").Object
            .Clear
                Select Case dato
                    Case "Vertical"
                        .AddItem "Superior"
                    Case "Horizontal"
                        .AddItem "NA"
            End Select
        End With
    End Sub
    
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
                Case "Cancamo_Frontal", "Cancamo_Lateral", "Celosidas_Grande", "Celosidas_Peque�a", _
                        "Chapa_Antivandalica", "Hablador_Lateral", "Angulo_Frontal", "Angulo Lateral", _
                        "Cubierta Frontal"
                    .AddItem "NA"
                Case Else
                    .AddItem "No se reconoce"
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
                    .AddItem "Ventilaci�n"
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
                Case "Ventilaci�n", "Descarga", "Riesgo_electrico2"
                    .AddItem "Pequeno"
                    .AddItem "Mediano"
                    .AddItem "Grande"
                Case "Placa_GR", "Ventilaci�n", "Descarga", "Riesgo_electrico2", "Riesgo_electricoV1", "Riesgo_electricoV2", _
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
        ActiveSheet.Shapes("Menu Acotar").Visible = False
        ActiveSheet.Shapes("Menu Color").Visible = False
    
    '============== Envolventes ============================
    '============== Unidades Funcionales ===================
        With Me.OLEObjects("TipoUF").Object
            .Clear
            Select Case IP
                Case "30"
                .AddItem "SFI - Frontal Interno"
                .AddItem "SFC - Frontal Compacto"
                .AddItem "SIA - Integral Abatible"
                .AddItem "SFA - Fijo Atornillado"
            Case Else
                .AddItem "SFI - Frontal Interno"
                .AddItem "SIA - Integral Abatible"
                .AddItem "SFA - Fijo Atornillado"
            End Select
        End With
        
    '============== Unidades Funcionales Lisas ===================
        With Me.OLEObjects("TipoLSS").Object
            .Clear
            If IP = "30" Then
                .AddItem "SFI - Frontal Interno"
                .AddItem "SIA - Integral Abatible"
                .AddItem "SFC - Frontal Compacto"
                .AddItem "SIC - Placa de Control"
            Else
                .AddItem "SFI - Frontal Interno"
                .AddItem "SIA - Integral Abatible"
                .AddItem "SFA - Fijo Atornillado"
                .AddItem "SIC - Placa de Control"
            End If
        End With
        'With  Me.OleObjects("AnchoLSS").Object
          ' .Clear
           '.AddItem "400"
           ' .AddItem "720"
        
        'End With
    '============== Contactores ===================
        With Me.OLEObjects("EquipoCT").Object
            .Clear
            '.AddItem "A-Unit (Barra + Mini)"
            .AddItem "Contactores"
        End With
    '============== DPS ===================
        With Me.OLEObjects("TipoUNDPS").Object
            .Clear
            If IP = "30" Then
                .AddItem "SFI - Frontal Interno"
                .AddItem "SIA - Integral Abatible"
                .AddItem "SFC - Frontal Compacto"
            Else
                .AddItem "SFI - Frontal Interno"
                .AddItem "SIA - Integral Abatible"
            End If
        End With
        With Me.OLEObjects("FabriUNDPS").Object
            .Clear
            .AddItem "Citel"
            .AddItem "HWA"
            .AddItem "Ciprotec"
            .AddItem "Schneider Electric"
            .AddItem "Phoenix Contac"
        End With
    '============== Unidades Funcionales de Control ===================
        Me.OLEObjects("TipoCTR").Object.Value = "SFI - Frontal Interno"
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
    '============== MENU ETOS =================
        With Me.OLEObjects("TipoU_ETO").Object
            .Clear
            .AddItem "SFI - Frontal Interno"
            .AddItem "SFC - Frontal Compacto"
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
    '============== EQUIPOS INDIVIDUALES ===================
        With Me.OLEObjects("EquipoIND").Object
            .Clear
            .AddItem "Contactor"
            .AddItem "Medio_Contactor"
            .AddItem "Amperimetro"
            .AddItem "Voltimetro"
            .AddItem "TH110"
            .AddItem "Transformador"
        End With
    '============== SERVICIOS AUXILIARES ===================
        With Me.OLEObjects("MontajeSV").Object
            .Clear
            .AddItem "A-Unit"
            '.AddItem "Cubiculo"
        End With
    '============== PLACAS DE CONTROL ===================
        With Me.OLEObjects("TamanoPC").Object
            .Clear
            .AddItem "Peque�a"
            '.AddItem "Mediana"
            .AddItem "Grande"
        End With
        With Me.OLEObjects("VistaPC").Object
            .Clear
            .AddItem "Frontal"
            .AddItem "Lateral"
        End With
    '============== COMPENSACION ==========================
        Me.OLEObjects("TipoCOM").Object.Value = "SFI - Frontal Interno"
        With Me.OLEObjects("PresentacionCOM").Object
            .Clear
            .AddItem "Cubiculo"
            .AddItem "Columna"
        End With
    '============== BARRAJE =============================
        With Me.OLEObjects("OrientacionBRR").Object
            .Clear
            .AddItem "Horizontal Bus Bar"
            .AddItem "Horizontal Tierra Neutro"
            .AddItem "Vertical Bus Bar"
        End With
    '============== TRANSFERENCIA =============================
        With Me.OLEObjects("TipoUnidadTRF").Object
            .Clear
            .AddItem "SFI - Frontal Interno"
        End With
        With Me.OLEObjects("OrientacionTRF").Object
            .Clear
            .AddItem "Horizontal"
            .AddItem "Vertical"
        End With
        With Me.OLEObjects("AnchoTFR").Object
            .Clear
            .AddItem "720"
        End With
    '============== Accesorios MTM  =============================
        With Me.OLEObjects("TipoACC").Object
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
            .AddItem "Celosidas_Peque�a"
            .AddItem "Celosidas_Grande"
            .AddItem "Chapa_Antivandalica"
        End With
    '============== ETIQUETA  =============================
        With Me.OLEObjects("ClaseETQ").Object
            .Clear
            .AddItem "Etiqueta"
            .AddItem "Linea de Producto"
            .AddItem "Logo de clientes"
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
    LlenarCombo Me.OLEObjects("AnchoLSS").Object, AnchoLISAS
    'msgbox MenuUnidades
    
End Sub
Public Sub Envolventes()
    '========================
        Celda = "NS82"
        ThisWorkbook.Sheets("DISENO").Range("NP82:QJ182").Clear
        Set wsActual = ThisWorkbook.Sheets("DISENO")
    '========================
    Dim dato1 As String, dato2 As String
    dato1 = Me.OLEObjects("AlturaENV").Object.Value
    dato2 = Me.OLEObjects("AnchoENV").Object.Value
    ResultadoENV = "EFM_" & dato1 & "X" & dato2
    MoverBloque ResultadoENV, "ENVOLVENTES", wsActual, Celda
End Sub
Public Sub UnidadesFuncionales()
    '========================
        Celda = "NS132"
        ThisWorkbook.Sheets("DISENO").Range("NP82:QJ182").Clear
        Set wsActual = ThisWorkbook.Sheets("DISENO")
    '========================
    Dim dato1 As String, dato2 As String, dato3 As String, dato4 As String
    Dim dato5 As String, dato6 As String

    dato1 = Me.OLEObjects("TipoUF").Object.Value
    Select Case dato1
        Case "SFI - Frontal Interno": dato1 = "SFI_"
        Case "SFC - Frontal Compacto": dato1 = "SFC_"
        Case "SIA - Integral Abatible": dato1 = "SIA_"
        Case "SFA - Fijo Atornillado": dato1 = "SFA_"
    End Select
    dato2 = Me.OLEObjects("FrameUF").Object.Value
    dato3 = Me.OLEObjects("MedidaUF").Object.Value
    dato4 = Me.OLEObjects("OrientacionUF").Object.Value
    Select Case dato4
        Case "Horizontal": dato4 = "H"
        Case "H.Compuesta": dato4 = "HC"
        Case "Vertical": dato4 = "V"
        Case "V.Extendida": dato4 = "VE"
        Case "V.Compuesta": dato4 = "VC"
    End Select
    dato5 = Me.OLEObjects("ConexionUF").Object.Value
    Select Case dato5
        Case "Superior": dato5 = "_NA"
        Case "Izquierda": dato5 = "_IZ"
        Case "Derecha": dato5 = "_DR"
    End Select
    dato6 = Me.OLEObjects("MandoUF").Object.Value
    Select Case dato6
        Case "Toggle": dato6 = "To"
        Case "Prolongado": dato6 = "P"
    End Select

    ResultadoUF = dato1 & dato2 & dato6 & "X" & dato3 & dato4 & dato5
    MsgBox ResultadoUF
    MoverBloque ResultadoUF, "UF-SFI", wsActual, Celda
End Sub
Public Sub Unidades_funcionales_lisas()
    '========================
        Celda = "NS132"
        ThisWorkbook.Sheets("DISENO").Range("NP82:QJ182").Clear
        Set wsActual = ThisWorkbook.Sheets("DISENO")
    '========================
    Dim dato1 As String, dato2 As String, dato3 As String
    dato1 = Me.OLEObjects("TipoLSS").Object.Value
    Select Case dato1
        Case "SFI - Frontal Interno": dato1 = "SFI"
        Case "SFC - Frontal Compacto": dato1 = "SFC"
        Case "SIA - Integral Abatible", "SFA - Fijo Atornillado": dato1 = "SFA"
        Case "SIC - Placa de Control": dato1 = "SIC_"
    End Select
    dato3 = Me.OLEObjects("UnidadesLSS").Object.Value
    dato2 = Me.OLEObjects("AnchoLSS").Object.Value
    ResultadoLSS = dato1 & dato2 & "X" & dato3
    MoverBloque ResultadoLSS, "LISAS", wsActual, Celda
End Sub
Public Sub Contactores()
    '========================
        Celda = "NS132"
        ThisWorkbook.Sheets("DISENO").Range("NP82:QJ182").Clear
        Set wsActual = ThisWorkbook.Sheets("DISENO")
    '========================

    Select Case CStr(Me.OLEObjects("EquipoCT").Object.Value & "")
        Case "A-Unit (Barra + Mini)": dato1 = "MBXAUNIT"
        Case "Contactores":           dato1 = "Cont"
    End Select

    Select Case CStr(Me.OLEObjects("Tama�oCT").Object.Value & "")
        Case "Talla 1-2": dato2 = "T1_2"
        Case "Talla 3":   dato2 = "T1_3"
        Case "4u":        dato2 = "4u"      ' <-- faltaba
    End Select

    Select Case CStr(Me.OLEObjects("VistaCT").Object.Value & "")
        Case "Con frente muerto": dato3 = "FM"
        Case "Sin frente muerto": dato3 = "SFM"
    End Select

    If dato1 = "" Or dato2 = "" Or dato3 = "" Then
        MsgBox "Faltan selecciones en Equipo, Tama�o o Vista.", vbExclamation
        Exit Sub
    End If

    ResultadoCT = "UN_" & dato1 & "_" & dato3 & "X" & dato2
    MoverBloque ResultadoCT, "Ufuncionales", wsActual, Celda
End Sub
Public Sub DPS()
    '========================
        Celda = "NS132"
        ThisWorkbook.Sheets("DISENO").Range("NP82:QJ182").Clear
        Set wsActual = ThisWorkbook.Sheets("DISENO")
    '========================
    Dim dato1 As String, dato2 As String, dato3 As String, dato4 As String, dato5 As String
    dato1 = Me.OLEObjects("TipoUNDPS").Object.Value
    Select Case dato1
        Case "SFI - Frontal Interno": dato1 = "SFC_D"
        Case "SFC - Frontal Compacto": dato1 = "SFC_D"
        Case "SIA - Integral Abatible": dato1 = "SIA_D"
    End Select
    dato2 = Me.OLEObjects("FabriUNDPS").Object.Value
    Select Case dato2
        Case "Citel": dato2 = "CI"
        Case "HWA": dato2 = "HW"
        Case "Ciprotec": dato2 = "CP"
        Case "Schneider Electric": dato2 = "SE"
        Case "Phoenix Contac": dato2 = "PC"
    End Select
    dato3 = Me.OLEObjects("AnchoDPS").Object.Value
    dato4 = Me.OLEObjects("EquipoDPS").Object.Value
    dato5 = Me.OLEObjects("MandoDPS").Object.Value
    Select Case dato5
        Case "Toggle": dato5 = "To"
        Case "Prolongado": dato5 = "P"
    End Select

    ResultadoDPS = dato1 & dato2 & "X" & dato3 & "_" & dato4 & dato5
    MoverBloque ResultadoDPS, "DPS", wsActual, Celda
End Sub
Public Sub Equipos_Control()
    '========================
        Celda = "NS132"
        ThisWorkbook.Sheets("DISENO").Range("NP82:QJ182").Clear
        Set wsActual = ThisWorkbook.Sheets("DISENO")
    '========================
    Dim dato1 As String, dato2 As String
    dato1 = Me.OLEObjects("EquipoCTR").Object.Value
    dato2 = Me.OLEObjects("MedidasCTR").Object.Value
    ResultadoCT = "SFI_" & dato1 & "_" & dato2
    MoverBloque ResultadoCT, "UCONTROL", wsActual, Celda
End Sub
Public Sub Accesorios_CTR()
    '========================
        Celda = "NS132"
        ThisWorkbook.Sheets("DISENO").Range("NP82:QJ182").Clear
        Set wsActual = ThisWorkbook.Sheets("DISENO")
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

    ResultadoACC = "AC_" & dato1 & "_" & dato2
    MoverBloque ResultadoACC, "UCONTROL", wsActual, Celda
End Sub
Public Sub Unidades_ETO()
    '========================
        Celda = "NS132"
        ThisWorkbook.Sheets("DISENO").Range("NP82:QJ182").Clear
        Set wsActual = ThisWorkbook.Sheets("DISENO")
    '========================
    Dim dato1 As String, dato2 As String

    dato1 = Me.OLEObjects("Ancho_ETO").Object.Value
    dato2 = Me.OLEObjects("Unidades_ETO").Object.Value

    ResultadoETO = "ETO" & dato1 & "X" & dato2
    MoverBloque ResultadoETO, "ETO_SFC_SFI", wsActual, Celda
End Sub
Public Sub Equipo_Individual()
    '========================
        Celda = "NS132"
        ThisWorkbook.Sheets("DISENO").Range("NP82:QJ182").Clear
        Set wsActual = ThisWorkbook.Sheets("DISENO")
    '========================
    Dim dato1 As String
    dato1 = Me.OLEObjects("EquipoIND").Object.Value
    ResultadoEQ = "EQ_" & dato1
    MoverBloque ResultadoEQ, "UCONTROL", wsActual, Celda
End Sub
Public Sub Servicios_Auxiliares()
    '========================
        Celda = "NS132"
        ThisWorkbook.Sheets("DISENO").Range("NP82:QJ182").Clear
        Set wsActual = ThisWorkbook.Sheets("DISENO")
    '========================
    Dim dato1 As String, dato2 As String, dato3 As String, dato4 As String
    dato1 = Me.OLEObjects("MontajeSV").Object.Value
    Select Case dato1
        Case "A-Unit": dato1 = "AU_"
        Case "Cubiculo": dato1 = "CB_"
    End Select
    dato2 = Me.OLEObjects("EquipoSV").Object.Value
    Select Case dato2
        Case "Cambio Tension": dato2 = "CambioTension"
        Case "Pont I/O": dato2 = "PontI_O"
        Case "Toma Corriente": dato2 = "TomaCorriente"
    End Select
    dato3 = Me.OLEObjects("AnchoSV").Object.Value
    dato4 = Me.OLEObjects("UnidadesSV").Object.Value
    If dato1 = "AU_" Then
        ResultadoSV = dato1 & dato2
    Else
        ResultadoSV = dato1 & dato2 & "X" & dato3 & "_" & dato4
    End If
    MoverBloque ResultadoSV, "SERVICIOS", wsActual, Celda
End Sub
Public Sub Placas_Control()
    '========================
        Celda = "NS132"
        ThisWorkbook.Sheets("DISENO").Range("NP82:QJ182").Clear
        Set wsActual = ThisWorkbook.Sheets("DISENO")
    '========================
    Dim dato1 As String, dato2 As String, dato3 As String, dato4 As String
    dato1 = Me.OLEObjects("TamanoPC").Object.Value
    Select Case dato1
        Case "Peque�a": dato1 = "PQ"
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
Public Sub Compensacion()
    '========================
        Celda = "NU107"
        ThisWorkbook.Sheets("DISENO").Range("NP82:QJ182").Clear
        Set wsActual = ThisWorkbook.Sheets("DISENO")
    '========================
    Dim dato1 As String, dato2 As String, dato3 As String
    dato1 = Me.OLEObjects("EquipoCOM").Object.Value
    dato2 = Me.OLEObjects("MedidaCOM").Object.Value
    dato3 = Me.OLEObjects("UnidadCOM").Object.Value
    ResultadoCOM = "UNC_" & dato1 & "X" & dato2 & "X" & dato3
    MoverBloque ResultadoCOM, "COMPENSACION", wsActual, Celda
End Sub
Public Sub BARRAS()
    '========================
        Celda = "NS104"
        ThisWorkbook.Sheets("DISENO").Range("NP82:QJ182").Clear
        Set wsActual = ThisWorkbook.Sheets("DISENO")
    '========================
    Dim dato1 As String, dato2 As String, dato3 As String
    dato1 = Me.OLEObjects("OrientacionBRR").Object.Value
    Select Case dato1
        Case "Horizontal Bus Bar": dato1 = "HBB"
        Case "Horizontal Tierra Neutro": dato1 = "HTN"
        Case "Vertical Bus Bar": dato1 = "VBB"
    End Select
    dato2 = Me.OLEObjects("AmperajeBRR").Object.Value
    dato3 = Me.OLEObjects("MedidaBRR").Object.Value
    ResultadoBAR = dato1 & "X" & dato3 & "X" & dato2
    MoverBloque ResultadoBAR, "BARRAJE", wsActual, Celda
End Sub
Public Sub Transferencia()
    '========================
        Celda = "NU127"
        ThisWorkbook.Sheets("DISENO").Range("NP82:QJ182").Clear
        Set wsActual = ThisWorkbook.Sheets("DISENO")
    '========================
    Dim dato1 As String, dato2 As String, dato3 As String, dato4 As String, dato5 As String, dato6 As String
    dato1 = Me.OLEObjects("TipoUnidadTRF").Object.Value
    Select Case dato1
        Case "SFI - Frontal Interno": dato1 = "SFI"
    End Select
    dato2 = Me.OLEObjects("OrientacionTRF").Object.Value
    Select Case dato2
        Case "Horizontal": dato2 = "H"
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
        Case "NA": dato5 = "NA"
    End Select
    dato6 = Me.OLEObjects("EquipoTRF").Object.Value

    ResultadoTRF = dato1 & "_UT_" & dato6 & dato3 & "X" & dato2 & "X" & dato4 & "X" & dato5
    
    MoverBloque ResultadoTRF, "ATS_SFI", wsActual, Celda

End Sub
Public Sub ACESORIOS_MTM()
    '========================
        Celda = "NS104"
        ThisWorkbook.Sheets("DISENO").Range("NP82:QJ182").Clear
        Set wsActual = ThisWorkbook.Sheets("DISENO")
    '========================
    Dim dato1 As String, dato2 As String
    dato1 = Me.OLEObjects("TipoACC").Object.Value
    dato2 = Me.OLEObjects("MedidaACC").Object.Value
    ResultadoACM = "ACC_" & dato1 & "X" & dato2
    MoverBloque ResultadoACM, "ACCESOSRIOS MTM", wsActual, Celda
End Sub
Public Sub ETIQUETAS()
    '========================
        Celda = "NS104"
        ThisWorkbook.Sheets("DISENO").Range("NP82:QJ182").Clear
        Set wsActual = ThisWorkbook.Sheets("DISENO")
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
Public Sub COTAS()
    '========================
        Celda = "NS104"
        ThisWorkbook.Sheets("DISENO").Range("NP82:QJ182").Clear
        Set wsActual = ThisWorkbook.Sheets("DISENO")
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
    
    ' Obtener el valor del color desde la portada
    valorColor = Me.OLEObjects("Colorselec").Object.Value
    Set ws = ThisWorkbook.Sheets("Vista frontal")
    ' Definir color destino seg�n el valor
    Select Case valorColor
        Case "RAL 7035": colorDestino = RGB(197, 199, 196)
        Case "RAL 7032": colorDestino = RGB(181, 176, 161)
        Case "RAL 9003": colorDestino = RGB(236, 236, 231)
        Case "RAL 9004": colorDestino = RGB(43, 43, 44)
        Case "Naranja": colorDestino = RGB(201, 133, 12)
        Case "Amarillo": colorDestino = RGB(235, 198, 21)
    End Select
    
    ' Recorrer el rango y cambiar solo los colores espec�ficos
    For Each Celda In ws.Range("BZ38:MX155")
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

