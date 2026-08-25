'========================== DISENO TREN =======================
Public Altura as string
Public Ancho as string
Public Profundidad as string
Public IP as string, Color as string

'UNIDADES FUNCIONALES
    Private bloqueado As Boolean

    ' --- helpers -------------------------------------------------
    Private Sub LimpiarCombo(cbo As Object)
        cbo.Clear
        On Error Resume Next          ' según el Style, uno de los dos aplica
        cbo.ListIndex = -1
        cbo.Text = ""
        On Error GoTo 0
    End Sub

    Private Sub Llenar(cbo As Object, items As String)
        Dim s As Variant
        For Each s In Split(items, "|")
            cbo.AddItem s
        Next
    End Sub

    Private Function Combo(nombre As String) As Object
        Set Combo = Me.OLEObjects(nombre).Object
    End Function

    ' --- eventos -------------------------------------------------
    private sub TipoUF_Change()
        CargarOrientacion
        Me.OleObjects("Carga").Object.Value = false
        Me.OleObjects("Totalizador").Object.Value = false
    end sub
    Private Sub Totalizador_Click()
        CargarOrientacion 
    End Sub
    Private Sub Carga_Click()
        CargarOrientacion
    End Sub
    Private Sub OrientacionUF_Change()
        If bloqueado Then Exit Sub
        bloqueado = True
        CargarEquipoYConexion
        bloqueado = False
    End Sub
    ' --- lógica --------------------------------------------------
    Public Sub CargarOrientacion()
        dim dato As String
        dato = Me.OleObjects("TipoUF").Object.Value
        If bloqueado Then Exit Sub
        bloqueado = True

        With Combo("OrientacionUF")
            LimpiarCombo Combo("OrientacionUF")
            select case dato
                case "SFI - Frontal Interno","SFC - Frontal Compacto": Llenar Combo("OrientacionUF"), "Horizontal|Horizontal Compuesta"
                case "SIA - Integral Abatible": Llenar Combo("OrientacionUF"), "Horizontal Abatible|Horizontal Compuesta Abatible"
            end select
            If Me.Carga.Value Then
                select case dato
                    case "SFI - Frontal Interno","SFC - Frontal Compacto": Llenar Combo("OrientacionUF"), "Vertical|Vertical Compacta"
                    case "SIA - Integral Abatible": Llenar Combo("OrientacionUF"), "Vertical Compacta Abatible"
                end select
            End If
        End With

        LimpiarCombo Combo("EquipoUF")
        LimpiarCombo Combo("ConexionUF")

        bloqueado = False
    End Sub

    Private Sub CargarEquipoYConexion()
        Dim dato As String
        dato = CStr(Combo("OrientacionUF").Value & "")

        LimpiarCombo Combo("EquipoUF")
        LimpiarCombo Combo("ConexionUF")

        If Not Me.Totalizador.Value And Not Me.Carga.Value Then Exit Sub
        If Len(dato) = 0 Then Exit Sub

        ' Equipo
        Select Case dato
            Case "Horizontal", "Horizontal Abatible"
                Llenar Combo("EquipoUF"), "GoPact100|GoPact250|EZC100|EZC250|EZC630|CVS250|CVS630|NSX250|NSX630"
            Case "Horizontal Compuesta", "Horizontal Compuesta Abatible"
                Llenar Combo("EquipoUF"), "GoPact100|GoPact250|EZC100|EZC250|EZC630|CVS250|CVS630|NSX250|NSX630"
            Case "Vertical"
                Llenar Combo("EquipoUF"), "Guardamotor"
            Case "Vertical Compacta"
                Llenar Combo("EquipoUF"), "MiniBreaker|GoPact100|EZC100"
            Case "Vertical Compacta Abatible"    
                Llenar Combo("EquipoUF"), "GoPact100|EZC100"
        End Select

        ' Conexión
        If Me.Totalizador.Value Then
            Llenar Combo("ConexionUF"), "Derecha"
        Else
            Select Case dato
                Case "Horizontal", "Horizontal Compuesta", "Horizontal Abatible", "Horizontal Compuesta Abatible"
                    Llenar Combo("ConexionUF"), "Izquierda|Derecha"
                Case "Vertical", "Vertical Compacta", "Vertical Compacta Abatible"
                    Llenar Combo("ConexionUF"), "NA"
            End Select
        End If
    End Sub  
'UNIDADES FUNCIONALES LISAS
    Private sub TipoLSS_Change()
        dim dato As String
        dato = Me.OleObjects("TipoLSS").Object.Value
        with Me.OleObjects("UnidadesLSS").Object
            .Clear
            Select case dato
                case "SFI - Frontal Interno"
                    .AddItem "1u"
                    .AddItem "2u"
                    .AddItem "3u"
                    .AddItem "4u"
                    .AddItem "5u"
                    .AddItem "6u"
                    .AddItem "7u"
                    .AddItem "8u"
                case "SIA - Integral Abatible" 
                    .AddItem "5u"  
                    .AddItem "6u"
                    .AddItem "7u"
                    .AddItem "8u"
                case "SFA - Fijo Atornillado"
                    .AddItem "2u" 
                    .AddItem "3u"
                    .AddItem "4u"  
                case "SIC - Placa de Control"
                    .AddItem "5u"     
                    .AddItem "6u"
                    .AddItem "7u"
                    .AddItem "8u"
                    .AddItem "9u"
                    .AddItem "10u"
                    .AddItem "11u"
                    .AddItem "12u"
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
    Dim dato As String
    dato = Me.OleObjects("FabriUNDPS").Object.Value
    With Me.OleObjects("EquipoDPS").Object   
    .Clear
    select case dato
        case "HWA"
            .AddItem "NSX100"
            .AddItem "EZC100"
            .AddItem "CVS100"
        case else    
            .AddItem "MiniBreaker"
            .AddItem "NSX100"
            .AddItem "EZC100"
            .AddItem "CVS100"
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
' Accesorios de control  
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
                case "Cubiculo"
                    .AddItem "BusBar"
                    .AddItem "BusBar_MINI"
                    .AddItem "Ventilacion"
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
                        .AddItem "ATS PREMIUM"
                        .AddItem "ATS ESTANDAR"
                        .AddItem "ATS BASICA"
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
                    IF dato2 = "Pequeña" then
                        .AddItem "Interna"
                    elseif dato2 = "Grande" then
                        .AddItem "Externa"
                    end if
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
' BARRAJE
    Private sub OrientacionBRR_Change()
        Dim dato As String
        dato = Me.OleObjects("OrientacionBRR").Object.Value
        With Me.OleObjects("TipoBRR").Object
            .Clear
            select case dato
                case "Vertical"
                    .AddItem "BusBar"
                case "Horizontal"
                    .AddItem "BusBar"
                    .AddItem "Tierra Neutro" 
            end select
        end with
        With Me.OleObjects("MedidaBRR").Object
            .Clear
            select case dato
                case "Vertical"
                    .AddItem "1600"
                    .AddItem "1800"
                    .AddItem "2000"
                case "Horizontal"   
                    .AddItem "400" 
                    .AddItem "720"
            end select
        end with
    end sub 
    Private sub TipoBRR_Change()
        Dim dato As String,dato2 As String
        dato = Me.OleObjects("OrientacionBRR").Object.Value
        dato2 = Me.OleObjects("TipoBRR").Object.Value
        With Me.OleObjects("AmperajeBRR").Object
            .Clear
            select case dato2
                case "BusBar"
                    if dato = "Vertical" then  
                        .AddItem "1250A"
                    elseif dato = "Horizontal" then
                        .AddItem "400A"
                        .AddItem "400X1250A"
                    end if
                case "Tierra Neutro"
                    .AddItem "NA"
            end select
        end with
    end sub
'TRANFERENCIAS
    Private sub OrientacionTRF_Change()
        Tranferencia_activo
        dim dato as string
        dato = Me.OleObjects("OrientacionTRF").Object.Value
        With Me.OleObjects("EquipoTRF").Object
            .Clear
            select case dato
                case "Vertical"
                    .AddItem "Trf. EZC250 Vertical entrada Superior"
                    .AddItem "Trf. EZC250 Vertical entrada Inferior"
                    .AddItem "Trf. EZC630 Vertical entrada Superior"
                    .AddItem "Trf. EZC630 Vertical entrada Inferior"
                    .AddItem "Trf. CVS250 Vertical entrada Superior"
                    .AddItem "Trf. CVS250 Vertical entrada Inferior"
                    .AddItem "Trf. CVS630 Vertical entrada Superior"
                    .AddItem "Trf. CVS630 Vertical entrada Inferior"
                    .AddItem "Trf. NSX250 Vertical entrada Superior"
                    .AddItem "Trf. NSX250 Vertical entrada Inferior"
                    .AddItem "Trf. NSX630 Vertical entrada Superior"
                    .AddItem "Trf. NSX630 Vertical entrada Inferior"
                case "Horizontal", "Horizontal Abatible"
                    .AddItem "EZC250X2"
                    .AddItem "CVS250X2"
                    .AddItem "NSX250X2"
            end select
        end with
    end sub
    Private sub EquipoTRF_Change()
        Dim dato As String
        dato = Me.OleObjects("EquipoTRF").Object.Value
        Dim texto As String
        Select Case  dato
            Case "Trf. EZC250 Vertical entrada Superior"
                texto =  "+C1 | Liso | | | | | W: 720/ U:4 | Unidad Funcional  Lisa / SFI - Frontal Interno | |" & vbCrLf
                texto = texto & "+C2 | Equipada | EZC250 | 3 | Vertical | EZC250N| W: 720 / U: 7 | SFI - Frontal Interno / Vertical/ Cant.equipos: 2 | Superior |" & vbCrLf
                texto = texto & "+C3 | Liso | | | | | W: 720 / U:  4 | Unidad Funcional  Lisa / SFI - Frontal Interno | |" & vbCrLf
            Case "Trf. EZC250 Vertical entrada Inferior"
                texto =  "+C1 | Liso | | | | | W: 720/ U:4 | Unidad Funcional  Lisa / SFI - Frontal Interno | |" & vbCrLf
                texto = texto & "+C2 | Liso | | | | | W: 720 / U:  4 | Unidad Funcional  Lisa / SFI - Frontal Interno | |" & vbCrLf
                texto = texto & "+C3 | Equipada | EZC250 | 3 | Vertical | EZC250N| W: 720 / U: 7 | SFI - Frontal Interno / Vertical/ Cant.equipos: 2 | Inferior |" & vbCrLf
                texto = texto & "+C4 | Liso | | | | | W: 720 / U:  4 | Unidad Funcional  Lisa / SFI - Frontal Interno | |" & vbCrLf
            case "Trf. EZC630 Vertical entrada Superior"    
                texto =  "+C1 | Liso | | | | | W: 720/ U:4 | Unidad Funcional  Lisa / SFI - Frontal Interno | |" & vbCrLf
                texto = texto & "+C2 | Equipada | EZC630 | 3 | Vertical | EZC630N| W: 720 / U: 7 | SFI - Frontal Interno / Vertical/ Cant.equipos: 2 | Superior |" & vbCrLf
                texto = texto & "+C3 | Liso | | | | | W: 720 / U:  4 | Unidad Funcional  Lisa / SFI - Frontal Interno | |" & vbCrLf
            case "Trf. EZC630 Vertical entrada Inferior"
                texto =  "+C1 | Liso | | | | | W: 720/ U:4 | Unidad Funcional  Lisa / SFI - Frontal Interno | |" & vbCrLf
                texto = texto & "+C2 | Liso | | | | | W: 720 / U:  4 | Unidad Funcional  Lisa / SFI - Frontal Interno | |" & vbCrLf
                texto = texto & "+C3 | Equipada | EZC630 | 3 | Vertical | EZC630N| W: 720 / U: 7 | SFI - Frontal Interno / Vertical/ Cant.equipos: 2 | Inferior |" & vbCrLf
                texto = texto & "+C4 | Liso | | | | | W: 720 / U:  4 | Unidad Funcional  Lisa / SFI - Frontal Interno | |" & vbCrLf
            Case "Trf. CVS250 Vertical entrada Superior"
                Texto = "+C1 | Liso | | | | | W: 720/ U:4 | Unidad Funcional  Lisa / SFI - Frontal Interno | |" & vbCrLf
                Texto = Texto & "+C2 | Equipada | CVS250 | 3 | Vertical | CVS250N| W: 720 / U: 7 | SFI - Frontal Interno / Vertical/ Cant.equipos: 2 | Superior |" & vbCrLf
                Texto = Texto & "+C3 | Liso | | | | | W: 720 / U:  4 | Unidad Funcional  Lisa / SFI - Frontal Interno | |" & vbCrLf
            Case "Trf. CVS250 Vertical entrada Inferior"
                Texto = "+C1 | Liso | | | | | W: 720/ U:4 | Unidad Funcional  Lisa / SFI - Frontal Interno | |" & vbCrLf
                Texto = Texto & "+C2 | Liso | | | | | W: 720 / U:  4 | Unidad Funcional  Lisa / SFI - Frontal Interno | |" & vbCrLf
                Texto = Texto & "+C3 | Equipada | CVS250 | 3 | Vertical | CVS250N| W: 720 / U: 7 | SFI - Frontal Interno / Vertical/ Cant.equipos: 2 | Inferior |" & vbCrLf
                Texto = Texto & "+C4 | Liso | | | | | W: 720 / U:  4 | Unidad Funcional  Lisa / SFI - Frontal Interno | |" & vbCrLf
            case "Trf. CVS630 Vertical entrada Superior"
                Texto = "+C1 | Liso | | | | | W: 720/ U:4 | Unidad Funcional  Lisa / SFI - Frontal Interno | |" & vbCrLf
                Texto = Texto & "+C2 | Equipada | CVS630 | 3 | Vertical | CVS630N| W: 720 / U: 7 | SFI - Frontal Interno / Vertical/ Cant.equipos: 2 | Superior |" & vbCrLf
                Texto = Texto & "+C3 | Liso | | | | | W: 720 / U:  4 | Unidad Funcional  Lisa / SFI - Frontal Interno | |" & vbCrLf
            case "Trf. CVS630 Vertical entrada Inferior"
                Texto = "+C1 | Liso | | | | | W: 720/ U:4 | Unidad Funcional  Lisa / SFI - Frontal Interno | |" & vbCrLf
                Texto = Texto & "+C2 | Liso | | | | | W: 720 / U:  4 | Unidad Funcional  Lisa / SFI - Frontal Interno | |" & vbCrLf
                Texto = Texto & "+C3 | Equipada | CVS630 | 3 | Vertical | CVS630N| W: 720 / U: 7 | SFI - Frontal Interno / Vertical/ Cant.equipos: 2 | Inferior |" & vbCrLf
                Texto = Texto & "+C4 | Liso | | | | | W: 720 / U:  4 | Unidad Funcional  Lisa / SFI - Frontal Interno | |" & vbCrLf
            case "Trf. NSX250 Vertical entrada Superior"  
                Texto = "+C1 | Liso | | | | | W: 720/ U:4 | Unidad Funcional  Lisa / SFI - Frontal Interno | |" & vbCrLf
                Texto = Texto & "+C2 | Equipada | NSX250 | 3 | Vertical | NSX250N| W: 720 / U: 7 | SFI - Frontal Interno / Vertical/ Cant.equipos: 2 | Superior |" & vbCrLf
                Texto = Texto & "+C3 | Liso | | | | | W: 720 / U:  4 | Unidad Funcional  Lisa / SFI - Frontal Interno | |" & vbCrLf  
            Case "Trf. NSX250 Vertical entrada Inferior"
                Texto = "+C1 | Liso | | | | | W: 720/ U:4 | Unidad Funcional  Lisa / SFI - Frontal Interno | |" & vbCrLf
                Texto = Texto & "+C2 | Liso | | | | | W: 720 / U:  4 | Unidad Funcional  Lisa / SFI - Frontal Interno | |" & vbCrLf
                Texto = Texto & "+C3 | Equipada | NSX250 | 3 | Vertical | NSX250N| W: 720 / U: 7 | SFI - Frontal Interno / Vertical/ Cant.equipos: 2 | Inferior |" & vbCrLf
                Texto = Texto & "+C4 | Liso | | | | | W: 720 / U:  4 | Unidad Funcional  Lisa / SFI - Frontal Interno | |" & vbCrLf
            case "Trf. NSX630 Vertical entrada Superior"
                Texto = "+C1 | Liso | | | | | W: 720/ U:4 | Unidad Funcional  Lisa / SFI - Frontal Interno | |" & vbCrLf
                Texto = Texto & "+C2 | Equipada | NSX630 | 3 | Vertical | NSX630N| W: 720 / U: 7 | SFI - Frontal Interno / Vertical/ Cant.equipos: 2 | Superior |" & vbCrLf
                Texto = Texto & "+C3 | Liso | | | | | W: 720 / U:  4 | Unidad Funcional  Lisa / SFI - Frontal Interno | |" & vbCrLf
            case "Trf. NSX630 Vertical entrada Inferior"
                Texto = "+C1 | Liso | | | | | W: 720/ U:4 | Unidad Funcional  Lisa / SFI - Frontal Interno | |" & vbCrLf
                Texto = Texto & "+C2 | Liso | | | | | W: 720 / U:  4 | Unidad Funcional  Lisa / SFI - Frontal Interno | |" & vbCrLf
                Texto = Texto & "+C3 | Equipada | NSX630 | 3 | Vertical | NSX630N| W: 720 / U: 7 | SFI - Frontal Interno / Vertical/ Cant.equipos: 2 | Inferior |" & vbCrLf
                Texto = Texto & "+C4 | Liso | | | | | W: 720 / U:  4 | Unidad Funcional  Lisa / SFI - Frontal Interno | |" & vbCrLf    
            case else 
             texto = "Seleccion no existente"   
        end select
         Me.OleObjects("Descripcion").Object.value = texto  
    end sub
    Public sub Tranferencia_activo()
        Dim dato As String
        dato = Me.OleObjects("OrientacionTRF").Object.Value
        select case dato
            case "Vertical"
                Me.OleObjects("Descripcion").Object.Enabled = true
            case "Horizontal", "Horizontal Abatible"
                Me.OleObjects("Descripcion").Object.Enabled = False
        end select
    end sub
' Accesorio MTM
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
    Me.OleObjects("Carga").Object.Value = false
    Me.OleObjects("Totalizador").Object.Value = false
    '============== Envolventes ===================
        With  Me.OleObjects("AlturaENV").Object
            .AddItem Altura
        End With
        With  Me.OleObjects("AnchoENV").Object
            .clear
            .AddItem "200"
            .Additem "400"
            .Additem "720"
        End With  
    '============== Unidades Funcionales ===================
        With  Me.OleObjects("TipoUF").Object
            .Clear
            select case IP
                Case "30"
                .AddItem "SFI - Frontal Interno"
                .AddItem "SIA - Integral Abatible"
                .AddItem "SFC - Frontal Compacto"
                
            case else
                .AddItem "SFI - Frontal Interno"
                .AddItem "SIA - Integral Abatible" 
            end select  
        End With
    '============== Unidades Funcionales Lisas ===================
        With  Me.OleObjects("TipoLSS").Object
            .Clear
            .AddItem "SFI - Frontal Interno"
            .AddItem "SIA - Integral Abatible"
            .AddItem "SFA - Fijo Atornillado"
            .AddItem "SIC - Placa de Control"
        End With
        With  Me.OleObjects("AnchoLSS").Object
            .Clear
            .AddItem "400"
            .AddItem "720"
        End With
    '============== Contactores ===================
        with Me.OleObjects("EquipoCT").Object
            .Clear
            '.AddItem "A-Unit (Barra + Mini)"
            .AddItem "Contactores"
        End With
    '============== DPS ===================
        With Me.OleObjects("TipoUNDPS").Object
            .Clear
            .AddItem "SFI - Frontal Interno"
            .AddItem "SIA - Integral Abatible"
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
            .AddItem "Cubiculo"
        End With
    '============== PLACAS DE CONTROL ===================  
        With Me.OleObjects("TamanoPC").Object
            .Clear
            .AddItem "Pequeña"
            .AddItem "Mediana"
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
            .AddItem "Horizontal"
            .AddItem "Vertical"
        End With
    '============== TRANSFERENCIA =============================
        with Me.OleObjects("OrientacionTRF").Object
            .Clear
            .AddItem "Horizontal"
            .AddItem "Horizontal Abatible"
            .AddItem "Vertical"
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
        Me.oleObjects("ColorPTD").Object.Value = Color
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
Public sub Extraer_datos()
    'ESTRAER DATOS DE DIMENSIONES
        Altura = Range("A1").Value
        Ancho = Range("A2").Value
        Profundidad = Range("A3").Value
        IP = Range("A4").Value
        Color = Range("A5").Value
    
End Sub

Public sub Envolventes()
    '========================
        celda = "NS98"
        thisWorkbook.Sheets("DISENO").Range("NS98:QJ182").Clear
        Set wsActual = thisWorkbook.Sheets("DISENO")
    '========================
    dim dato1 as string, dato2 as string
    dato1 = Me.OleObjects("AlturaENV").Object.Value
    dato2 = Me.OleObjects("AnchoENV").Object.Value
    ResultadoENV= "EFM_" & dato1 & "X" & dato2
    MoverBloque ResultadoENV,"ENVOLVENTES", wsActual, celda
End sub
Public sub Unidades_funcionales()
    '========================
        celda = "NS115"
        thisWorkbook.Sheets("DISENO").Range("NS115:QJ182").Clear
        Set wsActual = thisWorkbook.Sheets("DISENO")
    '========================
    dim dato1 as string, dato2 as string, dato3 as string, dato4 as string
    dato1 = Me.OleObjects("TipoUF").Object.Value
    select case dato1
        case "SFI - Frontal Interno","SFC - Frontal Compacto": dato1 = "U"
        Case "SIA - Integral Abatible": dato1 = "USIA_"
    END SELECT
    dato2 = Me.OleObjects("OrientacionUF").Object.Value 
    Select case dato2
        Case "Horizontal", "Horizontal Abatible": dato2 = "H"
        Case "Vertical": dato2 = "V"
        case "Horizontal Compuesta", "Horizontal Compuesta Abatible": dato2 = "HC"
        case "Vertical Compacta", "Vertical Compacta Abatible": dato2 = "VC"
    END SELECT  
    dato3 = Me.OleObjects("EquipoUF").Object.Value
    dato4 = Me.OleObjects("ConexionUF").Object.Value
    select case dato4
        case "Izquierda": dato4 = "_IZ"
        case "Derecha": dato4 = "_DR"
        case "NA": dato4 = "_NA"
    END SELECT
  
    ResultadoUF= dato1  & dato3 & "X720" & dato2 & dato4
    MoverBloque ResultadoUF,"Ufuncionales", wsActual, celda
End sub
Public sub Unidades_funcionales_lisas()
    '========================
        celda = "NS132"
        thisWorkbook.Sheets("DISENO").Range("NS132:QJ182").Clear
        Set wsActual = thisWorkbook.Sheets("DISENO")
    '========================
    dim dato1 as string, dato2 as string, dato3 as string
    dato1 = Me.OleObjects("TipoLSS").Object.Value
    select case dato1
        case "SFI - Frontal Interno": dato1 = "L_"
        Case "SIA - Integral Abatible", "SFA - Fijo Atornillado": dato1 = "LA_"
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
        thisWorkbook.Sheets("DISENO").Range("NS132:QJ182").Clear
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
        thisWorkbook.Sheets("DISENO").Range("NS132:QJ182").Clear
        Set wsActual = thisWorkbook.Sheets("DISENO")
    '========================
    dim dato1 as string, dato2 as string, dato3 as string, dato4 as string
    dato1 = Me.OleObjects("TipoUNDPS").Object.Value
    Select case dato1
        case "SFI - Frontal Interno": dato1 = "D"
        Case "SIA - Integral Abatible": dato1 = "DA"
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

    ResultadoDPS= dato1 & dato2 & "X" & dato3 & "_" & dato4
    MoverBloque ResultadoDPS,"DPS", wsActual, celda
End sub
Public sub Equipos_Control()
    '========================
        celda = "NS132"
        thisWorkbook.Sheets("DISENO").Range("NS132:QJ182").Clear
        Set wsActual = thisWorkbook.Sheets("DISENO")
    '========================
    dim dato1 as string, dato2 as string
    dato1 = Me.OleObjects("EquipoCTR").Object.Value
    dato2 = Me.OleObjects("MedidasCTR").Object.Value
    ResultadoCT = dato1 & "_" & dato2
    MoverBloque ResultadoCT,"UCONTROL", wsActual, celda
End sub
Public sub Accesorios_CTR()
    '========================
        celda = "NS132"
        thisWorkbook.Sheets("DISENO").Range("NS132:QJ182").Clear
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
Public sub Equipo_Individual()
    '========================
        celda = "NS132" 
        thisWorkbook.Sheets("DISENO").Range("NS132:QJ182").Clear
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
        thisWorkbook.Sheets("DISENO").Range("NS132:QJ182").Clear
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
        thisWorkbook.Sheets("DISENO").Range("NS132:QJ182").Clear
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
        thisWorkbook.Sheets("DISENO").Range("NU107:QJ182").Clear
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
        thisWorkbook.Sheets("DISENO").Range("NS104:QJ182").Clear
        Set wsActual = thisWorkbook.Sheets("DISENO")
    '========================
    Dim dato1 As String, dato2 As String, dato3 As String, dato4 As String
    dato1 = Me.OleObjects("OrientacionBRR").Object.Value
    Select case dato1
        case "Horizontal": dato1 = "H"
        case "Vertical": dato1 = "V"
    End Select
    dato2 = Me.OleObjects("TipoBRR").Object.Value
    Select case dato2
        case "BusBar": dato2 = "BB"
        case "Tierra Neutro": dato2 = "TN"
    End Select
    dato3 = Me.OleObjects("AmperajeBRR").Object.Value
    dato4 = Me.OleObjects("MedidaBRR").Object.Value
    ResultadoBAR = "UN_" & dato1 &  dato2 & "X" & dato4 & "_" & dato3
    MoverBloque ResultadoBAR,"BARRAJE", wsActual, celda
End sub
Public sub Transferencia()
    '========================
        celda = "NU107"
        thisWorkbook.Sheets("DISENO").Range("NU107:QJ182").Clear
        Set wsActual = thisWorkbook.Sheets("DISENO")
    '========================
    Dim dato1 As String, dato2 As String
    dato1 = Me.OleObjects("EquipoTRF").Object.Value
    select case dato1
        case "Trf. EZC250 Vertical entrada Superior": dato1 = "EZC250XVX720XSUP"
        case "Trf. EZC250 Vertical entrada Inferior": dato1 = "EZC250XVX720XINF"
        case "Trf. EZC630 Vertical entrada Superior": dato1 = "EZC630XVX720XSUP"
        case "Trf. EZC630 Vertical entrada Inferior": dato1 = "EZC630XVX720XINF"
        case "Trf. CVS250 Vertical entrada Superior": dato1 = "CVS250XVX720XSUP"
        case "Trf. CVS250 Vertical entrada Inferior": dato1 = "CVS250XVX720XINF"
        case "Trf. CVS630 Vertical entrada Superior": dato1 = "CVS630XVX720XSUP"
        case "Trf. CVS630 Vertical entrada Inferior": dato1 = "CVS630XVX720XINF"
        case "Trf. NSX250 Vertical entrada Superior": dato1 = "NSX250XVX720XSUP"
        case "Trf. NSX250 Vertical entrada Inferior": dato1 = "NSX250XVX720XINF"
        case "Trf. NSX630 Vertical entrada Superior": dato1 = "NSX630XVX720XSUP"
        case "Trf. NSX630 Vertical entrada Inferior": dato1 = "NSX630XVX720XINF"
    end select
    dato2 = Me.OleObjects("OrientacionTRF").Object.Value
    select case dato2
        case "Vertical"
        ResultadoTRF = "UT_" &  dato1
        case "Horizontal"
            dato2 = "H"
            ResultadoTRF = "UT_" & dato1 & "X" & dato2 & "X720"
        case "Horizontal Abatible"
            dato2 = "HA"
            ResultadoTRF = "UT_" & dato1 & "X" & dato2 & "X720"
    end select
    MoverBloque ResultadoTRF,"U_TRANSFERENCIA", wsActual, celda

End sub
Public sub ACESORIOS_MTM()
    '========================
        celda = "NS104"
        thisWorkbook.Sheets("DISENO").Range("NS104:QJ182").Clear
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
        thisWorkbook.Sheets("DISENO").Range("NS104:QJ182").Clear
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
        thisWorkbook.Sheets("DISENO").Range("NS104:QJ182").Clear
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