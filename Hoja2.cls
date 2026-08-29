'========================== DISENO TREN =======================
Public Altura As String
Public Ancho As String
Public Profundidad As String
Public IP As String, Color As String
Public Linea As String
Private bloqueado As Boolean
'Envolentes
    Public sub AlturaENV_Change()
        With Me.OLEObjects("AnchoENV").Object
            .Clear
            select case Linea
                case "Heavy"
                    .AddItem "200"
                    .AddItem "400"
                    .AddItem "720"
                    .AddItem "1120"
                case "Modular"
                    .AddItem "200"
                    .AddItem "400"
                    .AddItem "720"
                    .AddItem "800"
            end select            
        End With
    end sub    
'UNIDADES FUNCIONALES
Private Sub TipoUF_Change()
        CargarOrientacion
        Me.OLEObjects("Carga").Object.Value = False
        Me.OLEObjects("Totalizador").Object.Value = False
End Sub
Public Sub CargarOrientacion()
    dim dato As String
    dato = Me.OLEObjects("TipoUF").Object.Value
    With Me.OLEObjects("OrientacionUF").Object
        .Clear
        If Me.Totalizador.Value Then 
            Select Case dato
            case "SFI - Frontal Interno", "SFA - Fijo Atornillado"
                .AddItem "Horizontal"
                .AddItem "H. Compuesta"
            case "SIA - Integral Abatible"
                .AddItem "Horizontal Abatible"    
                .AddItem "H. Compuesta Abatible"
            End Select
        Elseif Not Me.Totalizador.Value Then
            Select Case dato
            case "SFI - Frontal Interno", "SFA - Fijo Atornillado" 
                .AddItem "Horizontal"
                .AddItem "H. Compuesta" 
                .AddItem "Vertical"  
                .AddItem "V. Compuesta"
            case "SIA - Integral Abatible"
                .AddItem "Horizontal Abatible"    
                .AddItem "H. Compuesta Abatible"
                .AddItem "Vertical Abatible"    
                .AddItem "V. Compuesta Abatible"
            End Select
        end if
    end with
end sub
Private Sub Totalizador_Click() 
    CargarOrientacion
End Sub
Private Sub Carga_Click()
    CargarOrientacion
End Sub
Private sub OrientacionUF_Change()
    dim dato As String
    dato = Me.OLEObjects("OrientacionUF").Object.Value
    With Me.OLEObjects("EquipoUF").Object
        .Clear
        Select Case dato
            Case "Horizontal", "Horizontal Abatible"
                .AddItem "GoPact100"
                .AddItem "GoPact250"
                .AddItem "EZC100"
                .AddItem "EZC250"
                .AddItem "EZC630"
                .AddItem "CVS250"
                .AddItem "CVS630"
                .AddItem "NSX250"
                .AddItem "NSX630"
                .AddItem "NS800"
                .AddItem "NS1250"
                .AddItem "NS_Motorizado"
            Case "H. Compuesta", "H. Compuesta Abatible"
                .AddItem "GoPact100"   
                .AddItem "GoPact250" 
                .AddItem "EZC100"
                .AddItem "EZC250"
                .AddItem "EZC630"
                .AddItem "NSX250"
                .AddItem "NSX630"
            case "Vertical"  
                .AddItem "CVS630"
                .AddItem "EZC630"
                .AddItem "Guardamotor"
                .AddItem "GoPact800"   
                .AddItem "NSX250"
                .AddItem "NSX630"
                .AddItem "NS"
                .AddItem "NS_Motorizado"
                .AddItem "NW_Fijo"
                .AddItem "NW_Extraible"
                .AddItem "MTZ_Fijo"
                .AddItem "MTZ_Extraible"
            case "Vertical Abatible"   
                .AddItem "GoPact800" 
                .AddItem "EZC630"
                .AddItem "CVS250"
                .AddItem "NSX250"
                .AddItem "NSX630"
                .AddItem "NS800"
                .AddItem "NS_Motorizado"
                .AddItem "NW_Fijo"
                .AddItem "NW_Extraible"
                .AddItem "MTZ_Fijo"
                .AddItem "MTZ_Extraible"
            case "V. Compuesta"   
                .AddItem "MiniBreaker"
                .AddItem "GoPact100"
                .AddItem "GoPact400"
                .AddItem "EZC100"
                .AddItem "EZC250"
                .AddItem "CVS100"
            case "V. Compuesta Abatible"   
                .AddItem "GoPact100"
                .AddItem "GoPact250"
                .AddItem "GoPact400" 
                .AddItem "EZC100"
                .AddItem "EZC250"
        end select
    end with   
    with Me.OLEObjects("ConexionUF").Object 
        .Clear
        If Me.Totalizador.Value Then 
            if dato= "Horizontal" or dato = "Horizontal Abatible" or dato = "H. Compuesta" or dato = "H. Compuesta Abatible" then
                .AddItem "Derecha"
            end if
        Elseif Not Me.Totalizador.Value Then
            if dato= "Horizontal" or dato = "Horizontal Abatible" or dato = "H. Compuesta"or dato = "H. Compuesta Abatible"  then
                .AddItem "Derecha"
                .AddItem "Izquierda"
            elseif dato= "Vertical" or dato = "Vertical Abatible" or dato = "V. Compuesta"or dato = "V. Compuesta Abatible" then
                .AddItem "NA"    
            end if  
        end if
    end with
end sub
Private sub EquipoUF_Change()
    Dim dato As String
    dato = Me.OLEObjects("EquipoUF").Object.Value
    With Me.OLEObjects("MedidaUF").Object
        .Clear
        Select case dato
            case "CVS250", "NS800", "GoPact400", "GoPact800", "CVS100"
                    .AddItem "400mm"
            case "MiniBreaker", "EZC100", "CVS250", _
                "CVS630", "NS1250", "NS_Motorizado","NS Motorizado","NS1250"
                    .AddItem "720mm"  
            case "GoPact100", "EZC250", "EZC630", "CVS250", "NSX250", "NSX630", "NS800", "NS", "GoPact250"
                    .AddItem "400mm"   
                    .AddItem "720mm" 
            Case"MTZ_Extraible", "MTZ_Fijo", "NW_Extraible","NW_Fijo"
                    .AddItem "720mm" 
                    .AddItem "1120mm"           
        end select
    end with
end sub                   
'UNIDADES FUNCIONALES LISAS
    Private Sub TipoLSS_Change()
        Dim dato As String
        dato = Me.OLEObjects("TipoLSS").Object.Value
        With Me.OLEObjects("UnidadesLSS").Object
            .Clear
            Select Case dato
                Case "SFI - Frontal Interno"
                    .AddItem "1u"
                    .AddItem "2u"
                    .AddItem "3u"
                    .AddItem "4u"
                    .AddItem "5u"
                    .AddItem "6u"
                    .AddItem "7u"
                    .AddItem "8u"
                Case "SIA - Integral Abatible"
                    .AddItem "5u"
                    .AddItem "6u"
                    .AddItem "7u"
                    .AddItem "8u"
                Case "SFA - Fijo Atornillado"
                    .AddItem "2u"
                    .AddItem "3u"
                    .AddItem "4u"
                Case "SIC - Placa de Control"
                    .AddItem "5u"
                    .AddItem "6u"
                    .AddItem "7u"
                    .AddItem "8u"
                    .AddItem "9u"
                    .AddItem "10u"
                    .AddItem "11u"
                    .AddItem "12u"
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
    Dim dato As String
    dato = Me.OLEObjects("FabriUNDPS").Object.Value
    With Me.OLEObjects("EquipoDPS").Object
    .Clear
    Select Case dato
        Case "HWA"
            .AddItem "NSX100"
            .AddItem "EZC100"
            .AddItem "CVS100"
        Case Else
            .AddItem "MiniBreaker"
            .AddItem "NSX100"
            .AddItem "EZC100"
            .AddItem "CVS100"
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
' Accesorios de control
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
                Case "Cubiculo"
                    .AddItem "BusBar"
                    .AddItem "BusBar_MINI"
                    .AddItem "Ventilacion"
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
                        .AddItem "ATS PREMIUM"
                        .AddItem "ATS ESTANDAR"
                        .AddItem "ATS BASICA"
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
                    If dato2 = "Peque�a" Then
                        .AddItem "Interna"
                    ElseIf dato2 = "Grande" Then
                        .AddItem "Externa"
                    End If
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
' BARRAJE
    Private Sub OrientacionBRR_Change()
        Dim dato As String
        dato = Me.OLEObjects("OrientacionBRR").Object.Value
        With Me.OLEObjects("TipoBRR").Object
            .Clear
            Select Case dato
                Case "Vertical"
                    .AddItem "BusBar"
                Case "Horizontal"
                    .AddItem "BusBar"
                    .AddItem "Tierra Neutro"
            End Select
        End With
        With Me.OLEObjects("MedidaBRR").Object
            .Clear
            Select Case dato
                Case "Vertical"
                    .AddItem "1600"
                    .AddItem "1800"
                    .AddItem "2000"
                Case "Horizontal"
                    .AddItem "400"
                    .AddItem "720"
            End Select
        End With
    End Sub
    Private Sub TipoBRR_Change()
        Dim dato As String, dato2 As String
        dato = Me.OLEObjects("OrientacionBRR").Object.Value
        dato2 = Me.OLEObjects("TipoBRR").Object.Value
        With Me.OLEObjects("AmperajeBRR").Object
            .Clear
            Select Case dato2
                Case "BusBar"
                    If dato = "Vertical" Then
                        .AddItem "1250A"
                    ElseIf dato = "Horizontal" Then
                        .AddItem "400A"
                        .AddItem "400X1250A"
                    End If
                Case "Tierra Neutro"
                    .AddItem "NA"
            End Select
        End With
    End Sub
'TRANFERENCIAS
    Private Sub OrientacionTRF_Change()
        Tranferencia_activo
        Dim dato As String
        dato = Me.OLEObjects("OrientacionTRF").Object.Value
        With Me.OLEObjects("EquipoTRF").Object
            .Clear
            Select Case dato
                Case "Vertical"
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
                Case "Horizontal", "Horizontal Abatible"
                    .AddItem "EZC250X2"
                    .AddItem "CVS250X2"
                    .AddItem "NSX250X2"
            End Select
        End With
    End Sub
    Private Sub EquipoTRF_Change()
        Dim dato As String
        dato = Me.OLEObjects("EquipoTRF").Object.Value
        Dim texto As String
        Select Case dato
            Case "Trf. EZC250 Vertical entrada Superior"
                texto = "+C1 | Liso | | | | | W: 720/ U:4 | Unidad Funcional  Lisa / SFI - Frontal Interno | |" & vbCrLf
                texto = texto & "+C2 | Equipada | EZC250 | 3 | Vertical | EZC250N| W: 720 / U: 7 | SFI - Frontal Interno / Vertical/ Cant.equipos: 2 | Superior |" & vbCrLf
                texto = texto & "+C3 | Liso | | | | | W: 720 / U:  4 | Unidad Funcional  Lisa / SFI - Frontal Interno | |" & vbCrLf
            Case "Trf. EZC250 Vertical entrada Inferior"
                texto = "+C1 | Liso | | | | | W: 720/ U:4 | Unidad Funcional  Lisa / SFI - Frontal Interno | |" & vbCrLf
                texto = texto & "+C2 | Liso | | | | | W: 720 / U:  4 | Unidad Funcional  Lisa / SFI - Frontal Interno | |" & vbCrLf
                texto = texto & "+C3 | Equipada | EZC250 | 3 | Vertical | EZC250N| W: 720 / U: 7 | SFI - Frontal Interno / Vertical/ Cant.equipos: 2 | Inferior |" & vbCrLf
                texto = texto & "+C4 | Liso | | | | | W: 720 / U:  4 | Unidad Funcional  Lisa / SFI - Frontal Interno | |" & vbCrLf
            Case "Trf. EZC630 Vertical entrada Superior"
                texto = "+C1 | Liso | | | | | W: 720/ U:4 | Unidad Funcional  Lisa / SFI - Frontal Interno | |" & vbCrLf
                texto = texto & "+C2 | Equipada | EZC630 | 3 | Vertical | EZC630N| W: 720 / U: 7 | SFI - Frontal Interno / Vertical/ Cant.equipos: 2 | Superior |" & vbCrLf
                texto = texto & "+C3 | Liso | | | | | W: 720 / U:  4 | Unidad Funcional  Lisa / SFI - Frontal Interno | |" & vbCrLf
            Case "Trf. EZC630 Vertical entrada Inferior"
                texto = "+C1 | Liso | | | | | W: 720/ U:4 | Unidad Funcional  Lisa / SFI - Frontal Interno | |" & vbCrLf
                texto = texto & "+C2 | Liso | | | | | W: 720 / U:  4 | Unidad Funcional  Lisa / SFI - Frontal Interno | |" & vbCrLf
                texto = texto & "+C3 | Equipada | EZC630 | 3 | Vertical | EZC630N| W: 720 / U: 7 | SFI - Frontal Interno / Vertical/ Cant.equipos: 2 | Inferior |" & vbCrLf
                texto = texto & "+C4 | Liso | | | | | W: 720 / U:  4 | Unidad Funcional  Lisa / SFI - Frontal Interno | |" & vbCrLf
            Case "Trf. CVS250 Vertical entrada Superior"
                texto = "+C1 | Liso | | | | | W: 720/ U:4 | Unidad Funcional  Lisa / SFI - Frontal Interno | |" & vbCrLf
                texto = texto & "+C2 | Equipada | CVS250 | 3 | Vertical | CVS250N| W: 720 / U: 7 | SFI - Frontal Interno / Vertical/ Cant.equipos: 2 | Superior |" & vbCrLf
                texto = texto & "+C3 | Liso | | | | | W: 720 / U:  4 | Unidad Funcional  Lisa / SFI - Frontal Interno | |" & vbCrLf
            Case "Trf. CVS250 Vertical entrada Inferior"
                texto = "+C1 | Liso | | | | | W: 720/ U:4 | Unidad Funcional  Lisa / SFI - Frontal Interno | |" & vbCrLf
                texto = texto & "+C2 | Liso | | | | | W: 720 / U:  4 | Unidad Funcional  Lisa / SFI - Frontal Interno | |" & vbCrLf
                texto = texto & "+C3 | Equipada | CVS250 | 3 | Vertical | CVS250N| W: 720 / U: 7 | SFI - Frontal Interno / Vertical/ Cant.equipos: 2 | Inferior |" & vbCrLf
                texto = texto & "+C4 | Liso | | | | | W: 720 / U:  4 | Unidad Funcional  Lisa / SFI - Frontal Interno | |" & vbCrLf
            Case "Trf. CVS630 Vertical entrada Superior"
                texto = "+C1 | Liso | | | | | W: 720/ U:4 | Unidad Funcional  Lisa / SFI - Frontal Interno | |" & vbCrLf
                texto = texto & "+C2 | Equipada | CVS630 | 3 | Vertical | CVS630N| W: 720 / U: 7 | SFI - Frontal Interno / Vertical/ Cant.equipos: 2 | Superior |" & vbCrLf
                texto = texto & "+C3 | Liso | | | | | W: 720 / U:  4 | Unidad Funcional  Lisa / SFI - Frontal Interno | |" & vbCrLf
            Case "Trf. CVS630 Vertical entrada Inferior"
                texto = "+C1 | Liso | | | | | W: 720/ U:4 | Unidad Funcional  Lisa / SFI - Frontal Interno | |" & vbCrLf
                texto = texto & "+C2 | Liso | | | | | W: 720 / U:  4 | Unidad Funcional  Lisa / SFI - Frontal Interno | |" & vbCrLf
                texto = texto & "+C3 | Equipada | CVS630 | 3 | Vertical | CVS630N| W: 720 / U: 7 | SFI - Frontal Interno / Vertical/ Cant.equipos: 2 | Inferior |" & vbCrLf
                texto = texto & "+C4 | Liso | | | | | W: 720 / U:  4 | Unidad Funcional  Lisa / SFI - Frontal Interno | |" & vbCrLf
            Case "Trf. NSX250 Vertical entrada Superior"
                texto = "+C1 | Liso | | | | | W: 720/ U:4 | Unidad Funcional  Lisa / SFI - Frontal Interno | |" & vbCrLf
                texto = texto & "+C2 | Equipada | NSX250 | 3 | Vertical | NSX250N| W: 720 / U: 7 | SFI - Frontal Interno / Vertical/ Cant.equipos: 2 | Superior |" & vbCrLf
                texto = texto & "+C3 | Liso | | | | | W: 720 / U:  4 | Unidad Funcional  Lisa / SFI - Frontal Interno | |" & vbCrLf
            Case "Trf. NSX250 Vertical entrada Inferior"
                texto = "+C1 | Liso | | | | | W: 720/ U:4 | Unidad Funcional  Lisa / SFI - Frontal Interno | |" & vbCrLf
                texto = texto & "+C2 | Liso | | | | | W: 720 / U:  4 | Unidad Funcional  Lisa / SFI - Frontal Interno | |" & vbCrLf
                texto = texto & "+C3 | Equipada | NSX250 | 3 | Vertical | NSX250N| W: 720 / U: 7 | SFI - Frontal Interno / Vertical/ Cant.equipos: 2 | Inferior |" & vbCrLf
                texto = texto & "+C4 | Liso | | | | | W: 720 / U:  4 | Unidad Funcional  Lisa / SFI - Frontal Interno | |" & vbCrLf
            Case "Trf. NSX630 Vertical entrada Superior"
                texto = "+C1 | Liso | | | | | W: 720/ U:4 | Unidad Funcional  Lisa / SFI - Frontal Interno | |" & vbCrLf
                texto = texto & "+C2 | Equipada | NSX630 | 3 | Vertical | NSX630N| W: 720 / U: 7 | SFI - Frontal Interno / Vertical/ Cant.equipos: 2 | Superior |" & vbCrLf
                texto = texto & "+C3 | Liso | | | | | W: 720 / U:  4 | Unidad Funcional  Lisa / SFI - Frontal Interno | |" & vbCrLf
            Case "Trf. NSX630 Vertical entrada Inferior"
                texto = "+C1 | Liso | | | | | W: 720/ U:4 | Unidad Funcional  Lisa / SFI - Frontal Interno | |" & vbCrLf
                texto = texto & "+C2 | Liso | | | | | W: 720 / U:  4 | Unidad Funcional  Lisa / SFI - Frontal Interno | |" & vbCrLf
                texto = texto & "+C3 | Equipada | NSX630 | 3 | Vertical | NSX630N| W: 720 / U: 7 | SFI - Frontal Interno / Vertical/ Cant.equipos: 2 | Inferior |" & vbCrLf
                texto = texto & "+C4 | Liso | | | | | W: 720 / U:  4 | Unidad Funcional  Lisa / SFI - Frontal Interno | |" & vbCrLf
            Case Else
             texto = "Seleccion no existente"
        End Select
         Me.OLEObjects("Descripcion").Object.Value = texto
    End Sub
    Public Sub Tranferencia_activo()
        Dim dato As String
        dato = Me.OLEObjects("OrientacionTRF").Object.Value
        Select Case dato
            Case "Vertical"
                Me.OLEObjects("Descripcion").Object.Enabled = True
            Case "Horizontal", "Horizontal Abatible"
                Me.OLEObjects("Descripcion").Object.Enabled = False
        End Select
    End Sub
' Accesorio MTM
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
    Me.OLEObjects("Carga").Object.Value = False
    Me.OLEObjects("Totalizador").Object.Value = False
    '============== Envolventes ===================
        With Me.OLEObjects("AlturaENV").Object
            .Clear
            .AddItem Altura
        End With
    '============== Unidades Funcionales ===================
        With Me.OLEObjects("TipoUF").Object
            .Clear
            .AddItem "SFI - Frontal Interno"
            .AddItem "SIA - Integral Abatible"
            .AddItem "SFA - Fijo Atornillado"
        End With
    '============== Unidades Funcionales Lisas ===================
        With Me.OLEObjects("TipoLSS").Object
            .Clear
            .AddItem "SFI - Frontal Interno"
            .AddItem "SIA - Integral Abatible"
            .AddItem "SFA - Fijo Atornillado"
            .AddItem "SIC - Placa de Control"
        End With
        With Me.OLEObjects("AnchoLSS").Object
            .Clear
            .AddItem "400"
            .AddItem "720"
        End With
    '============== Contactores ===================
        With Me.OLEObjects("EquipoCT").Object
            .Clear
            '.AddItem "A-Unit (Barra + Mini)"
            .AddItem "Contactores"
        End With
    '============== DPS ===================
        With Me.OLEObjects("TipoUNDPS").Object
            .Clear
            .AddItem "SFI - Frontal Interno"
            .AddItem "SIA - Integral Abatible"
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
            .AddItem "Cubiculo"
        End With
    '============== PLACAS DE CONTROL ===================
        With Me.OLEObjects("TamanoPC").Object
            .Clear
            .AddItem "Peque�a"
            .AddItem "Mediana"
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
            .AddItem "Horizontal"
            .AddItem "Vertical"
        End With
    '============== TRANSFERENCIA =============================
        With Me.OLEObjects("OrientacionTRF").Object
            .Clear
            .AddItem "Horizontal"
            .AddItem "Horizontal Abatible"
            .AddItem "Vertical"
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
        Me.OLEObjects("ColorPTD").Object.Value = Color
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
Public Sub Extraer_datos()
    'ESTRAER DATOS DE DIMENSIONES
        Altura = Range("A1").Value
        Ancho = Range("A2").Value
        Profundidad = Range("A3").Value
        IP = Range("A4").Value
        Color = Range("A5").Value
        Linea = Range("A6").Value
    
End Sub
Public Sub Envolventes()
    '========================
        Celda = "NS98"
        ThisWorkbook.Sheets("DISENO").Range("NS98:QJ182").Clear
        Set wsActual = ThisWorkbook.Sheets("DISENO")
    '========================
    Dim dato1 As String, dato2 As String
    dato1 = Me.OLEObjects("AlturaENV").Object.Value
    dato2 = Me.OLEObjects("AnchoENV").Object.Value
    ResultadoENV = "EFM_" & dato1 & "X" & dato2
    MoverBloque ResultadoENV, "ENVOLVENTES", wsActual, Celda
End Sub
Public Sub Unidades_funcionales()
    '========================
        Celda = "NS115"
        ThisWorkbook.Sheets("DISENO").Range("NS115:QJ182").Clear
        Set wsActual = ThisWorkbook.Sheets("DISENO")
    '========================
    Dim dato1 As String, dato2 As String, dato3 As String, dato4 As String,dato5 As String
    dato1 = Me.OLEObjects("TipoUF").Object.Value
    Select Case dato1
        Case "SFI - Frontal Interno", "SFC - Frontal Compacto": dato1 = "U"
        Case "SIA - Integral Abatible": dato1 = "USIA_"
    End Select
    dato2 = Me.OLEObjects("OrientacionUF").Object.Value
    Select Case dato2
        Case "Horizontal", "Horizontal Abatible": dato2 = "H"
        Case "Vertical": dato2 = "V"
        Case "Horizontal Compuesta", "Horizontal Compuesta Abatible": dato2 = "HC"
        Case "Vertical Compacta", "Vertical Compacta Abatible": dato2 = "VC"
    End Select
    dato3 = Me.OLEObjects("EquipoUF").Object.Value
    dato4 = Me.OLEObjects("ConexionUF").Object.Value
    Select Case dato4
        Case "Izquierda": dato4 = "_IZ"
        Case "Derecha": dato4 = "_DR"
        Case "NA": dato4 = "_NA"
    End Select
    dato5 = Me.OLEObjects("MedidaUF").Object.Value
    Select case dato5
        Case "720"
            ResultadoUF = dato1 & dato3 & "X720" & dato2 & dato4
        case "400"    
            ResultadoUF = dato1 & dato3 & "X" & dato5 & dato2 
    end select        
    MoverBloque ResultadoUF, "Ufuncionales", wsActual, Celda
End Sub
Public Sub Unidades_funcionales_lisas()
    '========================
        Celda = "NS132"
        ThisWorkbook.Sheets("DISENO").Range("NS132:QJ182").Clear
        Set wsActual = ThisWorkbook.Sheets("DISENO")
    '========================
    Dim dato1 As String, dato2 As String, dato3 As String
    dato1 = Me.OLEObjects("TipoLSS").Object.Value
    Select Case dato1
        Case "SFI - Frontal Interno": dato1 = "L_"
        Case "SIA - Integral Abatible", "SFA - Fijo Atornillado": dato1 = "LA_"
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
        ThisWorkbook.Sheets("DISENO").Range("NS132:QJ182").Clear
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
        ThisWorkbook.Sheets("DISENO").Range("NS132:QJ182").Clear
        Set wsActual = ThisWorkbook.Sheets("DISENO")
    '========================
    Dim dato1 As String, dato2 As String, dato3 As String, dato4 As String
    dato1 = Me.OLEObjects("TipoUNDPS").Object.Value
    Select Case dato1
        Case "SFI - Frontal Interno": dato1 = "D"
        Case "SIA - Integral Abatible": dato1 = "DA"
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

    ResultadoDPS = dato1 & dato2 & "X" & dato3 & "_" & dato4
    MoverBloque ResultadoDPS, "DPS", wsActual, Celda
End Sub
Public Sub Equipos_Control()
    '========================
        Celda = "NS132"
        ThisWorkbook.Sheets("DISENO").Range("NS132:QJ182").Clear
        Set wsActual = ThisWorkbook.Sheets("DISENO")
    '========================
    Dim dato1 As String, dato2 As String
    dato1 = Me.OLEObjects("EquipoCTR").Object.Value
    dato2 = Me.OLEObjects("MedidasCTR").Object.Value
    ResultadoCT = dato1 & "_" & dato2
    MoverBloque ResultadoCT, "UCONTROL", wsActual, Celda
End Sub
Public Sub Accesorios_CTR()
    '========================
        Celda = "NS132"
        ThisWorkbook.Sheets("DISENO").Range("NS132:QJ182").Clear
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
Public Sub Equipo_Individual()
    '========================
        Celda = "NS132"
        ThisWorkbook.Sheets("DISENO").Range("NS132:QJ182").Clear
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
        ThisWorkbook.Sheets("DISENO").Range("NS132:QJ182").Clear
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
        ThisWorkbook.Sheets("DISENO").Range("NS132:QJ182").Clear
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
        ThisWorkbook.Sheets("DISENO").Range("NU107:QJ182").Clear
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
        ThisWorkbook.Sheets("DISENO").Range("NS104:QJ182").Clear
        Set wsActual = ThisWorkbook.Sheets("DISENO")
    '========================
    Dim dato1 As String, dato2 As String, dato3 As String, dato4 As String
    dato1 = Me.OLEObjects("OrientacionBRR").Object.Value
    Select Case dato1
        Case "Horizontal": dato1 = "H"
        Case "Vertical": dato1 = "V"
    End Select
    dato2 = Me.OLEObjects("TipoBRR").Object.Value
    Select Case dato2
        Case "BusBar": dato2 = "BB"
        Case "Tierra Neutro": dato2 = "TN"
    End Select
    dato3 = Me.OLEObjects("AmperajeBRR").Object.Value
    dato4 = Me.OLEObjects("MedidaBRR").Object.Value
    ResultadoBAR = "UN_" & dato1 & dato2 & "X" & dato4 & "_" & dato3
    MoverBloque ResultadoBAR, "BARRAJE", wsActual, Celda
End Sub
Public Sub Transferencia()
    '========================
        Celda = "NU107"
        ThisWorkbook.Sheets("DISENO").Range("NU107:QJ182").Clear
        Set wsActual = ThisWorkbook.Sheets("DISENO")
    '========================
    Dim dato1 As String, dato2 As String
    dato1 = Me.OLEObjects("EquipoTRF").Object.Value
    Select Case dato1
        Case "Trf. EZC250 Vertical entrada Superior": dato1 = "EZC250XVX720XSUP"
        Case "Trf. EZC250 Vertical entrada Inferior": dato1 = "EZC250XVX720XINF"
        Case "Trf. EZC630 Vertical entrada Superior": dato1 = "EZC630XVX720XSUP"
        Case "Trf. EZC630 Vertical entrada Inferior": dato1 = "EZC630XVX720XINF"
        Case "Trf. CVS250 Vertical entrada Superior": dato1 = "CVS250XVX720XSUP"
        Case "Trf. CVS250 Vertical entrada Inferior": dato1 = "CVS250XVX720XINF"
        Case "Trf. CVS630 Vertical entrada Superior": dato1 = "CVS630XVX720XSUP"
        Case "Trf. CVS630 Vertical entrada Inferior": dato1 = "CVS630XVX720XINF"
        Case "Trf. NSX250 Vertical entrada Superior": dato1 = "NSX250XVX720XSUP"
        Case "Trf. NSX250 Vertical entrada Inferior": dato1 = "NSX250XVX720XINF"
        Case "Trf. NSX630 Vertical entrada Superior": dato1 = "NSX630XVX720XSUP"
        Case "Trf. NSX630 Vertical entrada Inferior": dato1 = "NSX630XVX720XINF"
    End Select
    dato2 = Me.OLEObjects("OrientacionTRF").Object.Value
    Select Case dato2
        Case "Vertical"
        ResultadoTRF = "UT_" & dato1
        Case "Horizontal"
            dato2 = "H"
            ResultadoTRF = "UT_" & dato1 & "X" & dato2 & "X720"
        Case "Horizontal Abatible"
            dato2 = "HA"
            ResultadoTRF = "UT_" & dato1 & "X" & dato2 & "X720"
    End Select
    MoverBloque ResultadoTRF, "U_TRANSFERENCIA", wsActual, Celda

End Sub
Public Sub ACESORIOS_MTM()
    '========================
        Celda = "NS104"
        ThisWorkbook.Sheets("DISENO").Range("NS104:QJ182").Clear
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
        ThisWorkbook.Sheets("DISENO").Range("NS104:QJ182").Clear
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
        ThisWorkbook.Sheets("DISENO").Range("NS104:QJ182").Clear
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
Public Sub Cambiodecolor() 
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

