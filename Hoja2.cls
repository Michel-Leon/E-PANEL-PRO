'========================== DISENO TREN =======================
Public Altura as string
Public Ancho as string
Public Profundidad as string
Public IP as string

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
Private Sub Worksheet_Activate()
    Extraer_datos
    Me.OleObjects("Carga").Object.Value = false
    Me.OleObjects("Totalizador").Object.Value = false
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

END Sub

Public sub Extraer_datos()
    'ESTRAER DATOS DE DIMENSIONES
    Altura = Range("A1").Value
    Ancho = Range("A2").Value
    IP = Range("A4").Value
    Profundidad = Range("A3").Value
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