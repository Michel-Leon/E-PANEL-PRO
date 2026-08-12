'============== vista columna ===================
Public valordimensiones As String
'======= CONDICIONES DE SELECCION =======
'BARRAJE
    Private sub TipoBRR_Change()
        dim dato as string
        dato = Me.OleObjects("TipoBRR").Object.Value
        ' CAMBIA CNTENIDO DE UNA CUADRO DE TEXTO  concepto_dato3 DEPENDIENDO DE LA SELECCION DEL COMBOBOX
        Select case dato
            case "Tierra/Neutro"
                ActiveSheet.Shapes("concepto_dato3").TextFrame2.TextRange.Text = "Medida:"
            case "BusBar"
                IF Me.OleObjects("OrientacionBRR").Object.Value = "Horizontal" Then
                    ActiveSheet.Shapes("concepto_dato3").TextFrame2.TextRange.Text = "Medida:"
                ElseIF Me.OleObjects("OrientacionBRR").Object.Value = "Vertical" Then
                    ActiveSheet.Shapes("concepto_dato3").TextFrame2.TextRange.Text = "Interruptor:"
                End If
        end select
        with Me.oleobjects("DATO3BRR").Object
            .Clear
            Select case dato
                case "Tierra/Neutro"
                    .AddItem "720"
                    .AddItem "400"
                case "BusBar"
                IF Me.OleObjects("OrientacionBRR").Object.Value = "Vertical" Then
                    .AddItem "EZC400"
                    .AddItem "CVS400"
                    .AddItem "NSX400"
                ELSEIF Me.OleObjects("OrientacionBRR").Object.Value = "Horizontal" Then
                    .AddItem "720"
                End If    
            end select
        end with
    End sub
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
        If bloqueado Then Exit Sub
        bloqueado = True

        With Combo("OrientacionUF")
            LimpiarCombo Combo("OrientacionUF")
            Llenar Combo("OrientacionUF"), "Horizontal|Horizontal Compuesta"
            If Me.Carga.Value Then
                Llenar Combo("OrientacionUF"), "Vertical|Vertical Compacta"
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
            Case "Horizontal"
                Llenar Combo("EquipoUF"), "GoPact100|GoPact250|EZC100|EZC250|CVS250|NSX250"
            Case "Horizontal Compuesta"
                Llenar Combo("EquipoUF"), "GoPact100|GoPact250|EZC100|EZC250|CVS250"
            Case "Vertical"
                Llenar Combo("EquipoUF"), "Mini Breaker|Guardamotor"
            Case "Vertical Compacta"
                Llenar Combo("EquipoUF"), "GoPact100|EZC100"
        End Select

        ' Conexión
        If Me.Totalizador.Value Then
            Llenar Combo("ConexionUF"), "Derecha"
        Else
            Select Case dato
                Case "Horizontal", "Horizontal Compuesta"
                    Llenar Combo("ConexionUF"), "Izquierda|Derecha"
                Case "Vertical", "Vertical Compacta"
                    Llenar Combo("ConexionUF"), "Superior|NA"
            End Select
        End If
    End Sub
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
'ACESORIOS DE CONTROL
Private sub elemaCTR_Change()
    Dim dato As String
    dato = Me.OleObjects("elemaCTR").Object.Value
    with Me.oleobjects("VistaCTR").Object   
        .Clear
        select case dato
            case "Clavija Roja", "Clavija Azul"
                .AddItem "Frontal"
                .AddItem "Derecha_Abierta"
                .AddItem "Derecha_Cerrada"
                .AddItem "Izquierda_Abierta"
                .AddItem "Izquierda_Cerrada"
            case else
                .AddItem "NA"
        end select
    End With
end sub 
' MTO DE TRANSFERENCIA
Private sub TipoTransferencia_change()
    dim dato as string
    Dim texto As String
    dato = Me.OleObjects("TipoTransferencia").Object.Value
    Select Case  dato
        Case "Tranf. EZC250 Vertical entrada superior | Tipo I"
            texto = "+N1| | | | | | H: 1100 / W: 720 / D: 400 | Seccion de Envolvente frente muerto | |" & vbCrLf
            texto = texto & "+C1 | Liso | | | | | W: 720/ U:  3 | Unidad Funcional  Lisa / SFI - Frontal Interno | |" & vbCrLf
            texto = texto & "+C2 | Liso | | | | | W: 720 / U:  4 | Unidad Funcional  Lisa / SFI - Frontal Interno | |" & vbCrLf
            texto = texto & "+C3 | Equipada | EZC250 | 3 | Vertical | EZC250N| W: 720 / U: 7 | SFI - Frontal Interno / Vertical/ Cant.equipos: 2 | Superior |" & vbCrLf
            texto = texto & "+C4 | Liso | | | | | W: 720 / U:  4 | Unidad Funcional  Lisa / SFI - Frontal Interno | |" & vbCrLf
        Case "Tranf. EZC250 Vertical entrada superior | Tipo II"
            texto = "+N1| | | | | | H: 900 / W: 720 / D: 400 | Seccion de Envolvente frente muerto | |" & vbCrLf
            texto = texto & "+C1 | Liso | | | | | W: 720/ U:  3 | Unidad Funcional  Lisa / SFI - Frontal Interno | |" & vbCrLf
            texto = texto & "+C2 | Equipada | EZC250 | 3 | Vertical | EZC250N| W: 720 / U: 7 | SFI - Frontal Interno / Vertical/ Cant.equipos: 2 | Superior |" & vbCrLf
            texto = texto & "+C3 | Liso | | | | | W: 720 / U:  4 | Unidad Funcional  Lisa / SFI - Frontal Interno | |" & vbCrLf
        
        Case "Tranf. EZC250 Vertical entrada inferior"
            texto = "+N1| | | | | | H: 1100 / W: 720 / D: 400 | Seccion de Envolvente frente muerto | |" & vbCrLf
            texto = texto & "+C1 | Liso | | | | | W: 720/ U:  3 | Unidad Funcional  Lisa / SFI - Frontal Interno | |" & vbCrLf
            texto = texto & "+C2 | Liso | | | | | W: 720 / U:  4 | Unidad Funcional  Lisa / SFI - Frontal Interno | |" & vbCrLf
            texto = texto & "+C3 | Equipada | EZC250 | 3 | Vertical | EZC250N| W: 720 / U: 7 | SFI - Frontal Interno / Vertical/ Cant.equipos: 2 | Inferior |" & vbCrLf
            texto = texto & "+C4 | Liso | | | | | W: 720 / U:  4 | Unidad Funcional  Lisa / SFI - Frontal Interno | |" & vbCrLf
        Case "Tranf. EZC100 Vertical entrada superior | Tipo I"
            texto = "+N1| | | | | | H: 1100 / W: 720 / D: 400 | Seccion de Envolvente frente muerto | |" & vbCrLf
            texto = texto & "+C1 | Liso | | | | | W: 720/ U:  3 | Unidad Funcional  Lisa / SFI - Frontal Interno | |" & vbCrLf
            texto = texto & "+C2 | Liso | | | | | W: 720 / U:  4 | Unidad Funcional  Lisa / SFI - Frontal Interno | |" & vbCrLf
            texto = texto & "+C2 | Liso | | | | | W: 720 / U:  4 | Unidad Funcional  Lisa / SFI - Frontal Interno | |" & vbCrLf
            texto = texto & "+C4 | Equipada | EZC100 | 3 | Vertical | EZC100N| W: 720 / U: 7 | SFI - Frontal Interno / Vertical/ Cant.equipos: 2 | Superior |" & vbCrLf
        case "Tranf. EZC100 Vertical entrada superior | Tipo II"
            texto = "+N1| | | | | | H: 900 / W: 720 / D: 400 | Seccion de Envolvente frente muerto | |" & vbCrLf
            texto = texto & "+C1 | Liso | | | | | W: 720/ U:  3 | Unidad Funcional  Lisa / SFI - Frontal Interno | |" & vbCrLf
            texto = texto & "+C2 | Equipada | EZC100 | 3 | Vertical | EZC100N| W: 720 / U: 7 | SFI - Frontal Interno / Vertical/ Cant.equipos: 2 | Superior |" & vbCrLf
            texto = texto & "+C3 | Liso | | | | | W: 720 / U:  4 | Unidad Funcional  Lisa / SFI - Frontal Interno | |" & vbCrLf
        case "Tranf. EZC100 Vertical entrada inferior"
            texto = "+N1| | | | | | H: 1100 / W: 720 / D: 400 | Seccion de Envolvente frente muerto | |" & vbCrLf
            texto = texto & "+C1 | Liso | | | | | W: 720/ U:  3 | Unidad Funcional  Lisa / SFI - Frontal Interno | |" & vbCrLf
            texto = texto & "+C2 | Liso | | | | | W: 720 / U:  4 | Unidad Funcional  Lisa / SFI - Frontal Interno | |" & vbCrLf
             texto = texto & "+C3 | Equipada | EZC100 | 3 | Vertical | EZC100N| W: 720 / U: 7 | SFI - Frontal Interno / Vertical/ Cant.equipos: 2 | Inferior |" & vbCrLf
            texto = texto & "+C4 | Liso | | | | | W: 720 / U:  4 | Unidad Funcional  Lisa / SFI - Frontal Interno | |" & vbCrLf
        case "Tranf. CVS250 Vertical entrada superior | Tipo I"
            texto = "+N1| | | | | | H: 1100 / W: 720 / D: 400 | Seccion de Envolvente frente muerto | |" & vbCrLf
            texto = texto & "+C1 | Liso | | | | | W: 720/ U:  3 | Unidad Funcional  Lisa / SFI - Frontal Interno | |" & vbCrLf
            texto = texto & "+C2 | Liso | | | | | W: 720 / U:  4 | Unidad Funcional  Lisa / SFI - Frontal Interno | |" & vbCrLf
            texto = texto & "+C3 | Equipada | CVS250 | 3 | Vertical | CVS250N| W: 720 / U: 7 | SFI - Frontal Interno / Vertical/ Cant.equipos: 2 | Superior |" & vbCrLf
            texto = texto & "+C4 | Liso | | | | | W: 720 / U:  4 | Unidad Funcional  Lisa / SFI - Frontal Interno | |" & vbCrLf 
        case "Tranf. CVS250 Vertical entrada superior | Tipo II"
            texto = "+N1| | | | | | H: 900 / W: 720 / D: 400 | Seccion de Envolvente frente muerto | |" & vbCrLf
            texto = texto & "+C1 | Liso | | | | | W: 720/ U:  3 | Unidad Funcional  Lisa / SFI - Frontal Interno | |" & vbCrLf
            texto = texto & "+C2 | Equipada | CVS250 | 3 | Vertical | CVS250N| W: 720 / U: 7 | SFI - Frontal Interno / Vertical/ Cant.equipos: 2 | Superior |" & vbCrLf
            texto = texto & "+C3 | Liso | | | | | W: 720 / U:  4 | Unidad Funcional  Lisa / SFI - Frontal Interno | |" & vbCrLf
        case "Tranf. CVS250 Vertical entrada inferior"
            texto = "+N1| | | | | | H: 1100 / W: 720 / D: 400 | Seccion de Envolvente frente muerto | |" & vbCrLf
            texto = texto & "+C1 | Liso | | | | | W: 720/ U:  3 | Unidad Funcional  Lisa / SFI - Frontal Interno | |" & vbCrLf
            texto = texto & "+C2 | Liso | | | | | W: 720 / U:  4 | Unidad Funcional  Lisa / SFI - Frontal Interno | |" & vbCrLf
            texto = texto & "+C3 | Equipada | CVS250 | 3 | Vertical | CVS250N| W: 720 / U: 7 | SFI - Frontal Interno / Vertical/ Cant.equipos: 2 | Inferior |" & vbCrLf
            texto = texto & "+C4 | Liso | | | | | W: 720 / U:  4 | Unidad Funcional  Lisa / SFI - Frontal Interno | |" & vbCrLf 
        case "Tranf. NSX250 Vertical entrada superior | Tipo I"
            texto = "+N1| | | | | | H: 1100 / W: 720 / D: 400 | Seccion de Envolvente frente muerto | |" & vbCrLf
            texto = texto & "+C1 | Liso | | | | | W: 720/ U:  3 | Unidad Funcional  Lisa / SFI - Frontal Interno | |" & vbCrLf
            texto = texto & "+C2 | Liso | | | | | W: 720 / U:  4 | Unidad Funcional  Lisa / SFI - Frontal Interno | |" & vbCrLf
            texto = texto & "+C3 | Equipada | NSX250 | 3 | Vertical | NSX250N| W: 720 / U: 7 | SFI - Frontal Interno / Vertical/ Cant.equipos: 2 | Superior |" & vbCrLf
            texto = texto & "+C4 | Liso | | | | | W: 720 / U:  4 | Unidad Funcional  Lisa / SFI - Frontal Interno | |" & vbCrLf 
        case "Tranf. NSX250 Vertical entrada superior | Tipo II"
            texto = "+N1| | | | | | H: 900 / W: 720 / D: 400 | Seccion de Envolvente frente muerto | |" & vbCrLf
            texto = texto & "+C1 | Liso | | | | | W: 720/ U:  3 | Unidad Funcional  Lisa / SFI - Frontal Interno | |" & vbCrLf
            texto = texto & "+C2 | Equipada | NSX250 | 3 | Vertical | NSX250N| W: 720 / U: 7 | SFI - Frontal Interno / Vertical/ Cant.equipos: 2 | Superior |" & vbCrLf
            texto = texto & "+C3 | Liso | | | | | W: 720 / U:  4 | Unidad Funcional  Lisa / SFI - Frontal Interno | |" & vbCrLf
        case "Tranf. NSX250 Vertical entrada inferior"
            texto = "+N1| | | | | | H: 1100 / W: 720 / D: 400 | Seccion de Envolvente frente muerto | |" & vbCrLf
            texto = texto & "+C1 | Liso | | | | | W: 720/ U:  3 | Unidad Funcional  Lisa / SFI - Frontal Interno | |" & vbCrLf
            texto = texto & "+C2 | Liso | | | | | W: 720 / U:  4 | Unidad Funcional  Lisa / SFI - Frontal Interno | |" & vbCrLf
            texto = texto & "+C3 | Equipada | NSX250 | 3 | Vertical | NSX250N| W: 720 / U: 7 | SFI - Frontal Interno / Vertical/ Cant.equipos: 2 | Inferior |" & vbCrLf
            texto = texto & "+C4 | Liso | | | | | W: 720 / U:  4 | Unidad Funcional  Lisa / SFI - Frontal Interno | |" & vbCrLf   
        Case else
            texto = "Seleccion no existente"    
    End Select
           
    Me.OleObjects("Descripcion").Object.value = texto  
end sub        
Private Sub Worksheet_Activate()
    extraerDatos
    '============== vista columna ===================
    With Me.OLEObjects("Dimensiones").Object
        .Clear
        .AddItem valordimensiones
    End With
    '============== DPS ===================
    With Me.OleObjects("FabricanteDPS").Object
        .Clear
        .AddItem "Ciprotec"
        .AddItem "HWA"
        .AddItem "Phoenix Contac"
        .AddItem "Citel"
        .AddItem "Schneider Electric"
    End With
    With Me.OleObjects("EquipoDPS").Object
        .Clear
        .AddItem "CVS"
        .AddItem "EZC"
        .AddItem "MINI"
        .AddItem "NSX"
    End With
    '============== Servicios ===================
    With Me.OleObjects("Clase_servicio").Object
        .Clear
        .AddItem "Iluminacion"
        .AddItem "Cambio_Tension"
        .AddItem "Switch"
        .AddItem "GateWay"
        .AddItem "Toma_Corriente"
        .AddItem "Batery.ST"
        .AddItem "Pont_IO"
        .AddItem "A_UnitBB"
    End With 
    '============== Barraje ===================
    With Me.OleObjects("OrientacionBRR").Object
        .Clear
        .AddItem "Horizontal"
        .AddItem "Vertical"
    End With
    with Me.OleObjects("TipoBRR").Object
        .Clear
        .AddItem "Tierra/Neutro"
        .AddItem "BusBar"
    End With
    '============== Lisas ===================
    with Me.oleobjects("MedidaLSS").Object
        .Clear
        .AddItem "720"
        .AddItem "400"
    End With
    with Me.oleobjects("UnidadLSS").Object
        .Clear
        .AddItem "1"
        .AddItem "2"
        .AddItem "3"
        .AddItem "4"
        .AddItem "5"
        .AddItem "6"
    End With
    '============== Unidades Funcionales ===================
    Me.OleObjects("TipoUnidad").Object.Value = "SFI - Frontal Interno"
    Me.Carga.Value = false
    Me.Totalizador.Value = false
    '============== Contactores ===================
    with Me.OleObjects("EquipoCT").Object
        .Clear
        .AddItem "A-Unit (Barra + Mini)"
        .AddItem "Contactores"
    End With
    '============== UNIDADES DE CONTROL ===================
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
    With Me.OleObjects("TamanoCTR").Object
        .Clear
        .AddItem "400"
        .AddItem "720"
    End With
    '============== ACCESORIOS DE CONTROL ===================
    With Me.OleObjects("elemaCTR").Object
        .Clear
        .AddItem "Clavija Roja"
        .AddItem "Clavija Azul"
        .AddItem "Pilotos"
        .AddItem "Parada Emergencia"
        .AddItem "Conmutador"
        .AddItem "Selector 1"
        .AddItem "Selector 2"
        .AddItem "Pulsador"
        .AddItem "Piloto verde"
        .AddItem "Piloto rojo"
        .AddItem "Piloto amarillo"
        .AddItem "Piloto azul"
        .AddItem "Piloto gris"
        .AddItem "Piloto gris x3"
    end With 
    '============== EQUIPOS INDIVIDUALES ===================
    With Me.OleObjects("EquipoEI").Object
        .Clear
        .AddItem "MiniBreaker (3P)"
        .AddItem "MiniBreaker (2P)"
        .AddItem "MiniBreaker (1P)"
        .AddItem "EZC100 (3P)"
        .AddItem "EZC100 (2P)"
        .AddItem "Bloque de distribuccion"
        .AddItem "Contactor (Talla 1-2)"
        .AddItem "Contactor (Talla 3)"
        .AddItem "Transformador"
    end With
    '============== MTO TRANSFERENCIAS ===================
    With Me.OleObjects("TipoTransferencia").Object
        .Clear
        .AddItem "Tranf. EZC250 Vertical entrada superior | Tipo I" '1100x720
        .AddItem "Tranf. EZC250 Vertical entrada superior | Tipo II"'900x720
        .AddItem "Tranf. EZC250 Vertical entrada inferior"
        .AddItem "Tranf. EZC100 Vertical entrada superior | Tipo I"'1100x720
        .AddItem "Tranf. EZC100 Vertical entrada superior | Tipo II"'900x720
        .AddItem "Tranf. EZC100 Vertical entrada inferior"
        .AddItem "Tranf. CVS250 Vertical entrada superior | Tipo I"'1100x720
        .AddItem "Tranf. CVS250 Vertical entrada superior | Tipo II"'900x720
        .AddItem "Tranf. CVS250 Vertical entrada inferior"
        .AddItem "Tranf. NSX250 Vertical entrada superior | Tipo I"'1100x720
        .AddItem "Tranf. NSX250 Vertical entrada superior | Tipo II"'900x720
    end With
End Sub

Sub extraerDatos()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Portada_Mestra")
    Dim partes() As String

    valordimensiones = ws.Range("C14").Value

    If InStr(1, valordimensiones, "X", vbTextCompare) = 0 Then
        MsgBox "Portada_Mestra!C14 no tiene el formato esperado ('NúmeroXNúmero'). Valor actual: '" & valordimensiones & "'", vbExclamation
        Exit Sub
    End If

    partes = Split(valordimensiones, "X")
    If UBound(partes) < 1 Then
        MsgBox "Portada_Mestra!C14 no tiene el formato esperado ('NúmeroXNúmero'). Valor actual: '" & valordimensiones & "'", vbExclamation
        Exit Sub
    End If

    valordimensiones = partes(0) & "X" & partes(1)   ' ahora sí queda "900X520"
End Sub
Public sub Envolventes()
    On Error GoTo errHandler
    With ThisWorkbook.Sheets("Vista FM").Range("BX44:IX144")
        .Clear
        .UnMerge
    End With
    dim dato as string
    dato = Me.oleobjects("Dimensiones").Object.Value
    Linea = thisWorkbook.Sheets("Cabecera").Range("E3").Value
    Set wsActual = thisWorkbook.Sheets("Vista FM")
    Select case Linea
        case "G-Flex Basic":Linea = "Basic"
        Case "G-Flex Neo":Linea = "Neo"
    end select
    Select case thisWorkbook.Sheets("Vista FM").Range("A1").Value  
        Case "600":Celda = "BX97"
        Case "700":Celda = "BX89"
        Case "900":Celda ="BX73"
        Case "800":Celda = "BX81"
        Case "110":Celda = "BX57"
    End select    

    ResultadoENV = "EFM_" & Linea & "X" & dato
    'MSGBOX ResultadoENV
    MoverBloque ResultadoENV,"ENVOLVENTES_FM", wsActual, celda
    '===== ARMAR LAS DEMAS VISTAS =====
    Dim datoH As String, datoW As String, datoD As String
    '------------------------------------------
    ResultadoENVP = "EP_" & Linea & "X" & dato
    Vista_puerta ResultadoENVP
    '------------------------------------------
    datoH = ThisWorkbook.Sheets("Vista FM").Range("A1").Value
    datoD = ThisWorkbook.Sheets("Vista FM").Range("A2").Value
    ResultadoENVL = "EL_" & Linea & "X" & datoH & "X" & datoD
    Vista_lateral ResultadoENVL
    '------------------------------------------
    ResultadoENVPS = "EPT_" & Linea & "X" & dato
    Vista_Posterior ResultadoENVPS
    'Msgbox ResultadoENVPS
    '------------------------------------------
    'MSGBOX ResultadoENVL & "/" & ResultadoENVP
    Exit Sub
    errHandler:
        MsgBox "Error " & Err.Number & " en Envolventes: " & Err.Description, vbCritical
End Sub
Public Sub DPS_Buscar()
    '========================
        celda = "JT107"
        thisWorkbook.Sheets("Vista FM").Range("JT107:NO182").Clear
        Set wsActual = thisWorkbook.Sheets("Vista FM")
    '========================
    dim Fabricante as string
    select case Me.oleobjects("FabricanteDPS").Object.Value
        case "Ciprotec":Fabricante = "CP"
        case "HWA":Fabricante = "HW"
        case "Phoenix Contac":Fabricante = "PC"
        case "Citel":Fabricante = "CT"
        case "Schneider Electric":Fabricante = "SH"
    end select
    dim Equipo as string
    select case Me.oleobjects("EquipoDPS").Object.Value
        case "CVS":Equipo = "CVS"
        case "EZC":Equipo = "EZC"
        case "MINI":Equipo = "MB"
        case "NSX":Equipo = "NSX"
    end select
    ResultadoDPS= "DPS_" & Fabricante & "X" & Equipo & "X" & "720X4"
    MoverBloque ResultadoDPS,"BD_DPS", wsActual, celda
End Sub
Public sub Servicios()
    '========================
        celda = "JT107"
        thisWorkbook.Sheets("Vista FM").Range("JT107:NO182").Clear
        Set wsActual = thisWorkbook.Sheets("Vista FM")
    '========================
    dim servicio as string
    servicio = Me.oleobjects("Clase_servicio").Object.Value
    ResultadoSERV= servicio
    MoverBloque ResultadoSERV,"SERVICIOS", wsActual, celda
End Sub
Public Sub Barraje()
    Dim wsActual As Worksheet
    Dim celda As String
    Dim dato1 As String, dato2 As String, dato3 As String
    Dim ResultadoBRR As String
    '========================
    celda = "JT95"
    Set wsActual = ThisWorkbook.Sheets("Vista FM")
    wsActual.Range("JT95:NO182").Clear
    '========================
    dato1 = Trim$(Me.OLEObjects("OrientacionBRR").Object.Value)
    dato2 = Trim$(Me.OLEObjects("TipoBRR").Object.Value)
    dato3 = Trim$(Me.OLEObjects("DATO3BRR").Object.Value)

    If dato1 = "" Or dato2 = "" Or dato3 = "" Then
        MsgBox "Completa las tres selecciones antes de generar el barraje.", vbExclamation
        Exit Sub
    End If

    Select Case dato1
        Case "Horizontal": dato1 = "H"
        Case "Vertical":   dato1 = "V"
    End Select

    Select Case dato2
        Case "Tierra/Neutro": dato2 = "TN"
        Case "BusBar":        dato2 = "BB"
    End Select

    ResultadoBRR = dato1 & dato2 & "X" & dato3

    MoverBloque ResultadoBRR, "BD_BARRAJE", wsActual, celda
End Sub
Public sub Unidades_lisas()
    '========================
        celda = "JT107"
        thisWorkbook.Sheets("Vista FM").Range("JT107:NO182").Clear
        Set wsActual = thisWorkbook.Sheets("Vista FM")
    '========================
    dim dato1 as string, dato2 as string
    dato1 = Me.oleobjects("MedidaLSS").Object.Value
    dato2 = Me.oleobjects("UnidadLSS").Object.Value
    ResultadoLSS= "LX" & dato1 & "X" & dato2
    MoverBloque ResultadoLSS,"UF_LISAS", wsActual, celda
End Sub
public sub Unidades_funcionales()
    '========================
        celda = "JT107"
        thisWorkbook.Sheets("Vista FM").Range("JT107:NO182").Clear
        Set wsActual = thisWorkbook.Sheets("Vista FM")
    '========================
    dim dato1 as string, dato2 as string, dato3 as string

    select case Me.oleobjects("OrientacionUF").Object.Value
        case "Horizontal":dato1 = "H"
        case "Horizontal Compuesta":dato1 = "HC"
        case "Vertical":dato1 = "V"
        case "Vertical Compacta":dato1 = "VC"
    end select
    select case Me.oleobjects("EquipoUF").Object.Value
        case "GoPact100":dato2 = "GoPact100"
        case "GoPact250":dato2 = "GoPact250"
        case "EZC100":dato2 = "EZC100"
        case "EZC250":dato2 = "EZC250"
        case "CVS250":dato2 = "CVS250"
        case "NSX250":dato2 = "NSX250"
        case "Mini Breaker":dato2 = "MB"
        case "Guardamotor":dato2 = "GM"
    end select
    select case Me.oleobjects("ConexionUF").Object.Value
        Case "Izquierda":dato3 = "IZ"
        Case "Derecha":dato3 = "DR"
        Case "Superior":dato3 = "SP"
        Case "NA":dato3 = "NA"
    end select
    ResultadoUF= "UN_" & dato2 & "X" & dato1 & "X720" & "_" & dato3
    MoverBloque ResultadoUF,"U_FUNCIONALES", wsActual, celda
End Sub
Public Sub Contactores()
    Dim wsActual As Worksheet
    Dim celda As String
    Dim dato1 As String, dato2 As String, dato3 As String
    Dim ResultadoCT As String

    Set wsActual = ThisWorkbook.Sheets("Vista FM")
    celda = "JT107"

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
    Debug.Print ResultadoCT          ' revísalo en Ventana Inmediato (Ctrl+G)

    If Not NombreExiste(ResultadoCT) Then
        MsgBox "No existe el bloque: " & ResultadoCT, vbExclamation
        Exit Sub
    End If

    wsActual.Range("JT107:NO182").Clear
    wsActual.Activate
    wsActual.Range("A1").Select

    MoverBloque ResultadoCT, "U_FUNCIONALES", wsActual, celda
End Sub
Private Function NombreExiste(n As String) As Boolean
    On Error Resume Next
    NombreExiste = Not ThisWorkbook.Names(n) Is Nothing
    On Error GoTo 0
End Function
Public sub Unidad_control()
    '========================
        celda = "JT107"
        thisWorkbook.Sheets("Vista FM").Range("JT107:NO182").Clear
        Set wsActual = thisWorkbook.Sheets("Vista FM")
    '========================
    dim dato1 as string, dato2 as string
    dato1 = Me.oleobjects("EquipoCTR").Object.Value
    dato2 = Me.oleobjects("TamanoCTR").Object.Value
    ResultadoUC=  dato1 & "_" & dato2
    MoverBloque ResultadoUC,"Control", wsActual, celda
End sub
Public sub Accesorios_Control()
    '========================
        celda = "JT107"
        thisWorkbook.Sheets("Vista FM").Range("JT107:NO182").Clear
        Set wsActual = thisWorkbook.Sheets("Vista FM")
    '========================
    dim dato1 as string, dato2 as string
    Select case Me.OleObjects("elemaCTR").Object.Value
        Case "Clavija Roja":dato1 = "ClavijaRoja"
        Case "Clavija Azul":dato1 = "ClavijaAzul"
        Case "Pilotos":dato1 = "Pilotos"
        Case "Parada Emergencia":dato1 = "Parada_Emergencia"
        Case "Conmutador":dato1 = "Conmutador"
        Case "Selector 1":dato1 = "Selector1"
        Case "Selector 2":dato1 = "Selector2"
        Case "Pulsador":dato1 = "Pulsador"
        Case "Piloto verde":dato1 = "Piloto_Verde"
        Case "Piloto rojo":dato1 = "Piloto_Rojo"
        Case "Piloto amarillo":dato1 = "Piloto_Amarillo"
        Case "Piloto azul":dato1 = "Piloto_Azul"
        Case "Piloto gris":dato1 = "Piloto_gris"
        Case "Piloto gris x3":dato1 = "Piloto_grisx3"
    End Select
    Select case Me.OleObjects("VistaCTR").Object.Value
        case "Frontal": dato2 = "Frontal"
        case "Derecha_Abierta": dato2 = "DA"
        case "Derecha_Cerrada": dato2 = "DC"
        case "Izquierda_Abierta": dato2 = "IA"
        case "Izquierda_Cerrada": dato2 = "IC"
        case "NA": dato2 = "NA"
    End Select
    ResultadoACC= "AC_" & dato1 & "_" & dato2
    MoverBloque ResultadoACC,"Control", wsActual, celda
End sub
Public sub Equipos_individuales()
    '========================
        celda = "JT107"
        thisWorkbook.Sheets("Vista FM").Range("JT107:NO182").Clear
        Set wsActual = thisWorkbook.Sheets("Vista FM")
    '========================
    Dim dato1 As String
    select case  Me.OleObjects("EquipoEI").Object.Value
        Case "MiniBreaker (3P)":dato1 = "MBX3P"
        Case "MiniBreaker (2P)":dato1 = "MBX2P"
        Case "MiniBreaker (1P)":dato1 = "MBX1P"
        Case "EZC100 (3P)":dato1 = "EZC100X3P"
        Case "EZC100 (2P)":dato1 = "EZC100X2P"
        Case "Bloque de distribuccion":dato1 = "BD"
        Case "Contactor (Talla 1-2)":dato1 = "ContactorT1_2"
        Case "Contactor (Talla 3)":dato1 = "ContactorT3"
        Case "Transformador":dato1 = "Transformador"
    end select
    ResultadoACC = "EQ_" & dato1
    MoverBloque ResultadoACC, "Control", wsActual, celda
End sub
