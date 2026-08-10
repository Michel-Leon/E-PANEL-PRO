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
    with Me.oleobjects("MedidaUF").Object
        .Clear
        .AddItem "720"
        .AddItem "400"
    End With
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


