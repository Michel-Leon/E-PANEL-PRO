'============== vista columna ===================
Public valordimensiones As String

Private Sub Worksheet_Activate()
    extraerDatos
    With Me.OLEObjects("Dimensiones").Object
        .Clear
        .AddItem valordimensiones
    End With

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
    MoverBloque ResultadoENV,"ENVOLVENTES_FM"
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
    MoverBloque ResultadoDPS,"BD_DPS"
End Sub




