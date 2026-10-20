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
                    .AddItem "1u"
                    .AddItem "2u"
                    .AddItem "3u"
                    .AddItem "4u"
                    .AddItem "5u"
                    .AddItem "6u"
                    .AddItem "7u"
                    .AddItem "8u"
                case "720"
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
Public Sub Worksheet_Activate()
    Extraer_datos
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
        Me.OleObjects("TipoLSS").Object.value = "SFI - Frontal Interno"
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
Public sub Extraer_datos()
    ' EXTRAER DATOS DE DIMENSIONES
    Altura = CStr(Me.Range("A1").Value)
    Ancho = CStr(Me.Range("A2").Value)
    IP = CStr(Me.Range("A4").Value)
    AnchoLISAS= CStr(Me.Range("A5").Value)

    LlenarCombo Me.OLEObjects("AlturaENV").Object, Altura
    'LlenarCombo Me.OLEObjects("AnchoENV").Object, Ancho
    'LlenarCombo Me.OLEObjects("AnchoLSS").Object, AnchoLISAS
    'msgbox MenuUnidades
    
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
   