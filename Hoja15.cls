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
    Private Sub FrameUF_Change()  
      Dim dato As String
      dato = Me.OleObjects("FrameUF").Object.Value
        with Me.OleObjects("MedidaUF").Object
          .Clear
            select case dato
                case "MiniBreaker", "GoPact100", "GoPact150", "GoPact200", "GoPact250", _
                 "EZC100", "EZC250", "CVS100", "NSX100", "NSX250"
                 .AddItem "400"
                Case "CVS250", "MiniBreaker_ONKA"
                    .AddItem "400"
                    .AddItem "720"
                case else 
                .AddItem "720"
            end select
        end with
        With Me.OleObjects("CorienteUF").Object
            .Clear
            Select Case Dato
                Case "MiniBreaker", "MiniBreaker_ONKA"
                    .AddItem "1-63A"
                Case "EZC100", "CVS100", "NSX100"
                    .AddItem "100A"
                Case "GoPact150", "GoPact100"
                    .AddItem "100A"
                    .AddItem "125A"
                    .AddItem "150A"
                Case "GoPact200"
                    .AddItem "200A"
                Case "EZC250", "CVS250", "NSX250", "GoPact250"
                    .AddItem "250A"
                Case "GoPact400"
                    .AddItem "400A"
            End Select
        End With
        WITH Me.OleObjects("OrientacionUF").Object
            .Clear
            Select case dato
                Case "CVS100", "GoPact150", "MiniBreaker", "MiniBreaker_ONKA"
                    .AddItem "Vertical"
                case "CVS250", "EZC100", "EZC250", "NSX250", _
                    "GoPact100", "GoPact250"
                    .AddItem "Vertical"
                    .AddItem "Horizontal"
                    .AddItem "H.Compuesta"
                    .AddItem "V.Compacta"
                Case "GoPact200"  
                    .AddItem "Vertical"
                    .AddItem "Horizontal"
                    .AddItem "H.Compuesta" 
                case "GoPact400"  
                    .AddItem "Horizontal"
                    .AddItem "H.Compuesta"
                Case "NSX100"    
                    .AddItem "V.Compacta"
            end select
        end with
    End Sub
    Private sub OrientacionUF_Change()
        Dim dato As String
        dato = Me.OleObjects("OrientacionUF").Object.Value
        With Me.OleObjects("FuncionUF").Object
            .Clear
            Select Case dato
                Case "Vertical", "V.Compacta", "Horizontal"    
                    .AddItem "Carga"
                case "H.Compuesta"
                    .AddItem "Totalizador"
            end select
        End With
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
        With  Me.OleObjects("TipoUF").Object
            .Clear
            .AddItem "SFI - Frontal Interno"
        End With
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