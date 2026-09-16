'========= Vista_Lat_Post ==========
Public Altura As String
Public Ancho As String, Profundidad As String
Public IP As String, Color As String 
'ENVOLVENTES
    Private sub Ancho_panel_Change()
        dim dato As String
        dato = Me.OLEObjects("Ancho_panel").Object.Value
        with Me.OLEObjects("Disponibilidad_panel").Object
            .Clear
            select case dato
                case "200"
                    .AddItem "No Disponible"
                case else
                    .AddItem "Acometida"
                    .AddItem "No Disponible" 
            end select
        end with
    End sub
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
                Case "Cancamo_Frontal", "Cancamo_Lateral", "Celosidas_Grande", "Celosidas_Pequeña", _
                        "Chapa_Antivandalica", "Hablador_Lateral", "Angulo_Frontal", "Angulo Lateral", _
                        "Cubierta Frontal","Techo_Frontal", "Ventilador_Frontal"
                    .AddItem "NA"
                Case Else
                    .AddItem "No se reconoce" 
                End Select
        End With
    End Sub    

Public Sub Worksheet_Activate()
    Extraer_datos
    '============== Envolventes ============================
        With Me.OLEObjects("Vista_Panel").Object
            .Clear
            .AddItem "Vista superior"
            .AddItem "Vista inferior"
        end With
    '============== Accesorios MTM  =============================
        With Me.OLEObjects("TipoACC").Object
            .Clear
            .AddItem "Techo_Frontal"
            '.AddItem "Techo_Laterial"
            .AddItem "Cancamo_Frontal"
            '.AddItem "Cancamo_Lateral"
            .AddItem "Hablador_Frontal"
            '.AddItem "Hablador_Lateral"
            .AddItem "Angulo_Frontal"
            '.AddItem "Angulo_Lateral"
            .AddItem "Cubierta_Frontal"
            '.AddItem "Cubierta_Lateral"
            .AddItem "Ventilador_Frontal"
            .AddItem "Ventilador_Superior"
            '.AddItem "Ventilador_Lateral"
            .AddItem "Celosidas_Pequeña"
            .AddItem "Celosidas_Grande"
            .AddItem "Chapa_Antivandalica"
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
    Ancho = CStr(Me.Range("A2").Value)
    IP = CStr(Me.Range("A4").Value)
    Profundidad = CStr(Me.Range("A3").Value)

    LlenarCombo Me.OLEObjects("Ancho_panel").Object, Ancho
    LlenarCombo Me.OLEObjects("Profundidad_panel").Object, Profundidad
    'msgbox MenuUnidades
    
End Sub
Public Sub Envolventes()
    '========================
        Celda = "KR90"
        ThisWorkbook.Sheets("Vista_Sup_Inf").Range("KR80:NJ185").Clear
        Set wsActual = ThisWorkbook.Sheets("Vista_Sup_Inf")
    '========================
    Dim dato1 As String, dato2 As String, dato3 As String, dato4 As String
    dato1 = Me.OLEObjects("Vista_Panel").Object.Value
    select case dato1
        case "Vista superior": dato1 = "ES"
        case "Vista inferior": dato1 = "EI"
    end select
    dato2 = Me.OLEObjects("Ancho_panel").Object.Value
    dato3 = Me.OLEObjects("Profundidad_panel").Object.Value
    dato4 = Me.OLEObjects("Disponibilidad_panel").Object.Value
    select case dato4
        case "No Disponible": dato4 = "ND"
        case "Acometida": dato4 = "AC"
    end select
    ResultadoS_I = dato1 &  dato2 & "X" & dato3 & "_" & dato4
    MoverBloque ResultadoS_I, "ENV_S_I_P_L", wsActual, Celda
End Sub
Public Sub ACESORIOS_MTM()
    '========================
        Celda = "KP104"
        ThisWorkbook.Sheets("Vista_Sup_Inf").Range("KR60:NJ185").Clear
        Set wsActual = ThisWorkbook.Sheets("Vista_Sup_Inf")
    '========================
    Dim dato1 As String, dato2 As String
    dato1 = Me.OLEObjects("TipoACC").Object.Value
    dato2 = Me.OLEObjects("MedidaACC").Object.Value
    ResultadoACM = "ACC_" & dato1 & "X" & dato2
    MoverBloque ResultadoACM, "ACCESOSRIOS MTM", wsActual, Celda
End Sub
Public Sub COTAS()
    '========================
        Celda = "KP104"
        ThisWorkbook.Sheets("Vista_Sup_Inf").Range("KR83:NJ185").Clear
        Set wsActual = ThisWorkbook.Sheets("Vista_Sup_Inf")
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
    valorColor = Me.OLEObjects("ColorConf").Object.Value
    Set ws = ThisWorkbook.Sheets("Vista_Sup_Inf")
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