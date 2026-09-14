'=================== VISTA PUERTA =================
Public Altura As String
Public Ancho As String, AnchoLISAS As String
Public Profundidad As String
Public IP As String, Color As String

Public Sub Worksheet_Activate()
    Extraer_datos
    '============== PLACAS DE CONTROL ===================
        With Me.OLEObjects("TamanoPC").Object
            .Clear
            .AddItem "Pequeña"
            '.AddItem "Mediana"
            .AddItem "Grande"
        End With
        With Me.OLEObjects("VistaPC").Object
            .Clear
            .AddItem "Frontal"
            .AddItem "Lateral"
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
            .AddItem "Celosidas_Pequeña"
            .AddItem "Celosidas_Grande"
            .AddItem "Chapa_Antivandalica"
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
        Celda = "KR83"
        ThisWorkbook.Sheets("FrontalFM").Range("KR83:NJ185").Clear
        Set wsActual = ThisWorkbook.Sheets("FrontalFM")
    '========================
    Dim dato1 As String, dato2 As String
    dato1 = Me.OLEObjects("AlturaENV").Object.Value
    dato2 = Me.OLEObjects("AnchoENV").Object.Value
    ResultadoENV = "EFM_" & dato1 & "X" & dato2
    MoverBloque ResultadoENV, "ENVOLVENTES", wsActual, Celda
End Sub
Public Sub ACESORIOS_MTM()
    '========================
        Celda = "NS104"
        ThisWorkbook.Sheets("FrontalFM").Range("KR83:NJ185").Clear
        Set wsActual = ThisWorkbook.Sheets("FrontalFM")
    '========================
    Dim dato1 As String, dato2 As String
    dato1 = Me.OLEObjects("TipoACC").Object.Value
    dato2 = Me.OLEObjects("MedidaACC").Object.Value
    ResultadoACM = "ACC_" & dato1 & "X" & dato2
    MoverBloque ResultadoACM, "ACCESOSRIOS MTM", wsActual, Celda
End Sub
Public Sub Placas_Control()
    '========================
        Celda = "NS132"
        ThisWorkbook.Sheets("FrontalFM").Range("KR83:NJ185").Clear
        Set wsActual = ThisWorkbook.Sheets("FrontalFM")
    '========================
    Dim dato1 As String, dato2 As String, dato3 As String, dato4 As String
    dato1 = Me.OLEObjects("TamanoPC").Object.Value
    Select Case dato1
        Case "Pequeña": dato1 = "PQ"
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