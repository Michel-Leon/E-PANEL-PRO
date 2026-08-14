'========================== DISENO TREN =======================
Public Altura as string
Public Ancho as string
Public Profundidad as string

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
Private Sub Worksheet_Activate()
    Extraer_datos
    With  Me.OleObjects("AlturaENV").Object
        .AddItem Altura
    End With
    With  Me.OleObjects("AnchoENV").Object
        .AddItem "200"
        .Additem "400"
        .Additem "720"
    End With  
    With  Me.OleObjects("TipoUF").Object
        .AddItem "SFI - Frontal Interno"
        .AddItem "SIA - Integral Abatible"
        .AddItem "SFC - Frontal Compacto"
        .AddItem "SFA - Fijo Atornillado"
    End With
    
END Sub

Public sub Extraer_datos()
    'ESTRAER DATOS DE DIMENSIONES
    Altura = Range("A1").Value
    Ancho = Range("A2").Value
    Profundidad = Range("A3").Value
End Sub

Public sub Envolventes()
    '========================
        celda = "NS98"
        thisWorkbook.Sheets("DISEÑO").Range("NS98:QJ182").Clear
        Set wsActual = thisWorkbook.Sheets("DISEÑO")
    '========================
    dim dato1 as string, dato2 as string
    dato1 = Me.OleObjects("AlturaENV").Object.Value
    dato2 = Me.OleObjects("AnchoENV").Object.Value
    ResultadoENV= "EFM_" & dato1 & "X" & dato2
    MoverBloque ResultadoENV,"ENVOLVENTES", wsActual, celda
End sub