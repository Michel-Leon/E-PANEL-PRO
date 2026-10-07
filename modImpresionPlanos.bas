Attribute VB_Name = "modImpresionPlanos"

Option Explicit

'==============================================================
'  PAQUETE DE IMPRESION / PDF DE PLANOS
'  Vistas: AREA_n (zona 1) y AREA_n_1 (zona 2), n = 1..6
'==============================================================

Public Const PREFIJO As String = "AREA_"
Public Const NUM_VISTAS As Long = 6
Public Const NOMBRE_BOTON As String = "btnPaqueteImpresion"

'--- Nombres de las vistas (en el mismo orden que AREA_1 ... AREA_6)
Public Function NombresVistas() As Variant
    NombresVistas = Array("Plano Puerta", "Frente Muerto", "Vista Superior", _
                          "Vista Inferior", "Vista Lateral", "Vista Posterior")
End Function

'--- Opciones del combo de cada vista (el Indice importa)
'    0 = No incluir, 1 = Solo zona 1, 2 = Solo zona 2, 3 = Ambas, 4 = Automatico
Public Function OpcionesZona() As Variant
    OpcionesZona = Array("No incluir", "Solo zona 1", "Solo zona 2", _
                         "Ambas zonas", "Automatico (con contenido)")
End Function

'--- Macro que se asigna al boton
Public Sub MostrarPanelImpresion()
    frmImprimirVistas.Show
End Sub

'--- Crea el boton en la hoja activa, en la celda seleccionada
Public Sub CrearBotonImpresion()
    Dim ws As Worksheet, shp As Shape
    Set ws = ActiveSheet

    On Error Resume Next
    ws.Shapes(NOMBRE_BOTON).Delete
    On Error GoTo 0

    Set shp = ws.Shapes.AddShape(msoShapeRoundedRectangle, _
                                 ActiveCell.Left, ActiveCell.Top, 180, 34)
    With shp
        .Name = NOMBRE_BOTON
        .OnAction = "MostrarPanelImpresion"
        .Fill.ForeColor.RGB = RGB(0, 101, 187)
        .Line.Visible = msoFalse
        .Placement = xlFreeFloating
        With .TextFrame2
            .VerticalAnchor = msoAnchorMiddle
            With .TextRange
                .Text = "Imprimir / PDF planos"
                .Font.Bold = msoTrue
                .Font.Size = 11
                .Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
                .ParagraphFormat.Alignment = msoAlignCenter
            End With
        End With
    End With

    'Que el boton no salga impreso
    On Error Resume Next
    shp.DrawingObject.PrintObject = False
    On Error GoTo 0
End Sub

'--- Nombre definido de cada zona
Public Function NombreZona(ByVal vista As Long, ByVal zona As Long) As String
    If zona = 1 Then
        NombreZona = PREFIJO & vista
    Else
        NombreZona = PREFIJO & vista & "_1"
    End If
End Function

'--- Devuelve el rango de un nombre (de libro o de hoja). Nothing si no existe
Public Function ObtenerRango(ByVal nombre As String) As Range
    Dim ws As Worksheet, r As Range
    On Error Resume Next
    Set r = ThisWorkbook.Names(nombre).RefersToRange
    If r Is Nothing Then
        For Each ws In ThisWorkbook.Worksheets
            Set r = ws.Names(nombre).RefersToRange
            If Not r Is Nothing Then Exit For
        Next ws
    End If
    On Error GoTo 0
    Set ObtenerRango = r
End Function

'--- La zona tiene algo? (valores, formulas con texto visible, imagenes o dibujos)
Public Function TieneContenido(ByVal rng As Range) As Boolean
    Dim a As Range, shp As Shape, toca As Boolean

    If rng Is Nothing Then Exit Function

    For Each a In rng.areas
        If a.CountLarge - Application.WorksheetFunction.CountBlank(a) > 0 Then
            TieneContenido = True
            Exit Function
        End If
    Next a

    For Each shp In rng.Worksheet.Shapes
        Select Case shp.Type
            Case msoFormControl, msoOLEControlObject, msoComment
                'botones y comentarios no cuentan
            Case Else
                If shp.Visible And shp.Name <> NOMBRE_BOTON Then
                    toca = False
                    On Error Resume Next
                    toca = Not Intersect(shp.TopLeftCell, rng) Is Nothing
                    If Not toca Then toca = Not Intersect(shp.BottomRightCell, rng) Is Nothing
                    On Error GoTo 0
                    If toca Then
                        TieneContenido = True
                        Exit Function
                    End If
                End If
        End Select
    Next shp
End Function

'--- Texto de estado para el formulario
Public Function EstadoZona(ByVal vista As Long, ByVal zona As Long) As String
    Dim r As Range
    Set r = ObtenerRango(NombreZona(vista, zona))
    If r Is Nothing Then
        EstadoZona = "no existe"
    ElseIf TieneContenido(r) Then
        EstadoZona = "con contenido"
    Else
        EstadoZona = "vacia"
    End If
End Function

'--- Arma la lista de rangos a imprimir, en orden de vista y zona
Public Function AreasSeleccionadas(modos() As Long, ByVal omitirVacias As Boolean) As Collection
    Dim col As New Collection, v As Long, z As Long
    Dim incluir As Boolean, r As Range

    For v = 1 To NUM_VISTAS
        For z = 1 To 2
            Select Case modos(v)
                Case 1: incluir = (z = 1)
                Case 2: incluir = (z = 2)
                Case 3, 4: incluir = True
                Case Else: incluir = False
            End Select

            If incluir Then
                Set r = ObtenerRango(NombreZona(v, z))
                If Not r Is Nothing Then
                    If modos(v) = 4 Or omitirVacias Then
                        If TieneContenido(r) Then col.Add r
                    Else
                        col.Add r
                    End If
                End If
            End If
        Next z
    Next v

    Set AreasSeleccionadas = col
End Function

'--- Nombre sugerido para el PDF
Private Function NombrePDFSugerido() As String
    Dim carpeta As String, base As String
    carpeta = ThisWorkbook.Path
    If carpeta = "" Then carpeta = CurDir
    base = ThisWorkbook.Name
    If InStrRev(base, ".") > 0 Then base = Left$(base, InStrRev(base, ".") - 1)
    NombrePDFSugerido = carpeta & Application.PathSeparator & base & "_Planos_" & _
                        Format(Now, "yyyy-mm-dd_hhmm") & ".pdf"
End Function

'==============================================================
'  FORMATO DE PaGINA DEL PAQUETE
'  Horizontal, margenes 0, centrado horizontal y vertical,
'  cada zona ajustada a 1 pagina
'==============================================================
Private Function GuardarFormato(ByVal ws As Worksheet) As Variant
    With ws.PageSetup
        GuardarFormato = Array(.PrintArea, .Orientation, .LeftMargin, .RightMargin, _
                               .TopMargin, .BottomMargin, .HeaderMargin, .FooterMargin, _
                               .CenterHorizontally, .CenterVertically, _
                               .Zoom, .FitToPagesWide, .FitToPagesTall)
    End With
End Function

Private Sub AplicarFormato(ByVal ws As Worksheet, ByVal areaImpresion As String)
    With ws.PageSetup
        .PrintArea = areaImpresion
        .Orientation = xlLandscape
        .LeftMargin = 0
        .RightMargin = 0
        .TopMargin = 0
        .BottomMargin = 0
        .HeaderMargin = 0
        .FooterMargin = 0
        .CenterHorizontally = True
        .CenterVertically = True
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = 1
    End With
End Sub

Private Sub RestaurarFormato(ByVal ws As Worksheet, ByVal f As Variant)
    With ws.PageSetup
        .PrintArea = f(0)
        .Orientation = f(1)
        .LeftMargin = f(2)
        .RightMargin = f(3)
        .TopMargin = f(4)
        .BottomMargin = f(5)
        .HeaderMargin = f(6)
        .FooterMargin = f(7)
        .CenterHorizontally = f(8)
        .CenterVertically = f(9)
        If f(10) = False Then
            .Zoom = False
            .FitToPagesWide = f(11)
            .FitToPagesTall = f(12)
        Else
            .Zoom = f(10)
        End If
    End With
End Sub

'--- Genera UN solo paquete (PDF o impresion) con todas las zonas elegidas
Public Sub GenerarPaquete(ByVal areas As Collection, ByVal aPDF As Boolean, _
                          ByVal abrirPDF As Boolean, ByVal copias As Long)
    Dim direcciones As Object, originales As Object
    Dim r As Range, k As Variant, i As Long
    Dim hojaInicial As Object, ruta As Variant
    Dim errNum As Long, errDesc As String

    Set direcciones = CreateObject("Scripting.Dictionary")
    Set originales = CreateObject("Scripting.Dictionary")

    'Agrupar las zonas por hoja, conservando el orden de las vistas
    For Each r In areas
        k = r.Worksheet.Name
        If direcciones.Exists(k) Then
            direcciones(k) = direcciones(k) & "," & r.Address
        Else
            direcciones.Add k, r.Address
        End If
    Next r

    'Destino
    If aPDF Then
        ruta = Application.GetSaveAsFilename( _
                    InitialFileName:=NombrePDFSugerido(), _
                    FileFilter:="Archivo PDF (*.pdf), *.pdf", _
                    Title:="Guardar paquete de planos en PDF")
        If VarType(ruta) = vbBoolean Then Exit Sub   'cancelo
    Else
        If Not Application.Dialogs(xlDialogPrinterSetup).Show Then Exit Sub
    End If

    ThisWorkbook.Activate
    Set hojaInicial = ActiveSheet
    Application.ScreenUpdating = False
    On Error GoTo Restaurar

    'Formato temporal: zonas elegidas, horizontal, margenes 0, centrado, 1 pagina por zona
    Application.PrintCommunication = False
    For Each k In direcciones.Keys
        originales(k) = GuardarFormato(ThisWorkbook.Worksheets(k))
        AplicarFormato ThisWorkbook.Worksheets(k), direcciones(k)
    Next k
    Application.PrintCommunication = True

    'Seleccionar juntas todas las hojas involucradas -> un solo paquete
    i = 0
    For Each k In direcciones.Keys
        ThisWorkbook.Worksheets(k).Select Replace:=(i = 0)
        i = i + 1
    Next k

    If aPDF Then
        ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, Filename:=CStr(ruta), _
            Quality:=xlQualityStandard, IncludeDocProperties:=True, _
            IgnorePrintAreas:=False, OpenAfterPublish:=abrirPDF
    Else
        ActiveWindow.SelectedSheets.PrintOut Copies:=copias, Collate:=True
    End If

Restaurar:
    errNum = Err.Number
    errDesc = Err.Description
    On Error Resume Next
    Application.PrintCommunication = False
    For Each k In originales.Keys
        RestaurarFormato ThisWorkbook.Worksheets(k), originales(k)
    Next k
    Application.PrintCommunication = True
    hojaInicial.Select
    Application.ScreenUpdating = True
    On Error GoTo 0

    If errNum <> 0 Then
        MsgBox "No se pudo completar la operacion:" & vbCrLf & errDesc, vbExclamation
    ElseIf aPDF Then
        If Not abrirPDF Then MsgBox "PDF generado con " & areas.Count & " zona(s):" & _
                                    vbCrLf & ruta, vbInformation
    Else
        MsgBox "Enviadas " & areas.Count & " zona(s) a la impresora.", vbInformation
    End If
End Sub


