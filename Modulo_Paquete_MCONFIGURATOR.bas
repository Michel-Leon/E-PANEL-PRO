Attribute VB_Name = "Modulo_Paquete_MCONFIGURATOR"
Option Explicit

' =====================================================================
'  CREAR PAQUETE MCONFIGURATOR
'  1. Pregunta el destinatario (APM o Cliente) y elige las hojas.
'  2. Copia SOLO el área de impresión de cada hoja a un libro nuevo.
'     - Todo se copia normal (formatos, combinadas, imágenes, etc.)
'     - Las celdas con FÓRMULA se detectan solas y se pegan como VALOR.
'  3. Al final pregunta si se desea guardar y dónde.
' =====================================================================

Public Sub Crear_paquete_MCONFIGURATOR()
    Dim wbOrigen As Workbook, wbDestino As Workbook
    Dim wsOrigen As Worksheet, wsDestino As Worksheet
    Dim hojas As Variant, nombreHoja As Variant
    Dim destinatario As String
    Dim primeraHoja As Boolean
    Dim omitidas As String
    Dim hojaActual As String

    Set wbOrigen = ThisWorkbook

    ' ---------- 1. DESTINATARIO ----------
    destinatario = PedirDestinatario()
    If destinatario = "" Then
        MsgBox "Operación cancelada por el usuario.", vbInformation
        Exit Sub
    End If

    Select Case destinatario
        Case "APM"
            hojas = Array("Panel Inicio M-Confi", "FrontalPuerta", "VistaFrontal", _
                          "Lateral", "Posterior","B_Conf")
        Case "CLIENTE"
            hojas = Array("Panel Inicio M-Confi", "Vista_puerta_cliente", "Vista_FM_CLIENTE", _
                          "Lateral", "Posterior")
    End Select

    On Error GoTo ErrorGeneral
    Application.ScreenUpdating = False
    Application.Calculate                       ' asegura valores actualizados

    ' Libro nuevo con una sola hoja
    Set wbDestino = Workbooks.Add(xlWBATWorksheet)
    primeraHoja = True

    ' ---------- 2. COPIAR ÁREAS DE IMPRESIÓN ----------
    For Each nombreHoja In hojas
        hojaActual = CStr(nombreHoja)

        If Not ExisteHoja(wbOrigen, hojaActual) Then
            omitidas = omitidas & vbCrLf & "• " & hojaActual & " (no existe)"

        ElseIf wbOrigen.Worksheets(hojaActual).PageSetup.PrintArea = "" Then
            omitidas = omitidas & vbCrLf & "• " & hojaActual & " (sin área de impresión)"

        Else
            Set wsOrigen = wbOrigen.Worksheets(hojaActual)

            If primeraHoja Then
                Set wsDestino = wbDestino.Worksheets(1)
                primeraHoja = False
            Else
                Set wsDestino = wbDestino.Worksheets.Add( _
                                After:=wbDestino.Worksheets(wbDestino.Worksheets.Count))
            End If
            wsDestino.Name = wsOrigen.Name

            CopiarAreaImpresion wsOrigen, wsDestino
        End If
    Next nombreHoja

    ' Si no se copió nada, no tiene sentido dejar el libro
    If primeraHoja Then
        wbDestino.Close SaveChanges:=False
        Application.ScreenUpdating = True
        MsgBox "No se copió ninguna hoja:" & omitidas, vbExclamation
        Exit Sub
    End If

    RomperVinculos wbDestino
    wbDestino.Worksheets(1).Activate
    Application.ScreenUpdating = True

    If omitidas <> "" Then
        MsgBox "Estas hojas fueron omitidas:" & omitidas, vbExclamation
    End If

    ' ---------- 3. GUARDADO ----------
    GuardarPaquete wbDestino, destinatario, wbOrigen.Path
    Exit Sub

ErrorGeneral:
    Application.CutCopyMode = False
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    MsgBox "Error procesando la hoja '" & hojaActual & "':" & vbCrLf & _
           "Error " & Err.Number & ": " & Err.Description, vbCritical
End Sub

' =====================================================================
'  Pregunta para quién es el paquete. Devuelve "APM", "CLIENTE" o ""
' =====================================================================
Private Function PedirDestinatario() As String
    Dim r As Variant
    Do
        r = Application.InputBox( _
                "¿Para quién va dirigido el paquete?" & vbCrLf & vbCrLf & _
                "   1 = Departamento APM" & vbCrLf & _
                "   2 = Cliente", _
                "Destinatario del paquete", "1", Type:=2)

        If VarType(r) = vbBoolean Then Exit Function      ' Cancelar

        Select Case UCase(Trim(CStr(r)))
            Case "1", "APM":     PedirDestinatario = "APM": Exit Function
            Case "2", "CLIENTE": PedirDestinatario = "CLIENTE": Exit Function
        End Select

        MsgBox "Opción no válida. Escriba 1 o 2.", vbExclamation
    Loop
End Function

' =====================================================================
'  Copia el área de impresión completa y congela SOLO las fórmulas
' =====================================================================
Private Sub CopiarAreaImpresion(wsO As Worksheet, wsD As Worksheet)
    Dim rngO As Range, rngD As Range
    Dim celdasFormula As Range, area As Range, c As Range
    Dim oCelda As Object
    Dim i As Long

    Set rngO = wsO.Range(wsO.PageSetup.PrintArea).Areas(1)
    Set rngD = wsD.Range("A1").Resize(rngO.Rows.Count, rngO.Columns.Count)

    ' a) Anchos de columna y altos de fila reales del origen
    For i = 1 To rngO.Columns.Count
        rngD.Columns(i).ColumnWidth = rngO.Columns(i).ColumnWidth
    Next i
    For i = 1 To rngO.Rows.Count
        rngD.Rows(i).RowHeight = rngO.Rows(i).RowHeight
    Next i

    ' b) Copia completa (valores, formatos, combinadas, imágenes, etc.)
    rngO.Copy Destination:=rngD
    Application.CutCopyMode = False

    ' c) Detectar celdas con fórmula en el origen
    On Error Resume Next
    Set celdasFormula = rngO.SpecialCells(xlCellTypeFormulas)
    On Error GoTo 0

    ' d) Reemplazar esas celdas por su valor absoluto
    If Not celdasFormula Is Nothing Then
        For Each area In celdasFormula.Areas
            On Error Resume Next
            Equivalente(rngO, wsD, area).Value = area.Value
            If Err.Number <> 0 Then                 ' ej. fórmulas matriciales
                Err.Clear
                For Each c In area.Cells
                    If c.HasArray Then
                        Equivalente(rngO, wsD, Intersect(c.CurrentArray, rngO)).Value = _
                            Intersect(c.CurrentArray, rngO).Value
                    Else
                        Equivalente(rngO, wsD, c).Value = c.Value
                    End If
                Next c
            End If
            On Error GoTo 0
        Next area

        ' e) Fórmulas con desbordamiento (Excel 365): pegar todo el rango derramado
        On Error Resume Next
        For Each c In celdasFormula.Cells
            Set oCelda = c
            If oCelda.HasSpill Then
                Equivalente(rngO, wsD, Intersect(oCelda.SpillingToRange, rngO)).Value = _
                    Intersect(oCelda.SpillingToRange, rngO).Value
            End If
        Next c
        On Error GoTo 0
    End If

    ' f) Configuración de impresión igual a la del origen
    CopiarConfigImpresion wsO, wsD, rngD

    ' g) Ocultar líneas de cuadrícula para que se vea como el formato
    wsD.Activate
    ActiveWindow.DisplayGridlines = False
    wsD.Range("A1").Select
End Sub

' Devuelve en la hoja destino el rango equivalente a "r" del origen
Private Function Equivalente(rngBase As Range, wsD As Worksheet, r As Range) As Range
    Set Equivalente = wsD.Cells(r.Row - rngBase.Row + 1, r.Column - rngBase.Column + 1) _
                         .Resize(r.Rows.Count, r.Columns.Count)
End Function

' =====================================================================
'  Copia la configuración de página del origen
' =====================================================================
Private Sub CopiarConfigImpresion(wsO As Worksheet, wsD As Worksheet, rngD As Range)
    On Error Resume Next
    Application.PrintCommunication = False
    With wsD.PageSetup
        .PrintArea = rngD.Address
        .Orientation = wsO.PageSetup.Orientation
        .PaperSize = wsO.PageSetup.PaperSize
        .LeftMargin = wsO.PageSetup.LeftMargin
        .RightMargin = wsO.PageSetup.RightMargin
        .TopMargin = wsO.PageSetup.TopMargin
        .BottomMargin = wsO.PageSetup.BottomMargin
        .HeaderMargin = wsO.PageSetup.HeaderMargin
        .FooterMargin = wsO.PageSetup.FooterMargin
        .CenterHorizontally = wsO.PageSetup.CenterHorizontally
        .CenterVertically = wsO.PageSetup.CenterVertically
        If wsO.PageSetup.Zoom = False Then
            .Zoom = False
            .FitToPagesWide = wsO.PageSetup.FitToPagesWide
            .FitToPagesTall = wsO.PageSetup.FitToPagesTall
        Else
            .Zoom = wsO.PageSetup.Zoom
        End If
    End With
    Application.PrintCommunication = True
    On Error GoTo 0
End Sub

' =====================================================================
'  Elimina cualquier vínculo que haya quedado apuntando al libro origen
' =====================================================================
Private Sub RomperVinculos(wb As Workbook)
    Dim vinculos As Variant, i As Long
    vinculos = wb.LinkSources(xlExcelLinks)
    If Not IsEmpty(vinculos) Then
        For i = LBound(vinculos) To UBound(vinculos)
            wb.BreakLink Name:=vinculos(i), Type:=xlLinkTypeExcelLinks
        Next i
    End If
End Sub

' =====================================================================
'  Pregunta si guardar y dónde (al final del proceso)
' =====================================================================
Private Sub GuardarPaquete(wb As Workbook, destinatario As String, carpetaInicial As String)
    Dim ruta As Variant
    Dim rutaFinal As String
    Dim nombreSugerido As String

    If MsgBox("El paquete para " & destinatario & " se creó correctamente." & vbCrLf & vbCrLf & _
              "¿Desea guardarlo ahora?", vbYesNo + vbQuestion, "Guardar paquete") = vbNo Then
        MsgBox "El libro queda abierto sin guardar. Puede revisarlo y guardarlo manualmente.", vbInformation
        Exit Sub
    End If

    nombreSugerido = "Paquete_" & destinatario & "_" & Format(Now, "yyyymmdd_hhmmss")
    If carpetaInicial <> "" Then nombreSugerido = carpetaInicial & Application.PathSeparator & nombreSugerido

    Do
        ruta = Application.GetSaveAsFilename( _
                   InitialFileName:=nombreSugerido, _
                   FileFilter:="Libro de Excel (*.xlsx), *.xlsx", _
                   Title:="Guardar paquete " & destinatario & " como...")

        If VarType(ruta) = vbBoolean Then
            MsgBox "Guardado cancelado. El libro queda abierto sin guardar.", vbInformation
            Exit Sub
        End If

        rutaFinal = CStr(ruta)
        If LCase(Right(rutaFinal, 5)) <> ".xlsx" Then rutaFinal = rutaFinal & ".xlsx"

        If Dir(rutaFinal) = "" Then Exit Do
        If MsgBox("Ya existe un archivo con ese nombre:" & vbCrLf & rutaFinal & vbCrLf & vbCrLf & _
                  "¿Desea reemplazarlo?", vbYesNo + vbExclamation) = vbYes Then Exit Do
    Loop

    On Error GoTo ErrorGuardado
    Application.DisplayAlerts = False
    wb.SaveAs Filename:=rutaFinal, FileFormat:=xlOpenXMLWorkbook
    Application.DisplayAlerts = True

    MsgBox "Archivo guardado correctamente en:" & vbCrLf & rutaFinal, vbInformation, "Proceso completado"
    Exit Sub

ErrorGuardado:
    Application.DisplayAlerts = True
    MsgBox "Error al guardar el archivo (el libro sigue abierto):" & vbCrLf & _
           "Ruta: " & rutaFinal & vbCrLf & _
           "Error " & Err.Number & ": " & Err.Description, vbCritical
End Sub

' =====================================================================
'  Utilidad
' =====================================================================
Private Function ExisteHoja(wb As Workbook, nombre As String) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Worksheets(nombre)
    On Error GoTo 0
    ExisteHoja = Not ws Is Nothing
End Function

' =====================================================================
'  MACRO DE APOYO: muestra los rangos exactos con fórmula
'  dentro del área de impresión de cada hoja de los paquetes
' =====================================================================
Public Sub Mapear_Formulas()
    Dim hojas As Variant, h As Variant
    Dim ws As Worksheet, wbRep As Workbook, wsRep As Worksheet
    Dim rngArea As Range, celdas As Range, a As Range
    Dim fila As Long

    hojas = Array("Panel Inicio M-Confi", "FrontalPuerta", "VistaFrontal", "Vista_P_CLIENTE", _
                  "Vista_FM_CLIENTE", "Vista_Sup_Inf", "Vista_Lat_Post")

    Set wbRep = Workbooks.Add(xlWBATWorksheet)
    Set wsRep = wbRep.Worksheets(1)
    wsRep.Name = "Mapa_Formulas"
    wsRep.Range("A1:D1").Value = Array("Hoja", "Área de impresión", "Rango con fórmula", "Nº celdas")
    wsRep.Range("A1:D1").Font.Bold = True
    fila = 2

    For Each h In hojas
        If ExisteHoja(ThisWorkbook, CStr(h)) Then
            Set ws = ThisWorkbook.Worksheets(CStr(h))
            If ws.PageSetup.PrintArea <> "" Then
                Set rngArea = ws.Range(ws.PageSetup.PrintArea).Areas(1)
                Set celdas = Nothing
                On Error Resume Next
                Set celdas = rngArea.SpecialCells(xlCellTypeFormulas)
                On Error GoTo 0

                If celdas Is Nothing Then
                    wsRep.Cells(fila, 1).Resize(1, 4).Value = _
                        Array(ws.Name, rngArea.Address(0, 0), "(sin fórmulas)", 0)
                    fila = fila + 1
                Else
                    For Each a In celdas.Areas
                        wsRep.Cells(fila, 1).Resize(1, 4).Value = _
                            Array(ws.Name, rngArea.Address(0, 0), a.Address(0, 0), a.Cells.Count)
                        fila = fila + 1
                    Next a
                End If
            Else
                wsRep.Cells(fila, 1).Resize(1, 3).Value = Array(ws.Name, "(sin área)", "-")
                fila = fila + 1
            End If
        End If
    Next h

    wsRep.Columns("A:D").AutoFit
    MsgBox "Mapa de fórmulas generado.", vbInformation
End Sub
