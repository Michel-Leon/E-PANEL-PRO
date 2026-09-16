Attribute VB_Name = "M�dulo3"
Public Sub Puerta_Menu_Envolventes()
    Dim shp As Shape
    Dim groupname As String
    groupname = "Menu Envolventes"
    
    On Error Resume Next
    Set shp = ActiveSheet.Shapes(groupname)
    On Error GoTo 0
    
    If Not shp Is Nothing Then
        ' Alterna la visibilidad del grupo seleccionado
        shp.Visible = Not shp.Visible
        ' Oculta los dem s grupos
        ActiveSheet.Shapes("Menu Placas de control").Visible = False
        ActiveSheet.Shapes("Menu AccesoriosMTM").Visible = False
        ActiveSheet.Shapes("Menu Etiqueta").Visible = False
        ActiveSheet.Shapes("Menu Acotar").Visible = False
        ActiveSheet.Shapes("Menu Color").Visible = False
    End If
End Sub
Public sub Puerta_Menu_Placas()
    Dim shp As Shape
    Dim groupname As String
    groupname = "Menu Placas de control"
    
    On Error Resume Next
    Set shp = ActiveSheet.Shapes(groupname)
    On Error GoTo 0
    
    If Not shp Is Nothing Then
        ' Alterna la visibilidad del grupo seleccionado
        shp.Visible = Not shp.Visible
        ' Oculta los dem s grupos
        ActiveSheet.Shapes("Menu Envolventes").Visible = False
        ActiveSheet.Shapes("Menu AccesoriosMTM").Visible = False
        ActiveSheet.Shapes("Menu Etiqueta").Visible = False
        ActiveSheet.Shapes("Menu Acotar").Visible = False
        ActiveSheet.Shapes("Menu Color").Visible = False
    End If
End sub
Public sub Puerta_Menu_AccesoriosMTM()
    Dim shp As Shape
    Dim groupname As String
    groupname = "Menu AccesoriosMTM"
    
    On Error Resume Next
    Set shp = ActiveSheet.Shapes(groupname)
    On Error GoTo 0
    
    If Not shp Is Nothing Then
        ' Alterna la visibilidad del grupo seleccionado
        shp.Visible = Not shp.Visible
        ' Oculta los dem s grupos
        ActiveSheet.Shapes("Menu Envolventes").Visible = False
        ActiveSheet.Shapes("Menu Placas de control").Visible = False
        ActiveSheet.Shapes("Menu Etiqueta").Visible = False
        ActiveSheet.Shapes("Menu Acotar").Visible = False
        ActiveSheet.Shapes("Menu Color").Visible = False
    End If
End sub
Public sub Puerta_Menu_Etiqueta()
    Dim shp As Shape
    Dim groupname As String
    groupname = "Menu Etiqueta"
    
    On Error Resume Next
    Set shp = ActiveSheet.Shapes(groupname)
    On Error GoTo 0
    
    If Not shp Is Nothing Then
        ' Alterna la visibilidad del grupo seleccionado
        shp.Visible = Not shp.Visible
        ' Oculta los dem s grupos
        ActiveSheet.Shapes("Menu Envolventes").Visible = False
        ActiveSheet.Shapes("Menu Placas de control").Visible = False
        ActiveSheet.Shapes("Menu AccesoriosMTM").Visible = False
        ActiveSheet.Shapes("Menu Acotar").Visible = False
        ActiveSheet.Shapes("Menu Color").Visible = False
    End If
End sub
public sub Puerta_Menu_Acotar()
    Dim shp As Shape
    Dim groupname As String
    groupname = "Menu Acotar"
    
    On Error Resume Next
    Set shp = ActiveSheet.Shapes(groupname)
    On Error GoTo 0
    
    If Not shp Is Nothing Then
        ' Alterna la visibilidad del grupo seleccionado
        shp.Visible = Not shp.Visible
        ' Oculta los dem s grupos
        ActiveSheet.Shapes("Menu Envolventes").Visible = False
        ActiveSheet.Shapes("Menu Placas de control").Visible = False
        ActiveSheet.Shapes("Menu AccesoriosMTM").Visible = False
        ActiveSheet.Shapes("Menu Etiqueta").Visible = False
        ActiveSheet.Shapes("Menu Color").Visible = False
    End If
End sub
public sub Puerta_Menu_Color()
    Dim shp As Shape
    Dim groupname As String
    groupname = "Menu Color"
    
    On Error Resume Next
    Set shp = ActiveSheet.Shapes(groupname)
    On Error GoTo 0
    
    If Not shp Is Nothing Then
        ' Alterna la visibilidad del grupo seleccionado
        shp.Visible = Not shp.Visible
        ' Oculta los dem s grupos
        ActiveSheet.Shapes("Menu Envolventes").Visible = False
        ActiveSheet.Shapes("Menu Placas de control").Visible = False
        ActiveSheet.Shapes("Menu AccesoriosMTM").Visible = False
        ActiveSheet.Shapes("Menu Etiqueta").Visible = False
        ActiveSheet.Shapes("Menu Acotar").Visible = False
    End If
End sub

Public Sub Puerta_area_trabajo2()
    ThisWorkbook.Sheets("FrontalPuerta").Range("KO83:NJ185").Clear
End Sub
Public Sub Puerta_area_trabajo1()
    ThisWorkbook.Sheets("FrontalPuerta").Range("CG48:FI146").Clear
End Sub
'=======================================================================
Public Sub Superior_Inferior_Menu_Envolventes()
    Dim shp As Shape
    Dim groupname As String
    groupname = "Menu Envolventes"
    
    On Error Resume Next
    Set shp = ActiveSheet.Shapes(groupname)
    On Error GoTo 0
    
    If Not shp Is Nothing Then
        ' Alterna la visibilidad del grupo seleccionado
        shp.Visible = Not shp.Visible
        ' Oculta los dem s grupos
        ActiveSheet.Shapes("Menu AccesoriosMTM").Visible = False
        ActiveSheet.Shapes("Menu Acotar").Visible = False
        ActiveSheet.Shapes("Menu Color").Visible = False
    End If
End Sub
Public Sub Superior_Inferior_Menu_AccesoriosMTM()
    Dim shp As Shape
    Dim groupname As String
    groupname = "Menu AccesoriosMTM"
    
    On Error Resume Next
    Set shp = ActiveSheet.Shapes(groupname)
    On Error GoTo 0
    
    If Not shp Is Nothing Then
        ' Alterna la visibilidad del grupo seleccionado
        shp.Visible = Not shp.Visible
        ' Oculta los dem s grupos
        ActiveSheet.Shapes("Menu Envolventes").Visible = False
        ActiveSheet.Shapes("Menu Acotar").Visible = False
        ActiveSheet.Shapes("Menu Color").Visible = False
    End If
End Sub
Public Sub Superior_Inferior_Menu_Acotar()
    Dim shp As Shape
    Dim groupname As String
    groupname = "Menu Acotar"
    
    On Error Resume Next
    Set shp = ActiveSheet.Shapes(groupname)
    On Error GoTo 0
    
    If Not shp Is Nothing Then
        ' Alterna la visibilidad del grupo seleccionado
        shp.Visible = Not shp.Visible
        ' Oculta los dem s grupos
        ActiveSheet.Shapes("Menu Envolventes").Visible = False
        ActiveSheet.Shapes("Menu AccesoriosMTM").Visible = False
        ActiveSheet.Shapes("Menu Color").Visible = False
    End If
End Sub
Public Sub Superior_Inferior_Menu_Color()
    Dim shp As Shape
    Dim groupname As String
    groupname = "Menu Color"
    
    On Error Resume Next
    Set shp = ActiveSheet.Shapes(groupname)
    On Error GoTo 0
    
    If Not shp Is Nothing Then
        ' Alterna la visibilidad del grupo seleccionado
        shp.Visible = Not shp.Visible
        ' Oculta los dem s grupos
        ActiveSheet.Shapes("Menu Envolventes").Visible = False
        ActiveSheet.Shapes("Menu AccesoriosMTM").Visible = False
        ActiveSheet.Shapes("Menu Acotar").Visible = False
    End If
End Sub
Public Sub Superior_Inferior_area_trabajo2()
    ThisWorkbook.Sheets("Vista_Sup_Inf").Range("KO83:NJ185").Clear
End Sub
Public Sub Superior_Inferior_area_trabajo1()
    Dim ws As Worksheet
    Dim resp As String
    Dim opcion As Long

    Set ws = ThisWorkbook.Sheets("Vista_Sup_Inf")
    Do
        resp = InputBox( _
            "¿Qué vista desea limpiar?" & vbCrLf & vbCrLf & _
            "   1 = Vista SUPERIOR" & vbCrLf & _
            "   2 = Vista INFERIOR" & vbCrLf & vbCrLf & _
            "Escriba 1 o 2 y presione Aceptar." & vbCrLf & _
            "Presione Cancelar para abortar la operación.", _
            "Limpiar área de trabajo")
        ' StrPtr = 0 significa que se presionó CANCELAR
        If StrPtr(resp) = 0 Then
            MsgBox "Operación cancelada. No se borró ningún dato.", vbInformation, "Cancelado"
            Exit Sub
        End If
        resp = Trim(resp)
        If resp = "1" Or resp = "2" Then
            opcion = CLng(resp)
            Exit Do
        Else
            MsgBox "Opción no válida." & vbCrLf & _
                   "Debe escribir 1 (Superior) o 2 (Inferior).", vbExclamation, "Dato incorrecto"
        End If
    Loop
    Application.ScreenUpdating = False
    Select Case opcion
        Case 1  ' ----- VISTA SUPERIOR -----
            ws.Range("CG45:FR122").Clear   ' área 1 vista superior
            ws.Range("GE45:JJ122").Clear   ' área 2 vista superior
            MsgBox "Vista SUPERIOR limpiada correctamente.", vbInformation, "Listo"
        Case 2  ' ----- VISTA INFERIOR -----
            ws.Range("CG130:FR207").Clear  ' área 1 vista inferior  <-- AJUSTAR RANGO
            ws.Range("GE130:JJ207").Clear  ' área 2 vista inferior  <-- AJUSTAR RANGO
            MsgBox "Vista INFERIOR limpiada correctamente.", vbInformation, "Listo"
    End Select
    Application.ScreenUpdating = True

End Sub