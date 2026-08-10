Attribute VB_Name = "M�dulo2"
Public Sub Seleccion_Menu_Envolventes()
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
        ActiveSheet.Shapes("Menu DPS").Visible = False
        ActiveSheet.Shapes("Menu_Servicios").Visible = False
        ActiveSheet.Shapes("Menu_BARRAS").Visible = False
        ActiveSheet.Shapes("Menu_Lisas").ZOrder msoBringToFront
        
    End If
End Sub
Public Sub Seleccion_Menu_DPS()
    Dim shp As Shape
    Dim groupname As String
    groupname = "Menu DPS"
    
    On Error Resume Next
    Set shp = ActiveSheet.Shapes(groupname)
    On Error GoTo 0
    
    If Not shp Is Nothing Then
        ' Alterna la visibilidad del grupo seleccionado
        shp.Visible = Not shp.Visible
        ' Oculta los dem s grupos
        ActiveSheet.Shapes("Menu Envolventes").Visible = False
        ActiveSheet.Shapes("Menu_Servicios").Visible = False
        ActiveSheet.Shapes("Menu_BARRAS").Visible = False
        ActiveSheet.Shapes("Menu_Lisas").ZOrder msoBringToFront
    End If
End Sub
public Sub Seleccion_Menu_Servicios()
    Dim shp As Shape
    Dim groupname As String
    groupname = "Menu_Servicios"
    
    On Error Resume Next
    Set shp = ActiveSheet.Shapes(groupname)
    On Error GoTo 0
    
    If Not shp Is Nothing Then
        ' Alterna la visibilidad del grupo seleccionado
        shp.Visible = Not shp.Visible
        ' Oculta los dem s grupos
        ActiveSheet.Shapes("Menu Envolventes").Visible = False
        ActiveSheet.Shapes("Menu DPS").Visible = False
        ActiveSheet.Shapes("Menu_BARRAS").Visible = False
        ActiveSheet.Shapes("Menu_Lisas").ZOrder msoBringToFront
    End If
End Sub
Public Sub Seleccion_Menu_BARRAS()
    Dim shp As Shape
    Dim groupname As String
    groupname = "Menu_BARRAS"
    
    On Error Resume Next
    Set shp = ActiveSheet.Shapes(groupname)
    On Error GoTo 0
    
    If Not shp Is Nothing Then
        ' Alterna la visibilidad del grupo seleccionado
        shp.Visible = Not shp.Visible
        ' Oculta los dem s grupos
        ActiveSheet.Shapes("Menu Envolventes").Visible = False
        ActiveSheet.Shapes("Menu DPS").Visible = False
        ActiveSheet.Shapes("Menu_Servicios").Visible = False
        ActiveSheet.Shapes("Menu_Lisas").ZOrder msoBringToFront
    End If
End Sub
Public Sub Seleccion_Menu_Lisas()
    Dim shp As Shape
    Dim groupname As String
    groupname = "Menu_Lisas"
    
    On Error Resume Next
    Set shp = ActiveSheet.Shapes(groupname)
    On Error GoTo 0
    
    If Not shp Is Nothing Then
        ' Alterna la visibilidad del grupo seleccionado
        shp.Visible = Not shp.Visible
        ' Oculta los dem s grupos
        ActiveSheet.Shapes("Menu Envolventes").Visible = False
        ActiveSheet.Shapes("Menu DPS").Visible = False
        ActiveSheet.Shapes("Menu_Servicios").Visible = False
        ActiveSheet.Shapes("Menu_BARRAS").ZOrder msoBringToFront
    End If
End Sub