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
        ActiveSheet.Shapes("Menu UnidadFuncional").Visible = False
        ActiveSheet.Shapes("Menu Lisas").Visible = False
        ActiveSheet.Shapes("Menu Contactores").Visible = False
        ActiveSheet.Shapes("Menu DPS").Visible = False
        ActiveSheet.Shapes("Menu Equipo CTR").Visible = False
        ActiveSheet.Shapes("Menu ACCESORIOS CTR").Visible = False
        ActiveSheet.Shapes("Menu Equipos IND").Visible = False
    end If
End Sub  
Public sub Seleccion_Menu_UnidadesFuncionales()
    Dim shp As Shape
    Dim groupname As String
    groupname = "Menu UnidadFuncional"
    
    On Error Resume Next
    Set shp = ActiveSheet.Shapes(groupname)
    On Error GoTo 0
    
    If Not shp Is Nothing Then
        ' Alterna la visibilidad del grupo seleccionado
        shp.Visible = Not shp.Visible
        ' Oculta los dem s grupos
        ActiveSheet.Shapes("Menu Envolventes").Visible = False
        ActiveSheet.Shapes("Menu Lisas").Visible = False
        ActiveSheet.Shapes("Menu Contactores").Visible = False
        ActiveSheet.Shapes("Menu DPS").Visible = False
        ActiveSheet.Shapes("Menu Equipo CTR").Visible = False
        ActiveSheet.Shapes("Menu ACCESORIOS CTR").Visible = False
        ActiveSheet.Shapes("Menu Equipos IND").Visible = False
    end If
End Sub  
Public sub Seleccion_Menu_Lisas()
    Dim shp As Shape
    Dim groupname As String
    groupname = "Menu Lisas"
    
    On Error Resume Next
    Set shp = ActiveSheet.Shapes(groupname)
    On Error GoTo 0
    
    If Not shp Is Nothing Then
        ' Alterna la visibilidad del grupo seleccionado
        shp.Visible = Not shp.Visible
        ' Oculta los dem s grupos
        ActiveSheet.Shapes("Menu Envolventes").Visible = False
        ActiveSheet.Shapes("Menu UnidadFuncional").Visible = False
        ActiveSheet.Shapes("Menu Contactores").Visible = False
        ActiveSheet.Shapes("Menu DPS").Visible = False
        ActiveSheet.Shapes("Menu Equipo CTR").Visible = False
        ActiveSheet.Shapes("Menu ACCESORIOS CTR").Visible = False
        ActiveSheet.Shapes("Menu Equipos").Visible = False
    end If
End Sub
Public sub Seleccion_Menu_Contactores()
    Dim shp As Shape
    Dim groupname As String
    groupname = "Menu Contactores"
    
    On Error Resume Next
    Set shp = ActiveSheet.Shapes(groupname)
    On Error GoTo 0
    
    If Not shp Is Nothing Then
        ' Alterna la visibilidad del grupo seleccionado
        shp.Visible = Not shp.Visible
        ' Oculta los dem s grupos
        ActiveSheet.Shapes("Menu Envolventes").Visible = False
        ActiveSheet.Shapes("Menu UnidadFuncional").Visible = False
        ActiveSheet.Shapes("Menu Lisas").Visible = False
        ActiveSheet.Shapes("Menu DPS").Visible = False
        ActiveSheet.Shapes("Menu Equipo CTR").Visible = False
        ActiveSheet.Shapes("Menu ACCESORIOS CTR").Visible = False
        ActiveSheet.Shapes("Menu Equipos").Visible = False
    end If
End Sub
Public sub seleccion_menu_dps()
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
        ActiveSheet.Shapes("Menu UnidadFuncional").Visible = False
        ActiveSheet.Shapes("Menu Lisas").Visible = False
        ActiveSheet.Shapes("Menu Contactores").Visible = False
        ActiveSheet.Shapes("Menu Equipo CTR").Visible = False
        ActiveSheet.Shapes("Menu ACCESORIOS CTR").Visible = False
        ActiveSheet.Shapes("Menu Equipos").Visible = False
    end If
End Sub
Public sub Seleccion_Menu_EquipoCTR()
    Dim shp As Shape
    Dim groupname As String
    groupname = "Menu Equipo CTR"
    
    On Error Resume Next
    Set shp = ActiveSheet.Shapes(groupname)
    On Error GoTo 0
    
    If Not shp Is Nothing Then
        ' Alterna la visibilidad del grupo seleccionado
        shp.Visible = Not shp.Visible
        ' Oculta los dem s grupos
        ActiveSheet.Shapes("Menu Envolventes").Visible = False
        ActiveSheet.Shapes("Menu UnidadFuncional").Visible = False
        ActiveSheet.Shapes("Menu Lisas").Visible = False
        ActiveSheet.Shapes("Menu Contactores").Visible = False
        ActiveSheet.Shapes("Menu DPS").Visible = False
        ActiveSheet.Shapes("Menu ACCESORIOS CTR").Visible = False
        ActiveSheet.Shapes("Menu Equipos").Visible = False
    end If
End Sub
public sub Seleccion_Menu_AccesoriosCTR()
    Dim shp As Shape
    Dim groupname As String
    groupname = "Menu ACCESORIOS CTR"
    
    On Error Resume Next
    Set shp = ActiveSheet.Shapes(groupname)
    On Error GoTo 0
    
    If Not shp Is Nothing Then
        ' Alterna la visibilidad del grupo seleccionado
        shp.Visible = Not shp.Visible
        ' Oculta los dem s grupos
        ActiveSheet.Shapes("Menu Envolventes").Visible = False
        ActiveSheet.Shapes("Menu UnidadFuncional").Visible = False
        ActiveSheet.Shapes("Menu Lisas").Visible = False
        ActiveSheet.Shapes("Menu Contactores").Visible = False
        ActiveSheet.Shapes("Menu DPS").Visible = False
        ActiveSheet.Shapes("Menu Equipo CTR").Visible = False
        ActiveSheet.Shapes("Menu Equipos").Visible = False
    end If
End Sub
Public sub Seleccion_Menu_EquiposIND()
    Dim shp As Shape
    Dim groupname As String
    groupname = "Menu Equipos"
    
    On Error Resume Next    
    Set shp = ActiveSheet.Shapes(groupname)
    On Error GoTo 0
    
    If Not shp Is Nothing Then
        ' Alterna la visibilidad del grupo seleccionado
        shp.Visible = Not shp.Visible
        ' Oculta los dem s grupos
        ActiveSheet.Shapes("Menu Envolventes").Visible = False
        ActiveSheet.Shapes("Menu UnidadFuncional").Visible = False
        ActiveSheet.Shapes("Menu Lisas").Visible = False
        ActiveSheet.Shapes("Menu Contactores").Visible = False
        ActiveSheet.Shapes("Menu DPS").Visible = False
        ActiveSheet.Shapes("Menu Equipo CTR").Visible = False
        ActiveSheet.Shapes("Menu ACCESORIOS CTR").Visible = False
    end If
End Sub
