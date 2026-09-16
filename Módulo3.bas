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