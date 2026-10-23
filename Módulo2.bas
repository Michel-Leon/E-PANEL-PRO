Attribute VB_Name = "Módulo2"
Public Sub Seleccion_Menu_EnvolventesGFL()
    Dim shp As Shape
    Dim groupname As String
    groupname = "Menu EnvolventesGFL"
    
    On Error Resume Next
    Set shp = ActiveSheet.Shapes(groupname)
    On Error GoTo 0
    
    If Not shp Is Nothing Then
        ' Alterna la visibilidad del grupo seleccionado
        shp.Visible = Not shp.Visible
        ' Oculta los dem s grupos
        ActiveSheet.Shapes("Menu UnidadFuncionalGFL").Visible = False
        ActiveSheet.Shapes("Menu LisasGFL").Visible = False
        ActiveSheet.Shapes("Menu DPSGFL").Visible = False
        ActiveSheet.Shapes("Menu Equipo CTRGFL").Visible = False
        ActiveSheet.Shapes("Menu ACCESORIOS CTRGFL").Visible = False
        ActiveSheet.Shapes("Menu ETOGFL").Visible = False
        ActiveSheet.Shapes("Menu CONMUTADORESGFL").Visible = False
        ActiveSheet.Shapes("Menu AccesoriosMTMGFL").Visible = False
        ActiveSheet.Shapes("Menu BarrajeGFL").Visible = False
        ActiveSheet.Shapes("Menu TransferenciasGFL").Visible = False
        ActiveSheet.Shapes("Menu ServiciosGFL").Visible = False
        ActiveSheet.Shapes("Menu ColorGFL").Visible = False
        ActiveSheet.Shapes("Menu AcotarGFL").Visible = False
        ActiveSheet.Shapes("Menu EtiquetaGFL").Visible = False
        ActiveSheet.Shapes("Menu ArranqueM GF").Visible = False
    End If
End Sub
Public Sub Seleccion_Menu_UnidadFuncionalGFL()
    Dim shp As Shape
    Dim groupname As String
    groupname = "Menu UnidadFuncionalGFL"
    
    On Error Resume Next
    Set shp = ActiveSheet.Shapes(groupname)
    On Error GoTo 0
    
    If Not shp Is Nothing Then
        ' Alterna la visibilidad del grupo seleccionado
        shp.Visible = Not shp.Visible
        ' Oculta los dem s grupos
        ActiveSheet.Shapes("Menu EnvolventesGFL").Visible = False
        ActiveSheet.Shapes("Menu LisasGFL").Visible = False
        ActiveSheet.Shapes("Menu DPSGFL").Visible = False
        ActiveSheet.Shapes("Menu Equipo CTRGFL").Visible = False
        ActiveSheet.Shapes("Menu ACCESORIOS CTRGFL").Visible = False
        ActiveSheet.Shapes("Menu ETOGFL").Visible = False
        ActiveSheet.Shapes("Menu CONMUTADORESGFL").Visible = False
        ActiveSheet.Shapes("Menu AccesoriosMTMGFL").Visible = False
        ActiveSheet.Shapes("Menu BarrajeGFL").Visible = False
        ActiveSheet.Shapes("Menu TransferenciasGFL").Visible = False
        ActiveSheet.Shapes("Menu ServiciosGFL").Visible = False
        ActiveSheet.Shapes("Menu ColorGFL").Visible = False
        ActiveSheet.Shapes("Menu AcotarGFL").Visible = False
        ActiveSheet.Shapes("Menu EtiquetaGFL").Visible = False
        ActiveSheet.Shapes("Menu ArranqueM GF").Visible = False
    End If
End Sub
Public Sub seleccion_menu_lisasGFL()
    Dim shp As Shape
    Dim groupname As String
    groupname = "Menu LisasGFL"
    
    On Error Resume Next
    Set shp = ActiveSheet.Shapes(groupname)
    On Error GoTo 0
    
    If Not shp Is Nothing Then
        ' Alterna la visibilidad del grupo seleccionado
        shp.Visible = Not shp.Visible
        ' Oculta los dem s grupos
        ActiveSheet.Shapes("Menu EnvolventesGFL").Visible = False
        ActiveSheet.Shapes("Menu UnidadFuncionalGFL").Visible = False
        ActiveSheet.Shapes("Menu DPSGFL").Visible = False
        ActiveSheet.Shapes("Menu Equipo CTRGFL").Visible = False
        ActiveSheet.Shapes("Menu ACCESORIOS CTRGFL").Visible = False
        ActiveSheet.Shapes("Menu ETOGFL").Visible = False
        ActiveSheet.Shapes("Menu CONMUTADORESGFL").Visible = False
        ActiveSheet.Shapes("Menu AccesoriosMTMGFL").Visible = False
        ActiveSheet.Shapes("Menu BarrajeGFL").Visible = False
        ActiveSheet.Shapes("Menu TransferenciasGFL").Visible = False
        ActiveSheet.Shapes("Menu ServiciosGFL").Visible = False
        ActiveSheet.Shapes("Menu ColorGFL").Visible = False
        ActiveSheet.Shapes("Menu AcotarGFL").Visible = False
        ActiveSheet.Shapes("Menu EtiquetaGFL").Visible = False
        ActiveSheet.Shapes("Menu ArranqueM GF").Visible = False
    End If
End Sub
Public Sub seleccion_menu_DPSGFL()
    Dim shp As Shape
    Dim groupname As String
    groupname = "Menu DPSGFL"
    
    On Error Resume Next
    Set shp = ActiveSheet.Shapes(groupname)
    On Error GoTo 0
    
    If Not shp Is Nothing Then
        ' Alterna la visibilidad del grupo seleccionado
        shp.Visible = Not shp.Visible
        ' Oculta los dem s grupos
        ActiveSheet.Shapes("Menu EnvolventesGFL").Visible = False
        ActiveSheet.Shapes("Menu UnidadFuncionalGFL").Visible = False
        ActiveSheet.Shapes("Menu LisasGFL").Visible = False
        ActiveSheet.Shapes("Menu Equipo CTRGFL").Visible = False
        ActiveSheet.Shapes("Menu ACCESORIOS CTRGFL").Visible = False
        ActiveSheet.Shapes("Menu ETOGFL").Visible = False
        ActiveSheet.Shapes("Menu CONMUTADORESGFL").Visible = False
        ActiveSheet.Shapes("Menu AccesoriosMTMGFL").Visible = False
        ActiveSheet.Shapes("Menu BarrajeGFL").Visible = False
        ActiveSheet.Shapes("Menu TransferenciasGFL").Visible = False
        ActiveSheet.Shapes("Menu ServiciosGFL").Visible = False
        ActiveSheet.Shapes("Menu ColorGFL").Visible = False
        ActiveSheet.Shapes("Menu AcotarGFL").Visible = False
        ActiveSheet.Shapes("Menu EtiquetaGFL").Visible = False
        ActiveSheet.Shapes("Menu ArranqueM GF").Visible = False
    End If
End Sub
Public Sub seleccion_menu_EquipoCTRGFL()
    Dim shp As Shape
    Dim groupname As String
    groupname = "Menu Equipo CTRGFL"
    
    On Error Resume Next
    Set shp = ActiveSheet.Shapes(groupname)
    On Error GoTo 0
    
    If Not shp Is Nothing Then
        ' Alterna la visibilidad del grupo seleccionado
        shp.Visible = Not shp.Visible
        ' Oculta los dem s grupos
        ActiveSheet.Shapes("Menu EnvolventesGFL").Visible = False
        ActiveSheet.Shapes("Menu UnidadFuncionalGFL").Visible = False
        ActiveSheet.Shapes("Menu LisasGFL").Visible = False
        ActiveSheet.Shapes("Menu DPSGFL").Visible = False
        ActiveSheet.Shapes("Menu ACCESORIOS CTRGFL").Visible = False
        ActiveSheet.Shapes("Menu ETOGFL").Visible = False
        ActiveSheet.Shapes("Menu CONMUTADORESGFL").Visible = False
        ActiveSheet.Shapes("Menu AccesoriosMTMGFL").Visible = False
        ActiveSheet.Shapes("Menu BarrajeGFL").Visible = False
        ActiveSheet.Shapes("Menu TransferenciasGFL").Visible = False
        ActiveSheet.Shapes("Menu ServiciosGFL").Visible = False
        ActiveSheet.Shapes("Menu ColorGFL").Visible = False
        ActiveSheet.Shapes("Menu AcotarGFL").Visible = False
        ActiveSheet.Shapes("Menu EtiquetaGFL").Visible = False
        ActiveSheet.Shapes("Menu ArranqueM GF").Visible = False
    End If
End Sub
Public Sub seleccion_menu_ACCESORIOSCTRGFL()
    Dim shp As Shape
    Dim groupname As String
    groupname = "Menu ACCESORIOS CTRGFL"
    
    On Error Resume Next
    Set shp = ActiveSheet.Shapes(groupname)
    On Error GoTo 0
    
    If Not shp Is Nothing Then
        ' Alterna la visibilidad del grupo seleccionado
        shp.Visible = Not shp.Visible
        ' Oculta los dem s grupos
        ActiveSheet.Shapes("Menu EnvolventesGFL").Visible = False
        ActiveSheet.Shapes("Menu UnidadFuncionalGFL").Visible = False
        ActiveSheet.Shapes("Menu LisasGFL").Visible = False
        ActiveSheet.Shapes("Menu DPSGFL").Visible = False
        ActiveSheet.Shapes("Menu Equipo CTRGFL").Visible = False
        ActiveSheet.Shapes("Menu ETOGFL").Visible = False
        ActiveSheet.Shapes("Menu CONMUTADORESGFL").Visible = False
        ActiveSheet.Shapes("Menu AccesoriosMTMGFL").Visible = False
        ActiveSheet.Shapes("Menu BarrajeGFL").Visible = False
        ActiveSheet.Shapes("Menu TransferenciasGFL").Visible = False
        ActiveSheet.Shapes("Menu ServiciosGFL").Visible = False
        ActiveSheet.Shapes("Menu ColorGFL").Visible = False
        ActiveSheet.Shapes("Menu AcotarGFL").Visible = False
        ActiveSheet.Shapes("Menu EtiquetaGFL").Visible = False
        ActiveSheet.Shapes("Menu ArranqueM GF").Visible = False
    End If
End Sub
Public Sub seleccion_menu_ETOGFL()
    Dim shp As Shape
    Dim groupname As String
    groupname = "Menu ETOGFL"
    
    On Error Resume Next
    Set shp = ActiveSheet.Shapes(groupname)
    On Error GoTo 0
    
    If Not shp Is Nothing Then
        ' Alterna la visibilidad del grupo seleccionado
        shp.Visible = Not shp.Visible
        ' Oculta los dem s grupos
        ActiveSheet.Shapes("Menu EnvolventesGFL").Visible = False
        ActiveSheet.Shapes("Menu UnidadFuncionalGFL").Visible = False
        ActiveSheet.Shapes("Menu LisasGFL").Visible = False
        ActiveSheet.Shapes("Menu DPSGFL").Visible = False
        ActiveSheet.Shapes("Menu Equipo CTRGFL").Visible = False
        ActiveSheet.Shapes("Menu ACCESORIOS CTRGFL").Visible = False
        ActiveSheet.Shapes("Menu CONMUTADORESGFL").Visible = False
        ActiveSheet.Shapes("Menu AccesoriosMTMGFL").Visible = False
        ActiveSheet.Shapes("Menu BarrajeGFL").Visible = False
        ActiveSheet.Shapes("Menu TransferenciasGFL").Visible = False
        ActiveSheet.Shapes("Menu ServiciosGFL").Visible = False
        ActiveSheet.Shapes("Menu ColorGFL").Visible = False
        ActiveSheet.Shapes("Menu AcotarGFL").Visible = False
        ActiveSheet.Shapes("Menu EtiquetaGFL").Visible = False
        ActiveSheet.Shapes("Menu ArranqueM GF").Visible = False
    End If
End Sub
Public Sub seleccion_menu_CONMUTADORESGFL()
    Dim shp As Shape
    Dim groupname As String
    groupname = "Menu CONMUTADORESGFL"
    
    On Error Resume Next
    Set shp = ActiveSheet.Shapes(groupname)
    On Error GoTo 0
    
    If Not shp Is Nothing Then
        ' Alterna la visibilidad del grupo seleccionado
        shp.Visible = Not shp.Visible
        ' Oculta los dem s grupos
        ActiveSheet.Shapes("Menu EnvolventesGFL").Visible = False
        ActiveSheet.Shapes("Menu UnidadFuncionalGFL").Visible = False
        ActiveSheet.Shapes("Menu LisasGFL").Visible = False
        ActiveSheet.Shapes("Menu DPSGFL").Visible = False
        ActiveSheet.Shapes("Menu Equipo CTRGFL").Visible = False
        ActiveSheet.Shapes("Menu ACCESORIOS CTRGFL").Visible = False
        ActiveSheet.Shapes("Menu ETOGFL").Visible = False
        ActiveSheet.Shapes("Menu AccesoriosMTMGFL").Visible = False
        ActiveSheet.Shapes("Menu BarrajeGFL").Visible = False
        ActiveSheet.Shapes("Menu TransferenciasGFL").Visible = False
        ActiveSheet.Shapes("Menu ServiciosGFL").Visible = False
        ActiveSheet.Shapes("Menu ColorGFL").Visible = False
        ActiveSheet.Shapes("Menu AcotarGFL").Visible = False
        ActiveSheet.Shapes("Menu EtiquetaGFL").Visible = False
        ActiveSheet.Shapes("Menu ArranqueM GF").Visible = False
    End If
End Sub
Public Sub seleccion_menu_ACCESORIOSMTMGFL()
    Dim shp As Shape
    Dim groupname As String
    groupname = "Menu AccesoriosMTMGFL"
    
    On Error Resume Next
    Set shp = ActiveSheet.Shapes(groupname)
    On Error GoTo 0
    
    If Not shp Is Nothing Then
        ' Alterna la visibilidad del grupo seleccionado
        shp.Visible = Not shp.Visible
        ' Oculta los dem s grupos
        ActiveSheet.Shapes("Menu EnvolventesGFL").Visible = False
        ActiveSheet.Shapes("Menu UnidadFuncionalGFL").Visible = False
        ActiveSheet.Shapes("Menu LisasGFL").Visible = False
        ActiveSheet.Shapes("Menu DPSGFL").Visible = False
        ActiveSheet.Shapes("Menu Equipo CTRGFL").Visible = False
        ActiveSheet.Shapes("Menu ACCESORIOS CTRGFL").Visible = False
        ActiveSheet.Shapes("Menu ETOGFL").Visible = False
        ActiveSheet.Shapes("Menu CONMUTADORESGFL").Visible = False
        ActiveSheet.Shapes("Menu BarrajeGFL").Visible = False
        ActiveSheet.Shapes("Menu TransferenciasGFL").Visible = False
        ActiveSheet.Shapes("Menu ServiciosGFL").Visible = False
        ActiveSheet.Shapes("Menu ColorGFL").Visible = False
        ActiveSheet.Shapes("Menu AcotarGFL").Visible = False
        ActiveSheet.Shapes("Menu EtiquetaGFL").Visible = False
        ActiveSheet.Shapes("Menu ArranqueM GF").Visible = False
    End If
End Sub
Public Sub seleccion_menu_BarrajeGFL()
    Dim shp As Shape
    Dim groupname As String
    groupname = "Menu BarrajeGFL"
    
    On Error Resume Next
    Set shp = ActiveSheet.Shapes(groupname)
    On Error GoTo 0
    
    If Not shp Is Nothing Then
        ' Alterna la visibilidad del grupo seleccionado
        shp.Visible = Not shp.Visible
        ' Oculta los dem s grupos
        ActiveSheet.Shapes("Menu EnvolventesGFL").Visible = False
        ActiveSheet.Shapes("Menu UnidadFuncionalGFL").Visible = False
        ActiveSheet.Shapes("Menu LisasGFL").Visible = False
        ActiveSheet.Shapes("Menu DPSGFL").Visible = False
        ActiveSheet.Shapes("Menu Equipo CTRGFL").Visible = False
        ActiveSheet.Shapes("Menu ACCESORIOS CTRGFL").Visible = False
        ActiveSheet.Shapes("Menu ETOGFL").Visible = False
        ActiveSheet.Shapes("Menu CONMUTADORESGFL").Visible = False
        ActiveSheet.Shapes("Menu AccesoriosMTMGFL").Visible = False
        ActiveSheet.Shapes("Menu TransferenciasGFL").Visible = False
        ActiveSheet.Shapes("Menu ServiciosGFL").Visible = False
        ActiveSheet.Shapes("Menu ColorGFL").Visible = False
        ActiveSheet.Shapes("Menu AcotarGFL").Visible = False
        ActiveSheet.Shapes("Menu EtiquetaGFL").Visible = False
        ActiveSheet.Shapes("Menu ArranqueM GF").Visible = False
    End If
End Sub
Public Sub seleccion_menu_TransferenciasGFL()
    Dim shp As Shape
    Dim groupname As String
    groupname = "Menu TransferenciasGFL"
    
    On Error Resume Next
    Set shp = ActiveSheet.Shapes(groupname)
    On Error GoTo 0
    
    If Not shp Is Nothing Then
        ' Alterna la visibilidad del grupo seleccionado
        shp.Visible = Not shp.Visible
        ' Oculta los dem s grupos
        ActiveSheet.Shapes("Menu EnvolventesGFL").Visible = False
        ActiveSheet.Shapes("Menu UnidadFuncionalGFL").Visible = False
        ActiveSheet.Shapes("Menu LisasGFL").Visible = False
        ActiveSheet.Shapes("Menu DPSGFL").Visible = False
        ActiveSheet.Shapes("Menu Equipo CTRGFL").Visible = False
        ActiveSheet.Shapes("Menu ACCESORIOS CTRGFL").Visible = False
        ActiveSheet.Shapes("Menu ETOGFL").Visible = False
        ActiveSheet.Shapes("Menu CONMUTADORESGFL").Visible = False
        ActiveSheet.Shapes("Menu AccesoriosMTMGFL").Visible = False
        ActiveSheet.Shapes("Menu BarrajeGFL").Visible = False
        ActiveSheet.Shapes("Menu ServiciosGFL").Visible = False
        ActiveSheet.Shapes("Menu ColorGFL").Visible = False
        ActiveSheet.Shapes("Menu AcotarGFL").Visible = False
        ActiveSheet.Shapes("Menu EtiquetaGFL").Visible = False
        ActiveSheet.Shapes("Menu ArranqueM GF").Visible = False
    End If
End Sub
Public Sub seleccion_menu_ServiciosGFL()
    Dim shp As Shape
    Dim groupname As String
    groupname = "Menu ServiciosGFL"
    
    On Error Resume Next
    Set shp = ActiveSheet.Shapes(groupname)
    On Error GoTo 0
    
    If Not shp Is Nothing Then
        ' Alterna la visibilidad del grupo seleccionado
        shp.Visible = Not shp.Visible
        ' Oculta los dem s grupos
        ActiveSheet.Shapes("Menu EnvolventesGFL").Visible = False
        ActiveSheet.Shapes("Menu UnidadFuncionalGFL").Visible = False
        ActiveSheet.Shapes("Menu LisasGFL").Visible = False
        ActiveSheet.Shapes("Menu DPSGFL").Visible = False
        ActiveSheet.Shapes("Menu Equipo CTRGFL").Visible = False
        ActiveSheet.Shapes("Menu ACCESORIOS CTRGFL").Visible = False
        ActiveSheet.Shapes("Menu ETOGFL").Visible = False
        ActiveSheet.Shapes("Menu CONMUTADORESGFL").Visible = False
        ActiveSheet.Shapes("Menu AccesoriosMTMGFL").Visible = False
        ActiveSheet.Shapes("Menu BarrajeGFL").Visible = False
        ActiveSheet.Shapes("Menu TransferenciasGFL").Visible = False
        ActiveSheet.Shapes("Menu ColorGFL").Visible = False
        ActiveSheet.Shapes("Menu AcotarGFL").Visible = False
        ActiveSheet.Shapes("Menu EtiquetaGFL").Visible = False
        ActiveSheet.Shapes("Menu ArranqueM GF").Visible = False
    End If
End Sub
Public Sub seleccion_menu_ColorGFL()
    Dim shp As Shape
    Dim groupname As String
    groupname = "Menu ColorGFL"
    
    On Error Resume Next
    Set shp = ActiveSheet.Shapes(groupname)
    On Error GoTo 0
    
    If Not shp Is Nothing Then
        ' Alterna la visibilidad del grupo seleccionado
        shp.Visible = Not shp.Visible
        ' Oculta los dem s grupos
        ActiveSheet.Shapes("Menu EnvolventesGFL").Visible = False
        ActiveSheet.Shapes("Menu UnidadFuncionalGFL").Visible = False
        ActiveSheet.Shapes("Menu LisasGFL").Visible = False
        ActiveSheet.Shapes("Menu DPSGFL").Visible = False
        ActiveSheet.Shapes("Menu Equipo CTRGFL").Visible = False
        ActiveSheet.Shapes("Menu ACCESORIOS CTRGFL").Visible = False
        ActiveSheet.Shapes("Menu ETOGFL").Visible = False
        ActiveSheet.Shapes("Menu CONMUTADORESGFL").Visible = False
        ActiveSheet.Shapes("Menu AccesoriosMTMGFL").Visible = False
        ActiveSheet.Shapes("Menu BarrajeGFL").Visible = False
        ActiveSheet.Shapes("Menu TransferenciasGFL").Visible = False
        ActiveSheet.Shapes("Menu ServiciosGFL").Visible = False
        ActiveSheet.Shapes("Menu AcotarGFL").Visible = False
        ActiveSheet.Shapes("Menu EtiquetaGFL").Visible = False
        ActiveSheet.Shapes("Menu ArranqueM GF").Visible = False
    End If
End Sub
Public Sub seleccion_menu_AcotarGFL()
    Dim shp As Shape
    Dim groupname As String
    groupname = "Menu AcotarGFL"
    
    On Error Resume Next
    Set shp = ActiveSheet.Shapes(groupname)
    On Error GoTo 0
    
    If Not shp Is Nothing Then
        ' Alterna la visibilidad del grupo seleccionado
        shp.Visible = Not shp.Visible
        ' Oculta los dem s grupos
        ActiveSheet.Shapes("Menu EnvolventesGFL").Visible = False
        ActiveSheet.Shapes("Menu UnidadFuncionalGFL").Visible = False
        ActiveSheet.Shapes("Menu LisasGFL").Visible = False
        ActiveSheet.Shapes("Menu DPSGFL").Visible = False
        ActiveSheet.Shapes("Menu Equipo CTRGFL").Visible = False
        ActiveSheet.Shapes("Menu ACCESORIOS CTRGFL").Visible = False
        ActiveSheet.Shapes("Menu ETOGFL").Visible = False
        ActiveSheet.Shapes("Menu CONMUTADORESGFL").Visible = False
        ActiveSheet.Shapes("Menu AccesoriosMTMGFL").Visible = False
        ActiveSheet.Shapes("Menu BarrajeGFL").Visible = False
        ActiveSheet.Shapes("Menu TransferenciasGFL").Visible = False
        ActiveSheet.Shapes("Menu ServiciosGFL").Visible = False
        ActiveSheet.Shapes("Menu ColorGFL").Visible = False
        ActiveSheet.Shapes("Menu EtiquetaGFL").Visible = False
        ActiveSheet.Shapes("Menu ArranqueM GF").Visible = False
    End If
End Sub
Public Sub seleccion_menu_EtiquetaGFL()
    Dim shp As Shape
    Dim groupname As String
    groupname = "Menu EtiquetaGFL"
    
    On Error Resume Next
    Set shp = ActiveSheet.Shapes(groupname)
    On Error GoTo 0
    
    If Not shp Is Nothing Then
        ' Alterna la visibilidad del grupo seleccionado
        shp.Visible = Not shp.Visible
        ' Oculta los dem s grupos
        ActiveSheet.Shapes("Menu EnvolventesGFL").Visible = False
        ActiveSheet.Shapes("Menu UnidadFuncionalGFL").Visible = False
        ActiveSheet.Shapes("Menu LisasGFL").Visible = False
        ActiveSheet.Shapes("Menu DPSGFL").Visible = False
        ActiveSheet.Shapes("Menu Equipo CTRGFL").Visible = False
        ActiveSheet.Shapes("Menu ACCESORIOS CTRGFL").Visible = False
        ActiveSheet.Shapes("Menu ETOGFL").Visible = False
        ActiveSheet.Shapes("Menu CONMUTADORESGFL").Visible = False
        ActiveSheet.Shapes("Menu AccesoriosMTMGFL").Visible = False
        ActiveSheet.Shapes("Menu BarrajeGFL").Visible = False
        ActiveSheet.Shapes("Menu TransferenciasGFL").Visible = False
        ActiveSheet.Shapes("Menu ServiciosGFL").Visible = False
        ActiveSheet.Shapes("Menu ColorGFL").Visible = False
        ActiveSheet.Shapes("Menu AcotarGFL").Visible = False
        ActiveSheet.Shapes("Menu ArranqueM GF").Visible = False
    End If
End Sub
Public Sub seleccion_menu_ArranqueM_GFL()
    Dim shp As Shape
    Dim groupname As String
    groupname = "Menu ArranqueM GF"
    
    On Error Resume Next
    Set shp = ActiveSheet.Shapes(groupname)
    On Error GoTo 0
    
    If Not shp Is Nothing Then
        ' Alterna la visibilidad del grupo seleccionado
        shp.Visible = Not shp.Visible
        ' Oculta los dem s grupos
        ActiveSheet.Shapes("Menu EnvolventesGFL").Visible = False
        ActiveSheet.Shapes("Menu UnidadFuncionalGFL").Visible = False
        ActiveSheet.Shapes("Menu LisasGFL").Visible = False
        ActiveSheet.Shapes("Menu DPSGFL").Visible = False
        ActiveSheet.Shapes("Menu Equipo CTRGFL").Visible = False
        ActiveSheet.Shapes("Menu ACCESORIOS CTRGFL").Visible = False
        ActiveSheet.Shapes("Menu ETOGFL").Visible = False
        ActiveSheet.Shapes("Menu CONMUTADORESGFL").Visible = False
        ActiveSheet.Shapes("Menu AccesoriosMTMGFL").Visible = False
        ActiveSheet.Shapes("Menu BarrajeGFL").Visible = False
        ActiveSheet.Shapes("Menu TransferenciasGFL").Visible = False
        ActiveSheet.Shapes("Menu ServiciosGFL").Visible = False
        ActiveSheet.Shapes("Menu ColorGFL").Visible = False
        ActiveSheet.Shapes("Menu AcotarGFL").Visible = False
        ActiveSheet.Shapes("Menu EtiquetaGFL").Visible = False
    End If
End Sub

'================================================
Public Sub Area_GFL_1()
    ThisWorkbook.Sheets("DISENO_GF").Range("BZ39:KJ155").Clear
End Sub
Public Sub Area_GFL_2()
    ThisWorkbook.Sheets("DISENO_GF").Range("LJ89:QI182").Clear
End Sub


