Attribute VB_Name = "M�dulo1"

'============== vARIABLES GLOBALES ===================
Public Linea As String
Public Celda As String
Public wsActual As Worksheet
Public ResultadoENV As String
Public ResultadoENVP As String
Public ResultadoENVL As String
Public ResultadoENVPS As String
'============== Mover Bloque ===================
Sub MoverBloque(ref As String, BD As String)
    Dim origen As Range
    Dim destino As Range
    Dim wsDatos As Worksheet
    

    Set wsDatos = ThisWorkbook.Sheets(BD)
    

    On Error GoTo errHandler
    Set origen = wsDatos.Range(ref)
    Set destino = wsActual.Range(Celda)

    origen.Copy Destination:=destino

    'wsActual.Range(destino, destino.Offset(origen.Rows.Count - 1, origen.Columns.Count - 1)).Select
    Exit Sub

    errHandler:
            MsgBox "Referencia no encontrada " & ref, vbCritical

End Sub
'======================================================================================================
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
        'ActiveSheet.Shapes("Complementos de dibujo").Visible = False
        
    End If
End Sub
Sub Vista_puerta(ByVal resEnvP As String)

    WITH ThisWorkbook.Sheets("Vista Puerta").Range("BX44:IX144")
        .Clear
        .UnMerge
    End With
    Set wsActual = ThisWorkbook.Sheets("Vista Puerta")
    Select case thisWorkbook.Sheets("Vista FM").Range("A1").Value  
        Case "600":Celda = "BX97"
        Case "700":Celda = "BX89"
        Case "900":Celda ="BX73"
        Case "800":Celda = "BX81"
        Case "110":Celda = "BX57"
    End select 
    'MsgBox resEnvP
    MoverBloque resEnvP, "ENVOLVENTES_PUERTA"
End Sub
sub Vista_lateral(ByVal resEnvL As String)
    With ThisWorkbook.Sheets("Vista Lateral").Range("CG49:IR145")
        .Clear
        .UnMerge
    End With
    Set wsActual = ThisWorkbook.Sheets("Vista Lateral")
    Select case thisWorkbook.Sheets("Vista FM").Range("A1").Value  
        Case "600":Celda = "CH98"
        Case "700":Celda = "CH89"
        Case "900":Celda ="CH74"
        Case "800":Celda = "CH82"
        Case "110":Celda = "CH58"
    End select    
    'MsgBox resEnvL
    MoverBloque resEnvL, "ENVOLVENTES L-Post"
End Sub
Sub Vista_Posterior(ByVal resEnvPS As String)
    With ThisWorkbook.Sheets("Vista Posterior").Range("CH49:IX144")
        .Clear
        .UnMerge
    End With
    Set wsActual = ThisWorkbook.Sheets("Vista Posterior")
    Select case thisWorkbook.Sheets("Vista FM").Range("A1").Value  
        Case "600":Celda = "BX97"
        Case "700":Celda = "BX89"
        Case "900":Celda ="BX73"
        Case "800":Celda = "BX81"
        Case "110":Celda = "BX57"
    End select    
    'MsgBox resEnvPS
    MoverBloque resEnvPS, "ENVOLVENTES L-Post"
End Sub

