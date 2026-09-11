Attribute VB_Name = "Módulo3"
'=================================
'ESCRIBIR CARACTERISTIOCAS DEL PRODUCTO
'=================================
Public MenuUnidades As String
Public Sub Escribir_caracteristicas_TB()
    Dim wsActual As Worksheet
   Set wsActual = ThisWorkbook.Sheets("INICIO")
    With wsActual
        .Range("A1").Value = "true"
        .Range("DK89").Value = "1600"
        .Range("DK97").Value = "200,400,720"
        .Range("DK105").Value = "400"
        .Range("DK113").Value = "30,31"
        .Range("DK134").Value = "630"
        .Range("DK142").Value = "25"
    End With
  MenuUnidades = "Opcion1"
End Sub
Public Sub Escribir_caracteristicas_TC()
    Dim wsActual As Worksheet
   Set wsActual = ThisWorkbook.Sheets("INICIO")
    With wsActual
        .Range("A1").Value = "true"
        .Range("DK89").Value = "1800,2000"
        .Range("DK97").Value = "200,400,720"
        .Range("DK105").Value = "400"
        .Range("DK113").Value = "30, 31, 42, 54"
        .Range("DK134").Value = "1250"
        .Range("DK142").Value = "25-40"
    End With
    MenuUnidades = "Opcion1"
End Sub
Public Sub Escribir_caracteristicas_TM()
    Dim wsActual As Worksheet
     Set wsActual = ThisWorkbook.Sheets("INICIO")
    With wsActual
        .Range("A1").Value = "true"
        .Range("DK89").Value = "1800,2000,2200"
        .Range("DK97").Value = "200,400,720,800"
        .Range("DK105").Value = "600,800,1000"
        .Range("DK113").Value = "31,42,54,55,65"
        .Range("DK134").Value = "<=4000"
        .Range("DK142").Value = "40-65"
    End With
    MenuUnidades = "Opcion2"
End Sub
Public Sub Escribir_caracteristicas_TH()
    Dim wsActual As Worksheet
   Set wsActual = ThisWorkbook.Sheets("INICIO")
    With wsActual
        .Range("A1").Value = "true"
        .Range("DK89").Value = "2000,2200"
        .Range("DK97").Value = "400,720,1120"
        .Range("DK105").Value = "1000,400+1000"
        .Range("DK113").Value = "31,42,54"
        .Range("DK134").Value = ">4000"
        .Range("DK142").Value = "40-85"
    End With
    Menu2Unidades = "Opcion2"
End Sub
