Attribute VB_Name = "M�dulo3"
'=================================
'ESCRIBIR CARACTERISTIOCAS DEL PRODUCTO
'=================================
Public sub Escribir_caracteristicas_TB()
    dim wsActual As Worksheet
   set  wsActual = thisWorkbook.Sheets("INICIO")
    With wsActual
        .Range("DK89").value ="1600"
        .Range("DK97").value ="200,400,720"
        .Range("DK105").value ="400"
        .Range("DK113").value ="30,31"
        .Range("DK134").value ="630"
        .Range("DK142").value ="25"
    end with
    Activesheet.shapes("Diseño").Visible = true
End sub   
Public sub Escribir_caracteristicas_TC()
    dim wsActual As Worksheet
   set  wsActual = thisWorkbook.Sheets("INICIO")
    With wsActual
        .Range("DK89").value ="1800,2000"
        .Range("DK97").value ="200,400,720"
        .Range("DK105").value ="400"
        .Range("DK113").value ="30, 31, 42, 54"
        .Range("DK134").value ="1250"
        .Range("DK142").value ="25-40"
    end with
    Activesheet.shapes("Diseño").Visible = true
End sub  
Public sub Escribir_caracteristicas_TM()
    dim wsActual As Worksheet
     set  wsActual = thisWorkbook.Sheets("INICIO")
    With wsActual
        .Range("DK89").value ="1800,2000,2200"
        .Range("DK97").value ="200,400,720,800"
        .Range("DK105").value ="600,800,1000"
        .Range("DK113").value ="31,42,54,55,65"
        .Range("DK134").value ="<=4000"
        .Range("DK142").value ="40-65"
    end with
    Activesheet.shapes("Diseño").Visible = true
End sub  
Public sub Escribir_caracteristicas_TH()
    dim wsActual As Worksheet
   set  wsActual = thisWorkbook.Sheets("INICIO")
    With wsActual
        .Range("DK89").value ="1800,2000,2200"
        .Range("DK97").value ="200,400,720,800"
        .Range("DK105").value ="600,800,1000"
        .Range("DK113").value ="31,42,54"
        .Range("DK134").value =">4000"
        .Range("DK142").value ="40-85"
    end with
    Activesheet.shapes("Diseño").Visible = true
End sub  