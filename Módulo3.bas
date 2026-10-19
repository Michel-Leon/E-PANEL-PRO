Attribute VB_Name = "M�dulo3"
'=================================
'ESCRIBIR CARACTERISTIOCAS DEL PRODUCTO
'=================================
public MenuUnidades as string
Public sub Escribir_caracteristicas_TB()
    dim wsActual As Worksheet
   set  wsActual = thisWorkbook.Sheets("INICIO")
    With wsActual
        .Range("A1").value = "true"
        .Range("DK89").value ="1600"
        .Range("DK97").value ="200,400,720"
        .Range("DK105").value ="400"
        .Range("DK113").value ="30,31"
        .Range("DK134").value ="630"
        .Range("DK142").value ="25"
    end with
  MenuUnidades = "Opcion0"
  CrearVinculoAHoja
End sub   
Public sub Escribir_caracteristicas_TC()
    dim wsActual As Worksheet
   set  wsActual = thisWorkbook.Sheets("INICIO")
    With wsActual
        .Range("A1").value = "true"
        .Range("DK89").value ="1800,2000"
        .Range("DK97").value ="200,400,720"
        .Range("DK105").value ="400"
        .Range("DK113").value ="30, 31, 42, 54"
        .Range("DK134").value ="1250"
        .Range("DK142").value ="25-40"
    end with
    MenuUnidades = "Opcion1"
    CrearVinculoAHoja
End sub  
Public sub Escribir_caracteristicas_TM()
    dim wsActual As Worksheet
     set  wsActual = thisWorkbook.Sheets("INICIO")
    With wsActual 
        .Range("A1").value = "true"
        .Range("DK89").value ="1800,2000,2200"
        .Range("DK97").value ="200,400,720,800"
        .Range("DK105").value ="600,800,1000"
        .Range("DK113").value ="31,42,54"
        .Range("DK134").value ="<=4000"
        .Range("DK142").value ="40-65"
    end with
    MenuUnidades = "Opcion2"
    CrearVinculoAHoja
End sub  
Public sub Escribir_caracteristicas_TH()
    dim wsActual As Worksheet
   set  wsActual = thisWorkbook.Sheets("INICIO")
    With wsActual
        .Range("A1").value = "true"
        .Range("DK89").value ="2000,2200"
        .Range("DK97").value ="400,720,1120"
        .Range("DK105").value ="1000,400+1000"
        .Range("DK113").value ="31,42,54"
        .Range("DK134").value =">4000"
        .Range("DK142").value ="40-85"
    end with
    Menu2Unidades = "Opcion2"
    CrearVinculoAHoja
End sub  
Public sub Escribir_caracteristicas_Neo()
    dim wsActual As Worksheet
   set  wsActual = thisWorkbook.Sheets("INICIO")
    With wsActual
        .Range("A1").value = "true"
        .Range("A2").value ="Neo"
        .Range("DK89").value ="1100,900,600"
        .Range("DK97").value ="400+720,720,400"
        .Range("DK105").value ="400"
        .Range("DK113").value ="31,42,54,55,65,66"
        .Range("DK134").value ="400"
        .Range("DK142").value ="25"
    end with
    MenuUnidades = "Opcion3"
    CrearVinculoAHoja
End sub
Public sub Escribir_caracteristicas_Basic()
    dim wsActual As Worksheet
   set  wsActual = thisWorkbook.Sheets("INICIO")
    With wsActual
        .Range("A1").value = "true"
        .Range("A2").value ="Basic"
        .Range("DK89").value ="900,800,700,600"
        .Range("DK97").value =" 720,620,520,400"
        .Range("DK105").value ="300"
        .Range("DK113").value ="31,42,54,55,65,66"
        .Range("DK134").value ="MENOR A 100A"
        .Range("DK142").value ="10"
    end with
    MenuUnidades = "Opcion4"
    CrearVinculoAHoja
End sub
Sub CrearVinculoAHoja()
    Dim nombreHoja As String
    
    Select Case MenuUnidades
        Case "Opcion0", "Opcion1", "Opcion2"
            nombreHoja = "DISENO"
        Case "Opcion3", "Opcion4"
            nombreHoja = "DISENO_GF"
        Case Else
            Exit Sub
    End Select
    
    If Not ExisteHoja(nombreHoja) Then
        MsgBox "La hoja '" & nombreHoja & "' no existe.", vbExclamation
        Exit Sub
    End If
    
    With ThisWorkbook.Sheets("Inicio")
        .Range("GS151:GY151").Hyperlinks.Delete
        .Hyperlinks.Add Anchor:=.Range("GS151:GY151"), _
            Address:="", _
            SubAddress:="'" & nombreHoja & "'!A1", _
            TextToDisplay:="Ir a " & nombreHoja
         ' Formato: fondo azul #0065BB y texto blanco
        With .Range("GS151:GY151")
            .Interior.Color = RGB(0, 101, 187)
            .Font.Color = RGB(255, 255, 255)
            .Font.Underline = xlUnderlineStyleNone
        End With    
    End With
End Sub

Function ExisteHoja(nombre As String) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(nombre)
    ExisteHoja = Not ws Is Nothing
End Function