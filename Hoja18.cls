'===================
Public Sub Worksheet_Activate()

    Activesheet.shapes("Diseño").Visible = false
    dim wsActual As Worksheet
     set  wsActual = thisWorkbook.Sheets("INICIO")
    With wsActual
        .Range("DK89").value =""
        .Range("DK97").value =""
        .Range("DK105").value =""
        .Range("DK113").value =""
        .Range("DK134").value =""
        .Range("DK142").value =""
    end with
End Sub
