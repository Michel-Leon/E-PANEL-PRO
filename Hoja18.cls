'===================
Public Sub Worksheet_Activate()
    Dim wsActual As Worksheet
     Set wsActual = ThisWorkbook.Sheets("INICIO")

    With wsActual
        .Range("A1").Value = ""
        .Range("DK89").Value = ""
        .Range("DK97").Value = ""
        .Range("DK105").Value = ""
        .Range("DK113").Value = ""
        .Range("DK134").Value = ""
        .Range("DK142").Value = ""
    End With
End Sub

