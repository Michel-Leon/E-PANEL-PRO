'lOGIN
Public Sub Worksheet_Activate()
 With ThisWorkbook.Sheets("LOGIN")
    .Range("IU83").Value = ""
    .Range("IU98").Value = ""
 End With
End Sub
