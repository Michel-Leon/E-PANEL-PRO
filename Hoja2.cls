'============== vista columna ===================
private sub Worksheet_Activate()
    dim ws As Worksheet
    set ws = ThisWorkbook.Sheets("Portada_Mestra")
    dim valordimensiones As String
    valordimensiones = ws.Range("C14").Value
    Linea = thisWorkbook.Sheets("Cabecera").Range("E3").Value
    with Me.oleobjects("Dimensiones").Object
        .clear
        .additem valordimensiones
    end with
End Sub

Public sub Envolventes ()
    ThisWorkbook.Sheets("Vista FM").Range("JR87:NP182").Clear
    dim dato as string
    dato = Me.oleobjects("Dimensiones").Object.Value
    Select case Linea
    ResultadoENV = "EFM_" & dato
End Sub
