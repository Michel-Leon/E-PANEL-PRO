'Vista puerta
Sub BorrarCeldas()

    'Borrar contenido y formato de relleno
    With Range("BW51:JG147,JY51:RI147")
        .ClearContents
        .Interior.Pattern = xlNone
        .Borders.LineStyle = xlNone
    End With

    'Desagrupar filas
    On Error Resume Next
    Rows("51:147").Ungroup

    'Desagrupar columnas
    Range("BW:JG").Columns.Ungroup
    Range("JY:RI").Columns.Ungroup
    On Error GoTo 0

End Sub
