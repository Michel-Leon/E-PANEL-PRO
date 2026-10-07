Sub BorrarCeldas()

    'Borrar contenido y formato de relleno
    With Range("BW51:JG122,JY51:RI122")
        .ClearContents
        .Interior.Pattern = xlNone
    End With

    'Desagrupar filas
    On Error Resume Next
    Rows("51:122").Ungroup

    'Desagrupar columnas
    Range("BW:JG").Columns.Ungroup
    Range("JY:RI").Columns.Ungroup
    On Error GoTo 0

End Sub


