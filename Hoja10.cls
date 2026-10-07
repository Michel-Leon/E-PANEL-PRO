Sub BorrarCeldas()

    'Borrar contenido y formato de relleno
    With Range("CH51:JG147,KJ51:RI147")
        .ClearContents
        .Interior.Pattern = xlNone
    End With

    'Desagrupar filas
    On Error Resume Next
    Rows("51:147").Ungroup

    'Desagrupar columnas
    Range("CH:JG").Columns.Ungroup
    Range("KJ:RI").Columns.Ungroup
    On Error GoTo 0

End Sub

