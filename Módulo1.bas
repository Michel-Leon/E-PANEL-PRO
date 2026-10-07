Attribute VB_Name = "M�dulo1"
Sub BorrarCeldas()

    'Confirmacion antes de borrar
    If MsgBox("¿Esta seguro de borrar el contenido?", vbYesNo + vbQuestion, "Confirmar borrado") = vbNo Then Exit Sub

    'Borrar contenido, color y bordes en una sola linea
    ActiveSheet.Range("BW51:JG147,JY51:RI147").Clear

    'Desagrupar filas y columnas
    On Error Resume Next
    ActiveSheet.Rows("51:147").Ungroup
    ActiveSheet.Range("BW:JG").Columns.Ungroup
    ActiveSheet.Range("JY:RI").Columns.Ungroup
    On Error GoTo 0

End Sub

Sub BorrarCeldas_SUP_IF()

    'Confirmacion antes de borrar
    If MsgBox("¿Esta seguro de borrar el contenido?", vbYesNo + vbQuestion, "Confirmar borrado") = vbNo Then Exit Sub

    'Borrar contenido, color y bordes en una sola linea
    ActiveSheet.Range("BX51:JC122,KB51:RG122").Clear

    'Desagrupar filas y columnas
    On Error Resume Next
    ActiveSheet.Rows("51:147").Ungroup
    ActiveSheet.Range("BW:JG").Columns.Ungroup
    ActiveSheet.Range("JY:RI").Columns.Ungroup
    On Error GoTo 0

End Sub

Sub BorrarCeldas_LAT()

    'Confirmacion antes de borrar
    If MsgBox("¿Esta seguro de borrar el contenido?", vbYesNo + vbQuestion, "Confirmar borrado") = vbNo Then Exit Sub

    'Borrar contenido, color y bordes en una sola linea
    ActiveSheet.Range("CH51:JC147,KL51:RG147").Clear

    'Desagrupar filas y columnas
    On Error Resume Next
    ActiveSheet.Rows("51:147").Ungroup
    ActiveSheet.Range("BW:JG").Columns.Ungroup
    ActiveSheet.Range("JY:RI").Columns.Ungroup
    On Error GoTo 0

End Sub