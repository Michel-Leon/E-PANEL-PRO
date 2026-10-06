Attribute VB_Name = "Color_Base_RAL7035"
Public Sub CambiodecolorBASE()
    Dim ws As Worksheet
    Dim celda As Range
    Dim valorColor As String
    Dim colorDestino As Long
    
    ' Obtener el valor del color desde la portada
    valorColor = "RAL 7035"
    Set ws = ThisWorkbook.Sheets("LISAS")
    ' Definir color destino seg?n el valor
    Select Case valorColor
        Case "RAL 7035": colorDestino = RGB(197, 199, 196)
        Case "RAL 7032": colorDestino = RGB(181, 176, 161)
        Case "RAL 9003": colorDestino = RGB(236, 236, 231)
        Case "RAL 9004": colorDestino = RGB(43, 43, 44)
        Case "Naranja": colorDestino = RGB(201, 133, 12)
        Case "Amarillo": colorDestino = RGB(235, 198, 21)
    End Select
    
    ' Recorrer el rango y cambiar solo los colores espec?ficos
    For Each celda In ws.Range("I1:CB2370")
        If valorColor = "RAL 7035" Then
            If celda.Interior.Color = RGB(200, 200, 200) Then
                celda.Interior.Color = RGB(197, 199, 196)
            End If
        End If
    Next celda
    
    MsgBox "fINALIZADO"
End Sub
