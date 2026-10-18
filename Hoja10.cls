'=========== VISTA_FM_CLIENTES ===========
Sub Traercolumna()
    ' Copia el rango A1:C10 y lo pega en E1
    Sheets("FrontalFM").Range("BV48:FJ155").Copy
    Sheets("Vista_FM_CLIENTE").Range("BV48").PasteSpecial xlPasteAll
    
    ' Limpia el portapapeles
    Application.CutCopyMode = False
End Sub
