'========= VISTA_P_CLIENTE ==========
Sub Traercolumna()
    ' Copia el rango A1:C10 y lo pega en E1
    Sheets("FrontalPuerta").Range("BV48:FJ155").Copy
    Sheets("Vista_P_CLIENTE").Range("BV48").PasteSpecial xlPasteAll
    
    ' Limpia el portapapeles
    Application.CutCopyMode = False
End Sub
