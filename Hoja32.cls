'=========== VISTA_FM_CLIENTES ===========
Sub Traercolumna()
    ' Copia el rango A1:C10 y lo pega en E1
    Sheets("VistaFrontal").Range("CE48:FV155").Copy
    Sheets("Vista_FM_CLIENTE").Range("CE48").PasteSpecial xlPasteAll
    
    ' Limpia el portapapeles
    Application.CutCopyMode = False
End Sub
Public Sub Area_GFL_1()
    ThisWorkbook.Sheets("Vista_FM_Cliente").Range("CF46:GA146").Clear
    ThisWorkbook.Sheets("Vista_FM_Cliente").Range("CE150:FV155").Clear
End Sub