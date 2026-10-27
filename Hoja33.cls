'=========== VISTA_VP_CLIENTES ===========
Sub Traercolumna()
    ' Copia el rango A1:C10 y lo pega en E1
    Sheets("FrontalPuerta").Range("CE48:FV155").Copy
    Sheets("Vista_puerta_cliente").Range("CE48").PasteSpecial xlPasteAll
    
    ' Limpia el portapapeles
    Application.CutCopyMode = False
End Sub
Public Sub Area_GFL_1()
    ThisWorkbook.Sheets("Vista_puerta_cliente").Range("CF46:GA146").Clear
    ThisWorkbook.Sheets("Vista_puerta_cliente").Range("CE150:FV155").Clear
End Sub