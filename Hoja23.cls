'==== Diseno B-Configurator
Sub Traercolumna()
    ' Copia el rango A1:C10 y lo pega en E1
    Sheets("VistaFrontal").Range("CE48:FV155").Copy
    Sheets("B_Conf").Range("CE48").PasteSpecial xlPasteAll
    
    ' Limpia el portapapeles
    Application.CutCopyMode = False
End Sub
Public Sub Worksheet_Activate()
    'Olcultar menus
        ActiveSheet.Shapes("Menu Barraje Vertical").Visible = False
        ActiveSheet.Shapes("Menu Barraje Horizontal").Visible = False
        ActiveSheet.Shapes("Menu Barraje Tierra").Visible = False
        ActiveSheet.Shapes("Menu Barraje Neutro").Visible = False
        ActiveSheet.Shapes("Menu Main Horizontal").Visible = False
    '====VBB
        With Me.OLEObjects("VBBInteruptor").Object
            .clear
            .AddItem "EZC400"
            .AddItem "CVS400"
            .AddItem "NSX400"
        end With
    '====HBB
        With Me.OLEObjects("HBBAmperaje").Object
            .clear
            .AddItem "400A"
        end With
    '====GBB
        With Me.OLEObjects("GBBAncho").Object
            .clear
            .AddItem "400"
            .AddItem "720"
        end With
        With Me.OLEObjects("GBB_ICC").Object
            .clear
            .addItem "25"
        end With
    '====NBB
        With Me.OLEObjects("NBBAncho").Object  
            .clear
            .AddItem "400"
            .AddItem "720"
        end With
        With Me.OLEObjects("NBBPorcentaje").Object
            .clear
            .addItem "100"
        end With
    '====MAIN
        With Me.OLEObjects("MAINInterruptor").Object
            .clear
            .AddItem "GoPact100"
            .AddItem "GoPact200"
            .AddItem "GoPact250"
            .AddItem "EZC100"
            .AddItem "EZC250"
            .AddItem "CVS250"
            .AddItem "NSX250"
        end With
        With Me.OLEObjects("MAINOrientacion").Object
            .clear
            .AddItem "HC"
        end With
End Sub
Public Sub Verticales()
    '========================
        celda = "KR91"
        ThisWorkbook.Sheets("B_Conf").Range("KP82:OJ184").Clear
        Set wsActual = ThisWorkbook.Sheets("B_Conf")
    '========================
    Dim Dato1 As String
    Dato1 = Me.OLEObjects("VBBInteruptor").Object.Value
    ResultadoVBB = "VBB_" & Dato1 & "X400"
    'MSGBOX ResultadoVBB
    MoverBloque ResultadoVBB, "Barras", wsActual, celda
End Sub
Public Sub Horizontales()
    '========================
        celda = "KR91"
        ThisWorkbook.Sheets("B_Conf").Range("KP82:OJ184").Clear
        Set wsActual = ThisWorkbook.Sheets("B_Conf")
    '========================
    Dim Dato1 As String
    Dato1 = Me.OLEObjects("HBBAmperaje").Object.Value
    ResultadoHBB = "HBBX720_" & Dato1
    MoverBloque ResultadoHBB, "Barras", wsActual, celda
End Sub
public Sub GBB()
    '========================
        celda = "KR102"
        ThisWorkbook.Sheets("B_Conf").Range("KP82:OJ184").Clear
        Set wsActual = ThisWorkbook.Sheets("B_Conf")
    '========================
    Dim Dato1 As String
    Dim Dato2 As String
    Dato1 = Me.OLEObjects("GBBAncho").Object.Value
    Dato2 = Me.OLEObjects("GBB_ICC").Object.Value
    ResultadoGBB = "GBBX" & Dato1 & "X" & Dato2
   ' MSGBOX ResultadoGBB
    MoverBloque ResultadoGBB, "Barras", wsActual, celda
End Sub
Public Sub NBB()
    '========================   
        celda = "KR102"
        ThisWorkbook.Sheets("B_Conf").Range("KP82:OJ184").Clear
        Set wsActual = ThisWorkbook.Sheets("B_Conf")
    '========================
    Dim Dato1 As String
    Dim Dato2 As String
    Dato1 = Me.OLEObjects("NBBAncho").Object.Value
    Dato2 = Me.OLEObjects("NBBPorcentaje").Object.Value
    ResultadoNBB = "NBBX" & Dato1 & "X" & Dato2
    MoverBloque ResultadoNBB, "Barras", wsActual, celda
End Sub
Public Sub MAIN()
    '========================
        celda = "KR99"
        ThisWorkbook.Sheets("B_Conf").Range("KP82:OJ184").Clear
        Set wsActual = ThisWorkbook.Sheets("B_Conf")
    '========================
    Dim Dato1 As String , Dato2 As String   
    Dato1 = Me.OLEObjects("MAINInterruptor").Object.Value
    Dato2 = Me.OLEObjects("MAINOrientacion").Object.Value
    ResultadoMAIN = "MSFI_" & Dato1 & "ToX720" & Dato2
    MoverBloque ResultadoMAIN, "Main", wsActual, celda
End Sub
'======================= CARGA DE DATOS
    Public sub Bucar_datos_VBB()
        N_INICIO = 64
        N_FIN = 64
        Celda_cubiculo ="KU91"
        Informe_Barras ResultadoVBB
    end sub
    Public sub Bucar_datos_HBB()
        N_INICIO = 76
        N_FIN = 76
        Celda_cubiculo ="KV91"
        Informe_Barras ResultadoHBB
    end sub
    Public sub Bucar_datos_GBB()
        N_INICIO = 88
        N_FIN = 91
        Celda_cubiculo ="KR102"
        Informe_Barras_GBB ResultadoGBB
    end sub
    Public sub Bucar_datos_NBB()
        N_INICIO = 88
        N_FIN = 91
        'Celda_cubiculo =""
        Informe_Barras_GBB ResultadoNBB
    end sub
    Public sub Bucar_datos_MAIN()
        N_INICIO = 106
        N_FIN = 145
        Celda_cubiculo ="KU99"
        Informe_Barras_MAIN ResultadoMAIN
    end sub
