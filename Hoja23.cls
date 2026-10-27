'==== Diseno B-Configurator

Public Sub Worksheet_Activate()
    '====VBB
        With Me.OLEObjects("VBBInteruptor").Object
            .clear
            .AddItem "EZC400"
            .AddItem "CVS400"
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

