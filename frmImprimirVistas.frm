VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmImprimirVistas 
   ClientHeight    =   2330
   ClientLeft      =   -30
   ClientTop       =   -150
   ClientWidth     =   3340
   OleObjectBlob   =   "frmImprimirVistas.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "frmImprimirVistas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'==============================================================
'  C?DIGO DEL USERFORM  "frmImprimirVistas"
'  (El formulario se deja VAC?O: los controles se crean solos)
'==============================================================
Option Explicit

Private WithEvents btnPDF As MSForms.CommandButton
Attribute btnPDF.VB_VarHelpID = -1
Private WithEvents btnImprimir As MSForms.CommandButton
Attribute btnImprimir.VB_VarHelpID = -1
Private WithEvents btnCancelar As MSForms.CommandButton
Attribute btnCancelar.VB_VarHelpID = -1
Private WithEvents cboTodas As MSForms.ComboBox
Attribute cboTodas.VB_VarHelpID = -1

Private cboVista(1 To NUM_VISTAS) As MSForms.ComboBox
Private chkOmitir As MSForms.CheckBox
Private chkAbrir As MSForms.CheckBox
Private txtCopias As MSForms.TextBox

Private Sub UserForm_Initialize()
    Dim vistas As Variant, v As Long, y As Single

    vistas = NombresVistas()
    Me.Caption = "Paquete de impresion de planos"
    Me.Width = 350                      'antes 480
    y = 10

    'Encabezados
    AgregarLabel "Vista", 12, y, 110, True
    AgregarLabel "Zonas a incluir", 125, y, 205, True
    y = y + 18

    'Una fila por vista
    For v = 1 To NUM_VISTAS
        AgregarLabel vistas(v - 1), 12, y + 3, 110, False

        Set cboVista(v) = Me.Controls.Add("Forms.ComboBox.1", "cboVista" & v)
        With cboVista(v)
            .Left = 125: .Top = y: .Width = 205      'antes 160
            .Style = fmStyleDropDownList
            .List = OpcionesZona()
            .ListIndex = 4          'Automatico por defecto
        End With
        '(se elimino el label de estado Z1 / Z2)
        y = y + 22
    Next v

    'Aplicar a todas
    y = y + 6
    AgregarLabel "Aplicar a todas:", 12, y + 3, 110, True
    Set cboTodas = Me.Controls.Add("Forms.ComboBox.1", "cboTodas")
    With cboTodas
        .Left = 125: .Top = y: .Width = 205          'antes 160
        .Style = fmStyleDropDownList
        .List = OpcionesZona()
    End With
    y = y + 30

    'Opciones
    Set chkOmitir = Me.Controls.Add("Forms.CheckBox.1", "chkOmitir")
    With chkOmitir
        .Caption = "Omitir siempre las zonas vacias (aunque esten seleccionadas)"
        .Left = 12: .Top = y: .Width = 320: .Height = 18
        .Value = True
    End With
    y = y + 22

    Set chkAbrir = Me.Controls.Add("Forms.CheckBox.1", "chkAbrir")
    With chkAbrir
        .Caption = "Abrir el PDF al terminar"
        .Left = 12: .Top = y: .Width = 150: .Height = 18
        .Value = True
    End With

    AgregarLabel "Copias (impresora):", 190, y + 3, 85, False
    Set txtCopias = Me.Controls.Add("Forms.TextBox.1", "txtCopias")
    With txtCopias
        .Left = 290: .Top = y: .Width = 40: .Height = 18
        .Text = "1"
    End With
    y = y + 34

    'Botones (3 x 100 de ancho, repartidos en el nuevo ancho)
    Set btnPDF = AgregarBoton("Exportar a PDF", 12, y)
    Set btnImprimir = AgregarBoton("Imprimir", 121, y)
    Set btnCancelar = AgregarBoton("Cancelar", 230, y)
    btnPDF.Default = True
    btnCancelar.Cancel = True
    btnCancelar.BackColor = RGB(128, 128, 128)
    Me.Height = y + 64
End Sub

Private Sub AgregarLabel(ByVal texto As String, ByVal x As Single, ByVal y As Single, _
                         ByVal w As Single, ByVal negrita As Boolean)
    Dim l As MSForms.Label
    Set l = Me.Controls.Add("Forms.Label.1")
    With l
        .Caption = texto
        .Left = x: .Top = y: .Width = w: .Height = 16
        .Font.Bold = negrita
    End With
End Sub

Private Function AgregarBoton(ByVal texto As String, ByVal x As Single, _
                              ByVal y As Single) As MSForms.CommandButton
    Const ANCHO As Single = 100
    Const ALTO As Single = 24
    Const GROSOR As Single = 1         'grosor del borde gris
    Dim borde As MSForms.Label

    'Borde gris (se crea primero para que quede detrás)
    Set borde = Me.Controls.Add("Forms.Label.1")
    With borde
        .Left = x - GROSOR: .Top = y - GROSOR
        .Width = ANCHO + GROSOR * 2: .Height = ALTO + GROSOR * 2
        .BackColor = RGB(160, 160, 160)
        .BackStyle = fmBackStyleOpaque
        .Caption = ""
    End With

    'Botón
    Set AgregarBoton = Me.Controls.Add("Forms.CommandButton.1")
    With AgregarBoton
        .Caption = texto
        .Left = x: .Top = y: .Width = ANCHO: .Height = ALTO
        .Font.Bold = True
        .BackColor = RGB(0, 101, 187)       'azul #0065BB
        .ForeColor = RGB(255, 255, 255)     'texto blanco
    End With
End Function

Private Sub cboTodas_Change()
    Dim v As Long
    If cboTodas.ListIndex < 0 Then Exit Sub
    For v = 1 To NUM_VISTAS
        cboVista(v).ListIndex = cboTodas.ListIndex
    Next v
End Sub

Private Sub btnPDF_Click()
    Ejecutar True
End Sub

Private Sub btnImprimir_Click()
    Ejecutar False
End Sub

Private Sub btnCancelar_Click()
    Unload Me
End Sub

Private Sub Ejecutar(ByVal aPDF As Boolean)
    Dim modos(1 To NUM_VISTAS) As Long
    Dim v As Long, areas As Collection, copias As Long

    For v = 1 To NUM_VISTAS
        modos(v) = cboVista(v).ListIndex
    Next v

    Set areas = AreasSeleccionadas(modos, chkOmitir.Value)
    If areas.Count = 0 Then
        MsgBox "No hay zonas para enviar." & vbCrLf & _
               "Revise la seleccion (las zonas vacias o inexistentes se omiten).", vbInformation
        Exit Sub
    End If

    copias = Val(txtCopias.Text)
    If copias < 1 Then copias = 1

    Me.Hide
    GenerarPaquete areas, aPDF, chkAbrir.Value, copias
    Unload Me
End Sub

