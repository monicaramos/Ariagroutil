VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmContabFactCV 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Contabilización de Facturas "
   ClientHeight    =   5730
   ClientLeft      =   45
   ClientTop       =   2430
   ClientWidth     =   6360
   Icon            =   "frmContabFactCV.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5730
   ScaleWidth      =   6360
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog cd1 
      Left            =   7680
      Top             =   5160
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame FrameCobros 
      Height          =   5760
      Left            =   0
      TabIndex        =   10
      Top             =   -30
      Width           =   6375
      Begin VB.ComboBox Combo1 
         Height          =   315
         Index           =   1
         ItemData        =   "frmContabFactCV.frx":000C
         Left            =   1620
         List            =   "frmContabFactCV.frx":0016
         Style           =   2  'Dropdown List
         TabIndex        =   26
         Tag             =   "Tipo Factura|N|N|0|5|factsocio|tipofact|||"
         Top             =   1020
         Width           =   1665
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   6
         Left            =   2730
         Locked          =   -1  'True
         TabIndex        =   24
         Text            =   "Text5"
         Top             =   2055
         Width           =   3135
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   6
         Left            =   1590
         MaxLength       =   10
         TabIndex        =   1
         Tag             =   "Código Propio|N|N|1|99|sbanco|codbanpr|00|S|"
         Top             =   2055
         Width           =   1050
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   7
         Left            =   1605
         MaxLength       =   10
         TabIndex        =   0
         Tag             =   "Código Postal|T|S|||clientes|codposta|||"
         Top             =   1560
         Width           =   1050
      End
      Begin VB.TextBox txtCodigo 
         Height          =   285
         Index           =   5
         Left            =   3720
         MaxLength       =   3
         TabIndex        =   3
         Text            =   "AAA"
         Top             =   2700
         Width           =   405
      End
      Begin VB.TextBox txtCodigo 
         Height          =   285
         Index           =   4
         Left            =   1605
         MaxLength       =   3
         TabIndex        =   2
         Text            =   "AAA"
         Top             =   2685
         Width           =   405
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   3
         Left            =   3735
         MaxLength       =   10
         TabIndex        =   7
         Tag             =   "Código Postal|T|S|||clientes|codposta|||"
         Top             =   3645
         Width           =   1050
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   2
         Left            =   1605
         MaxLength       =   10
         TabIndex        =   6
         Tag             =   "Código Postal|T|S|||clientes|codposta|||"
         Top             =   3615
         Width           =   1050
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Left            =   4755
         TabIndex        =   9
         Top             =   5205
         Width           =   975
      End
      Begin VB.CommandButton cmdAceptar 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   3570
         TabIndex        =   8
         Top             =   5205
         Width           =   975
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   0
         Left            =   1605
         MaxLength       =   7
         TabIndex        =   4
         Tag             =   "Nº de Factura|N|N|0|9999999|schfac|numfactu|0000000|S|"
         Top             =   3150
         Width           =   830
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   1
         Left            =   3720
         MaxLength       =   7
         TabIndex        =   5
         Tag             =   "Nº de Factura|N|N|0|9999999|schfac|numfactu|0000000|S|"
         Top             =   3150
         Width           =   830
      End
      Begin MSComctlLib.ProgressBar Pb1 
         Height          =   345
         Left            =   330
         TabIndex        =   20
         Top             =   4140
         Width           =   5370
         _ExtentX        =   9472
         _ExtentY        =   609
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.Label Label4 
         Caption         =   "Tipo de Factura"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00972E0B&
         Height          =   255
         Index           =   5
         Left            =   330
         TabIndex        =   27
         Top             =   1020
         Width           =   1815
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   6
         Left            =   1260
         MouseIcon       =   "frmContabFactCV.frx":002C
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar cuenta"
         Top             =   2055
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Cta. Banco"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00972E0B&
         Height          =   195
         Index           =   3
         Left            =   330
         TabIndex        =   25
         Top             =   2070
         Width           =   795
      End
      Begin VB.Label lblProgres 
         Height          =   285
         Index           =   1
         Left            =   330
         TabIndex        =   23
         Top             =   4815
         Width           =   5295
      End
      Begin VB.Label lblProgres 
         Height          =   285
         Index           =   0
         Left            =   360
         TabIndex        =   22
         Top             =   4500
         Width           =   5265
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   7
         Left            =   1290
         Picture         =   "frmContabFactCV.frx":017E
         ToolTipText     =   "Buscar fecha"
         Top             =   1560
         Width           =   240
      End
      Begin VB.Label Label4 
         Caption         =   "Fecha Vto"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00972E0B&
         Height          =   255
         Index           =   4
         Left            =   330
         TabIndex        =   21
         Top             =   1560
         Width           =   1815
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Letra de Serie"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00972E0B&
         Height          =   195
         Index           =   2
         Left            =   330
         TabIndex        =   19
         Top             =   2445
         Width           =   1005
      End
      Begin VB.Label Label4 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   1
         Left            =   2880
         TabIndex        =   18
         Top             =   2700
         Width           =   420
      End
      Begin VB.Label Label4 
         Caption         =   "Desde"
         Height          =   195
         Index           =   0
         Left            =   720
         TabIndex        =   17
         Top             =   2685
         Width           =   465
      End
      Begin VB.Label Label4 
         Caption         =   "Fecha Factura"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00972E0B&
         Height          =   255
         Index           =   16
         Left            =   330
         TabIndex        =   16
         Top             =   3375
         Width           =   1815
      End
      Begin VB.Label Label4 
         Caption         =   "Desde"
         Height          =   195
         Index           =   15
         Left            =   720
         TabIndex        =   15
         Top             =   3615
         Width           =   465
      End
      Begin VB.Label Label4 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   14
         Left            =   2880
         TabIndex        =   14
         Top             =   3645
         Width           =   420
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   2
         Left            =   1290
         Picture         =   "frmContabFactCV.frx":0209
         ToolTipText     =   "Buscar fecha"
         Top             =   3615
         Width           =   240
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   3
         Left            =   3420
         Picture         =   "frmContabFactCV.frx":0294
         ToolTipText     =   "Buscar fecha"
         Top             =   3645
         Width           =   240
      End
      Begin VB.Label Label4 
         Caption         =   "Desde"
         Height          =   195
         Index           =   13
         Left            =   720
         TabIndex        =   13
         Top             =   3150
         Width           =   465
      End
      Begin VB.Label Label4 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   12
         Left            =   2880
         TabIndex        =   12
         Top             =   3195
         Width           =   420
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Factura"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00972E0B&
         Height          =   195
         Index           =   11
         Left            =   330
         TabIndex        =   11
         Top             =   2910
         Width           =   555
      End
   End
End
Attribute VB_Name = "frmContabFactCV"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' +-+-+-+-+-+-+-+-+-+-+-
' +-+- Autor MANOLO +-+-
' +-+-+-+-+-+-+-+-+-+-+-

Option Explicit

Public NumCod As String 'Para indicar cod. Traspaso,Movimiento, etc. que llama
                        'Para indicar nº oferta a imprimir

Public CadTag As String 'Cadena con el Tag del campo que se va a poner en D/H en los listados
                        'Se necesita si el tipo de codigo es texto

    
Private Conexion As Byte
'1.- Conexión a BD Ariges  2.- Conexión a BD Conta

Private HaDevueltoDatos As Boolean

Private WithEvents frmCta As frmCtasConta 'Ctas contables
Attribute frmCta.VB_VarHelpID = -1
Private WithEvents frmC As frmCal 'calendario fechas
Attribute frmC.VB_VarHelpID = -1
Private WithEvents frmFPa As frmForpaConta 'Formas de Pago Contables
Attribute frmFPa.VB_VarHelpID = -1
Private WithEvents frmTIva As frmTipIVAConta 'Tipos de Iva Contables
Attribute frmTIva.VB_VarHelpID = -1

'GENERALES PARA PASARLE A CRYSTAL REPORT
Private cadFormula As String 'Cadena con la FormulaSelection para Crystal Report
Private cadParam As String 'Cadena con los parametros para Crystal Report
Private numParam As Byte 'Numero de parametros que se pasan a Crystal Report
Private cadSelect As String 'Cadena para comprobar si hay datos antes de abrir Informe
Private cadTitulo As String 'Titulo para la ventana frmImprimir
Private cadNombreRPT As String 'Nombre del informe

Dim indCodigo As Integer 'indice para txtCodigo
Dim indFrame As Single 'nº de frame en el que estamos
 
'Se inicializan para cada Informe (tabla de BD a la que hace referencia
Dim tabla As String
Dim Codigo As String 'Código para FormulaSelection de Crystal Report
Dim TipCod As String
Dim Orden1 As String 'Campo de Ordenacion (por codigo) para Cristal Report
Dim Orden2 As String 'Campo de Ordenacion (por nombre) para Cristal Report

Dim PrimeraVez As Boolean

Private Sub KEYpress(KeyAscii As Integer)
    If KeyAscii = 13 Then 'ENTER
        KeyAscii = 0
        SendKeys "{tab}"
    ElseIf KeyAscii = 27 Then Unload Me  'ESC
    End If
End Sub

Private Sub cmdAceptar_Click()
Dim cDesde As String, cHasta As String 'cadena codigo Desde/Hasta
Dim nDesde As String, nHasta As String 'cadena Descripcion Desde/Hasta
Dim cadTABLA As String, cOrden As String
Dim cadMen As String
Dim i As Byte
Dim SQL As String
Dim tipo As Byte
Dim Nregs As Long
Dim NumError As Long

    If Not DatosOk Then Exit Sub
    
    cadSelect = tabla & ".intconta=0 "
    
    'D/H Fecha factura
    cDesde = Trim(txtCodigo(2).Text)
    cHasta = Trim(txtCodigo(3).Text)
    If Not (cDesde = "" And cHasta = "") Then
        'Cadena para seleccion Desde y Hasta
        Codigo = "{" & tabla & ".fecfactu}"
        TipCod = "F"
        If Not PonerDesdeHasta(cDesde, cHasta, "", "", "pDHfechaFac= """) Then Exit Sub
    End If
    
    'D/H letra de serie
    cDesde = Trim(txtCodigo(4).Text)
    cHasta = Trim(txtCodigo(5).Text)
    If Not (cDesde = "" And cHasta = "") Then
        'Cadena para seleccion Desde y Hasta
        Codigo = "{" & tabla & ".letraser}"
        TipCod = "T"
        If Not PonerDesdeHasta(cDesde, cHasta, "", "", "pDHColec= """) Then Exit Sub
    End If
    
    'D/H numero de factura
    cDesde = Trim(txtCodigo(0).Text)
    cHasta = Trim(txtCodigo(1).Text)
    If Not (cDesde = "" And cHasta = "") Then
        'Cadena para seleccion Desde y Hasta
        Codigo = "{" & tabla & ".numfactu}"
        TipCod = "N"
        If Not PonerDesdeHasta(cDesde, cHasta, "", "", "pDHColec= """) Then Exit Sub
    End If
    
    If Not AnyadirAFormula(cadSelect, "{cvfacturas.tipofactu} = " & Combo1(1).ListIndex) Then Exit Sub
    
    If Not HayRegParaInforme(tabla, cadSelect) Then Exit Sub
    
    If HayFacturasIncorrectas(tabla, cadSelect) Then Exit Sub
    
    
    ContabilizarFacturas tabla, cadSelect
     'Eliminar la tabla TMP
    BorrarTMPFacturas
    'Desbloqueamos ya no estamos contabilizando facturas
    DesBloqueoManual ("FCVCON") 'Facturas Coarval CONtabilizar
    
eError:
    If Err.Number <> 0 Or NumError <> 0 Then
        MuestraError Err.Number, "No se ha realizado el proceso de contabilización. Llame a soporte."
    End If
    
    Pb1.visible = False
    lblProgres(0).Caption = ""
    lblProgres(1).Caption = ""
    cmdCancel_Click
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub Combo1_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub


Private Sub Combo1_Click(Index As Integer)
    Select Case Combo1(1).ListIndex
        Case 0
            txtCodigo(6).Text = vParamAplic.CtaBancoCVV
        Case 1, 2
            txtCodigo(6).Text = vParamAplic.CtaBancoCV
    End Select
    txtCodigo_LostFocus (6)
End Sub

Private Sub Combo1_LostFocus(Index As Integer)
    Select Case Combo1(1).ListIndex
        Case 0
            txtCodigo(6).Text = vParamAplic.CtaBancoCVV
        Case 1, 2
            txtCodigo(6).Text = vParamAplic.CtaBancoCV
    End Select
    txtCodigo_LostFocus (6)

End Sub

Private Sub Form_Activate()
    If PrimeraVez Then
        PrimeraVez = False
        ValoresPorDefecto
        PonerFoco txtCodigo(7)
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
Dim h As Integer, w As Integer
Dim List As Collection

    PrimeraVez = True
    Limpiar Me

    'IMAGES para busqueda
     Me.imgBuscar(6).Picture = frmPpal.imgListImages16.ListImages(1).Picture
     
    '###Descomentar
'    CommitConexion
         
    CargaCombo
         
    FrameCobrosVisible True, h, w
    indFrame = 5
    tabla = "cvfacturas"
    
    Pb1.visible = False
    
    'Esto se consigue poneinedo el cancel en el opcion k corresponda
    Me.cmdCancel.Cancel = True
    Me.Width = w
    Me.Height = h
    
'14/02/2007 lo dejo donde estaba
'
'     '07022007   lo he quitado de contbilizar facturas y lo he puesto aqui para acotar registros
'     'comprobar que se han rellenado los dos campos de fecha
'     'sino rellenar con fechaini o fechafin del ejercicio
'     'que guardamos en vbles Orden1,Orden2
'
'     Orden1 = ""
'     Orden1 = DevuelveDesdeBDNew(cConta, "parametros", "fechaini", "", "", "", "", "", "", "", "", "", "")
'
'     Orden2 = ""
'     Orden2 = DevuelveDesdeBDNew(cConta, "parametros", "fechafin", "", "", "", "", "", "", "", "", "", "")
     
'     If txtCodigo(2).Text = "" Then
'        txtCodigo(2).Text = Orden1 'fechaini del ejercicio de la conta
'     End If
'
'     If txtCodigo(3).Text = "" Then
'        txtCodigo(3).Text = Orden2 'fecha fin del ejercicio de la conta
'     End If
End Sub

Private Sub frmC_Selec(vFecha As Date)
 'Fecha
    txtCodigo(CByte(imgFec(2).Tag)).Text = Format(vFecha, "dd/MM/yyyy")
End Sub

Private Sub frmCta_DatoSeleccionado(CadenaSeleccion As String)
'Form de Consulta de Cuentas contables
    txtCodigo(indCodigo).Text = RecuperaValor(CadenaSeleccion, 1) 'cta contable
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2) 'nombre ctacontable
End Sub


Private Sub imgFec_Click(Index As Integer)
'FEchas
    Dim esq, dalt As Long
    Dim obj As Object
    
    Set frmC = New frmCal

    esq = imgFec(Index).Left
    dalt = imgFec(Index).Top

    Set obj = imgFec(Index).Container

    While imgFec(Index).Parent.Name <> obj.Name
        esq = esq + obj.Left
        dalt = dalt + obj.Top
        Set obj = obj.Container
    Wend
       
    ' es desplega dalt i cap a la esquerra
    frmC.Left = esq + imgFec(Index).Parent.Left + 30
    frmC.Top = dalt + imgFec(Index).Parent.Top + imgFec(Index).Height + 420 + 30

    ' ***canviar l'index de imgFec pel 1r index de les imagens de buscar data***
    imgFec(2).Tag = Index 'independentment de les dates que tinga, sempre pose l'index en la 27
    If txtCodigo(Index).Text <> "" Then frmC.NovaData = txtCodigo(Index).Text

    frmC.Show vbModal
    Set frmC = Nothing
    PonerFoco txtCodigo(CByte(imgFec(2).Tag) + 2)
    ' ***************************
End Sub

Private Sub ListView1_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub imgBuscar_Click(Index As Integer)
   Select Case Index
        
        Case 6 ' ctas contables ventas y banco
            AbrirFrmCtasConta (Index)
        
        
    End Select
    PonerFoco txtCodigo(indCodigo)
End Sub

Private Sub txtCodigo_GotFocus(Index As Integer)
    ConseguirFoco txtCodigo(Index), 3
End Sub

Private Sub txtCodigo_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    KEYdown KeyCode
End Sub

Private Sub txtCodigo_KeyPress(Index As Integer, KeyAscii As Integer)
'14/02/2007 antes estaba esto
'    KEYpress KeyAscii
' ahora
    If KeyAscii = teclaBuscar Then
        Select Case Index
            Case 6: KEYBusqueda KeyAscii, 6 'cta contable
            Case 8: KEYBusqueda KeyAscii, 8 'cta contable
            Case 9: KEYBusqueda KeyAscii, 9 'forma de pago
            Case 10: KEYBusqueda KeyAscii, 10 'tipo de iva
            Case 2: KEYFecha KeyAscii, 2 'fecha desde
            Case 3: KEYFecha KeyAscii, 3 'fecha hasta
            Case 7: KEYFecha KeyAscii, 7 'fecha de vencimiento
        End Select
    Else
        KEYpress KeyAscii
    End If

End Sub

Private Sub KEYBusqueda(KeyAscii As Integer, indice As Integer)
    KeyAscii = 0
    imgBuscar_Click (indice)
End Sub

Private Sub KEYFecha(KeyAscii As Integer, indice As Integer)
    KeyAscii = 0
    imgFec_Click (indice)
End Sub

Private Sub txtCodigo_LostFocus(Index As Integer)
Dim cad As String, cadTipo As String 'tipo cliente

    'Quitar espacios en blanco por los lados
    txtCodigo(Index).Text = Trim(txtCodigo(Index).Text)
    
    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub

    
    Select Case Index
        Case 6, 8 ' CTAS CONTABLES
            If txtCodigo(Index).Text = "" Then Exit Sub
            If Combo1(1).ListIndex = 0 Then
                txtNombre(Index).Text = PonerNombreCuenta(txtCodigo(Index), 1, , cContaCVV)
            Else
                txtNombre(Index).Text = PonerNombreCuenta(txtCodigo(Index), 1, , cContaCV)
            End If
            
        Case 2, 3, 7  'FECHAS
            If txtCodigo(Index).Text <> "" Then PonerFormatoFecha txtCodigo(Index)
        
        Case 0, 1 ' NUMERO DE FACTURA
            If txtCodigo(Index).Text <> "" Then PonerFormatoEntero txtCodigo(Index)
        
        Case 4, 5 ' LETRA DE SERIE
            If txtCodigo(Index).Text <> "" Then txtCodigo(Index).Text = UCase(txtCodigo(Index).Text)
        
    End Select
End Sub

Private Sub FrameCobrosVisible(visible As Boolean, ByRef h As Integer, ByRef w As Integer)
    Me.FrameCobros.visible = visible
    If visible = True Then
        Me.FrameCobros.Top = -90
        Me.FrameCobros.Left = 0
        Me.FrameCobros.Height = 6120
        Me.FrameCobros.Width = 6450
        w = Me.FrameCobros.Width
        h = Me.FrameCobros.Height
    End If
End Sub

Private Sub ValoresPorDefecto()
    Combo1(1).ListIndex = 0
    txtCodigo(7).Text = Format(Now, "dd/mm/yyyy")
    txtCodigo(6).Text = vParamAplic.CtaBancoCVV
    If txtCodigo(6).Text <> "" Then txtNombre(6).Text = PonerNombreCuenta(txtCodigo(6), 1, , cContaCVV)
End Sub

Private Sub InicializarVbles()
    cadFormula = ""
    cadSelect = ""
    cadParam = ""
    numParam = 0
End Sub

Private Function PonerDesdeHasta(codD As String, codH As String, nomD As String, nomH As String, param As String) As Boolean
'IN: codD,codH --> codigo Desde/Hasta
'    nomD,nomH --> Descripcion Desde/Hasta
'Añade a cadFormula y cadSelect la cadena de seleccion:
'       "(codigo>=codD AND codigo<=codH)"
' y añade a cadParam la cadena para mostrar en la cabecera informe:
'       "codigo: Desde codD-nomd Hasta: codH-nomH"
Dim devuelve As String
Dim devuelve2 As String

    PonerDesdeHasta = False
    devuelve = CadenaDesdeHasta(codD, codH, Codigo, TipCod)
    If devuelve = "Error" Then Exit Function
    If Not AnyadirAFormula(cadFormula, devuelve) Then Exit Function
    If TipCod <> "F" Then 'Fecha
        If Not AnyadirAFormula(cadSelect, devuelve) Then Exit Function
    Else
        devuelve2 = CadenaDesdeHastaBD(codD, codH, Codigo, TipCod)
        If devuelve2 = "Error" Then Exit Function
        If Not AnyadirAFormula(cadSelect, devuelve2) Then Exit Function
    End If
    If devuelve <> "" Then
        If param <> "" Then
            'Parametro Desde/Hasta
            cadParam = cadParam & AnyadirParametroDH(param, codD, codH, nomD, nomH)
            numParam = numParam + 1
        End If
        PonerDesdeHasta = True
    End If
End Function

Private Sub AbrirFrmCtasConta(indice As Integer)
    indCodigo = indice
    Set frmCta = New frmCtasConta
    frmCta.DatosADevolverBusqueda = "0|1|"
    frmCta.CodigoActual = txtCodigo(indCodigo)
    Select Case Combo1(1).ListIndex
        Case 0
            frmCta.Conexion = cContaCVV
        Case 1, 2
            frmCta.Conexion = cContaCV
    End Select
    frmCta.Show vbModal
    Set frmCta = Nothing
End Sub


Private Function DatosOk() As Boolean
Dim b As Boolean
Dim cadG As String
    b = True

    If txtCodigo(7).Text = "" And b Then
        MsgBox "Debe introducir obligatoriamente una Fecha de Vencimiento.", vbExclamation
        b = False
        PonerFoco txtCodigo(7)
    End If
    
    If txtCodigo(6).Text = "" And b Then
        MsgBox "Debe introducir obligatoriamente una Cta Contable de Banco.", vbExclamation
        b = False
        PonerFoco txtCodigo(6)
    End If
    
    Select Case Combo1(1).ListIndex
        Case 0
            FechasEjercicioContaCVV FIniCVV, FFinCVV
            Orden1 = FIniCVV
            Orden2 = FFinCVV
        Case 1, 2
            FechasEjercicioContaCV FIniCV, FFinCV
            Orden1 = FIniCV
            Orden2 = FFinCV
    End Select
    
    '07022007 he añadido esto tambien aquí
     If txtCodigo(2).Text = "" Then
        txtCodigo(2).Text = Orden1 'fechaini del ejercicio de la conta
     End If
     
     If txtCodigo(3).Text = "" Then
        txtCodigo(3).Text = Format(Day(CDate(Orden2)), "00") & "/" & Format(Month(CDate(Orden2)), "00") & "/" & Format(Year(CDate(Orden2)) + 1, "0000") 'fecha fin del ejercicio de la conta
     End If


    DatosOk = b
End Function

' copiado del ariges
Private Sub ContabilizarFacturas(cadTABLA As String, cadwhere As String)
'Contabiliza Facturas de Clientes o de Proveedores
Dim SQL As String
Dim b As Boolean
Dim tmpErrores As Boolean 'Indica si se creo correctamente la tabla de errores
Dim CCoste As String
Dim cadwhere1 As String

    SQL = "FCVCON" 'contabilizar facturas de coarval

    'Bloquear para que nadie mas pueda contabilizar
    DesBloqueoManual (SQL)
    If Not BloqueoManual(SQL, "1") Then
        MsgBox "No se pueden Contabilizar Facturas. Hay otro usuario contabilizando.", vbExclamation
        Screen.MousePointer = vbDefault
        Exit Sub
    End If

    'comprobar si existen en Ariagroutil (coarval) facturas anteriores al periodo solicitado
    'sin contabilizar
    If Me.txtCodigo(2).Text <> "" Then
        SQL = "SELECT COUNT(*) FROM " & cadTABLA
        SQL = SQL & " WHERE fecfactu <"
        SQL = SQL & DBSet(txtCodigo(2), "F") & " AND intconta=0 and tipofactu = " & Combo1(1).ListIndex
        
        If RegistrosAListar(SQL) > 0 Then
            MsgBox "Hay Facturas anteriores sin contabilizar.", vbExclamation
            Exit Sub
        End If
    End If
    
    'Visualizar la barra de Progreso
    Me.Pb1.visible = True
'    Me.Pb1.Top = 3350
    
    
    '==========================================================
    'REALIZAR COMPROBACIONES ANTES DE CONTABILIZAR FACTURAS
    '==========================================================
    
    Me.lblProgres(0).Caption = "Comprobaciones: "
    CargarProgres Me.Pb1, 100
        
    BorrarTMPFacturas
    'Cargar tabla TEMP con las Facturas que vamos a Trabajar
    b = CrearTMPFacturasCV(cadTABLA, cadwhere, False, False)
    If Not b Then Exit Sub
            
    BorrarTMPErrComprob
    ' nuevo
    b = CrearTMPErrComprob()
    If Not b Then Exit Sub
    
    'comprobar que todas las LETRAS SERIE existen en la contabilidad y en Ariagroutil
    '-----------------------------------------------------------------------------
    IncrementarProgres Me.Pb1, 10
    Me.lblProgres(1).Caption = "Comprobando letras de serie ..."
    Select Case Combo1(1).ListIndex
        Case 0
            b = ComprobarLetraSerieCV(0)
        Case 1
            b = ComprobarLetraSerieCV(1)
        Case 2 ' son las facturas de compra
            b = True
    End Select
    IncrementarProgres Me.Pb1, 30
    Me.Refresh
    If Not b Then
        frmMensaje.OpcionMensaje = 1
        frmMensaje.Show vbModal
        Exit Sub
    End If
    
    'comprobar que todos el TIPO IVA
    'existe en la Conta: vparamaplic.codivagas IN (conta.tiposiva.codigiva)
    '--------------------------------------------------------------------------
    Me.lblProgres(1).Caption = "Comprobando Tipos de IVA en contabilidad ..."
    b = ComprobarTiposIVACV(Combo1(1).ListIndex)
    IncrementarProgres Me.Pb1, 10
    Me.Refresh
    If Not b Then
        frmMensaje.OpcionMensaje = 3
        frmMensaje.Show vbModal
        Exit Sub
    End If
    
    'comprobar que no haya Nº FACTURAS en la contabilidad para esa fecha
    'que ya existan
    '-----------------------------------------------------------------------
    If Combo1(1).ListIndex <= 1 Then
        Me.lblProgres(1).Caption = "Comprobando Nº Facturas en contabilidad ..."
        SQL = "anofaccl>=" & Year(txtCodigo(2).Text) & " AND anofaccl<= " & Year(txtCodigo(3).Text)
        If Combo1(1).ListIndex = 0 Then
            b = ComprobarNumFacturas(cContaCVV, SQL)
        Else
            b = ComprobarNumFacturas(cContaCV, SQL)
        End If
    Else
        b = True
    End If
    
    IncrementarProgres Me.Pb1, 20
    Me.Refresh
    If Not b Then
        frmMensaje.OpcionMensaje = 1
        frmMensaje.Show vbModal
        Exit Sub
    End If

    'comprobar que todas las CUENTAS de los distintos clientes que vamos a
    'contabilizar existen en la Conta: cvfacturas.codmactasoc IN (conta.cuentas.codmacta)
    '-----------------------------------------------------------------------------
    Me.lblProgres(1).Caption = "Comprobando Cuentas Contables en contabilidad ..."
    Select Case Combo1(1).ListIndex
        Case 0
            b = ComprobarCtaContable(cadTABLA, 12, , cContaCVV, 0)
        Case 1, 2
            b = ComprobarCtaContable(cadTABLA, 12, , cContaCV, Combo1(1).ListIndex)
    End Select
    IncrementarProgres Me.Pb1, 10
    Me.Refresh
    If Not b Then
        frmMensaje.OpcionMensaje = 2
        frmMensaje.Show vbModal
        Exit Sub
    End If
    
    
    'comprobar que todas las CUENTAS de los distintos clientes que vamos a
    'contabilizar existen en la Conta: cvfacturas.codmactavta IN (conta.cuentas.codmacta)
    '-----------------------------------------------------------------------------
    Me.lblProgres(1).Caption = "Comprobando Cuentas Contables en contabilidad ..."
    Select Case Combo1(1).ListIndex
        Case 0
            b = ComprobarCtaContable(cadTABLA, 13, , cContaCVV, 0)
        Case 1, 2
            b = ComprobarCtaContable(cadTABLA, 13, , cContaCV, Combo1(1).ListIndex)
    End Select
    IncrementarProgres Me.Pb1, 20
    Me.Refresh
    If Not b Then
        frmMensaje.OpcionMensaje = 2
        frmMensaje.Show vbModal
        Exit Sub
    End If
    
    
    
    '===========================================================================
    'CONTABILIZAR FACTURAS
    '===========================================================================
    Me.lblProgres(0).Caption = "Contabilizar Facturas: "
    CargarProgres Me.Pb1, 10
    Me.lblProgres(1).Caption = "Insertando Facturas en Contabilidad..."
       
    
    'Crear tabla TEMP para los posible errores de facturas
    tmpErrores = CrearTMPErrFact(cadTABLA)
    
    
    b = PasarFacturasAContab(cadTABLA, txtCodigo(7).Text, txtCodigo(6).Text, CCoste)
    
    If Not b Then
        If tmpErrores Then
            'Cargar un listview con la tabla TEMP de Errores y mostrar
            'las facturas que fallaron
            frmMensaje.OpcionMensaje = 10
            frmMensaje.Show vbModal
        Else
            MsgBox "No pueden mostrarse los errores.", vbInformation
        End If
    Else
        MsgBox "El proceso ha finalizado correctamente.", vbInformation
    End If
    
    'Eliminar tabla TEMP de Errores
    BorrarTMPErrFact
    BorrarTMPErrComprob
End Sub


Private Function PasarFacturasAContab(cadTABLA As String, FecVenci As String, Banpr As String, CCoste As String) As Boolean
Dim SQL As String
Dim RS As ADODB.Recordset
Dim b As Boolean
Dim i As Integer
Dim numfactu As Integer
Dim codigo1 As String

    On Error GoTo EPasarFac

    PasarFacturasAContab = False
    
    'Total de Facturas a Insertar en la contabilidad
    SQL = "SELECT count(*) "
    SQL = SQL & " FROM " & cadTABLA & " INNER JOIN tmpfactu "
    codigo1 = "numserie"
    SQL = SQL & " ON " & cadTABLA & ".letraser=tmpfactu." & codigo1
    SQL = SQL & " AND " & cadTABLA & ".numfactu=tmpfactu.numfactu AND " & cadTABLA & ".fecfactu=tmpfactu.fecfactu "
    
    
    Set RS = New ADODB.Recordset
    RS.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not RS.EOF Then
        numfactu = RS.Fields(0)
    Else
        numfactu = 0
    End If
    RS.Close
    Set RS = Nothing

    If numfactu > 0 Then
        CargarProgres Me.Pb1, numfactu
        
        SQL = "SELECT * "
        SQL = SQL & " FROM tmpfactu "
        SQL = SQL & " order by fecfactu, numfactu"
            
        Set RS = New ADODB.Recordset
        RS.Open SQL, conn, adOpenStatic, adLockPessimistic, adCmdText
        i = 1

        b = True
        'contabilizar cada una de las facturas seleccionadas
        While Not RS.EOF And b
            SQL = cadTABLA & ".letraser=" & DBSet(Trim(RS.Fields(0)), "T") & " and numfactu=" & DBSet(RS!numfactu, "T")
            SQL = SQL & " and fecfactu=" & DBSet(RS!fecfactu, "F")
            If PasarFacturaCV(SQL, FecVenci, Combo1(1).ListIndex, txtCodigo(6).Text, CCoste) = False And b Then b = False
            
            IncrementarProgres Me.Pb1, 1
            Me.lblProgres(1).Caption = "Insertando Facturas en Contabilidad...   (" & i & " de " & numfactu & ")"
            Me.Refresh
            i = i + 1
            RS.MoveNext
        Wend
        RS.Close
        Set RS = Nothing
    End If
    
EPasarFac:
    If Err.Number <> 0 Then b = False
    
    If b Then
        PasarFacturasAContab = True
    Else
        PasarFacturasAContab = False
    End If
End Function


Public Function HayFacturasIncorrectas(cTabla As String, cWhere As String) As Boolean
'Comprobar si hay registros a Mostrar antes de abrir el Informe
Dim SQL As String
Dim RS As ADODB.Recordset
Dim PorIva As Currency

Dim BaseT As Currency
Dim IvaT As Currency
Dim TotalT As Currency

Dim IvaCal As Currency
Dim TotalCal As Currency

Dim cad As String
Dim cad1 As String

    On Error GoTo eHayFacturasIncorrectas
    
    HayFacturasIncorrectas = True
    

    If cWhere <> "" Then
        cWhere = QuitarCaracterACadena(cWhere, "{")
        cWhere = QuitarCaracterACadena(cWhere, "}")
        cWhere = QuitarCaracterACadena(cWhere, "_1")
        SQL = SQL & " WHERE " & cWhere
    End If
    
    SQL = " SELECT letraser,numfactu,fecfactu,codmactasoc, year(fecfactu) as anofaccl"
    SQL = SQL & " FROM cvfacturas "
    SQL = SQL & " WHERE " & cWhere
    SQL = SQL & " and erronea = 1 "
    SQL = SQL & " order by letraser, numfactu, fecfactu"
    
    Set RS = New ADODB.Recordset
    RS.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    cad = ""
    
    While Not RS.EOF
        cad = cad & DBLet(RS!numfactu, "N") & ", "
        
        RS.MoveNext
    Wend
    
    Set RS = Nothing
    
    If cad <> "" Then
        cad1 = "Las siguientes facturas tienen importes incorrectos. Revise." & vbCrLf & vbCrLf & Mid(cad, 1, Len(cad) - 2)
        MsgBox cad1, vbExclamation
        HayFacturasIncorrectas = True
    Else
        HayFacturasIncorrectas = False
    End If
    
eHayFacturasIncorrectas:
    If Err.Number <> 0 Then
        MuestraError Err.Number, "Comprobando Facturas Incorrectas", Err.Description
    End If
End Function



Private Sub CargaCombo()
Dim Ini As Integer
Dim Fin As Integer
Dim i As Integer

    ' *** neteje els combos, els pose valor i seleccione el valor per defecte ***
    'tipo de factura
    Combo1(1).Clear
    
    Combo1(1).AddItem "Varias"
    Combo1(1).ItemData(Combo1(1).NewIndex) = 0
    Combo1(1).AddItem "Ventas Tienda"
    Combo1(1).ItemData(Combo1(1).NewIndex) = 1
    Combo1(1).AddItem "Compras Tienda"
    Combo1(1).ItemData(Combo1(1).NewIndex) = 2

End Sub


