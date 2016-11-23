VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmContaFacSoc 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Integración Contable de Facturas de Socios"
   ClientHeight    =   6750
   ClientLeft      =   45
   ClientTop       =   2430
   ClientWidth     =   6585
   Icon            =   "frmContaFacSoc.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6750
   ScaleWidth      =   6585
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
      Height          =   6450
      Left            =   150
      TabIndex        =   11
      Top             =   180
      Width           =   6330
      Begin VB.Frame Frame1 
         Caption         =   "Datos para Selección"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   2490
         Left            =   90
         TabIndex        =   12
         Top             =   225
         Width           =   6090
         Begin VB.TextBox txtCodigo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   9
            Left            =   1575
            MaxLength       =   10
            TabIndex        =   5
            Tag             =   "Cta.Contable|T|S|||sparam|ctaconta|||"
            Top             =   2070
            Width           =   1080
         End
         Begin VB.TextBox txtNombre 
            BackColor       =   &H80000018&
            Enabled         =   0   'False
            Height          =   285
            Index           =   9
            Left            =   2700
            Locked          =   -1  'True
            TabIndex        =   31
            Top             =   2070
            Width           =   3180
         End
         Begin VB.TextBox txtCodigo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   0
            Left            =   1575
            MaxLength       =   10
            TabIndex        =   4
            Tag             =   "Cta.Contable|T|S|||sparam|ctaconta|||"
            Top             =   1710
            Width           =   1080
         End
         Begin VB.TextBox txtNombre 
            BackColor       =   &H80000018&
            Enabled         =   0   'False
            Height          =   285
            Index           =   0
            Left            =   2700
            Locked          =   -1  'True
            TabIndex        =   28
            Top             =   1710
            Width           =   3180
         End
         Begin VB.TextBox txtCodigo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   8
            Left            =   3705
            MaxLength       =   7
            TabIndex        =   1
            Tag             =   "Nº de Factura|N|N|0|9999999|schfac|numfactu|0000000|S|"
            Top             =   690
            Width           =   830
         End
         Begin VB.TextBox txtCodigo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   7
            Left            =   1590
            MaxLength       =   7
            TabIndex        =   0
            Tag             =   "Nº de Factura|N|N|0|9999999|schfac|numfactu|0000000|S|"
            Top             =   690
            Width           =   830
         End
         Begin VB.TextBox txtCodigo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   6
            Left            =   3690
            MaxLength       =   10
            TabIndex        =   3
            Tag             =   "Código Postal|T|S|||clientes|codposta|||"
            Top             =   1170
            Width           =   1050
         End
         Begin VB.TextBox txtCodigo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   5
            Left            =   1575
            MaxLength       =   10
            TabIndex        =   2
            Tag             =   "Código Postal|T|S|||clientes|codposta|||"
            Top             =   1170
            Width           =   1050
         End
         Begin VB.Label Label4 
            Caption         =   "Hasta"
            Height          =   195
            Index           =   4
            Left            =   675
            TabIndex        =   32
            Top             =   2070
            Width           =   465
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   1
            Left            =   1260
            ToolTipText     =   "Buscar Cuenta Contable"
            Top             =   2070
            Width           =   240
         End
         Begin VB.Label Label4 
            Caption         =   "Desde"
            Height          =   195
            Index           =   3
            Left            =   675
            TabIndex        =   30
            Top             =   1710
            Width           =   465
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   0
            Left            =   1260
            ToolTipText     =   "Buscar Cuenta Contable"
            Top             =   1710
            Width           =   240
         End
         Begin VB.Label Label1 
            Caption         =   "Cuenta Contable"
            ForeColor       =   &H00972E0B&
            Height          =   195
            Index           =   1
            Left            =   315
            TabIndex        =   29
            Top             =   1440
            Width           =   1395
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
            Left            =   345
            TabIndex        =   27
            Top             =   450
            Width           =   555
         End
         Begin VB.Label Label4 
            Caption         =   "Hasta"
            Height          =   195
            Index           =   12
            Left            =   2865
            TabIndex        =   26
            Top             =   735
            Width           =   420
         End
         Begin VB.Label Label4 
            Caption         =   "Desde"
            Height          =   195
            Index           =   13
            Left            =   705
            TabIndex        =   25
            Top             =   690
            Width           =   465
         End
         Begin VB.Image imgFec 
            Height          =   240
            Index           =   6
            Left            =   3405
            Picture         =   "frmContaFacSoc.frx":000C
            ToolTipText     =   "Buscar fecha"
            Top             =   1185
            Width           =   240
         End
         Begin VB.Image imgFec 
            Height          =   240
            Index           =   5
            Left            =   1275
            Picture         =   "frmContaFacSoc.frx":0097
            ToolTipText     =   "Buscar fecha"
            Top             =   1155
            Width           =   240
         End
         Begin VB.Label Label4 
            Caption         =   "Hasta"
            Height          =   195
            Index           =   14
            Left            =   2865
            TabIndex        =   24
            Top             =   1185
            Width           =   420
         End
         Begin VB.Label Label4 
            Caption         =   "Desde"
            Height          =   195
            Index           =   15
            Left            =   705
            TabIndex        =   23
            Top             =   1155
            Width           =   465
         End
         Begin VB.Label Label4 
            Caption         =   "Fecha Factura"
            ForeColor       =   &H00972E0B&
            Height          =   255
            Index           =   2
            Left            =   315
            TabIndex        =   22
            Top             =   915
            Width           =   1815
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Datos para la contabilización"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   1755
         Left            =   90
         TabIndex        =   13
         Top             =   2790
         Width           =   6075
         Begin VB.TextBox txtCodigo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   3
            Left            =   1980
            MaxLength       =   10
            TabIndex        =   8
            Tag             =   "Cta.Contable|T|S|||sparam|ctaconta|||"
            Top             =   1170
            Width           =   1125
         End
         Begin VB.TextBox txtNombre 
            BackColor       =   &H80000018&
            Enabled         =   0   'False
            Height          =   285
            Index           =   3
            Left            =   3195
            Locked          =   -1  'True
            TabIndex        =   20
            Top             =   1170
            Width           =   2685
         End
         Begin VB.TextBox txtCodigo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   1
            Left            =   1980
            MaxLength       =   10
            TabIndex        =   7
            Tag             =   "Código Postal|T|S|||clientes|codposta|||"
            Top             =   810
            Width           =   1080
         End
         Begin VB.TextBox txtNombre 
            BackColor       =   &H80000018&
            Enabled         =   0   'False
            Height          =   285
            Index           =   4
            Left            =   3195
            Locked          =   -1  'True
            TabIndex        =   14
            Top             =   450
            Width           =   2685
         End
         Begin VB.TextBox txtCodigo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   4
            Left            =   1980
            MaxLength       =   10
            TabIndex        =   6
            Tag             =   "Cta.Contable|T|S|||sparam|ctaconta|||"
            Top             =   450
            Width           =   1125
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   3
            Left            =   1710
            ToolTipText     =   "Buscar Forma Pago"
            Top             =   1170
            Width           =   240
         End
         Begin VB.Label Label1 
            Caption         =   "Forma de Pago"
            ForeColor       =   &H00972E0B&
            Height          =   195
            Index           =   0
            Left            =   180
            TabIndex        =   21
            Top             =   1215
            Width           =   1395
         End
         Begin VB.Label Label4 
            Caption         =   "Fecha Vencimiento"
            ForeColor       =   &H00972E0B&
            Height          =   255
            Index           =   0
            Left            =   180
            TabIndex        =   19
            Top             =   855
            Width           =   1425
         End
         Begin VB.Image imgFec 
            Height          =   240
            Index           =   1
            Left            =   1710
            Picture         =   "frmContaFacSoc.frx":0122
            ToolTipText     =   "Buscar fecha"
            Top             =   810
            Width           =   240
         End
         Begin VB.Label Label1 
            Caption         =   "Cta.Banco Prevista"
            ForeColor       =   &H00972E0B&
            Height          =   195
            Index           =   24
            Left            =   180
            TabIndex        =   15
            Top             =   495
            Width           =   1395
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   4
            Left            =   1710
            ToolTipText     =   "Buscar Cuenta Contable"
            Top             =   450
            Width           =   240
         End
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Left            =   5040
         TabIndex        =   10
         Top             =   5940
         Width           =   975
      End
      Begin VB.CommandButton cmdAceptar 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   3855
         TabIndex        =   9
         Top             =   5940
         Width           =   975
      End
      Begin MSComctlLib.ProgressBar Pb1 
         Height          =   345
         Left            =   135
         TabIndex        =   16
         Top             =   4725
         Width           =   6045
         _ExtentX        =   10663
         _ExtentY        =   609
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.TextBox txtCodigo 
         Height          =   285
         Index           =   2
         Left            =   2070
         MaxLength       =   30
         TabIndex        =   33
         Top             =   3960
         Width           =   3870
      End
      Begin VB.Label lblProgres 
         Height          =   285
         Index           =   0
         Left            =   90
         TabIndex        =   18
         Top             =   5220
         Width           =   5940
      End
      Begin VB.Label lblProgres 
         Height          =   285
         Index           =   1
         Left            =   90
         TabIndex        =   17
         Top             =   5535
         Width           =   5925
      End
      Begin VB.Label Label4 
         Caption         =   "Concepto "
         ForeColor       =   &H00972E0B&
         Height          =   255
         Index           =   1
         Left            =   270
         TabIndex        =   34
         Top             =   4005
         Width           =   1050
      End
   End
End
Attribute VB_Name = "frmContaFacSoc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' +-+-+-+-+-+-+-+-+-+-+-
' +-+- Autor MONICA +-+-
' +-+-+-+-+-+-+-+-+-+-+-

Option Explicit

Public NumCod As String 'Para indicar cod. Traspaso,Movimiento, etc. que llama
                        'Para indicar nº oferta a imprimir

Public CadTag As String 'Cadena con el Tag del campo que se va a poner en D/H en los listados
                        'Se necesita si el tipo de codigo es texto

    
Private Conexion As Byte
'1.- Conexión a BD Ariges  2.- Conexión a BD Conta

Private HaDevueltoDatos As Boolean

Private WithEvents frmC As frmCal 'calendario fechas
Attribute frmC.VB_VarHelpID = -1
Private WithEvents frmCtas As frmCtasConta 'cuentas de contabilidad
Attribute frmCtas.VB_VarHelpID = -1
Private WithEvents frmFPa As frmForpaConta 'formas de pago de la contabilidad
Attribute frmFPa.VB_VarHelpID = -1

'GENERALES PARA PASARLE A CRYSTAL REPORT
Private cadFormula As String 'Cadena con la FormulaSelection para Crystal Report
Private cadParam As String 'Cadena con los parametros para Crystal Report
Private numParam As Byte 'Numero de parametros que se pasan a Crystal Report
Private cadSelect As String 'Cadena para comprobar si hay datos antes de abrir Informe
Private cadTitulo As String 'Titulo para la ventana frmImprimir
Private cadNombreRPT As String 'Nombre del informe

Dim indCodigo As Integer 'indice para txtCodigo
 
'Se inicializan para cada Informe (tabla de BD a la que hace referencia
Dim Codigo As String 'Código para FormulaSelection de Crystal Report
Dim TipCod As String

Dim PrimeraVez As Boolean

Private Sub KEYpress(KeyAscii As Integer)
    If KeyAscii = 13 Then 'ENTER
        KeyAscii = 0
        SendKeys "{tab}"
    ElseIf KeyAscii = 27 Then Unload Me  'ESC
    End If
End Sub

Private Sub cmdAceptar_Click()
Dim SQL As String
Dim i As Byte
Dim cadwhere As String
Dim cDesde As String
Dim cHasta As String

    If Not DatosOk Then Exit Sub
    
    cadSelect = "factsocio.intconta=0 "
    
    'D/H Fecha factura
    cDesde = Trim(txtCodigo(5).Text)
    cHasta = Trim(txtCodigo(6).Text)
    If Not (cDesde = "" And cHasta = "") Then
        'Cadena para seleccion Desde y Hasta
        Codigo = "{factsocio.fecfactu}"
        TipCod = "F"
        If Not PonerDesdeHasta(cDesde, cHasta, "", "", "pDHfechaFac= """) Then Exit Sub
    End If
    
    'D/H cuenta contable socio
    cDesde = Trim(txtCodigo(0).Text)
    cHasta = Trim(txtCodigo(9).Text)
    If Not (cDesde = "" And cHasta = "") Then
        'Cadena para seleccion Desde y Hasta
        Codigo = "{factsocio.codmacta}"
        TipCod = "T"
        If Not PonerDesdeHasta(cDesde, cHasta, "", "", "pDHCtaConta= """) Then Exit Sub
    End If
    
    'D/H numero de factura
    cDesde = Trim(txtCodigo(7).Text)
    cHasta = Trim(txtCodigo(8).Text)
    If Not (cDesde = "" And cHasta = "") Then
        'Cadena para seleccion Desde y Hasta
        Codigo = "{factsocio.numfactu}"
        TipCod = "N"
        If Not PonerDesdeHasta(cDesde, cHasta, "", "", "pDHColec= """) Then Exit Sub
    End If
    
    If Not HayRegParaInforme("factsocio", cadSelect) Then Exit Sub
    
    ContabilizarFacturas "factsocio", cadSelect
     'Eliminar la tabla TMP
    BorrarTMPFacturas
    'Desbloqueamos ya no estamos contabilizando facturas
    DesBloqueoManual ("CONSOC") 'CONtabilizar facturas SOCios
    
eError:
    If Err.Number <> 0 Then
        MuestraError Err.Number, "No se ha realizado el proceso de contabilización de facturas de socio. Llame a soporte."
    End If
    
    Pb1.visible = False
    lblProgres(0).Caption = ""
    lblProgres(1).Caption = ""
    cmdCancel_Click
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
    If PrimeraVez Then
        PrimeraVez = False
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
     Me.imgBuscar(4).Picture = frmPpal.imgListImages16.ListImages(1).Picture
     Me.imgBuscar(3).Picture = frmPpal.imgListImages16.ListImages(1).Picture
     Me.imgBuscar(0).Picture = frmPpal.imgListImages16.ListImages(1).Picture
     Me.imgBuscar(1).Picture = frmPpal.imgListImages16.ListImages(1).Picture
     txtCodigo(5).Text = Format(Now, "dd/mm/yyyy") ' fecha de factura desde
     txtCodigo(6).Text = Format(Now, "dd/mm/yyyy") ' fecha de factura hasta
     txtCodigo(1).Text = Format(Now, "dd/mm/yyyy") ' fecha de vencimiento

    '###Descomentar
'    CommitConexion
         
    FrameCobrosVisible True, h, w
    Pb1.visible = False
    
    'Esto se consigue poneinedo el cancel en el opcion k corresponda
    Me.cmdCancel.Cancel = True
'   Me.Width = w + 70
'   Me.Height = h + 350
End Sub

Private Sub frmC_Selec(vFecha As Date)
 'Fecha
    txtCodigo(CByte(imgFec(1).Tag)).Text = Format(vFecha, "dd/MM/yyyy")
End Sub

Private Sub frmTDia_DatoSeleccionado(CadenaSeleccion As String)
'Form de Consulta de Clientes
    txtCodigo(indCodigo).Text = Format(RecuperaValor(CadenaSeleccion, 1), "00")
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmCtas_DatoSeleccionado(CadenaSeleccion As String)
'Form de Consulta de Cuentas contables
    txtCodigo(indCodigo).Text = RecuperaValor(CadenaSeleccion, 1)
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmFpa_DatoSeleccionado(CadenaSeleccion As String)
'Form de Consulta de formas de pago de contabilidad
    txtCodigo(indCodigo).Text = RecuperaValor(CadenaSeleccion, 1)
    txtCodigo(indCodigo).Text = Format(txtCodigo(indCodigo).Text, "000")
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
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
    imgFec(1).Tag = Index 'independentment de les dates que tinga, sempre pose l'index en la 27
    If txtCodigo(Index).Text <> "" Then frmC.NovaData = txtCodigo(Index).Text

    frmC.Show vbModal
    Set frmC = Nothing
    PonerFoco txtCodigo(CByte(imgFec(1).Tag))
    ' ***************************
End Sub

Private Sub ListView1_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub imgBuscar_Click(Index As Integer)
   Select Case Index
        Case 3 ' forma de pago de la tesoreria
            AbrirFrmForpaConta (Index)
        Case 4 'cuenta contable banco
            AbrirFrmCuentas (Index)
        Case 0 ' cuenta contable desde
            AbrirFrmCuentas (Index)
        Case 1 ' cuenta contable hasta
            AbrirFrmCuentas (9)
    End Select
    PonerFoco txtCodigo(indCodigo)
End Sub

Private Sub Optcodigo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then 'ENTER
        PonerFocoBtn Me.cmdAceptar
    ElseIf KeyAscii = 27 Then Unload Me  'ESC
    End If
End Sub

Private Sub OptNombre_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then 'ENTER
        PonerFocoBtn Me.cmdAceptar
    ElseIf KeyAscii = 27 Then Unload Me  'ESC
    End If
End Sub

Private Sub txtCodigo_GotFocus(Index As Integer)
    ConseguirFoco txtCodigo(Index), 3
End Sub

Private Sub txtCodigo_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    KEYdown KeyCode
End Sub

Private Sub txtCodigo_KeyPress(Index As Integer, KeyAscii As Integer)
'14/02/2007
'    KEYpress KeyAscii
' ahora
    If KeyAscii = teclaBuscar Then
        Select Case Index
            Case 0: KEYBusqueda KeyAscii, 0 'cta contable desde
            Case 9: KEYBusqueda KeyAscii, 1 'cta contable hasta
            Case 5: KEYFecha KeyAscii, 2 'fecha desde factura
            Case 6: KEYFecha KeyAscii, 3 'fecha hasta factura
            Case 1: KEYFecha KeyAscii, 1 'fecha vencimiento
            Case 4: KEYBusqueda KeyAscii, 4 'cta contable banco
            Case 3: KEYBusqueda KeyAscii, 3 'forma de pago
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
Dim Cad As String, cadTipo As String 'tipo cliente

    'Quitar espacios en blanco por los lados
    txtCodigo(Index).Text = Trim(txtCodigo(Index).Text)
    
    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub

    
    Select Case Index
        Case 3 ' FORMA DE PAGO DE LA CONTABILIDAD
            If vParamAplic.ContabilidadNueva Then
                If txtCodigo(Index).Text <> "" Then txtNombre(Index).Text = DevuelveDesdeBDNew(cContaFacSoc, "formapago", "nomforpa", "codforpa", txtCodigo(3).Text, "N")
            Else
                If txtCodigo(Index).Text <> "" Then txtNombre(Index).Text = DevuelveDesdeBDNew(cContaFacSoc, "sforpa", "nomforpa", "codforpa", txtCodigo(3).Text, "N")
            End If
            If txtNombre(Index).Text = "" Then
                MsgBox "Forma de Pago  no existe en la contabilidad. Reintroduzca.", vbExclamation
            End If
            
        Case 4 ' CUENTA CONTABLE
            If txtCodigo(Index).Text <> "" Then txtNombre(Index).Text = PonerNombreCuenta(txtCodigo(Index), 2, , cContaFacSoc, False)
            If txtNombre(Index).Text = "" Then
                MsgBox "Número de Cuenta contable no existe en la contabilidad. Reintroduzca.", vbExclamation
            End If

        Case 5, 6 'FECHAS
            If txtCodigo(Index).Text <> "" Then
                If PonerFormatoFecha(txtCodigo(Index)) Then
                    If Index = 5 Then
                        txtCodigo(6).Text = txtCodigo(5).Text
                    End If
                End If
            End If
            
        Case 1 'FECHAS de vencimiento
            If txtCodigo(Index).Text <> "" Then PonerFormatoFecha txtCodigo(Index)
            
            
        Case 0, 9 ' CUENTAs CONTABLEs
            If txtCodigo(Index).Text <> "" Then txtNombre(Index).Text = PonerNombreCuenta(txtCodigo(Index), 2, , cContaFacSoc, False)
            
    End Select
End Sub

Private Sub FrameCobrosVisible(visible As Boolean, ByRef h As Integer, ByRef w As Integer)
    Me.FrameCobros.visible = visible
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

Private Sub AbrirFrmCuentas(indice As Integer)
    indCodigo = indice
    Set frmCtas = New frmCtasConta
    frmCtas.DatosADevolverBusqueda = "0|1|"
    frmCtas.CodigoActual = txtCodigo(indCodigo)
    frmCtas.Conexion = cContaFacSoc
    frmCtas.Facturas = False
    frmCtas.Show vbModal
    Set frmCtas = Nothing
End Sub

Private Sub AbrirFrmForpaConta(indice As Integer)
    indCodigo = indice
    Set frmFPa = New frmForpaConta
    frmFPa.DatosADevolverBusqueda = "0|1|"
    frmFPa.CodigoActual = txtCodigo(indCodigo)
    frmFPa.Conexion = cContaFacSoc
    frmFPa.Show vbModal
    Set frmFPa = Nothing
End Sub
 

Private Sub ContabilizarIntereses(cadwhere As String)
'Contabiliza Facturas de Clientes o de Proveedores
Dim SQL As String
Dim b As Boolean
Dim tmpErrores As Boolean 'Indica si se creo correctamente la tabla de errores
Dim CCoste As String
Dim cadTABLA As String

    SQL = "CONINT" 'contabilizar CALCULO DE INTERESES

    'Bloquear para que nadie mas pueda contabilizar
    DesBloqueoManual (SQL)
    If Not BloqueoManual(SQL, "1") Then
        MsgBox "No se pueden Contabilizar Cálculo de Intereses. Hay otro usuario contabilizándolo.", vbExclamation
        Screen.MousePointer = vbDefault
        Exit Sub
    End If

    
    'Visualizar la barra de Progreso
    Me.Pb1.visible = True
'    Me.Pb1.Top = 3350
    
    
    '==========================================================
    'REALIZAR COMPROBACIONES ANTES DE CONTABILIZAR FACTURAS
    '==========================================================
    
    Me.lblProgres(0).Caption = "Comprobaciones: "
    CargarProgres Me.Pb1, 100
        
    ' nuevo
    b = CrearTMPErrComprob()
    If Not b Then Exit Sub
    
    
    'comprobar que todas las CUENTAS de codigos avnic existen
    'en la Conta: savnic.codmacta IN (conta.cuentas.codmacta)
    '-----------------------------------------------------------------------------
    Me.lblProgres(1).Caption = "Comprobando Cuenta Ctble Retención ..."
    b = ComprobarCtaContable(cadTABLA, 1, "movim.fechamov = " & DBSet(txtCodigo(0).Text, "F") & " and intconta = 0", cConta)
    IncrementarProgres Me.Pb1, 33
    Me.Refresh
    If Not b Then
        frmMensaje.OpcionMensaje = 2
        frmMensaje.Show vbModal
        Exit Sub
    End If

    
    'comprobar que todas las CUENTAS de gasto existen
    'en la Conta: sparam.ctagasto IN (conta.cuentas.codmacta)
    '-----------------------------------------------------------------------------
    Me.lblProgres(1).Caption = "Comprobando Cuenta Ctble Gasto ..."
    b = ComprobarCtaContable(cadTABLA, 6, , cConta)
    IncrementarProgres Me.Pb1, 33
    Me.Refresh
    If Not b Then
        frmMensaje.OpcionMensaje = 2
        frmMensaje.Show vbModal
        Exit Sub
    End If
    'comprobar que todas las CUENTAS de retencion existen
    'en la Conta: sparam.ctareten IN (conta.cuentas.codmacta)
    '-----------------------------------------------------------------------------
    Me.lblProgres(1).Caption = "Comprobando Cuenta Ctble Retención ..."
    b = ComprobarCtaContable(cadTABLA, 7, , cConta)
    IncrementarProgres Me.Pb1, 33
    Me.Refresh
    If Not b Then
        frmMensaje.OpcionMensaje = 2
        frmMensaje.Show vbModal
        Exit Sub
    End If
     
     
     
    
    '===========================================================================
    'CONTABILIZAR CIERRE
    '===========================================================================
    Me.lblProgres(0).Caption = "Contabilizar Cierre: "
    CargarProgres Me.Pb1, 10
    Me.lblProgres(1).Caption = "Insertando Asiento en Contabilidad..."
    
    
    cadwhere = "fechamov = " & DBSet(txtCodigo(0).Text, "F") & " and intconta = 0"
    b = PasarCalculoAContab(cadwhere)
    
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
    
End Sub

Private Function DatosOk() As Boolean
Dim b As Boolean
Dim Orden1 As String
Dim Orden2 As String
Dim FFin As Date
Dim Cta As String

   b = True

   

   If txtCodigo(6).Text = "" Then
        MsgBox "Introduzca la Fecha de Factura a contabilizar.", vbExclamation
        b = False
        PonerFoco txtCodigo(6)
   Else
        ' comprobamos que la contabilizacion se encuentre en los ejercicios contables
         Orden1 = ""
         Orden1 = vEmpresaFacSoc.FechaIni '  DevuelveDesdeBDNew(cContaFacSoc, "parametros", "fechaini", "", "", "", "", "", "", "", "", "", "")
    
         Orden2 = ""
         Orden2 = vEmpresaFacSoc.FechaFin ' DevuelveDesdeBDNew(cContaFacSoc, "parametros", "fechafin", "", "", "", "", "", "", "", "", "", "")
         FIni = CDate(Orden1)
         FFin = CDate(Orden2)
         If Not (CDate(Orden1) <= CDate(txtCodigo(6).Text) And CDate(txtCodigo(6).Text) < CDate(Day(FIni) & "/" & Month(FIni) & "/" & Year(FIni) + 2)) Then
            MsgBox "La Fecha de la contabilización no es del ejercicio actual ni del siguiente. Reintroduzca.", vbExclamation
            b = False
            PonerFoco txtCodigo(6)
         End If
   End If
    
   If txtCodigo(1).Text = "" And b Then
        MsgBox "Introduzca la Fecha de Vencimiento a contabilizar.", vbExclamation
        b = False
        PonerFoco txtCodigo(1)
   End If
    
   If txtCodigo(3).Text = "" And b Then
        MsgBox "Introduzca la Forma de Pago para contabilizar.", vbExclamation
        b = False
        PonerFoco txtCodigo(3)
   End If
   
   'cta contable de banco
   If b Then
        If txtCodigo(4).Text = "" Then
             MsgBox "Introduzca la Cta.Contable d Banco para contabilizar.", vbExclamation
             b = False
             PonerFoco txtCodigo(4)
        Else
             Cta = ""
             Cta = DevuelveDesdeBDNew(cContaFacSoc, "cuentas", "codmacta", "codmacta", txtCodigo(4).Text, "T")
             If Cta = "" Then
                 MsgBox "La cuenta contable de Banco no existe. Reintroduzca.", vbExclamation
                 b = False
                 PonerFoco txtCodigo(4)
             End If
        End If
    End If
   DatosOk = b
   
End Function

Private Function PasarCalculoAContab(cadwhere As String) As Boolean
Dim SQL As String
Dim Rs As ADODB.Recordset
Dim b As Boolean
Dim i As Integer
Dim NumLinea As Integer
Dim Mc As CContadorContab
Dim numdocum As String
Dim ampliacion As String
Dim ampliaciond As String
Dim ampliacionh As String
Dim ImporteD As Currency
Dim ImporteH As Currency
Dim Diferencia As Currency
Dim Obs As String
Dim cadMen As String
Dim Cad As String
Dim CtaDifer As String
Dim Codmacta As String

    On Error GoTo EPasarCal

    PasarCalculoAContab = False
    
    'Total de lineas de asiento a Insertar en la contabilidad
    SQL = "SELECT count(*)" & _
          " FROM movim " & _
          "WHERE " & cadwhere
             
    NumLinea = TotalRegistros(SQL)
    
    If NumLinea = 0 Then Exit Function
    
    NumLinea = NumLinea * 3
    
    If NumLinea > 0 Then
        NumLinea = NumLinea + 1
        
        CargarProgres Me.Pb1, NumLinea
        
        ConnConta.BeginTrans
        conn.BeginTrans
        
        Set Mc = New CContadorContab
        
        If Mc.ConseguirContador("0", (CDate(txtCodigo(0).Text) <= CDate(FFin)), True, cConta) = 0 Then
        
        Obs = "Contabilizacion de Cálculo de Intereses AVNICS de fecha " & Format(txtCodigo(0).Text, "dd/mm/yyyy")

    
        'Insertar en la conta Cabecera Asiento
        b = InsertarCabAsientoDia("1", Mc.Contador, txtCodigo(0).Text, Obs, cadMen, cConta)
        cadMen = "Insertando Cab. Asiento: " & cadMen
        
        If b Then
            SQL = "SELECT codavnic, timporte, timport1, timport2 " & _
                  " FROM movim " & _
                  " WHERE " & cadwhere
            
            Set Rs = New ADODB.Recordset
            
            Rs.Open SQL, conn, adOpenDynamic, adLockOptimistic, adCmdText
            
            i = 0
            ImporteD = 0
            ImporteH = 0
            
            ampliacion = "Int.AVNICS AriagroUtil"
            ampliaciond = Trim(DevuelveDesdeBDNew(cConta, "conceptos", "nomconce", "codconce", vParamAplic.ConceDebe, "N")) & " " & ampliacion
            ampliacionh = Trim(DevuelveDesdeBDNew(cConta, "conceptos", "nomconce", "codconce", vParamAplic.ConceHaber, "N")) & " " & ampliacion
            
            
            If Not Rs.EOF Then Rs.MoveFirst
            While Not Rs.EOF And b
                Codmacta = ""
                Codmacta = DevuelveDesdeBDNew(cPTours, "avnic", "codmacta", "codavnic", Rs.Fields(0).Value, "N", , "anoejerc", Year(CDate(txtCodigo(0).Text)), "N")
                
                numdocum = "Av-" & Format(DBLet(Rs!codavnic, "N"), "000000")
                ' ******************IMPORTE BRUTO
                i = i + 1
                
                Cad = "1," & DBSet(txtCodigo(0).Text, "F") & "," & DBSet(Mc.Contador, "N") & ","
                Cad = Cad & DBSet(i, "N") & "," & DBSet(vParamAplic.CtaGasto, "T") & "," & DBSet(numdocum, "T") & ","
                
                ' COMPROBAMOS EL SIGNO DEL IMPORTE PQ NO PERMITIMOS INTRODUCIR APUNTES CON IMPORTES NEGATIVOS
                If Rs.Fields(2).Value > 0 Then
                    ' importe al debe en positivo
                    Cad = Cad & DBSet(vParamAplic.ConceDebe, "N") & "," & DBSet(ampliaciond, "T") & "," & DBSet(Rs.Fields(2).Value, "N") & ","
                    Cad = Cad & ValorNulo & "," & ValorNulo & "," & DBSet(Codmacta, "T") & "," & ValorNulo & ",0"
                
                    ImporteD = ImporteD + CCur(Rs.Fields(2).Value)
                Else
                    ' importe al haber en positivo, cambiamos el signo
                    Cad = Cad & DBSet(vParamAplic.ConceHaber, "N") & "," & DBSet(ampliacionh, "T") & "," & ValorNulo & ","
                    Cad = Cad & DBSet((Rs.Fields(2).Value * -1), "N") & "," & ValorNulo & "," & DBSet(Codmacta, "T") & "," & ValorNulo & ",0"
                
                    ImporteH = ImporteH + (CCur(Rs.Fields(2).Value) * (-1))
                End If
                
                Cad = "(" & Cad & ")"
                
                b = InsertarLinAsientoDia(Cad, cadMen, cConta)
                cadMen = "Insertando Lin. Asiento: " & i
            
                IncrementarProgres Me.Pb1, 1
                Me.lblProgres(1).Caption = "Insertando línea de Asiento en Contabilidad...   (" & i & " de " & NumLinea & ")"
                Me.Refresh
                
                ' ******************RETENCION
                i = i + 1
                
                Cad = "1," & DBSet(txtCodigo(0).Text, "F") & "," & DBSet(Mc.Contador, "N") & ","
                Cad = Cad & DBSet(i, "N") & "," & DBSet(vParamAplic.CtaReten, "T") & "," & DBSet(numdocum, "T") & ","
                
                ' COMPROBAMOS EL SIGNO DEL IMPORTE PQ NO PERMITIMOS INTRODUCIR APUNTES CON IMPORTES NEGATIVOS
                If Rs.Fields(3).Value > 0 Then
                    ' importe al haber en positivo
                    Cad = Cad & DBSet(vParamAplic.ConceHaber, "N") & "," & DBSet(ampliacionh, "T") & "," & ValorNulo & ","
                    Cad = Cad & DBSet(Rs.Fields(3).Value, "N") & "," & ValorNulo & "," & DBSet(Codmacta, "T") & "," & ValorNulo & ",0"
                
                    ImporteH = ImporteH + CCur(Rs.Fields(3).Value)
                Else
                    ' importe al debe en positivo, cambiamos el signo
                    Cad = Cad & DBSet(vParamAplic.ConceDebe, "N") & "," & DBSet(ampliaciond, "T") & "," & DBSet((Rs.Fields(3).Value * -1), "N") & ","
                    Cad = Cad & ValorNulo & "," & ValorNulo & "," & DBSet(Codmacta, "T") & "," & ValorNulo & ",0"
                
                    ImporteH = ImporteH + (CCur(Rs.Fields(3).Value) * (-1))
                End If
                
                Cad = "(" & Cad & ")"
                
                b = InsertarLinAsientoDia(Cad, cadMen, cConta)
                cadMen = "Insertando Lin. Asiento: " & i
            
                IncrementarProgres Me.Pb1, 1
                Me.lblProgres(1).Caption = "Insertando línea de Asiento en Contabilidad...   (" & i & " de " & NumLinea & ")"
                Me.Refresh
                
                ' ******************IMPORTE NETO
                i = i + 1
                
                Cad = "1," & DBSet(txtCodigo(0).Text, "F") & "," & DBSet(Mc.Contador, "N") & ","
                Cad = Cad & DBSet(i, "N") & "," & DBSet(Codmacta, "T") & "," & DBSet(numdocum, "T") & ","
                
                ' COMPROBAMOS EL SIGNO DEL IMPORTE PQ NO PERMITIMOS INTRODUCIR APUNTES CON IMPORTES NEGATIVOS
                If Rs.Fields(1).Value > 0 Then
                    ' importe al haber en positivo
                    Cad = Cad & DBSet(vParamAplic.ConceHaber, "N") & "," & DBSet(ampliacionh, "T") & "," & ValorNulo & ","
                    Cad = Cad & DBSet(Rs.Fields(1).Value, "N") & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & ",0"
                
                    ImporteH = ImporteH + CCur(Rs.Fields(1).Value)
                Else
                    ' importe al debe en positivo, cambiamos el signo
                    Cad = Cad & DBSet(vParamAplic.ConceDebe, "N") & "," & DBSet(ampliaciond, "T") & "," & DBSet((Rs.Fields(1).Value * -1), "N") & ","
                    Cad = Cad & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & ",0"
                
                    ImporteH = ImporteH + (CCur(Rs.Fields(1).Value) * (-1))
                End If
                
                Cad = "(" & Cad & ")"
                
                b = InsertarLinAsientoDia(Cad, cadMen, cConta)
                cadMen = "Insertando Lin. Asiento: " & i
            
                b = InsertarEnTesoreriaNew(txtCodigo(0).Text, txtCodigo(1).Text, Rs.Fields(0).Value, Year(CDate(txtCodigo(0).Text)), txtCodigo(4).Text, txtCodigo(2).Text, txtCodigo(3).Text, cadMen)
                cadMen = "Insertando en Tesoreria: "
               
            
                IncrementarProgres Me.Pb1, 1
                Me.lblProgres(1).Caption = "Insertando línea de Asiento en Contabilidad...   (" & i & " de " & NumLinea & ")"
                Me.Refresh
                
            
                Rs.MoveNext
            Wend
            Rs.Close
            
            If b Then
                'Poner intconta=1 en ariagroutil.movim
                b = ActualizarMovimientos(cadwhere, cadMen)
                cadMen = "Actualizando Movimientos: " & cadMen
            End If
        End If
    End If
   End If
   
EPasarCal:
    If Err.Number <> 0 Then
        b = False
        MuestraError Err.Number, "Integrando Asiento a Contabilidad", Err.Description
    End If
    If b Then
        ConnConta.CommitTrans
        conn.CommitTrans
        PasarCalculoAContab = True
    Else
        ConnConta.RollbackTrans
        conn.RollbackTrans
        PasarCalculoAContab = False
    End If
End Function


' copiado del ariges
Private Sub ContabilizarFacturas(cadTABLA As String, cadwhere As String)
'Contabiliza Facturas de Clientes o de Proveedores
Dim SQL As String
Dim b As Boolean
Dim tmpErrores As Boolean 'Indica si se creo correctamente la tabla de errores
Dim CCoste As String

    SQL = "CONSOC" 'contabilizar facturas de socios

    'Bloquear para que nadie mas pueda contabilizar
    DesBloqueoManual (SQL)
    If Not BloqueoManual(SQL, "1") Then
        MsgBox "No se pueden Contabilizar Facturas Socios. Hay otro usuario contabilizando.", vbExclamation
        Screen.MousePointer = vbDefault
        Exit Sub
    End If

'14/02/2007 lo he descomentado
     'comprobar que se han rellenado los dos campos de fecha
     'sino rellenar con fechaini o fechafin del ejercicio
     'que guardamos en vbles Orden1,Orden2

'     Orden1 = ""
'     Orden1 = DevuelveDesdeBDNew(cContaTel, "parametros", "fechaini", "", "", "", "", "", "", "", "", "", "")
'
'     Orden2 = ""
'     Orden2 = DevuelveDesdeBDNew(cContaTel, "parametros", "fechafin", "", "", "", "", "", "", "", "", "", "")
'
'     If txtCodigo(2).Text = "" Then
'        txtCodigo(2).Text = Orden1 'fechaini del ejercicio de la conta
'     End If
'
'     If txtCodigo(3).Text = "" Then
'        txtCodigo(3).Text = Orden2 'fecha fin del ejercicio de la conta
'     End If
'14/02/2007 hasta aqui lo he descomentado

    'comprobar si existen en Ariagroutil (facturas socios) facturas anteriores al periodo solicitado
    'sin contabilizar
    If Me.txtCodigo(6).Text <> "" Then
        SQL = "SELECT COUNT(*) FROM " & cadTABLA
        SQL = SQL & " WHERE fecfactu <"
        SQL = SQL & DBSet(txtCodigo(5), "F") & " AND intconta=0 "
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
    b = CrearTMPFacturasProveedor(cadTABLA, cadwhere)
    If Not b Then Exit Sub
            
    ' nuevo
    BorrarTMPErrComprob
    
    b = CrearTMPErrComprob()
    If Not b Then Exit Sub
    
    
    'comprobar que todas las CUENTAS de los distintos prov. que vamos a
    'contabilizar existen en la Conta: factsocio.codmacta IN (conta.cuentas.codmacta)
    '-----------------------------------------------------------------------------
    Me.lblProgres(1).Caption = "Comprobando Cuentas Contables de Socios en contabilidad ..."
    b = ComprobarCtaContable(cadTABLA, 8, , cContaFacSoc)
    IncrementarProgres Me.Pb1, 30
    Me.Refresh
    If Not b Then
        frmMensaje.OpcionMensaje = 2
        frmMensaje.Show vbModal
        Exit Sub
    End If
    
    'comprobar que todas las CUENTAS de  las variedades que vamos a
    'contabilizar existen en la Conta: (variedad.codmacta) IN (conta.cuentas.codmacta)
    '-----------------------------------------------------------------------------
    Me.lblProgres(1).Caption = "Comprobando Cuentas Ctbles de Variedades en contabilidad ..."
    b = ComprobarCtaContable(cadTABLA, 9, , cContaFacSoc)
    IncrementarProgres Me.Pb1, 20
    Me.Refresh
    If Not b Then
        frmMensaje.OpcionMensaje = 2
        frmMensaje.Show vbModal
        Exit Sub
    End If

    'comprobar la CUENTA de retencion
    'existe en la Conta: (vparamaplic.CtaRetenFacSoc) IN (conta.cuentas.codmacta)
    '-----------------------------------------------------------------------------
    Me.lblProgres(1).Caption = "Comprobando Cuentas Ctbles de Variedades en contabilidad ..."
    b = ComprobarCtaContable(cadTABLA, 10, , cContaFacSoc)
    IncrementarProgres Me.Pb1, 20
    Me.Refresh
    If Not b Then
        frmMensaje.OpcionMensaje = 2
        frmMensaje.Show vbModal
        Exit Sub
    End If

    'comprobar que todos las TIPO IVA de las distintas fecturas que vamos a
    'contabilizar existen en la Conta: scafac.codigiv1,codigiv2,codigiv3 IN (conta.tiposiva.codigiva)
    '--------------------------------------------------------------------------
    Me.lblProgres(1).Caption = "Comprobando Tipos de IVA en contabilidad ..."
    b = ComprobarTiposIVAFacSoc()
    IncrementarProgres Me.Pb1, 10
    Me.Refresh
    If Not b Then
        frmMensaje.OpcionMensaje = 2
        frmMensaje.Show vbModal
        Exit Sub
    End If


    'comprobar si hay contabilidad ANALITICA: conta.parametros.autocoste=1
    'y verificar que las cuentas de sfamia.ctaventa empiezan por el digito
    'de conta.parametros.grupogto o conta.parametros.grupovta
    'obtener el centro de coste del usuario para insertarlo en linfact
    If vEmpresaFacSoc.TieneAnalitica Then  'hay contab. analitica
       Me.lblProgres(1).Caption = "Comprobando Contabilidad Analítica ..."
       b = ComprobarCtaContable(cadTABLA, 11, , cContaFacSoc)
       IncrementarProgres Me.Pb1, 10
       Me.Refresh
       If Not b Then
            frmMensaje.OpcionMensaje = 2
            frmMensaje.Show vbModal
            Exit Sub
       End If
    End If
    


    b = ComprobarFormadePago(txtCodigo(3).Text)
    IncrementarProgres Me.Pb1, 10
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
    tmpErrores = CrearTMPErrFact("schfac")
    
    
    b = PasarFacturasAContab(cadTABLA, "")
    
    If Not b Then
        If tmpErrores Then
            'Cargar un listview con la tabla TEMP de Errores y mostrar
            'las facturas que fallaron
            frmMensaje.OpcionMensaje = 11
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


Private Function PasarFacturasAContab(cadTABLA As String, CCoste As String) As Boolean
Dim SQL As String
Dim Rs As ADODB.Recordset
Dim b As Boolean
Dim i As Integer
Dim numfactu As Integer
Dim codigo1 As String

    On Error GoTo EPasarFac

    PasarFacturasAContab = False

    '---- Obtener el total de Facturas a Insertar en la contabilidad
    SQL = "SELECT count(*) "
    SQL = SQL & " FROM " & cadTABLA & " INNER JOIN tmpFactu "
    SQL = SQL & " ON " & cadTABLA & ".codmacta = tmpFactu.codmacta "
    SQL = SQL & " AND " & cadTABLA & ".numfactu=tmpFactu.numfactu AND " & cadTABLA & ".fecfactu=tmpFactu.fecfactu "

    Set Rs = New ADODB.Recordset
    Rs.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not Rs.EOF Then
        numfactu = Rs.Fields(0)
    Else
        numfactu = 0
    End If
    Rs.Close
    Set Rs = Nothing


'    'Modificacion como David
'    '-----------------------------------------------------------
'    ' Mosrtaremos para cada factura de PROVEEDOR
'    ' que numregis le ha asignado
'    SQL = "DELETE FROM tmpinformes WHERE codusu = " & vUsu.Codigo
'    conn.Execute SQL

    '---- Pasar cada una de las facturas seleccionadas a la Conta
    If numfactu > 0 Then
        CargarProgres Me.Pb1, numfactu

        'seleccinar todas las facturas que hemos insertado en la temporal (las que vamos a contabilizar)
        SQL = "SELECT * "
        SQL = SQL & " FROM tmpFactu "

        Set Rs = New ADODB.Recordset
        Rs.Open SQL, conn, adOpenStatic, adLockPessimistic, adCmdText
        i = 1

        b = True
        'pasar a contabilidad cada una de las facturas seleccionadas
        While Not Rs.EOF
            SQL = cadTABLA & ".codmacta= " & DBSet(Rs!Codmacta, "T") & " and numfactu=" & DBSet(Rs!numfactu, "N")
            SQL = SQL & " and fecfactu=" & DBSet(Rs!fecfactu, "F")
            If PasarFacturaSoc(SQL, txtCodigo(1).Text, txtCodigo(4).Text, txtCodigo(3).Text) = False And b Then b = False

            '---- Laura 26/10/2006
            'Al pasar cada factura al hacer el commit desbloqueamos los registros
            'que teniamos bloqueados y los volvemos a bloquear
            'Laura: 11/10/2006 bloquear los registros q vamos a contabilizar
            SQL = cadTABLA & " INNER JOIN tmpFactu ON " & cadTABLA & ".codmacta=tmpFactu.codmacta AND " & cadTABLA & ".numfactu=tmpFactu.numfactu AND " & cadTABLA & ".fecfactu=tmpFactu.fecfactu "
            If Not BloqueaRegistro(SQL, cadTABLA & ".codmacta=tmpFactu.codmacta AND " & cadTABLA & ".numfactu=tmpFactu.numfactu AND " & cadTABLA & ".fecfactu=tmpFactu.fecfactu") Then
'                MsgBox "No se pueden Contabilizar Facturas. Hay registros bloqueados.", vbExclamation
'                Screen.MousePointer = vbDefault
'                Exit Sub
            End If
            '----

            IncrementarProgres Me.Pb1, 1
            Me.lblProgres(1).Caption = "Insertando Facturas en Contabilidad...   (" & i & " de " & numfactu & ")"
            Me.Refresh
            i = i + 1
            Rs.MoveNext
        Wend

        Rs.Close
        Set Rs = Nothing
    End If

EPasarFac:
    If Err.Number <> 0 Then b = False

    If b Then
        PasarFacturasAContab = True
    Else
        PasarFacturasAContab = False
    End If
End Function

