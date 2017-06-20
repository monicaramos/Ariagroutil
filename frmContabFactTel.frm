VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmContabFactTel 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Contabilización de Facturas "
   ClientHeight    =   5955
   ClientLeft      =   45
   ClientTop       =   2430
   ClientWidth     =   6645
   Icon            =   "frmContabFactTel.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5955
   ScaleWidth      =   6645
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
      Left            =   90
      TabIndex        =   13
      Top             =   120
      Width           =   6375
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   10
         Left            =   2445
         Locked          =   -1  'True
         TabIndex        =   33
         Text            =   "Text5"
         Top             =   2070
         Width           =   3450
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   10
         Left            =   1575
         MaxLength       =   6
         TabIndex        =   4
         Tag             =   "Código Propio|N|N|1|99|sbanco|codbanpr|00|S|"
         Top             =   2070
         Width           =   735
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   9
         Left            =   2445
         Locked          =   -1  'True
         TabIndex        =   31
         Text            =   "Text5"
         Top             =   1710
         Width           =   3450
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   9
         Left            =   1575
         MaxLength       =   6
         TabIndex        =   3
         Tag             =   "Código Propio|N|N|1|99|sbanco|codbanpr|00|S|"
         Top             =   1710
         Width           =   735
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   6
         Left            =   2760
         Locked          =   -1  'True
         TabIndex        =   29
         Text            =   "Text5"
         Top             =   1335
         Width           =   3135
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   6
         Left            =   1620
         MaxLength       =   10
         TabIndex        =   2
         Tag             =   "Código Propio|N|N|1|99|sbanco|codbanpr|00|S|"
         Top             =   1335
         Width           =   1050
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   8
         Left            =   1590
         MaxLength       =   10
         TabIndex        =   1
         Tag             =   "Código Propio|N|N|1|99|sbanco|codbanpr|00|S|"
         Top             =   945
         Width           =   1050
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   8
         Left            =   2760
         Locked          =   -1  'True
         TabIndex        =   24
         Text            =   "Text5"
         Top             =   945
         Width           =   3135
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   7
         Left            =   1605
         MaxLength       =   10
         TabIndex        =   0
         Tag             =   "Código Postal|T|S|||clientes|codposta|||"
         Top             =   480
         Width           =   1050
      End
      Begin VB.TextBox txtCodigo 
         Height          =   285
         Index           =   5
         Left            =   3720
         MaxLength       =   3
         TabIndex        =   6
         Top             =   2700
         Width           =   405
      End
      Begin VB.TextBox txtCodigo 
         Height          =   285
         Index           =   4
         Left            =   1605
         MaxLength       =   3
         TabIndex        =   5
         Top             =   2685
         Width           =   405
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   3
         Left            =   3735
         MaxLength       =   10
         TabIndex        =   10
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
         TabIndex        =   9
         Tag             =   "Código Postal|T|S|||clientes|codposta|||"
         Top             =   3615
         Width           =   1050
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Left            =   4755
         TabIndex        =   12
         Top             =   5205
         Width           =   975
      End
      Begin VB.CommandButton cmdAceptar 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   3570
         TabIndex        =   11
         Top             =   5205
         Width           =   975
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   0
         Left            =   1605
         MaxLength       =   7
         TabIndex        =   7
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
         TabIndex        =   8
         Tag             =   "Nº de Factura|N|N|0|9999999|schfac|numfactu|0000000|S|"
         Top             =   3150
         Width           =   830
      End
      Begin MSComctlLib.ProgressBar Pb1 
         Height          =   345
         Left            =   330
         TabIndex        =   23
         Top             =   4140
         Width           =   5370
         _ExtentX        =   9472
         _ExtentY        =   609
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   10
         Left            =   1290
         MouseIcon       =   "frmContabFactTel.frx":000C
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar iva"
         Top             =   2070
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Tipo Iva"
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
         Index           =   7
         Left            =   360
         TabIndex        =   34
         Top             =   2085
         Width           =   585
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   9
         Left            =   1290
         MouseIcon       =   "frmContabFactTel.frx":015E
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar forma pago"
         Top             =   1710
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Forma Pago"
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
         Index           =   6
         Left            =   360
         TabIndex        =   32
         Top             =   1725
         Width           =   855
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   6
         Left            =   1290
         MouseIcon       =   "frmContabFactTel.frx":02B0
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar cuenta"
         Top             =   1335
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
         Left            =   360
         TabIndex        =   30
         Top             =   1350
         Width           =   795
      End
      Begin VB.Label lblProgres 
         Height          =   285
         Index           =   1
         Left            =   330
         TabIndex        =   28
         Top             =   4815
         Width           =   5295
      End
      Begin VB.Label lblProgres 
         Height          =   285
         Index           =   0
         Left            =   360
         TabIndex        =   27
         Top             =   4500
         Width           =   5265
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Cta. Ventas"
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
         Index           =   5
         Left            =   360
         TabIndex        =   26
         Top             =   960
         Width           =   855
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   8
         Left            =   1290
         MouseIcon       =   "frmContabFactTel.frx":0402
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar cuenta"
         Top             =   945
         Width           =   240
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   7
         Left            =   1290
         Picture         =   "frmContabFactTel.frx":0554
         ToolTipText     =   "Buscar fecha"
         Top             =   480
         Width           =   240
      End
      Begin VB.Label Label4 
         Caption         =   "Fecha Vto"
         ForeColor       =   &H00972E0B&
         Height          =   255
         Index           =   4
         Left            =   360
         TabIndex        =   25
         Top             =   480
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
         TabIndex        =   22
         Top             =   2445
         Width           =   1005
      End
      Begin VB.Label Label4 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   1
         Left            =   2880
         TabIndex        =   21
         Top             =   2700
         Width           =   420
      End
      Begin VB.Label Label4 
         Caption         =   "Desde"
         Height          =   195
         Index           =   0
         Left            =   720
         TabIndex        =   20
         Top             =   2685
         Width           =   465
      End
      Begin VB.Label Label4 
         Caption         =   "Fecha Factura"
         ForeColor       =   &H00972E0B&
         Height          =   255
         Index           =   16
         Left            =   330
         TabIndex        =   19
         Top             =   3375
         Width           =   1815
      End
      Begin VB.Label Label4 
         Caption         =   "Desde"
         Height          =   195
         Index           =   15
         Left            =   720
         TabIndex        =   18
         Top             =   3615
         Width           =   465
      End
      Begin VB.Label Label4 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   14
         Left            =   2880
         TabIndex        =   17
         Top             =   3645
         Width           =   420
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   2
         Left            =   1290
         Picture         =   "frmContabFactTel.frx":05DF
         ToolTipText     =   "Buscar fecha"
         Top             =   3615
         Width           =   240
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   3
         Left            =   3420
         Picture         =   "frmContabFactTel.frx":066A
         ToolTipText     =   "Buscar fecha"
         Top             =   3645
         Width           =   240
      End
      Begin VB.Label Label4 
         Caption         =   "Desde"
         Height          =   195
         Index           =   13
         Left            =   720
         TabIndex        =   16
         Top             =   3150
         Width           =   465
      End
      Begin VB.Label Label4 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   12
         Left            =   2880
         TabIndex        =   15
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
         Left            =   360
         TabIndex        =   14
         Top             =   2910
         Width           =   555
      End
   End
End
Attribute VB_Name = "frmContabFactTel"
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
Dim cContaFra As cContabilizarFacturas


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
Dim Sql As String
Dim Tipo As Byte
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
        Codigo = "{" & tabla & ".numserie}"
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
    
    If Not HayRegParaInforme(tabla, cadSelect) Then Exit Sub
    
    If HayFacturasIncorrectas(tabla, cadSelect) Then Exit Sub
    
    
    ContabilizarFacturas tabla, cadSelect
     'Eliminar la tabla TMP
    BorrarTMPFacturas
    'Desbloqueamos ya no estamos contabilizando facturas
    DesBloqueoManual ("TELCON") 'VENtas CONtabilizar
    
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
     Me.imgBuscar(8).Picture = frmPpal.imgListImages16.ListImages(1).Picture
     Me.imgBuscar(9).Picture = frmPpal.imgListImages16.ListImages(1).Picture
     Me.imgBuscar(10).Picture = frmPpal.imgListImages16.ListImages(1).Picture
     
    '###Descomentar
'    CommitConexion
         
    FrameCobrosVisible True, h, w
    indFrame = 5
    tabla = "telmovil"
    
    Pb1.visible = False
    
    'Esto se consigue poneinedo el cancel en el opcion k corresponda
    Me.cmdCancel.Cancel = True
    Me.Width = w + 70
    Me.Height = h + 350
    
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

Private Sub frmFpa_DatoSeleccionado(CadenaSeleccion As String)
'Form de Consulta de forma de pago
    txtCodigo(indCodigo).Text = RecuperaValor(CadenaSeleccion, 1) 'forma de pago contable
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2) 'nombre forma de pago
End Sub

Private Sub frmTIva_DatoSeleccionado(CadenaSeleccion As String)
'Form de Consulta de tipos de iva
    txtCodigo(indCodigo).Text = RecuperaValor(CadenaSeleccion, 1) 'tipos de iva contable
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2) 'nombre tipos de iva
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
        
        Case 6, 8 ' ctas contables ventas y banco
            AbrirFrmCtasConta (Index)
        
        Case 9 ' forma de pago
            AbrirFrmForpaConta (Index)
        
        Case 10 ' tipo de iva de contabilidad
            AbrirFrmTipIvaConta (Index)
        
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
Dim Cad As String, cadTipo As String 'tipo cliente

    'Quitar espacios en blanco por los lados
    txtCodigo(Index).Text = Trim(txtCodigo(Index).Text)
    
    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub

    
    Select Case Index
        Case 6, 8 ' CTAS CONTABLES
            If txtCodigo(Index).Text = "" Then Exit Sub
            txtNombre(Index).Text = PonerNombreCuenta(txtCodigo(Index), 1, , cContaTel)
            
        Case 2, 3, 7  'FECHAS
            If txtCodigo(Index).Text <> "" Then PonerFormatoFecha txtCodigo(Index)
        
        Case 9 ' FORMA DE PAGO
            If txtCodigo(Index).Text = "" Then Exit Sub
            txtNombre(Index).Text = PonerNombreFPago(txtCodigo(Index), cContaTel)
        
        Case 10 ' TIPO DE IVA
            If txtCodigo(Index).Text = "" Then Exit Sub
            txtNombre(Index).Text = PonerNombreTIva(txtCodigo(Index), cContaTel)
        
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
        Me.FrameCobros.Height = 6015
        Me.FrameCobros.Width = 6555
        w = Me.FrameCobros.Width
        h = Me.FrameCobros.Height
    End If
End Sub

Private Sub ValoresPorDefecto()
    txtCodigo(7).Text = Format(Now, "dd/mm/yyyy")
    txtCodigo(6).Text = vParamAplic.CtaBancoTel
    If txtCodigo(6).Text <> "" Then txtNombre(6).Text = PonerNombreCuenta(txtCodigo(6), 1, , cContaTel)
    txtCodigo(8).Text = vParamAplic.CtaVentaTel
    If txtCodigo(8).Text <> "" Then txtNombre(8).Text = PonerNombreCuenta(txtCodigo(8), 1, , cContaTel)
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
    frmCta.Conexion = cContaTel
    frmCta.Show vbModal
    Set frmCta = Nothing
End Sub

Private Sub AbrirFrmForpaConta(indice As Integer)
    indCodigo = indice
    Set frmFPa = New frmForpaConta
    frmFPa.DatosADevolverBusqueda = "0|1|"
    frmFPa.CodigoActual = txtCodigo(indCodigo)
    frmFPa.Conexion = cContaTel
    frmFPa.Show vbModal
    Set frmFPa = Nothing
End Sub

Private Sub AbrirFrmTipIvaConta(indice As Integer)
    indCodigo = indice
    Set frmTIva = New frmTipIVAConta
    frmTIva.DatosADevolverBusqueda = "0|1|"
    frmTIva.CodigoActual = txtCodigo(indCodigo)
    frmTIva.Conexion = cContaTel
    frmTIva.Show vbModal
    Set frmTIva = Nothing
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
    If b Then
        If txtCodigo(8).Text = "" Then
            MsgBox "Debe introducir obligatoriamente una Cta Contable de Ventas.", vbExclamation
            b = False
            PonerFoco txtCodigo(8)
        Else
            ' comprobamos que la cta contable es del grupo de ventas
            cadG = DevuelveDesdeBDNew(cContaTel, "parametros", "grupovta", "", "", "")
            If Mid(txtCodigo(8).Text, 1, 1) <> cadG Then
                MsgBox "La Cuenta debe de ser del Grupo de Ventas. Reintroduzca.", vbExclamation
                b = False
                PonerFoco txtCodigo(8)
            End If
        End If
    End If
    
    If txtCodigo(9).Text = "" And b Then
        MsgBox "Debe introducir obligatoriamente una Forma de Pago.", vbExclamation
        b = False
        PonerFoco txtCodigo(9)
    End If

    If txtCodigo(10).Text = "" And b Then
        MsgBox "Debe introducir obligatoriamente un Tipo de Iva.", vbExclamation
        b = False
        PonerFoco txtCodigo(10)
    End If
     
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
Dim Sql As String
Dim b As Boolean
Dim tmpErrores As Boolean 'Indica si se creo correctamente la tabla de errores
Dim CCoste As String

    Sql = "TELCON" 'contabilizar facturas de venta

    'Bloquear para que nadie mas pueda contabilizar
    DesBloqueoManual (Sql)
    If Not BloqueoManual(Sql, "1") Then
        MsgBox "No se pueden Contabilizar Facturas. Hay otro usuario contabilizando.", vbExclamation
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

    'comprobar si existen en Ariagroutil (telefonia) facturas anteriores al periodo solicitado
    'sin contabilizar
    If Me.txtCodigo(2).Text <> "" Then
        Sql = "SELECT COUNT(*) FROM " & cadTABLA
        Sql = Sql & " WHERE fecfactu <"
        Sql = Sql & DBSet(txtCodigo(2), "F") & " AND intconta=0 "
        If RegistrosAListar(Sql) > 0 Then
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
    b = CrearTMPFacturas(cadTABLA, cadwhere, False, True)
    If Not b Then Exit Sub
            
    ' nuevo
    b = CrearTMPErrComprob()
    If Not b Then Exit Sub
    
    'comprobar que todas las LETRAS SERIE existen en la contabilidad y en Ariagroutil
    '-----------------------------------------------------------------------------
    IncrementarProgres Me.Pb1, 10
    Me.lblProgres(1).Caption = "Comprobando letras de serie ..."
    b = ComprobarLetraSerie(cContaTel)
    IncrementarProgres Me.Pb1, 30
    Me.Refresh
    If Not b Then
        frmMensaje.OpcionMensaje = 1
        frmMensaje.Show vbModal
        Exit Sub
    End If
    
    'comprobar que no haya Nº FACTURAS en la contabilidad para esa fecha
    'que ya existan
    '-----------------------------------------------------------------------
    Me.lblProgres(1).Caption = "Comprobando Nº Facturas en contabilidad ..."
    Sql = "anofaccl>=" & Year(txtCodigo(2).Text) & " AND anofaccl<= " & Year(txtCodigo(3).Text)
    b = ComprobarNumFacturas(cContaTel, Sql)
    IncrementarProgres Me.Pb1, 30
    Me.Refresh
    If Not b Then
        frmMensaje.OpcionMensaje = 1
        frmMensaje.Show vbModal
        Exit Sub
    End If
    
    'comprobar que todas las CUENTAS de los distintos clientes que vamos a
    'contabilizar existen en la Conta: telmovil.codmacta IN (conta.cuentas.codmacta)
    '-----------------------------------------------------------------------------
    Me.lblProgres(1).Caption = "Comprobando Cuentas Contables en contabilidad ..."
    b = ComprobarCtaContable(cadTABLA, 4, , cContaTel)
    IncrementarProgres Me.Pb1, 30
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
    
    
    b = PasarFacturasAContab(cadTABLA, txtCodigo(7).Text, txtCodigo(8).Text, CCoste)
    
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
Dim Sql As String
Dim Rs As ADODB.Recordset
Dim b As Boolean
Dim i As Integer
Dim numfactu As Integer
Dim codigo1 As String

    On Error GoTo EPasarFac

    PasarFacturasAContab = False
    
    'Total de Facturas a Insertar en la contabilidad
    Sql = "SELECT count(*) "
    Sql = Sql & " FROM " & cadTABLA & " INNER JOIN tmpfactu "
    codigo1 = "numserie"
    Sql = Sql & " ON " & cadTABLA & "." & codigo1 & "=tmpfactu." & codigo1
    Sql = Sql & " AND " & cadTABLA & ".numfactu=tmpfactu.numfactu AND " & cadTABLA & ".fecfactu=tmpfactu.fecfactu "
    
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not Rs.EOF Then
        numfactu = Rs.Fields(0)
    Else
        numfactu = 0
    End If
    Rs.Close
    Set Rs = Nothing

    If numfactu > 0 Then
        Set cContaFra = New cContabilizarFacturas
        
        If Not cContaFra.EstablecerValoresInciales(ConnContaTel) Then
            'NO ha establcedio los valores de la conta.  Le dejaremos seguir, avisando que
            ' obviamente, no va a contabilizar las FRAS
            Sql = "Si continua, las facturas se insertaran en el registro, pero no serán contabilizadas" & vbCrLf
            Sql = Sql & "en este momento. Deberán ser contabilizadas desde el ARICONTA" & vbCrLf & vbCrLf
            Sql = Sql & Space(50) & "¿Continuar?"
            If MsgBox(Sql, vbQuestion + vbYesNoCancel) <> vbYes Then Exit Function
        End If
    
        CargarProgres Me.Pb1, numfactu
        
        Sql = "SELECT * "
        Sql = Sql & " FROM tmpfactu "
            
        Set Rs = New ADODB.Recordset
        Rs.Open Sql, conn, adOpenStatic, adLockPessimistic, adCmdText
        i = 1

        b = True
        'contabilizar cada una de las facturas seleccionadas
        While Not Rs.EOF
            Sql = cadTABLA & "." & codigo1 & "=" & DBSet(Rs.Fields(0), "T") & " and numfactu=" & DBLet(Rs!numfactu, "N")
            Sql = Sql & " and fecfactu=" & DBSet(Rs!fecfactu, "F")
            If PasarFactura(Sql, FecVenci, txtCodigo(8).Text, txtCodigo(6).Text, txtCodigo(9).Text, txtCodigo(10), CCoste, cContaFra) = False And b Then b = False
            
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


Public Function HayFacturasIncorrectas(cTabla As String, cWhere As String) As Boolean
'Comprobar si hay registros a Mostrar antes de abrir el Informe
Dim Sql As String
Dim Rs As ADODB.Recordset
Dim PorIva As Currency

Dim BaseT As Currency
Dim IvaT As Currency
Dim TotalT As Currency

Dim IvaCal As Currency
Dim TotalCal As Currency

Dim Cad As String
Dim cad1 As String

    On Error GoTo eHayFacturasIncorrectas
    
    HayFacturasIncorrectas = True
    
    PorIva = 0
    Sql = DevuelveDesdeBDNew(cContaTel, "tiposiva", "porceiva", "codigiva", txtCodigo(10).Text, "N")
    If Sql <> "" Then PorIva = CCur(Sql)

    If cWhere <> "" Then
        cWhere = QuitarCaracterACadena(cWhere, "{")
        cWhere = QuitarCaracterACadena(cWhere, "}")
        cWhere = QuitarCaracterACadena(cWhere, "_1")
        Sql = Sql & " WHERE " & cWhere
    End If
    
    Sql = " SELECT numserie,numfactu,fecfactu,codmacta, year(fecfactu) as anofaccl,"
    Sql = Sql & "baseimpo,cuotaiva,totalfac "
    Sql = Sql & " FROM telmovil "
    Sql = Sql & " WHERE " & cWhere
    Sql = Sql & " order by numserie, numfactu, fecfactu"
    
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    Cad = ""
    
    While Not Rs.EOF
        BaseT = DBLet(Rs!BaseImpo, "N")
        IvaT = DBLet(Rs!CuotaIva, "N")
        TotalT = DBLet(Rs!TotalFac, "N")
        
        IvaCal = Round2(BaseT * PorIva / 100, 2)
        TotalCal = BaseT + IvaT

'[Monica]15/09/2011: solo voy a comprobar que base + iva = total factura
'                    sustituyo la siguiente instruccion por la de abajo
'        If IvaT <> IvaCal Or TotalT <> TotalCal Then cad = cad & DBLet(Rs!NumFactu, "N") & ", "
        If BaseT + IvaT <> TotalT Then Cad = Cad & DBLet(Rs!numfactu, "N") & ", "
        
        Rs.MoveNext
    Wend
    
    Set Rs = Nothing
    
    If Cad <> "" Then
        cad1 = "Las siguientes facturas tienen importes incorrectos. Revise." & vbCrLf & vbCrLf & Mid(Cad, 1, Len(Cad) - 2)
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

