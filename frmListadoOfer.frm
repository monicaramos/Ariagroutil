VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmListadoOfer 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Informes"
   ClientHeight    =   6360
   ClientLeft      =   45
   ClientTop       =   2430
   ClientWidth     =   10245
   Icon            =   "frmListadoOfer.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6360
   ScaleWidth      =   10245
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog cd1 
      Left            =   6435
      Top             =   5985
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame FrameEnvioFacMail 
      Height          =   6015
      Left            =   0
      TabIndex        =   16
      Top             =   0
      Width           =   10215
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   8
         Left            =   1335
         MaxLength       =   6
         TabIndex        =   0
         Top             =   1080
         Width           =   830
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   8
         Left            =   2250
         Locked          =   -1  'True
         TabIndex        =   34
         Text            =   "Text5"
         Top             =   1080
         Width           =   3135
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   1
         Left            =   3840
         MaxLength       =   3
         TabIndex        =   8
         Text            =   "wwwwwww"
         Top             =   4980
         Width           =   525
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   0
         Left            =   1320
         MaxLength       =   3
         TabIndex        =   7
         Top             =   5010
         Width           =   405
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   107
         Left            =   3810
         MaxLength       =   7
         TabIndex        =   6
         Top             =   4230
         Width           =   1365
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   106
         Left            =   1290
         MaxLength       =   7
         TabIndex        =   5
         Text            =   "wwwwwww"
         Top             =   4230
         Width           =   1365
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   108
         Left            =   1290
         MaxLength       =   10
         TabIndex        =   3
         Top             =   3255
         Width           =   1080
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   109
         Left            =   3810
         MaxLength       =   10
         TabIndex        =   4
         Top             =   3255
         Width           =   1080
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   18
         Left            =   9000
         TabIndex        =   12
         Top             =   5370
         Width           =   975
      End
      Begin VB.TextBox Text1 
         Height          =   320
         Index           =   0
         Left            =   5640
         TabIndex        =   9
         Text            =   "Text1"
         Top             =   1950
         Width           =   4335
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   110
         Left            =   2580
         Locked          =   -1  'True
         TabIndex        =   18
         Text            =   "Text5"
         Top             =   1875
         Width           =   2655
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   110
         Left            =   1320
         MaxLength       =   10
         TabIndex        =   1
         Top             =   1875
         Width           =   1215
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   111
         Left            =   2580
         Locked          =   -1  'True
         TabIndex        =   17
         Text            =   "Text5"
         Top             =   2370
         Width           =   2655
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   111
         Left            =   1320
         MaxLength       =   10
         TabIndex        =   2
         Top             =   2370
         Width           =   1215
      End
      Begin VB.TextBox Text1 
         Height          =   2355
         Index           =   1
         Left            =   5640
         MultiLine       =   -1  'True
         TabIndex        =   10
         Text            =   "frmListadoOfer.frx":000C
         Top             =   2760
         Width           =   4335
      End
      Begin VB.CommandButton cmdEnvioMail 
         Caption         =   "Enviar"
         Height          =   375
         Left            =   7920
         TabIndex        =   11
         Top             =   5370
         Width           =   975
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Sección"
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
         Left            =   240
         TabIndex        =   35
         Top             =   1050
         Width           =   540
      End
      Begin VB.Image imgBuscarOfer 
         Height          =   240
         Index           =   8
         Left            =   1020
         MouseIcon       =   "frmListadoOfer.frx":0012
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar Sección"
         Top             =   1080
         Width           =   240
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "Desde"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   1
         Left            =   600
         TabIndex        =   33
         Top             =   5010
         Width           =   450
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "Hasta"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   3360
         TabIndex        =   32
         Top             =   5010
         Width           =   420
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "Hasta"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   13
         Left            =   3330
         TabIndex        =   31
         Top             =   4215
         Width           =   420
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "Desde"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   14
         Left            =   570
         TabIndex        =   30
         Top             =   4215
         Width           =   450
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "Nº Factura"
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
         Index           =   15
         Left            =   240
         TabIndex        =   29
         Top             =   3930
         Width           =   780
      End
      Begin VB.Label Label14 
         Caption         =   "Envio facturas por mail"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   345
         Index           =   16
         Left            =   240
         TabIndex        =   28
         Top             =   360
         Width           =   4455
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "Desde"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   17
         Left            =   570
         TabIndex        =   27
         Top             =   3300
         Width           =   450
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Fact."
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
         Index           =   1
         Left            =   240
         TabIndex        =   26
         Top             =   3000
         Width           =   855
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   33
         Left            =   1050
         Picture         =   "frmListadoOfer.frx":0164
         Top             =   3285
         Width           =   240
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   34
         Left            =   3570
         Picture         =   "frmListadoOfer.frx":01EF
         Top             =   3285
         Width           =   240
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "Hasta"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   18
         Left            =   3090
         TabIndex        =   25
         Top             =   3300
         Width           =   420
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "Asunto"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   19
         Left            =   5640
         TabIndex        =   24
         Top             =   1710
         Width           =   510
      End
      Begin VB.Image imgBuscarOfer 
         Height          =   240
         Index           =   56
         Left            =   1020
         Top             =   1875
         Width           =   240
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Cliente"
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
         Index           =   32
         Left            =   240
         TabIndex        =   23
         Top             =   1500
         Width           =   495
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Desde"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   33
         Left            =   540
         TabIndex        =   22
         Top             =   1875
         Width           =   450
      End
      Begin VB.Image imgBuscarOfer 
         Height          =   240
         Index           =   57
         Left            =   1020
         Top             =   2370
         Width           =   240
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Hasta"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   34
         Left            =   540
         TabIndex        =   21
         Top             =   2370
         Width           =   420
      End
      Begin VB.Label Label14 
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
         Index           =   20
         Left            =   240
         TabIndex        =   20
         Top             =   4710
         Width           =   1005
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "Mensaje"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   21
         Left            =   5640
         TabIndex        =   19
         Top             =   2520
         Width           =   600
      End
   End
   Begin VB.Frame FrameEnvioMail 
      Height          =   1215
      Left            =   0
      TabIndex        =   13
      Top             =   30
      Visible         =   0   'False
      Width           =   6615
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   375
         Left            =   360
         TabIndex        =   14
         Top             =   360
         Width           =   5895
         _ExtentX        =   10398
         _ExtentY        =   661
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "Preparando datos envio"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   22
         Left            =   360
         TabIndex        =   15
         Top             =   840
         Width           =   5805
      End
   End
End
Attribute VB_Name = "frmListadoOfer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Event DatoSeleccionado(CadenaSeleccion As String)

Public OpcionListado As Integer
    '(ver opciones en frmListado)
        
        
        
    '315:  Envio por mail de las facturas
        
Public NumCod As String 'Para indicar cod. Traspaso,Movimiento, etc. que llama
                        'Para indicar nº oferta/pedido a imprimir

Public CodClien As String 'Para seleccionar inicialmente las ofertas del Cliente
                          'en el listado de Recordatorio de Ofertas y de Valoracion de Ofertas

Public FecEntre As String 'Para pasar inicialmente la fecha de entrega de la Oferta que se va a pasar a pedido
                          'como la fecha de entega del PEdido
                          
Dim BdConta As Integer ' numero de la contabilidad donde se hace conexion

Private NomTabla As String
Private NomTablaLin As String

Dim tabla As String
Dim TipCod As String


'Private WithEvents frmMtoCartasOfe As frmFacCartasOferta
Private WithEvents frmCtas As frmCtasConta
Attribute frmCtas.VB_VarHelpID = -1
Private WithEvents frmSec As frmManSecciones
Attribute frmSec.VB_VarHelpID = -1
Private WithEvents frmF As frmCal 'Calendario de Fechas
Attribute frmF.VB_VarHelpID = -1

'Private WithEvents frmCP As frmCPostal 'codigo postal
'Private WithEvents frmMen As frmMensaje  'Form Mensajes para mostrar las etiquetas a imprimir

'----- Variables para el INforme ----
Private cadFormula As String 'Cadena con la FormulaSelection para Crystal Report
Private cadParam As String 'cadena con los parametros q se pasan a Crystal R.
Private numParam As Byte
Private cadSelect As String 'Cadena para comprobar si hay datos antes de abrir Informe
Private Titulo As String 'Titulo informe que se pasa a frmImprimir
Private nomRPT As String 'nombre del fichero .rpt a imprimir
Private conSubRPT As Boolean 'si tiene subinformes para enlazarlos a las tablas correctas
'-------------------------------------

Dim indCodigo As Byte 'indice para txtCodigo
Dim Codigo As String 'Código para FormulaSelection de Crystal Report

Dim PrimeraVez As Boolean


Private Sub KEYpress(KeyAscii As Integer)
Dim cerrar As Boolean

    KEYpressGnral KeyAscii, 2, cerrar
    If cerrar Then Unload Me
End Sub

Private Sub chkEmail_KeyPress(Index As Integer, KeyAscii As Integer)
     KEYpress KeyAscii
End Sub


Private Sub cmdAceptarAlbCom_Click()
'Solicitar datos para Generar Albaran  a partir de Pedido de Compras
Dim Cad As String

    Cad = "" 'txtCodigo(47).Text & "|"
    Cad = Cad & txtCodigo(48).Text & "|"
    Cad = Cad & txtCodigo(49).Text & "|"
    RaiseEvent DatoSeleccionado(Cad)
    Unload Me
End Sub




Private Sub cmdCancel_Click(Index As Integer)
    Unload Me
End Sub


Private Sub cmdEnvioMail_Click()
Dim cDesde As String, cHasta As String 'cadena codigo Desde/Hasta
Dim nDesde As String, nHasta As String 'cadena Descripcion Desde/Hasta
Dim cadTABLA As String, cOrden As String
Dim i As Byte
Dim b As Boolean

Dim Rs As ADODB.Recordset


    'El proceso constara de varias fases.
    'Fase 1: Montar el select y ver si hay registros
    'Fase 2: Preparar carpetas para los pdf
    'Fase 3: Generar para cada factura (una a una) del select su pdf
    'Fase 4: Enviar por mail, adjuntando los archivos correspondientes
    
    If Text1(0).Text = "" Then
        MsgBox "Ponga el asunto", vbExclamation
        Exit Sub
    End If
    
    'Cadena para seleccion Desde y Hasta FECHA
    '--------------------------------------------
    InicializarVbles
    cadFormula = ""
    cadSelect = ""
    
    InicializarVbles
    
    If txtCodigo(8).Text = "" Then
        MsgBox "Introduzca la Sección.", vbExclamation
        Exit Sub
    End If
    
    If Not AnyadirAFormula(cadFormula, "{" & tabla & ".codsecci} = " & txtCodigo(8).Text) Then Exit Sub
    If Not AnyadirAFormula(cadSelect, tabla & ".codsecci = " & DBSet(txtCodigo(8).Text, "N")) Then Exit Sub
    
    '========= PARAMETROS  =============================
    'Añadir el parametro de Empresa
    cadParam = cadParam & "|pEmpresa=""" & vEmpresa.nomEmpre & """|"
    numParam = numParam + 1
    
    'D/H Cuenta contable
    cDesde = Trim(txtCodigo(110).Text)
    cHasta = Trim(txtCodigo(111).Text)
    nDesde = txtNombre(110).Text
    nHasta = txtNombre(111).Text
    If Not (cDesde = "" And cHasta = "") Then
        'Cadena para seleccion Desde y Hasta
        Codigo = "" & tabla & ".ctaclien"
        TipCod = "T"
        If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHcliente= """) Then Exit Sub
    End If
    
    'D/H Fecha factura
    cDesde = Trim(txtCodigo(108).Text)
    cHasta = Trim(txtCodigo(109).Text)
    If Not (cDesde = "" And cHasta = "") Then
        'Cadena para seleccion Desde y Hasta
        Codigo = "" & tabla & ".fecfactu"
        TipCod = "F"
        If Not PonerDesdeHasta(cDesde, cHasta, "", "", "pDHfechaFac= """) Then Exit Sub
    End If
    
    'D/H Serie
    cDesde = Trim(txtCodigo(0).Text)
    cHasta = Trim(txtCodigo(1).Text)
    If Not (cDesde = "" And cHasta = "") Then
        'Cadena para seleccion Desde y Hasta
        Codigo = "" & tabla & ".letraser"
        TipCod = "T"
        If Not PonerDesdeHasta(cDesde, cHasta, "", "", "pDHSerie= """) Then Exit Sub
    End If
    
    'Factura
    cDesde = Trim(txtCodigo(106).Text)
    cHasta = Trim(txtCodigo(107).Text)
    If Not (cDesde = "" And cHasta = "") Then
        'Cadena para seleccion Desde y Hasta
        Codigo = "" & tabla & ".numfactu"
        TipCod = "N"
        If Not PonerDesdeHasta(cDesde, cHasta, "", "", "pDHFact= """) Then Exit Sub
    End If
    
    Screen.MousePointer = vbHourglass
    
    'Eliminamos temporales
    conn.Execute "DELETE from tmpinformes where codusu =" & vSesion.Codigo
    
    If cadSelect <> "" Then cadSelect = " WHERE " & cadSelect
    
    Set Rs = New ADODB.Recordset
    DoEvents
        
    'Ahora insertare en la tabla temporal tmpinformes las facturas que voy a generar pdf
'    Codigo = "insert into tmpinformes (codusu,numalbar,codprove,codartic,numlinea,fechaalb,codalmac,cantidad) "
                                            'ctaclien,numfactu, letraser,fecfactu,totalfac
    Codigo = "insert into tmpinformes (codusu,nombre1,importe1, nombre2, fecha1, importe2) "
    Codigo = Codigo & " values ( " & vSesion.Codigo & ","
    
    If Not PrepararCarpetasEnvioMail Then Exit Sub
        
    Screen.MousePointer = vbHourglass

    'Vamos a meter todas las facturas en la tabla temporal para comprobar si tienen mail
    'los clientes
    
    NomTabla = "Select letraser,numfactu,ctaclien,fecfactu,totalfac from cabfact  " & cadSelect
    'El orden vamos a hacerlo por: Tipo documento
    NomTabla = NomTabla & " ORDER BY letraser, numfactu, fecfactu "
    Rs.Open NomTabla, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    NumRegElim = 0
    While Not Rs.EOF
        NomTabla = DBSet(Rs!ctaclien, "T") & "," & Rs!numfactu & ",'" & Trim(Rs!letraser) & "','" & Format(Rs!fecfactu, FormatoFecha)
        
        'El tipo de informe lo guardare en el ultimo campo
        'El report es el = 12
        NomTabla = NomTabla & "'," & TransformaComasPuntos(CStr(DBLet(Rs!TotalFac, "N"))) & ")"
        conn.Execute Codigo & NomTabla
        NumRegElim = NumRegElim + 1
        Rs.MoveNext
    Wend
    Rs.Close
    
    
    If NumRegElim = 0 Then
        MsgBox "Ningun dato a enviar por mail", vbExclamation
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    
    'Numero de registros
    NomTabla = NumRegElim
    
    'AHora ya tengo todos los datos de las facturas que voy  a imprimir
    'Entonces copruebo si para los clientes si tienen puesto el campo mail o no
    cadSelect = "Select nombre1 ,maidatos"
    cadSelect = cadSelect & " as email from tmpinformes,conta" & BdConta & ".cuentas cuentas where codusu = " & vSesion.Codigo & " and cuentas.codmacta=nombre1"
    cadSelect = cadSelect & " group by codmacta having email is null"
    Rs.Open cadSelect, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    NumRegElim = 0
    While Not Rs.EOF
        NumRegElim = NumRegElim + 1
        Rs.MoveNext
    Wend
    Rs.Close
    
    If NumRegElim > 0 Then
        If MsgBox("Tiene clientes sin mail. Continuar sin sus datos?", vbQuestion + vbYesNo) = vbNo Then
            Screen.MousePointer = vbDefault
            Exit Sub
        End If
            
        'Si no salimos borramos
        Rs.Open cadSelect, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        cadSelect = "DELETE from tmpinformes where codusu =" & vSesion.Codigo & " and nombre1 ="
        While Not Rs.EOF
            conn.Execute cadSelect & DBSet(Rs!nombre1, "T")
            Rs.MoveNext
        Wend
        Rs.Close
        
        
        cadSelect = "Select count(*) from tmpinformes where codusu =" & vSesion.Codigo
        Rs.Open cadSelect, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        NumRegElim = 0
        If Not Rs.EOF Then
            If Not IsNull(Rs.Fields(0)) Then NumRegElim = DBLet(Rs.Fields(0), "N")
            
        End If
        Rs.Close
        
        If NumRegElim = 0 Then
            'NO hay datos para enviar
            
            Screen.MousePointer = vbDefault
            MsgBox "No hay datos para enviar por mail", vbExclamation
            Exit Sub
        Else
            cadSelect = "Hay " & NumRegElim & " facturas para enviar por mail." & vbCrLf & "¿Continuar?"
            If MsgBox(cadSelect, vbQuestion + vbYesNo) = vbNo Then NumRegElim = 0
        End If
        If NumRegElim = 0 Then
            Set Rs = Nothing
            Screen.MousePointer = vbDefault
            Exit Sub
        End If
        NomTabla = NumRegElim
    
    End If
        
    PonerTamnyosMail True
    MDIppal.visible = False
    'Voy arriesgar.
    'Confio en que no envien por mail mas de 32000 facturas (un integer)
    Label14(22).Caption = "Preparando datos"
    Me.ProgressBar1.Max = CInt(NomTabla)
    Me.ProgressBar1.Value = 0
    
    
    
    NumRegElim = 0
    If GeneracionEnvioMail(Rs) Then NumRegElim = 1
        
    
    'Si ha ido todo bien entonces numregelim=1
    If NumRegElim = 1 Then
        cadSelect = "Select nommacta, maidatos"
        cadSelect = cadSelect & " as email,tmpinformes.* from tmpinformes,conta" & BdConta & ".cuentas cuentas where codusu = " & vSesion.Codigo & " and cuentas.codmacta=nombre1"
'        cadSelect = cadSelect & " group by codclien having email is null"

        
        frmEMail.DatosEnvio = Text1(0).Text & "|" & Text1(1).Text & "|1|" & cadSelect & "|"
        frmEMail.Opcion = 4 'Multienvio de facturacion
        frmEMail.Show vbModal
        
        
        'Para tranquilizar las pantallas, borrar los ficheros generados
        'Confio en que no envien por mail mas de 32000 facturas (un integer)
        Label14(22).Caption = "Restaurando ...."
        Me.ProgressBar1.visible = False
        Me.Refresh
        DoEvents
        espera 1
        PrepararCarpetasEnvioMail
        Me.ProgressBar1.visible = True
        
        
    End If
    
    
    
    
    'Es para evitar la cantidad de pantallas abriendose y cerrandose
    Me.visible = False
    PonerTamnyosMail False
    espera 1
    Unload Me
    MDIppal.Show

    Screen.MousePointer = vbDefault
End Sub

Private Sub PonerTamnyosMail(peque As Boolean)
    If peque Then
        Me.Height = Me.FrameEnvioMail.Height + 60
        Me.Width = Me.FrameEnvioMail.Width
    Else
        Me.Height = Me.FrameEnvioFacMail.Height
        Me.Width = Me.FrameEnvioFacMail.Width
    End If
    Me.Height = Me.Height + 420
    Me.Width = Me.Width + 120
    Me.FrameEnvioMail.visible = peque
    Me.FrameEnvioFacMail.visible = Not peque
    DoEvents
    Me.Refresh
End Sub


Private Function GeneracionEnvioMail(ByRef Rs As ADODB.Recordset) As Boolean

    On Error GoTo EGeneracionEnvioMail
    GeneracionEnvioMail = False

    
    cadSelect = "Select * from tmpinformes where codusu =" & vSesion.Codigo & " ORDER BY importe1,nombre2"
    Rs.Open cadSelect, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    CodClien = ""
    While Not Rs.EOF
        
        If Dir(App.path & "\docum.pdf", vbArchive) <> "" Then Kill App.path & "\docum.pdf"
    
        Label14(22).Caption = "Factura: " & Rs!importe1 & " " & Rs!nombre2
        Label14(22).Refresh
        
        Dim indRPT As Byte 'Indica el tipo de Documento en la tabla "scryst"
        Dim nomDocu As String 'Nombre de Informe rpt de crystal
        
        indRPT = 1 'Facturas Varias
        
        '[Monica]26/05/2016: otro report para materna
        If EsSeccionMaterna(txtCodigo(8).Text) Then indRPT = 4
        
       If Not PonerParamRPT(1, cadParam, numParam, nomDocu) Then Exit Function
        
       cadFormula = "({cabfact.codsecci}=" & CLng(txtCodigo(8).Text) & ") "
       cadFormula = cadFormula & " AND ({cabfact.letraser}='" & Trim(Rs!nombre2) & "') "
       cadFormula = cadFormula & " AND ({cabfact.numfactu}=" & Rs!importe1 & ") "
       cadFormula = cadFormula & " AND ({cabfact.fecfactu}= Date(" & Year(Rs!fecha1) & "," & Month(Rs!fecha1) & "," & Day(Rs!fecha1) & "))"

   
        With frmImprimir
            .FormulaSeleccion = cadFormula
            .OtrosParametros = cadParam
            .NumeroParametros = numParam
            .SoloImprimir = False
            .EnvioEMail = True
            .Titulo = "" 'cadTitulo
            .NombreRPT = nomDocu
            .Opcion = 1
            .Facturas = True
            .Contabilidad = BdConta
            .Show vbModal
        End With
    
                    
        'Subo el progress bar
        Label14(22).Caption = "Generando PDF"
        Label14(22).Refresh
        Me.ProgressBar1.Value = Me.ProgressBar1.Value + 1
        If (Me.ProgressBar1.Value Mod 25) = 24 Then
            Me.Refresh
            DoEvents
            espera 1
        End If
        Me.Refresh
        DoEvents
        
        
        
        'FileCopy App.Path & "\docum.pdf", App.Path & "\temp\" & RS!NumAlbar & Format(RS!codProve, "0000000") & Format(RS!codArtic, "0000000") & Format(RS!FechaAlb, "yymmdd") & ".pdf"
        FileCopy App.path & "\docum.pdf", App.path & "\temp\" & Trim(Rs!nombre2) & Format(Rs!importe1, "0000000") & ".pdf"
        
        Rs.MoveNext
    Wend
    Rs.Close
    
    Set Rs = Nothing
    GeneracionEnvioMail = True
    Exit Function
EGeneracionEnvioMail:
       MuestraError Err.Number
End Function


Private Sub Form_Activate()
    If PrimeraVez Then
        PrimeraVez = False
        BdConta = 0
        Select Case OpcionListado
            Case 315 ' envio de facturas por email
                PonerFoco txtCodigo(8)
                
        End Select
    End If
    Screen.MousePointer = vbDefault
End Sub


Private Sub Form_Load()
Dim h As Integer, w As Integer
Dim indFrame As Single
Dim devuelve As String
    
'    'Icono del formulario
'    Me.Icon = frmPpal.Icon
'
    PrimeraVez = True
    Limpiar Me
    indCodigo = 0
    NomTabla = ""

    'Ocultar todos los Frames de Formulario
    Me.FrameEnvioFacMail.visible = False
    
    CargarIconos
    
    Select Case OpcionListado
        'LISTADOS DE FACTURACION
        '-----------------------
        Case 315
            tabla = "cabfact"
            indFrame = 18
            h = FrameEnvioFacMail.Height
            w = FrameEnvioFacMail.Width
            PonerFrameVisible FrameEnvioFacMail, True, h, w

    End Select
    
    'Esto se consigue poneinedo el cancel en el opcion k corresponda
    Me.cmdCancel(indFrame).Cancel = True
    Me.Width = w + 70
    Me.Height = h + 350
    
End Sub

Private Sub frmCtas_DatoSeleccionado(CadenaSeleccion As String)
    txtCodigo(indCodigo).Text = RecuperaValor(CadenaSeleccion, 1) 'codmacta
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2) 'nommacta
End Sub

Private Sub frmSec_DatoSeleccionado(CadenaSeleccion As String)
Dim Cad As String
    txtCodigo(indCodigo).Text = RecuperaValor(CadenaSeleccion, 1) 'codsecci
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2) 'nomsecci
    
    Cad = RecuperaValor(CadenaSeleccion, 5)  'numconta
    If Cad <> "" Then BdConta = CByte(Cad)  'numero de conta
End Sub

Private Sub frmF_Selec(vFecha As Date)
'Calendario de Fecha
    txtCodigo(indCodigo).Text = Format(vFecha, "dd/mm/yyyy")
End Sub


Private Sub frmMen_DatoSeleccionado(CadenaSeleccion As String)
'formulario de mensajes: frmMensajes
    If CadenaSeleccion <> "" Then
        If OpcionListado = 305 Or OpcionListado = 306 Then 'Proveedores
            cadFormula = "{proveedor.codprove} IN [" & CadenaSeleccion & "]"
            cadSelect = "proveedor.codprove IN (" & CadenaSeleccion & ")"
        Else 'clientes
            cadFormula = "{sclien.codclien} IN [" & CadenaSeleccion & "]"
            cadSelect = "sclien.codclien IN (" & CadenaSeleccion & ")"
        End If
    Else 'no seleccionamos ningun cliente
        cadFormula = ""
        cadSelect = ""
    End If
End Sub


Private Sub imgBuscarOfer_Click(Index As Integer)
    Select Case Index
        Case 8 ' seccion
            AbrirFrmSeccion (Index)
            
        Case 56, 57 'Cod. CLIENTE
            If BdConta = 0 Then Exit Sub
            
            AbrirFrmCuentas (Index + 54)
    End Select
    PonerFoco txtCodigo(indCodigo)
End Sub


Private Sub imgFecha_Click(Index As Integer)
   Screen.MousePointer = vbHourglass
   Set frmF = New frmCal
   
   '++monica
    Dim esq As Long
    Dim dalt As Long
    Dim menu As Long
    Dim obj As Object

    Set frmF = New frmCal
    
    esq = imgFecha(Index).Left
    dalt = imgFecha(Index).Top
    
    Set obj = imgFecha(Index).Container

    While imgFecha(Index).Parent.Name <> obj.Name
        esq = esq + obj.Left
        dalt = dalt + obj.Top
        Set obj = obj.Container
    Wend
    
    menu = Me.Height - Me.ScaleHeight 'ací tinc el heigth del menú i de la toolbar

    frmF.Left = esq + imgFecha(Index).Parent.Left + 30
    frmF.Top = dalt + imgFecha(Index).Parent.Top + imgFecha(Index).Height + menu - 40

   frmF.NovaData = Now
   
   Select Case Index
        Case 1 'frameOfertas (indFrame=6)
            indCodigo = 3 'Desde
        Case 2 'frameOfertas (indFrame=6)
            indCodigo = 4 'Hasta
        Case 3 'frameRecordatorio Oferta
            indCodigo = 7 '(Desde)
        Case 4 'frameRecordatorio Oferta
            indCodigo = 8 '(Hasta)
        Case 5 'frameEfectuadas
            indCodigo = 16 'Desde
        Case 6 'frameEfectuadas
            indCodigo = 17 'Hasta
        Case 7 'frameTraspasoHco
            indCodigo = 22 'Desde
        Case 8 'frameTraspasoHco
            indCodigo = 23 'hasta
        Case 9, 10 'FrameGenerarPedido
            indCodigo = Index + 16
        Case 11, 12 'Frame Clientes Inactivos
            indCodigo = 20 + Index
        Case 13 'frame pasar pedido a Albaran de compras (a proveedor)
            indCodigo = 49
        Case 14
            indCodigo = 50
        Case 15, 16
            indCodigo = Index + 54
        Case 17 'Frame Factura Rectificariva
            indCodigo = 72
        Case 18, 19 'Ped. Compras
            indCodigo = Index + 56
        Case 20, 21 'Carta Pedidos
            indCodigo = Index + 57
        Case 22: indCodigo = Index + 60
        Case 23, 24 'Reimprimir facturas
            indCodigo = Index + 62
        Case 25, 26 'Cierre caja TPV
            indCodigo = Index + 63
        Case 27, 28 'Listados estadistica compras
            indCodigo = Index + 65
        Case 29, 30 'Estadistica ventas por familia
            indCodigo = Index + 69
   
        Case 31, 32 'Impresion etiq. clientes. Desde / hasta factura
            indCodigo = Index + 73
        Case 33, 34
            indCodigo = Index + 75
   End Select
   
   
   PonerFormatoFecha txtCodigo(indCodigo)
   If txtCodigo(indCodigo).Text <> "" Then frmF.NovaData = CDate(txtCodigo(indCodigo).Text)

   Screen.MousePointer = vbDefault
   frmF.Show vbModal
   Set frmF = Nothing
   PonerFoco txtCodigo(indCodigo)
End Sub

Private Sub Text1_GotFocus(Index As Integer)
    ConseguirFoco Text1(Index), 3
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub txtCodigo_GotFocus(Index As Integer)
    ConseguirFoco txtCodigo(Index), 3
End Sub

Private Sub txtCodigo_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If Index = 33 And KeyCode = 38 Then Exit Sub
    KEYdown KeyCode
End Sub

Private Sub txtCodigo_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = teclaBuscar Then
        Select Case Index
            Case 110: KEYBusqueda KeyAscii, 56 'cuenta desde
            Case 111: KEYBusqueda KeyAscii, 57 'cuenta hasta
            Case 108: KEYFecha KeyAscii, 33 'fecha desde
            Case 109: KEYFecha KeyAscii, 34 'fecha hasta
            Case 8: KEYBusqueda KeyAscii, 8 'seccion
        End Select
    Else
        KEYpress KeyAscii
    End If
End Sub

Private Sub KEYBusqueda(KeyAscii As Integer, indice As Integer)
    KeyAscii = 0
    imgBuscarOfer_Click (indice)
End Sub

Private Sub KEYFecha(KeyAscii As Integer, indice As Integer)
    KeyAscii = 0
    imgFecha_Click (indice)
End Sub


Private Sub txtCodigo_LostFocus(Index As Integer)
Dim tabla As String
Dim codCampo As String, nomcampo As String
Dim TipCampo As String, Formato As String
Dim Titulo As String
Dim EsNomCod As Boolean
Dim Cad As String

    'Quitar espacios en blanco por los lados
    txtCodigo(Index).Text = Trim(txtCodigo(Index).Text)
    
    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub
    
    EsNomCod = False
    TipCampo = "N" 'Casi todos son numericos
    
    Select Case Index
        'FECHA Desde Hasta
        Case 108, 109
            If txtCodigo(Index).Text = "" Then Exit Sub
            PonerFormatoFecha txtCodigo(Index)
            
        Case 106, 107 'Nº de OFERTA/FACTURA
            If PonerFormatoEntero(txtCodigo(Index)) Then
                txtCodigo(Index).Text = Format(txtCodigo(Index).Text, "0000000")
            End If
        
        Case 8 'Seccion
            BdConta = 0
            txtNombre(Index).Text = PonerNombreDeCod(txtCodigo(Index), "seccion", "nomsecci", "codsecci", "N")
            If txtCodigo(Index).Text <> "" Then
                txtCodigo(Index).Text = Format(txtCodigo(Index).Text, "000")
            
                Cad = DevuelveDesdeBDNew(cPTours, "seccion", "numconta", "codsecci", txtCodigo(8).Text, "N") 'numconta
                If Cad <> "" Then BdConta = CByte(Cad)  'numero de conta
            Else
                MsgBox "Debe introducir un código existente en la sección.", vbExclamation
            End If
                    
            
        Case 110, 111 'cuenta contable
            If txtCodigo(Index).Text = "" Then Exit Sub
            
            If BdConta = 0 Then Exit Sub
            
            If AbrirConexionContaFac(vParamAplic.UsuarioContaFac, vParamAplic.PasswordContaFac, BdConta) Then
                Set vEmpresaFac = New CempresaFac
                If vEmpresaFac.LeerNiveles Then
                    txtNombre(Index).Text = PonerNombreCuenta(txtCodigo(Index), 0, , BdConta, True)
                End If
                Set vEmpresaFac = Nothing
                CerrarConexionContaFac
            End If

    End Select
    
    If EsNomCod Then
        If TipCampo = "N" Then
            If PonerFormatoEntero(txtCodigo(Index)) Then
                txtNombre(Index).Text = PonerNombreDeCod(txtCodigo(Index), tabla, nomcampo, codCampo, TipCampo)
                If txtCodigo(Index).Text <> "" Then txtCodigo(Index).Text = Format(txtCodigo(Index).Text, Formato)
            Else
                txtNombre(Index).Text = ""
            End If
        Else
            txtNombre(Index).Text = PonerNombreDeCod(txtCodigo(Index), tabla, nomcampo, codCampo, TipCampo)
        End If
    End If
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

Private Sub LlamarImprimir()
     With frmImprimir
        .FormulaSeleccion = cadFormula
        .OtrosParametros = cadParam
        .NumeroParametros = numParam
        .SoloImprimir = False
        .EnvioEMail = False
        .Opcion = OpcionListado
        .Titulo = Titulo
        .NombreRPT = nomRPT
        .ConSubInforme = conSubRPT
        .Show vbModal
    End With
End Sub


Private Sub CargarIconos()
Dim i As Integer
    
    For i = 8 To 8
        Me.imgBuscarOfer(i).Picture = frmPpal.imgListImages16.ListImages(1).Picture
    Next i
    For i = 56 To 57
        Me.imgBuscarOfer(i).Picture = frmPpal.imgListImages16.ListImages(1).Picture
    Next i
End Sub


Private Sub BorrarTempInformes()
Dim Sql As String

    On Error GoTo EBorrar
    
    Sql = "DELETE FROM tmpinformes WHERE codusu=" & vSesion.Codigo
    conn.Execute Sql
    
EBorrar:
    If Err.Number <> 0 Then Err.Clear
End Sub


Private Sub AbrirFrmCuentas(indice As Integer)
    If BdConta = 0 Then Exit Sub
    
    If AbrirConexionContaFac(vParamAplic.UsuarioContaFac, vParamAplic.PasswordContaFac, BdConta) Then
        Set vEmpresaFac = New CempresaFac
        If vEmpresaFac.LeerNiveles Then
            indCodigo = indice
            Set frmCtas = New frmCtasConta
            frmCtas.Conexion = BdConta
            frmCtas.Facturas = True
            frmCtas.CadBusqueda = DevuelveDesdeBDNew(cPTours, "seccion", "raizcta", "codsecci", txtCodigo(8).Text, "N")
            frmCtas.NumDigit = DevuelveDesdeBDNewFac("empresa", "numdigi" & vEmpresaFac.numNivel, "", "", "")
            frmCtas.DatosADevolverBusqueda = "0|1|"
            frmCtas.CodigoActual = txtCodigo(indice).Text
            frmCtas.Show vbModal
            Set frmCtas = Nothing
            PonerFoco txtCodigo(indice)
        End If
        Set vEmpresaFac = Nothing
        CerrarConexionContaFac
    End If
End Sub
 
Private Sub AbrirFrmSeccion(indice As Integer)
    indCodigo = indice
    Set frmSec = New frmManSecciones
    frmSec.DatosADevolverBusqueda = "0|1|2|3|4|"
    frmSec.CodigoActual = txtCodigo(8).Text
    frmSec.Show vbModal
    Set frmSec = Nothing
End Sub
 

