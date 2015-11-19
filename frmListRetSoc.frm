VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmListRetSoc 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Informes"
   ClientHeight    =   5205
   ClientLeft      =   45
   ClientTop       =   2430
   ClientWidth     =   6720
   Icon            =   "frmListRetSoc.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5205
   ScaleWidth      =   6720
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog cd1 
      Left            =   7680
      Top             =   5160
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame FrameGrabacionModelos 
      Height          =   5145
      Left            =   0
      TabIndex        =   19
      Top             =   0
      Width           =   6675
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   5
         Left            =   1710
         MaxLength       =   10
         TabIndex        =   29
         Tag             =   "Código Postal|T|S|||clientes|codposta|||"
         Top             =   1155
         Width           =   1050
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   6
         Left            =   3825
         MaxLength       =   10
         TabIndex        =   30
         Tag             =   "Código Postal|T|S|||clientes|codposta|||"
         Top             =   1155
         Width           =   1050
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   7
         Left            =   2835
         Locked          =   -1  'True
         TabIndex        =   28
         Top             =   1785
         Width           =   3180
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   7
         Left            =   1710
         MaxLength       =   10
         TabIndex        =   31
         Tag             =   "Cta.Contable|T|S|||sparam|ctaconta|||"
         Top             =   1785
         Width           =   1080
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   8
         Left            =   2835
         Locked          =   -1  'True
         TabIndex        =   27
         Top             =   2145
         Width           =   3180
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   8
         Left            =   1710
         MaxLength       =   10
         TabIndex        =   32
         Tag             =   "Cta.Contable|T|S|||sparam|ctaconta|||"
         Top             =   2145
         Width           =   1080
      End
      Begin VB.Frame FrameContacto 
         Caption         =   "Persona de Contacto"
         ForeColor       =   &H00972E0B&
         Height          =   915
         Left            =   390
         TabIndex        =   20
         Top             =   3105
         Width           =   5865
         Begin VB.TextBox txtCodigo 
            Height          =   285
            Index           =   9
            Left            =   150
            MaxLength       =   40
            TabIndex        =   34
            Tag             =   "Código Postal|T|S|||clientes|codposta|||"
            Top             =   510
            Width           =   4260
         End
         Begin VB.TextBox txtCodigo 
            Height          =   285
            Index           =   10
            Left            =   4470
            MaxLength       =   9
            TabIndex        =   35
            Tag             =   "Telefono|N|S|||clientes|codposta|000000000||"
            Top             =   510
            Width           =   1050
         End
         Begin VB.Label Label4 
            Caption         =   "Apellidos y Nombre"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   29
            Left            =   210
            TabIndex        =   22
            Top             =   300
            Width           =   2595
         End
         Begin VB.Label Label4 
            Caption         =   "Teléfono"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   36
            Left            =   4530
            TabIndex        =   21
            Top             =   300
            Width           =   705
         End
      End
      Begin VB.CommandButton CmdCancelModelo 
         Caption         =   "&Cancelar"
         Height          =   375
         Left            =   5400
         TabIndex        =   37
         Top             =   4425
         Width           =   975
      End
      Begin VB.CommandButton CmdAcepModelo 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   4140
         TabIndex        =   36
         Top             =   4425
         Width           =   975
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   4
         Left            =   1710
         MaxLength       =   13
         TabIndex        =   33
         Tag             =   "Nro.Justificante|N|S|||clientes|codposta|0000000000000||"
         Top             =   2625
         Width           =   1380
      End
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   90
         Top             =   180
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   30
         Left            =   3690
         MaxLength       =   13
         TabIndex        =   23
         Top             =   3240
         Width           =   1380
      End
      Begin VB.Label Label4 
         Caption         =   "Desde"
         Height          =   195
         Index           =   1
         Left            =   810
         TabIndex        =   43
         Top             =   1170
         Width           =   465
      End
      Begin VB.Label Label4 
         Caption         =   "Fecha Factura"
         ForeColor       =   &H00972E0B&
         Height          =   255
         Index           =   8
         Left            =   450
         TabIndex        =   42
         Top             =   900
         Width           =   1815
      End
      Begin VB.Label Label4 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   6
         Left            =   3000
         TabIndex        =   41
         Top             =   1170
         Width           =   420
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   5
         Left            =   1410
         Picture         =   "frmListRetSoc.frx":000C
         ToolTipText     =   "Buscar fecha"
         Top             =   1140
         Width           =   240
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   6
         Left            =   3540
         Picture         =   "frmListRetSoc.frx":0097
         ToolTipText     =   "Buscar fecha"
         Top             =   1170
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Cuenta Contable"
         ForeColor       =   &H00972E0B&
         Height          =   195
         Index           =   1
         Left            =   450
         TabIndex        =   40
         Top             =   1515
         Width           =   1395
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   3
         Left            =   1395
         ToolTipText     =   "Buscar Cuenta Contable"
         Top             =   2160
         Width           =   240
      End
      Begin VB.Label Label4 
         Caption         =   "Desde"
         Height          =   195
         Index           =   3
         Left            =   810
         TabIndex        =   39
         Top             =   1785
         Width           =   465
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   2
         Left            =   1395
         ToolTipText     =   "Buscar Cuenta Contable"
         Top             =   1800
         Width           =   240
      End
      Begin VB.Label Label4 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   4
         Left            =   810
         TabIndex        =   38
         Top             =   2145
         Width           =   465
      End
      Begin VB.Label Label9 
         Caption         =   "Grabación Modelo 190"
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
         Left            =   420
         TabIndex        =   25
         Top             =   270
         Width           =   5160
      End
      Begin VB.Label Label4 
         Caption         =   "Nro.Justific."
         ForeColor       =   &H00972E0B&
         Height          =   195
         Index           =   28
         Left            =   420
         TabIndex        =   24
         Top             =   2655
         Width           =   945
      End
   End
   Begin VB.Frame FrameCobros 
      Height          =   5130
      Left            =   0
      TabIndex        =   8
      Top             =   45
      Width           =   6690
      Begin VB.CheckBox Check1 
         Caption         =   "Detalle de Facturas"
         Height          =   240
         Index           =   1
         Left            =   315
         TabIndex        =   4
         Top             =   3420
         Width           =   2220
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Saltar página por Socio"
         Height          =   240
         Index           =   0
         Left            =   315
         TabIndex        =   5
         Top             =   3870
         Width           =   2220
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   3
         Left            =   1530
         MaxLength       =   10
         TabIndex        =   3
         Tag             =   "Código Postal|T|S|||clientes|codposta|||"
         Top             =   2730
         Width           =   1050
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   2
         Left            =   1530
         MaxLength       =   10
         TabIndex        =   2
         Tag             =   "Código Postal|T|S|||clientes|codposta|||"
         Top             =   2400
         Width           =   1050
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Left            =   5220
         TabIndex        =   7
         Top             =   4425
         Width           =   975
      End
      Begin VB.CommandButton cmdAceptar 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   4005
         TabIndex        =   6
         Top             =   4425
         Width           =   975
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   0
         Left            =   1530
         MaxLength       =   10
         TabIndex        =   0
         Top             =   1290
         Width           =   1230
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   1
         Left            =   1530
         MaxLength       =   10
         TabIndex        =   1
         Top             =   1665
         Width           =   1230
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   0
         Left            =   2835
         Locked          =   -1  'True
         TabIndex        =   15
         Text            =   "Text5"
         Top             =   1305
         Width           =   3360
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   1
         Left            =   2835
         Locked          =   -1  'True
         TabIndex        =   16
         Text            =   "Text5"
         Top             =   1665
         Width           =   3360
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   1710
         Left            =   2880
         TabIndex        =   17
         Top             =   2385
         Width           =   3330
         _ExtentX        =   5874
         _ExtentY        =   3016
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         HideColumnHeaders=   -1  'True
         Checkboxes      =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Appearance      =   1
         NumItems        =   0
      End
      Begin VB.Label Label1 
         Caption         =   "Listado de Retenciones"
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
         Index           =   0
         Left            =   315
         TabIndex        =   26
         Top             =   360
         Width           =   5160
      End
      Begin VB.Image Image1 
         Height          =   240
         Index           =   0
         Left            =   5715
         Picture         =   "frmListRetSoc.frx":0122
         ToolTipText     =   "Marcar todos"
         Top             =   2070
         Width           =   240
      End
      Begin VB.Image Image1 
         Height          =   240
         Index           =   1
         Left            =   5955
         Picture         =   "frmListRetSoc.frx":6974
         ToolTipText     =   "Desmarcar todos"
         Top             =   2070
         Width           =   240
      End
      Begin VB.Label Label4 
         Caption         =   "Tipo de Factura"
         ForeColor       =   &H00972E0B&
         Height          =   255
         Index           =   0
         Left            =   2880
         TabIndex        =   18
         Top             =   2115
         Width           =   1815
      End
      Begin VB.Label Label4 
         Caption         =   "Fecha"
         ForeColor       =   &H00972E0B&
         Height          =   255
         Index           =   16
         Left            =   285
         TabIndex        =   14
         Top             =   2130
         Width           =   1815
      End
      Begin VB.Label Label4 
         Caption         =   "Desde"
         Height          =   195
         Index           =   15
         Left            =   645
         TabIndex        =   13
         Top             =   2370
         Width           =   465
      End
      Begin VB.Label Label4 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   14
         Left            =   645
         TabIndex        =   12
         Top             =   2730
         Width           =   420
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   2
         Left            =   1215
         Picture         =   "frmListRetSoc.frx":7376
         ToolTipText     =   "Buscar fecha"
         Top             =   2370
         Width           =   240
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   3
         Left            =   1215
         Picture         =   "frmListRetSoc.frx":7401
         ToolTipText     =   "Buscar fecha"
         Top             =   2730
         Width           =   240
      End
      Begin VB.Label Label4 
         Caption         =   "Desde"
         Height          =   195
         Index           =   13
         Left            =   645
         TabIndex        =   11
         Top             =   1290
         Width           =   465
      End
      Begin VB.Label Label4 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   12
         Left            =   645
         TabIndex        =   10
         Top             =   1665
         Width           =   420
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Cuenta Contable"
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
         Left            =   285
         TabIndex        =   9
         Top             =   1035
         Width           =   1215
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   0
         Left            =   1245
         MouseIcon       =   "frmListRetSoc.frx":748C
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar Cuenta Contable"
         Top             =   1290
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   1
         Left            =   1215
         MouseIcon       =   "frmListRetSoc.frx":75DE
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar Cuenta Contable"
         Top             =   1665
         Width           =   240
      End
   End
End
Attribute VB_Name = "frmListRetSoc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' +-+-+-+-+-+-+-+-+-+-+-
' +-+- Autor MANOLO +-+-
' +-+-+-+-+-+-+-+-+-+-+-

Option Explicit

Public OpcionListado As Byte
' =====FACTURAS SOCIOS====
' 0 = Listado de Retenciones
' 1 = Grabacion del Modelo 190
' 2 = Grabación del Modelo 346 ( de momento no )
                        

Public CadTag As String 'Cadena con el Tag del campo que se va a poner en D/H en los listados
                        'Se necesita si el tipo de codigo es texto

    
Private Conexion As Byte
'1.- Conexión a BD Ariges  2.- Conexión a BD Conta

Private HaDevueltoDatos As Boolean

'Private WithEvents frmcli As frmManClien 'Clientes
'Private WithEvents frmCol As frmManCoope 'Colectivo
Private WithEvents frmC As frmCal 'calendario fechas
Attribute frmC.VB_VarHelpID = -1
Private WithEvents frmCtas As frmCtasConta
Attribute frmCtas.VB_VarHelpID = -1

'GENERALES PARA PASARLE A CRYSTAL REPORT
Private cadFormula As String 'Cadena con la FormulaSelection para Crystal Report
Private cadParam As String 'Cadena con los parametros para Crystal Report
Private numParam As Byte 'Numero de parametros que se pasan a Crystal Report
Private cadSelect As String 'Cadena para comprobar si hay datos antes de abrir Informe
Private cadTitulo As String 'Titulo para la ventana frmImprimir
Private cadNombreRPT As String 'Nombre del informe

Dim indCodigo As Integer 'indice para txtCodigo
Dim indFrame As Single 'nº de frame en el que estamos

Dim BdConta As Integer ' numero de la contabilidad donde se hace conexion

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


Private Sub Check1_Click(Index As Integer)
'    If Index = 1 Then
'        If Check1(Index).Value = 0 Then
'            Check1(0).Value = 0
'            Check1(0).Enabled = False
'        Else
'            Check1(0).Enabled = True
'        End If
'    End If
End Sub

Private Sub CmdAcepModelo_Click()
Dim cDesde As String, cHasta As String 'cadena codigo Desde/Hasta
Dim nDesde As String, nHasta As String 'cadena Descripcion Desde/Hasta
Dim i As Byte
Dim nTabla As String

Dim vWhere As String
Dim b As Boolean
Dim tipo As Byte
Dim Fecfin As String
Dim FecIni As String


    InicializarVbles
    
    If Not DatosOk Then Exit Sub


    'D/H Socios
    cDesde = Trim(txtCodigo(7).Text)
    cHasta = Trim(txtCodigo(8).Text)
    nDesde = txtNombre(7).Text
    nHasta = txtNombre(8).Text
    If Not (cDesde = "" And cHasta = "") Then
        'Cadena para seleccion Desde y Hasta
        Codigo = "{" & tabla & ".codmacta}"
        TipCod = "T"
        If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHCliente= """) Then Exit Sub
    End If
    
    'D/H Fecha factura
    cDesde = Trim(txtCodigo(5).Text)
    cHasta = Trim(txtCodigo(6).Text)
    If Not (cDesde = "" And cHasta = "") Then
        'Cadena para seleccion Desde y Hasta
        Codigo = "{" & tabla & ".fecfactu}"
        TipCod = "F"
        If Not PonerDesdeHasta(cDesde, cHasta, "", "", "pDHFecha= """) Then Exit Sub
    End If
   
    nTabla = "factsocio"
    
    FecIni = CDate(txtCodigo(5).Text)
    
    txtCodigo(30).Text = Format(Year(FecIni), "0000") ' inicio del año natural
    
    
    Select Case OpcionListado
        Case 1 'modelo 190
            If Not AnyadirAFormula(cadFormula, "{factsocio.impreten} <> 0") Then Exit Sub
            If Not AnyadirAFormula(cadSelect, "{factsocio.impreten} <> 0") Then Exit Sub
            
            If Not AnyadirAFormula(cadFormula, "{factsocio.tipofact} in [0,1,2,3]") Then Exit Sub
            If Not AnyadirAFormula(cadSelect, "{factsocio.tipofact} in (0,1,2,3)") Then Exit Sub
        
        Case 2 'modelo 346
            ' seleccionamos tipodocu: 5 = subvencion
            '                         6 = siniestro
            If Not AnyadirAFormula(cadFormula, "{stipom.tipodocu} in [4,5]") Then Exit Sub
            If Not AnyadirAFormula(cadSelect, "{stipom.tipodocu} in (4,5)") Then Exit Sub
    
            If Not AnyadirAFormula(cadFormula, "{rfactsoc_variedad.imporvar} <> 0") Then Exit Sub
            If Not AnyadirAFormula(cadSelect, "{rfactsoc_variedad.imporvar} <> 0") Then Exit Sub
            
            nTabla = "(" & nTabla & ") INNER JOIN rfactsoc_variedad ON rfactsoc.codtipom = rfactsoc_variedad.codtipom "
            nTabla = nTabla & " and rfactsoc.numfactu = rfactsoc_variedad.numfactu "
            nTabla = nTabla & " and rfactsoc.fecfactu = rfactsoc_variedad.fecfactu "
            
    End Select
    
    If HayRegParaInforme(nTabla, cadSelect) Then
        b = GeneraFicheroModelo(OpcionListado, nTabla, cadSelect)
        If b Then
            If CopiarFichero Then
                MsgBox "Proceso realizado correctamente", vbExclamation
                CmdCancelModelo_Click
            End If
        End If
    End If
        
End Sub


Private Sub cmdAceptar_Click()
Dim cDesde As String, cHasta As String 'cadena codigo Desde/Hasta
Dim nDesde As String, nHasta As String 'cadena Descripcion Desde/Hasta
Dim cadTABLA As String, cOrden As String
Dim i As Byte
Dim b As Boolean
Dim Tipos As String

    InicializarVbles
    
    
    '========= PARAMETROS  =============================
    'Añadir el parametro de Empresa
    cadParam = cadParam & "|pEmpresa=""" & vEmpresa.nomEmpre & """|"
    numParam = numParam + 1
    
    'D/H Cuenta contable
    cDesde = Trim(txtCodigo(0).Text)
    cHasta = Trim(txtCodigo(1).Text)
    nDesde = txtNombre(0).Text
    nHasta = txtNombre(1).Text
    If Not (cDesde = "" And cHasta = "") Then
        'Cadena para seleccion Desde y Hasta
        Codigo = "{" & tabla & ".codmacta}"
        TipCod = "T"
        If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHcliente= """) Then Exit Sub
    End If
    
    'D/H Fecha
    cDesde = Trim(txtCodigo(2).Text)
    cHasta = Trim(txtCodigo(3).Text)
    If Not (cDesde = "" And cHasta = "") Then
        'Cadena para seleccion Desde y Hasta
        Codigo = "{" & tabla & ".fecfactu}"
        TipCod = "F"
        If Not PonerDesdeHasta(cDesde, cHasta, "", "", "pDHfechaFac= """) Then Exit Sub
    End If
    
    'saltar pagina por socio
    If CBool(Me.Check1(0).Value) Then
        cadParam = cadParam & "pSaltaPagina=1|"
    Else
        cadParam = cadParam & "pSaltaPagina=0|"
    End If
    numParam = numParam + 1
    
    
    'resumen o no
    If CBool(Me.Check1(1).Value) Then
        cadParam = cadParam & "pResumen=1|"
    Else
        cadParam = cadParam & "pResumen=0|"
    End If
    numParam = numParam + 1
    

    'cargamos en la formula que tipo de factura vamos a seleccionar
    Tipos = ""
    For i = 1 To 6
        If ListView1.ListItems(i).Checked Then
            Tipos = Tipos & i - 1 & ","
        End If
    Next i
    
    If Tipos = "" Then
        MsgBox "Debe seleccionar al menos un tipo de factura.", vbExclamation
        Exit Sub
    Else
        ' quitamos la ultima coma
        Tipos = "{factsocio.tipofact} in (" & Mid(Tipos, 1, Len(Tipos) - 1) & ")"
        If Not AnyadirAFormula(cadSelect, Tipos) Then Exit Sub
        Tipos = Replace(Replace(Tipos, "(", "["), ")", "]")
        If Not AnyadirAFormula(cadFormula, Tipos) Then Exit Sub
    End If
        
    
    cadTABLA = "factsocio"
    If HayRegParaInforme(cadTABLA, cadSelect) Then
       cadTitulo = "Listado de Retenciones a Socios"
       
       cadNombreRPT = "rInfRetSocios.rpt"
       LlamarImprimir
    End If
    
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub CmdCancelModelo_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
    If PrimeraVez Then
        PrimeraVez = False
        BdConta = 0
        
        Select Case OpcionListado
            Case 0
                PonerFoco txtCodigo(0)
            Case 1
                PonerFoco txtCodigo(5)
        End Select
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
Dim h As Integer, w As Integer
Dim List As Collection
Dim i As Integer

    PrimeraVez = True
    Limpiar Me

    'IMAGES para busqueda
    For i = 0 To 3
        Me.imgBuscar(i).Picture = frmPpal.imgListImages16.ListImages(1).Picture
    Next i

    Me.FrameCobros.visible = False
    Me.FrameGrabacionModelos.visible = False
         
         
    Select Case OpcionListado
        Case 0
            CargarListView
                 
            FrameCobrosVisible True, h, w
        
            Check1(1).Value = 1
        
        Case 1
            FrameGrabacionModelosVisible True, h, w
        
            txtCodigo(5).Text = "01/01/" & Format(Year(Now) - 1, "0000")
            txtCodigo(6).Text = "31/12/" & Format(Year(Now) - 1, "0000")
    End Select
    indFrame = 5
    tabla = "factsocio"
            
            
    'Esto se consigue poneinedo el cancel en el opcion k corresponda
    Me.cmdCancel.Cancel = True

End Sub

Private Sub frmC_Selec(vFecha As Date)
 'Fecha
    txtCodigo(CByte(imgFec(2).Tag)).Text = Format(vFecha, "dd/MM/yyyy")
End Sub

Private Sub frmCtas_DatoSeleccionado(CadenaSeleccion As String)
    txtCodigo(indCodigo).Text = RecuperaValor(CadenaSeleccion, 1) 'codmacta
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2) 'nommacta
End Sub


Private Sub Image1_Click(Index As Integer)
Dim i As Integer
Dim TotalCant As Currency
Dim TotalImporte As Currency

    Screen.MousePointer = vbHourglass
    
    Select Case Index
        Case 0
            For i = 1 To ListView1.ListItems.Count
                ListView1.ListItems(i).Checked = True
            Next i
        Case 1
            For i = 1 To ListView1.ListItems.Count
                ListView1.ListItems(i).Checked = False
            Next i
    End Select
    
    Screen.MousePointer = vbDefault

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

Private Sub imgBuscar_Click(Index As Integer)
   Select Case Index
        Case 0, 1 'Cuentas
            AbrirFrmCuentas (Index)
        Case 2, 3 'cuentas de modelo 190
            AbrirFrmCuentas (Index + 5)
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
'15/02/2007
'    KEYpress KeyAscii
'ahora
    If KeyAscii = teclaBuscar Then
        Select Case Index
            Case 0: KEYBusqueda KeyAscii, 0 'cuenta desde
            Case 1: KEYBusqueda KeyAscii, 1 'cuenta hasta
            Case 7: KEYBusqueda KeyAscii, 2 'cuenta desde
            Case 8: KEYBusqueda KeyAscii, 3 'cuenta hasta
            Case 2: KEYFecha KeyAscii, 3 'fecha desde
            Case 3: KEYFecha KeyAscii, 3 'fecha hasta
            Case 5: KEYFecha KeyAscii, 5 'fecha desde
            Case 6: KEYFecha KeyAscii, 6 'fecha hasta
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
        Case 0, 1, 7, 8 'Cuenta Cliente
            If txtCodigo(Index).Text = "" Then Exit Sub
            txtNombre(Index).Text = PonerNombreCuenta(txtCodigo(Index), 0, , cContaFacSoc, False)
            'DevuelveDesdeBDNew(cContaFacSoc, "cuentas", "nommacta", "codmacta", txtCodigo(Index).Text, "T")
            
        Case 2, 3, 5, 6 'FECHAS
            If txtCodigo(Index).Text <> "" Then PonerFormatoFecha txtCodigo(Index)
            
        Case 4, 10 'Justificante y Telefono de mod.190
            If txtCodigo(Index).Text <> "" Then PonerFormatoEntero txtCodigo(Index)
            
    End Select
End Sub

Private Sub FrameCobrosVisible(visible As Boolean, ByRef h As Integer, ByRef w As Integer)
    Me.FrameCobros.visible = visible
    If visible = True Then
        Me.FrameCobros.Top = -90
        Me.FrameCobros.Left = 0
        Me.FrameCobros.Height = 5595
        Me.FrameCobros.Width = 6810
        w = Me.FrameCobros.Width
        h = Me.FrameCobros.Height
    End If
End Sub

Private Sub FrameGrabacionModelosVisible(visible As Boolean, ByRef h As Integer, ByRef w As Integer)
'Frame para la grabacion de modelos
    Me.FrameGrabacionModelos.visible = visible
    If visible = True Then
        Me.FrameGrabacionModelos.Top = -90
        Me.FrameGrabacionModelos.Left = 0
        Me.FrameGrabacionModelos.Height = 5595
        Me.FrameGrabacionModelos.Width = 6810
        w = Me.FrameGrabacionModelos.Width
        h = Me.FrameGrabacionModelos.Height
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
        .Titulo = cadTitulo
        .NombreRPT = cadNombreRPT
        .Opcion = 1
        .Facturas = False
        .EnvioEMail = False
        .Contabilidad = cContaFacSoc
        .Show vbModal
    End With
End Sub

Private Sub AbrirFrmCuentas(indice As Integer)
            
    indCodigo = indice
    Set frmCtas = New frmCtasConta
    frmCtas.Conexion = cContaFacSoc
    frmCtas.Facturas = False
    frmCtas.CadBusqueda = vParamAplic.RaizCtaFacSoc
    frmCtas.NumDigit = vEmpresaFacSoc.DigitosUltimoNivel
    frmCtas.DatosADevolverBusqueda = "0|1|"
    frmCtas.CodigoActual = txtCodigo(indice).Text
    frmCtas.Show vbModal
    Set frmCtas = Nothing
    PonerFoco txtCodigo(indice)
End Sub
 
 
Private Sub AbrirVisReport()
    Screen.MousePointer = vbHourglass
    CadenaDesdeOtroForm = ""
    With frmVisReport
        .FormulaSeleccion = cadFormula
'        .SoloImprimir = (Me.OptVisualizar(indFrame).Value = 1)
        .OtrosParametros = cadParam
        .NumeroParametros = numParam
        '##descomen
'        .MostrarTree = MostrarTree
'        .Informe = MIPATH & Nombre
'        .InfConta = InfConta
        '##
        
'        If NombreSubRptConta <> "" Then
'            .SubInformeConta = NombreSubRptConta
'        Else
'            .SubInformeConta = ""
'        End If
        '##descomen
'        .ConSubInforme = ConSubInforme
        '##
        .Opcion = ""
'       .ExportarPDF = (chkEMAIL.Value = 1)
        .Show vbModal
    End With
    
'    If Me.chkEMAIL.Value = 1 Then
'    '####Descomentar
'        If CadenaDesdeOtroForm <> "" Then frmEMail.Show vbModal
'    End If
    Unload Me
End Sub

Private Sub AbrirEMail()
    If CadenaDesdeOtroForm <> "" Then frmEMail.Show vbModal
End Sub

Private Sub CargarListView()
'Muestra la lista Detallada de Facturas que dieron error al contabilizar
'en un ListView
Dim RS As ADODB.Recordset
Dim ItmX As ListItem
Dim SQL As String

    On Error GoTo ECargarList

    'Los encabezados
    ListView1.ColumnHeaders.Clear

    ListView1.ColumnHeaders.Add , , "Tipo Factura", 2000
            
    Set ItmX = ListView1.ListItems.Add
    ItmX.Text = "Anticipo"
    Set ItmX = ListView1.ListItems.Add
    ItmX.Text = "Liquidación"
    Set ItmX = ListView1.ListItems.Add
    ItmX.Text = "Retirada"
    Set ItmX = ListView1.ListItems.Add
    ItmX.Text = "Industria"
    Set ItmX = ListView1.ListItems.Add
    ItmX.Text = "Subvención"
    Set ItmX = ListView1.ListItems.Add
    ItmX.Text = "Siniestro"

ECargarList:
    If Err.Number <> 0 Then
        Err.Clear
    End If
End Sub



Private Function GeneraFicheroModelo(tipo As Byte, pTabla As String, pWhere As String) As Boolean
Dim NFic As Integer
Dim Regs As Integer
Dim CodigoOrdenante As String
Dim RS As ADODB.Recordset
Dim Rs4 As ADODB.Recordset
Dim Aux As String
Dim Aux2 As String
Dim Cad As String
Dim Pagos As Boolean
Dim Concepto As Byte
Dim b As Boolean
Dim Nregs As Long
Dim total As Variant
Dim Sql4 As String
Dim cTabla As String
Dim vWhere As String


    On Error GoTo EGen
    GeneraFicheroModelo = False
    
    cTabla = pTabla
    vWhere = pWhere
    
    cTabla = QuitarCaracterACadena(cTabla, "{")
    cTabla = QuitarCaracterACadena(cTabla, "}")
    cTabla = QuitarCaracterACadena(cTabla, "_1")
    If vWhere <> "" Then
        vWhere = QuitarCaracterACadena(vWhere, "{")
        vWhere = QuitarCaracterACadena(vWhere, "}")
        vWhere = QuitarCaracterACadena(vWhere, "_1")
    End If
    
    NFic = FreeFile
    
    Open App.path & "\modelo.txt" For Output As #NFic
    
    Select Case tipo
        Case 1 ' MODELO 190
            Aux = "select count(*), sum(factsocio.basereten), sum(factsocio.impreten) "
            Aux = Aux & " from " & cTabla
            Aux = Aux & " where " & vWhere
                
            Set RS = New ADODB.Recordset
            RS.Open Aux, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                
            'CABECERA
' [Monica] 14/01/2010 : no sé de donde he copiado que habían dos cabeceras
'            Cabecera190a NFic, CLng(DBLet(Rs.Fields(0).Value, "N"))
            Cabecera190b NFic, CLng(DBLet(RS.Fields(0).Value, "N")), CCur(DBLet(RS.Fields(1).Value, "N")), CCur(DBLet(RS.Fields(2).Value, "N"))
            
            Set RS = Nothing
            
            'Imprimimos las lineas
            Aux = "select factsocio.codmacta, sum(factsocio.basereten), sum(factsocio.impreten) "
            Aux = Aux & " from " & cTabla
            Aux = Aux & " where " & vWhere
            Aux = Aux & " group by 1 "
            Aux = Aux & " having sum(factsocio.basereten) <> 0 "
            Aux = Aux & " order by 1 "
            
            Set RS = New ADODB.Recordset
            RS.Open Aux, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            
            If RS.EOF Then
                'No hayningun registro
            Else
                b = True
                Regs = 0
                While Not RS.EOF And b
                    Regs = Regs + 1
                    
                    Sql4 = "select nifdatos, nommacta, codposta from cuentas where codmacta = " & DBLet(RS!Codmacta, "T")
                    Set Rs4 = New ADODB.Recordset
                    
                    Rs4.Open Sql4, ConnContaFacSoc, adOpenForwardOnly, adLockPessimistic, adCmdText
                    
                    If Not Rs4.EOF Then
                        Linea190 NFic, Rs4, RS
                    Else
                        b = False
                    End If
                    Set Rs4 = Nothing
                    
                    RS.MoveNext
                Wend
            End If
            RS.Close
            Set RS = Nothing
            
'        Case 2 ' MODELO 346
''            cTabla = "(" & cTabla & ") INNER JOIN variedades ON rfactsoc_variedad.codvarie = variedades.codvarie "
''            cTabla = "(" & cTabla & ") INNER JOIN productos ON variedades.codprodu = productos.codprodu "
''            cTabla = "(" & cTabla & ") INNER JOIN grupopro ON productos.codgrupo = grupopro.codgrupo "
''
''            Aux = "select rfactsoc.codsocio, grupopro.codgrupo, sum(rfactsoc_variedad.imporvar) "
''            Aux = Aux & " from " & cTabla
''            Aux = Aux & " where " & vWhere & " and grupopro.codgrupo in (4,5) " ' algarrobos y olivos
''            Aux = Aux & " group by rfactsoc.codsocio, grupopro.codgrupo "
''            Aux = Aux & "  union "
''            Aux = Aux & " select rfactsoc.codsocio, 0, sum(rfactsoc_variedad.imporvar) "
''            Aux = Aux & " from " & cTabla
''            Aux = Aux & " where " & vWhere & " and not grupopro.codgrupo in (4,5) " ' el resto
''            Aux = Aux & " group by rfactsoc.codsocio, grupopro.codgrupo "
''            Aux = Aux & " order by 1,2"
'
'            Aux = "select tmp346.codsocio, tmp346.codgrupo, sum(tmp346.importe) "
'            Aux = Aux & " from tmp346 "
'            Aux = Aux & " where " & vWhere & " and tmp346.codgrupo in (4,5) " ' algarrobos y olivos
'            Aux = Aux & " group by tmp346.codsocio, tmp346.codgrupo "
'            Aux = Aux & "  union "
'            Aux = Aux & " select tmp346.codsocio, 0, sum(tmp346.importe) "
'            Aux = Aux & " from tmp346 "
'            Aux = Aux & " where " & vWhere & " and not tmp346.codgrupo in (4,5) " ' el resto
'            Aux = Aux & " group by tmp346.codsocio, tmp346.codgrupo "
'            Aux = Aux & " order by 1,2"
'
'
'
'            Nregs = TotalRegistrosConsulta(Aux)
'
'            If Nregs <> 0 Then
'                Aux2 = "select sum(tmp346.importe) from tmp346 "
'                Aux2 = Aux2 & " where " & vWhere
'
'                total = DevuelveValor(Aux2)
'
'                Cabecera346 NFic, Nregs, CCur(total)
'
'                Set RS = New ADODB.Recordset
'                RS.Open Aux, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
'
'                If RS.EOF Then
'                    'No hayningun registro
'                Else
'                    b = True
'                    Regs = 0
'                    While Not RS.EOF And b
'                        Regs = Regs + 1
'                        Set vSocio = New CSocio
'
'                        If vSocio.LeerDatos(DBLet(RS!Codsocio, "N")) Then
'                            Linea346 NFic, vSocio, RS
'                        Else
'                            b = False
'                        End If
'
'                        Set vSocio = Nothing
'                        RS.MoveNext
'                    Wend
'                End If
'                RS.Close
'                Set RS = Nothing
'
'            End If
    End Select
    Close (NFic)
    
    If Regs > 0 Then GeneraFicheroModelo = True
    Exit Function
    
EGen:
    Set RS = Nothing
    Close (NFic)
    MuestraError Err.Number, Err.Description
End Function


Private Sub Cabecera190b(NFich As Integer, Nregs As Currency, ImpReten As Currency, BaseReten As Currency)
Dim Cad As String

'TIPO DE REGISTRO 1:REGISTRO DEL RETENEDOR}
    
    Cad = "1190"                                                  'p.1
    Cad = Cad & Format(txtCodigo(30).Text, "0000")                'p.5 año de ejercicio
    Cad = Cad & RellenaABlancos(vEmpresa.CifEmpresa, True, 9)        'p.9 cif empresa
    Cad = Cad & RellenaABlancos(vEmpresa.nomEmpre, True, 40)   'p.18 nombre de empresa
    Cad = Cad & "D"                                               'p.58
    Cad = Cad & RellenaAceros(txtCodigo(10).Text, True, 9)        'p.59 telefono
    Cad = Cad & RellenaABlancos(txtCodigo(9).Text, True, 40)     'p.68 persona de contacto
    Cad = Cad & RellenaAceros(txtCodigo(4).Text, True, 13)       'p.108 nro de justificante
    Cad = Cad & Space(2)                                          'p.121 ni es complementaria ni sustitutiva
    Cad = Cad & RellenaAceros("0", True, 13)                      'p.123 13 ceros (justificante de la complementaria o sustitutiva)
    Cad = Cad & Format(Nregs, "000000000")                        'p.136 nro de registros

    If BaseReten < 0 Then
        Cad = Cad & "N"                                           'p.145 signo de retenciones
        Cad = Cad & RellenaAceros(ImporteSinFormato(CStr(BaseReten * (-1) * 100)), False, 15)    'p.146
    Else
        Cad = Cad & " "                                           'p.145
        Cad = Cad & RellenaAceros(ImporteSinFormato(CStr(BaseReten * 100)), False, 15)           'p.146
    End If
              
    If ImpReten < 0 Then                                          'p.161
        Cad = Cad & RellenaAceros(ImporteSinFormato(CStr(ImpReten * (-1) * 100)), False, 15)
    Else
        Cad = Cad & RellenaAceros(ImporteSinFormato(CStr(ImpReten * 100)), False, 15)
    End If
    Cad = Cad & Space(322) 'p.176 a 487                    'antes:  Space(62)                             'p.176
    Cad = Cad & Space(3)   'p.488 a 500 firma digital      'antes:  Space(13)                                         'p.238

    Print #NFich, Cad

End Sub


Private Sub Linea190(NFich As Integer, ByRef Rs4 As ADODB.Recordset, ByRef RS As ADODB.Recordset)
Dim Cad As String

    Cad = "2190"                                                'p.1
    Cad = Cad & Format(txtCodigo(30).Text, "0000")              'p.5 año ejercicio
    Cad = Cad & RellenaABlancos(vEmpresa.CifEmpresa, True, 9)     'p.9 cif empresa
    Cad = Cad & RellenaABlancos(Rs4!nifdatos, True, 9)            'p.18 nifsocio
    Cad = Cad & Space(9)                                        'p.27 nif del representante legal
    Cad = Cad & RellenaABlancos(Rs4!Nommacta, True, 40)        'p.36 nombre socio
    Cad = Cad & RellenaABlancos(Mid(Rs4!Codposta, 1, 2), True, 2) 'p.76 codpobla[1,2] codigo de provincia
    Cad = Cad & "H"                                             'p.78 clave de percepcion H=actividades agrícolas, ganaderas y forestales
    Cad = Cad & "01"                                            'p.79 subclave:
'                                                                       01 =  Se consignará esta subclave cuando se trate de percepciones
'                                                                        a las que resulte aplicable el tipo de retención establecido
'                                                                        con carácter general en el artículo 95.4.2º del Reglamento
'                                                                        del Impuesto.
   
'[Monica]: 14/01/2010
' antes no estaba en el if de abajo siempre era un blanco lo he cambiado según el signo.
'    cad = cad & " "                                             'p.81
    
    If DBLet(RS.Fields(1).Value, "N") < 0 Then                  'p.82 base de retencion
        Cad = Cad & "N"                                             'p.81
        Cad = Cad & RellenaAceros(ImporteSinFormato(CStr(DBLet(RS.Fields(1).Value, "N") * (-1) * 100)), False, 13)
    Else
        Cad = Cad & " "                                             'p.81
        Cad = Cad & RellenaAceros(ImporteSinFormato(CStr(DBLet(RS.Fields(1).Value, "N") * 100)), False, 13)
    End If
    
    If DBLet(RS.Fields(2).Value, "N") < 0 Then                  'p.95 importe de retencion
        Cad = Cad & RellenaAceros(ImporteSinFormato(CStr(DBLet(RS.Fields(2).Value, "N") * (-1) * 100)), False, 13)
    Else
        Cad = Cad & RellenaAceros(ImporteSinFormato(CStr(DBLet(RS.Fields(2).Value, "N") * 100)), False, 13)
    End If
    
    Cad = Cad & " "                                             'p.108
    Cad = Cad & RellenaAceros("0", True, 13)                    'p.109
    Cad = Cad & RellenaAceros("0", True, 13)                    'p.122
    Cad = Cad & RellenaAceros("0", True, 13)                    'p.135
    Cad = Cad & RellenaAceros("0", True, 4)                     'p.148
    Cad = Cad & "0"                                             'p.152
    Cad = Cad & RellenaAceros("0", True, 5)                     'p.153
    Cad = Cad & RellenaABlancos(" ", True, 9)                   'p.158
    Cad = Cad & String(88, "0")                                 'p.167  antes eran 84 ceros
    Cad = Cad & Space(246)                                      'p.255 - 500 se rellenan a blancos
    
    Print #NFich, Cad
End Sub


Private Function CopiarFichero() As Boolean
Dim nomFich As String

On Error GoTo ecopiarfichero

    CopiarFichero = False
    ' abrimos el commondialog para indicar donde guardarlo
'    Me.CommonDialog1.InitDir = App.path

    Me.CommonDialog1.DefaultExt = "txt"
    
    CommonDialog1.Filter = "Archivos txt|txt|"
    CommonDialog1.FilterIndex = 1
    
    ' copiamos el primer fichero
    Select Case OpcionListado
        Case 1
            CommonDialog1.FileName = "modelo190.txt"
        Case 2
            CommonDialog1.FileName = "modelo346.txt"
    End Select
        
    Me.CommonDialog1.ShowSave
    
    If CommonDialog1.FileName <> "" Then
        FileCopy App.path & "\modelo.txt", CommonDialog1.FileName
    End If
    
    CopiarFichero = True
    Exit Function

ecopiarfichero:
    If Err.Number <> 0 Then
        MuestraError Err.Number, Err.Description
    End If
    Err.Clear
End Function



Private Function DatosOk() As Boolean
Dim b As Boolean
Dim SQL As String
Dim sql2 As String
Dim vClien As CSocio
' añadido
Dim Mens As String
Dim numfactu As String
Dim numser As String
Dim Fecha As Date
Dim tipoMov As String

    b = True
    Select Case OpcionListado
        Case 1
            '1 - grabacion de modelo 190
            If txtCodigo(5).Text = "" Or txtCodigo(6) = "" Then
                MsgBox "Debe introducir obligatoriamente el rango de fechas.", vbExclamation
                b = False
                PonerFoco txtCodigo(5)
            Else
                If Mid(txtCodigo(5).Text, 7, 4) = Mid(txtCodigo(6).Text, 7, 4) Then
                    If Not (Mid(txtCodigo(5).Text, 1, 5) = "01/01" And Mid(txtCodigo(6).Text, 1, 5) = "31/12") Then
                        MsgBox "El rango de fechas debe de corresponderse con un año natural. Revise.", vbExclamation
                        b = False
                        PonerFoco txtCodigo(5)
                    End If
                Else
                    MsgBox "El rango de fechas debe ser el de año natural. Revise.", vbExclamation
                    b = False
                    PonerFoco txtCodigo(5)
                End If
            End If
    End Select
    DatosOk = b

End Function

