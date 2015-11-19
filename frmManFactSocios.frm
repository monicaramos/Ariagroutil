VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmManFactSocios 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Facturas Socios"
   ClientHeight    =   5820
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9060
   Icon            =   "frmManFactSocios.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5820
   ScaleWidth      =   9060
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Height          =   2550
      Index           =   0
      Left            =   240
      TabIndex        =   20
      Top             =   480
      Width           =   8580
      Begin VB.ComboBox Combo1 
         Height          =   315
         Index           =   1
         ItemData        =   "frmManFactSocios.frx":000C
         Left            =   6435
         List            =   "frmManFactSocios.frx":0016
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Tag             =   "Tipo Factura|N|N|0|5|factsocio|tipofact|||"
         Top             =   1890
         Width           =   1665
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Index           =   0
         ItemData        =   "frmManFactSocios.frx":002C
         Left            =   3690
         List            =   "frmManFactSocios.frx":0036
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Tag             =   "Tipo Socio|N|N|0|2|factsocio|tiposoci|||"
         Top             =   1890
         Width           =   1665
      End
      Begin VB.TextBox text1 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   4
         Left            =   1530
         MaxLength       =   10
         TabIndex        =   4
         Tag             =   "Kilos Facturados|N|N|||factsocio|kilosfac|###,##0||"
         Top             =   1935
         Width           =   1080
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   3
         Left            =   2655
         TabIndex        =   36
         Top             =   1395
         Width           =   5415
      End
      Begin VB.TextBox text1 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   3
         Left            =   1530
         MaxLength       =   10
         TabIndex        =   3
         Tag             =   "Codigo Variedad|N|N|||factsocio|codvarie|000000||"
         Top             =   1395
         Width           =   1080
      End
      Begin VB.CheckBox Check1 
         Height          =   255
         Index           =   1
         Left            =   7785
         TabIndex        =   34
         Tag             =   "Contabilizada|N|N|0|1|factsocio|intconta|||"
         Top             =   405
         Width           =   255
      End
      Begin VB.TextBox text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   2
         Left            =   1530
         MaxLength       =   10
         TabIndex        =   2
         Tag             =   "Cta.Contable|T|N|||factsocio|codmacta||S|"
         Text            =   "1234567890"
         Top             =   1080
         Width           =   1065
      End
      Begin VB.TextBox text1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000013&
         Height          =   285
         Index           =   0
         Left            =   225
         MaxLength       =   7
         TabIndex        =   0
         Tag             =   "Nº de Factura|N|S|0|9999999|factsocio|numfactu|0000000|S|"
         Top             =   585
         Width           =   795
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   2
         Left            =   2655
         TabIndex        =   23
         Top             =   1080
         Width           =   5415
      End
      Begin VB.TextBox text1 
         BackColor       =   &H80000013&
         Height          =   285
         Index           =   1
         Left            =   1530
         MaxLength       =   10
         TabIndex        =   1
         Tag             =   "Fecha Factura|F|N|||factsocio|fecfactu|dd/mm/yyyy|S|"
         Top             =   585
         Width           =   1080
      End
      Begin VB.Label Label2 
         Caption         =   "Tipo Factura"
         Height          =   255
         Index           =   0
         Left            =   5490
         TabIndex        =   43
         Top             =   1935
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "Tipo Socio"
         Height          =   255
         Index           =   1
         Left            =   2790
         TabIndex        =   42
         Top             =   1935
         Width           =   900
      End
      Begin VB.Label Label1 
         Caption         =   "Kilos Facturados"
         Height          =   255
         Index           =   0
         Left            =   225
         TabIndex        =   41
         Top             =   1935
         Width           =   1395
      End
      Begin VB.Label Label1 
         Caption         =   "Variedad"
         Height          =   255
         Index           =   5
         Left            =   225
         TabIndex        =   37
         Top             =   1440
         Width           =   945
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   1
         Left            =   1215
         Tag             =   "-1"
         ToolTipText     =   "Buscar Variedad"
         Top             =   1395
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Contabilizada"
         Height          =   255
         Index           =   7
         Left            =   6750
         TabIndex        =   35
         Top             =   450
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Nº Factura"
         Height          =   255
         Index           =   4
         Left            =   225
         TabIndex        =   24
         Top             =   315
         Width           =   855
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   1
         Left            =   2385
         Picture         =   "frmManFactSocios.frx":004C
         ToolTipText     =   "Buscar fecha"
         Top             =   270
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   0
         Left            =   1215
         Tag             =   "-1"
         ToolTipText     =   "Buscar Cta Contable"
         Top             =   1080
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Fecha Fact."
         Height          =   255
         Index           =   1
         Left            =   1530
         TabIndex        =   22
         Top             =   315
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Cta.Contable"
         Height          =   255
         Index           =   3
         Left            =   225
         TabIndex        =   21
         Top             =   1125
         Width           =   945
      End
   End
   Begin VB.Frame FrameTotFactu 
      Caption         =   "Total Factura"
      ForeColor       =   &H00972E0B&
      Height          =   1575
      Left            =   225
      TabIndex        =   25
      Top             =   3195
      Width           =   8580
      Begin VB.TextBox text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   11
         Left            =   3735
         MaxLength       =   15
         TabIndex        =   13
         Tag             =   "Importe Retención|N|S|||factsocio|impreten|#,###,###,##0.00|N|"
         Top             =   1125
         Width           =   1635
      End
      Begin VB.TextBox text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   9
         Left            =   225
         MaxLength       =   10
         TabIndex        =   11
         Tag             =   "Base Retencion|N|S|||factsocio|basereten|#,###,###,##0.00||"
         Text            =   "1234567890"
         Top             =   1125
         Width           =   1575
      End
      Begin VB.TextBox text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   10
         Left            =   2925
         MaxLength       =   6
         TabIndex        =   12
         Tag             =   "% Ret|N|S|0|100.00|factsocio|porcreten|##0.00|N|"
         Text            =   "99.99"
         Top             =   1125
         Width           =   645
      End
      Begin VB.TextBox text1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00CAE3FD&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   12
         Left            =   5805
         MaxLength       =   15
         TabIndex        =   14
         Tag             =   "Total Factura|N|S|||factsocio|totalfac|#,###,###,##0.00|N|"
         Top             =   495
         Width           =   2280
      End
      Begin VB.TextBox text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   6
         Left            =   2250
         MaxLength       =   2
         TabIndex        =   8
         Tag             =   "Tipo IVA 1|N|S|0|99|factsocio|tipoiva|00||"
         Text            =   "12"
         Top             =   510
         Width           =   525
      End
      Begin VB.TextBox text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   7
         Left            =   2940
         MaxLength       =   6
         TabIndex        =   9
         Tag             =   "% IVA 1|N|S|0|100.00|factsocio|porciva|##0.00|N|"
         Text            =   "99.99"
         Top             =   510
         Width           =   645
      End
      Begin VB.TextBox text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   8
         Left            =   3735
         MaxLength       =   15
         TabIndex        =   10
         Tag             =   "Importe IVA 1|N|N|||factsocio|cuotaiva|#,###,###,##0.00|N|"
         Top             =   495
         Width           =   1635
      End
      Begin VB.TextBox text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   5
         Left            =   225
         MaxLength       =   15
         TabIndex        =   7
         Tag             =   "Base IVA 1|N|S|||factsocio|baseimpo|#,###,###,##0.00|N|"
         Text            =   "575757575757557"
         Top             =   495
         Width           =   1605
      End
      Begin VB.Label Label1 
         Caption         =   "Importe Retención"
         Height          =   255
         Index           =   18
         Left            =   3735
         TabIndex        =   40
         Top             =   855
         Width           =   1545
      End
      Begin VB.Label Label1 
         Caption         =   "Base Retención"
         Height          =   255
         Index           =   17
         Left            =   225
         TabIndex        =   39
         Top             =   855
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "% Ret."
         Height          =   255
         Index           =   12
         Left            =   2925
         TabIndex        =   38
         Top             =   900
         Width           =   615
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   2
         Left            =   1965
         MousePointer    =   4  'Icon
         Tag             =   "-1"
         ToolTipText     =   "Buscar tipo de IVA"
         Top             =   510
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Total Factura"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   11
         Left            =   5805
         TabIndex        =   30
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "Tipo IVA"
         Height          =   255
         Index           =   14
         Left            =   2265
         TabIndex        =   29
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "% IVA"
         Height          =   255
         Index           =   15
         Left            =   2970
         TabIndex        =   28
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "Importe IVA"
         Height          =   255
         Index           =   16
         Left            =   3735
         TabIndex        =   27
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Base Imponible"
         Height          =   255
         Index           =   10
         Left            =   240
         TabIndex        =   26
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Frame Frame1 
      Height          =   555
      Index           =   0
      Left            =   225
      TabIndex        =   18
      Top             =   4950
      Width           =   2865
      Begin VB.Label lblIndicador 
         Alignment       =   2  'Center
         Caption         =   "Label2"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   120
         TabIndex        =   19
         Top             =   180
         Width           =   2655
      End
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   7785
      TabIndex        =   16
      Top             =   5040
      Width           =   1035
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   6465
      TabIndex        =   15
      Top             =   5025
      Width           =   1035
   End
   Begin MSAdodcLib.Adodc Data1 
      Height          =   375
      Left            =   3870
      Top             =   5040
      Visible         =   0   'False
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.CommandButton cmdRegresar 
      Caption         =   "&Regresar"
      Height          =   375
      Left            =   7785
      TabIndex        =   17
      Top             =   5040
      Visible         =   0   'False
      Width           =   1035
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   32
      Top             =   0
      Width           =   9060
      _ExtentX        =   15981
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      Style           =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   19
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Style           =   3
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Buscar"
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Ver todos"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Style           =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Nuevo"
            Object.Tag             =   "2"
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Modificar"
            Object.Tag             =   "2"
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Eliminar"
            Object.Tag             =   "2"
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.Visible         =   0   'False
            Style           =   3
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Imprimir"
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Salir"
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Style           =   3
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Primero"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Anterior"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Siguiente"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button19 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Último"
            ImageIndex      =   9
         EndProperty
      EndProperty
      BorderStyle     =   1
      Begin VB.CheckBox chkVistaPrevia 
         Caption         =   "Vista previa"
         Height          =   195
         Index           =   0
         Left            =   7290
         TabIndex        =   33
         Top             =   120
         Width           =   1215
      End
   End
   Begin VB.Label Label1 
      Caption         =   "Empresa"
      Height          =   255
      Index           =   6
      Left            =   360
      TabIndex        =   31
      Top             =   720
      Width           =   615
   End
   Begin VB.Menu mnOpciones 
      Caption         =   "&Opciones"
      Begin VB.Menu mnBuscar 
         Caption         =   "&Buscar"
         Shortcut        =   ^F
      End
      Begin VB.Menu mnVerTodos 
         Caption         =   "&Ver Todos"
         Shortcut        =   ^B
      End
      Begin VB.Menu mnBarra1 
         Caption         =   "-"
      End
      Begin VB.Menu mnNuevo 
         Caption         =   "&Nuevo"
         HelpContextID   =   2
         Shortcut        =   ^N
      End
      Begin VB.Menu mnModificar 
         Caption         =   "&Modificar"
         HelpContextID   =   2
         Shortcut        =   ^M
      End
      Begin VB.Menu mnEliminar 
         Caption         =   "&Eliminar"
         HelpContextID   =   2
         Shortcut        =   ^E
      End
      Begin VB.Menu mnBarra2 
         Caption         =   "-"
      End
      Begin VB.Menu mnImprimir 
         Caption         =   "&Imprimir"
         Shortcut        =   ^I
      End
      Begin VB.Menu mnBarra3 
         Caption         =   "-"
      End
      Begin VB.Menu mnSalir 
         Caption         =   "&Salir"
         Shortcut        =   ^S
      End
   End
End
Attribute VB_Name = "frmManFactSocios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' +-+-+-+-+-+-+-+-+-+-+-+-
' +-+- Autor: MONICA  +-+-
' +-+-+-+-+-+-+-+-+-+-+-+-

Option Explicit

Public DatosADevolverBusqueda As String    'Tindrà el nº de text que vol que torne, empipat
Public Event DatoSeleccionado(CadenaSeleccion As String)


Public numfactu As Long
Public LetraSerie As String
Public tipo As Byte ' 0 schfac normal
                    ' 1 schfacr ajena para el Regaixo

Private Modo As Byte
'*************** MODOS ********************
'   0.-  Formulari net sense cap camp ple
'   1.-  Preparant per a fer la búsqueda
'   2.-  Ja tenim registres i els anem a recorrer
'        ,podem editar-los Edició del camp
'   3.-  Inserció de nou registre
'   4.-  Modificar
'   5.-  Manteniment Llinies

'***Variables comuns a tots els formularis*****

Dim ModoLineas As Byte
'1.- Añadir,  2.- Modificar,  3.- Borrar,  0.-Pasar control a Lineas

Dim NumTabMto As Integer 'Indica quin nº de Tab està en modo Mantenimient
Dim TituloLinea As String 'Descripció de la llínia que està en Mantenimient
Dim PrimeraVez As Boolean

Private CadenaConsulta As String 'SQL de la taula principal del formulari
Private Ordenacion As String
Private NomTabla As String  'Nom de la taula

Private kCampo As Integer
'-------------------------------------------------------------------------
Private HaDevueltoDatos As Boolean

Dim btnPrimero As Byte 'Variable que indica el nº del Botó PrimerRegistro en la Toolbar1
'Dim CadAncho() As Boolean  'array, per a quan cridem al form de llínies
Dim indice As Byte 'Indice del text1 donde se ponen los datos devueltos desde otros Formularios de Mtos

Private WithEvents frmC As frmCal 'calendario fechas
Attribute frmC.VB_VarHelpID = -1
Private WithEvents frmB As frmBuscaGrid
Attribute frmB.VB_VarHelpID = -1
Private WithEvents frmZ As frmZoom  'Zoom para campos Text
Attribute frmZ.VB_VarHelpID = -1

Private WithEvents frmVar As frmManVariedad
Attribute frmVar.VB_VarHelpID = -1
Private WithEvents frmCtas As frmCtasConta
Attribute frmCtas.VB_VarHelpID = -1
Private WithEvents frmTipIVA As frmTipIVAConta  'Tipos de IVA de la contabilidad
Attribute frmTipIVA.VB_VarHelpID = -1

Dim CtaAnt As String
Dim FormaPagoAnt As String
Dim ModoModificar As Boolean
Dim ModificaImportes As Boolean ' variable que me indica q hay que modificar lineas de la factura de contabilidad
                                ' y cobros en la tesoreria

Dim BdConta As Integer
Dim BdConta1 As Integer

'GENERALES PARA PASARLE A CRYSTAL REPORT
Private cadFormula As String 'Cadena con la FormulaSelection para Crystal Report
Private cadParam As String 'Cadena con los parametros para Crystal Report
Private numParam As Byte 'Numero de parametros que se pasan a Crystal Report
Private cadSelect As String 'Cadena para comprobar si hay datos antes de abrir Informe
Private cadTitulo As String 'Titulo para la ventana frmImprimir
Private cadNombreRPT As String 'Nombre del informe

Dim TipForpa As String
Dim TipForpaAnt As String

' utilizado para buscar por checks
Private BuscaChekc As String

Dim CadenaBorrado As String


Private Sub Check1_Click(Index As Integer)
    If Modo = 1 Then
        'Buscqueda
        If InStr(1, BuscaChekc, "check1(" & Index & ")") = 0 Then BuscaChekc = BuscaChekc & "check1(" & Index & ")|"
    End If
End Sub

Private Sub Check1_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub cmdAceptar_Click()
Dim b As Boolean
Dim vSec As CSeccion 'Clase Seccion
Dim vTabla As String
Dim CtaClie As String
Dim Cad As String

' variables para el recalculo de iva y totales
    Dim i As Integer
    Dim Imptot(2)
    Dim Tipiva(2)
    Dim Impbas(2) As Currency
    Dim ImpIva(2) As Currency
    Dim PorIva(2) As Currency
    Dim ImpRec(2) As Currency
    Dim PorRec(2) As Currency
    Dim TotFac As Currency

' retencion
    Dim PorRet As Currency
    Dim ImpRet As Currency
    
    
    Screen.MousePointer = vbHourglass
    On Error GoTo Error1
    
    ModoModificar = False
    b = True
    Select Case Modo
        Case 1  'BUSQUEDA
            HacerBusqueda
            
        Case 3 'INSERTAR
            If DatosOk Then
                If InsertarDesdeForm2(Me, 1) Then
                    Data1.RecordSource = "Select * from " & NomTabla & Ordenacion
                    Cad = "numfactu = " & text1(0).Text & " and fecfactu = " & DBSet(text1(1).Text, "F")
                    Cad = Cad & " and codmacta = " & DBSet(text1(2).Text, "T")
                    PosicionarData Cad
                    PonerModo 2
                End If
            Else
                ModoLineas = 0
            End If

        Case 4  'MODIFICAR
            If Not DatosOk Then
                ModoLineas = 0
                Screen.MousePointer = vbDefault
                Exit Sub
            Else
                ModoModificar = True
                
'                conn.BeginTrans
'                If BdConta <> 0 Then
'                    If AbrirConexionContaFac(vParamAplic.UsuarioContaFac, vParamAplic.PasswordContaFac, BdConta) Then
'                        ConnContaFac.BeginTrans
'                        Set vEmpresaFac = New CempresaFac
'                        If vEmpresaFac.LeerNiveles Then
'                            PorRet = 0
'                            If text1(26).Text <> "" Then PorRet = CCur(ImporteSinFormato(text1(26).Text))
'                            AdoAux(0).Recordset.MoveFirst
'                            RecalculoBasesIvaFactura AdoAux(0).Recordset, Imptot, Tipiva, Impbas, impiva, PorIva, TotFac, ImpRec, PorRec, PorRet, ImpRet
'
'                            text1(28).Text = ""
'                            If ImpRet <> 0 Then text1(28).Text = Format(ImpRet, "#,###,###,##0.00")
'                            text1(24).Text = Format(TotFac, "#,###,###,##0.00")
'
'                            If text1(8).Text = "" Then text1(8).Text = "0,00"
'                            If text1(9).Text = "" Then text1(9).Text = "0,00"
'                        End If
'                        Set vEmpresaFac = Nothing
'                    End If
'                End If
'
'                If CadenaBorrado <> "" Then
'                    conn.Execute CadenaBorrado
'                    CadenaBorrado = ""
'                    EliminarLinea
'                End If
                
                
                If ModificaDesdeFormulario2(Me, 1) Then
                    If Check1(1).Value = 1 Then
                        MsgBox "Los cambios realizados recuerde hacerlos en la Contabilidad y Cartera correspondiente.", vbExclamation
                        
'12/02/2008: lo he quitado porque los cambios los haran ellos en la contabilidad
'                        'solo en el caso de que este contabilizada
'                        If Val(CtaAnt) <> Val(text1(4).Text) Then
''                            CtaClie = ""
''                            CtaClie = DevuelveDesdeBDNew(cPTours, "ssocio", "codmacta", "codsocio", text1(3).Text, "N")
'                            b = ModificaCtaClienteFacturaContabilidad(text1(0).Text, text1(1).Text, text1(2).Text, text1(4).Text)
'                        End If
'' 09022007 ya no dejo modificar la forma de pago
''                        If Val(FormaPagoAnt) <> Val(Text1(5).Text) Then _
''                            ModificaFormaPagoTesoreria Text1(0).Text, Text1(1).Text, Text1(2).Text, Text1(5).Text, FormaPagoAnt, TipForpa, TipForpaAnt
'
'                        If ModificaImportes And b Then
'                            BorrarTMPErrFact
'                            vTabla = "cabfact"
'' cuando aclare temas de contabilizacion en tesoreria se tiene que realizar esta funcion
''                            b = ModificaImportesFacturaContabilidad(text1(0).Text, text1(1).Text, text1(2).Text, text1(18).Text, text1(5).Text, vTabla)
'                            ModificaImportes = False
'                        End If
                    End If
                    TerminaBloquear
                    PosicionarData "numfactu = " & DBSet(text1(0).Text, "N") & " and fecfactu = " & DBSet(text1(1).Text, "F") & " and codmacta = " & DBSet(text1(2).Text, "T")
                End If
            End If
            
            
    End Select
    Screen.MousePointer = vbDefault
 
Error1:
    Screen.MousePointer = vbDefault
    If Err.Number <> 0 Or Not b Then
        MsgBox Err.Number & ": " & Err.Description, vbExclamation
'        If ModoModificar Then
'            conn.RollbackTrans
'            ConnContaFac.RollbackTrans
'            ModoModificar = False
'        End If
'    Else
'        If ModoModificar Then
'            conn.CommitTrans
'            ConnContaFac.CommitTrans
'            ModoModificar = False
'        End If
    End If
End Sub

Private Sub Form_Activate()
    Screen.MousePointer = vbDefault
    If PrimeraVez Then PrimeraVez = False
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub Form_Unload(Cancel As Integer)
    CheckValueGuardar Me.Name, Me.chkVistaPrevia(0).Value
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
Dim i As Integer
Dim sql2 As String

    PrimeraVez = True

    ' ICONITOS DE LA BARRA
    btnPrimero = 16 'index del botó "primero"
    With Me.Toolbar1
        .HotImageList = frmPpal.imgListComun_OM
        .DisabledImageList = frmPpal.imgListComun_BN
        .ImageList = frmPpal.imgListComun
        'el 1 i el 2 son separadors
        .Buttons(3).Image = 1   'Buscar
        .Buttons(4).Image = 2   'Todos
        'el 5 i el 6 son separadors
        .Buttons(7).Image = 3   'Insertar
        .Buttons(8).Image = 4   'Modificar
        .Buttons(9).Image = 5   'Borrar
'        .Buttons(10).Image = 16 ' Rectificativas
        'el 10 i el 11 son separadors
        .Buttons(12).Image = 10  'Imprimir
        .Buttons(13).Image = 11  'Salir
        'el 14 i el 15 son separadors
        .Buttons(btnPrimero).Image = 6  'Primero
        .Buttons(btnPrimero + 1).Image = 7 'Anterior
        .Buttons(btnPrimero + 2).Image = 8 'Siguiente
        .Buttons(btnPrimero + 3).Image = 9 'Último
    End With
    
    
    'cargar IMAGES de busqueda
    For i = 0 To Me.imgBuscar.Count - 1
        Me.imgBuscar(i).Picture = frmPpal.imgListImages16.ListImages(1).Picture
    Next i
   
    
    LimpiarCampos   'Limpia los campos TextBox
    
    CargaCombo
    
    '## A mano
    NomTabla = "factsocio"
    Ordenacion = " ORDER BY numfactu, fecfactu, codmacta "
    
    'Vemos como esta guardado el valor del check
    chkVistaPrevia(0).Value = CheckValueLeer(Name)
    
    'ASignamos un SQL al DATA1
    Data1.ConnectionString = conn
    sql2 = "Select * from " & NomTabla & " where numfactu = -1"
    Data1.RecordSource = sql2
    Data1.Refresh
        
    If DatosADevolverBusqueda = "" Then
        PonerModo 0
    Else
        PonerModo 1
        text1(0).BackColor = vbYellow 'letraser
    End If
    
    ModoLineas = 0
    
    
    If numfactu <> 0 Then
        text1(0).Text = numfactu
'        text1(1).Text = fecfactu
        PonerModo 1
        cmdAceptar_Click
    End If


End Sub

Private Sub LimpiarCampos()
    On Error Resume Next

    Limpiar Me   'Metodo general: Limpia los controles TextBox
    lblIndicador.Caption = ""
    
    Me.Check1(1).Value = 0
    
    Combo1(0).ListIndex = -1
    Combo1(1).ListIndex = -1
    'Aqui va el especifico de cada form es
    '### a mano
    If Err.Number <> 0 Then Err.Clear
End Sub

'----------------------------------------------------------------
'----------------------------------------------------------------
'   En PONERMODO se habilitan, o no, los diverso campos del
'   formulario en funcion del modo en k vayamos a trabajar
Private Sub PonerModo(Kmodo As Byte)
Dim i As Integer, NumReg As Byte
Dim b As Boolean
On Error GoTo EPonerModo
 
    Modo = Kmodo
    BuscaChekc = ""
    
    PonerIndicador lblIndicador, Modo, ModoLineas
       
    'Modo 2. Hay datos y estamos visualizandolos
    '=========================================
    'Ponemos visible, si es formulario de busqueda, el boton regresar cuando hay datos
    If DatosADevolverBusqueda <> "" Then
        cmdRegresar.visible = (Modo = 2)
    Else
        cmdRegresar.visible = False
    End If
    
    '=======================================
    b = (Modo = 2)
    'Poner Flechas de desplazamiento visibles
    NumReg = 1
    If Not Data1.Recordset.EOF Then
        If Data1.Recordset.RecordCount > 1 Then NumReg = 2 'Solo es para saber q hay + de 1 registro
    End If
    DesplazamientoVisible Me.Toolbar1, btnPrimero, b, NumReg
    

    b = Modo <> 0 And Modo <> 2
    cmdCancelar.visible = b
    cmdAceptar.visible = b
    
    '---------------------------------------------
    
    'Bloquea los campos Text1 si no estamos modificando/Insertando Datos
    'Si estamos en Insertar además limpia los campos Text1
    BloquearText1 Me, Modo
    BloquearChecks Me, Modo
    
    For i = 0 To text1.Count - 1
        text1(i).Enabled = (Modo = 3 Or Modo = 4 Or Modo = 1)
    Next i
    
    BloquearImgBuscar Me, Modo, ModoLineas
       
    'Bloquear los campos de clave primaria, NO se puede modificar
    b = Not (Modo = 1 Or Modo = 3) 'solo al insertar/buscar estará activo
    For i = 0 To 2
        BloquearTxt text1(i), b, True
        text1(i).Enabled = Not b
    Next i
    
    ' el importe de retencion solo se puede consultar
'[Monica] 25/01/2010 dejo introducir y modificar el importe de retencion y el total factura
'    BloquearTxt Text1(11), Not (Modo = 1)
'    Text1(11).Enabled = (Modo = 1)
    
    ' el importe total de factura solo se puede consultar
'[Monica] 25/01/2010 dejo introducir y modificar el importe de retencion y el total factura
'    BloquearTxt Text1(12), Not (Modo = 1)
'    Text1(12).Enabled = (Modo = 1)
    text1(12).BackColor = &HCAE3FD
    
    
    'Los % de IVA siempre bloqueados
    BloquearTxt text1(7), True
    text1(7).Enabled = (Modo = 1)
    

    ' **** si n'hi han imagens de buscar en la capçalera *****
    BloquearImgBuscar Me, Modo, ModoLineas
'    BloquearImgZoom Me, Modo, ModoLineas
    ' ********************************************************

    b = (Modo = 3) Or (Modo = 1)
    Me.imgBuscar(0).Enabled = b
    Me.imgBuscar(0).visible = b
    
    b = (Modo = 3) Or (Modo = 1) Or (Modo = 4)
    Me.imgBuscar(2).Enabled = b
    Me.imgBuscar(2).visible = b
    
    
    'Imagen Calendario fechas
    Me.imgFec(1).Enabled = (Modo = 3 Or Modo = 1) 'es clave, solo al insertar o buscar
    Me.imgFec(1).visible = (Modo = 3 Or Modo = 1) 'es clave, solo al insertar o buscar
            
    chkVistaPrevia(0).Enabled = (Modo <= 2)
    
    'Poner el tamaño de los campos. Si es modo Busqueda el MaxLength del campo
    'debe ser mayor.
    PonerLongCampos
                          
    ' solo podremos tocar el campo de contabilizado si estamos buscando
    Check1(1).Enabled = (Modo = 1)
    
    BloquearCombo Me, Modo
    
    PonerModoOpcionesMenu (Modo) 'Activar opciones menu según modo
    PonerOpcionesMenu   'Activar opciones de menu según nivel
                        'de permisos del usuario

EPonerModo:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Sub

Private Sub PonerLongCampos()
'Modificar el MaxLength del campo en funcion de si es modo de búsqueda o no
'para los campos que permitan introducir criterios más largos del tamaño del campo
'Si maxlength=6 para codprove, en modo busqueda Maxlenth=13 'el doble + 1
    
    'para los TEXT1
    PonerLongCamposGnral Me, Modo, 1
End Sub

Private Sub PonerOpcionesMenu()
'Activas unas Opciones de Menu y Toolbar según el nivel de usuario
    PonerOpcionesMenuGeneral Me
    PonerOpcionesMenuGeneralNew Me
End Sub

Private Sub PonerModoOpcionesMenu(Modo)
'Activas unas Opciones de Menu y Toolbar según el modo en que estemos
Dim b As Boolean, bAux As Boolean
Dim i As Byte

    '-----  TOOLBAR DE LA CABECERA
    '------------------------------------------
    'b = (Modo = 2 Or Modo = 0 Or Modo = 1)
    b = (Modo = 2 Or Modo = 0)
    'Buscar
    Toolbar1.Buttons(3).Enabled = b
    Me.mnBuscar.Enabled = b
    'Ver Todos
    Toolbar1.Buttons(4).Enabled = b
    Me.mnVerTodos.Enabled = b
    'Insertar
    Toolbar1.Buttons(7).Enabled = b
    Me.mnNuevo.Enabled = b
    
    b = (Modo = 2 And Data1.Recordset.RecordCount > 0) 'And (Check1(1).Value = 0)
    'Modificar
    Toolbar1.Buttons(8).Enabled = b
    Me.mnModificar.Enabled = b
    'eliminar
    Toolbar1.Buttons(9).Enabled = b
    Me.mnEliminar.Enabled = b
 
    'Imprimir
    'VRS:2.0.1(3)
    Toolbar1.Buttons(12).Enabled = (Modo = 2)
    Me.mnImprimir.Enabled = (Modo = 2)
End Sub

Private Sub Desplazamiento(Index As Integer)
'Botones de Desplazamiento de la Toolbar
'Para desplazarse por los registros de control Data
    If Data1.Recordset.EOF Then Exit Sub
    DesplazamientoData Data1, Index
    PonerCampos
End Sub

Private Sub frmB_Selecionado(CadenaDevuelta As String)
    Dim CadB As String
    Dim Aux As String
    
    If CadenaDevuelta <> "" Then
        HaDevueltoDatos = True
        Screen.MousePointer = vbHourglass
        'Sabemos que campos son los que nos devuelve
        'Creamos una cadena consulta y ponemos los datos
        CadB = ""
        Aux = ValorDevueltoFormGrid(text1(0), CadenaDevuelta, 1) 'numfactu
        CadB = Aux
        Aux = ValorDevueltoFormGrid(text1(1), CadenaDevuelta, 2) 'fecfactu
        CadB = CadB & " AND " & Aux
        Aux = ValorDevueltoFormGrid(text1(2), CadenaDevuelta, 3) 'codvarie
        CadB = CadB & " AND " & Aux
        '   Como la clave principal es unica, con poner el sql apuntando
        '   al valor devuelto sobre la clave ppal es suficiente
        CadenaConsulta = "select * from " & NomTabla & " WHERE " & CadB & " " & Ordenacion
        PonerCadenaBusqueda
        Screen.MousePointer = vbDefault
    End If
End Sub

Private Sub frmC_Selec(vFecha As Date)
    'Fecha
    text1(CByte(imgFec(1).Tag)).Text = Format(vFecha, "dd/mm/yyyy")
End Sub

Private Sub frmCtas_DatoSeleccionado(CadenaSeleccion As String)
    text1(indice).Text = RecuperaValor(CadenaSeleccion, 1) 'codmacta
    Text2(indice).Text = RecuperaValor(CadenaSeleccion, 2) 'nommacta
End Sub

Private Sub frmVar_DatoSeleccionado(CadenaSeleccion As String)
Dim Cad As String
    text1(3).Text = RecuperaValor(CadenaSeleccion, 1) 'codvarie
    Text2(3).Text = RecuperaValor(CadenaSeleccion, 2) 'nomvarie
End Sub

Private Sub frmTipIVA_DatoSeleccionado(CadenaSeleccion As String)
'Tipos de IVA (de la Contabilidad)
    text1(indice).Text = RecuperaValor(CadenaSeleccion, 1) 'codigiva
    FormateaCampo text1(indice)
    text1(indice + 1).Text = RecuperaValor(CadenaSeleccion, 3) '% iva
    If Modo <> 1 Then
        text1(indice + 3).Text = RecuperaValor(CadenaSeleccion, 4) '% rec
    End If
End Sub

Private Sub frmZ_Actualizar(vCampo As String)
     text1(indice).Text = vCampo
End Sub

Private Sub imgBuscar_Click(Index As Integer)
   'Screen.MousePointer = vbHourglass
    TerminaBloquear
    
    Select Case Index
        Case 0 'Cuenta Contable
            If vParamAplic.NumeroContaFacSoc = 0 Then Exit Sub
            
            indice = Index + 2
            Set frmCtas = New frmCtasConta
            frmCtas.NumDigit = 0
            frmCtas.Facturas = False
            frmCtas.CadBusqueda = vParamAplic.RaizCtaFacSoc
            frmCtas.DatosADevolverBusqueda = "0|1|"
            frmCtas.CodigoActual = Text2(indice).Text
            frmCtas.Conexion = cContaFacSoc
            frmCtas.Show vbModal
            Set frmCtas = Nothing
            PonerFoco text1(indice)

        Case 1 'Variedad
            indice = 3
            Set frmVar = New frmManVariedad
            frmVar.DatosADevolverBusqueda = "0|1|"
            frmVar.CodigoActual = text1(3).Text
            frmVar.Show vbModal
            Set frmVar = Nothing
            
        Case 2 'tipo de iva
            If vParamAplic.NumeroContaFacSoc = 0 Then Exit Sub
            
            indice = Index + 4
            
            Set frmTipIVA = New frmTipIVAConta
            frmTipIVA.Facturas = False
            frmTipIVA.Conexion = cContaFacSoc
            frmTipIVA.DatosADevolverBusqueda = "0|1|"
            frmTipIVA.CodigoActual = text1(indice).Text
            frmTipIVA.Show vbModal
            Set frmTipIVA = Nothing
            PonerFoco text1(indice)
            
    End Select
    
    If Modo = 4 Then BLOQUEADesdeFormulario2 Me, Data1, 1
End Sub

Private Sub imgFec_Click(Index As Integer)
    Dim esq As Long
    Dim dalt As Long
    Dim menu As Long
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
       
    menu = Me.Height - Me.ScaleHeight 'ací tinc el heigth del menú i de la toolbar

    ' es desplega baix i cap a la dreta
    'frmC.Left = esq + imgFec(Index).Parent.Left + 30
    'frmC.Top = dalt + imgFec(Index).Parent.Top + imgFec(Index).Height + menu - 40
    
    ' es desplega dalt i cap a la esquerra
    frmC.Left = esq + imgFec(Index).Parent.Left - frmC.Width + imgFec(Index).Width + 40
    frmC.Top = dalt + imgFec(Index).Parent.Top - frmC.Height + menu - 25
       
    ' ***canviar l'index de imgFec pel 1r index de les imagens de buscar data***
    imgFec(1).Tag = Index 'independentment de les dates que tinga, sempre pose l'index en la 27
    If text1(Index).Text <> "" Then frmC.NovaData = text1(Index).Text

    frmC.Show vbModal
    Set frmC = Nothing
    PonerFoco text1(CByte(imgFec(1).Tag))
    ' ***************************
End Sub

Private Sub mnBuscar_Click()
    BotonBuscar
    Me.Check1(1).Value = 0
End Sub

Private Sub mnEliminar_Click()
    BotonEliminar
End Sub

Private Sub mnImprimir_Click()
    'VRS:2.0.1(3): añadido el boton de imprimir
    cadTitulo = "Reimpresion de Facturas"

    ' ### [Monica] 11/09/2006
    '****************************
    Dim indRPT As Byte 'Indica el tipo de Documento en la tabla "scryst"
    Dim nomDocu As String 'Nombre de Informe rpt de crystal

    indRPT = 3 'Facturas Socios

    If Not PonerParamRPT(indRPT, cadParam, numParam, nomDocu) Then Exit Sub
    'Nombre fichero .rpt a Imprimir
    frmImprimir.NombreRPT = nomDocu
    ' he añadido estas dos lineas para que llame al rpt correspondiente

    cadNombreRPT = nomDocu
    cadFormula = "({" & NomTabla & ".numfactu} = " & text1(0).Text & ") AND "
    cadFormula = cadFormula & "{" & NomTabla & ".codmacta} = """ & text1(2).Text & """ AND {" & NomTabla & ".fecfactu} = cdate(""" & text1(1).Text & """) "
    
    '23022007 Monica: la separacion de la bonificacion solo la quieren en Alzira
'    If vParamAplic.Cooperativa = 1 Then cadFormula = cadFormula & " and {slhfac.numalbar} <> 'BONIFICA'" ' AND ({ssocio.impfactu}<=1)"
    
    cadParam = "|pEmpresa=" & vEmpresa.nomEmpre '& "|pCodigoISO="11112"|pCodigoRev="01"|
    LlamarImprimir
End Sub

Private Sub mnModificar_Click()

    'Comprobaciones
    '--------------
    If Data1.Recordset.EOF Then Exit Sub
    If Data1.Recordset.RecordCount < 1 Then Exit Sub
    
    'El registre de codi 0 no es pot Modificar ni Eliminar
    ' ### [Monica] 27/09/2006
    ' quitamos el control de no poder modificar ni eliminar si es 0
    'If EsCodigoCero(CStr(Data1.Recordset.Fields(0).Value), FormatoCampo(text1(0))) Then Exit Sub
    
    ' ### [Monica] 27/09/2006
    ' solo podemos modificar en el caso de que haya contabilidad si la factura es modificable
'12/02/2008: lo he quitado porque lo modificaran ellos manualmente en la contabilidad
'    If vParamAplic.NumeroConta <> 0 And Not FacturaModificable(text1(0).Text, text1(1).Text, text1(2).Text, Check1(1).Value) Then Exit Sub
    
    
    'Preparar para modificar
    '-----------------------
    If BLOQUEADesdeFormulario2(Me, Data1, 1) Then BotonModificar
End Sub

Private Sub mnNuevo_Click()
     BotonAnyadir
End Sub


Private Sub mnSalir_Click()
    Unload Me
End Sub

Private Sub mnVerTodos_Click()
    BotonVerTodos
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim Cad As String
    
    
    Select Case Button.Index
        Case 3  'Buscar
           mnBuscar_Click
        Case 4  'Todos
            mnVerTodos_Click
        Case 7  'Nuevo
            mnNuevo_Click
        Case 8  'Modificar
            mnModificar_Click
        Case 9  'Borrar
            mnEliminar_Click
        Case 12 'Imprimir
            mnImprimir_Click
        Case 13    'Salir
            mnSalir_Click
        Case btnPrimero To btnPrimero + 3 'Flechas Desplazamiento
            Desplazamiento (Button.Index - btnPrimero)
    End Select
End Sub

Private Sub BotonBuscar()
    'Buscar
    
    If Modo <> 1 Then
        LimpiarCampos
        'Poner los grid sin apuntar a nada
        'LimpiarDataGrids
        PonerModo 1
        'Si pasamos el control aqui lo ponemos en amarillo
        PonerFoco text1(0)
        text1(0).BackColor = vbYellow
    Else
        HacerBusqueda
        If Data1.Recordset.EOF Then
            text1(kCampo).Text = ""
            text1(kCampo).BackColor = vbYellow
            PonerFoco text1(kCampo)
        End If
    End If
End Sub

Private Sub HacerBusqueda()
Dim CadB As String

    CadB = ObtenerBusqueda2(Me, BuscaChekc)
    
    If chkVistaPrevia(0) = 1 Then
        MandaBusquedaPrevia CadB
    ElseIf CadB <> "" Then
        'Se muestran en el mismo form
        CadenaConsulta = "select * from " & NomTabla & " WHERE " & CadB & " " & Ordenacion
        PonerCadenaBusqueda
    Else
        PonerFoco text1(0)
    End If
End Sub

Private Sub MandaBusquedaPrevia(CadB As String)
Dim Cad As String
        'Llamamos a al form
        '##A mano
        Cad = ""
        Cad = Cad & "Factura|" & NomTabla & ".numfactu|N|" & FormatoCampo(text1(0)) & "|15·"
        Cad = Cad & ParaGrid(text1(1), 15, "Fecha")
        Cad = Cad & "Cta.Contable|" & NomTabla & ".codmacta|T|" & text1(2) & "|15·"
        Cad = Cad & "Variedad|" & NomTabla & ".codvarie|T|" & text1(3) & "|15·"
        Cad = Cad & "Descripcion|variedad.nomvarie|T|" & Text2(3) & "|40·"
        If Cad <> "" Then
            Screen.MousePointer = vbHourglass
            Set frmB = New frmBuscaGrid
            frmB.vCampos = Cad
            frmB.vTabla = NomTabla & " INNER JOIN variedad ON " & NomTabla & ".codvarie=variedad.codvarie "
            frmB.vSQL = CadB
            HaDevueltoDatos = False
            '###A mano
            frmB.vDevuelve = "0|1|2|"
            frmB.vTitulo = "Facturas Socios"
            frmB.vSelElem = 0
            '#
            frmB.Show vbModal
            Set frmB = Nothing
            'Si ha puesto valores y tenemos que es formulario de busqueda entonces
            'tendremos que cerrar el form lanzando el evento
            If HaDevueltoDatos Then
                If (Not Data1.Recordset.EOF) And DatosADevolverBusqueda <> "" Then _
                    cmdRegresar_Click
            Else   'de ha devuelto datos, es decir NO ha devuelto datos
                PonerFoco text1(kCampo)
            End If
        End If
End Sub

Private Sub cmdRegresar_Click()
Dim Cad As String
Dim Aux As String
Dim i As Integer
Dim J As Integer

    If Data1.Recordset.EOF Then
        MsgBox "Ningún registro devuelto.", vbExclamation
        Exit Sub
    End If
    
    Cad = ""
    i = 0
    Do
        J = i + 1
        i = InStr(J, DatosADevolverBusqueda, "|")
        If i > 0 Then
            Aux = Mid(DatosADevolverBusqueda, J, i - J)
            J = Val(Aux)
            Cad = Cad & text1(J).Text & "|"
        End If
    Loop Until i = 0
    RaiseEvent DatoSeleccionado(Cad)
    Unload Me
End Sub

Private Sub PonerCadenaBusqueda()
    Screen.MousePointer = vbHourglass
    On Error GoTo EEPonerBusq
    
    Data1.RecordSource = CadenaConsulta
    Data1.Refresh
    If Data1.Recordset.RecordCount <= 0 Then
        MsgBox "No hay ningún registro en la tabla " & NomTabla, vbInformation
        Screen.MousePointer = vbDefault
        Exit Sub
    Else
        PonerModo 2
        'Data1.Recordset.MoveLast
        Data1.Recordset.MoveFirst
        PonerCampos
    End If
    
    Screen.MousePointer = vbDefault
    Exit Sub
    
EEPonerBusq:
    MuestraError Err.Number, "PonerCadenaBusqueda"
    PonerModo 0
    Screen.MousePointer = vbDefault
End Sub

Private Sub BotonVerTodos()
'Ver todos
Dim i As Integer

    LimpiarCampos 'Limpia los Text1
    
    If chkVistaPrevia(0).Value = 1 Then
        MandaBusquedaPrevia ""
    Else
        CadenaConsulta = "Select * from " & NomTabla & Ordenacion
        PonerCadenaBusqueda
    End If
End Sub

Private Sub BotonAnyadir()
'Añadir registro en tabla de expedientes individuales: expincab (Cabecera)

    LimpiarCampos 'Vacía los TextBox
    'Poner los grid sin apuntar a nada
'    LimpiarDataGrids

    'Añadiremos el boton de aceptar y demas objetos para insertar
    PonerModo 3

    'Quan afegixc pose en Fecha
    text1(1).Text = Format(Now, "dd/mm/yyyy")
    text1(10).Text = Format(vParamAplic.PorcreteFacSoc, "##0.00")
    

    'Total Factura (por defecto=0)
'    text1(18).Text = "0"
'    text1(19).Text = "0"

    'em posicione en el 1r tab
    PonerFoco text1(0)
End Sub

Private Sub BotonModificar()
Dim vSec As CSeccion
Dim Cad As String

    '++monica:12/02/2008
    If CByte(Data1.Recordset!intconta) = 1 Then
       Cad = "   Se dispone a realizar cambios en los datos de la Factura.     " & vbCrLf & vbCrLf & _
             "Recuerde modificar la Contabilidad y Tesoreria correspondiente!!!"
       MsgBox Cad, vbExclamation
    End If
    '++

    'Añadiremos el boton de aceptar y demas objetos para insertar
    PonerModo 4
    
    'Quan modifique pose en la F.Modificación la data actual
    PonerFoco text1(3)
End Sub

Private Sub BotonEliminar()
Dim Cad As String
Dim vSec As CSeccion
Dim NumFacElim As Long 'Numero de la Factura que se ha Eliminado
Dim NumSecElim As Integer 'Numero de la Seccion que se ha eliminado

    On Error GoTo EEliminar


    'Ciertas comprobaciones
    If Data1.Recordset.EOF Then Exit Sub
    
    '++monica:12/02/2008
    If CByte(Data1.Recordset!intconta) = 1 Then
       Cad = "No se permite eliminar una Factura Contabilizada!!!"
       MsgBox Cad, vbExclamation
       Exit Sub
    End If
    '++
    
'    'El registre de codi 0 no es pot Modificar ni Eliminar
'    If EsCodigoCero(CStr(Data1.Recordset.Fields(1).Value), FormatoCampo(text1(1))) Then Exit Sub

    Cad = "¿Seguro que desea eliminar la factura?"
    Cad = Cad & vbCrLf & "Factura: " & Format(Data1.Recordset!numfactu, FormatoCampo(text1(0)))
    Cad = Cad & vbCrLf & "Fecha: " & Data1.Recordset.Fields("fecfactu")
    Cad = Cad & vbCrLf & "Cta.Contable: " & Data1.Recordset!Codmacta
    
    'Borramos
    If MsgBox(Cad, vbQuestion + vbYesNo) = vbYes Then
        'Hay que eliminar
        On Error GoTo EEliminar
        Screen.MousePointer = vbHourglass
        NumRegElim = Data1.Recordset.AbsolutePosition
        If Not eliminar Then
            Screen.MousePointer = vbDefault
            Exit Sub
        Else
            If SituarDataTrasEliminar(Data1, NumRegElim) Then
                PonerCampos
            Else
                LimpiarCampos
                'Poner los grid sin apuntar a nada
                'LimpiarDataGrids
                PonerModo 0
            End If
        End If
    End If
    Screen.MousePointer = vbDefault
    
EEliminar:
    Screen.MousePointer = vbDefault
    If Err.Number > 0 Then MuestraError Err.Number, "Eliminar Factura", Err.Description
End Sub

Private Sub PonerCampos()
Dim i As Integer
Dim vSec As CSeccion
    If Data1.Recordset.EOF Then Exit Sub
    PonerCamposForma2 Me, Data1, 1 'opcio=1: pone el formato o los campos de la cabecera
    
    'Recuperar Descripciones de los campos de Codigo
    '--------------------------------------------------
    Text2(3).Text = PonerNombreDeCod(text1(3), "variedad", "nomvarie")
    Text2(2).Text = PonerNombreCuenta(text1(2), Modo, , cContaFacSoc, False)
    
    '-- Esto permanece para saber donde estamos
    lblIndicador.Caption = Data1.Recordset.AbsolutePosition & " de " & Data1.Recordset.RecordCount
    PonerModoOpcionesMenu (Modo)
End Sub

Private Sub cmdCancelar_Click()
Dim i As Integer
Dim V

    Select Case Modo
        Case 1, 3 'Busqueda, Insertar
                LimpiarCampos
                If Data1.Recordset.EOF Then
                    PonerModo 0
                Else
                    PonerModo 2
                    PonerCampos
                End If
                PonerFoco text1(0)

        Case 4  'Modificar
                TerminaBloquear
                PonerModo 2
                PonerCampos
                PonerFoco text1(0)
        
    End Select
End Sub

Private Function DatosOk() As Boolean
Dim b As Boolean
Dim Cad As String

    On Error GoTo EDatosOK

    DatosOk = False
    b = CompForm2(Me, 1) 'Comprobar formato datos ok
    If Not b Then Exit Function
    
    If Mid(text1(2).Text, 1, vEmpresaFacSoc.DigitosNivelAnterior) <> vParamAplic.RaizCtaFacSoc Then
        Cad = "La Cuenta contable no coincide con la Raiz del Socio." & vbCrLf & vbCrLf & "         ¿Desea continuar?"
        If MsgBox(Cad, vbQuestion + vbYesNo + vbDefaultButton1) = vbNo Then
            b = False
            PonerFoco text1(2)
        End If
    End If
    
    DatosOk = b
    
EDatosOK:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Function

Private Sub PosicionarData(Cad As String)
'Dim cad As String
Dim Indicador As String
    
  '  cad = ""
    If SituarDataMULTI(Data1, Cad, Indicador) Then
        If ModoLineas <> 1 Then
            PonerModo 2
        End If
       
       lblIndicador.Caption = Indicador
    Else
       LimpiarCampos
       'Poner los grid sin apuntar a nada
       'LimpiarDataGrids
       PonerModo 0
    End If
End Sub

Private Function eliminar() As Boolean
Dim vWhere As String

    On Error GoTo FinEliminar
        
    eliminar = False
        
    vWhere = ObtenerWhereCab(True)

    conn.Execute "Delete from " & NomTabla & vWhere
        
    eliminar = True
    Exit Function
    
FinEliminar:
    If Err.Number <> 0 Then
        MuestraError Err.Number, "Eliminar"
    End If
End Function

Private Sub Text1_GotFocus(Index As Integer)
    kCampo = Index
    ConseguirFoco text1(Index), Modo
End Sub

Private Sub Text1_LostFocus(Index As Integer)
Dim Cad As String, Datos As String
Dim Suma As Currency
Dim i As Integer

    If Not PerderFocoGnral(text1(Index), Modo) Then Exit Sub
    
    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub
    
    Select Case Index
        Case 0 'Nº factura
            If text1(Index).Text <> "" Then FormateaCampo text1(Index)
                        
        Case 1 'Fecha
            If text1(Index).Text <> "" Then PonerFormatoFecha text1(Index)
            
        Case 3 'Variedad
            If text1(Index).Text <> "" Then
                If PonerFormatoEntero(text1(3)) Then
                    Text2(Index).Text = PonerNombreDeCod(text1(Index), "variedad", "nomvarie", "codvarie", "N")
                    If Text2(Index).Text = "" Then
                        Cad = "No existe la Variedad: " & text1(Index).Text & vbCrLf
                        Cad = Cad & "¿Desea crearla?" & vbCrLf
                        If MsgBox(Cad, vbQuestion + vbYesNo) = vbYes Then
                            Set frmVar = New frmManVariedad
                            frmVar.DatosADevolverBusqueda = "0|1|"
                            text1(Index).Text = ""
                            TerminaBloquear
                            frmVar.Show vbModal
                            Set frmVar = Nothing
                            If Modo = 4 Then BLOQUEADesdeFormulario2 Me, Data1, 1
                        Else
                            text1(Index).Text = ""
                        End If
                        PonerFoco text1(Index)
                    End If
                Else
                    Text2(Index).Text = ""
                End If
            End If
            
        Case 2 'Cta Contable
            If text1(Index).Text = "" Then Exit Sub
            Text2(Index).Text = PonerNombreCuenta(text1(Index), Modo, , cContaFacSoc)
        
        Case 4 'kilos facturados
            PonerFormatoEntero text1(Index)
            
        Case 5, 8, 9 ', 11, 12 'IMPORTES Base, IVA
            If text1(Index).Text = "" Then Exit Sub
            If Modo = 1 Then Exit Sub
            If PonerFormatoDecimal(text1(Index), 1) Then CalculoTotales
            
'[Monica] 25/01/2010 lo he quitado de arriba y lo he dejado modificar
        Case 11, 12  'IMPORTES Base, IVA
            If Modo = 1 Then Exit Sub
            PonerFormatoDecimal text1(Index), 1
            
        Case 7, 10 'porcentajes
            If text1(Index).Text = "" Then Exit Sub
            If Modo = 1 Then Exit Sub
            If PonerFormatoDecimal(text1(Index), 7) Then CalculoTotales
        
        Case 6 'cod. IVA
           If text1(Index).Text = "" Then
              text1(Index + 1).Text = ""
           Else
              text1(Index + 1).Text = DevuelveDesdeBDNew(cContaFacSoc, "tiposiva", "porceiva", "codigiva", text1(Index).Text, "N")
           End If
           CalculoTotales
    End Select
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = teclaBuscar Then
        If Modo = 1 Or Modo = 3 Or Modo = 4 Then
            Select Case Index
                Case 2: KEYFecha KeyAscii, 2
'                Case 3: KEYBusqueda KeyAscii, 0
'                Case 4: KEYBusqueda KeyAscii, 1
'                Case 5: KEYBusqueda KeyAscii, 2
'                Case 7: KEYBusqueda KeyAscii, 3
'                Case 11: KEYBusqueda KeyAscii, 4
'                Case 15: KEYBusqueda KeyAscii, 5
'               ' Case 1: KEYFecha KeyAscii, 1
            End Select
        End If
    Else
        If Not text1(Index).MultiLine Then
            KEYpress KeyAscii
        End If
    End If
End Sub

Private Sub Text1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
'Avanzar/Retroceder los campos con las flechas de desplazamiento del teclado.
    KEYdown KeyCode
End Sub

Private Sub KEYpress(KeyAscii As Integer)
    If KeyAscii = 13 Then 'ENTER
        KeyAscii = 0
        SendKeys "{tab}"
    ElseIf KeyAscii = 27 Then 'ESC
        If (Modo = 0 Or Modo = 2) Then Unload Me
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

Private Function ObtenerWhereCab(conW As Boolean) As String
Dim vWhere As String
    vWhere = ""
    If conW Then vWhere = " WHERE "
    vWhere = vWhere & " numfactu = " & text1(0).Text
    vWhere = vWhere & " AND fecfactu= '" & Format(text1(1).Text, FormatoFecha) & "'"
    vWhere = vWhere & " AND codmacta= '" & Trim(text1(2).Text) & "'"
    ObtenerWhereCab = vWhere
End Function

' *** si n'hi han combos a la capçalera ***
Private Sub Combo1_GotFocus(Index As Integer)
    If Modo = 1 Then Combo1(Index).BackColor = vbYellow
End Sub

Private Sub Combo1_LostFocus(Index As Integer)
    If Combo1(Index).BackColor = vbYellow Then Combo1(Index).BackColor = vbWhite
    If Index = 0 And (Modo = 3 Or Modo = 4) Then
        PonerCamposRet
        CalculoTotales
    End If
End Sub

Private Sub Combo1_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Function FacturaModificable(letraser As String, numfactu As String, fecfactu As String, Contabil As String) As Boolean

    FacturaModificable = False
    
    If Contabil = 0 Then
        FacturaModificable = True
    Else
        ' si la factura esta contabilizada tenemos que ver si en la contabilidad esta contabilizada y
        ' si en la tesoreria esta remesada o cobrada en estos casos la factura no puede ser modificada
        If FacturaContabilizada(letraser, numfactu, Year(CDate(fecfactu))) Then
            MsgBox "Factura contabilizada en la Contabilidad, no puede modificarse ni eliminarse."
            Exit Function
        End If
        
        If FacturaRemesada(letraser, numfactu, fecfactu) Then
            MsgBox "Factura Remesada, no puede modificarse ni eliminarse."
            Exit Function
        End If
        
        If FacturaCobrada(letraser, numfactu, fecfactu) Then
            MsgBox "Factura Cobrada, no puede modificarse ni eliminarse."
            Exit Function
        End If
           
        FacturaModificable = True
    End If

End Function

'VRS:2.0.1(3)
Private Sub LlamarImprimir()
    With frmImprimir
        .FormulaSeleccion = cadFormula
        .OtrosParametros = cadParam
        .NumeroParametros = 2
        .SoloImprimir = False
        .Titulo = cadTitulo
        .NombreRPT = cadNombreRPT
        .Facturas = False
        .EnvioEMail = False
        .Contabilidad = cContaFacSoc
        .Opcion = 1
        .Show vbModal
    End With
End Sub


' ********* si n'hi han combos a la capçalera ************
Private Sub CargaCombo()
Dim Ini As Integer
Dim Fin As Integer
Dim i As Integer

    ' *** neteje els combos, els pose valor i seleccione el valor per defecte ***
    ' tipo de socio
    Combo1(0).Clear
    
    Combo1(0).AddItem "Módulos"
    Combo1(0).ItemData(Combo1(0).NewIndex) = 0
    Combo1(0).AddItem "E.Directa"
    Combo1(0).ItemData(Combo1(0).NewIndex) = 1
    Combo1(0).AddItem "Entidad"
    Combo1(0).ItemData(Combo1(0).NewIndex) = 2

    'tipo de factura
    Combo1(1).Clear
    
    Combo1(1).AddItem "Anticipo"
    Combo1(1).ItemData(Combo1(1).NewIndex) = 0
    Combo1(1).AddItem "Liquidación"
    Combo1(1).ItemData(Combo1(1).NewIndex) = 1
    Combo1(1).AddItem "Retirada"
    Combo1(1).ItemData(Combo1(1).NewIndex) = 2
    Combo1(1).AddItem "Industria"
    Combo1(1).ItemData(Combo1(1).NewIndex) = 3
    Combo1(1).AddItem "Subvención"
    Combo1(1).ItemData(Combo1(1).NewIndex) = 4
    Combo1(1).AddItem "Siniestro"
    Combo1(1).ItemData(Combo1(1).NewIndex) = 5

End Sub

Private Sub CalculoTotales()
Dim Base As Currency
Dim Tiva As Currency
Dim PorIva As Currency
Dim ImpIva As Currency
Dim BaseReten As Currency
Dim PorRet As Currency
Dim ImpRet As Currency
Dim TotFac As Currency

    Base = CCur(ComprobarCero(text1(5).Text))
    PorIva = CCur(ComprobarCero(text1(7).Text))
    ImpIva = Round2(Base * PorIva / 100, 2)
    
    Select Case Combo1(0).ListIndex
        Case 0
            BaseReten = Base + ImpIva
        Case 1
            BaseReten = Base
        Case 2
            BaseReten = 0
    End Select
    
    'solo en el caso de que estemos insertando y modificando y no haya % de retencion
    'le daremos el que hay en parametros
    If text1(10).Text = "" And Combo1(0).ListIndex <> 2 And (Modo = 3 Or Modo = 4) Then
        text1(10).Text = CCur(ComprobarCero(vParamAplic.PorcreteFacSoc))
    End If
    
    ' calculo de la retencion
    PorRet = CCur(ComprobarCero(text1(10).Text))
    ImpRet = Round2(BaseReten * PorRet / 100, 2)
    
    TotFac = Base + ImpIva - ImpRet

    text1(8).Text = Format(ImpIva, "#,###,###,##0.00")
    
    If BaseReten = 0 Then
        text1(9).Text = ""
    Else
        text1(9).Text = Format(BaseReten, "#,###,###,##0.00")
    End If
    
    If ImpRet = 0 Then
        text1(11).Text = ""
    Else
        text1(11).Text = ImpRet
    End If
    
    If TotFac = 0 Then
        text1(12).Text = ""
    Else
        text1(12).Text = Format(TotFac, "#,###,###,##0.00")
    End If
End Sub

Private Sub PonerCamposRet()
Dim i As Integer
    If Not (Modo = 3 Or Modo = 4) Then Exit Sub
    
    For i = 9 To 11
        text1(i).Enabled = (Combo1(0).ListIndex <> 2)
        If (Combo1(0).ListIndex = 2) Then
             text1(i).Text = ""
        End If
    Next i
    
    
End Sub
