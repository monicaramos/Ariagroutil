VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmManFactCV1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Facturas Coarval y Varias"
   ClientHeight    =   5745
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12270
   Icon            =   "frmManFactCV1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5745
   ScaleWidth      =   12270
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtAux2 
      BackColor       =   &H80000018&
      Enabled         =   0   'False
      Height          =   285
      Index           =   2
      Left            =   4860
      Locked          =   -1  'True
      MaxLength       =   40
      TabIndex        =   50
      Top             =   4560
      Width           =   945
   End
   Begin VB.CheckBox chkAbonos 
      Caption         =   "Errónea"
      Height          =   255
      Index           =   1
      Left            =   11160
      TabIndex        =   49
      Tag             =   "Factura Errónea|N|N|0|1|cvfacturas|erronea|||"
      Top             =   4800
      Width           =   960
   End
   Begin VB.TextBox txtAux 
      Alignment       =   1  'Right Justify
      Height          =   290
      Index           =   18
      Left            =   10710
      MaxLength       =   30
      TabIndex        =   18
      Tag             =   "Base Imponible3|N|S|||cvfacturas|baseimpo3|##,###,#0.00||"
      Top             =   3780
      Width           =   1395
   End
   Begin VB.TextBox txtAux 
      Alignment       =   1  'Right Justify
      Height          =   290
      Index           =   17
      Left            =   10710
      MaxLength       =   30
      TabIndex        =   21
      Tag             =   "Cuota Iva3|N|S|||cvfacturas|cuotaiva3|##,###,#0.00||"
      Top             =   4410
      Width           =   1395
   End
   Begin VB.TextBox txtAux 
      Alignment       =   1  'Right Justify
      Height          =   285
      Index           =   16
      Left            =   11370
      MaxLength       =   30
      TabIndex        =   20
      Tag             =   "Porcentaje Iva3|N|S|||cvfacturas|porciva3|##0.00||"
      Top             =   4080
      Width           =   735
   End
   Begin VB.TextBox txtAux 
      Alignment       =   1  'Right Justify
      Height          =   290
      Index           =   15
      Left            =   10710
      MaxLength       =   30
      TabIndex        =   19
      Tag             =   "Código Iva3|N|S|||cvfacturas|codiva3|00||"
      Top             =   4080
      Width           =   465
   End
   Begin VB.TextBox txtAux 
      Alignment       =   1  'Right Justify
      Height          =   290
      Index           =   14
      Left            =   10710
      MaxLength       =   30
      TabIndex        =   14
      Tag             =   "Base Imponible 2|N|S|||cvfacturas|baseimpo2|##,###,#0.00||"
      Top             =   2850
      Width           =   1395
   End
   Begin VB.TextBox txtAux 
      Alignment       =   1  'Right Justify
      Height          =   290
      Index           =   13
      Left            =   10710
      MaxLength       =   30
      TabIndex        =   17
      Tag             =   "Cuota Iva2|N|S|||cvfacturas|cuotaiva2|##,###,#0.00||"
      Top             =   3450
      Width           =   1395
   End
   Begin VB.TextBox txtAux 
      Alignment       =   1  'Right Justify
      Height          =   285
      Index           =   12
      Left            =   11370
      MaxLength       =   30
      TabIndex        =   16
      Tag             =   "Porcentaje Iva2|N|S|||cvfacturas|porciva2|##0.00||"
      Top             =   3150
      Width           =   735
   End
   Begin VB.TextBox txtAux 
      Alignment       =   1  'Right Justify
      Height          =   290
      Index           =   11
      Left            =   10710
      MaxLength       =   30
      TabIndex        =   15
      Tag             =   "Código Iva2|N|S|||cvfacturas|codiva2|00||"
      Top             =   3150
      Width           =   465
   End
   Begin VB.TextBox txtAux 
      Height          =   285
      Index           =   7
      Left            =   10800
      MaxLength       =   10
      TabIndex        =   8
      Tag             =   "Cuenta Contable|T|S|||cvfacturas|codmactavta|||"
      Top             =   900
      Width           =   1275
   End
   Begin VB.TextBox txtAux 
      Alignment       =   1  'Right Justify
      Height          =   290
      Index           =   4
      Left            =   10710
      MaxLength       =   30
      TabIndex        =   10
      Tag             =   "Base Imponible|N|N|||cvfacturas|baseimpo|##,###,#0.00||"
      Top             =   1950
      Width           =   1395
   End
   Begin VB.TextBox txtAux 
      Alignment       =   1  'Right Justify
      Height          =   290
      Index           =   5
      Left            =   10710
      MaxLength       =   30
      TabIndex        =   13
      Tag             =   "Cuota Iva|N|N|||cvfacturas|cuotaiva|##,###,#0.00||"
      Top             =   2550
      Width           =   1395
   End
   Begin VB.TextBox txtAux 
      Alignment       =   1  'Right Justify
      Height          =   285
      Index           =   9
      Left            =   11370
      MaxLength       =   30
      TabIndex        =   12
      Tag             =   "Porcentaje Iva|N|N|||cvfacturas|porciva|##0.00||"
      Top             =   2250
      Width           =   735
   End
   Begin VB.TextBox txtAux 
      Alignment       =   1  'Right Justify
      Height          =   290
      Index           =   10
      Left            =   10710
      MaxLength       =   30
      TabIndex        =   11
      Tag             =   "Código Iva|N|N|||cvfacturas|codiva|00||"
      Top             =   2250
      Width           =   465
   End
   Begin VB.CheckBox chkAbonos 
      Caption         =   "Contabilizada"
      Height          =   255
      Index           =   0
      Left            =   9630
      TabIndex        =   22
      Tag             =   "Int.Contable|N|N|0|1|cvfacturas|intconta|||"
      Top             =   4800
      Width           =   1410
   End
   Begin VB.TextBox txtAux 
      Height          =   290
      Index           =   8
      Left            =   10800
      MaxLength       =   10
      TabIndex        =   7
      Tag             =   "Nif Socio|T|S|||cvfacturas|nifsocio|||"
      Top             =   600
      Width           =   1275
   End
   Begin VB.TextBox txtAux 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   290
      Index           =   6
      Left            =   5850
      MaxLength       =   30
      TabIndex        =   6
      Tag             =   "Total Factura|N|N|||cvfacturas|totalfac|##,###,#0.00||"
      Top             =   4560
      Width           =   1395
   End
   Begin VB.CommandButton btnBuscar 
      Appearance      =   0  'Flat
      Caption         =   "+"
      Height          =   300
      Index           =   0
      Left            =   4620
      MaskColor       =   &H00000000&
      TabIndex        =   40
      ToolTipText     =   "Buscar cuenta contable"
      Top             =   4560
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.TextBox txtAux 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   290
      Index           =   3
      Left            =   2580
      MaxLength       =   30
      TabIndex        =   4
      Tag             =   "Fecha de Factura|F|N|||cvfacturas|fecfactu|dd/mm/yyyy|S|"
      Top             =   4560
      Width           =   885
   End
   Begin VB.CommandButton btnBuscar 
      Appearance      =   0  'Flat
      Caption         =   "+"
      Height          =   300
      Index           =   1
      Left            =   3480
      MaskColor       =   &H00000000&
      TabIndex        =   39
      ToolTipText     =   "Buscar fecha"
      Top             =   4560
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.TextBox txtAux 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   290
      Index           =   2
      Left            =   3720
      MaxLength       =   10
      TabIndex        =   5
      Tag             =   "Cuenta Contable|T|S|||cvfacturas|codmactasoc|||"
      Top             =   4560
      Width           =   885
   End
   Begin VB.TextBox txtAux 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   290
      Index           =   1
      Left            =   1440
      MaxLength       =   30
      TabIndex        =   3
      Tag             =   "Numero de Factura|T|N|||cvfacturas|numfactu||S|"
      Top             =   4530
      Width           =   1065
   End
   Begin VB.TextBox txtAux 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   290
      Index           =   0
      Left            =   600
      MaxLength       =   3
      TabIndex        =   2
      Tag             =   "Letra Serie|T|N|||cvfacturas|letraser||S|"
      Top             =   4530
      Width           =   800
   End
   Begin VB.ComboBox Combo1 
      Appearance      =   0  'Flat
      Height          =   315
      Index           =   0
      ItemData        =   "frmManFactCV1.frx":000C
      Left            =   -150
      List            =   "frmManFactCV1.frx":000E
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Tag             =   "Tipo Factura|N|N|0|2|cvfacturas|tipofactu||S|"
      Top             =   4560
      Width           =   735
   End
   Begin VB.TextBox txtAux 
      Height          =   285
      Index           =   19
      Left            =   10260
      MaxLength       =   10
      TabIndex        =   9
      Tag             =   "Forma Pago|N|S|||cvfacturas|codforpa|000||"
      Top             =   1560
      Width           =   435
   End
   Begin VB.TextBox txtAux2 
      BackColor       =   &H80000018&
      Enabled         =   0   'False
      Height          =   285
      Index           =   7
      Left            =   9420
      Locked          =   -1  'True
      MaxLength       =   40
      TabIndex        =   36
      Top             =   1230
      Width           =   2655
   End
   Begin VB.TextBox txtAux2 
      BackColor       =   &H80000018&
      Enabled         =   0   'False
      Height          =   285
      Index           =   19
      Left            =   10740
      Locked          =   -1  'True
      MaxLength       =   40
      TabIndex        =   35
      Top             =   1560
      Width           =   1365
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   9930
      TabIndex        =   23
      Tag             =   "   "
      Top             =   5220
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   11040
      TabIndex        =   24
      Top             =   5220
      Visible         =   0   'False
      Width           =   1095
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "frmManFactCV1.frx":0010
      Height          =   4410
      Left            =   90
      TabIndex        =   27
      Top             =   630
      Width           =   9210
      _ExtentX        =   16245
      _ExtentY        =   7779
      _Version        =   393216
      AllowUpdate     =   0   'False
      BorderStyle     =   0
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         AllowFocus      =   0   'False
         AllowRowSizing  =   0   'False
         AllowSizing     =   0   'False
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdRegresar 
      Caption         =   "&Regresar"
      Height          =   375
      Left            =   11040
      TabIndex        =   29
      Top             =   5220
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Height          =   555
      Index           =   1
      Left            =   90
      TabIndex        =   25
      Top             =   5130
      Width           =   2385
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
         Height          =   255
         Left            =   40
         TabIndex        =   26
         Top             =   240
         Width           =   2295
      End
   End
   Begin MSAdodcLib.Adodc adodc1 
      Height          =   375
      Left            =   2205
      Top             =   0
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
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   28
      Top             =   0
      Width           =   12270
      _ExtentX        =   21643
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      Style           =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   12
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Style           =   3
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Buscar"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Ver todos"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Nuevo"
            Object.Tag             =   "2"
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Modificar"
            Object.Tag             =   "2"
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Eliminar"
            Object.Tag             =   "2"
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Style           =   3
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Imprimir"
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Salir"
         EndProperty
      EndProperty
      BorderStyle     =   1
      Begin VB.CheckBox chkVistaPrevia 
         Caption         =   "Vista previa"
         Height          =   195
         Left            =   3735
         TabIndex        =   0
         Top             =   90
         Visible         =   0   'False
         Width           =   1215
      End
   End
   Begin VB.Label Label1 
      Caption         =   "Iva 3"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   225
      Index           =   6
      Left            =   9390
      TabIndex        =   48
      Top             =   3810
      Width           =   645
   End
   Begin VB.Label Label1 
      Caption         =   "Iva 2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   225
      Index           =   5
      Left            =   9420
      TabIndex        =   47
      Top             =   2910
      Width           =   585
   End
   Begin VB.Label Label1 
      Caption         =   "Iva 1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   225
      Index           =   4
      Left            =   9420
      TabIndex        =   46
      Top             =   1980
      Width           =   585
   End
   Begin VB.Label Label5 
      Caption         =   "Cuota"
      Height          =   255
      Left            =   10110
      TabIndex        =   45
      Top             =   4440
      Width           =   585
   End
   Begin VB.Label Label1 
      Caption         =   "Tipo"
      Height          =   255
      Index           =   3
      Left            =   10110
      TabIndex        =   44
      Top             =   4095
      Width           =   345
   End
   Begin VB.Image imgBuscar 
      Height          =   240
      Index           =   4
      Left            =   10470
      ToolTipText     =   "Buscar iva"
      Top             =   4110
      Width           =   240
   End
   Begin VB.Label Label1 
      Caption         =   "Base"
      Height          =   285
      Index           =   2
      Left            =   10110
      TabIndex        =   43
      Top             =   3810
      Width           =   525
   End
   Begin VB.Label Label4 
      Caption         =   "Cuota"
      Height          =   255
      Left            =   10110
      TabIndex        =   42
      Top             =   3480
      Width           =   525
   End
   Begin VB.Label Label1 
      Caption         =   "Tipo"
      Height          =   255
      Index           =   0
      Left            =   10110
      TabIndex        =   41
      Top             =   3195
      Width           =   345
   End
   Begin VB.Image imgBuscar 
      Height          =   240
      Index           =   3
      Left            =   10440
      ToolTipText     =   "Buscar iva"
      Top             =   3180
      Width           =   240
   End
   Begin VB.Image imgBuscar 
      Height          =   240
      Index           =   2
      Left            =   10440
      ToolTipText     =   "Buscar iva"
      Top             =   2250
      Width           =   240
   End
   Begin VB.Image imgBuscar 
      Height          =   240
      Index           =   0
      Left            =   9990
      ToolTipText     =   "Buscar F.Pago"
      Top             =   1590
      Width           =   240
   End
   Begin VB.Image imgBuscar 
      Height          =   240
      Index           =   1
      Left            =   10500
      ToolTipText     =   "Buscar cuenta"
      Top             =   930
      Width           =   240
   End
   Begin VB.Label Label9 
      Caption         =   "Cuenta Base"
      Height          =   285
      Left            =   9420
      TabIndex        =   37
      Top             =   930
      Width           =   1050
   End
   Begin VB.Label Label1 
      Caption         =   "Tipo"
      Height          =   255
      Index           =   9
      Left            =   10110
      TabIndex        =   34
      Top             =   2265
      Width           =   375
   End
   Begin VB.Label Label2 
      Caption         =   "Cuota"
      Height          =   255
      Left            =   10110
      TabIndex        =   33
      Top             =   2580
      Width           =   555
   End
   Begin VB.Label Label1 
      Caption         =   "Base"
      Height          =   255
      Index           =   7
      Left            =   10110
      TabIndex        =   32
      Top             =   1980
      Width           =   645
   End
   Begin VB.Label Label3 
      Caption         =   "CIF"
      Height          =   255
      Left            =   9420
      TabIndex        =   31
      Top             =   630
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "Base"
      Height          =   285
      Index           =   1
      Left            =   10110
      TabIndex        =   30
      Top             =   2910
      Width           =   465
   End
   Begin VB.Label Label13 
      Caption         =   "F.Pago"
      Height          =   285
      Left            =   9420
      TabIndex        =   38
      Top             =   1560
      Width           =   555
   End
   Begin VB.Menu mnOpciones 
      Caption         =   "&Opciones"
      Begin VB.Menu mnBuscar 
         Caption         =   "&Buscar"
         Shortcut        =   ^F
      End
      Begin VB.Menu mnVerTodos 
         Caption         =   "&Ver todos"
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
Attribute VB_Name = "frmManFactCV1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' +-+-+-+-+-+-+-+-+-+-+-+-
' +-+- Autor: MANOLO  +-+-
' +-+-+-+-+-+-+-+-+-+-+-+-

' **************** PER A QUE FUNCIONE EN UN ATRE MANTENIMENT ********************
' 0. Posar-li l'atribut Datasource a "adodc1" del Datagrid1. Canviar el Caption
'    del formulari
' 1. Canviar els TAGs i els Maxlength de TextAux(0) i TextAux(1)
' 2. En PonerModo(vModo) repasar els indexs del botons, per si es canvien
' 3. En la funció BotonAnyadir() canviar la taula i el camp per a SugerirCodigoSiguienteStr
' 4. En la funció BotonBuscar() canviar el nom de la clau primaria
' 5. En la funció BotonEliminar() canviar la pregunta, les descripcions de la
'    variable SQL i el contingut del DELETE
' 6. En la funció PonerLongCampos() posar els camps als que volem canviar el MaxLength quan busquem
' 7. En Form_Load() repasar la barra d'iconos (per si es vol canviar algún) i
'    canviar la consulta per a vore tots els registres
' 8. En Toolbar1_ButtonClick repasar els indexs de cada botó per a que corresponguen
' 9. En la funció CargaGrid canviar l'ORDER BY (normalment per la clau primaria);
'    canviar ademés els noms dels camps, el format i si fa falta la cantitat;
'    repasar els index dels botons modificar i eliminar.
'    NOTA: si en Form_load ya li he posat clausula WHERE, canviar el `WHERE` de
'    `SQL = CadenaConsulta & " WHERE " & vSQL` per un `AND`
' 10. En txtAux_LostFocus canviar el mensage i el format del camp
' 11. En la funció DatosOk() canviar els arguments de DevuelveDesdeBD i el mensage
'    en cas d'error
' 12. En la funció SepuedeBorrar() canviar les comprovacions per a vore si es pot
'    borrar el registre
' *******************************SI N'HI HA COMBO*******************************
' 0. Comprovar que en el SQL de Form_Load() es faça referència a la taula del Combo
' 1. Pegar el Combo1 al  costat dels TextAux. Canviar-li el TAG
' 2. En BotonModificar() canviar el camp del Combo
' 3. En CargaCombo() canviar la consulta i els noms del camps, o posar els valor
'    a ma si no es llig de cap base de datos els valors del Combo

Option Explicit

Public DatosADevolverBusqueda As String    'Tendra el nº de text que quiere que devuelva, empipados
Public Event DatoSeleccionado(CadenaSeleccion As String)

'codi per al registe que s'afegix al cridar des d'atre formulari.
'Obrir en modo Insertar i tornar datos del registre insertat
Public NuevoCodigo As String
Public CodigoActual As String
Public DeConsulta As Boolean

Private CadenaConsulta As String
Private CadenaConsulta1 As String
Private CadB As String

Private WithEvents frmCtas As frmCtasConta 'cuentas contables
Attribute frmCtas.VB_VarHelpID = -1
Private WithEvents frmTIva As frmTipIVAConta 'Tipos de iva de conta
Attribute frmTIva.VB_VarHelpID = -1
Private WithEvents frmFPa As frmForpaConta 'formas de pago de la contabilidad
Attribute frmFPa.VB_VarHelpID = -1
Private WithEvents frmC As frmCal
Attribute frmC.VB_VarHelpID = -1
'Private WithEvents frmTra As frmManTraba 'trabajadores


Private BuscaChekc As String

Dim vSeccion As CSeccion
Dim Modo As Byte
'----------- MODOS --------------------------------
'   0.-  Formulari net sense cap camp ple
'   1.-  Preparant per a fer la búsqueda
'   2.-  Ja tenim registres i els anem a recorrer
'        ,podem editar-los Edició del camp
'   3.-  Inserció de nou registre
'   4.-  Modificar
'--------------------------------------------------
Dim PrimeraVez As Boolean
Dim indice As Byte 'Index del text1 on es poses els datos retornats des d'atres Formularis de Mtos
Dim i As Integer
Dim indCodigo As Integer

Private Sub btnBuscar_Click(Index As Integer)
 TerminaBloquear
    
    Select Case Index
        Case 0 'Cuentas Contables (de contabilidad)
            If vParamAplic.NumeroContaCV = 0 Then Exit Sub
            
            indice = 2
            Set frmCtas = New frmCtasConta
            frmCtas.NumDigit = 0
            frmCtas.DatosADevolverBusqueda = "0|1|"
            frmCtas.CodigoActual = txtAux(indice).Text
            frmCtas.Conexion = cContaCV
            frmCtas.Show vbModal
            Set frmCtas = Nothing
            PonerFoco txtAux(indice)
        
        
        Case 1 ' fecha de factura
            Dim esq As Long
            Dim dalt As Long
            Dim menu As Long
            Dim obj As Object
        
            Set frmC = New frmCal
            
            esq = btnBuscar(Index).Left
            dalt = btnBuscar(Index).Top
                
            Set obj = btnBuscar(Index).Container
              
              While btnBuscar(Index).Parent.Name <> obj.Name
                    esq = esq + obj.Left
                    dalt = dalt + obj.Top
                    Set obj = obj.Container
              Wend
            
            menu = Me.Height - Me.ScaleHeight 'ací tinc el heigth del menú i de la toolbar
        
            frmC.Left = esq + btnBuscar(Index).Parent.Left + 30
            frmC.Top = dalt + btnBuscar(Index).Parent.Top + btnBuscar(Index).Height + menu - 40
        
            btnBuscar(Index).Tag = Index + 1 '<===
            ' *** repasar si el camp es txtAux o Text1 ***
            If txtAux(3).Text <> "" Then frmC.NovaData = txtAux(3).Text
            ' ********************************************
        
            frmC.Show vbModal
            Set frmC = Nothing
            ' *** repasar si el camp es txtAux o Text1 ***
            PonerFoco txtAux(3) '<===
            ' ********************************************
        
            
    End Select
    
    If Modo = 4 Then BLOQUEADesdeFormulario2 Me, Me.adodc1, 1

End Sub

Private Sub chkAbonos_GotFocus(Index As Integer)
    PonerFocoChk Me.chkAbonos(Index)
End Sub

Private Sub chkAbonos_Click(Index As Integer)
    If Modo = 1 Then
        'Buscqueda
        If InStr(1, BuscaChekc, "chkAbonos(" & Index & ")") = 0 Then BuscaChekc = BuscaChekc & "chkAbonos(" & Index & ")|"
    Else
        If Index = 0 And (Modo = 3 Or Modo = 4) Then
            If chkAbonos(Index).Value = 0 Then
                chkAbonos(1).Value = 0
                chkAbonos(1).Enabled = False
            Else
                chkAbonos(1).Enabled = True
            End If
        End If
    End If
End Sub

Private Sub chkAbonos_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub chkAbonos_LostFocus(Index As Integer)
'    If Index = 1 And (Modo = 3 Or Modo = 4) Then
'        If chkAbonos(Index).Value = 1 Then Text1(25).Text = ""
'    End If
End Sub

Private Sub PonerContRegIndicador()
'si estamos en modo ver registros muestra el numero de registro en el que estamos
'situados del total de registros mostrados: 1 de 24
Dim cadReg As String

    If (Modo = 2 Or Modo = 0) Then
        cadReg = PonerContRegistros(Me.adodc1)
        If CadB = "" Then
            lblIndicador.Caption = cadReg
        Else
            lblIndicador.Caption = "BUSQUEDA: " & cadReg
        End If
    End If
End Sub

Private Sub PonerModo(vModo)
Dim b As Boolean

    Modo = vModo
    
    BuscaChekc = ""
    
    b = (Modo = 2)
    If b Then
        PonerContRegIndicador
    Else
        PonerIndicador lblIndicador, Modo
    End If
    
    b = (Modo = 2 Or Modo = 0)
    
    For i = 0 To 3
        txtAux(i).visible = Not b
    Next i
    txtAux(6).visible = Not b
    
    txtAux2(2).visible = Not b
    
    For i = 0 To 3
        If i <> 2 Then
            txtAux(i).Enabled = (Modo = 1 Or Modo = 3)
        Else
            txtAux(i).Enabled = (Modo = 1 Or Modo = 3 Or Modo = 4)
        End If
    Next i
    
    b = (Modo = 2)
    
    ' campos de fuera del grid
    For i = 4 To 19
        If i <> 6 Then
            BloquearTxt txtAux(i), b
        End If
    Next i
    
    
    ' **** si n'hi han camps fora del grid, bloquejar-los ****
    BloquearCmb Me.Combo1(0), b
    Combo1(0).visible = Not (Modo = 0 Or Modo = 2)
    Combo1(0).Enabled = (Modo = 1 Or Modo = 3)

    Me.btnBuscar(0).visible = Not b
    Me.btnBuscar(0).Enabled = Modo = 1 Or Modo = 3 Or Modo = 4
    Me.btnBuscar(1).visible = Not b
    Me.btnBuscar(1).Enabled = Modo = 1 Or Modo = 3
    
    For i = 0 To 4
        Me.imgBuscar(i).Enabled = Not b
        Me.imgBuscar(i).visible = Not b
    Next i
    
    BloquearChk Me.chkAbonos(0), (Modo <> 1)
    BloquearChk Me.chkAbonos(1), (Modo <> 1)
    
    
    
    
'    For i = 0 To Me.cmdAux.Count - 1
'        cmdAux(i).visible = Not b
'        cmdAux(i).Enabled = Not b
'    Next i
    
'    Combo1(0).visible = Not b
'    Combo1(0).Enabled = Not b
'
    
    cmdAceptar.visible = Not b
    cmdCancelar.visible = Not b
    DataGrid1.Enabled = b
    
    'Si es regresar
    If DatosADevolverBusqueda <> "" Then cmdRegresar.visible = b
    
    PonerLongCampos
    PonerModoOpcionesMenu 'Activar/Desact botones de menu segun Modo
    PonerOpcionesMenu  'En funcion del usuario
    
    'Si estamos modo Modificar bloquear clave primaria
    BloquearTxt txtAux(0), (Modo = 4)
End Sub



Private Sub PonerModoOpcionesMenu()
'Activa/Desactiva botones del la toobar y del menu, segun el modo en que estemos
Dim b As Boolean

    b = (Modo = 2)
    'Busqueda
    Toolbar1.Buttons(2).Enabled = b
    Me.mnBuscar.Enabled = b
    'Ver Todos
    Toolbar1.Buttons(3).Enabled = b
    Me.mnVerTodos.Enabled = b
    
    'Insertar
    Toolbar1.Buttons(6).Enabled = b And Not DeConsulta
    Me.mnNuevo.Enabled = b And Not DeConsulta
    
    b = (b And adodc1.Recordset.RecordCount > 0) And Not DeConsulta
    'Modificar
    Toolbar1.Buttons(7).Enabled = b
    Me.mnModificar.Enabled = b
    'Eliminar
    Toolbar1.Buttons(8).Enabled = b
    Me.mnEliminar.Enabled = b
    'Imprimir
    Toolbar1.Buttons(11).Enabled = b
    Me.mnImprimir.Enabled = b
    
End Sub

Private Sub BotonAnyadir()
    Dim NumF As String
    Dim anc As Single
    
'    CargaGrid 'primer de tot carregue tot el grid
'    CadB = ""
    '******************** canviar taula i camp **************************
    '********************************************************************
    'Situamos el grid al final
    AnyadirLinea DataGrid1, adodc1
         
    anc = DataGrid1.Top
    If DataGrid1.Row < 0 Then
        anc = anc + 206
    Else
        anc = anc + DataGrid1.RowTop(DataGrid1.Row) + 5
    End If
    For i = 0 To txtAux.Count - 1
        txtAux(i).Text = ""
    Next i
    Me.chkAbonos(0).Value = 0

    For i = 2 To 2
        txtAux2(i).Text = ""
    Next i
    
    Combo1(0).ListIndex = 0
    
    LLamaLineas anc, 3 'Pone el form en Modo=3, Insertar
       
    'Ponemos el foco
    PonerFocoCmb Combo1(0)
End Sub

Private Sub BotonVerTodos()
    CadB = ""
    CargaGrid ""
    PonerModo 2
End Sub

Private Sub BotonBuscar()
    ' ***************** canviar per la clau primaria ********
    CargaGrid "cvfacturas.numfactu = '-1'"
    '*******************************************************************************
    'Buscar
    For i = 0 To txtAux.Count - 1
        txtAux(i).Text = ""
    Next i

    Me.Combo1(0).ListIndex = -1
    Me.chkAbonos(0).Value = 0
    Me.chkAbonos(1).Value = 0
    
    txtAux2(2).Text = ""
    txtAux2(7).Text = ""
    txtAux2(19).Text = ""
    
'    PosicionarCombo Combo1, "724"
    LLamaLineas DataGrid1.Top + 206, 1 'Pone el form en Modo=1, Buscar
    PonerFocoCmb Combo1(0)
End Sub

Private Sub BotonModificar()
    Dim anc As Single
    Dim i As Integer
    
    Screen.MousePointer = vbHourglass
    
    If DataGrid1.Bookmark < DataGrid1.FirstRow Or DataGrid1.Bookmark > (DataGrid1.FirstRow + DataGrid1.VisibleRows - 1) Then
        i = DataGrid1.Bookmark - DataGrid1.FirstRow
        DataGrid1.Scroll 0, i
        DataGrid1.Refresh
    End If
    
    If DataGrid1.Row < 0 Then
        anc = 320
    Else
        anc = DataGrid1.RowTop(DataGrid1.Row) + 545
    End If

    'Llamamos al form
    txtAux(0).Text = DataGrid1.Columns(2).Text
    txtAux(1).Text = DataGrid1.Columns(3).Text
    txtAux(2).Text = DataGrid1.Columns(19).Text
    txtAux(3).Text = DataGrid1.Columns(4).Text
    txtAux2(2).Text = DataGrid1.Columns(20).Text
    txtAux(6).Text = DataGrid1.Columns(21).Text
    
    

    PosicionarCombo Me.Combo1(0), DataGrid1.Columns(0).Text
    

    LLamaLineas anc, 4 'Pone el form en Modo=4, Modificar
   
    'Como es modificar
    PonerFoco txtAux(2)
    Screen.MousePointer = vbDefault
End Sub

Private Sub LLamaLineas(alto As Single, xModo As Byte)
Dim jj As Integer

    DeseleccionaGrid Me.DataGrid1
    PonerModo xModo
    
    'Fijamos el ancho
    For i = 0 To 3
        txtAux(i).Top = alto
    Next i
    txtAux(6).Top = alto
    txtAux2(2).Top = alto

    For jj = 0 To btnBuscar.Count - 1
        btnBuscar(jj).visible = (Modo = 1 Or Modo = 3 Or Modo = 4)
        btnBuscar(jj).Top = txtAux(3).Top
        btnBuscar(jj).Height = txtAux(3).Height
    Next jj
    Combo1(0).visible = (Modo = 1 Or Modo = 3 Or Modo = 4)
    Combo1(0).Top = txtAux(3).Top

End Sub

Private Sub BotonEliminar()
Dim sql As String
Dim temp As Boolean

    On Error GoTo Error2
    'Ciertas comprobaciones
    If adodc1.Recordset.EOF Then Exit Sub
'    If Not SepuedeBorrar Then Exit Sub
        
    ' ### [Monica] 26/09/2006 dejamos modificar y eliminar el codigo 0
    ' *** repasar el nom de l'adodc, l'index del Field i el camp que te la PK ***
    'El registre de codi 0 no es pot Modificar ni Eliminar
    'If EsCodigoCero(CStr(adodc1.Recordset.Fields(0).Value), FormatoCampo(txtAux(0))) Then Exit Sub
    ' ***************************************************************************
    
    '*************** canviar els noms i el DELETE **********************************
    sql = "¿Seguro que desea eliminar la Factura?"
    sql = sql & vbCrLf & "Tipo: " & adodc1.Recordset.Fields(1)
    sql = sql & vbCrLf & "Serie: " & adodc1.Recordset.Fields(2)
    sql = sql & vbCrLf & "Número: " & adodc1.Recordset.Fields(3)
    sql = sql & vbCrLf & "Fecha Factura: " & adodc1.Recordset.Fields(4)
    
    If MsgBox(sql, vbQuestion + vbYesNo) = vbYes Then
        'Hay que eliminar
        NumRegElim = adodc1.Recordset.AbsolutePosition
        sql = "Delete from cvfacturas where letraser=" & DBSet(adodc1.Recordset!letraser, "T")
        sql = sql & " and numfactu = " & DBSet(adodc1.Recordset!NumFactu, "T")
        sql = sql & " and fecfactu = " & DBSet(adodc1.Recordset!fecfactu, "F")
        sql = sql & " and tipofactu = " & adodc1.Recordset!TipoFactu
        
        conn.Execute sql
        CargaGrid CadB
        temp = SituarDataTrasEliminar(adodc1, NumRegElim, True)
        PonerModoOpcionesMenu
        adodc1.Recordset.Cancel
    End If
    Exit Sub

    
Error2:
    Screen.MousePointer = vbDefault
    MuestraError Err.Number, "Eliminando registro", Err.Description
End Sub

Private Sub PonerLongCampos()
'Modificar el MaxLength del campo en funcion de si es modo de búsqueda o no
'para los campos que permitan introducir criterios más largos del tamaño del campo
'Si maxlength=6 para codprove, en modo busqueda Maxlenth=13 'el doble + 1
    
    'para los txtAux
    PonerLongCamposGnral Me, Modo, 3
End Sub



Private Sub cmdAceptar_Click()
Dim temp As Boolean
Dim i As String

    Select Case Modo
        Case 1 'BUSQUEDA
'            CadB = ObtenerBusqueda(Me)
            CadB = ObtenerBusqueda2(Me, BuscaChekc, 1)
           
            If CadB <> "" Then
                CargaGrid CadB
                PonerModo 2
'                lblIndicador.Caption = "BUSQUEDA: " & PonerContRegistros(Me.adodc1)
                PonerFocoGrid Me.DataGrid1
            End If
            
        Case 3 'INSERTAR
            If DatosOk Then
                If InsertarDesdeForm(Me) Then
                    CargaGrid CadB
                    If (DatosADevolverBusqueda <> "") And NuevoCodigo <> "" Then
                        cmdCancelar_Click
                        If Not adodc1.Recordset.EOF Then
                            SituarDataMULTI adodc1, "tipofactu = " & Combo1(0).ListIndex & " and letraser = '" & txtAux(0).Text & "' and numfactu = " & txtAux(1).Text & " and fecfactu = " & DBSet(txtAux(3).Text, "F"), "" ' Find (adodc1.Recordset.Fields(2).Name & " =" & NuevoCodigo)
                        End If
                        cmdRegresar_Click
                    Else
                        BotonAnyadir
                    End If
                    CadB = ""
                End If
            End If
            
        Case 4 'MODIFICAR
            If DatosOk Then
                If ModificaDesdeFormulario(Me) Then
                    TerminaBloquear
'                    PosicionarData
'                    CargaGrid CadB

                    NumRegElim = adodc1.Recordset.AbsolutePosition
                    CargaGrid CadB
                    temp = SituarDataTrasEliminar(adodc1, NumRegElim, True)
                    PonerModo 2
                    If temp Then
                        CargaForaGrid
                    Else
                        LimpiarCampos
                    End If
                    adodc1.Recordset.Cancel
                    PonerFocoGrid Me.DataGrid1
                End If
            End If
    End Select
End Sub

Private Sub cmdAux_Click(Index As Integer)
    Select Case Index
        Case 0
            If vSeccion Is Nothing Then Exit Sub
            
            indice = 4
            Set frmFPa = New frmForpaConta
            frmFPa.DatosADevolverBusqueda = "0|1|"
            frmFPa.CodigoActual = txtAux(indice)
        '    frmFpa.Conexion = cContaFacSoc
            frmFPa.Show vbModal
            Set frmFPa = Nothing
            PonerFoco txtAux(indice)
        
        Case 1 'cuentas contables de y proveedor
            
            indice = 5
            Set frmCtas = New frmCtasConta
            frmCtas.NumDigit = 0
            frmCtas.DatosADevolverBusqueda = "0|1|"
            frmCtas.CodigoActual = txtAux(indice).Text
            frmCtas.Show vbModal
            Set frmCtas = Nothing
            PonerFoco txtAux(indice)
        
        
     End Select
End Sub

Private Sub cmdCancelar_Click()
    On Error Resume Next
    
    Select Case Modo
        Case 1 'búsqueda
            CargaGrid CadB
        Case 3 'insertar
            DataGrid1.AllowAddNew = False
            'CargaGrid
            If Not adodc1.Recordset.EOF Then adodc1.Recordset.MoveFirst
        Case 4 'modificar
            TerminaBloquear
    End Select
    
    PonerModo 2
    
'    If CadB <> "" Then
'        lblIndicador.Caption = "BUSQUEDA: " & PonerContRegistros(Me.adodc1)
''    Else
''        lblIndicador.Caption = ""
'    End If
    
    PonerFocoGrid Me.DataGrid1
    If Err.Number <> 0 Then Err.Clear
End Sub

Private Sub cmdRegresar_Click()
Dim cad As String
Dim i As Integer
Dim J As Integer
Dim Aux As String

    If adodc1.Recordset.EOF Then
        MsgBox "Ningún registro devuelto.", vbExclamation
        Exit Sub
    End If
    cad = ""
    i = 0
    Do
        J = i + 1
        i = InStr(J, DatosADevolverBusqueda, "|")
        If i > 0 Then
            Aux = Mid(DatosADevolverBusqueda, J, i - J)
            J = Val(Aux)
            cad = cad & adodc1.Recordset.Fields(J) & "|"
        End If
    Loop Until i = 0
    RaiseEvent DatoSeleccionado(cad)
    Unload Me
End Sub

Private Sub Combo1_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub DataGrid1_DblClick()
    If cmdRegresar.visible Then cmdRegresar_Click
End Sub

Private Sub DataGrid1_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub DataGrid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    
    If Modo = 2 Then
        PonerContRegIndicador
        CargaForaGrid
    End If
    
End Sub

Private Sub Form_Activate()
    Screen.MousePointer = vbDefault

    If PrimeraVez Then
        PrimeraVez = False
        If (DatosADevolverBusqueda <> "") And NuevoCodigo <> "" Then
            BotonAnyadir
        Else
            CargaGrid
            PonerModo 2
             If Me.CodigoActual <> "" Then
                SituarData Me.adodc1, "numfactu=" & DBSet(CodigoActual, "T"), "", True
            End If
        End If
    End If
End Sub

Private Sub Form_Load()
    PrimeraVez = True

    With Me.Toolbar1
        .HotImageList = frmPpal.imgListComun_OM
        .DisabledImageList = frmPpal.imgListComun_BN
        .ImageList = frmPpal.imgListComun
        'el 1 es separadors
        .Buttons(2).Image = 1   'Buscar
        .Buttons(3).Image = 2   'Todos
        'el 4 i el 5 son separadors
        .Buttons(6).Image = 3   'Insertar
        .Buttons(7).Image = 4   'Modificar
        .Buttons(8).Image = 5   'Borrar
        'el 9 i el 10 son separadors
        .Buttons(11).Image = 10  'imprimir
        .Buttons(12).Image = 11  'Salir
    End With

    '## A mano
'    chkVistaPrevia.Value = CheckValueLeer(Name)
    
    For i = 0 To Me.imgBuscar.Count - 1
        Me.imgBuscar(i).Picture = frmPpal.imgListImages16.ListImages(1).Picture
    Next i
    
    
    '****************** canviar la consulta *********************************+
    CadenaConsulta = "SELECT cvfacturas.tipofactu, "
    CadenaConsulta = CadenaConsulta & " CASE cvfacturas.tipofactu WHEN 0 THEN ""Varias"" WHEN 1 THEN ""Vta.Tienda"" WHEN 2 THEN ""Compra"" END, "
    CadenaConsulta = CadenaConsulta & "cvfacturas.letraser , cvfacturas.numfactu, cvfacturas.fecfactu, cvfacturas.nifsocio, cvfacturas.codiva, cvfacturas.porciva, "
    CadenaConsulta = CadenaConsulta & "cvfacturas.baseimpo, cvfacturas.cuotaiva,cvfacturas.codiva2, cvfacturas.porciva2,cvfacturas.baseimpo2, cvfacturas.cuotaiva2, "
    CadenaConsulta = CadenaConsulta & "cvfacturas.codiva3, cvfacturas.porciva3,cvfacturas.baseimpo3, cvfacturas.cuotaiva3, "
    CadenaConsulta = CadenaConsulta & "cvfacturas.intconta, cvfacturas.codmactasoc, conta" & vParamAplic.NumeroContaCVV & ".cuentas.nommacta, cvfacturas.totalfac, "
    CadenaConsulta = CadenaConsulta & "cvfacturas.codmactavta, cvfacturas.codforpa, cvfacturas.erronea "
    CadenaConsulta = CadenaConsulta & "FROM cvfacturas left join conta" & DBSet(vParamAplic.NumeroContaCVV, "N")
    CadenaConsulta = CadenaConsulta & ".cuentas on (cvfacturas.codmactasoc = conta" & DBSet(vParamAplic.NumeroContaCVV, "N") & ".cuentas.codmacta)"
    CadenaConsulta = CadenaConsulta & " WHERE 1 = 1 and tipofactu = 0 "
'    CadenaConsulta = CadenaConsulta & " union "
    CadenaConsulta1 = " SELECT cvfacturas.tipofactu, "
    CadenaConsulta1 = CadenaConsulta1 & " CASE cvfacturas.tipofactu WHEN 0 THEN ""Varias"" WHEN 1 THEN ""Vta.Tienda"" WHEN 2 THEN ""Compra"" END, "
    CadenaConsulta1 = CadenaConsulta1 & "cvfacturas.letraser , cvfacturas.numfactu, cvfacturas.fecfactu, cvfacturas.nifsocio, cvfacturas.codiva, cvfacturas.porciva, "
    CadenaConsulta1 = CadenaConsulta1 & "cvfacturas.baseimpo, cvfacturas.cuotaiva,cvfacturas.codiva2, cvfacturas.porciva2,cvfacturas.baseimpo2, cvfacturas.cuotaiva2, "
    CadenaConsulta1 = CadenaConsulta1 & "cvfacturas.codiva3, cvfacturas.porciva3,cvfacturas.baseimpo3, cvfacturas.cuotaiva3, "
    CadenaConsulta1 = CadenaConsulta1 & "cvfacturas.intconta, cvfacturas.codmactasoc, conta" & vParamAplic.NumeroContaCV & ".cuentas.nommacta, cvfacturas.totalfac, "
    CadenaConsulta1 = CadenaConsulta1 & "cvfacturas.codmactavta, cvfacturas.codforpa, cvfacturas.erronea "
    CadenaConsulta1 = CadenaConsulta1 & "FROM cvfacturas left join conta" & DBSet(vParamAplic.NumeroContaCV, "N")
    CadenaConsulta1 = CadenaConsulta1 & ".cuentas on (cvfacturas.codmactasoc = conta" & DBSet(vParamAplic.NumeroContaCV, "N") & ".cuentas.codmacta)"
    CadenaConsulta1 = CadenaConsulta1 & " WHERE 1 = 1 and tipofactu > 0 "
    
    
'    CadenaConsulta = "SELECT cvfacturas.tipofactu, "
'    CadenaConsulta = CadenaConsulta & " CASE cvfacturas.tipofactu WHEN 0 THEN ""Varias"" WHEN 1 THEN ""Vta.Tienda"" WHEN 2 THEN ""Compra"" END, "
'    CadenaConsulta = CadenaConsulta & "cvfacturas.letraser , cvfacturas.numfactu, cvfacturas.fecfactu, cvfacturas.nifsocio, cvfacturas.codiva, cvfacturas.porciva, "
'    CadenaConsulta = CadenaConsulta & "cvfacturas.baseimpo, cvfacturas.cuotaiva,cvfacturas.codiva2, cvfacturas.porciva2,cvfacturas.baseimpo2, cvfacturas.cuotaiva2, "
'    CadenaConsulta = CadenaConsulta & "cvfacturas.codiva3, cvfacturas.porciva3,cvfacturas.baseimpo3, cvfacturas.cuotaiva3, "
'    CadenaConsulta = CadenaConsulta & "cvfacturas.intconta, cvfacturas.codmactasoc, cvfacturas.totalfac, "
'    CadenaConsulta = CadenaConsulta & "cvfacturas.codmactavta, cvfacturas.codforpa, cvfacturas.erronea "
'    CadenaConsulta = CadenaConsulta & "FROM cvfacturas "
'    CadenaConsulta = CadenaConsulta & " WHERE 1 = 1 "
'
    '************************************************************************
    
    CargaCombo
    
    CadB = ""
    CargaGrid
        
End Sub

Private Sub Form_Unload(Cancel As Integer)
'    CheckValueGuardar Me.Name, Me.chkVistaPrevia.Value
    Screen.MousePointer = vbDefault
    If Modo = 4 Then TerminaBloquear
End Sub

Private Sub frmC_Selec(vFecha As Date)
    ' *** repasar si el camp es txtAux o Text1 ***
    txtAux(3).Text = Format(vFecha, "dd/mm/yyyy") '<===
    ' ********************************************
End Sub

Private Sub frmCtas_DatoSeleccionado(CadenaSeleccion As String)
'Cuentas contables de la Contabilidad
    txtAux(indice).Text = RecuperaValor(CadenaSeleccion, 1) 'codiva
    FormateaCampo txtAux(indice)
    txtAux2(indice).Text = RecuperaValor(CadenaSeleccion, 2) 'nomiva
End Sub

Private Sub frmFpa_DatoSeleccionado(CadenaSeleccion As String)
    txtAux(indice).Text = RecuperaValor(CadenaSeleccion, 1) 'codforpa
    FormateaCampo txtAux(indice)
    txtAux2(indice).Text = RecuperaValor(CadenaSeleccion, 2) 'nomforpa
End Sub

Private Sub frmTIva_DatoSeleccionado(CadenaSeleccion As String)
'Mantenimiento Tipo de iva
    txtAux(indice).Text = RecuperaValor(CadenaSeleccion, 1) 'codiva
    FormateaCampo txtAux(indice)
    
    Select Case indice
        Case 10
            txtAux(9).Text = RecuperaValor(CadenaSeleccion, 2) 'porceiva
        Case 11
            txtAux(12).Text = RecuperaValor(CadenaSeleccion, 2) 'porceiva
        Case 15
            txtAux(16).Text = RecuperaValor(CadenaSeleccion, 2) 'porceiva
    End Select
End Sub

Private Sub imgBuscar_Click(Index As Integer)
    Select Case Index
        Case 0
            indice = 19
            Set frmFPa = New frmForpaConta
            frmFPa.DatosADevolverBusqueda = "0|1|"
            frmFPa.CodigoActual = txtAux(indice)
            frmFPa.Conexion = cContaCV
            frmFPa.Show vbModal
            Set frmFPa = Nothing
            PonerFoco txtAux(indice)
        
        Case 1 'cuenta contable de ventas
            If vParamAplic.NumeroContaCV = 0 Then Exit Sub
            
            indice = 7
            Set frmCtas = New frmCtasConta
            frmCtas.NumDigit = 0
            frmCtas.DatosADevolverBusqueda = "0|1|"
            frmCtas.CodigoActual = txtAux(indice).Text
            If Combo1(1).ListIndex = 0 Then
                frmCtas.Conexion = cContaCVV
            Else
                frmCtas.Conexion = cContaCV
            End If
            frmCtas.Show vbModal
            Set frmCtas = Nothing
            PonerFoco txtAux(indice)
            
        Case 2, 3 'codigo de iva 1 y 2
            indice = Index + 8
            AbrirFrmTipIvaConta (indice)
            
        Case 4 ' codigo de iva 3
            indice = 15
            AbrirFrmTipIvaConta (indice)
        
     End Select
End Sub

Private Sub mnBuscar_Click()
    BotonBuscar
End Sub

Private Sub mnEliminar_Click()
    BotonEliminar
End Sub

Private Sub mnImprimir_Click()
    frmListFactCV.OpcionListado = 0 ' Listado de Facturas Coarval
    frmListFactCV.Show vbModal
End Sub

Private Sub mnModificar_Click()
    'Comprobaciones
    '--------------
    If adodc1.Recordset.EOF Then Exit Sub
    
    If adodc1.Recordset.RecordCount < 1 Then Exit Sub
    
    ' ### [Monica] 26/09/2006 dejamos modificar y eliminar el codigo 0
    ' *** repasar el nom de l'adodc, l'index del Field i el camp que te la PK ***
    'El registre de codi 0 no es pot Modificar ni Eliminar
    'If EsCodigoCero(CStr(adodc1.Recordset.Fields(0).Value), FormatoCampo(txtAux(0))) Then Exit Sub
    
    
    'Preparamos para modificar
    '-------------------------
    If BLOQUEADesdeFormulario2(Me, adodc1, 1) Then BotonModificar
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
    Select Case Button.Index
        Case 2
                mnBuscar_Click
        Case 3
                mnVerTodos_Click
        Case 6
                mnNuevo_Click
        Case 7
                mnModificar_Click
        Case 8
                mnEliminar_Click
        Case 11
                'MsgBox "Imprimir...under construction"
                mnImprimir_Click
        Case 12
                mnSalir_Click
    End Select
End Sub

Private Sub CargaGrid(Optional vSQL As String)
    Dim sql As String
    Dim tots As String
    
'    adodc1.ConnectionString = Conn
    If vSQL <> "" Then
        sql = CadenaConsulta & " AND " & vSQL
        sql = sql & " union "
        sql = sql & CadenaConsulta1 & " and " & vSQL
    Else
        sql = CadenaConsulta
        sql = sql & " union "
        sql = sql & CadenaConsulta1
    End If
    '********************* canviar el ORDER BY *********************++
    sql = sql & " ORDER BY 3, 4"
    '**************************************************************++
    
    CargaGridGnral Me.DataGrid1, Me.adodc1, sql, PrimeraVez
    
    ' *******************canviar els noms i si fa falta la cantitat********************
    tots = "N||||0|;S|Combo1(0)|C|Tipo|1180|;"
    tots = tots & "S|txtAux(0)|T|Ser.|450|;S|txtAux(1)|T|Factura|1000|;"
    tots = tots & "S|txtAux(3)|T|Fecha|1150|;"
    tots = tots & "S|btnBuscar(1)|B|||;"
    
    tots = tots & "N||||0|;N||||0|;N||||0|;N||||0|;N||||0|;N||||0|;N||||0|;"
    tots = tots & "N||||0|;N||||0|;N||||0|;N||||0|;N||||0|;N||||0|;N||||0|;"
    tots = tots & "S|txtAux(2)|T|Cta.Cl/Pr|1000|;"
    tots = tots & "S|btnBuscar(0)|B|||;S|txtAux2(2)|T|Nombre|2600|;"
    tots = tots & "S|txtAux(6)|T|Total Factura|1250|;"
    tots = tots & "N||||0|;N||||0|;N||||0|;"
    
    arregla tots, DataGrid1, Me
    
    DataGrid1.ScrollBars = dbgAutomatic
    DataGrid1.Columns(0).Alignment = dbgLeft

    If (Not adodc1.Recordset.EOF) Then 'per a que pose els valors de les arees de text la primera volta
        
        CargaForaGrid
        
    Else
        txtAux2(2).Text = ""
        txtAux2(7).Text = ""
        txtAux2(19).Text = ""
    End If

'   DataGrid1.Columns(2).Alignment = dbgRight
End Sub

Private Sub txtAux_GotFocus(Index As Integer)
    ConseguirFocoLin txtAux(Index)
End Sub


Private Sub txtAux_LostFocus(Index As Integer)
    If Not PerderFocoGnral(txtAux(Index), Modo) Then Exit Sub
    
    Select Case Index
        Case 0
            If txtAux(0).Text <> "" Then txtAux(0).Text = UCase(txtAux(0).Text)
        Case 10, 11, 15
            PonerFormatoEntero txtAux(Index)
        Case 3
            PonerFormatoFecha txtAux(Index)
        Case 2, 7 'cuenta contable
            If txtAux(Index).Text = "" Then Exit Sub
            txtAux2(Index).Text = PonerNombreCuenta(txtAux(Index), Modo, , cContaCV)
        Case 4, 5, 6, 9, 12, 13, 14, 16, 17, 18
            If Modo = 1 Then Exit Sub
            PonerFormatoDecimal txtAux(Index), 3
        Case 19 ' forma de pago
            If txtAux(Index).Text <> "" Then
                txtAux2(Index) = DevuelveDesdeBDNew(cContaCV, "sforpa", "nomforpa", "codforpa", txtAux(Index).Text, "N")
                If txtAux2(Index).Text = "" Then
                    MsgBox "Forma de pago no existe. Reintroduzca.", vbExclamation
                    PonerFoco txtAux(Index)
                End If
            End If
        
    End Select
    
End Sub

Private Function DatosOk() As Boolean
'Dim Datos As String
Dim b As Boolean
Dim sql As String
Dim Mens As String
Dim Cta As String
Dim cadMen As String

    b = CompForm(Me)
    If Not b Then Exit Function
    
    If Modo = 3 Then   'Estamos insertando
        sql = "select count(*) from cvfacturas where tipofactu = " & DBSet(Combo1(0).ListIndex, "N") & " and letraser = " & DBSet(txtAux(0).Text, "T")
        sql = sql & " and numfactu = " & DBSet(txtAux(1).Text, "N") & " and fecfactu = " & DBSet(txtAux(3).Text, "F")
        If TotalRegistros(sql) > 1 Then
            MsgBox "Ya existe esta factura. Revise.", vbExclamation
            b = False
        End If
    End If
    
    If b And (Modo = 3 Or Modo = 4) Then
        If txtAux(1).Text = "" Then
            MsgBox "El Número de factura debe tener un valor. Revise.", vbExclamation
            PonerFoco txtAux(1)
            b = False
        End If
        ' comprobamos que cuadran los importes bases + ivas = total factura
        If b Then
            If Not FacturaCorrecta Then
                MsgBox "Las suma de bases más ivas no coincide con el total factura. Revise.", vbExclamation
                b = False
            End If
        End If
        If Modo = 4 And b Then Me.chkAbonos(1).Value = 0
    End If
    
        
    DatosOk = b
End Function

Private Function FacturaCorrecta() As Boolean

Dim c_Base1 As Currency
Dim c_Base2 As Currency
Dim c_Base3 As Currency
Dim c_Iva1 As Currency
Dim c_Iva2 As Currency
Dim c_Iva3 As Currency
Dim c_Total As Currency


    On Error Resume Next
    
    c_Base1 = CCur(ImporteSinFormato(ComprobarCero(txtAux(4).Text)))
    c_Base2 = CCur(ImporteSinFormato(ComprobarCero(txtAux(14).Text)))
    c_Base3 = CCur(ImporteSinFormato(ComprobarCero(txtAux(18).Text)))
    c_Iva1 = CCur(ImporteSinFormato(ComprobarCero(txtAux(5).Text)))
    c_Iva2 = CCur(ImporteSinFormato(ComprobarCero(txtAux(13).Text)))
    c_Iva3 = CCur(ImporteSinFormato(ComprobarCero(txtAux(17).Text)))
    c_Total = CCur(ImporteSinFormato(ComprobarCero(txtAux(6).Text)))

    FacturaCorrecta = ((c_Base1 + c_Base2 + c_Base3 + c_Iva1 + c_Iva2 + c_Iva3) = c_Total)

End Function


Private Sub PonerOpcionesMenu()
    PonerOpcionesMenuGeneral Me
    PonerOpcionesMenuGeneralNew Me
End Sub


Private Sub printNou()
    With frmImprimir2
        .cadTabla2 = "cvfacturas"
        .Informe2 = "rManCV.rpt"
        If CadB <> "" Then
            '.cadRegSelec = Replace(SQL2SF(CadB), "clientes", "clientes_1")
            .cadRegSelec = SQL2SF(CadB)
        Else
            .cadRegSelec = ""
        End If
        ' *** repasar el nom de l'adodc ***
        '.cadRegActua = Replace(POS2SF(Data1, Me), "clientes", "clientes_1")
        .cadRegActua = POS2SF(adodc1, Me)
        ' *** repasar codEmpre ***
        .cadTodosReg = ""
        '.cadTodosReg = "{itinerar.codempre} = " & codEmpre
        ' *** repasar si li pose ordre o no ****
        .OtrosParametros2 = "pEmpresa='" & vEmpresa.nomEmpre & "'|" '|pOrden={cvfacturas.codforpa}|"
        '.OtrosParametros2 = "pEmpresa='" & vEmpresa.NomEmpre & "'|"
        ' *** posar el nº de paràmetres que he posat en OtrosParametros2 ***
        '.NumeroParametros2 = 1
        .NumeroParametros2 = 2
        ' ******************************************************************
        .MostrarTree2 = False
        .InfConta2 = False
        .ConSubInforme2 = True
        .Contabilidad2 = cContaCV
        .Show vbModal
    End With
End Sub

'' ### [DavidV] 26/04/2006: Activar/desactivar la rueda del ratón.
'Private Sub DataGrid1_GotFocus()
'  WheelHook DataGrid1
'End Sub
'Private Sub DataGrid1_Lostfocus()
'  WheelUnHook
'End Sub

Private Sub txtAux_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub txtAux_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
'Alvançar/Retrocedir els camps en les fleches de desplaçament del teclat.
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

Private Sub CargaForaGrid()
Dim i As Integer

    If DataGrid1.Columns.Count <= 2 Then Exit Sub
    
    ' *** posar als camps de fora del grid el valor de la columna corresponent ***
    
    'Llamamos al form
    txtAux(2).Text = DataGrid1.Columns(19).Text
    
    txtAux(8).Text = DataGrid1.Columns(5).Text ' nif socio
    txtAux(4).Text = DataGrid1.Columns(8).Text ' baseimponible
    txtAux(10).Text = DataGrid1.Columns(6).Text ' codigoiva
    txtAux(9).Text = DataGrid1.Columns(7).Text ' porcentaje iva
    txtAux(5).Text = DataGrid1.Columns(9).Text ' importe iva
    
    txtAux(14).Text = DataGrid1.Columns(12).Text ' baseimponible 2
    txtAux(11).Text = DataGrid1.Columns(10).Text ' codigoiva 2
    txtAux(12).Text = DataGrid1.Columns(11).Text ' porcentaje iva 2
    txtAux(13).Text = DataGrid1.Columns(13).Text ' importe iva 2
    
    txtAux(18).Text = DataGrid1.Columns(16).Text ' baseimponible 3
    txtAux(15).Text = DataGrid1.Columns(14).Text ' codigoiva 3
    txtAux(16).Text = DataGrid1.Columns(15).Text ' porcentaje iva 3
    txtAux(17).Text = DataGrid1.Columns(17).Text ' importe iva 3
    
    txtAux(19).Text = DataGrid1.Columns(23).Text ' forma de pago
    
    txtAux(7).Text = DataGrid1.Columns(22).Text ' cuenta contable ventas
    
    If DataGrid1.Columns(18).Text <> "" Then
        Me.chkAbonos(0).Value = DataGrid1.Columns(18).Text
    End If
    
    If DataGrid1.Columns(24).Text <> "" Then
        Me.chkAbonos(1).Value = DataGrid1.Columns(24).Text
    End If
    
    If DataGrid1.Columns(0).Text = 0 Then
'        txtAux2(2).Text = ""
'        If txtAux(2).Text <> "" Then
'            txtAux2(2) = DevuelveDesdeBDNew(cContaCVV, "cuentas", "nommacta", "codmacta", txtAux(2).Text, "T")
'        End If
        
        txtAux2(7).Text = ""
        If txtAux(7).Text <> "" Then
            txtAux2(7) = DevuelveDesdeBDNew(cContaCVV, "cuentas", "nommacta", "codmacta", txtAux(7).Text, "T")
        End If
        
        txtAux2(19).Text = ""
        If txtAux(19).Text <> "" Then
            txtAux2(19) = DevuelveDesdeBDNew(cContaCVV, "sforpa", "nomforpa", "codforpa", txtAux(19).Text, "N")
        End If
    Else
'        txtAux2(2).Text = ""
'        If txtAux(2).Text <> "" Then
'            txtAux2(2) = DevuelveDesdeBDNew(cContaCV, "cuentas", "nommacta", "codmacta", txtAux(2).Text, "T")
'        End If
        
        txtAux2(7).Text = ""
        If txtAux(7).Text <> "" Then
            txtAux2(7) = DevuelveDesdeBDNew(cContaCV, "cuentas", "nommacta", "codmacta", txtAux(7).Text, "T")
        End If
        
        txtAux2(19).Text = ""
        If txtAux(19).Text <> "" Then
            txtAux2(19) = DevuelveDesdeBDNew(cContaCV, "sforpa", "nomforpa", "codforpa", txtAux(19).Text, "N")
        End If
    
    
    End If
    PonerCamposForma Me, Me.adodc1
    
    If ComprobarCero(txtAux(5).Text) = 0 Then txtAux(5).Text = ""
    If ComprobarCero(txtAux(13).Text) = 0 Then txtAux(13).Text = ""
    If ComprobarCero(txtAux(14).Text) = 0 Then txtAux(14).Text = ""
    If ComprobarCero(txtAux(17).Text) = 0 Then txtAux(17).Text = ""
    If ComprobarCero(txtAux(18).Text) = 0 Then txtAux(18).Text = ""
    
    
End Sub

Private Sub CargaCombo()

    On Error GoTo ErrCarga
    
    'Tipo de factura
    Combo1(0).Clear
    
    Combo1(0).AddItem "Varias"
    Combo1(0).ItemData(Combo1(0).NewIndex) = 0
    Combo1(0).AddItem "Vta.Tienda"
    Combo1(0).ItemData(Combo1(0).NewIndex) = 1
    Combo1(0).AddItem "Compra"
    Combo1(0).ItemData(Combo1(0).NewIndex) = 2
    
ErrCarga:
    If Err.Number <> 0 Then MuestraError Err.Number, "Cargar datos combo.", Err.Description
End Sub

Private Sub AbrirFrmTipIvaConta(indice As Integer)
    indCodigo = indice
    Set frmTIva = New frmTipIVAConta
    frmTIva.DatosADevolverBusqueda = "0|2|"
    frmTIva.CodigoActual = txtAux(indCodigo)
    If Combo1(1).ListIndex = 0 Then
        frmTIva.Conexion = cContaCVV
    Else
        frmTIva.Conexion = cContaCV
    End If
    frmTIva.Show vbModal
    Set frmTIva = Nothing
End Sub

Private Sub LimpiarDataGrids()
'Pone los Grids sin datos, apuntando a ningún registro
On Error Resume Next
    
    PrimeraVez = True
    CargaGrid CadB
    If Err.Number <> 0 Then Err.Clear
End Sub



Private Sub PosicionarData()
Dim Indicador As String
Dim vWhere As String

    If Not Me.adodc1.Recordset.EOF Then
        'Hay datos en el Data1 bien porque se ha hecho VerTodos o una Busqueda
         vWhere = "(" & ObtenerWhereCP(False) & ")"
         If SituarDataMULTI(adodc1, vWhere, Indicador) Then
             PonerModo 2
             lblIndicador.Caption = Indicador
        Else
             LimpiarCampos
             'Poner los grid sin apuntar a nada
             LimpiarDataGrids
             PonerModo 2
         End If
    Else
        'El Data esta vacio, desde el modo de inicio se pulsa Insertar
'        PonerCadenaBusqueda
        CargaGrid CadB
    End If
End Sub


Private Function ObtenerWhereCP(conWhere As Boolean) As String
Dim sql As String

    On Error Resume Next
    
    sql = " tipofactu= " & Combo1(0).ListIndex & ""
    sql = sql & " and letraser = '" & Trim(txtAux(0).Text) & "'"
    sql = sql & " and numfactu = '" & Trim(txtAux(1).Text) & "'"
    sql = sql & " and fecfactu = " & DBSet(txtAux(3).Text, "F")

    If conWhere Then sql = " WHERE " & sql
    ObtenerWhereCP = sql
    
    If Err.Number <> 0 Then MuestraError Err.Number, "Obteniendo cadena WHERE.", Err.Description
End Function


Private Sub LimpiarCampos()
Dim i As Integer

    On Error Resume Next
    
    Limpiar Me   'Metodo general: Limpia los controles TextBox
    lblIndicador.Caption = ""
    'Aqui va el especifico de cada form es
    '### a mano
    Me.Combo1(0).ListIndex = -1
    Me.chkAbonos(0).Value = 0
    Me.chkAbonos(1).Value = 0
    txtAux2(2).Text = ""
    txtAux2(5).Text = ""
    txtAux2(7).Text = ""
    
    If Err.Number <> 0 Then Err.Clear
End Sub


