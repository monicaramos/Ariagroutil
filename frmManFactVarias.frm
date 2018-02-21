VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmManFactVarias 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Facturas Varias"
   ClientHeight    =   8430
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12555
   Icon            =   "frmManFactVarias.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8430
   ScaleWidth      =   12555
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      Caption         =   "Retención"
      ForeColor       =   &H00972E0B&
      Height          =   645
      Left            =   225
      TabIndex        =   82
      Top             =   4275
      Width           =   12210
      Begin VB.TextBox text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   26
         Left            =   855
         MaxLength       =   6
         TabIndex        =   36
         Tag             =   "% Ret|N|S|0|100.00|cabfact|retfaccl|##0.00|N|"
         Text            =   "99.99"
         Top             =   225
         Width           =   645
      End
      Begin VB.TextBox text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   27
         Left            =   3240
         MaxLength       =   10
         TabIndex        =   37
         Tag             =   "Cta.Contable|T|S|||cabfact|cuereten|||"
         Text            =   "1234567890"
         Top             =   225
         Width           =   1035
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   27
         Left            =   4320
         TabIndex        =   83
         Top             =   225
         Width           =   2895
      End
      Begin VB.TextBox text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   28
         Left            =   8775
         MaxLength       =   15
         TabIndex        =   38
         Tag             =   "Importe Retención|N|S|||cabfact|trefaccl|#,###,###,##0.00|N|"
         Top             =   225
         Width           =   1635
      End
      Begin VB.Label Label1 
         Caption         =   "% Ret."
         Height          =   255
         Index           =   12
         Left            =   345
         TabIndex        =   86
         Top             =   225
         Width           =   615
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   6
         Left            =   2970
         Tag             =   "-1"
         ToolTipText     =   "Buscar Cta Contable"
         Top             =   225
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Cta.Contable"
         Height          =   255
         Index           =   17
         Left            =   2025
         TabIndex        =   85
         Top             =   225
         Width           =   945
      End
      Begin VB.Label Label1 
         Caption         =   "Importe Retención"
         Height          =   255
         Index           =   18
         Left            =   7335
         TabIndex        =   84
         Top             =   225
         Width           =   1545
      End
   End
   Begin VB.Frame Frame2 
      Height          =   2145
      Index           =   0
      Left            =   240
      TabIndex        =   53
      Top             =   480
      Width           =   12195
      Begin VB.TextBox Text2 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   34
         Left            =   4320
         TabIndex        =   93
         Top             =   1755
         Width           =   1770
      End
      Begin VB.TextBox text1 
         Height          =   285
         Index           =   34
         Left            =   3960
         MaxLength       =   2
         TabIndex        =   12
         Tag             =   "Cod.Pais|T|S|||cabfact|codpais|||"
         Text            =   "12"
         Top             =   1755
         Width           =   315
      End
      Begin VB.TextBox text1 
         Height          =   285
         Index           =   33
         Left            =   1035
         MaxLength       =   30
         TabIndex        =   11
         Tag             =   "Provincia|T|S|||cabfact|desprovi|||"
         Text            =   "1234567890"
         Top             =   1755
         Width           =   2115
      End
      Begin VB.TextBox text1 
         Height          =   285
         Index           =   32
         Left            =   2610
         MaxLength       =   50
         TabIndex        =   10
         Tag             =   "Poblacion|T|S|||cabfact|despobla|||"
         Text            =   "1234567890"
         Top             =   1440
         Width           =   3465
      End
      Begin VB.TextBox text1 
         Height          =   285
         Index           =   31
         Left            =   1035
         MaxLength       =   6
         TabIndex        =   9
         Tag             =   "C.Postal|T|S|||cabfact|codposta|||"
         Text            =   "123456"
         Top             =   1440
         Width           =   765
      End
      Begin VB.TextBox text1 
         Height          =   285
         Index           =   30
         Left            =   1035
         MaxLength       =   50
         TabIndex        =   7
         Tag             =   "Direccion|T|S|||cabfact|dirdatos|||"
         Text            =   "1234567890"
         Top             =   1125
         Width           =   3420
      End
      Begin VB.TextBox text1 
         Height          =   285
         Index           =   29
         Left            =   5040
         MaxLength       =   15
         TabIndex        =   8
         Tag             =   "Nif|T|S|||cabfact|nifdatos|||"
         Text            =   "1234567890"
         Top             =   1125
         Width           =   1035
      End
      Begin VB.CheckBox Check1 
         Alignment       =   1  'Right Justify
         Height          =   255
         Index           =   0
         Left            =   11625
         TabIndex        =   14
         Tag             =   "Contabilizada|N|N|0|1|cabfact|pasaridoc|||"
         Top             =   540
         Width           =   255
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   0
         Left            =   6975
         TabIndex        =   80
         Top             =   450
         Width           =   3255
      End
      Begin VB.TextBox text1 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   25
         Left            =   6300
         MaxLength       =   10
         TabIndex        =   13
         Tag             =   "Forma de Pago|N|N|||cabfact|codforpa|000||"
         Top             =   450
         Width           =   675
      End
      Begin VB.TextBox text1 
         Height          =   915
         Index           =   5
         Left            =   6300
         MaxLength       =   250
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   16
         Tag             =   "Observaciones|T|S|||cabfact|observac|||"
         Top             =   1125
         Width           =   5640
      End
      Begin VB.CheckBox Check1 
         Height          =   255
         Index           =   1
         Left            =   11700
         TabIndex        =   15
         Tag             =   "Contabilizada|N|N|0|1|cabfact|intconta|||"
         Top             =   225
         Width           =   255
      End
      Begin VB.TextBox text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   4
         Left            =   1440
         MaxLength       =   10
         TabIndex        =   5
         Tag             =   "Cta.Contable|T|N|||cabfact|ctaclien|||"
         Text            =   "1234567890"
         Top             =   765
         Width           =   1035
      End
      Begin VB.TextBox text1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000013&
         Height          =   285
         Index           =   1
         Left            =   4020
         MaxLength       =   7
         TabIndex        =   3
         Tag             =   "Nº de Factura|N|S|0|9999999|cabfact|numfactu|0000000|S|"
         Top             =   450
         Width           =   795
      End
      Begin VB.TextBox text1 
         Alignment       =   2  'Center
         BackColor       =   &H80000013&
         Height          =   285
         Index           =   3
         Left            =   120
         MaxLength       =   2
         TabIndex        =   1
         Tag             =   "Seccion|N|N|0|99|cabfact|codsecci|00|S|"
         Top             =   450
         Width           =   900
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   3
         Left            =   1035
         TabIndex        =   57
         Top             =   450
         Width           =   2445
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   4
         Left            =   2610
         MaxLength       =   60
         TabIndex        =   6
         Tag             =   "Nombre Cuenta|T|S|||cabfact|nommacta|||"
         Top             =   765
         Width           =   3480
      End
      Begin VB.TextBox text1 
         BackColor       =   &H80000013&
         Height          =   285
         Index           =   0
         Left            =   3555
         MaxLength       =   3
         TabIndex        =   2
         Tag             =   "Letra Serie|T|S|||cabfact|letraser||S|"
         Top             =   450
         Width           =   405
      End
      Begin VB.TextBox text1 
         BackColor       =   &H80000013&
         Height          =   285
         Index           =   2
         Left            =   4905
         MaxLength       =   10
         TabIndex        =   4
         Tag             =   "Fecha Factura|F|N|||cabfact|fecfactu|dd/mm/yyyy|S|"
         Top             =   450
         Width           =   1170
      End
      Begin VB.Label Label1 
         Caption         =   "País"
         Height          =   255
         Index           =   24
         Left            =   3240
         TabIndex        =   94
         Top             =   1800
         Width           =   390
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   7
         Left            =   3645
         Tag             =   "-1"
         ToolTipText     =   "Buscar País"
         Top             =   1800
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Provincia"
         Height          =   255
         Index           =   23
         Left            =   135
         TabIndex        =   92
         Top             =   1800
         Width           =   720
      End
      Begin VB.Label Label1 
         Caption         =   "Población"
         Height          =   255
         Index           =   22
         Left            =   1845
         TabIndex        =   91
         Top             =   1440
         Width           =   720
      End
      Begin VB.Label Label1 
         Caption         =   "C.Postal"
         Height          =   255
         Index           =   21
         Left            =   135
         TabIndex        =   90
         Top             =   1485
         Width           =   945
      End
      Begin VB.Label Label1 
         Caption         =   "Domicilio"
         Height          =   255
         Index           =   20
         Left            =   135
         TabIndex        =   89
         Top             =   1125
         Width           =   945
      End
      Begin VB.Label Label1 
         Caption         =   "NIF"
         Height          =   255
         Index           =   19
         Left            =   4590
         TabIndex        =   88
         Top             =   1125
         Width           =   360
      End
      Begin VB.Label Label1 
         Caption         =   "Aridoc"
         Height          =   225
         Index           =   9
         Left            =   11115
         TabIndex        =   87
         Top             =   555
         Width           =   585
      End
      Begin VB.Label Label1 
         Caption         =   "Forma Pago"
         Height          =   255
         Index           =   5
         Left            =   6300
         TabIndex        =   81
         Top             =   180
         Width           =   900
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   5
         Left            =   7245
         Tag             =   "-1"
         ToolTipText     =   "Buscar Forma de Pago"
         Top             =   180
         Width           =   240
      End
      Begin VB.Image imgZoom 
         Height          =   240
         Index           =   0
         Left            =   7425
         ToolTipText     =   "Zoom descripción"
         Top             =   810
         Width           =   240
      End
      Begin VB.Label Label29 
         Caption         =   "Observaciones"
         Height          =   255
         Left            =   6300
         TabIndex        =   78
         Top             =   855
         Width           =   1125
      End
      Begin VB.Label Label1 
         Caption         =   "Contabilizada"
         Height          =   225
         Index           =   7
         Left            =   10650
         TabIndex        =   75
         Top             =   255
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Sección"
         Height          =   255
         Index           =   13
         Left            =   135
         TabIndex        =   74
         Top             =   195
         Width           =   645
      End
      Begin VB.Label Label1 
         Caption         =   "Nº Factura"
         Height          =   255
         Index           =   4
         Left            =   4020
         TabIndex        =   58
         Top             =   180
         Width           =   855
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   0
         Left            =   810
         MousePointer    =   4  'Icon
         Tag             =   "-1"
         ToolTipText     =   "Buscar Sección"
         Top             =   180
         Width           =   240
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   2
         Left            =   5805
         Picture         =   "frmManFactVarias.frx":000C
         ToolTipText     =   "Buscar fecha"
         Top             =   135
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   1
         Left            =   1125
         Tag             =   "-1"
         ToolTipText     =   "Buscar Cta Contable"
         Top             =   765
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Serie"
         Height          =   255
         Index           =   2
         Left            =   3510
         TabIndex        =   56
         Top             =   180
         Width           =   450
      End
      Begin VB.Label Label1 
         Caption         =   "Fecha Fact."
         Height          =   255
         Index           =   1
         Left            =   4950
         TabIndex        =   55
         Top             =   180
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Cta.Contable"
         Height          =   255
         Index           =   3
         Left            =   135
         TabIndex        =   54
         Top             =   810
         Width           =   945
      End
   End
   Begin VB.Frame FrameAux0 
      Caption         =   "Lineas Factura"
      ForeColor       =   &H00972E0B&
      Height          =   2760
      Left            =   225
      TabIndex        =   65
      Top             =   4995
      Width           =   12225
      Begin VB.TextBox txtAux 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   290
         Index           =   10
         Left            =   10140
         MaxLength       =   15
         TabIndex        =   47
         Tag             =   "Precio|N|S|||linfact|precio|###,##0.0000||"
         Text            =   "precio"
         Top             =   1920
         Visible         =   0   'False
         Width           =   675
      End
      Begin VB.TextBox txtAux 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   290
         Index           =   9
         Left            =   9330
         MaxLength       =   15
         TabIndex        =   46
         Tag             =   "Cantidad|N|S|||linfact|cantidad|##,###,##0.00||"
         Text            =   "cantidad"
         Top             =   1920
         Visible         =   0   'False
         Width           =   765
      End
      Begin VB.TextBox txtAux 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   285
         Index           =   8
         Left            =   5160
         TabIndex        =   79
         Tag             =   "Iva|N|N|0|99|linfact|tipoiva|00||"
         Top             =   1920
         Visible         =   0   'False
         Width           =   630
      End
      Begin VB.TextBox txtAux 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   290
         Index           =   6
         Left            =   5850
         MaxLength       =   50
         TabIndex        =   45
         Tag             =   "Ampliación|T|S|||linfact|ampliaci|||"
         Text            =   "Ampliacion"
         Top             =   1920
         Visible         =   0   'False
         Width           =   3435
      End
      Begin VB.TextBox txtAux 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   290
         Index           =   0
         Left            =   135
         MaxLength       =   10
         TabIndex        =   49
         Tag             =   "Seccion|N|N|||linfact|codsecci|00|S|"
         Text            =   "Seccion"
         Top             =   1920
         Visible         =   0   'False
         Width           =   795
      End
      Begin VB.CommandButton btnBuscar 
         Appearance      =   0  'Flat
         Caption         =   "+"
         Height          =   300
         Index           =   0
         Left            =   3420
         MaskColor       =   &H00000000&
         TabIndex        =   70
         ToolTipText     =   "Buscar Concepto"
         Top             =   1920
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.TextBox txtAux 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   290
         Index           =   5
         Left            =   2940
         MaxLength       =   6
         TabIndex        =   44
         Tag             =   "Concepto|N|N|0|999|linfact|codconce|000||"
         Text            =   "Concep"
         Top             =   1920
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.TextBox txtAux 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   290
         Index           =   1
         Left            =   1020
         MaxLength       =   3
         TabIndex        =   0
         Tag             =   "Letra Serie|T|N|||linfact|letraser||S|"
         Text            =   "L"
         Top             =   1920
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.TextBox txtAux 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   290
         Index           =   4
         Left            =   2580
         MaxLength       =   2
         TabIndex        =   43
         Tag             =   "Número de línea|N|N|1|99|linfact|numlinea|00|S|"
         Text            =   "li"
         Top             =   1920
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.TextBox txtAux 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   290
         Index           =   2
         Left            =   1380
         MaxLength       =   7
         TabIndex        =   41
         Tag             =   "Nº Factura|N|N|0|9999999|linfact|numfactu|0000000|S|"
         Text            =   "Fac"
         Top             =   1920
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.TextBox txtAux 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   290
         Index           =   3
         Left            =   1980
         MaxLength       =   10
         TabIndex        =   42
         Tag             =   "Fecha Factura|F|N|||linfact|fecfactu|dd/mm/yyyy|S|"
         Text            =   "fecfactu"
         Top             =   1920
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.TextBox txtAux 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   290
         Index           =   7
         Left            =   10860
         MaxLength       =   15
         TabIndex        =   48
         Tag             =   "Importe|N|N|||linfact|importe|##,###,##0.00||"
         Text            =   "Importe"
         Top             =   1920
         Visible         =   0   'False
         Width           =   825
      End
      Begin VB.TextBox txtAux2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   285
         Index           =   0
         Left            =   3645
         TabIndex        =   66
         Top             =   1935
         Visible         =   0   'False
         Width           =   1455
      End
      Begin MSAdodcLib.Adodc AdoAux 
         Height          =   375
         Index           =   0
         Left            =   4560
         Top             =   240
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
         Caption         =   "AdoAux(0)"
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
      Begin MSComctlLib.Toolbar ToolAux 
         Height          =   390
         Index           =   0
         Left            =   240
         TabIndex        =   67
         Top             =   270
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   688
         ButtonWidth     =   609
         ButtonHeight    =   582
         AllowCustomize  =   0   'False
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   3
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Nuevo"
               Object.Tag             =   "2"
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Modificar"
               Object.Tag             =   "2"
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Eliminar"
               Object.Tag             =   "2"
            EndProperty
         EndProperty
         Begin VB.CheckBox Check2 
            Caption         =   "Vista previa"
            Height          =   195
            Index           =   1
            Left            =   8400
            TabIndex        =   68
            Top             =   120
            Visible         =   0   'False
            Width           =   1215
         End
      End
      Begin MSDataGridLib.DataGrid DataGridAux 
         Height          =   1905
         Index           =   0
         Left            =   240
         TabIndex        =   69
         Top             =   735
         Width           =   11790
         _ExtentX        =   20796
         _ExtentY        =   3360
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
   End
   Begin VB.Frame FrameTotFactu 
      Caption         =   "Total Factura"
      ForeColor       =   &H00972E0B&
      Height          =   1575
      Left            =   240
      TabIndex        =   59
      Top             =   2685
      Width           =   12195
      Begin VB.TextBox text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   22
         Left            =   6000
         MaxLength       =   6
         TabIndex        =   33
         Tag             =   "% REC 3|N|S|0|100.00|cabfact|porcrec3|##0.00|N|"
         Top             =   1170
         Width           =   645
      End
      Begin VB.TextBox text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   16
         Left            =   6000
         MaxLength       =   6
         TabIndex        =   27
         Tag             =   "% REC 2|N|S|0|100.00|cabfact|porcrec2|##0.00|N|"
         Top             =   855
         Width           =   645
      End
      Begin VB.TextBox text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   10
         Left            =   6000
         MaxLength       =   6
         TabIndex        =   21
         Tag             =   "% REC 1|N|S|0|100.00|cabfact|porcrec1|##0.00|N|"
         Text            =   "99.99"
         Top             =   495
         Width           =   645
      End
      Begin VB.TextBox text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   23
         Left            =   6810
         MaxLength       =   15
         TabIndex        =   34
         Tag             =   "Importe REC 3|N|S|||cabfact|imporec3|#,###,###,##0.00|N|"
         Top             =   1185
         Width           =   1635
      End
      Begin VB.TextBox text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   17
         Left            =   6810
         MaxLength       =   15
         TabIndex        =   28
         Tag             =   "Importe REC 2|N|S|||cabfact|imporec2|#,###,###,##0.00|N|"
         Top             =   840
         Width           =   1635
      End
      Begin VB.TextBox text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   11
         Left            =   6810
         MaxLength       =   15
         TabIndex        =   22
         Tag             =   "Importe Rec 1|N|S|||cabfact|imporec1|#,###,###,##0.00|N|"
         Top             =   495
         Width           =   1635
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
         Index           =   24
         Left            =   8790
         MaxLength       =   15
         TabIndex        =   35
         Tag             =   "Total Factura|N|S|||cabfact|totalfac|#,###,###,##0.00|N|"
         Top             =   495
         Width           =   2280
      End
      Begin VB.TextBox text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   7
         Left            =   2430
         MaxLength       =   2
         TabIndex        =   18
         Tag             =   "Tipo IVA 1|N|S|0|99|cabfact|tipoiva1|00||"
         Text            =   "12"
         Top             =   510
         Width           =   525
      End
      Begin VB.TextBox text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   13
         Left            =   2445
         MaxLength       =   2
         TabIndex        =   24
         Tag             =   "Tipo IVA 2|N|S|0|99|cabfact|tipoiva2|00||"
         Top             =   840
         Width           =   525
      End
      Begin VB.TextBox text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   19
         Left            =   2445
         MaxLength       =   2
         TabIndex        =   30
         Tag             =   "Tipo IVA 3|N|S|0|99|cabfact|tipoiva3|00||"
         Top             =   1185
         Width           =   525
      End
      Begin VB.TextBox text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   8
         Left            =   3180
         MaxLength       =   6
         TabIndex        =   19
         Tag             =   "% IVA 1|N|S|0|100.00|cabfact|porciva1|##0.00|N|"
         Text            =   "99.99"
         Top             =   510
         Width           =   645
      End
      Begin VB.TextBox text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   14
         Left            =   3180
         MaxLength       =   6
         TabIndex        =   25
         Tag             =   "% IVA 2|N|S|0|100.00|cabfact|porciva2|##0.00|N|"
         Top             =   840
         Width           =   645
      End
      Begin VB.TextBox text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   20
         Left            =   3180
         MaxLength       =   6
         TabIndex        =   31
         Tag             =   "% IVA 3|N|S|0|100.00|cabfact|porciva3|##0.00|N|"
         Top             =   1185
         Width           =   645
      End
      Begin VB.TextBox text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   9
         Left            =   4005
         MaxLength       =   15
         TabIndex        =   20
         Tag             =   "Importe IVA 1|N|S|||cabfact|impoiva1|#,###,###,##0.00|N|"
         Top             =   495
         Width           =   1605
      End
      Begin VB.TextBox text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   15
         Left            =   4005
         MaxLength       =   15
         TabIndex        =   26
         Tag             =   "Importe IVA 2|N|S|||cabfact|impoiva2|#,###,###,##0.00|N|"
         Top             =   840
         Width           =   1605
      End
      Begin VB.TextBox text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   21
         Left            =   4005
         MaxLength       =   15
         TabIndex        =   32
         Tag             =   "Importe IVA 3|N|S|||cabfact|impoiva3|#,###,###,##0.00|N|"
         Top             =   1185
         Width           =   1605
      End
      Begin VB.TextBox text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   6
         Left            =   225
         MaxLength       =   15
         TabIndex        =   17
         Tag             =   "Base IVA 1|N|S|||cabfact|baseiva1|#,###,###,##0.00|N|"
         Text            =   "575757575757557"
         Top             =   495
         Width           =   1605
      End
      Begin VB.TextBox text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   12
         Left            =   240
         MaxLength       =   15
         TabIndex        =   23
         Tag             =   "Base IVA 2|N|S|||cabfact|baseiva2|#,###,###,##0.00|N|"
         Top             =   840
         Width           =   1605
      End
      Begin VB.TextBox text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   18
         Left            =   240
         MaxLength       =   15
         TabIndex        =   29
         Tag             =   "Base IVA 3|N|S|||cabfact|baseiva3|#,###,###,##0.00|N|"
         Top             =   1185
         Width           =   1605
      End
      Begin VB.Label Label1 
         Caption         =   "% Rec."
         Height          =   255
         Index           =   8
         Left            =   6030
         TabIndex        =   77
         Top             =   270
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "Importe Recargo"
         Height          =   255
         Index           =   0
         Left            =   6810
         TabIndex        =   76
         Top             =   270
         Width           =   1545
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   4
         Left            =   2145
         MousePointer    =   4  'Icon
         Tag             =   "-1"
         ToolTipText     =   "Buscar tipo de IVA"
         Top             =   1200
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   3
         Left            =   2145
         MousePointer    =   4  'Icon
         Tag             =   "-1"
         ToolTipText     =   "Buscar tipo de IVA"
         Top             =   840
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   2
         Left            =   2145
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
         Left            =   8790
         TabIndex        =   64
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "Tipo IVA"
         Height          =   255
         Index           =   14
         Left            =   2445
         TabIndex        =   63
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "% IVA"
         Height          =   255
         Index           =   15
         Left            =   3210
         TabIndex        =   62
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "Importe IVA"
         Height          =   255
         Index           =   16
         Left            =   4005
         TabIndex        =   61
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Base Imponible"
         Height          =   255
         Index           =   10
         Left            =   240
         TabIndex        =   60
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Frame Frame1 
      Height          =   555
      Index           =   0
      Left            =   240
      TabIndex        =   51
      Top             =   7785
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
         TabIndex        =   52
         Top             =   180
         Width           =   2655
      End
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   11010
      TabIndex        =   40
      Top             =   7935
      Width           =   1035
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   9750
      TabIndex        =   39
      Top             =   7935
      Width           =   1035
   End
   Begin MSAdodcLib.Adodc Data1 
      Height          =   375
      Left            =   4200
      Top             =   7770
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
      Left            =   11010
      TabIndex        =   50
      Top             =   7935
      Visible         =   0   'False
      Width           =   1035
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   72
      Top             =   0
      Width           =   12555
      _ExtentX        =   22146
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      Style           =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   20
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
            Object.ToolTipText     =   "Modificar Total Factura"
            Object.Tag             =   "2"
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Carga Masiva"
            Object.Tag             =   "2"
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Imprimir"
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Salir"
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Style           =   3
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Primero"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Anterior"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button19 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Siguiente"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button20 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Último"
            ImageIndex      =   9
         EndProperty
      EndProperty
      BorderStyle     =   1
      Begin VB.CheckBox chkVistaPrevia 
         Caption         =   "Vista previa"
         Height          =   195
         Index           =   0
         Left            =   8520
         TabIndex        =   73
         Top             =   120
         Width           =   1215
      End
   End
   Begin VB.Label Label1 
      Caption         =   "Empresa"
      Height          =   255
      Index           =   6
      Left            =   360
      TabIndex        =   71
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
      Begin VB.Menu mn_ModTotales 
         Caption         =   "&Mod.Totales"
         Enabled         =   0   'False
         Shortcut        =   ^T
         Visible         =   0   'False
      End
      Begin VB.Menu mnCargaMasiva 
         Caption         =   "&Carga Masiva"
         HelpContextID   =   2
         Shortcut        =   ^C
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
Attribute VB_Name = "frmManFactVarias"
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
'   6.-  Modificar totales
'***Variables comuns a tots els formularis*****

Dim ModoLineas As Byte
'1.- Añadir,  2.- Modificar,  3.- Borrar,  0.-Pasar control a Lineas

Dim NumTabMto As Integer 'Indica quin nº de Tab està en modo Mantenimient
Dim TituloLinea As String 'Descripció de la llínia que està en Mantenimient
Dim PrimeraVez As Boolean
Dim ModificarTotales As Boolean

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

Private WithEvents frmSec As frmManSecciones
Attribute frmSec.VB_VarHelpID = -1
Private WithEvents frmCon As frmManConceptos
Attribute frmCon.VB_VarHelpID = -1
Private WithEvents frmCtas As frmCtasConta
Attribute frmCtas.VB_VarHelpID = -1
Private WithEvents frmTipIVA As frmTipIVAConta  'Tipos de IVA de la contabilidad
Attribute frmTipIVA.VB_VarHelpID = -1
Private WithEvents frmFPa As frmForpaConta  'Formas de Pago de la tesoreria
Attribute frmFPa.VB_VarHelpID = -1
Private WithEvents frmPais As frmBasico2 ' pais de la contabilidad
Attribute frmPais.VB_VarHelpID = -1

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

Dim Seguir As Boolean



Private Sub btnBuscar_Click(Index As Integer)
    ' els formularis als que crida son d'una atra BDA
    TerminaBloquear
    
    Select Case Index
        Case 0 'Conceptos
            Set frmCon = New frmManConceptos
            frmCon.DatosADevolverBusqueda = "0|1|2|4|"
            frmCon.CodigoActual = txtAux(5).Text
            frmCon.Show vbModal
            Set frmCon = Nothing
            
    End Select
    
    PonerFoco txtAux(5)
    If Modo = 4 Then BLOQUEADesdeFormulario2 Me, Data1, 1
End Sub

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
    Dim impiva(2) As Currency
    Dim PorIva(2) As Currency
    Dim ImpRec(2) As Currency
    Dim PorRec(2) As Currency
    Dim TotFac As Currency

'retencion
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
                Set vSec = New CSeccion
                If vSec.leer(text1(3).Text) Then
                    text1(1).Text = vSec.ConseguirContador(text1(3).Text)
                    text1(0).Text = vSec.LetraSerie
                    If InsertarDesdeForm2(Me, 1) Then
                        If vSec.IncrementarContador(text1(3)) Then
                            Data1.RecordSource = "Select * from " & NomTabla & Ordenacion
                            Cad = "codsecci = " & text1(3).Text & " and letraser = " & DBSet(Trim(text1(0).Text), "T")
                            Cad = Cad & " and numfactu = " & DBSet(text1(1).Text, "N")
                            Cad = Cad & " and fecfactu = " & DBSet(text1(2).Text, "F")
                            PosicionarData Cad
                            PonerModo 2
                            BotonAnyadirLinea 0
                        End If
                    End If
                End If
                Set vSec = Nothing
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
                conn.BeginTrans
                If BdConta <> 0 Then
                    If AbrirConexionContaFac(vParamAplic.UsuarioContaFac, vParamAplic.PasswordContaFac, BdConta) Then
                        ConnContaFac.BeginTrans
                        Set vEmpresaFac = New CempresaFac
                        If vEmpresaFac.LeerNiveles Then
                            PorRet = 0
                            If text1(26).Text <> "" Then PorRet = CCur(ImporteSinFormato(text1(26).Text))
                            If AdoAux(0).Recordset.RecordCount > 0 Then AdoAux(0).Recordset.MoveFirst
                            RecalculoBasesIvaFactura AdoAux(0).Recordset, Imptot, Tipiva, Impbas, impiva, PorIva, TotFac, ImpRec, PorRec, PorRet, ImpRet

                            text1(28).Text = ""
                            If ImpRet <> 0 Then text1(28).Text = Format(ImpRet, "#,###,###,##0.00")
                            text1(24).Text = Format(TotFac, "#,###,###,##0.00")

                            If text1(8).Text = "" Then text1(8).Text = "0,00"
                            If text1(9).Text = "" Then text1(9).Text = "0,00"
                        End If
                        Set vEmpresaFac = Nothing
                    End If
                End If
                
                If CadenaBorrado <> "" Then
                    conn.Execute CadenaBorrado
                    CadenaBorrado = ""
                    EliminarLinea
                End If
                
                
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
                    PosicionarData "codsecci = " & DBSet(text1(3).Text, "N") & " and letraser = '" & Trim(text1(0).Text) & "' and numfactu = " & text1(1).Text & " and fecfactu = " & DBSet(text1(2).Text, "F")
                End If
            End If
            
        Case 5 'LLINIES
            Select Case ModoLineas
                Case 1 'afegir llinia
                    InsertarLinea
                Case 2 'modificar llinies
                    ModificarLinea
                    PosicionarData "codsecci = " & text1(3).Text & " and letraser = '" & Trim(text1(0).Text) & "' and numfactu = " & text1(1).Text & " and fecfactu = " & DBSet(text1(2).Text, "F")
                    Screen.MousePointer = vbDefault
                    Exit Sub
            End Select
            
            
        Case 6  'MODIFICAR TOTALES
            If Not DatosOk Then
                ModoLineas = 0
                Screen.MousePointer = vbDefault
                Exit Sub
            Else
                ModoModificar = True
                conn.BeginTrans
                
                If ModificaDesdeFormulario2(Me, 1) Then
                    If Check1(1).Value = 1 Then
                        MsgBox "Los cambios realizados recuerde hacerlos en la Contabilidad y Cartera correspondiente.", vbExclamation
                        
                    End If
                    TerminaBloquear
                    PosicionarData "codsecci = " & DBSet(text1(3).Text, "N") & " and letraser = '" & Trim(text1(0).Text) & "' and numfactu = " & text1(1).Text & " and fecfactu = " & DBSet(text1(2).Text, "F")
                End If
            End If
            
            
    End Select
    Screen.MousePointer = vbDefault
 
Error1:
    Screen.MousePointer = vbDefault
    If Err.Number <> 0 Or Not b Then
        MsgBox Err.Number & ": " & Err.Description, vbExclamation
        If ModoModificar Then
            conn.RollbackTrans
            ConnContaFac.RollbackTrans
            ModoModificar = False
        End If
    Else
        If ModoModificar Then
            conn.CommitTrans
            ConnContaFac.CommitTrans
            ModoModificar = False
        End If
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
    btnPrimero = 17 'index del botó "primero"
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
        .Buttons(10).Image = 13 ' Modificar totales
        .Buttons(11).Image = 16 ' Carga masiva de facturas
        
        .Buttons(13).Image = 10  'Imprimir
        .Buttons(14).Image = 11  'Salir
        'el 14 i el 15 son separadors
        .Buttons(btnPrimero).Image = 6  'Primero
        .Buttons(btnPrimero + 1).Image = 7 'Anterior
        .Buttons(btnPrimero + 2).Image = 8 'Siguiente
        .Buttons(btnPrimero + 3).Image = 9 'Último
    End With
    
    'ICONITOS DE LAS BARRAS EN LOS TABS DE LINEA
    For i = 0 To ToolAux.Count - 1
        With Me.ToolAux(i)
            '.ImageList = frmPpal.imgListComun_VELL
            '  ### [Monica] 02/10/2006 acabo de comentarlo
            '.HotImageList = frmPpal.imgListComun_OM16
            '.DisabledImageList = frmPpal.imgListComun_BN16
            .ImageList = frmPpal.imgListComun16
            .Buttons(1).Image = 3   'Insertar
            .Buttons(2).Image = 4   'Modificar
            .Buttons(3).Image = 5   'Borrar
        End With
    Next i
    
    'cargar IMAGES de busqueda
    For i = 0 To Me.imgBuscar.Count - 1
        Me.imgBuscar(i).Picture = frmPpal.imgListImages16.ListImages(1).Picture
    Next i
   
    'IMAGES para zoom
    For i = 0 To Me.imgZoom.Count - 1
        Me.imgZoom(i).Picture = frmPpal.imgListImages16.ListImages(3).Picture
    Next i
    
    LimpiarCampos   'Limpia los campos TextBox
    For i = 0 To DataGridAux.Count - 1 'neteje tots els grids de llinies
        DataGridAux(i).ClearFields
    Next i
    
    '## A mano
    NomTabla = "cabfact"
    Ordenacion = " ORDER BY codsecci, letraser, numfactu, fecfactu "
    
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
    
    For i = 0 To DataGridAux.Count - 1
        CargaGrid i, (Modo = 2) 'carregue els datagrids de llinies
    Next i
    
    If LetraSerie <> "" Then
        text1(0).Text = Trim(LetraSerie)
        text1(1).Text = numfactu
        PonerModo 1
        cmdAceptar_Click
    End If


End Sub

Private Sub LimpiarCampos()
    On Error Resume Next

    Limpiar Me   'Metodo general: Limpia los controles TextBox
    lblIndicador.Caption = ""
    '[Monica]18/11/2013: cambios por aridoc
    Me.Check1(0).Value = 0
    Me.Check1(1).Value = 0
    
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
Dim CtaMultiple As Boolean

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
    
    BloquearImgBuscar Me, Modo, ModoLineas
       
    'Bloquear los campos de clave primaria, NO se puede modificar
    b = Not (Modo = 1) 'solo al insertar/buscar estará activo
    For i = 0 To 1
        BloquearTxt text1(i), b, True
        text1(i).Enabled = Not b
    Next i
    b = (Modo = 4) Or (Modo = 0) Or (Modo = 2) Or (Modo = 5)
    For i = 2 To 3
        BloquearTxt text1(i), b, True
        text1(i).Enabled = Not b
    Next i
    
    '[Monica]27/11/2017: el nombre de la cuenta es bloqueado
    b = (Modo = 0) Or (Modo = 2) Or (Modo = 5)
    BloquearTxt Text2(4), b, False
    Text2(4).Enabled = Not b
    
    For i = 6 To 24
        BloquearTxt text1(i), Not (Modo = 1 Or (Modo = 4 And ModificarTotales))
    Next i
    
    
    ' el importe de retencion solo se puede consultar
    BloquearTxt text1(28), Not (Modo = 1 Or (Modo = 4 And ModificarTotales))
    text1(28).Enabled = (Modo = 1 Or (Modo = 4 And ModificarTotales))
    
'    'Los % de IVA siempre bloqueados
'    BloquearTxt text1(8), True
'    BloquearTxt text1(14), True
'    BloquearTxt text1(20), True
'    'Los % de REC siempre bloqueados
'    BloquearTxt text1(10), True
'    BloquearTxt text1(16), True
'    BloquearTxt text1(22), True
    'El total de la factura siempre bloqueado
'    BloquearTxt text1(24), True
    
    '09/02/2007 no dejo modificar la forma de pago
   b = ((Modo = 4) And Me.Check1(1).Value = 1) Or (Modo = 0) Or (Modo = 2) Or (Modo = 5)
   BloquearTxt text1(25), b
    
    
    text1(24).BackColor = &HCAE3FD

    ' **** si n'hi han imagens de buscar en la capçalera *****
    BloquearImgBuscar Me, Modo, ModoLineas
    BloquearImgZoom Me, Modo, ModoLineas
    ' ********************************************************

    b = (Modo = 3) Or (Modo = 1)
    Me.imgBuscar(0).Enabled = b
    Me.imgBuscar(0).visible = b
    
    b = (Modo = 3) Or (Modo = 1) Or (Modo = 4 And Me.Check1(1).Value = 0)
    Me.imgBuscar(5).Enabled = b
    Me.imgBuscar(5).visible = b
    
    
    'Imagen Calendario fechas
    b = (Modo = 3 Or Modo = 4 Or Modo = 1 Or (Modo = 5 And (ModoLineas = 1 Or ModoLineas = 2)))
    Me.imgFec(2).Enabled = (Modo = 3 Or Modo = 1) 'es clave, solo al insertar o buscar
    Me.imgFec(2).visible = (Modo = 3 Or Modo = 1) 'es clave, solo al insertar o buscar
            
    chkVistaPrevia(0).Enabled = (Modo <= 2)
    
    'Poner el tamaño de los campos. Si es modo Busqueda el MaxLength del campo
    'debe ser mayor.
    PonerLongCampos
                          
    If (Modo < 2) Or (Modo = 3) Then
        For i = 0 To DataGridAux.Count - 1
            CargaGrid i, False
        Next i
    End If
    
    b = (Modo = 4) Or (Modo = 2)
    For i = 0 To DataGridAux.Count - 1
        DataGridAux(i).Enabled = b
    Next i
    
    ' solo podremos tocar el campo de contabilizado si estamos buscando
    Check1(1).Enabled = (Modo = 1)
    '[Monica]18/11/2013: cambios por aridoc
    Check1(0).Enabled = (Modo = 1) And vParamAplic.HayAridoc
    
    
    'b = (Modo = 4)
    b = (Modo = 1) Or (Modo = 4 And ModificarTotales)
    FrameTotFactu.Enabled = b
    
    Frame2(0).Enabled = (Modo = 4 And Not ModificarTotales) Or (Modo <> 4)
    
    b = (Modo = 5)
    Me.FrameAux0.Enabled = (Modo = 2) Or (Modo = 5)
    
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
    'modificar totals
    Toolbar1.Buttons(11).Enabled = b
    Me.mnCargaMasiva.Enabled = b
    
    
    
    b = (Modo = 2 And Data1.Recordset.RecordCount > 0) 'And (Check1(1).Value = 0)
    'Modificar
    Toolbar1.Buttons(8).Enabled = b
    Me.mnModificar.Enabled = b
    'eliminar
    Toolbar1.Buttons(9).Enabled = b
    Me.mnEliminar.Enabled = b
    'modificar totals
    Toolbar1.Buttons(10).Enabled = b
    Me.mnEliminar.Enabled = b
    
    'Imprimir
    'VRS:2.0.1(3)
    Toolbar1.Buttons(13).Enabled = (Modo = 2)
    Me.mnImprimir.Enabled = (Modo = 2)
    
    '-----------  LINEAS
    ' *** MEU: botons de les llínies de cuentas bancarias,
    ' només es poden gastar quan inserte o modifique clients ****
    'b = (Modo = 3 Or Modo = 4)
    b = (Modo = 3 Or (Modo = 4 And Not ModificarTotales) Or Modo = 2) 'And (Check1(1).Value = 0)
    For i = 0 To ToolAux.Count - 1
        ToolAux(i).Buttons(1).Enabled = b
        If b Then bAux = (b And Me.AdoAux(i).Recordset.RecordCount > 0)
        ToolAux(i).Buttons(2).Enabled = bAux
        ToolAux(i).Buttons(3).Enabled = bAux
    Next i
    'Imprimir en pestaña Comisiones de Productos
'    ToolAux(2).Buttons(6).Enabled = (Modo = 2) Or (Modo = 3) Or (Modo = 4) Or (Modo = 5 And ModoLineas = 0)
    ' ************************************************************
End Sub

Private Sub Desplazamiento(Index As Integer)
'Botones de Desplazamiento de la Toolbar
'Para desplazarse por los registros de control Data
    If Data1.Recordset.EOF Then Exit Sub
    DesplazamientoData Data1, Index
    PonerCampos
End Sub

Private Function MontaSQLCarga(Index As Integer, enlaza As Boolean) As String
'--------------------------------------------------------------------
' MontaSQlCarga:
'   Basándose en la información proporcionada por el vector de campos
'   crea un SQl para ejecutar una consulta sobre la base de datos que los
'   devuelva.
' Si ENLAZA -> Enlaza con el data1
'           -> Si no lo cargamos sin enlazar a ningun campo
'--------------------------------------------------------------------
Dim SQL As String
Dim tabla As String
    
    Select Case Index
        Case 0 'Lineas de factura
                tabla = "linfact"
                SQL = "SELECT codsecci,letraser,numfactu,fecfactu,numlinea,linfact.codconce,concefact.nomconce, linfact.tipoiva, ampliaci,"
                SQL = SQL & "cantidad, precio, importe"
                SQL = SQL & " FROM linfact, concefact "
                SQL = SQL & " WHERE linfact.codconce = concefact.codconce "
    
                If enlaza Then
                    SQL = SQL & " AND " & ObtenerWhereCab(False)
                Else
                    SQL = SQL & " AND codsecci = -1"
                End If
                SQL = SQL & " ORDER BY " & tabla & ".numlinea "
    End Select
    MontaSQLCarga = SQL
End Function

Private Sub frmB_Selecionado(CadenaDevuelta As String)
    Dim CadB As String
    Dim Aux As String
    
    If CadenaDevuelta <> "" Then
        HaDevueltoDatos = True
        Screen.MousePointer = vbHourglass
        'Sabemos que campos son los que nos devuelve
        'Creamos una cadena consulta y ponemos los datos
        CadB = ""
        Aux = ValorDevueltoFormGrid(text1(3), CadenaDevuelta, 1) 'codsecci
        CadB = Aux
        Aux = ValorDevueltoFormGrid(text1(0), CadenaDevuelta, 2) 'letraser
        CadB = CadB & " AND " & Aux
        Aux = ValorDevueltoFormGrid(text1(1), CadenaDevuelta, 3) 'numfactu
        CadB = CadB & " AND " & Aux
        Aux = ValorDevueltoFormGrid(text1(2), CadenaDevuelta, 4) 'fecfactu
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
    text1(CByte(imgFec(2).Tag)).Text = Format(vFecha, "dd/mm/yyyy")
End Sub

Private Sub frmCtas_DatoSeleccionado(CadenaSeleccion As String)
    text1(indice).Text = RecuperaValor(CadenaSeleccion, 1) 'codmacta
    Text2(indice).Text = RecuperaValor(CadenaSeleccion, 2) 'nommacta
End Sub

Private Sub frmFpa_DatoSeleccionado(CadenaSeleccion As String)
Dim Cad As String
    text1(25).Text = RecuperaValor(CadenaSeleccion, 1) 'codforpa
    Text2(0).Text = RecuperaValor(CadenaSeleccion, 2) 'nomsecci
End Sub

Private Sub frmCon_DatoSeleccionado(CadenaSeleccion As String)
'Conceptos
Dim BdConta As String
Dim Tipiva As String

    txtAux(5).Text = RecuperaValor(CadenaSeleccion, 1) 'codconce
    txtAux2(0).Text = RecuperaValor(CadenaSeleccion, 2) 'nomartic
    BdConta = RecuperaValor(CadenaSeleccion, 3) 'base de datos de conta
    Tipiva = RecuperaValor(CadenaSeleccion, 4) 'tipo de iva
    If BdConta <> "" Then
        If AbrirConexionContaFac(vParamAplic.UsuarioContaFac, vParamAplic.PasswordContaFac, CInt(BdConta)) Then
            If Tipiva <> "" Then
                txtAux(8).Text = Tipiva 'DevuelveDesdeBDNewFac("tiposiva", "nombriva", "codigiva", Tipiva, "N") 'descripcion del tipo de iva
            End If
            CerrarConexionContaFac
        End If
    End If
End Sub

Private Sub frmPais_DatoSeleccionado(CadenaSeleccion As String)
    If CadenaSeleccion <> "" Then
        text1(34).Text = RecuperaValor(CadenaSeleccion, 1)
        Text2(34).Text = RecuperaValor(CadenaSeleccion, 2)
    End If
End Sub

Private Sub frmSec_DatoSeleccionado(CadenaSeleccion As String)
Dim Cad As String
    text1(3).Text = RecuperaValor(CadenaSeleccion, 1) 'codsecci
    Text2(3).Text = RecuperaValor(CadenaSeleccion, 2) 'nomsecci
'    Text1(0).Text = RecuperaValor(CadenaSeleccion, 4) 'letraser
'    Text1(1).Text = RecuperaValor(CadenaSeleccion, 3) 'numfactu
    
    Cad = RecuperaValor(CadenaSeleccion, 5)  'numconta
    If Cad <> "" Then BdConta = CInt(Cad)  'numero de conta
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

        Case 0 'Seccion
            indice = 3
            Set frmSec = New frmManSecciones
            frmSec.DatosADevolverBusqueda = "0|1|2|3|4|"
            frmSec.CodigoActual = text1(3).Text
            frmSec.Show vbModal
            Set frmSec = Nothing
            
        Case 1, 6 'Cuenta Contable
            If BdConta = 0 Then Exit Sub
            
            If AbrirConexionContaFac(vParamAplic.UsuarioContaFac, vParamAplic.PasswordContaFac, BdConta) Then
                Set vEmpresaFac = New CempresaFac
                If vEmpresaFac.LeerNiveles Then
'                    txtAux(6) = PonerNombreCuenta(txtAux(5), Modo, cContaFac)
                    Set frmCtas = New frmCtasConta
                    Select Case Index
                        Case 1
                            indice = 4
                            frmCtas.CadBusqueda = DevuelveDesdeBDNew(cPTours, "seccion", "raizcta", "codsecci", text1(3).Text, "N")
                        Case 6
                            indice = 27
                            frmCtas.CadBusqueda = DevuelveDesdeBDNew(cPTours, "seccion", "raizctaret", "codsecci", text1(3).Text, "N")
                    End Select
                    frmCtas.Conexion = BdConta
                    frmCtas.Facturas = True
'[Monica] 09/09/2009 mod pq fallaba en la insercion de facturas varias cuando buscas por nombre
'                    frmCtas.NumDigit = DevuelveDesdeBDNewFac("empresa", "numdigi" & vEmpresaFac.numNivel, "", "", "")
                    frmCtas.NumDigit = 0
                    frmCtas.DatosADevolverBusqueda = "0|1|"
                    frmCtas.CodigoActual = text1(indice).Text
                    frmCtas.Show vbModal
                    Set frmCtas = Nothing
                    PonerFoco text1(indice)
                End If
                Set vEmpresaFac = Nothing
                CerrarConexionContaFac
            End If
            
        Case 5 'forma de pago
            If BdConta = 0 Then Exit Sub
            
            If AbrirConexionContaFac(vParamAplic.UsuarioContaFac, vParamAplic.PasswordContaFac, BdConta) Then
                Set vEmpresaFac = New CempresaFac
                If vEmpresaFac.LeerNiveles Then
'                    txtAux(6) = PonerNombreCuenta(txtAux(5), Modo, cContaFac)
                    indice = 5
                    Set frmFPa = New frmForpaConta
                    frmFPa.Conexion = BdConta
                    frmFPa.Facturas = True
                    frmFPa.DatosADevolverBusqueda = "0|1|"
                    frmFPa.CodigoActual = text1(25).Text
                    frmFPa.Show vbModal
                    Set frmFPa = Nothing
                    PonerFoco text1(25)
                End If
                Set vEmpresaFac = Nothing
                CerrarConexionContaFac
            End If
            
            
        Case 2, 3, 4 'tiposd de IVA (de la contabilidad)
            If BdConta = 0 Then Exit Sub
            
            If AbrirConexionContaFac(vParamAplic.UsuarioContaFac, vParamAplic.PasswordContaFac, BdConta) Then
                Set vEmpresaFac = New CempresaFac
                If vEmpresaFac.LeerNiveles Then
                    If Index = 2 Then Let indice = 7
                    If Index = 3 Then Let indice = 13
                    If Index = 4 Then Let indice = 19
                    Set frmTipIVA = New frmTipIVAConta
                    frmTipIVA.Facturas = True
                    frmTipIVA.Conexion = BdConta
                    frmTipIVA.DatosADevolverBusqueda = "0|1|2|3|"
                    frmTipIVA.CodigoActual = text1(indice).Text
                    frmTipIVA.Show vbModal
                    Set frmTipIVA = Nothing
                    PonerFoco text1(indice)

                End If
            End If
            
        Case 7 'codigo de pais
            If BdConta = 0 Then Exit Sub
            
            If AbrirConexionContaFac(vParamAplic.UsuarioContaFac, vParamAplic.PasswordContaFac, BdConta) Then
                Set vEmpresaFac = New CempresaFac
                If vEmpresaFac.LeerNiveles Then
                    Set frmPais = New frmBasico2
                    AyudaPais frmPais, text1(34).Text, , cContaFac
                    Set frmPais = Nothing
                    PonerFoco text1(34)
                End If
                Set vEmpresaFac = Nothing
                CerrarConexionContaFac
            End If
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
    imgFec(2).Tag = Index 'independentment de les dates que tinga, sempre pose l'index en la 27
    If text1(Index).Text <> "" Then frmC.NovaData = text1(Index).Text

    frmC.Show vbModal
    Set frmC = Nothing
    PonerFoco text1(CByte(imgFec(2).Tag))
    ' ***************************
End Sub

Private Sub imgZoom_Click(Index As Integer)
    
    Set frmZ = New frmZoom

    If Index = 0 Then
        indice = 5
        frmZ.pTitulo = "Observaciones de la Factura"
        frmZ.pValor = text1(indice).Text
        frmZ.pModo = Modo
    
        frmZ.Show vbModal
        Set frmZ = Nothing
            
        PonerFoco text1(indice)
    End If
End Sub



Private Sub mn_ModTotales_Click()

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
    If BLOQUEADesdeFormulario2(Me, Data1, 1) Then BotonModificarTotales
End Sub

Private Sub mnBuscar_Click()
    BotonBuscar
    Me.Check1(1).Value = 0
    '[Monica]18/11/2013: cambios por aridoc
    Me.Check1(0).Value = 0
End Sub

Private Sub mnCargaMasiva_Click()
    BotonCargaMasiva
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

    indRPT = 1 'Facturas Varias
    
    '[Monica]26/05/2016: si es materna cogemos otra impresion de facturas varias
    If EsSeccionMaterna(text1(3).Text) Then indRPT = 4
    

    If Not PonerParamRPT(indRPT, cadParam, numParam, nomDocu) Then Exit Sub
    'Nombre fichero .rpt a Imprimir
    frmImprimir.NombreRPT = nomDocu
    ' he añadido estas dos lineas para que llame al rpt correspondiente

    cadNombreRPT = nomDocu  ' "rFactgas.rpt"
    cadFormula = "({" & NomTabla & ".codsecci} = " & text1(3).Text & ") AND "
    cadFormula = cadFormula & "({" & NomTabla & ".letraser} = """ & Trim(text1(0).Text) & """) AND ({" & NomTabla & ".numfactu} = " & text1(1).Text & ") and ({" & NomTabla & ".fecfactu} = cdate(""" & text1(2).Text & """)) "
    
    '23022007 Monica: la separacion de la bonificacion solo la quieren en Alzira
'    If vParamAplic.Cooperativa = 1 Then cadFormula = cadFormula & " and {slhfac.numalbar} <> 'BONIFICA'" ' AND ({ssocio.impfactu}<=1)"
    
    cadParam = "|pEmpresa=" & vEmpresa.nomEmpre & "|" '& "|pCodigoISO="11112"|pCodigoRev="01"|
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

'Private Sub mnRectificar_Click()
'
'    'Comprobaciones
'    '--------------
'    If Data1.Recordset.EOF Then Exit Sub
'    If Data1.Recordset.RecordCount < 1 Then Exit Sub
'
'    'El registre de codi 0 no es pot Modificar ni Eliminar
'    ' ### [Monica] 27/09/2006
'    ' quitamos el control de no poder modificar ni eliminar si es 0
'    'If EsCodigoCero(CStr(Data1.Recordset.Fields(0).Value), FormatoCampo(text1(0))) Then Exit Sub
'
'    ' ### [Monica] 27/09/2006
'    ' solo podemos modificar en el caso de que haya contabilidad si la factura es modificable
'    If vParamAplic.NumeroConta <> 0 And Not FacturaModificable(Text1(0).Text, Text1(1).Text, Text1(2).Text, Check1(1).Value) Then Exit Sub
'
'    'Preparar para modificar
'    '-----------------------
'    If Not FacturaModificable(Text1(0).Text, Text1(1).Text, Text1(2).Text, Check1(1).Value) Then Exit Sub
'    If BLOQUEADesdeFormulario2(Me, Data1, 1) Then BotonRectificar
'End Sub

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
            '++monica:12/02/2008
            If CByte(Data1.Recordset!intconta) = 1 Then
               Cad = "   Se dispone a realizar cambios en los datos de la Factura.     " & vbCrLf & vbCrLf & _
                     "Recuerde modificar la Contabilidad y Tesoreria correspondiente!!!"
               MsgBox Cad, vbExclamation
            End If
            '++
            mnModificar_Click
        Case 9  'Borrar
            '++monica:12/02/2008
            If CByte(Data1.Recordset!intconta) = 1 Then
               Cad = "No se permite eliminar una Factura Contabilizada!!!"
               MsgBox Cad, vbExclamation
            Else
            '++
                mnEliminar_Click
            End If
        Case 10 'Rectificativa
            mn_ModTotales_Click
        Case 11 'Carga Masiva de Facturas
            mnCargaMasiva_Click
        Case 13 'Imprimir
            mnImprimir_Click
        Case 14    'Salir
            mnSalir_Click
        Case btnPrimero To btnPrimero + 3 'Flechas Desplazamiento
            Desplazamiento (Button.Index - btnPrimero)
    End Select
End Sub

Private Sub BotonBuscar()
    'Buscar
    Seguir = True
    
    '[Monica]27/11/2017: datos fiscales
    BloquearDatosFiscales False

    If Modo <> 1 Then
        BdConta = 0
        LimpiarCampos
        'Poner los grid sin apuntar a nada
        'LimpiarDataGrids
        PonerModo 1
        'Si pasamos el control aqui lo ponemos en amarillo
        PonerFoco text1(3)
        text1(3).BackColor = vbYellow
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

    CadB = ObtenerBusqueda2(Me, BuscaChekc, 1)
    
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
        Cad = Cad & "Sección|" & NomTabla & ".codsecci|N|" & FormatoCampo(text1(3)) & "|10·"
        Cad = Cad & "Nom. Sección|nomsecci|T||50·"
        Cad = Cad & "Serie|" & NomTabla & ".letraser|T|" & text1(0) & "|10·"
        Cad = Cad & "Nº Fact.|" & NomTabla & ".numfactu|N|" & FormatoCampo(text1(1)) & "|16·"
        Cad = Cad & ParaGrid(text1(2), 15, "Fecha")
        If Cad <> "" Then
            Screen.MousePointer = vbHourglass
            Set frmB = New frmBuscaGrid
            frmB.vCampos = Cad
            frmB.vTabla = NomTabla & " INNER JOIN seccion ON " & NomTabla & ".codsecci=seccion.codsecci "
            frmB.vSQL = CadB
            HaDevueltoDatos = False
            '###A mano
            frmB.vDevuelve = "0|2|3|4|"
            frmB.vTitulo = "Facturas Varias"
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
    
    For i = 0 To DataGridAux.Count - 1 'Limpias los DataGrid
        CargaGrid i, False
    Next i
    
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

    '[Monica]27/11/2017: datos fiscales
    BloquearDatosFiscales True

    Seguir = True
    'Añadiremos el boton de aceptar y demas objetos para insertar
    PonerModo 3

    'Quan afegixc pose en Fecha
    text1(2).Text = Format(Now, "dd/mm/yyyy")

    'Total Factura (por defecto=0)
'    text1(18).Text = "0"
'    text1(19).Text = "0"



    'em posicione en el 1r tab
    PonerFoco text1(3)
End Sub

Private Sub BotonModificar()
Dim vSec As CSeccion
    Seguir = True

    'Añadiremos el boton de aceptar y demas objetos para insertar
    ModificarTotales = False
    PonerModo 4
    
   
    ' cargamos la base de datos a la que apunta la seccion
    BdConta = 0
    Set vSec = New CSeccion
    If vSec.leer(text1(3).Text) Then
        BdConta = vSec.BdConta
    End If
    Set vSec = Nothing
    
    
    '[Monica]27/11/2017: datos fiscales
    If vParamAplic.ContabilidadNueva Then
        BloquearDatosFiscales Not EsCuentaMultiple(text1(4).Text)
    End If

    
    
    ' ### [Monica] 27/09/2006
    ' me guardo los valores anteriores de cuenta contable
    CtaAnt = text1(4).Text
    
    'Quan modifique pose en la F.Modificación la data actual
    PonerFoco text1(4)
End Sub


Private Sub BotonModificarTotales()
Dim vSec As CSeccion
    Seguir = True

    'Añadiremos el boton de aceptar y demas objetos para insertar
    ModificarTotales = True
    PonerModo 4
    
    ' cargamos la base de datos a la que apunta la seccion
    BdConta = 0
    Set vSec = New CSeccion
    If vSec.leer(text1(3).Text) Then
        BdConta = vSec.BdConta
    End If
    Set vSec = Nothing
    
    
    'Quan modifique pose en la F.Modificación la data actual
    PonerFoco text1(4)
End Sub




'Private Sub BotonRectificar()
'
'    Set frmList = New frmListado
'    'Añadiremos el boton de aceptar y demas objetos para insertar
'    frmList.CadTag = Text1(0).Text & "|" & Text1(1).Text & "|" & Text1(2).Text & "|" & Text1(3).Text & "|" & Text2(3).Text & "|" & Format(Check1(1).Value, "0") & "|"
'    frmList.OpcionListado = 12
'    frmList.Show vbModal
'
'End Sub

Private Sub BotonEliminar()
Dim Cad As String
Dim vSec As CSeccion
Dim NumFacElim As Long 'Numero de la Factura que se ha Eliminado
Dim NumSecElim As Integer 'Numero de la Seccion que se ha eliminado

    On Error GoTo EEliminar

    'Ciertas comprobaciones
    If Data1.Recordset.EOF Then Exit Sub
    
'    'El registre de codi 0 no es pot Modificar ni Eliminar
'    If EsCodigoCero(CStr(Data1.Recordset.Fields(1).Value), FormatoCampo(text1(1))) Then Exit Sub

    Cad = "¿Seguro que desea eliminar la factura?"
    Cad = Cad & vbCrLf & "Sección: " & Format(Data1.Recordset!codsecci, FormatoCampo(text1(3)))
    Cad = Cad & vbCrLf & "Serie: " & Format(Data1.Recordset!letraser, FormatoCampo(text1(0)))
    Cad = Cad & vbCrLf & "Nº: " & Format(Data1.Recordset!numfactu, FormatoCampo(text1(1)))
    Cad = Cad & vbCrLf & "Fecha: " & Data1.Recordset.Fields("fecfactu")
    
    'Borramos
    If MsgBox(Cad, vbQuestion + vbYesNo) = vbYes Then
        'Hay que eliminar
        On Error GoTo EEliminar
        Screen.MousePointer = vbHourglass
        NumRegElim = Data1.Recordset.AbsolutePosition
        NumSecElim = Data1.Recordset.Fields(0)
        NumFacElim = Data1.Recordset.Fields(2)
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
            'Devolvemos contador, si no estamos actualizando
            Set vSec = New CSeccion
            vSec.DevolverContador CStr(NumSecElim), NumFacElim
            Set vSec = Nothing
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
    
    For i = 0 To DataGridAux.Count - 1
        CargaGrid i, True
    Next i
    
    'Recuperar Descripciones de los campos de Codigo
    '--------------------------------------------------
    Text2(3).Text = PonerNombreDeCod(text1(3), "seccion", "nomsecci")
'    BdConta = DevuelveDesdeBDNew(cPTours, "seccion", "numconta", "codsecci", Text1(3).Text, "N")
    
    ' cargamos la base de datos a la que apunta la seccion
    Set vSec = New CSeccion
    If vSec.leer(text1(3)) Then
        If vSec.BdConta <> 0 Then
            BdConta = vSec.BdConta
            If AbrirConexionContaFac(vParamAplic.UsuarioContaFac, vParamAplic.PasswordContaFac, CByte(BdConta)) Then
                Set vEmpresaFac = New CempresaFac
                If vEmpresaFac.LeerNiveles Then
'[Monica]27/11/2017: quito esta linea pq el nopmbre es un campo de la tabla
'                    Text2(4).Text = PonerNombreCuenta(Text1(4), Modo, , CByte(BdConta), True)
                    If vParamAplic.ContabilidadNueva Then
                        Text2(0).Text = DevuelveDesdeBDNewFac("formapago", "nomforpa", "codforpa", text1(25).Text, "N")
                    Else
                        Text2(0).Text = DevuelveDesdeBDNewFac("sforpa", "nomforpa", "codforpa", text1(25).Text, "N")
                    End If
                    Text2(27).Text = ""
                    If text1(27).Text <> "" Then
                        Text2(27).Text = PonerNombreCuenta(text1(27), Modo, , CByte(BdConta), True)
                    End If
                    '[Monica]27/11/2017: nombre del pais
                    Text2(34).Text = ""
                    If vParamAplic.ContabilidadNueva Then
                        If text1(34).Text <> "" Then
                            Text2(34).Text = DevuelveDesdeBDNewFac("paises", "nompais", "codpais", text1(34).Text, "T")
                        End If
                    End If
                End If
                Set vEmpresaFac = Nothing
                CerrarConexionContaFac
            End If
        End If
    End If
    Set vSec = Nothing
    
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
        
        Case 5 'LINEAS
            Select Case ModoLineas
                Case 1 'afegir llinia
                    ModoLineas = 0
                    DataGridAux(NumTabMto).AllowAddNew = False
'                    SituarTab (NumTabMto)
                    'If BLOQUEADesdeFormulario2(Me, Data1, 1) Then BotonModificar  'Modificar
                    LLamaLineas NumTabMto, ModoLineas 'ocultar txtAux
                    'If DataGridAux(NumTabMto).Enabled Then DataGridAux(NumTabMto).SetFocus
                    DataGridAux(NumTabMto).Enabled = True
                    DataGridAux(NumTabMto).SetFocus

                    If Not AdoAux(NumTabMto).Recordset.EOF Then
                        AdoAux(NumTabMto).Recordset.MoveFirst
                    End If

                Case 2 'modificar llinies
                    ModoLineas = 0
                    PonerModo 4
                    If Not AdoAux(NumTabMto).Recordset.EOF Then
                        V = AdoAux(NumTabMto).Recordset.Fields(3) 'el 1 es el nº de llinia
                        AdoAux(NumTabMto).Recordset.Find (AdoAux(NumTabMto).Recordset.Fields(1).Name & " =" & V)
                    End If
                    LLamaLineas NumTabMto, ModoLineas 'ocultar txtAux
            End Select
            
            PosicionarData "codsecci = " & Data1.Recordset.Fields(0) & " and letraser = '" & Trim(text1(0).Text) & "' and numfactu = " & text1(1).Text & " and fecfactu = " & DBSet(text1(2).Text, "F")
            
'            If Not AdoAux(NumTabMto).Recordset.EOF Then
'                DataGridAux_RowColChange NumTabMto, 1, 1
'            Else
'                LimpiarCamposFrame NumTabMto
'            End If
    End Select
End Sub

Private Function DatosOk() As Boolean
Dim b As Boolean
Dim Datos As String
Dim SQL As String
Dim UltNiv As Integer

    On Error GoTo EDatosOK




    DatosOk = False
    b = CompForm2(Me, 1) 'Comprobar formato datos ok
    If Not b Then Exit Function
    
    'cuenta contable
    If b And text1(4).Text <> "" Then
        If BdConta = 0 Then
            MsgBox "No hay conexion a la contabilidad de la seccion. Revise", vbExclamation
            b = False
        Else
            If AbrirConexionContaFac(vParamAplic.UsuarioContaFac, vParamAplic.PasswordContaFac, CByte(BdConta)) Then
                Set vEmpresaFac = New CempresaFac
                If vEmpresaFac.LeerNiveles Then
'[Monica]27/11/2017: cambiado por lo de abajo
'                    Text2(4) = PonerNombreCuenta(Text1(4), Modo, , BdConta, True)
'                    If Text2(4) = "" Then
                    text1(4).Text = DevuelveDesdeBDNewFac("cuentas", "codmacta", "codmacta", text1(4).Text, "T")
                    If text1(4).Text = "" Then
                        MsgBox "No existe la cuenta contable en la contabilidad asociada a la sección", vbExclamation
                        b = False
                    Else
                        Select Case vEmpresaFac.numNivel - 1
                            Case 1
                                UltNiv = vEmpresaFac.numDigi1
                            Case 2
                                UltNiv = vEmpresaFac.numDigi2
                            Case 3
                                UltNiv = vEmpresaFac.numDigi3
                            Case 4
                                UltNiv = vEmpresaFac.numDigi4
                            Case 5
                                UltNiv = vEmpresaFac.numDigi5
                            Case 6
                                UltNiv = vEmpresaFac.numDigi6
                            Case 7
                                UltNiv = vEmpresaFac.numDigi7
                            Case 8
                                UltNiv = vEmpresaFac.numDigi8
                            Case 9
                                UltNiv = vEmpresaFac.numDigi9
                            Case 10
                                UltNiv = vEmpresaFac.numDigi10
                        End Select
                        If Mid(text1(4), 1, UltNiv) <> DevuelveDesdeBDNew(cPTours, "seccion", "raizcta", "codsecci", text1(3), "N") Then
                            If MsgBox("La Cuenta Contable no tiene la misma raiz que la sección." & vbCrLf & "          ¿ Desea Continuar ?", vbQuestion + vbYesNo + vbDefaultButton1) = vbNo Then
                                b = False
                            End If
                        End If
                    End If
                End If
                Set vEmpresaFac = Nothing
                CerrarConexionContaFac
            End If
        End If
    End If
    
    
    '[Monica]20/06/2017: control de fechas que antes no estaba
    If b And text1(2).Text <> "" Then
        If BdConta = 0 Then
            MsgBox "No hay conexion a la contabilidad de la seccion. Revise", vbExclamation
            b = False
        Else
            If AbrirConexionContaFac(vParamAplic.UsuarioContaFac, vParamAplic.PasswordContaFac, CByte(BdConta)) Then
                Set vEmpresaFac = New CempresaFac
                If vEmpresaFac.LeerNiveles Then
                
                    ResultadoFechaContaOK = EsFechaOKConta(CDate(text1(2).Text))
                    If ResultadoFechaContaOK > 0 Then
                        If ResultadoFechaContaOK <> 4 Then MsgBox MensajeFechaOkConta, vbExclamation
                        b = False
                    End If
                End If
                Set vEmpresaFac = Nothing
                CerrarConexionContaFac
            End If
        End If
    End If
    
    
    'si hay porcentaje de retencion debe de haber cuenta de retencion e
    If b And text1(26).Text <> "" And text1(27).Text = "" Then
        If CInt(text1(26).Text) <> 0 Then
            MsgBox "Si hay porcentaje de retención debe introducir una cuenta contable asociada. Revise.", vbExclamation
            b = False
        End If
    End If
    
    'cuenta contable de retencion
    If b And text1(27).Text <> "" Then
        If BdConta = 0 Then
            MsgBox "No hay conexion a la contabilidad de la seccion. Revise", vbExclamation
            b = False
        Else
            If AbrirConexionContaFac(vParamAplic.UsuarioContaFac, vParamAplic.PasswordContaFac, CByte(BdConta)) Then
                Set vEmpresaFac = New CempresaFac
                If vEmpresaFac.LeerNiveles Then
                    Text2(27) = PonerNombreCuenta(text1(27), Modo, , BdConta, True)
                    If Text2(27) = "" Then
                        MsgBox "No existe la cuenta contable de Retención en la contabilidad asociada a la sección", vbExclamation
                        b = False
                    Else
                        Select Case vEmpresaFac.numNivel - 1
                            Case 1
                                UltNiv = vEmpresaFac.numDigi1
                            Case 2
                                UltNiv = vEmpresaFac.numDigi2
                            Case 3
                                UltNiv = vEmpresaFac.numDigi3
                            Case 4
                                UltNiv = vEmpresaFac.numDigi4
                            Case 5
                                UltNiv = vEmpresaFac.numDigi5
                            Case 6
                                UltNiv = vEmpresaFac.numDigi6
                            Case 7
                                UltNiv = vEmpresaFac.numDigi7
                            Case 8
                                UltNiv = vEmpresaFac.numDigi8
                            Case 9
                                UltNiv = vEmpresaFac.numDigi9
                            Case 10
                                UltNiv = vEmpresaFac.numDigi10
                        End Select
                        If Mid(text1(27), 1, UltNiv) <> DevuelveDesdeBDNew(cPTours, "seccion", "raizctaret", "codsecci", text1(3), "N") Then
                            If MsgBox("La Cuenta Contable de Retención no tiene la misma raiz que la sección." & vbCrLf & "         ¿ Desea Continuar ?", vbQuestion + vbYesNo + vbDefaultButton1) = vbNo Then
                                b = False
                            End If
                        End If
                    End If
                End If
                Set vEmpresaFac = Nothing
                CerrarConexionContaFac
            End If
        End If
    End If
    
    '[Monica]30/11/2017: volvemos a comprobar el nif y si es incorrecto preguntamos si continuar
    If b Then
        If text1(29).Text <> "" And Not ModificaImportes Then
            If Not ValidarNIF(text1(29).Text) Then
                If MsgBox("¿ Desea continuar ?", vbQuestion + vbYesNo + vbDefaultButton1) = vbNo Then b = False
            End If
        End If
    End If
    
    
'--monica: lo he quitado pq ha de recalcular
'    'Comprobamos que la suma de importes de las lineas es igual al total de la factura
'    If b And Modo <> 3 Then
'        Datos = SumaLineas("")
'
'        If CCur(Datos) > CCur(TransformaPuntosComas(DBSet(text1(6).Text, "N"))) + CCur(TransformaPuntosComas(DBSet(text1(12).Text, "N"))) + CCur(TransformaPuntosComas(DBSet(text1(18).Text, "N"))) Then
'            MsgBox "La suma de los importes de lineas es mayor que el total de la factura!!!", vbExclamation
'            b = False
'        ElseIf CCur(Datos) < CCur(TransformaPuntosComas(DBSet(text1(6).Text, "N"))) + CCur(TransformaPuntosComas(DBSet(text1(12).Text, "N"))) + CCur(TransformaPuntosComas(DBSet(text1(18).Text, "N"))) Then
'            MsgBox "La suma de los importes de lineas es menor que el total de la factura!!!", vbExclamation
'            b = False
'        End If
'    End If
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
        
    conn.BeginTrans
    vWhere = ObtenerWhereCab(True)

    'Eliminar las Lineas de facturas de proveedor
    conn.Execute "DELETE FROM linfact " & vWhere
    
    'Eliminar la CABECERA
    conn.Execute "Delete from " & NomTabla & vWhere
               
FinEliminar:
    If Err.Number <> 0 Then
        MuestraError Err.Number, "Eliminar"
        conn.RollbackTrans
        eliminar = False
    Else
        conn.CommitTrans
        eliminar = True
    End If
End Function

Private Sub Text1_GotFocus(Index As Integer)
    kCampo = Index
    ConseguirFoco text1(Index), Modo
    If Index = 4 Then CtaAnt = text1(4).Text
End Sub


Private Sub Text1_LostFocus(Index As Integer)
Dim Cad As String, Datos As String
Dim Suma As Currency
Dim i As Integer
Dim CtaMultiple As Boolean


    If Not PerderFocoGnral(text1(Index), Modo) Then Exit Sub
    
    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub
    
    Select Case Index
        Case 1 'Nº factura
            If text1(Index).Text <> "" Then FormateaCampo text1(Index)
                        
        Case 2 'Fecha
            If text1(Index).Text <> "" Then PonerFormatoFecha text1(Index)
            
        Case 3 'Seccion
            If text1(Index).Text <> "" Then
                If PonerFormatoEntero(text1(3)) Then
                    Text2(Index).Text = PonerNombreDeCod(text1(Index), "seccion", "nomsecci", "codsecci", "N")
                    If Text2(Index).Text = "" Then
                        Cad = "No existe la Sección: " & text1(Index).Text & vbCrLf
                        Cad = Cad & "¿Desea crearla?" & vbCrLf
                        If MsgBox(Cad, vbQuestion + vbYesNo) = vbYes Then
                            Set frmSec = New frmManSecciones
                            frmSec.DatosADevolverBusqueda = "0|1|"
                            text1(Index).Text = ""
                            TerminaBloquear
                            frmSec.Show vbModal
                            Set frmSec = Nothing
                            If Modo = 4 Then BLOQUEADesdeFormulario2 Me, Data1, 1
                        Else
                            text1(Index).Text = ""
                        End If
                        PonerFoco text1(Index)
                    Else
                        'recuperar el numero de contabilidad
                        BdConta = DevuelveDesdeBDNew(cPTours, "seccion", "numconta", "codsecci", text1(3).Text, "N")
                        If DBLet(BdConta, "N") = 0 Then
                            MsgBox "Esta seccion no está asociada a ninguna contabilidad. Revise.", vbExclamation
                            text1(Index).Text = ""
                            PonerFoco text1(Index)
                        End If
                        
                    End If
                Else
                    Text2(Index).Text = ""
                    
                End If
            End If
            
            
        Case 4 'Cta Contable
            If text1(Index).Text = "" Then Exit Sub
            
            If BdConta = 0 Then Exit Sub
            
            If AbrirConexionContaFac(vParamAplic.UsuarioContaFac, vParamAplic.PasswordContaFac, CByte(BdConta)) Then
                Set vEmpresaFac = New CempresaFac
                If vEmpresaFac.LeerNiveles Then
                    If CtaAnt <> text1(4).Text Then Text2(4) = PonerNombreCuenta(text1(4), Modo, , BdConta, True)
                    If text1(Index).Text = "" Then
                        PonerFoco text1(Index)
                    Else
                        '[Monica]27/11/2017: si la cuenta es multiple, deben de introducir los datos fiscales
                        If vParamAplic.ContabilidadNueva Then
                            CtaMultiple = EsCuentaMultiple(text1(4).Text)
                            BloquearDatosFiscales Not CtaMultiple
                            If (CtaAnt <> text1(4).Text And Modo <> 1) Or Not CtaMultiple Then TraerDatosCuenta text1(4).Text  'Modo <> 4 And
                            If CtaMultiple Then
                                PonerFoco Text2(4)
                            End If
                        Else
                            TraerDatosCuenta text1(4).Text
                        End If
                    End If
                End If
                Set vEmpresaFac = Nothing
                CerrarConexionContaFac
            End If
        
        '[Monica]27/11/2017: para el caso de que sea una cuenta multiple
        Case 34 ' codigo de pais
            If text1(Index).Text = "" Then Exit Sub
            
            If BdConta = 0 Then Exit Sub
            
            If AbrirConexionContaFac(vParamAplic.UsuarioContaFac, vParamAplic.PasswordContaFac, CByte(BdConta)) Then
                Set vEmpresaFac = New CempresaFac
                If vEmpresaFac.LeerNiveles Then
                    If text1(Index).Text <> "" Then
                        If vParamAplic.ContabilidadNueva Then
                            Text2(Index).Text = DevuelveDesdeBDNewFac("paises", "nompais", "codpais", text1(Index).Text, "T")
                            If Text2(Index) = "" Then
                                MsgBox "No existe el País. Reintroduzca.", vbExclamation
                                PonerFoco text1(Index)
                            End If
                        End If
                    Else
                        Text2(Index).Text = ""
                    End If
                End If
            End If
        
        Case 25 'Forma pago
            If text1(Index).Text = "" Then Exit Sub
            
            If BdConta = 0 Then Exit Sub
            
            If AbrirConexionContaFac(vParamAplic.UsuarioContaFac, vParamAplic.PasswordContaFac, CByte(BdConta)) Then
                Set vEmpresaFac = New CempresaFac
                If vEmpresaFac.LeerNiveles Then
                    If vParamAplic.ContabilidadNueva Then
                        Text2(0).Text = DevuelveDesdeBDNewFac("formapago", "nomforpa", "codforpa", text1(25).Text, "N")
                    Else
                        Text2(0).Text = DevuelveDesdeBDNewFac("sforpa", "nomforpa", "codforpa", text1(25).Text, "N")
                    End If
                    If Text2(0).Text = "" Then
                        MsgBox "No existe la Forma de Pago. Reintroduzca.", vbExclamation
                        Seguir = False
                        PonerFoco text1(Index)
                    Else
                        Seguir = True
                    End If
                End If
                Set vEmpresaFac = Nothing
                CerrarConexionContaFac
            End If
            
        Case 26 'porcentaje de retencion
            If Modo = 1 Then Exit Sub
            PonerFormatoDecimal text1(Index), 7
            
        Case 8, 10, 14, 16, 20, 22, 24
            If Modo = 1 Then Exit Sub
            PonerFormatoDecimal text1(Index), 7
            
            
        Case 5 'despues de las observaciones si estamos insertando despues he de ir al campo de retencion
            If Modo = 3 And Seguir Then PonerFoco text1(26)
            
        Case 6, 9, 11, 12, 15, 17, 18, 21, 23    'IMPORTES Base, IVA
            PonerFormatoDecimal text1(Index), 1
            
        Case 7, 13, 19 'cod. IVA
           If text1(Index).Text = "" Then
              text1(Index + 1).Text = ""
           Else
                If BdConta = 0 Then Exit Sub
                If AbrirConexionContaFac(vParamAplic.UsuarioContaFac, vParamAplic.PasswordContaFac, CByte(BdConta)) Then
                    Set vEmpresaFac = New CempresaFac
                    If vEmpresaFac.LeerNiveles Then
                        text1(Index + 1).Text = DevuelveDesdeBDNewFac("tiposiva", "porceiva", "codigiva", text1(Index).Text, "N")
                    End If
                    Set vEmpresaFac = Nothing
                    CerrarConexionContaFac
                End If
           End If
              
        Case 27 'cuenta de retencion
            Text2(Index).Text = ""
            If text1(Index).Text = "" Then Exit Sub
            
            If BdConta = 0 Then Exit Sub
            
            If AbrirConexionContaFac(vParamAplic.UsuarioContaFac, vParamAplic.PasswordContaFac, CByte(BdConta)) Then
                Set vEmpresaFac = New CempresaFac
                If vEmpresaFac.LeerNiveles Then
                    Text2(27) = PonerNombreCuenta(text1(27), Modo, , BdConta, True)
                    If Text2(Index).Text = "" Then
                        PonerFoco text1(Index)
                    End If
                End If
                Set vEmpresaFac = Nothing
                CerrarConexionContaFac
            End If
              
        Case 29 ' nif, se valida
            If text1(29).Text = "" Or Modo = 1 Then Exit Sub
            
            text1(Index).Text = UCase(text1(Index).Text)
            ValidarNIF text1(Index).Text
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
Dim cerrar As Boolean

    KEYpressGnral KeyAscii, Modo, cerrar
    If cerrar Then Unload Me
End Sub

Private Sub KEYBusqueda(KeyAscii As Integer, indice As Integer)
    KeyAscii = 0
    imgBuscar_Click (indice)
End Sub

Private Sub KEYFecha(KeyAscii As Integer, indice As Integer)
    KeyAscii = 0
    imgFec_Click (indice)
End Sub

Private Sub KEYBusquedaLin(KeyAscii As Integer, indice As Integer)
    KeyAscii = 0
    btnBuscar_Click (indice)
End Sub

Private Sub Text2_GotFocus(Index As Integer)
    kCampo = Index
    ConseguirFoco Text2(Index), Modo
End Sub

Private Sub Text2_LostFocus(Index As Integer)
Dim Cad As String, Datos As String
Dim Suma As Currency
Dim i As Integer
Dim CtaMultiple As Boolean


    If Not PerderFocoGnral(Text2(Index), Modo) Then Exit Sub
    
    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub
End Sub

Private Sub Text2_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
'Avanzar/Retroceder los campos con las flechas de desplazamiento del teclado.
    KEYdown KeyCode
End Sub

Private Sub Text2_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub






'************* LLINIES: ****************************
Private Sub ToolAux_ButtonClick(Index As Integer, ByVal Button As MSComctlLib.Button)
Dim Cad As String
'12/02/2008: lo he quitado porque lo modificaran ellos manualmente en la contabilidad
'    If vParamAplic.NumeroConta <> 0 And _
'       Not FacturaModificable(text1(0).Text, text1(1).Text, text1(2).Text, Check1(1).Value) Then Exit Sub
    '++monica:12/02/2008
     If CByte(Data1.Recordset!intconta) = 1 Then
        Cad = "   Se dispone a realizar cambios en los datos de la Factura.     " & vbCrLf & vbCrLf & _
              "Recuerde modificar la Contabilidad y Tesoreria correspondiente!!!"
        MsgBox Cad, vbExclamation
     End If
    '++
    
    
     Select Case Button.Index
        Case 1
'            TerminaBloquear
            BotonAnyadirLinea Index
        Case 2
'            TerminaBloquear
            BotonModificarLinea Index
        Case 3
'            TerminaBloquear
            BotonEliminarLinea Index
            If Modo = 4 Then
                If BLOQUEADesdeFormulario2(Me, Data1, 1) Then BotonModificar
            End If
        Case 6 'Imprimir
'            BotonImprimirLinea Index
    End Select
End Sub

Private Sub BotonEliminarLinea(Index As Integer)
Dim SQL As String
Dim eliminar As Boolean

    On Error GoTo Error2

    ModoLineas = 3 'Posem Modo Eliminar Llínia

    If Modo = 4 Then 'Modificar Capçalera
        cmdAceptar_Click
        If ModoLineas = 0 Then Exit Sub
    End If

    NumTabMto = Index
    PonerModo 5

'    If AdoAux(Index).Recordset.EOF Then Exit Sub
    If Not SepuedeBorrar(Index) Then Exit Sub
    
    If AdoAux(Index).Recordset.RecordCount = 1 Then
        MsgBox "No se puede borrar un única línea de factura, elimine la factura completa", vbExclamation
        PonerModo 2
        Exit Sub
    End If
    
    
    eliminar = False

    Select Case Index
        Case 0 'lineas de factura
            SQL = "¿Seguro que desea eliminar la línea?"
            SQL = SQL & vbCrLf & "Nº línea: " & Format(DBLet(AdoAux(Index).Recordset!NumLinea), FormatoCampo(txtAux(4)))
            SQL = SQL & vbCrLf & "Concepto: " & DBLet(AdoAux(Index).Recordset!codConce) '& "  " & txtAux(4).Text
            If MsgBox(SQL, vbQuestion + vbYesNo) = vbYes Then
                NumRegElim = AdoAux(Index).Recordset.AbsolutePosition
                eliminar = True
                SQL = "DELETE FROM linfact"
                SQL = SQL & ObtenerWhereCab(True) & " AND numlinea= " & AdoAux(Index).Recordset!NumLinea
            End If
    End Select

    If eliminar Then
        TerminaBloquear
'        conn.Execute Sql
        CadenaBorrado = SQL
        '16022007
        If BLOQUEADesdeFormulario2(Me, Data1, 1) Then
                ModificaImportes = True
                BotonModificar
                cmdAceptar_Click
                ModificaImportes = False
        End If
'        EliminarLinea
        
        
        'antes estaba debajo de situardata
        CargaGrid Index, True
        SituarDataTrasEliminar AdoAux(Index), NumRegElim, True
        
        
        
    End If

    ModoLineas = 0
    PosicionarData "codsecci = " & text1(3).Text & " and letraser = '" & Trim(text1(0).Text) & "' and numfactu = " & text1(1).Text & " and fecfactu = " & DBSet(text1(2).Text, "F")

    Exit Sub
Error2:
    Screen.MousePointer = vbDefault
    MuestraError Err.Number, "Eliminando linea", Err.Description
End Sub

Private Sub BotonAnyadirLinea(Index As Integer)
Dim NumF As String
Dim vWhere As String, vTabla As String
Dim anc As Single
Dim i As Integer
Dim SumLin As Currency
Dim vSec As CSeccion

    'Si no estaba modificando lineas salimos
    ' Es decir, si estaba insertando linea no podemos hacer otra cosa
    'If ModificaLineas = 2 Then Exit Sub
    ModoLineas = 1 'Ponemos Modo Añadir Linea

    If (Modo = 3) Or (Modo = 4) Then 'Insertar o Modifcar Cabecera
        cmdAceptar_Click
        'No se ha insertado la cabecera
        If ModoLineas = 0 Then Exit Sub
    End If

    NumTabMto = Index
    PonerModo 5
'    If b Then BloquearText1 Me, 4 'Si viene de Insertar Cabecera no bloquear los Text1


    'Obtener el numero de linea ha insertar
    Select Case Index
        Case 0: vTabla = "linfact"
    End Select
    'Obtener el sig. nº de linea a insertar
    vWhere = ObtenerWhereCab(False)
    NumF = SugerirCodigoSiguienteStr(vTabla, "numlinea", vWhere)

    'Situamos el grid al final
    AnyadirLinea DataGridAux(Index), AdoAux(Index)

    anc = DataGridAux(Index).Top
    If DataGridAux(Index).Row < 0 Then
        anc = anc + 210
    Else
        anc = anc + DataGridAux(Index).RowTop(DataGridAux(Index).Row) + 5
    End If

    LLamaLineas Index, ModoLineas, anc

    Select Case Index
        Case 0 'lineas factura
            txtAux(0).Text = text1(3).Text 'seccion
            txtAux(1).Text = text1(0).Text 'serie
            txtAux(2).Text = text1(1).Text 'factura
            txtAux(3).Text = text1(2).Text 'fecha
            txtAux(4).Text = NumF 'numlinea
'            FormateaCampo txtAux(3)
            For i = 5 To txtAux.Count - 1
                txtAux(i).Text = ""
            Next i
            txtAux2(0).Text = ""

            'desbloquear la linea (se bloquea al añadir)
'            BloquearTxt txtAux(3), False
            PonerFoco txtAux(5)
    End Select
End Sub

Private Sub BotonModificarLinea(Index As Integer)
    Dim anc As Single
    Dim i As Integer
    Dim J As Integer
    Dim vSec As CSeccion
    
    If AdoAux(Index).Recordset.EOF Then Exit Sub
    If AdoAux(Index).Recordset.RecordCount < 1 Then Exit Sub
    
    ModoLineas = 2 'Modificar llínia
    
    If Modo = 4 Then 'Modificar Cabecera
        cmdAceptar_Click
        If ModoLineas = 0 Then Exit Sub
    End If
    
    ' cargamos la base de datos a la que apunta la seccion
    BdConta = 0
    Set vSec = New CSeccion
    If vSec.leer(text1(3).Text) Then
        BdConta = vSec.BdConta
    End If
    Set vSec = Nothing
    
    NumTabMto = Index
    PonerModo 5
    
    If DataGridAux(Index).Bookmark < DataGridAux(Index).FirstRow Or DataGridAux(Index).Bookmark > (DataGridAux(Index).FirstRow + DataGridAux(Index).VisibleRows - 1) Then
        i = DataGridAux(Index).Bookmark - DataGridAux(Index).FirstRow
        DataGridAux(Index).Scroll 0, i
        DataGridAux(Index).Refresh
    End If
      
    anc = DataGridAux(Index).Top
    If DataGridAux(Index).Row < 0 Then
        anc = anc + 210
    Else
        anc = anc + DataGridAux(Index).RowTop(DataGridAux(Index).Row) + 5
    End If

    Select Case Index
        Case 0 'lineas de factura
            For J = 0 To 5
                txtAux(J).Text = DataGridAux(Index).Columns(J).Text
            Next J
            txtAux2(0).Text = DataGridAux(Index).Columns(6).Text 'DevuelveDesdeBDNew(cPTours, "concefact", "nomconce", "codconce", DataGridAux(Index).Columns(5).Text, "N")
            txtAux(8).Text = DataGridAux(Index).Columns(7).Text 'DevuelveDesdeBDNew(cPTours, "concefact", "tipoiva", "codconce", DataGridAux(Index).Columns(5).Text, "N")
            txtAux(6).Text = DataGridAux(Index).Columns(8).Text    ' ampliacion
            txtAux(7).Text = DataGridAux(Index).Columns(11).Text   ' importe
            txtAux(9).Text = DataGridAux(Index).Columns(9).Text    ' cantidad
            txtAux(10).Text = DataGridAux(Index).Columns(10).Text  ' precio
    End Select
    
    LLamaLineas Index, ModoLineas, anc
   
    Select Case Index
        Case 0 'lineas de factura
            PonerFoco txtAux(5)
    End Select
End Sub

Private Sub LLamaLineas(Index As Integer, xModo As Byte, Optional alto As Single)
Dim jj As Integer
Dim b As Boolean

    On Error GoTo ELLamaLin

    DeseleccionaGrid DataGridAux(Index)
    
    b = (xModo = 1 Or xModo = 2) 'Insertar o Modificar Lineas
    Select Case Index
        Case 0 'lineas de factura
            For jj = 5 To txtAux.Count - 1
                txtAux(jj).Top = alto
                txtAux(jj).visible = b
            Next jj
            txtAux(8).visible = False
            txtAux(8).Enabled = False
            
            txtAux2(0).Top = alto
            txtAux2(0).visible = b
            Me.btnBuscar(0).Top = alto
            Me.btnBuscar(0).visible = b
    End Select
    
ELLamaLin:
    Err.Clear
End Sub


Private Sub txtAux_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = teclaBuscar Then
        If Modo = 5 And (ModoLineas = 1 Or ModoLineas = 2) Then
            Select Case Index
                Case 5: KEYBusquedaLin KeyAscii, 0
                Case 6: KEYBusquedaLin KeyAscii, 1
            End Select
        End If
    Else
        KEYpress KeyAscii
    End If
End Sub

Private Sub txtAux_LostFocus(Index As Integer)
Dim cadMen As String
Dim SQL As String
    txtAux(Index).Text = Trim(txtAux(Index).Text)

    Select Case Index
        Case 6 ' Ampliacion
            txtAux(Index).Text = UCase(txtAux(Index).Text)
            


        Case 5 ' Concepto
            If PonerFormatoEntero(txtAux(Index)) Then
                txtAux2(0).Text = PonerNombreDeCod(txtAux(Index), "concefact", "nomconce", "codconce", "N")
                txtAux(8).Text = PonerNombreDeCod(txtAux(Index), "concefact", "tipoiva", "codconce", "N")
                If txtAux2(0).Text = "" Then
                    cadMen = "No existe el Concepto: " & txtAux(Index).Text & vbCrLf
                    cadMen = cadMen & "¿Desea crearlo?" & vbCrLf
                    If MsgBox(cadMen, vbQuestion + vbYesNo) = vbYes Then
                        Set frmCon = New frmManConceptos
                        frmCon.DatosADevolverBusqueda = "0|1|"
                        frmCon.NuevoCodigo = txtAux(Index).Text
                        txtAux(Index).Text = ""
                        TerminaBloquear
                        frmCon.Show vbModal
                        Set frmCon = Nothing
                    Else
                        txtAux(Index).Text = ""
                    End If
                    PonerFoco txtAux(Index)
                Else
                    BdConta1 = PonerNombreDeCod(txtAux(Index), "concefact", "numconta", "codconce", "N")
                    
                    If BdConta1 <> BdConta Then
                        MsgBox "La Conta de este concepto ha de ser la misma que la de la sección. Reintroduzca.", vbExclamation
                        txtAux(Index).Text = ""
                        PonerFoco txtAux(5)
                    End If
                End If
            Else
                txtAux2(0).Text = ""
            End If

        Case 9 ' cantidad
            If txtAux(Index).Text <> "" Then
                PonerFormatoDecimal txtAux(Index), 3
'                If txtAux(10).Text <> "" Then
                    txtAux(7).Text = Round2(CCur(txtAux(9).Text) * CCur(ComprobarCero(txtAux(10).Text)), 2)
                    PonerFormatoDecimal txtAux(7), 3
'                End If
            End If
            
        Case 10 ' precio
            If txtAux(Index).Text <> "" Then
                If PonerFormatoDecimal(txtAux(Index), 8) Then
'                If txtAux(9).Text <> "" Then
                    txtAux(7).Text = Round2(CCur(ComprobarCero(txtAux(9).Text)) * CCur(txtAux(10).Text), 2)
                    PonerFormatoDecimal txtAux(7), 3
'                End If
                End If
            End If
        
        Case 7 'Importe
'           If Trim(txtAux(Index).Text) = "" Then
'                PonerFocoBtn Me.cmdAceptar
'                Exit Sub
'           End If
           If Not EsNumerico(txtAux(Index).Text) Then
                MsgBox "El Importe debe ser numérico.", vbExclamation
                On Error Resume Next
                txtAux(Index).Text = ""
                PonerFoco txtAux(Index)
                Exit Sub
            End If
            'Es numerico
            PonerFormatoDecimal txtAux(Index), 3
            PonerFocoBtn Me.cmdAceptar
    End Select
    
End Sub

Private Sub txtAux_GotFocus(Index As Integer)
    ' si vamos a insertar el importe miramos si podemos calcularlo y no entrar en importe
    If Index = 7 And (txtAux(9).Text <> "" Or txtAux(10).Text <> "") And txtAux(Index).Text = "" Then
        txtAux(Index).Text = Round2(ComprobarCero(txtAux(9).Text) * ComprobarCero(txtAux(10).Text), 2)
'        cmdAceptar.SetFocus
        Exit Sub
    End If
    
    ConseguirFocoLin txtAux(Index)
End Sub

Private Function DatosOkLlin(nomFrame As String) As Boolean
Dim b As Boolean
Dim SumLin As Currency
    
    On Error GoTo EDatosOKLlin

    DatosOkLlin = False
        
    b = CompForm2(Me, 2, nomFrame) 'Comprobar formato datos ok
    If Not b Then Exit Function
    
' ### [Monica] 29/09/2006
' he quitado la parte de comprobar la suma de lineas
'    'Comprobar que el Importe del total de las lineas suma el total o menos de la factura
'    SumLin = CCur(SumaLineas(txtAux(4).Text))
'
'    'Le añadimos el importe de linea que vamos a insertar
'    SumLin = SumLin + CCur(txtAux(7).Text)
'
'    'comprobamos que no sobrepase el total de la factura
'    If SumLin > CCur(Text1(18).Text) Then
'        MsgBox "La suma del importe de las lineas no puede ser superior al total de la factura.", vbExclamation
'        b = False
'    End If
    
    DatosOkLlin = b
    
EDatosOKLlin:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Function

Private Function SepuedeBorrar(ByRef Index As Integer) As Boolean

    SepuedeBorrar = False
    If AdoAux(Index).Recordset.EOF Then Exit Function

    SepuedeBorrar = True
End Function

Private Sub CargaGrid(Index As Integer, enlaza As Boolean)
Dim b As Boolean
Dim tots As String

    On Error GoTo ECarga

    'b = DataGridAux(Index).Enabled
    'DataGridAux(Index).Enabled = False
    
    tots = MontaSQLCarga(Index, enlaza)
    
    CargaGridGnral Me.DataGridAux(Index), Me.AdoAux(Index), tots, PrimeraVez
    
    Select Case Index
        Case 0 'lineas de factura
            'si es visible|control|tipo campo|nombre campo|ancho control|formato campo|
            tots = "N||||0|;N||||0|;N||||0|;N||||0|;N||||0|;"
            tots = tots & "S|txtAux(5)|T|Código|700|;S|btnBuscar(0)|B|||;S|txtAux2(0)|T|Concepto|2100|;S|txtAux(8)|T|T.Iva|550|;"
            tots = tots & "S|txtAux(6)|T|Ampliación|4300|;S|txtAux(9)|T|Cantidad|1150|;S|txtAux(10)|T|Precio|1000|;S|txtAux(7)|T|Importe|1350|;"
            arregla tots, DataGridAux(Index), Me
'           DataGridAux(Index).Columns(6).Alignment = dbgCenter
'           DataGridAux(Index).Columns(9).Alignment = dbgRight
    End Select
    
    DataGridAux(Index).ScrollBars = dbgAutomatic
    
ECarga:
    If Err.Number <> 0 Then MuestraError Err.Number, "Cargando datos grid: " & DataGridAux(Index).Tag, Err.Description
End Sub

Private Sub InsertarLinea()
'Inserta registro en las tablas de Lineas: provbanc, provdpto
Dim nomFrame As String
Dim b As Boolean
Dim V As Integer

' variables para el recalculo de iva y totales
    Dim i As Integer
    Dim Imptot(2)
    Dim Tipiva(2)
    Dim Impbas(2) As Currency
    Dim impiva(2) As Currency
    Dim PorIva(2) As Currency
    Dim ImpRec(2) As Currency
    Dim PorRec(2) As Currency
    Dim TotFac As Currency

'retencion
    Dim PorRet As Currency
    Dim ImpRet As Currency

    On Error Resume Next

    Select Case NumTabMto
        Case 0: nomFrame = "FrameAux0" 'lineas de factura
    End Select

    If DatosOkLlin(nomFrame) Then
        TerminaBloquear
        If InsertarDesdeForm2(Me, 2, nomFrame) Then
            b = BLOQUEADesdeFormulario2(Me, Data1, 1)
            CargaGrid NumTabMto, True
            V = AdoAux(NumTabMto).Recordset.Fields(4) 'el 2 es el nº de llinia
            ModoLineas = 0
            CargaGrid NumTabMto, True
'            SituarTab (NumTabMto)
            DataGridAux(NumTabMto).SetFocus
            AdoAux(NumTabMto).Recordset.Find (AdoAux(NumTabMto).Recordset.Fields(4).Name & " =" & V)
            
'            ' ### [Monica] 29/09/2006
'            ' añadido el tema de de recalculo de bases
            If BdConta = 0 Then Exit Sub
            
            If AbrirConexionContaFac(vParamAplic.UsuarioContaFac, vParamAplic.PasswordContaFac, CByte(BdConta)) Then
                Set vEmpresaFac = New CempresaFac
                If vEmpresaFac.LeerNiveles Then
                    PorRet = 0
                    If text1(26).Text <> "" Then PorRet = CCur(ImporteSinFormato(text1(26).Text))
                    RecalculoBasesIvaFactura AdoAux(0).Recordset, Imptot, Tipiva, Impbas, impiva, PorIva, TotFac, ImpRec, PorRec, PorRet, ImpRet
                End If
                Set vEmpresaFac = Nothing
                CerrarConexionContaFac
            End If



            '13/02/2007 iniacializo los txt
            For i = 0 To 2
                text1(6 + (6 * i)).Text = ""
                text1(7 + (6 * i)).Text = ""
                text1(8 + (6 * i)).Text = ""
                text1(9 + (6 * i)).Text = ""
                text1(10 + (6 * i)).Text = ""
                text1(11 + (6 * i)).Text = ""
            Next i
            text1(26).Text = ""
            text1(28).Text = ""
            
            '13/02/2007 he añadido las condiciones del for antes solo estaban las sentencias
            For i = 0 To 2
                 If Tipiva(i) <> 0 Then
                    text1(6 + (6 * i)).Text = Impbas(i)
                    text1(7 + (6 * i)).Text = Tipiva(i)
                    text1(8 + (6 * i)).Text = PorIva(i)
                    text1(9 + (6 * i)).Text = impiva(i)
                    If PorRec(i) <> 0 Then text1(10 + (6 * i)).Text = PorRec(i)
                    If ImpRec(i) <> 0 Then text1(11 + (6 * i)).Text = ImpRec(i)
                 End If
'12/03/2007
'                 If Impbas(i) <> 0 Then text1(6 + (6 * i)).Text = Impbas(i)
'                 If PorIva(i) <> 0 Then text1(8 + (6 * i)).Text = PorIva(i)
'                 If impiva(i) <> 0 Then text1(9 + (6 * i)).Text = impiva(i)
'                 If PorRec(i) <> 0 Then text1(10 + (6 * i)).Text = PorRec(i)
'                 If ImpRec(i) <> 0 Then text1(11 + (6 * i)).Text = ImpRec(i)

                 'TotFac = Impbas(i) + impiva(i)
            Next i
            If PorRet <> 0 Then text1(26).Text = PorRet
            If ImpRet <> 0 Then text1(28).Text = ImpRet
            text1(24).Text = TotFac

            If text1(8).Text = "" Then text1(8).Text = "0,00"
            If text1(9).Text = "" Then text1(9).Text = "0,00"
            
            
'++monica: 10/03/2009
            PonerFormatos
'++
            
            If BLOQUEADesdeFormulario2(Me, Data1, 1) Then
'                Modo = 4
'                PonerModo Modo
'                ClienteAnt = Text1(3).Text
'                FormaPagoAnt = Text1(5).Text
                ModificaImportes = True
                BotonModificar
                cmdAceptar_Click
                ModificaImportes = False
            End If

            LLamaLineas NumTabMto, 0
            
            If b Then BotonAnyadirLinea NumTabMto
        End If
    End If
End Sub

Private Sub ModificarLinea()
'Modifica registro en las tablas de Lineas: provbanc, provdpto
Dim nomFrame As String
Dim V As Currency

' variables para el recalculo de iva y totales
    Dim i As Integer
    Dim Imptot(2)
    Dim Tipiva(2)
    Dim Impbas(2) As Currency
    Dim impiva(2) As Currency
    Dim PorIva(2) As Currency
    Dim ImpRec(2) As Currency
    Dim PorRec(2) As Currency
    Dim TotFac As Currency
    
    'retencion
    Dim PorRet As Currency
    Dim ImpRet As Currency


    On Error GoTo EModificarLin

    Select Case NumTabMto
        Case 0: nomFrame = "FrameAux0" 'lineas de factura
    End Select
    
    If DatosOkLlin(nomFrame) Then
        TerminaBloquear
'        conn.BeginTrans
        If ModificaDesdeFormulario2(Me, 2, nomFrame) Then
            If BLOQUEADesdeFormulario2(Me, Data1, 1) Then
            
            ' ### [Monica] 29/09/2006
            ' he quitado el boton modificar para recalcular bases e iva
            
            'BotonModificar
                

                
            End If
            V = AdoAux(NumTabMto).Recordset.Fields(4) 'el 2 es el nº de llinia
            ModoLineas = 0
            CargaGrid NumTabMto, True
'            SituarTab (NumTabMto)
            DataGridAux(NumTabMto).SetFocus
'            AdoAux(NumTabMto).Recordset.Find (AdoAux(NumTabMto).Recordset.Fields(1).Name & " =" & V)
            
'            ' ### [Monica] 29/09/2006
'            ' añadido el tema de de recalculo de bases
            If BdConta = 0 Then Exit Sub
            
            If AbrirConexionContaFac(vParamAplic.UsuarioContaFac, vParamAplic.PasswordContaFac, CByte(BdConta)) Then
                Set vEmpresaFac = New CempresaFac
                If vEmpresaFac.LeerNiveles Then
                    PorRet = 0
                    If text1(26).Text <> "" Then PorRet = CCur(ImporteSinFormato(text1(26).Text))
                
                    RecalculoBasesIvaFactura AdoAux(0).Recordset, Imptot, Tipiva, Impbas, impiva, PorIva, TotFac, ImpRec, PorRec, PorRet, ImpRet
                End If
                Set vEmpresaFac = Nothing
                CerrarConexionContaFac
            End If

            '13/02/2007 iniacializo los txt
            For i = 0 To 2
                text1(6 + (6 * i)).Text = ""
                text1(7 + (6 * i)).Text = ""
                text1(8 + (6 * i)).Text = ""
                text1(9 + (6 * i)).Text = ""
                text1(10 + (6 * i)).Text = ""
                text1(11 + (6 * i)).Text = ""
            Next i

            '13/02/2007 he añadido las condiciones del for antes solo estaban las sentencias
            For i = 0 To 2
                 If Impbas(i) <> 0 Then text1(6 + (6 * i)).Text = Impbas(i)
                 If Tipiva(i) <> 0 Then text1(7 + (6 * i)).Text = Tipiva(i)
                 If PorIva(i) <> 0 Then text1(8 + (6 * i)).Text = PorIva(i)
                 If impiva(i) <> 0 Then text1(9 + (6 * i)).Text = impiva(i)
                 If PorRec(i) <> 0 Then text1(10 + (6 * i)).Text = PorRec(i)
                 If ImpRec(i) <> 0 Then text1(11 + (6 * i)).Text = ImpRec(i)

                 'TotFac = Impbas(i) + impiva(i)
            Next i
            text1(24).Text = TotFac
            If ImpRet <> 0 Then text1(28).Text = ImpRet
            
            If text1(8).Text = "" Then text1(8).Text = "0,00"
            If text1(9).Text = "" Then text1(9).Text = "0,00"
            
'++monica: 10/03/2009
            PonerFormatos
'++
            
            
            If BLOQUEADesdeFormulario2(Me, Data1, 1) Then
'                Modo = 4
'                PonerModo Modo
'                ClienteAnt = Text1(3).Text
'                FormaPagoAnt = Text1(5).Text
                ModificaImportes = True
'--monica:10/03/2009
'                PonerCamposForma Me, Me.Data1
                BotonModificar
                cmdAceptar_Click
                ModificaImportes = False
            End If

            LLamaLineas NumTabMto, 0
        End If
    End If
    Exit Sub
    
EModificarLin:
    If Err.Number <> 0 Then MuestraError Err.Number, "Modificar Linea", Err.Description
End Sub

Private Function ObtenerWhereCab(conW As Boolean) As String
Dim vWhere As String
    vWhere = ""
    If conW Then vWhere = " WHERE "
    vWhere = vWhere & " codsecci = " & text1(3).Text
    vWhere = vWhere & " AND letraser='" & Trim(text1(0).Text) & "'"
    vWhere = vWhere & " AND numfactu= " & text1(1).Text & " AND fecfactu= '" & Format(text1(2).Text, FormatoFecha) & "'"
    ObtenerWhereCab = vWhere
End Function



Private Function SumaLineas(NumLin As String) As String
'Al Insertar o Modificar linea sumamos todas las lineas excepto la que estamos
'Insertando o modificando que su valor sera el del txtaux(4).text
'En el DatosOK de la factura sumamos todas las lineas
Dim SQL As String
Dim Rs As ADODB.Recordset
Dim SumLin As Currency

    SumLin = 0
    SQL = "SELECT SUM(importe) FROM linfact "
    SQL = SQL & ObtenerWhereCab(True)
    If NumLin <> "" Then SQL = SQL & " AND numlinea<>" & DBSet(txtAux(4).Text, "N") 'numlinea
    Set Rs = New ADODB.Recordset
    Rs.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not Rs.EOF Then
        'En SumLin tenemos la suma de las lineas ya insertadas
        SumLin = CCur(DBLet(Rs.Fields(0), "N"))
    End If
    Rs.Close
    Set Rs = Nothing
    SumaLineas = CStr(SumLin)
End Function


'' ### [DavidV] 26/04/2006: Activar/desactivar la rueda del ratón.
'Private Sub DataGridAux_GotFocus(Index As Integer)
'  WheelHook DataGridAux(Index)
'End Sub
'Private Sub DataGridAux_LostFocus(Index As Integer)
'  WheelUnHook
'End Sub


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
        'Nuevo. Febrero 2010
        .outClaveNombreArchiv = text1(0).Text & Format(text1(1).Text, "0000000")
        .outCodigoCliProv = text1(4).Text
        .outTipoDocumento = 1
    
        .FormulaSeleccion = cadFormula
        .OtrosParametros = cadParam
        .NumeroParametros = 2
        .SoloImprimir = False
        .Titulo = cadTitulo
        .NombreRPT = cadNombreRPT
        .Facturas = True
        .Contabilidad = BdConta
        .EnvioEMail = False
        .Opcion = 1
        .Show vbModal
    End With
End Sub


Private Sub ActivarFrameCobros()
Dim obj As Object

For Each obj In Me
    If TypeOf obj Is Frame Then
        If obj.Name = "FrameCobros" Then
            
            
        End If
        
    End If
Next obj

End Sub


Private Sub EliminarLinea()
Dim nomFrame As String
Dim V As Currency
Dim SQL As String

    
 
' variables para el recalculo de iva y totales
    Dim i As Integer
    Dim Imptot(2)
    Dim Tipiva(2)
    Dim Impbas(2) As Currency
    Dim impiva(2) As Currency
    Dim PorIva(2) As Currency
    Dim ImpRec(2) As Currency
    Dim PorRec(2) As Currency
    Dim TotFac As Currency

    'retencion
    Dim PorRet As Currency
    Dim ImpRet As Currency


    On Error GoTo EEliminarLin

    Select Case NumTabMto
        Case 0: nomFrame = "FrameAux0" 'lineas de factura
    End Select
    

    TerminaBloquear
'        conn.BeginTrans
    If BLOQUEADesdeFormulario2(Me, Data1, 1) Then

            ' ### [Monica] 29/09/2006
            ' he quitado el boton modificar para recalcular bases e iva

            'BotonModificar

            End If
            ModoLineas = 0
'            V = AdoAux(NumTabMto).Recordset.Fields(4) 'el 2 es el nº de llinia
            CargaGrid NumTabMto, True

'            SituarTab (NumTabMto)

' [Monica] 25/01/2010 Daba error cuando elimina linea he quitado el setfocus
'            DataGridAux(NumTabMto).SetFocus

'            AdoAux(NumTabMto).Recordset.Find (AdoAux(NumTabMto).Recordset.Fields(1).Name & " =" & V)

'            ' ### [Monica] 29/09/2006
'            ' añadido el tema de de recalculo de bases
            PorRet = 0
            If text1(26).Text <> "" Then PorRet = CCur(ImporteSinFormato(text1(26).Text))

            RecalculoBasesIvaFactura AdoAux(0).Recordset, Imptot, Tipiva, Impbas, impiva, PorIva, TotFac, ImpRec, PorRec, PorRet, ImpRet


            '13/02/2007 iniacializo los txt
            For i = 0 To 2
                text1(6 + (6 * i)).Text = ""
                text1(7 + (6 * i)).Text = ""
                text1(8 + (6 * i)).Text = ""
                text1(9 + (6 * i)).Text = ""
                text1(10 + (6 * i)).Text = ""
                text1(11 + (6 * i)).Text = ""
            Next i

            '13/02/2007 he añadido las condiciones del for antes solo estaban las sentencias
            For i = 0 To 2
                 If Impbas(i) <> 0 Then text1(6 + (6 * i)).Text = Impbas(i)
                 If Tipiva(i) <> 0 Then text1(7 + (6 * i)).Text = Tipiva(i)
                 If PorIva(i) <> 0 Then text1(8 + (6 * i)).Text = PorIva(i)
                 If impiva(i) <> 0 Then text1(9 + (6 * i)).Text = impiva(i)
                 If PorRec(i) <> 0 Then text1(10 + (6 * i)).Text = PorRec(i)
                 If ImpRec(i) <> 0 Then text1(11 + (6 * i)).Text = ImpRec(i)

                 'TotFac = Impbas(i) + impiva(i)
            Next i
            text1(24).Text = TotFac
            If ImpRet <> 0 Then text1(28).Text = ImpRet
            
            If text1(8).Text = "" Then text1(8).Text = "0,00"
            If text1(9).Text = "" Then text1(9).Text = "0,00"
            
            
'            If BLOQUEADesdeFormulario2(Me, Data1, 1) Then
'                ModificaImportes = True
'                BotonModificar
'                cmdAceptar_Click
'            End If

'++monica: 10/03/2009
            PonerFormatos
'++
            LLamaLineas NumTabMto, 0
    Exit Sub
    
EEliminarLin:
    If Err.Number <> 0 Then MuestraError Err.Number, "Eliminar Linea", Err.Description
End Sub

Private Sub PonerFormatos()
Dim mTag As CTag
Dim i As Integer

    Set mTag = New CTag
    For i = 6 To 24
        mTag.Cargar text1(i)
        If mTag.Formato <> "" And CStr(text1(i).Text) <> "" Then
             text1(i).Text = Format(text1(i).Text, mTag.Formato)
        End If
    Next i
    Set mTag = Nothing

End Sub

Private Sub BotonCargaMasiva()
    frmCargaFactVar.Show vbModal
End Sub


Private Function EsFechaOKConta(Fecha As Date) As Byte
Dim F2 As Date

    If vEmpresaFac.FechaIni > Fecha Then
        EsFechaOKConta = 1
    Else
        F2 = DateAdd("yyyy", 1, vEmpresaFac.FechaFin)
        If Fecha > F2 Then
            EsFechaOKConta = 2
        Else
            'OK. Dentro de los ejercicios contables
            EsFechaOKConta = 0
        End If
    End If
    '[Monica]20/06/2017: de david
    If EsFechaOKConta = 0 Then
        'Si tiene SII
        If vParamAplic.ContabilidadNueva Then
            If vEmpresaFac.TieneSII Then
'                If DateDiff("d", Fecha, Now) > vEmpresaFac.SIIDiasAviso Then
                '[Monica]19/02/2018: fines de semana
                If Fecha < UltimaFechaCorrectaSII(vEmpresaFac.SIIDiasAviso, Now) Then
                    MensajeFechaOkConta = "Fecha fuera de periodo de comunicación SII."
                    'LLEVA SII y han trascurrido los dias
                    If vSesion.Nivel = 0 Then
                        If MsgBox(MensajeFechaOkConta & vbCrLf & "¿Continuar?", vbQuestion + vbYesNoCancel) <> vbYes Then
                            EsFechaOKConta = 4
                        End If
                    Else
                        'NO tienen nivel
                        EsFechaOKConta = 5
                    End If
                End If
            End If
        End If
    Else
        MensajeFechaOkConta = "Fuera de ejercicios contables"
    End If

End Function


Private Function EsCuentaMultiple(Codmacta As String) As Boolean
Dim SQL As String
Dim Rs As ADODB.Recordset

    EsCuentaMultiple = False


    If BdConta = 0 Then Exit Function

    If AbrirConexionContaFac(vParamAplic.UsuarioContaFac, vParamAplic.PasswordContaFac, CByte(BdConta)) Then

        SQL = "select esctamultiple from cuentas where codmacta = " & DBSet(Codmacta, "T")
        
        Set Rs = New ADODB.Recordset
        Rs.Open SQL, ConnContaFac, adOpenForwardOnly, adLockReadOnly, adCmdText
        
        If Not Rs.EOF Then
            SQL = DBLet(Rs!esctamultiple, "N")
        Else
            SQL = 0
        End If
        EsCuentaMultiple = (SQL = 1)
        Set Rs = Nothing
        
    End If
    
End Function

Private Sub BloquearDatosFiscales(bloqueo As Boolean)
Dim i As Integer
    
    If Modo = 5 Then Exit Sub


    Text2(4).Enabled = Not bloqueo
    Text2(4).Locked = bloqueo
    
    For i = 29 To 34
        text1(i).Enabled = Not bloqueo
        text1(i).Locked = bloqueo
    Next i
    
    imgBuscar(7).Enabled = Not bloqueo
    imgBuscar(7).visible = Not bloqueo
    
    
    
    If Text2(4).Enabled Then ' blanco
        Text2(4).BackColor = &H80000005
    Else ' amarillo
        Text2(4).BackColor = &H80000018
    End If
End Sub


Private Sub TraerDatosCuenta(Cuenta As String)
Dim SQL As String
Dim Rs As ADODB.Recordset

    On Error GoTo eTraerDatosCuenta

    SQL = "select * from cuentas where codmacta = " & DBSet(Cuenta, "T")
    
    Set Rs = New ADODB.Recordset
    Rs.Open SQL, ConnContaFac, adOpenForwardOnly, adLockReadOnly, adCmdText
    
    If Not Rs.EOF Then
        text1(29).Text = DBLet(Rs!nifdatos, "T")
        text1(30).Text = DBLet(Rs!dirdatos, "T")
        text1(31).Text = DBLet(Rs!Codposta, "T")
        text1(32).Text = DBLet(Rs!desPobla, "T")
        text1(33).Text = DBLet(Rs!desProvi, "T")
        If vParamAplic.ContabilidadNueva Then
            text1(34).Text = DBLet(Rs!codpais, "T")
        Else
            text1(34).Text = Mid(DBLet(Rs!paise, "T"), 1, 2)
        End If
    End If
    Set Rs = Nothing
    Exit Sub
     
eTraerDatosCuenta:
    MuestraError Err.Number, "Traer Datos Cuenta", Err.Description
End Sub



