VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmHcoFactGas 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Histórico de Facturas de Gasolinera"
   ClientHeight    =   6285
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9075
   Icon            =   "frmHcoFactGas.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6285
   ScaleWidth      =   9075
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Height          =   885
      Index           =   0
      Left            =   240
      TabIndex        =   25
      Top             =   480
      Width           =   8625
      Begin VB.CheckBox Check1 
         Height          =   255
         Index           =   1
         Left            =   8280
         TabIndex        =   5
         Tag             =   "Contabilizada|N|N|0|1|gascabfac|intconta|||"
         Top             =   450
         Width           =   255
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   1
         Left            =   690
         TabIndex        =   1
         Tag             =   "Nº de Factura|N|N|||gascabfac|numfactu||S|"
         Top             =   450
         Width           =   855
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   3
         Left            =   2745
         MaxLength       =   6
         TabIndex        =   3
         Tag             =   "Socio|N|N|0|999999|gascabfac|codsocio|000000|N|"
         Top             =   450
         Width           =   720
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   4
         Left            =   3555
         TabIndex        =   4
         Tag             =   "Nombre|T|N|||gascabfac|nomsocio||N|"
         Top             =   450
         Width           =   3660
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   0
         Left            =   225
         MaxLength       =   3
         TabIndex        =   0
         Tag             =   "Letra Serie|T|N|||gascabfac|letraser||S|"
         Top             =   450
         Width           =   405
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   2
         Left            =   1575
         MaxLength       =   10
         TabIndex        =   2
         Tag             =   "Fecha Factura|F|N|||gascabfac|fecfactu|dd/mm/yyyy|S|"
         Top             =   450
         Width           =   1080
      End
      Begin VB.Label Label1 
         Caption         =   "Contabilizada"
         Height          =   255
         Index           =   7
         Left            =   7290
         TabIndex        =   44
         Top             =   450
         Width           =   1020
      End
      Begin VB.Label Label1 
         Caption         =   "Socio"
         Height          =   255
         Index           =   13
         Left            =   2745
         TabIndex        =   43
         Top             =   180
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Nº Factura"
         Height          =   255
         Index           =   4
         Left            =   690
         TabIndex        =   28
         Top             =   180
         Width           =   855
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   2
         Left            =   2385
         Picture         =   "frmHcoFactGas.frx":000C
         ToolTipText     =   "Buscar fecha"
         Top             =   180
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Serie"
         Height          =   255
         Index           =   2
         Left            =   225
         TabIndex        =   27
         Top             =   180
         Width           =   630
      End
      Begin VB.Label Label1 
         Caption         =   "Fecha "
         Height          =   255
         Index           =   1
         Left            =   1620
         TabIndex        =   26
         Top             =   180
         Width           =   855
      End
   End
   Begin VB.Frame FrameAux0 
      Caption         =   "Lineas Factura"
      ForeColor       =   &H00972E0B&
      Height          =   2760
      Left            =   225
      TabIndex        =   35
      Top             =   2610
      Width           =   8625
      Begin VB.TextBox txtAux 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   290
         Index           =   8
         Left            =   6885
         MaxLength       =   10
         TabIndex        =   18
         Tag             =   "Precio|N|N|||gaslinfac|preciove|#,##0.000||"
         Text            =   "Pre"
         Top             =   1935
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.TextBox txtAux 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   290
         Index           =   7
         Left            =   5940
         MaxLength       =   10
         TabIndex        =   17
         Tag             =   "Cantidad|N|N|||gaslinfac|cantidad|#,##0.00||"
         Text            =   "Can"
         Top             =   1935
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.TextBox txtAux 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   290
         Index           =   5
         Left            =   2070
         MaxLength       =   10
         TabIndex        =   15
         Tag             =   "Fecha Albaran|F|N|||gaslinfac|fecalbar|dd/mm/yyyy||"
         Text            =   "Fec"
         Top             =   1920
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.TextBox txtAux 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   290
         Index           =   4
         Left            =   3060
         MaxLength       =   6
         TabIndex        =   16
         Tag             =   "Artículo|N|N|0|999999|gaslinfac|codartic|000000||"
         Text            =   "Arti"
         Top             =   1935
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.TextBox txtAux 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   290
         Index           =   0
         Left            =   120
         MaxLength       =   3
         TabIndex        =   11
         Tag             =   "Letra Serie|T|N|||gaslinfac|letraser||S|"
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
         Index           =   3
         Left            =   1680
         MaxLength       =   2
         TabIndex        =   14
         Tag             =   "Número de línea|N|N|1|9999|gaslinfac|numlinea|0000|S|"
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
         Index           =   1
         Left            =   480
         MaxLength       =   7
         TabIndex        =   12
         Tag             =   "Nº Factura|N|N|0|9999999|gaslinfac|numfactu|0000000|S|"
         Text            =   "Fac"
         Top             =   1920
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.TextBox txtAux 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   290
         Index           =   2
         Left            =   1080
         MaxLength       =   10
         TabIndex        =   13
         Tag             =   "Fecha Factura|F|N|||gaslinfac|fecfactu|dd/mm/yyyy|S|"
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
         Index           =   9
         Left            =   7740
         MaxLength       =   15
         TabIndex        =   19
         Tag             =   "Importe|N|N|||gaslinfac|implinea|##,###,##0.00||"
         Text            =   "Importe"
         Top             =   1935
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.TextBox txtAux 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   285
         Index           =   6
         Left            =   3645
         TabIndex        =   36
         Text            =   "NomArtic"
         Top             =   1935
         Visible         =   0   'False
         Width           =   2175
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
      Begin MSDataGridLib.DataGrid DataGridAux 
         Height          =   2310
         Index           =   0
         Left            =   225
         TabIndex        =   39
         Top             =   360
         Width           =   8235
         _ExtentX        =   14526
         _ExtentY        =   4075
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
      Begin MSComctlLib.Toolbar ToolAux 
         Height          =   390
         Index           =   0
         Left            =   240
         TabIndex        =   37
         Top             =   450
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
               Enabled         =   0   'False
               Object.Visible         =   0   'False
               Object.Tag             =   "2"
               Style           =   3
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Enabled         =   0   'False
               Object.Visible         =   0   'False
               Object.ToolTipText     =   "Modificar"
               Object.Tag             =   "2"
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Enabled         =   0   'False
               Object.Visible         =   0   'False
               Object.ToolTipText     =   "Eliminar"
               Object.Tag             =   "2"
            EndProperty
         EndProperty
         Begin VB.CheckBox Check2 
            Caption         =   "Vista previa"
            Height          =   195
            Index           =   1
            Left            =   8400
            TabIndex        =   38
            Top             =   120
            Visible         =   0   'False
            Width           =   1215
         End
      End
   End
   Begin VB.Frame FrameTotFactu 
      Caption         =   "Total Factura"
      ForeColor       =   &H00972E0B&
      Height          =   1035
      Left            =   240
      TabIndex        =   29
      Top             =   1485
      Width           =   8625
      Begin VB.TextBox Text1 
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
         Index           =   7
         Left            =   5490
         MaxLength       =   15
         TabIndex        =   10
         Tag             =   "Total Factura|N|N|0|9999999999.99|gascabfac|total|#,###,###,##0.00|N|"
         Top             =   510
         Width           =   2595
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000014&
         Height          =   285
         Index           =   8
         Left            =   2040
         MaxLength       =   2
         TabIndex        =   7
         Tag             =   "Cod.Iva|N|N|||gascabfac|codiva|00|N|"
         Text            =   "12"
         Top             =   495
         Width           =   525
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000014&
         Height          =   285
         Index           =   9
         Left            =   2700
         MaxLength       =   6
         TabIndex        =   8
         Tag             =   "Porc.IVA|N|N|||gascabfac|porciva|##0.00|N|"
         Text            =   "99.99"
         Top             =   495
         Width           =   645
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   6
         Left            =   3555
         MaxLength       =   15
         TabIndex        =   9
         Tag             =   "Importe IVA|N|S|||gascabfac|iva|#,###,###,##0.00|N|"
         Top             =   495
         Width           =   1605
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   5
         Left            =   240
         MaxLength       =   15
         TabIndex        =   6
         Tag             =   "Base IVA|N|N|||gascabfac|base|#,###,###,##0.00|N|"
         Text            =   "575757575757557"
         Top             =   480
         Width           =   1605
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
         Left            =   5490
         TabIndex        =   34
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "Tipo IVA"
         Height          =   255
         Index           =   14
         Left            =   2040
         TabIndex        =   33
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "% IVA"
         Height          =   255
         Index           =   15
         Left            =   2745
         TabIndex        =   32
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "Importe IVA"
         Height          =   255
         Index           =   16
         Left            =   3555
         TabIndex        =   31
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Base Imponible"
         Height          =   255
         Index           =   10
         Left            =   240
         TabIndex        =   30
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Frame Frame1 
      Height          =   555
      Index           =   0
      Left            =   225
      TabIndex        =   23
      Top             =   5490
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
         TabIndex        =   24
         Top             =   180
         Width           =   2655
      End
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   7695
      TabIndex        =   21
      Top             =   5610
      Width           =   1035
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   6480
      TabIndex        =   20
      Top             =   5610
      Width           =   1035
   End
   Begin MSAdodcLib.Adodc Data1 
      Height          =   375
      Left            =   4185
      Top             =   5670
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
      Left            =   7695
      TabIndex        =   22
      Top             =   5580
      Visible         =   0   'False
      Width           =   1035
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   41
      Top             =   0
      Width           =   9075
      _ExtentX        =   16007
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
            Enabled         =   0   'False
            Object.Visible         =   0   'False
            Object.Tag             =   "2"
            Style           =   3
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.Visible         =   0   'False
            Object.ToolTipText     =   "Modificar"
            Object.Tag             =   "2"
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.Visible         =   0   'False
            Object.Tag             =   "2"
            Style           =   3
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.Visible         =   0   'False
            Object.ToolTipText     =   "Rectificar factura"
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.Visible         =   0   'False
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
         Left            =   6480
         TabIndex        =   42
         Top             =   90
         Width           =   1215
      End
   End
   Begin VB.Label Label1 
      Caption         =   "Empresa"
      Height          =   255
      Index           =   6
      Left            =   360
      TabIndex        =   40
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
         Enabled         =   0   'False
         HelpContextID   =   2
         Shortcut        =   ^N
         Visible         =   0   'False
      End
      Begin VB.Menu mnModificar 
         Caption         =   "&Modificar"
         Enabled         =   0   'False
         HelpContextID   =   2
         Shortcut        =   ^M
         Visible         =   0   'False
      End
      Begin VB.Menu mnEliminar 
         Caption         =   "&Eliminar"
         Enabled         =   0   'False
         HelpContextID   =   2
         Shortcut        =   ^E
         Visible         =   0   'False
      End
      Begin VB.Menu mnRectificar 
         Caption         =   "&Rectificar"
         Enabled         =   0   'False
         Shortcut        =   ^R
         Visible         =   0   'False
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
Attribute VB_Name = "frmHcoFactGas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' +-+-+-+-+-+-+-+-+-+-+-+-
' +-+- Autor: MANOLO  +-+-
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


Dim ClienteAnt As String
Dim FormaPagoAnt As String
Dim ModoModificar As Boolean
Dim ModificaImportes As Boolean ' variable que me indica q hay que modificar lineas de la factura de contabilidad
                                ' y cobros en la tesoreria

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
Dim vTabla As String
Dim CtaClie As String

    Screen.MousePointer = vbHourglass
    On Error GoTo Error1
    
    ModoModificar = False
    b = True
    Select Case Modo
        Case 1  'BUSQUEDA
            HacerBusqueda
            
'        Case 3 'INSERTAR
'            If DatosOk Then
'                If InsertarDesdeForm2(Me, 1) Then
'                    Data1.RecordSource = "Select * from " & NomTabla & Ordenacion
'                    PosicionarData
'                End If
'            Else
'                ModoLineas = 0
'            End If
'
'        Case 4  'MODIFICAR
'            If Not DatosOk Then
'                ModoLineas = 0
'                Screen.MousePointer = vbDefault
'                Exit Sub
'            Else
'                ModoModificar = True
'
'                conn.BeginTrans
'                If vParamAplic.NumeroConta <> 0 Then ConnConta.BeginTrans
'
'                If CadenaBorrado <> "" Then
'                    conn.Execute CadenaBorrado
'                    CadenaBorrado = ""
'                    EliminarLinea
'                End If
'
'
'                If ModificaDesdeFormulario2(Me, 1) Then
'                    If vParamAplic.NumeroConta <> 0 And Check1(1).Value = 1 Then
'                        'solo en el caso de que este contabilizada
'                        If Val(ClienteAnt) <> Val(Text1(3).Text) Then
'                            CtaClie = ""
'                            CtaClie = DevuelveDesdeBDNew(cPTours, "ssocio", "codmacta", "codsocio", Text1(3).Text, "N")
'                            b = ModificaClienteFacturaContabilidad(Text1(0).Text, Text1(1).Text, Text1(2).Text, CtaClie, tipo)
'                        End If
'' 09022007 ya no dejo modificar la forma de pago
''                        If Val(FormaPagoAnt) <> Val(Text1(5).Text) Then _
''                            ModificaFormaPagoTesoreria Text1(0).Text, Text1(1).Text, Text1(2).Text, Text1(5).Text, FormaPagoAnt, TipForpa, TipForpaAnt
'
'                        If ModificaImportes And b Then
'                            BorrarTMPErrFact
'                            If tipo = 0 Then
'                                vtabla = "schfac"
'                            Else
'                                vtabla = "schfacr"
'                            End If
'                            b = ModificaImportesFacturaContabilidad(Text1(0).Text, Text1(1).Text, Text1(2).Text, Text1(18).Text, Text1(5).Text, vtabla)
'                            ModificaImportes = False
'                        End If
'                    End If
'                    TerminaBloquear
'                    PosicionarData "letraser = '" & Text1(0).Text & "' and numfactu = " & Text1(1).Text & " and fecfactu = " & DBSet(Text1(2).Text, "F")
'                End If
'            End If
            
'        Case 5 'LLINIES
'            Select Case ModoLineas
''                Case 1 'afegir llinia
''                    InsertarLinea
'                Case 2 'modificar llinies
'                    ModificarLinea
'                    PosicionarData "letraser = '" & Text1(0).Text & "' and numfactu = " & Text1(1).Text & " and fecfactu = " & DBSet(Text1(2).Text, "F")
'                    Screen.MousePointer = vbDefault
'                    Exit Sub
'                Case 3 'eliminar llinies
'                    ModificarLinea
'                    PosicionarData "letraser = '" & Text1(0).Text & "' and numfactu = " & Text1(1).Text & " and fecfactu = " & DBSet(Text1(2).Text, "F")
'
'            End Select
    End Select
    Screen.MousePointer = vbDefault
 
Error1:
    Screen.MousePointer = vbDefault
    If Err.Number <> 0 Or Not b Then
        MsgBox Err.Number & ": " & Err.Description, vbExclamation
'        If ModoModificar Then
'            conn.RollbackTrans
'            ConnConta.RollbackTrans
'        End If
'    Else
'        If ModoModificar Then
'            conn.CommitTrans
'            ConnConta.CommitTrans
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
        .Buttons(10).Image = 16 ' Rectificativas
        'el 10 i el 11 son separadors
        .Buttons(12).Image = 10  'Imprimir
        .Buttons(13).Image = 11  'Salir
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
'    For i = 0 To Me.imgBuscar.Count - 1
'        Me.imgBuscar(i).Picture = frmPpal.imgListImages16.ListImages(1).Picture
'    Next i
   
    
    LimpiarCampos   'Limpia los campos TextBox
    For i = 0 To DataGridAux.Count - 1 'neteje tots els grids de llinies
        DataGridAux(i).ClearFields
    Next i
    
    '## A mano
    NomTabla = "gascabfac"
 
    Ordenacion = " ORDER BY letraser, numfactu, fecfactu "
    
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
        text1(0).Text = LetraSerie
        text1(1).Text = numfactu
        PonerModo 1
        cmdAceptar_Click
    End If


End Sub

Private Sub LimpiarCampos()
    On Error Resume Next

    Limpiar Me   'Metodo general: Limpia los controles TextBox
    lblIndicador.Caption = ""
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
    
'    BloquearImgBuscar Me, Modo, ModoLineas
       
    'Bloquear los campos de clave primaria, NO se puede modificar
    BloquearTxt text1(i), Not (Modo = 1)
'    b = (Modo = 3) Or (Modo = 1)  'solo al insertar/buscar estará activo
'    For i = 0 To 2
'        BloquearTxt Text1(i), Not b
'    Next i
'    'Los % de IVA siempre bloqueados
'    BloquearTxt Text1(8), True
'    BloquearTxt Text1(12), True
'    BloquearTxt Text1(16), True
'    'El total de la factura siempre bloqueado
'    BloquearTxt Text1(18), True
'    BloquearTxt Text1(19), True
'
'    '09/02/2007 no dejo modificar la forma de pago
'    BloquearTxt Text1(5), Not b
'
'
    text1(7).BackColor = &HCAE3FD
'    Text1(19).BackColor = &HC0C0FF

'    b = (Modo = 3) Or (Modo = 1) Or (Modo = 4)
'    Me.imgBuscar(0).Enabled = b
'    Me.imgBuscar(0).visible = b
'    Me.imgBuscar(1).Enabled = b
'    Me.imgBuscar(1).visible = b
'    Me.imgBuscar(2).Enabled = ((Modo = 3) Or (Modo = 1))
'    Me.imgBuscar(2).visible = ((Modo = 3) Or (Modo = 1))
'
    
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
    
    b = (Modo = 4)
    FrameTotFactu.Enabled = Not b
    
    b = (Modo = 5)
    
    
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
    
    
    b = (Modo = 2 And Data1.Recordset.RecordCount > 0)
    'Modificar
    Toolbar1.Buttons(8).Enabled = b
    Me.mnModificar.Enabled = b
    'eliminar
    Toolbar1.Buttons(9).Enabled = b
    Me.mnEliminar.Enabled = b
    'rectificativas
    Toolbar1.Buttons(10).Enabled = b And (tipo = 0) 'And EsFacturaRectificable(Text1(0).Text)
    Me.mnRectificar.Enabled = b And (tipo = 0) 'And EsFacturaRectificable(Text1(0).Text)
   
    'Imprimir
    'VRS:2.0.1(3)
    Toolbar1.Buttons(12).Enabled = (Modo = 2)
    Me.mnImprimir.Enabled = (Modo = 2)
    '-----------  LINEAS
    ' *** MEU: botons de les llínies de cuentas bancarias,
    ' només es poden gastar quan inserte o modifique clients ****
    'b = (Modo = 3 Or Modo = 4)
    b = (Modo = 3 Or Modo = 4 Or Modo = 2)
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
            tabla = "gaslinfac"
            SQL = "SELECT letraser,numfactu,fecfactu,numlinea,fecalbar,"
            SQL = SQL & "codartic, nomartic, cantidad, preciove, implinea "
            SQL = SQL & " FROM gaslinfac "
            SQL = SQL & " WHERE 1 = 1 "

            If enlaza Then
                SQL = SQL & " AND " & ObtenerWhereCab(False)
            Else
                SQL = SQL & " AND numfactu = -1"
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
        Aux = ValorDevueltoFormGrid(text1(0), CadenaDevuelta, 1) 'numfactu
        CadB = Aux
        Aux = ValorDevueltoFormGrid(text1(1), CadenaDevuelta, 2) 'fecfactu
        CadB = CadB & " AND " & Aux
        Aux = ValorDevueltoFormGrid(text1(2), CadenaDevuelta, 3) 'codprove
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

Private Sub mnBuscar_Click()
    BotonBuscar
    Me.Check1(1).Value = 0
End Sub

Private Sub mnEliminar_Click()
    BotonEliminar
End Sub

Private Sub mnImprimir_Click()
Dim SQL As String

    'VRS:2.0.1(3): añadido el boton de imprimir
    cadTitulo = "Reimpresion de Facturas"

    ' ### [Monica] 11/09/2006
    '****************************
    Dim indRPT As Byte 'Indica el tipo de Documento en la tabla "scryst"
    Dim nomDocu As String 'Nombre de Informe rpt de crystal

    indRPT = 2 'Facturas Socios de Gasolinera

    If Not PonerParamRPT(indRPT, cadParam, numParam, nomDocu) Then Exit Sub
    'Nombre fichero .rpt a Imprimir
        
'    nomDocu = Replace(nomDocu, ".rpt", "Aj" & "C" & Format(Data1.Recordset!codcoope, "00") & ".rpt")
    
    frmImprimir.NombreRPT = nomDocu
    ' he añadido estas dos lineas para que llame al rpt correspondiente

    cadNombreRPT = nomDocu  ' "rFactgas.rpt"
    cadFormula = "({" & NomTabla & ".letraser} = """ & text1(0).Text & """) AND ({" & NomTabla & ".numfactu} = " & text1(1).Text & ") and ({" & NomTabla & ".fecfactu} = cdate(""" & text1(2).Text & """)) "
    
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
'    If vParamAplic.NumeroConta <> 0 And Not FacturaModificable(Text1(0).Text, Text1(1).Text, Text1(2).Text, Check1(1).Value) Then Exit Sub
    
    'Preparar para modificar
    '-----------------------
    If BLOQUEADesdeFormulario2(Me, Data1, 1) Then BotonModificar
End Sub

'Private Sub mnNuevo_Click()
'     BotonAnyadir
'End Sub

Private Sub mnRectificar_Click()

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
'    If vParamAplic.NumeroConta <> 0 And Not FacturaModificable(Text1(0).Text, Text1(1).Text, Text1(2).Text, Check1(1).Value) Then Exit Sub
    
    'Preparar para modificar
    '-----------------------
'    If Not FacturaModificable(Text1(0).Text, Text1(1).Text, Text1(2).Text, Check1(1).Value) Then Exit Sub
    If BLOQUEADesdeFormulario2(Me, Data1, 1) Then BotonRectificar
End Sub


Private Sub mnSalir_Click()
    Unload Me
End Sub

Private Sub mnVerTodos_Click()
    BotonVerTodos
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
        Case 3  'Buscar
           mnBuscar_Click
        Case 4  'Todos
            mnVerTodos_Click
'        Case 7  'Nuevo
'            mnNuevo_Click
        Case 8  'Modificar
            mnModificar_Click
'        Case 9  'Borrar
'            mnEliminar_Click
        Case 10 'Rectificativa
            mnRectificar_Click
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
Dim cad As String
        'Llamamos a al form
        '##A mano
        cad = ""
        cad = cad & ParaGrid(text1(0), 7, "Serie")
        cad = cad & ParaGrid(text1(1), 16, "Nº Fact.")
        cad = cad & ParaGrid(text1(2), 15, "Fecha")
        'los ponemos a mano:
        cad = cad & "Socio.|codsocio|N|" & FormatoCampo(text1(3)) & "|12·"
        cad = cad & "Nom. Socio|nomsocio|T||50·"
        If cad <> "" Then
            Screen.MousePointer = vbHourglass
            Set frmB = New frmBuscaGrid
            frmB.vCampos = cad
            frmB.vTabla = NomTabla
            frmB.vSQL = CadB
            HaDevueltoDatos = False
            '###A mano
            frmB.vDevuelve = "0|1|2|4|"
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
Dim cad As String
Dim Aux As String
Dim i As Integer
Dim J As Integer

    If Data1.Recordset.EOF Then
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
            cad = cad & text1(J).Text & "|"
        End If
    Loop Until i = 0
    RaiseEvent DatoSeleccionado(cad)
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

'Private Sub BotonAnyadir()
''Añadir registro en tabla de expedientes individuales: expincab (Cabecera)
'
'    LimpiarCampos 'Vacía los TextBox
'    'Poner los grid sin apuntar a nada
''    LimpiarDataGrids
'
'    'Añadiremos el boton de aceptar y demas objetos para insertar
'    PonerModo 3
'
'    'Quan afegixc pose en Fecha
'    Text1(2).Text = Format(Now, "dd/mm/yyyy")
'
'    'Total Factura (por defecto=0)
'    Text1(18).Text = "0"
'    Text1(19).Text = "0"
'
'    'em posicione en el 1r tab
'    PonerFoco Text1(0)
'End Sub

Private Sub BotonModificar()
    
    'Añadiremos el boton de aceptar y demas objetos para insertar
    PonerModo 4
    
    ' ### [Monica] 27/09/2006
    ' me guardo los valores anteriores de cliente y forma de pago
    ClienteAnt = text1(3).Text
    FormaPagoAnt = text1(5).Tag
    
    'Quan modifique pose en la F.Modificación la data actual
    PonerFoco text1(3)
End Sub

Private Sub BotonRectificar()
    
'    Set frmList = New frmListado
'    'Añadiremos el boton de aceptar y demas objetos para insertar
'    frmList.CadTag = Text1(0).Text & "|" & Text1(1).Text & "|" & Text1(2).Text & "|" & Text1(3).Text & "|" & Text2(3).Text & "|" & Format(Check1(1).Value, "0") & "|"
'    frmList.OpcionListado = 12
'    frmList.Show vbModal

End Sub

Private Sub BotonEliminar()
Dim cad As String

    On Error GoTo EEliminar

    'Ciertas comprobaciones
    If Data1.Recordset.EOF Then Exit Sub
    
    'El registre de codi 0 no es pot Modificar ni Eliminar
    If EsCodigoCero(CStr(Data1.Recordset.Fields(1).Value), FormatoCampo(text1(1))) Then Exit Sub

    cad = "¿Seguro que desea eliminar la factura?"
    cad = cad & vbCrLf & "Nº: " & Format(Data1.Recordset!numfactu, FormatoCampo(text1(1)))
    cad = cad & vbCrLf & "Fecha: " & Data1.Recordset.Fields("fecfactu")
    
    'Borramos
    If MsgBox(cad, vbQuestion + vbYesNo) = vbYes Then
        'Hay que eliminar
        On Error GoTo EEliminar
        Screen.MousePointer = vbHourglass
        NumRegElim = Data1.Recordset.AbsolutePosition
        If Not eliminar Then
            Screen.MousePointer = vbDefault
            Exit Sub
        ElseIf SituarDataTrasEliminar(Data1, NumRegElim) Then
            PonerCampos
        Else
            LimpiarCampos
            'Poner los grid sin apuntar a nada
            'LimpiarDataGrids
            PonerModo 0
        End If
    End If
    Screen.MousePointer = vbDefault
    
EEliminar:
    Screen.MousePointer = vbDefault
    If Err.Number > 0 Then MuestraError Err.Number, "Eliminar Factura", Err.Description
End Sub

Private Sub PonerCampos()
Dim i As Integer

    If Data1.Recordset.EOF Then Exit Sub
    PonerCamposForma2 Me, Data1, 1 'opcio=1: pone el formato o los campos de la cabecera
    
    For i = 0 To DataGridAux.Count - 1
        CargaGrid i, True
    Next i
    
    'Recuperar Descripciones de los campos de Codigo
    '--------------------------------------------------
'    text2(0).Text = vParamAplic.CodIvaGas
'    text2(1).Text = DevuelveDesdeBDNew(cContaGas, "tiposiva", "porceiva", "codigiva", vParamAplic.CodIvaGas, "N")

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
'                Case 1 'afegir llinia
'                    ModoLineas = 0
'                    DataGridAux(NumTabMto).AllowAddNew = False
''                    SituarTab (NumTabMto)
'                    'If BLOQUEADesdeFormulario2(Me, Data1, 1) Then BotonModificar  'Modificar
'                    LLamaLineas NumTabMto, ModoLineas 'ocultar txtAux
'                    'If DataGridAux(NumTabMto).Enabled Then DataGridAux(NumTabMto).SetFocus
'                    DataGridAux(NumTabMto).Enabled = True
'                    DataGridAux(NumTabMto).SetFocus
'
'                    If Not AdoAux(NumTabMto).Recordset.EOF Then
'                        AdoAux(NumTabMto).Recordset.MoveFirst
'                    End If

                Case 2 'modificar llinies
                    ModoLineas = 0
                    PonerModo 4
                    If Not AdoAux(NumTabMto).Recordset.EOF Then
                        V = AdoAux(NumTabMto).Recordset.Fields(3) 'el 1 es el nº de llinia
                        AdoAux(NumTabMto).Recordset.Find (AdoAux(NumTabMto).Recordset.Fields(1).Name & " =" & V)
                    End If
                    LLamaLineas NumTabMto, ModoLineas 'ocultar txtAux
            End Select
            
            PosicionarData "letraser = '" & text1(0).Text & "' and numfactu = " & text1(1).Text & " and fecfactu = " & DBSet(text1(2).Text, "F")
            
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

    On Error GoTo EDatosOK

    DatosOk = False
    b = CompForm2(Me, 1) 'Comprobar formato datos ok
    If Not b Then Exit Function
    
    ' en caso de que haya contabilidad
    ' comprobamos si el cliente tiene cuenta contable existente en la contabilidad
    If vParamAplic.NumeroConta <> 0 And Check1(1).Value <> 0 Then
        SQL = ""
        SQL = DevuelveDesdeBD("codmacta", "ssocio", "codsocio", text1(3).Text, "N")
        If SQL = "" Then
            MsgBox "El cliente no tiene cuenta contable asociada.", vbExclamation
            Exit Function
        Else
            Datos = ""
            Datos = DevuelveDesdeBDNew(cConta, "cuentas", "codmacta", "codmacta", SQL, "T")
            If Datos = "" Then
                MsgBox "La cuenta contable asociada al cliente no está dada de alta en contabilidad. Revise.", vbExclamation
                Exit Function
            End If
        End If
    End If
    
    'Comprobamos que la suma de importes de las lineas es igual al total de la factura
    Datos = SumaLineas("")
    
    If CCur(Datos) > (CCur(text1(18).Text) + CCur(text1(19).Text)) Then
        MsgBox "La suma de los importes de lineas es mayor que el total de la factura!!!", vbExclamation
    ElseIf CCur(Datos) < CCur(text1(18).Text) Then
        MsgBox "La suma de los importes de lineas es menor que el total de la factura!!!", vbExclamation
    End If
         
    DatosOk = b
EDatosOK:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Function

Private Sub PosicionarData(cad As String)
'Dim cad As String
Dim Indicador As String
    
  '  cad = ""
    If SituarDataMULTI(Data1, cad, Indicador) Then
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
    conn.Execute "DELETE FROM slhfac " & vWhere
    
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
End Sub

Private Sub Text1_LostFocus(Index As Integer)
Dim cad As String, Datos As String
Dim Suma As Currency

    If Not PerderFocoGnral(text1(Index), Modo) Then Exit Sub
    
    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub
    
    Select Case Index
        Case 0 ' letra de serie
            If text1(Index).Text <> "" Then text1(Index).Text = UCase(text1(Index).Text)
            
        Case 1 'Nº factura
            If text1(Index).Text <> "" Then FormateaCampo text1(Index)
            
        Case 2 'Fecha
            If text1(Index).Text <> "" Then PonerFormatoFecha text1(Index)
            
        Case 3 'Socio
            If text1(Index).Text <> "" Then FormateaCampo text1(Index)
            
    End Select
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = teclaBuscar Then
        If Modo = 1 Or Modo = 3 Or Modo = 4 Then
            Select Case Index
                Case 2: KEYFecha KeyAscii, 2
            End Select
        End If
    Else
        KEYpress KeyAscii
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

Private Sub KEYFecha(KeyAscii As Integer, indice As Integer)
    KeyAscii = 0
    imgFec_Click (indice)
End Sub

Private Sub KEYBusquedaLin(KeyAscii As Integer, indice As Integer)
    KeyAscii = 0
'    btnBuscar_Click (indice)
End Sub

'************* LLINIES: ****************************
Private Sub ToolAux_ButtonClick(Index As Integer, ByVal Button As MSComctlLib.Button)

'    If vParamAplic.NumeroConta <> 0 And _
'       Not FacturaModificable(Text1(0).Text, Text1(1).Text, Text1(2).Text, Check1(1).Value) Then Exit Sub
    
    Select Case Button.Index
'        Case 1
''            TerminaBloquear
'            BotonAnyadirLinea Index
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

    If AdoAux(Index).Recordset.EOF Then Exit Sub
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
            SQL = SQL & vbCrLf & "Nº línea: " & Format(DBLet(AdoAux(Index).Recordset!NumLinea), FormatoCampo(txtAux(3)))
            SQL = SQL & vbCrLf & "Albaran: " & DBLet(AdoAux(Index).Recordset!numalbar) '& "  " & txtAux(4).Text
            If MsgBox(SQL, vbQuestion + vbYesNo) = vbYes Then
                NumRegElim = AdoAux(Index).Recordset.AbsolutePosition
                eliminar = True
                SQL = "DELETE FROM slhfac"
                SQL = SQL & ObtenerWhereCab(True) & " AND numlinea= " & AdoAux(Index).Recordset!NumLinea
            End If
    End Select

    If eliminar Then
        TerminaBloquear
'        conn.Execute sql
        CadenaBorrado = SQL
        '16022007
        If BLOQUEADesdeFormulario2(Me, Data1, 1) Then
                ModificaImportes = True
                BotonModificar
                cmdAceptar_Click
        End If
        'EliminarLinea
        
        
        'antes estaba debajo de situardata
        CargaGrid Index, True
        SituarDataTrasEliminar AdoAux(Index), NumRegElim, True
    End If

    ModoLineas = 0
    PosicionarData "letraser = '" & text1(0).Text & "' and numfactu = " & text1(1).Text & " and fecfactu = " & DBSet(text1(2).Text, "F")

    Exit Sub
Error2:
    Screen.MousePointer = vbDefault
    MuestraError Err.Number, "Eliminando linea", Err.Description
End Sub

'Private Sub BotonAnyadirLinea(Index As Integer)
'Dim NumF As String
'Dim vWhere As String, vTabla As String
'Dim anc As Single
'Dim i As Integer
'Dim SumLin As Currency
'
'    'Si no estaba modificando lineas salimos
'    ' Es decir, si estaba insertando linea no podemos hacer otra cosa
'    'If ModificaLineas = 2 Then Exit Sub
'    ModoLineas = 1 'Ponemos Modo Añadir Linea
'
'    If (Modo = 3) Or (Modo = 4) Then 'Insertar o Modifcar Cabecera
'        cmdAceptar_Click
'        'No se ha insertado la cabecera
'        If ModoLineas = 0 Then Exit Sub
'    End If
'
'    NumTabMto = Index
'    PonerModo 5
''    If b Then BloquearText1 Me, 4 'Si viene de Insertar Cabecera no bloquear los Text1
'
'
'    'Obtener el numero de linea ha insertar
'    Select Case Index
'        Case 0: vTabla = "slhfac"
'    End Select
'    'Obtener el sig. nº de linea a insertar
'    vWhere = ObtenerWhereCab(False)
'    NumF = SugerirCodigoSiguienteStr(vTabla, "numlinea", vWhere)
'
'    'Situamos el grid al final
'    AnyadirLinea DataGridAux(Index), AdoAux(Index)
'
'    anc = DataGridAux(Index).Top
'    If DataGridAux(Index).Row < 0 Then
'        anc = anc + 210
'    Else
'        anc = anc + DataGridAux(Index).RowTop(DataGridAux(Index).Row) + 5
'    End If
'
'    LLamaLineas Index, ModoLineas, anc
'
'    Select Case Index
'        Case 0 'lineas factura
'            txtAux(0).Text = Text1(0).Text 'serie
'            txtAux(1).Text = Text1(1).Text 'factura
'            txtAux(2).Text = Text1(2).Text 'fecha
'            txtAux(3).Text = NumF 'numlinea
'            FormateaCampo txtAux(3)
'            For i = 4 To 12
'                txtAux(i).Text = ""
'            Next i
'
'            'desbloquear la linea (se bloquea al añadir)
'            BloquearTxt txtAux(3), False
'            PonerFoco txtAux(4)
'    End Select
'End Sub

Private Sub BotonModificarLinea(Index As Integer)
'    Dim anc As Single
'    Dim i As Integer
'    Dim J As Integer
'
'    If AdoAux(Index).Recordset.EOF Then Exit Sub
'    If AdoAux(Index).Recordset.RecordCount < 1 Then Exit Sub
'
'    ModoLineas = 2 'Modificar llínia
'
'    If Modo = 4 Then 'Modificar Cabecera
'        cmdAceptar_Click
'        If ModoLineas = 0 Then Exit Sub
'    End If
'
'
'    NumTabMto = Index
'    PonerModo 5
'
'    If DataGridAux(Index).Bookmark < DataGridAux(Index).FirstRow Or DataGridAux(Index).Bookmark > (DataGridAux(Index).FirstRow + DataGridAux(Index).VisibleRows - 1) Then
'        i = DataGridAux(Index).Bookmark - DataGridAux(Index).FirstRow
'        DataGridAux(Index).Scroll 0, i
'        DataGridAux(Index).Refresh
'    End If
'
'    anc = DataGridAux(Index).Top
'    If DataGridAux(Index).Row < 0 Then
'        anc = anc + 210
'    Else
'        anc = anc + DataGridAux(Index).RowTop(DataGridAux(Index).Row) + 5
'    End If
'
'    Select Case Index
'        Case 0 'lineas de factura
'            For J = 0 To 9
'                txtAux(J).Text = DataGridAux(Index).Columns(J).Text
'            Next J
''            txtAux(6).Text = DataGridAux(Index).Columns(j).Text
''            txtAux(7).Text = DataGridAux(Index).Columns(j + 1).Text
'
'            txtAux2(0).Text = DataGridAux(Index).Columns(10).Text
'            For J = 10 To 12
'                txtAux(J) = DataGridAux(Index).Columns(J + 1).Text
'            Next J
'
''            txtAux2(6).Text = DataGridAux(Index).Columns(j + 1).Text
'
'    End Select
'
'    LLamaLineas Index, ModoLineas, anc
'
'    Select Case Index
'        Case 0 'lineas de factura
'            PonerFoco txtAux(4)
'    End Select
End Sub

Private Sub LLamaLineas(Index As Integer, xModo As Byte, Optional alto As Single)
Dim jj As Integer
Dim b As Boolean

    On Error GoTo ELLamaLin

    DeseleccionaGrid DataGridAux(Index)
    
    b = (xModo = 1 Or xModo = 2) 'Insertar o Modificar Lineas
    Select Case Index
        Case 0 'lineas de factura
            For jj = 4 To txtAux.Count - 1
                txtAux(jj).Top = alto
                txtAux(jj).visible = b
            Next jj
    End Select
    
ELLamaLin:
    Err.Clear
End Sub

Private Sub txtAux_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = teclaBuscar Then
'        If Modo = 5 And (ModoLineas = 1 Or ModoLineas = 2) Then
'            Select Case Index
'                Case 5: KEYBusquedaLin KeyAscii, 0
'                Case 6: KEYBusquedaLin KeyAscii, 1
'            End Select
'        End If
    Else
        KEYpress KeyAscii
    End If
End Sub

Private Sub txtAux_LostFocus(Index As Integer)
Dim cadMen As String

'    txtAux(Index).Text = Trim(txtAux(Index).Text)
'
'    Select Case Index
'        Case 4 ' albaran
'            txtAux(Index).Text = UCase(txtAux(Index).Text)
'
'        Case 5 ' fecha de albaran
'            PonerFormatoFecha txtAux(Index)
'
'        Case 6 ' hora
'            PonerFormatoHora txtAux(Index)
'
'        Case 7 ' turno
'            If Not EsNumerico(txtAux(Index).Text) Then
'                MsgBox "El turno debe ser numérico.", vbExclamation
'                On Error Resume Next
'                txtAux(Index).Text = ""
'                PonerFoco txtAux(Index)
'                Exit Sub
'            End If
'            FormateaCampo txtAux(Index)
'
'        Case 8 ' tarjeta
'            If Not EsNumerico(txtAux(Index).Text) Then
'                MsgBox "El número de tarjeta debe ser numérico.", vbExclamation
'                On Error Resume Next
'                txtAux(Index).Text = ""
'                PonerFoco txtAux(Index)
'                Exit Sub
'            End If
'            FormateaCampo txtAux(Index)
'
'        Case 9 ' articulo
'            If PonerFormatoEntero(txtAux(Index)) Then
'                txtAux2(0).Text = PonerNombreDeCod(txtAux(Index), "sartic", "nomartic", "codartic", "N")
'                If txtAux2(0).Text = "" Then
'                    cadMen = "No existe el Articulo: " & txtAux(Index).Text & vbCrLf
'                    cadMen = cadMen & "¿Desea crearlo?" & vbCrLf
'                    If MsgBox(cadMen, vbQuestion + vbYesNo) = vbYes Then
'                        Set frmArt = New frmManArtic
'                        frmArt.DatosADevolverBusqueda = "0|1|"
'                        frmArt.NuevoCodigo = txtAux(Index).Text
'                        txtAux(Index).Text = ""
'                        TerminaBloquear
'                        frmArt.Show vbModal
'                        Set frmArt = Nothing
'                    Else
'                        txtAux(Index).Text = ""
'                    End If
'                    PonerFoco txtAux(Index)
'                End If
'            Else
'                txtAux2(0).Text = ""
'            End If
'
'        Case 10 ' cantidad
'           If Not EsNumerico(txtAux(Index).Text) Then
'                MsgBox "La cantidad debe ser numérica.", vbExclamation
'                On Error Resume Next
'                txtAux(Index).Text = ""
'                PonerFoco txtAux(Index)
'                Exit Sub
'            End If
'            'Es numerico
'            PonerFormatoDecimal txtAux(Index), 2
'        Case 11 ' precio
'           If Not EsNumerico(txtAux(Index).Text) Then
'                MsgBox "El Precio debe ser numérico.", vbExclamation
'                On Error Resume Next
'                txtAux(Index).Text = ""
'                PonerFoco txtAux(Index)
'                Exit Sub
'            End If
'            'Es numerico
'            PonerFormatoDecimal txtAux(Index), 2
'
'        Case 12 'Importe
'           If Trim(txtAux(Index).Text) = "" Then
'                PonerFocoBtn Me.cmdAceptar
'                Exit Sub
'           End If
'           If Not EsNumerico(txtAux(Index).Text) Then
'                MsgBox "El Importe debe ser numérico.", vbExclamation
'                On Error Resume Next
'                txtAux(Index).Text = ""
'                PonerFoco txtAux(Index)
'                Exit Sub
'            End If
'            'Es numerico
'            PonerFormatoDecimal txtAux(Index), 3
'            PonerFocoBtn Me.cmdAceptar
'    End Select
'
'    CalcularImporteNue txtAux(10), txtAux(11), txtAux(12), Index - 10
    
End Sub

Private Sub txtAux_GotFocus(Index As Integer)
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
            tots = "N||||0|;N||||0|;N||||0|;N||||0|;"
            tots = tots & "S|txtAux(5)|T|Fecha|1000|;"
            tots = tots & "S|txtAux(4)|T|Articulo|800|;"
            tots = tots & "S|txtAux(6)|T|Denominación|2430|;S|txtAux(7)|T|Cantidad|1100|;"
            tots = tots & "S|txtAux(8)|T|Precio|1000|;S|txtAux(9)|T|Importe|1200|;"
            arregla tots, DataGridAux(Index), Me
    
    End Select
    
    DataGridAux(Index).ScrollBars = dbgAutomatic
    
ECarga:
    If Err.Number <> 0 Then MuestraError Err.Number, "Cargando datos grid: " & DataGridAux(Index).Tag, Err.Description
End Sub

'Private Sub InsertarLinea()
''Inserta registro en las tablas de Lineas: provbanc, provdpto
'Dim nomFrame As String
'Dim b As Boolean
'
'    On Error Resume Next
'
'    Select Case NumTabMto
'        Case 0: nomFrame = "FrameAux0" 'lineas de factura
'    End Select
'
'    If DatosOkLlin(nomFrame) Then
'        TerminaBloquear
'        If InsertarDesdeForm2(Me, 2, nomFrame) Then
'            b = BLOQUEADesdeFormulario2(Me, Data1, 1)
'            CargaGrid NumTabMto, True
'            If b Then BotonAnyadirLinea NumTabMto
'        End If
'    End If
'End Sub

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
    Dim TotFac As Currency
    Dim totimp As Currency



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
            V = AdoAux(NumTabMto).Recordset.Fields(1) 'el 2 es el nº de llinia
            ModoLineas = 0
            CargaGrid NumTabMto, True
'            SituarTab (NumTabMto)
            DataGridAux(NumTabMto).SetFocus
            AdoAux(NumTabMto).Recordset.Find (AdoAux(NumTabMto).Recordset.Fields(1).Name & " =" & V)
            
'            ' ### [Monica] 29/09/2006
'            ' añadido el tema de de recalculo de bases
'            RecalculoBasesIvaFactura AdoAux(0).Recordset, Imptot, Tipiva, Impbas, impiva, PorIva, TotFac, totimp

            
            '13/02/2007 iniacializo los txt
            For i = 0 To 2
                text1(7 + (4 * i)).Text = ""
                text1(6 + (4 * i)).Text = ""
                text1(8 + (4 * i)).Text = ""
                text1(9 + (4 * i)).Text = ""
            Next i
            
            '13/02/2007 he añadido las condiciones del for antes solo estaban las sentencias
            For i = 0 To 2
                 If Tipiva(i) <> 0 Then text1(7 + (4 * i)).Text = Tipiva(i)
                 If Impbas(i) <> 0 Then text1(6 + (4 * i)).Text = Impbas(i)
                 If PorIva(i) <> 0 Then text1(8 + (4 * i)).Text = PorIva(i)
                 If impiva(i) <> 0 Then text1(9 + (4 * i)).Text = impiva(i)
                 'TotFac = Impbas(i) + impiva(i)
            Next i
            text1(19).Text = totimp
            text1(18).Text = TotFac
            If BLOQUEADesdeFormulario2(Me, Data1, 1) Then
'                Modo = 4
'                PonerModo Modo
'                ClienteAnt = Text1(3).Text
'                FormaPagoAnt = Text1(5).Text
                ModificaImportes = True
                BotonModificar
                cmdAceptar_Click

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
    vWhere = vWhere & " letraser='" & text1(0).Text & "'"
    vWhere = vWhere & " AND numfactu= " & text1(1).Text & " AND fecfactu= '" & Format(text1(2).Text, FormatoFecha) & "'"
    ObtenerWhereCab = vWhere
End Function



Private Function SumaLineas(NumLin As String) As String
'Al Insertar o Modificar linea sumamos todas las lineas excepto la que estamos
'Insertando o modificando que su valor sera el del txtaux(4).text
'En el DatosOK de la factura sumamos todas las lineas
Dim SQL As String
Dim RS As ADODB.Recordset
Dim SumLin As Currency

    SumLin = 0
    If tipo = 0 Then
        SQL = "SELECT SUM(implinea) FROM slhfac "
    Else
        SQL = "SELECT SUM(implinea) FROM slhfacr "
    End If
    SQL = SQL & ObtenerWhereCab(True)
    If NumLin <> "" Then SQL = SQL & " AND numlinea<>" & DBSet(txtAux(4).Text, "N") 'numlinea
    Set RS = New ADODB.Recordset
    RS.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not RS.EOF Then
        'En SumLin tenemos la suma de las lineas ya insertadas
        SumLin = CCur(DBLet(RS.Fields(0), "N"))
    End If
    RS.Close
    Set RS = Nothing
    SumaLineas = CStr(SumLin)
End Function

'' ### [DavidV] 26/04/2006: Activar/desactivar la rueda del ratón.
'Private Sub DataGridAux_GotFocus(Index As Integer)
'  WheelHook DataGridAux(Index)
'End Sub
'Private Sub DataGridAux_LostFocus(Index As Integer)
'  WheelUnHook
'End Sub


'Private Function FacturaModificable(letraser As String, numfactu As String, fecfactu As String, Contabil As String) As Boolean
'
'    FacturaModificable = False
'
'    If Contabil = 0 Then
'        FacturaModificable = True
'    Else
'        ' si la factura esta contabilizada tenemos que ver si en la contabilidad esta contabilizada y
'        ' si en la tesoreria esta remesada o cobrada en estos casos la factura no puede ser modificada
'        If FacturaContabilizada(letraser, numfactu, Year(CDate(fecfactu))) Then
'            MsgBox "Factura contabilizada en la Contabilidad, no puede modificarse ni eliminarse."
'            Exit Function
'        End If
'
'        If FacturaRemesada(letraser, numfactu, fecfactu) Then
'            MsgBox "Factura Remesada, no puede modificarse ni eliminarse."
'            Exit Function
'        End If
'
'        If FacturaCobrada(letraser, numfactu, fecfactu) Then
'            MsgBox "Factura Cobrada, no puede modificarse ni eliminarse."
'            Exit Function
'        End If
'
'        FacturaModificable = True
'    End If
'
'End Function

'VRS:2.0.1(3)
Private Sub LlamarImprimir()
    With frmImprimir
        .FormulaSeleccion = cadFormula
        .OtrosParametros = cadParam
        .NumeroParametros = 2
        .SoloImprimir = False
        .Titulo = cadTitulo
        .NombreRPT = cadNombreRPT
        .Opcion = 1
        .Contabilidad = vParamAplic.NumeroContaGas
        .EnvioEMail = False
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

' variables para el recalculo de iva y totales
    Dim i As Integer
    Dim Imptot(2)
    Dim Tipiva(2)
    Dim Impbas(2) As Currency
    Dim impiva(2) As Currency
    Dim PorIva(2) As Currency
    Dim TotFac As Currency
    Dim totimp As Currency



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
            V = AdoAux(NumTabMto).Recordset.Fields(1) 'el 2 es el nº de llinia
            ModoLineas = 0
            CargaGrid NumTabMto, True
'            SituarTab (NumTabMto)
            DataGridAux(NumTabMto).SetFocus
            AdoAux(NumTabMto).Recordset.Find (AdoAux(NumTabMto).Recordset.Fields(1).Name & " =" & V)
            
'            ' ### [Monica] 29/09/2006
'            ' añadido el tema de de recalculo de bases
'            RecalculoBasesIvaFactura AdoAux(0).Recordset, Imptot, Tipiva, Impbas, impiva, PorIva, TotFac, totimp

            
            '13/02/2007 iniacializo los txt
            For i = 0 To 2
                text1(7 + (4 * i)).Text = ""
                text1(6 + (4 * i)).Text = ""
                text1(8 + (4 * i)).Text = ""
                text1(9 + (4 * i)).Text = ""
            Next i
            
            '13/02/2007 he añadido las condiciones del for antes solo estaban las sentencias
            For i = 0 To 2
                 If Tipiva(i) <> 0 Then text1(7 + (4 * i)).Text = Tipiva(i)
                 If Impbas(i) <> 0 Then text1(6 + (4 * i)).Text = Impbas(i)
                 If PorIva(i) <> 0 Then text1(8 + (4 * i)).Text = PorIva(i)
                 If impiva(i) <> 0 Then text1(9 + (4 * i)).Text = impiva(i)
                 'TotFac = Impbas(i) + impiva(i)
            Next i
            text1(19).Text = totimp
            text1(18).Text = TotFac
'            If BLOQUEADesdeFormulario2(Me, Data1, 1) Then
'                ModificaImportes = True
'                BotonModificar
'                cmdAceptar_Click
'            End If

            LLamaLineas NumTabMto, 0
    Exit Sub
    
EEliminarLin:
    If Err.Number <> 0 Then MuestraError Err.Number, "Eliminar Linea", Err.Description
End Sub

