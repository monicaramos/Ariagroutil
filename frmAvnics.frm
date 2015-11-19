VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmAvnics 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "A.V.N.I.C.S."
   ClientHeight    =   7455
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11715
   Icon            =   "frmAvnics.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7455
   ScaleWidth      =   11715
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Height          =   765
      Index           =   0
      Left            =   240
      TabIndex        =   34
      Top             =   480
      Width           =   11295
      Begin VB.ComboBox Combo1 
         Height          =   315
         Index           =   0
         ItemData        =   "frmAvnics.frx":000C
         Left            =   9240
         List            =   "frmAvnics.frx":0016
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Tag             =   "Clave Alta|N|N|0|2|avnic|codialta|||"
         Top             =   240
         Width           =   1935
      End
      Begin VB.TextBox text1 
         Height          =   285
         Index           =   2
         Left            =   7080
         MaxLength       =   10
         TabIndex        =   2
         Tag             =   "F.Alta|F|N|||avnic|fechalta|dd/mm/yyyy||"
         Top             =   240
         Width           =   1200
      End
      Begin VB.TextBox text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   0
         Left            =   1440
         MaxLength       =   6
         TabIndex        =   0
         Tag             =   "Código Avnics|N|N|0|999999|avnic|codavnic|000000|S|"
         Top             =   240
         Width           =   855
      End
      Begin VB.TextBox text1 
         Height          =   285
         Index           =   1
         Left            =   3480
         MaxLength       =   40
         TabIndex        =   1
         Tag             =   "Ejercicio|N|N|0|9999|avnic|anoejerc|0000|S|"
         Top             =   240
         Width           =   660
      End
      Begin VB.Label Label2 
         Caption         =   "Clave Alta"
         Height          =   255
         Index           =   1
         Left            =   8400
         TabIndex        =   53
         Top             =   270
         Width           =   975
      End
      Begin VB.Label Label20 
         Caption         =   "F.Alta"
         Height          =   255
         Left            =   6000
         TabIndex        =   51
         Top             =   255
         Width           =   615
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   0
         Left            =   6720
         Picture         =   "frmAvnics.frx":002C
         ToolTipText     =   "Buscar fecha"
         Top             =   240
         Width           =   240
      End
      Begin VB.Label Label4 
         Caption         =   "Ejercicio"
         Height          =   255
         Left            =   2640
         TabIndex        =   36
         Top             =   240
         Width           =   660
      End
      Begin VB.Label Label1 
         Caption         =   "Código Avnic"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   35
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Frame Frame1 
      Height          =   555
      Index           =   0
      Left            =   240
      TabIndex        =   31
      Top             =   6840
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
         TabIndex        =   32
         Top             =   180
         Width           =   2655
      End
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   10410
      TabIndex        =   30
      Top             =   6960
      Width           =   1035
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   9120
      TabIndex        =   29
      Top             =   6960
      Width           =   1035
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5460
      Left            =   240
      TabIndex        =   33
      Top             =   1320
      Width           =   11295
      _ExtentX        =   19923
      _ExtentY        =   9631
      _Version        =   393216
      Style           =   1
      Tabs            =   1
      TabsPerRow      =   6
      TabHeight       =   520
      ShowFocusRect   =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Datos básicos"
      TabPicture(0)   =   "frmAvnics.frx":00B7
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label5"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label6(0)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label1(26)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label7"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label29"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "imgZoom(0)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "imgBuscar(0)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label8"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Label6(1)"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "text1(3)"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "text1(5)"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "text1(8)"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "text1(7)"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "text1(6)"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "FrameDatosAlta"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "FrameDatosContacto"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "text1(26)"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "text1(9)"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "text2(9)"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "text1(4)"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).ControlCount=   20
      Begin VB.TextBox text1 
         Height          =   285
         Index           =   4
         Left            =   1440
         MaxLength       =   40
         TabIndex        =   5
         Tag             =   "Nombre|T|N|||avnic|nombrper|||"
         Top             =   960
         Width           =   4155
      End
      Begin VB.TextBox text2 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   9
         Left            =   2700
         TabIndex        =   48
         Top             =   2400
         Width           =   2895
      End
      Begin VB.TextBox text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   9
         Left            =   1440
         MaxLength       =   10
         TabIndex        =   10
         Tag             =   "Cta.Contable|T|S|||avnic|codmacta|||"
         Top             =   2400
         Width           =   1215
      End
      Begin VB.TextBox text1 
         Height          =   1905
         Index           =   26
         Left            =   5880
         MaxLength       =   250
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   28
         Tag             =   "Observaciones|T|S|||avnic|observac|||"
         Top             =   3390
         Width           =   5295
      End
      Begin VB.Frame FrameDatosContacto 
         Caption         =   "Datos Segundo Titular"
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
         Height          =   2490
         Left            =   120
         TabIndex        =   44
         Top             =   2850
         Width           =   5655
         Begin VB.TextBox text1 
            Height          =   285
            Index           =   11
            Left            =   4275
            MaxLength       =   9
            TabIndex        =   12
            Tag             =   "NIF / CIF|T|S|||avnic|nifrepre|||"
            Top             =   360
            Width           =   1200
         End
         Begin VB.TextBox text1 
            Height          =   285
            Index           =   10
            Left            =   1320
            MaxLength       =   9
            TabIndex        =   11
            Tag             =   "NIF / CIF|T|S|||avnic|nifpers1|||"
            Top             =   360
            Width           =   1200
         End
         Begin VB.TextBox text1 
            Height          =   285
            Index           =   13
            Left            =   1320
            MaxLength       =   40
            TabIndex        =   14
            Tag             =   "Domicilio|T|S|||avnic|nomcall1|||"
            Top             =   1185
            Width           =   4155
         End
         Begin VB.TextBox text1 
            Height          =   285
            Index           =   16
            Left            =   1320
            MaxLength       =   35
            TabIndex        =   17
            Tag             =   "Provincia|T|S|||avnic|provinc1|||"
            Top             =   1890
            Width           =   4155
         End
         Begin VB.TextBox text1 
            Height          =   285
            Index           =   15
            Left            =   2100
            MaxLength       =   35
            TabIndex        =   16
            Tag             =   "Población|T|S|||avnic|poblaci1|||"
            Top             =   1545
            Width           =   3375
         End
         Begin VB.TextBox text1 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   14
            Left            =   1320
            MaxLength       =   6
            TabIndex        =   15
            Tag             =   "C.Postal|T|S|||avnic|codpost1|||"
            Top             =   1545
            Width           =   735
         End
         Begin VB.TextBox text1 
            Height          =   285
            Index           =   12
            Left            =   1320
            MaxLength       =   40
            TabIndex        =   13
            Tag             =   "Nombre|T|S|||avnic|nombper1|||"
            Top             =   795
            Width           =   4155
         End
         Begin VB.Label Label12 
            Caption         =   "NIF Representante"
            Height          =   255
            Left            =   2760
            TabIndex        =   63
            Top             =   360
            Width           =   1455
         End
         Begin VB.Label Label10 
            Caption         =   "NIF"
            Height          =   255
            Left            =   120
            TabIndex        =   62
            Top             =   360
            Width           =   735
         End
         Begin VB.Label Label6 
            Caption         =   "Domicilio"
            Height          =   255
            Index           =   3
            Left            =   120
            TabIndex        =   61
            Top             =   1200
            Width           =   735
         End
         Begin VB.Label Label1 
            Caption         =   "Población"
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   60
            Top             =   1560
            Width           =   735
         End
         Begin VB.Label Label9 
            Caption         =   "Provincia"
            Height          =   255
            Left            =   120
            TabIndex        =   59
            Top             =   1920
            Width           =   735
         End
         Begin VB.Label Label6 
            Caption         =   "Nombre"
            Height          =   255
            Index           =   2
            Left            =   120
            TabIndex        =   58
            Top             =   810
            Width           =   735
         End
      End
      Begin VB.Frame FrameDatosAlta 
         Caption         =   "Datos Financieros"
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
         Height          =   2475
         Left            =   5880
         TabIndex        =   41
         Top             =   420
         Width           =   5295
         Begin VB.TextBox text1 
            Height          =   285
            Index           =   27
            Left            =   1200
            MaxLength       =   4
            TabIndex        =   19
            Tag             =   "IBAN|T|S|||avnic|iban|||"
            Top             =   585
            Width           =   600
         End
         Begin VB.TextBox text1 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   25
            Left            =   1200
            MaxLength       =   6
            TabIndex        =   27
            Tag             =   "% Int.|N|N|0|999.99|avnic|porcinte|##0.00||"
            Top             =   2040
            Width           =   600
         End
         Begin VB.TextBox text1 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   24
            Left            =   1200
            MaxLength       =   12
            TabIndex        =   26
            Tag             =   "Importe Ret.|N|S|0|9999999.99|avnic|imporret|#,###,##0.00||"
            Top             =   1680
            Width           =   1320
         End
         Begin VB.TextBox text1 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   23
            Left            =   1200
            MaxLength       =   12
            TabIndex        =   25
            Tag             =   "Importe Per.|N|S|0|9999999.99|avnic|imporper|#,###,##0.00||"
            Top             =   1320
            Width           =   1320
         End
         Begin VB.TextBox text1 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   22
            Left            =   1200
            MaxLength       =   12
            TabIndex        =   24
            Tag             =   "Importe|N|N|0|9999999.99|avnic|importes|#,###,##0.00||"
            Top             =   960
            Width           =   1320
         End
         Begin VB.TextBox text1 
            Height          =   285
            Index           =   21
            Left            =   3840
            MaxLength       =   10
            TabIndex        =   23
            Tag             =   "Cuenta|T|N|||avnic|cuentaba|||"
            Top             =   585
            Width           =   1320
         End
         Begin VB.TextBox text1 
            Height          =   285
            Index           =   20
            Left            =   3300
            MaxLength       =   2
            TabIndex        =   22
            Tag             =   "D.C.|T|N|||avnic|digcontr|||"
            Top             =   585
            Width           =   480
         End
         Begin VB.TextBox text1 
            Height          =   285
            Index           =   19
            Left            =   2580
            MaxLength       =   4
            TabIndex        =   21
            Tag             =   "Sucursal|N|N|0|9999|avnic|codsucur|0000||"
            Top             =   585
            Width           =   600
         End
         Begin VB.TextBox text1 
            Height          =   285
            Index           =   18
            Left            =   1860
            MaxLength       =   4
            TabIndex        =   20
            Tag             =   "Banco|N|N|0|9999|avnic|codbanco|0000||"
            Top             =   585
            Width           =   600
         End
         Begin VB.TextBox text1 
            Height          =   285
            Index           =   17
            Left            =   1200
            MaxLength       =   10
            TabIndex        =   18
            Tag             =   "F.Vto.|F|N|||avnic|fechavto|dd/mm/yyyy||"
            Top             =   240
            Width           =   1200
         End
         Begin VB.Label Label15 
            Caption         =   "% Interes"
            Height          =   255
            Left            =   120
            TabIndex        =   57
            Top             =   2070
            Width           =   975
         End
         Begin VB.Label Label13 
            Caption         =   "Imp. Retenc."
            Height          =   255
            Left            =   120
            TabIndex        =   56
            Top             =   1710
            Width           =   975
         End
         Begin VB.Label Label11 
            Caption         =   "Imp. Percep."
            Height          =   255
            Left            =   120
            TabIndex        =   55
            Top             =   1350
            Width           =   975
         End
         Begin VB.Label Label3 
            Caption         =   "Imp. Avnic"
            Height          =   255
            Left            =   120
            TabIndex        =   54
            Top             =   990
            Width           =   975
         End
         Begin VB.Image imgFec 
            Height          =   240
            Index           =   15
            Left            =   840
            Picture         =   "frmAvnics.frx":00D3
            ToolTipText     =   "Buscar fecha"
            Top             =   240
            Width           =   240
         End
         Begin VB.Label Label17 
            Caption         =   "IBAN Avnic"
            Height          =   255
            Left            =   120
            TabIndex        =   50
            Top             =   615
            Width           =   975
         End
         Begin VB.Label Label21 
            Caption         =   "F.Vto."
            Height          =   255
            Left            =   120
            TabIndex        =   43
            Top             =   255
            Width           =   615
         End
      End
      Begin VB.TextBox text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   6
         Left            =   1440
         MaxLength       =   6
         TabIndex        =   7
         Tag             =   "C.Postal|T|N|||avnic|codposta|||"
         Top             =   1710
         Width           =   735
      End
      Begin VB.TextBox text1 
         Height          =   285
         Index           =   7
         Left            =   2220
         MaxLength       =   35
         TabIndex        =   8
         Tag             =   "Población|T|N|||avnic|poblacio|||"
         Top             =   1710
         Width           =   3375
      End
      Begin VB.TextBox text1 
         Height          =   285
         Index           =   8
         Left            =   1440
         MaxLength       =   35
         TabIndex        =   9
         Tag             =   "Provincia|T|N|||avnic|provinci|||"
         Top             =   2055
         Width           =   4155
      End
      Begin VB.TextBox text1 
         Height          =   285
         Index           =   5
         Left            =   1440
         MaxLength       =   40
         TabIndex        =   6
         Tag             =   "Domicilio|T|N|||avnic|nomcalle|||"
         Top             =   1350
         Width           =   4155
      End
      Begin VB.TextBox text1 
         Height          =   285
         Index           =   3
         Left            =   1440
         MaxLength       =   9
         TabIndex        =   4
         Tag             =   "NIF / CIF|T|N|||avnic|nifperso|||"
         Top             =   520
         Width           =   1200
      End
      Begin VB.Label Label6 
         Caption         =   "Nombre"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   52
         Top             =   975
         Width           =   735
      End
      Begin VB.Label Label8 
         Caption         =   "Cta.Conta."
         Height          =   255
         Left            =   240
         TabIndex        =   49
         Top             =   2400
         Width           =   735
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   0
         Left            =   1080
         ToolTipText     =   "Buscar Cta.Contable"
         Top             =   2400
         Width           =   240
      End
      Begin VB.Image imgZoom 
         Height          =   240
         Index           =   0
         Left            =   7080
         ToolTipText     =   "Zoom descripción"
         Top             =   3030
         Width           =   240
      End
      Begin VB.Label Label29 
         Caption         =   "Observaciones"
         Height          =   255
         Left            =   5880
         TabIndex        =   47
         Top             =   3075
         Width           =   1215
      End
      Begin VB.Label Label7 
         Caption         =   "Provincia"
         Height          =   255
         Left            =   240
         TabIndex        =   40
         Top             =   2085
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Población"
         Height          =   255
         Index           =   26
         Left            =   240
         TabIndex        =   39
         Top             =   1725
         Width           =   735
      End
      Begin VB.Label Label6 
         Caption         =   "Domicilio"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   38
         Top             =   1365
         Width           =   735
      End
      Begin VB.Label Label5 
         Caption         =   "NIF"
         Height          =   255
         Left            =   240
         TabIndex        =   37
         Top             =   525
         Width           =   735
      End
   End
   Begin MSAdodcLib.Adodc Data1 
      Height          =   330
      Left            =   4200
      Top             =   6960
      Visible         =   0   'False
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   582
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
      TabIndex        =   45
      Top             =   0
      Width           =   11715
      _ExtentX        =   20664
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
            Object.Visible         =   0   'False
            Style           =   3
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.Visible         =   0   'False
            Object.ToolTipText     =   "Buscar Tarjeta"
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
         Left            =   8520
         TabIndex        =   46
         Top             =   120
         Width           =   1215
      End
   End
   Begin VB.CommandButton cmdRegresar 
      Caption         =   "&Regresar"
      Height          =   375
      Left            =   10440
      TabIndex        =   42
      Top             =   6960
      Visible         =   0   'False
      Width           =   1035
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
      Begin VB.Menu mnBuscarTarjeta 
         Caption         =   "Buscar &Tarjeta"
         Shortcut        =   ^T
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
Attribute VB_Name = "frmAvnics"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' +-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+
' +-+- Autor: MANOLO                   -+-+
' +-+- Menú: AVNICS                    -+-+
' +-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+

' +-+-+-+- DISSENY +-+-+-+-
' 1. Posar tots els controls al formulari
' 2. Posar els index correlativament
' 3. Si n'hi han botons de buscar repasar el ToolTipText
' 4. Alliniar els camps numérics a la dreta i el resto a l'esquerra
' 5. Posar els TAGs
' (si es INTEGER: si PK => mínim 1; si no PK => mínim 0; màxim => 99; format => 00)
' (si es DECIMAL; mínim => 0; màxim => 99.99; format => #,###,###,##0.00)
' (si es DATE; format => dd/mm/yyyy)
' 6. Posar els MAXLENGTHs
' 7. Posar els TABINDEXs

Option Explicit

'Dim T1 As Single

Public DatosADevolverBusqueda As String    'Tindrà el nº de text que vol que torne, empipat
Public Event DatoSeleccionado(CadenaSeleccion As String)
Public NuevoCodigo As String
Public CodigoActual As String
Public DeConsulta As Boolean

' *** declarar els formularis als que vaig a cridar ***
Private WithEvents frmB As frmBuscaGrid
Attribute frmB.VB_VarHelpID = -1
Private WithEvents frmC As frmCal 'calendario fecha
Attribute frmC.VB_VarHelpID = -1
Private WithEvents frmZ As frmZoom  'Zoom para campos Text
Attribute frmZ.VB_VarHelpID = -1
Private WithEvents frmCtas As frmCtasConta 'cuentas contables
Attribute frmCtas.VB_VarHelpID = -1
' *****************************************************


Private Modo As Byte
'*************** MODOS ********************
'   0.-  Formulari net sense cap camp ple
'   1.-  Preparant per a fer la búsqueda
'   2.-  Ja tenim registres i els anem a recorrer
'        ,podem editar-los Edició del camp
'   3.-  Inserció de nou registre
'   4.-  Modificar
'   5.-  Manteniment Llinies

'+-+-Variables comuns a tots els formularis+-+-+

Dim ModoLineas As Byte
'1.- Afegir,  2.- Modificar,  3.- Borrar,  0.-Passar control a Llínies

Dim NumTabMto As Integer 'Indica quin nº de Tab està en modo Mantenimient
Dim TituloLinea As String 'Descripció de la llínia que està en Mantenimient
Dim PrimeraVez As Boolean

Private CadenaConsulta As String 'SQL de la taula principal del formulari
Private Ordenacion As String
Private NombreTabla As String  'Nom de la taula

Private kCampo As Integer
'-------------------------------------------------------------------------
Private HaDevueltoDatos As Boolean

'Private VieneDeBuscar As Boolean
'Per a quan torna 2 poblacions en el mateix codi Postal. Si ve de pulsar prismatic
'de búsqueda posar el valor de població seleccionada i no tornar a recuperar de la Base de Datos

Dim btnPrimero As Byte 'Variable que indica el nº del Botó PrimerRegistro en la Toolbar1
'Dim CadAncho() As Boolean  'array, per a quan cridem al form de llínies
Dim indice As Byte 'Index del text1 on es poses els datos retornats des d'atres Formularis de Mtos
Dim CadB As String

Dim BuscaChekc As String


Private Sub cmbAux_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub cmdAceptar_Click()
    Screen.MousePointer = vbHourglass
    
    On Error GoTo Error1
    
    Select Case Modo
        Case 1  'BÚSQUEDA
            HacerBusqueda
            
        Case 3 'INSERTAR
            If DatosOk Then
                If InsertarDesdeForm2(Me, 1) Then
                    Data1.RecordSource = "Select * from " & NombreTabla & Ordenacion
                    PosicionarData
                End If
            Else
                ModoLineas = 0
            End If
            
        Case 4  'MODIFICAR
            If DatosOk Then
                If ModificaDesdeFormulario2(Me, 1) Then
                    TerminaBloquear
                    PosicionarData
                End If
            Else
                ModoLineas = 0
            End If
    End Select
    Screen.MousePointer = vbDefault
 
Error1:
    Screen.MousePointer = vbDefault
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Sub

Private Sub Combo1_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
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

    PrimeraVez = True
    
    ' ICONETS DE LA BARRA
    btnPrimero = 16 'index del botó "primero"
    With Me.Toolbar1
        .HotImageList = frmPpal.imgListComun_OM
        .DisabledImageList = frmPpal.imgListComun_BN
        .ImageList = frmPpal.imgListComun
        'l'1 i el 2 son separadors
        .Buttons(3).Image = 1   'Buscar
        .Buttons(4).Image = 2   'Totss
        'el 5 i el 6 son separadors
        .Buttons(7).Image = 3   'Insertar
        .Buttons(8).Image = 4   'Modificar
        .Buttons(9).Image = 5   'Borrar
        .Buttons(11).Image = 21   'Borrar
        'el 10 i el 11 son separadors
        .Buttons(12).Image = 10  'Imprimir
        .Buttons(13).Image = 11  'Eixir
        'el 13 i el 14 son separadors
        .Buttons(btnPrimero).Image = 6  'Primer
        .Buttons(btnPrimero + 1).Image = 7 'Anterior
        .Buttons(btnPrimero + 2).Image = 8 'Següent
        .Buttons(btnPrimero + 3).Image = 9 'Últim
    End With
    
    'cargar IMAGES de busqueda
    For i = 0 To Me.imgBuscar.Count - 1
        Me.imgBuscar(i).Picture = frmPpal.imgListImages16.ListImages(1).Picture
    Next i
    
    'IMAGES para zoom
    For i = 0 To Me.imgZoom.Count - 1
        Me.imgZoom(i).Picture = frmPpal.imgListImages16.ListImages(3).Picture
    Next i
    
    ' *** si n'hi han tabs, per a que per defecte sempre es pose al 1r***
    Me.SSTab1.Tab = 0
    ' *******************************************************************
    
    LimpiarCampos   'Neteja els camps TextBox
    
    '*** canviar el nom de la taula i l'ordenació de la capçalera ***
    NombreTabla = "avnic"
    Ordenacion = " ORDER BY codavnic"
    
    'Mirem com està guardat el valor del check
    chkVistaPrevia(0).Value = CheckValueLeer(Name)
    
    Data1.ConnectionString = conn
    '***** cambiar el nombre de la PK de la cabecera *************
    Data1.RecordSource = "Select * from " & NombreTabla & " where codavnic=-1"
    Data1.Refresh
       
    ModoLineas = 0
       
    ' *** si n'hi han combos (capçalera o llínies) ***
    CargaCombo 0
    
    If DatosADevolverBusqueda = "" Then
        PonerModo 0
    Else
        PonerModo 1 'búsqueda
        ' *** posar de groc els camps visibles de la clau primaria de la capçalera ***
        text1(0).BackColor = vbYellow 'codclien
        ' ****************************************************************************
    End If
End Sub

Private Sub LimpiarCampos()
    On Error Resume Next
    
    Limpiar Me   'Mètode general: Neteja els controls TextBox
    lblIndicador.Caption = ""
    
    ' *** si n'hi han combos a la capçalera ***
    Me.Combo1(0).ListIndex = -1

    If Err.Number <> 0 Then Err.Clear
End Sub

Private Sub LimpiarCamposLin(frameAux As String)
    On Error Resume Next
    
    LimpiarLin Me, frameAux  'Mètode general: Neteja els controls TextBox
    lblIndicador.Caption = ""

    If Err.Number <> 0 Then Err.Clear
End Sub

'----------------------------------------------------------------
'----------------------------------------------------------------
'   En PONERMODO s'habiliten, o no, els diversos camps del
'   formulari en funció del modo en que anem a treballar
Private Sub PonerModo(Kmodo As Byte, Optional indFrame As Integer)
Dim i As Integer, NumReg As Byte
Dim b As Boolean

    On Error GoTo EPonerModo
 
    'Actualisa Iconos Insertar,Modificar,Eliminar
    'ActualizarToolbar Modo, Kmodo
    Modo = Kmodo
    PonerIndicador lblIndicador, Modo, ModoLineas
       
    'Modo 2. N'hi han datos i estam visualisant-los
    '=========================================
    'Posem visible, si es formulari de búsqueda, el botó "Regresar" quan n'hi han datos
    If DatosADevolverBusqueda <> "" Then
        cmdRegresar.visible = (Modo = 2)
    Else
        cmdRegresar.visible = False
    End If
    
    '=======================================
    b = (Modo = 2)
    'Posar Fleches de desplasament visibles
    NumReg = 1
    If Not Data1.Recordset.EOF Then
        If Data1.Recordset.RecordCount > 1 Then NumReg = 2 'Només es per a saber que n'hi ha + d'1 registre
    End If
    DesplazamientoVisible Me.Toolbar1, btnPrimero, b, NumReg
    
    '---------------------------------------------
    b = Modo <> 0 And Modo <> 2
    cmdCancelar.visible = b
    cmdAceptar.visible = b
       
    'Bloqueja els camps Text1 si no estem modificant/Insertant Datos
    'Si estem en Insertar a més neteja els camps Text1
    BloquearText1 Me, Modo
    
    '*** si n'hi han combos a la capçalera ***
    BloquearCombo Me, Modo
    '**************************
    
    ' **** si n'hi han imagens de buscar en la capçalera *****
    BloquearImgBuscar Me, Modo, ModoLineas
    BloquearImgZoom Me, Modo, ModoLineas
    BloquearImgFec Me, 0, Modo
    BloquearImgFec Me, 15, Modo
    ' ********************************************************
    
    chkVistaPrevia(0).Enabled = (Modo <= 2)
    
    PonerLongCampos

'    If (Modo < 2) Or (Modo = 3) Then
'        CargaGrid 0, False
'        CargaGrid 1, False
'    End If
'
'    b = (Modo = 4) Or (Modo = 2)
'    DataGridAux(0).Enabled = b
'    DataGridAux(1).Enabled = b
      
    ' ****** si n'hi han combos a la capçalera ***********************
     If (Modo = 0) Or (Modo = 2) Or (Modo = 5) Then
        Combo1(0).Enabled = False
        Combo1(0).BackColor = &H80000018 'groc
    ElseIf (Modo = 1) Or (Modo = 3) Or (Modo = 4) Then
        Combo1(0).Enabled = True
        Combo1(0).BackColor = &H80000005 'blanc
    End If
    ' ****************************************************************
    
    PonerModoOpcionesMenu (Modo) 'Activar opcions menú según modo
    PonerOpcionesMenu   'Activar opcions de menú según nivell
                        'de permisos de l'usuari

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
    PonerOpcionesMenuGeneral Me
    PonerOpcionesMenuGeneralNew Me
End Sub

Private Sub PonerModoOpcionesMenu(Modo)
'Actives unes Opcions de Menú i Toolbar según el modo en que estem
Dim b As Boolean, bAux As Boolean
Dim i As Byte
    
    'Barra de CAPÇALERA
    '------------------------------------------
    'b = (Modo = 2 Or Modo = 0 Or Modo = 1)
    b = (Modo = 2 Or Modo = 0)
    'Buscar
    Toolbar1.Buttons(3).Enabled = b
    Me.mnBuscar.Enabled = b
    'Vore Tots
    Toolbar1.Buttons(4).Enabled = b
    Me.mnVerTodos.Enabled = b
    
    'Insertar
    Toolbar1.Buttons(7).Enabled = b And Not DeConsulta
    Me.mnNuevo.Enabled = b And Not DeConsulta
    
    b = (Modo = 2 And Data1.Recordset.RecordCount > 0) And Not DeConsulta
    'Modificar
    Toolbar1.Buttons(8).Enabled = b
    Me.mnModificar.Enabled = b
    'eliminar
    Toolbar1.Buttons(9).Enabled = b
    Me.mnEliminar.Enabled = b
    
    'Imprimir
    'Toolbar1.Buttons(12).Enabled = (b Or Modo = 0)
    Toolbar1.Buttons(12).Enabled = True And Not DeConsulta
       
    ' *** si n'hi han llínies que tenen grids (en o sense tab) ***
'    b = (Modo = 3 Or Modo = 4 Or Modo = 2) And Not DeConsulta
'    For i = 0 To ToolAux.Count - 1
'        ToolAux(i).Buttons(1).Enabled = b
'        If b Then bAux = (b And Me.AdoAux(i).Recordset.RecordCount > 0)
'        ToolAux(i).Buttons(2).Enabled = bAux
'        ToolAux(i).Buttons(3).Enabled = bAux
'    Next i
    
End Sub

Private Sub Desplazamiento(Index As Integer)
'Botons de Desplaçament; per a desplaçar-se pels registres de control Data
    If Data1.Recordset.EOF Then Exit Sub
    DesplazamientoData Data1, Index
    PonerCampos
End Sub

Private Sub frmB_Selecionado(CadenaDevuelta As String)
    Dim Aux As String
    
    If CadenaDevuelta <> "" Then
        HaDevueltoDatos = True
        Screen.MousePointer = vbHourglass
        'Sabem quins camps son els que mos torna
        'Creem una cadena consulta i posem els datos
        CadB = ""
        Aux = ValorDevueltoFormGrid(text1(0), CadenaDevuelta, 1)
        CadB = Aux
        '   Com la clau principal es única, en posar el sql apuntant
        '   al valor retornat sobre la clau ppal es suficient
        CadenaConsulta = "select * from " & NombreTabla & " WHERE " & CadB & " " & Ordenacion
        ' **********************************
        PonerCadenaBusqueda
        Screen.MousePointer = vbDefault
    End If
End Sub

Private Sub frmTra_Actualizar(vValor As Integer)
'Mantenimiento de Colectivos
    
    LimpiarCampos
    text1(0).Text = vValor
    
    FormateaCampo text1(0)
        Modo = 1
        cmdAceptar_Click
End Sub

Private Sub frmZ_Actualizar(vCampo As String)
     text1(indice).Text = vCampo
End Sub

' *** si n'hi ha buscar data, posar a les <=== el menor index de les imagens de buscar data ***
' NOTA: ha de coincidir l'index de la image en el del camp a on va a parar el valor
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

    frmC.Left = esq + imgFec(Index).Parent.Left + 30
    frmC.Top = dalt + imgFec(Index).Parent.Top + imgFec(Index).Height + menu - 40

    imgFec(15).Tag = Index '<===
    ' *** repasar si el camp es txtAux o Text1 ***
    If text1(Index + 2).Text <> "" Then frmC.NovaData = text1(Index + 2).Text
    ' ********************************************

    frmC.Show vbModal
    Set frmC = Nothing
    ' *** repasar si el camp es txtAux o Text1 ***
    PonerFoco text1(CByte(imgFec(15).Tag) + 2) '<===
    ' ********************************************
End Sub

Private Sub frmC_Selec(vFecha As Date)
    ' *** repasar si el camp es txtAux o Text1 ***
    text1(CByte(imgFec(15).Tag) + 2).Text = Format(vFecha, "dd/mm/yyyy") '<===
End Sub

Private Sub imgZoom_Click(Index As Integer)
    
    Set frmZ = New frmZoom

    If Index = 0 Then
        indice = 26
        frmZ.pTitulo = "Observaciones del Avnic."
        frmZ.pValor = text1(indice).Text
        frmZ.pModo = Modo
    
        frmZ.Show vbModal
        Set frmZ = Nothing
            
        PonerFoco text1(indice)
    End If
End Sub

Private Sub mnBuscar_Click()
    BotonBuscar
    Combo1(0).ListIndex = -1 'quan busque, per defecte no seleccione res.
End Sub

Private Sub mnEliminar_Click()
    BotonEliminar
End Sub

Private Sub mnImprimir_Click()
'    AbrirListado (10)
End Sub

Private Sub mnModificar_Click()
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

    Select Case Button.Index
        Case 3  'Búscar
            mnBuscar_Click
        Case 4  'Tots
            mnVerTodos_Click
        Case 7  'Nou
            mnNuevo_Click
        Case 8  'Modificar
            mnModificar_Click
        Case 9  'Borrar
            mnEliminar_Click
        Case 12 'Imprimir
            mnImprimir_Click
        Case 13    'Eixir
            mnSalir_Click
            
        Case btnPrimero To btnPrimero + 3 'Fleches Desplaçament
            Desplazamiento (Button.Index - btnPrimero)
    End Select
End Sub

Private Sub BotonBuscar()
Dim i As Integer
' ***** Si la clau primaria de la capçalera no es Text1(0), canviar-ho en <=== *****
    If Modo <> 1 Then
        LimpiarCampos
        PonerModo 1
        PonerFoco text1(0) ' <===
        text1(0).BackColor = vbYellow ' <===
        ' *** si n'hi han combos a la capçalera ***
        For i = 0 To Combo1.Count - 1
            Combo1(i).ListIndex = -1
        Next i
    Else
        HacerBusqueda
        If Data1.Recordset.EOF Then
            text1(kCampo).Text = ""
            text1(kCampo).BackColor = vbYellow
            PonerFoco text1(kCampo)
        End If
    End If
' ******************************************************************************
End Sub

Private Sub HacerBusqueda()

    CadB = ObtenerBusqueda2(Me, 1)
    
    If chkVistaPrevia(0) = 1 Then
        MandaBusquedaPrevia CadB
    ElseIf CadB <> "" Then
        CadenaConsulta = "select * from " & NombreTabla & " WHERE " & CadB & " " & Ordenacion
        PonerCadenaBusqueda
    Else
        ' *** foco al 1r camp visible de la capçalera que siga clau primaria ***
        PonerFoco text1(0)
        ' **********************************************************************
    End If
End Sub

Private Sub MandaBusquedaPrevia(CadB As String)
    Dim cad As String
        
    'Cridem al form
    ' **************** arreglar-ho per a vore lo que es desije ****************
    ' NOTA: el total d'amples de ParaGrid, ha de sumar 100
    cad = ""
    cad = cad & ParaGrid(text1(0), 15, "Cód.")
    cad = cad & ParaGrid(text1(3), 25, "N.I.F.")
    cad = cad & ParaGrid(text1(4), 60, "Nombre")
    If cad <> "" Then
        Screen.MousePointer = vbHourglass
        Set frmB = New frmBuscaGrid
        frmB.vCampos = cad
        frmB.vTabla = NombreTabla
        frmB.vSQL = CadB
        HaDevueltoDatos = False
        frmB.vDevuelve = "0|1|2|" '*** els camps que volen que torne ***
        frmB.vTitulo = "Avnics" ' ***** repasa açò: títol de BuscaGrid *****
        frmB.vSelElem = 1

        frmB.Show vbModal
        Set frmB = Nothing
        'Si ha posat valors i tenim que es formulari de búsqueda llavors
        'tindrem que tancar el form llançant l'event
        If HaDevueltoDatos Then
            If (Not Data1.Recordset.EOF) And DatosADevolverBusqueda <> "" Then _
                cmdRegresar_Click
        Else   'de ha retornat datos, es a decir NO ha retornat datos
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
        MsgBox "No hay ningún registro en la tabla " & NombreTabla, vbInformation
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
'Vore tots
    LimpiarCampos 'Neteja els Text1
    CadB = ""
    
    If chkVistaPrevia(0).Value = 1 Then
        MandaBusquedaPrevia ""
    Else
        CadenaConsulta = "Select * from " & NombreTabla & Ordenacion
        PonerCadenaBusqueda
    End If
End Sub

Private Sub BotonAnyadir()

    LimpiarCampos 'Huida els TextBox
    PonerModo 3
    
    ' ****** Valors per defecte a l'afegir, repasar si n'hi ha
    ' codEmpre i quins camps tenen la PK de la capçalera *******
    text1(0).Text = SugerirCodigoSiguienteStr("avnic", "codavnic")
    FormateaCampo text1(0)
    
    text1(1).Text = Format(Now, "yyyy") '
    text1(2).Text = Format(Now, "dd/mm/yyyy") ' Quan afegixc pose en F.Alta i F.Modificación la data actual
    PosicionarCombo Combo1(0), 1
        
    PonerFoco text1(0) '*** 1r camp visible que siga PK ***
    
    ' *** si n'hi han camps de descripció a la capçalera ***
    'PosarDescripcions

    ' *** si n'hi han tabs, em posicione al 1r ***
    Me.SSTab1.Tab = 0
End Sub

Private Sub BotonModificar()

    PonerModo 4

    ' *** bloquejar els camps visibles de la clau primaria de la capçalera ***
    BloquearTxt text1(0), True
    BloquearTxt text1(1), True
    
    ' *** foco al 1r camp visible que NO siga clau primaria ***
    PonerFoco text1(2)
End Sub

Private Sub BotonEliminar()
Dim cad As String

    On Error GoTo EEliminar

    'Ciertas comprobaciones
    If Data1.Recordset.EOF Then Exit Sub

    ' *************** canviar la pregunta ****************
    cad = "¿Seguro que desea eliminar el Avnics?"
    cad = cad & vbCrLf & "Código: " & Format(Data1.Recordset.Fields(0), FormatoCampo(text1(0)))
    cad = cad & vbCrLf & "Nombre: " & Data1.Recordset.Fields(1)
    
    If MsgBox(cad, vbQuestion + vbYesNo) = vbYes Then
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
            PonerModo 0
        End If
    End If
    Screen.MousePointer = vbDefault
    
EEliminar:
    Screen.MousePointer = vbDefault
    If Err.Number > 0 Then MuestraError Err.Number, "Eliminar Avnics", Err.Description
End Sub

Private Sub PonerCampos()
Dim i As Integer
Dim codPobla As String, desPobla As String
Dim CPostal As String, desProvi As String, desPais As String

    If Data1.Recordset.EOF Then Exit Sub
    PonerCamposForma2 Me, Data1, 1 'opcio=1: posa el format o els camps de la capçalera
    
    ' ************* configurar els camps de les descripcions de la capçalera *************
    Text2(9).Text = PonerNombreCuenta(text1(9), Modo, , 2)
    ' ********************************************************************************
    
    '-- Esto permanece para saber donde estamos
    lblIndicador.Caption = Data1.Recordset.AbsolutePosition & " de " & Data1.Recordset.RecordCount
    
'    PonerModoOpcionesMenu (Modo)
'    PonerOpcionesMenu
End Sub

Private Sub cmdCancelar_Click()
Dim i As Integer
Dim V

    Select Case Modo
        Case 1, 3 'Búsqueda, Insertar
                LimpiarCampos
                If Data1.Recordset.EOF Then
                    PonerModo 0
                Else
                    PonerModo 2
                    PonerCampos
                End If
                ' *** foco al primer camp visible de la capçalera ***
                PonerFoco text1(0)

        Case 4  'Modificar
                TerminaBloquear
                PonerModo 2
                PonerCampos
                ' *** primer camp visible de la capçalera ***
                PonerFoco text1(0)
        
'        Case 5 'LLÍNIES
'            Select Case ModoLineas
'                Case 1 'afegir llínia
'                    ModoLineas = 0
'                    ' *** les llínies que tenen datagrid (en o sense tab) ***
'                    If NumTabMto = 0 Or NumTabMto = 1 Or NumTabMto = 2 Or NumTabMto = 4 Then
'                        DataGridAux(NumTabMto).AllowAddNew = False
'                        ' **** repasar si es diu Data1 l'adodc de la capçalera ***
'                        'If BLOQUEADesdeFormulario2(Me, Data1, 1) Then BotonModificar 'Modificar
'                        LLamaLineas NumTabMto, ModoLineas 'ocultar txtAux
'                        DataGridAux(NumTabMto).Enabled = True
'                        DataGridAux(NumTabMto).SetFocus
'
'                        ' *** si n'hi han camps de descripció dins del grid, els neteje ***
'                        'txtAux2(2).text = ""
'
'                    End If
'
'                    ' *** si n'hi han tabs ***
'                    SituarTab (NumTabMto + 1)
'
'                    If Not AdoAux(NumTabMto).Recordset.EOF Then
'                        AdoAux(NumTabMto).Recordset.MoveFirst
'                    End If
'
'                Case 2 'modificar llínies
'                    ModoLineas = 0
'
'                    ' *** si n'hi han tabs ***
'                    SituarTab (NumTabMto + 1)
'                    LLamaLineas NumTabMto, ModoLineas 'ocultar txtAux
'                    PonerModo 4
'                    If Not AdoAux(NumTabMto).Recordset.EOF Then
'                        ' *** l'Index de Fields es el que canvie de la PK de llínies ***
'                        V = AdoAux(NumTabMto).Recordset.Fields(1) 'el 2 es el nº de llinia
'                        AdoAux(NumTabMto).Recordset.Find (AdoAux(NumTabMto).Recordset.Fields(1).Name & " =" & V)
'                        ' ***************************************************************
'                    End If

'           End Select
            
            PosicionarData
           
    End Select
End Sub

Private Function DatosOk() As Boolean
Dim b As Boolean
'Dim Datos As String
Dim cta As String
Dim CadMen As String


    On Error GoTo EDatosOK

    DatosOk = False
    b = CompForm2(Me, 1)
    If Not b Then Exit Function
    
    ' *** canviar els arguments de la funcio, el mensage i repasar si n'hi ha codEmpre ***
    If (Modo = 3) Then 'insertar
        'comprobar si existe ya el cod. del campo clave primaria
        If ExisteCP(text1(0)) Then b = False
    End If
    
    
        '[Monica]22/08/2013: añadida la comprobacion de que la cuenta contable sea correcta
        If text1(18).Text = "" Or text1(19).Text = "" Or text1(21).Text = "" Or text1(21).Text = "" Then
            '[Monica]20/11/2013: añadido el codigo de iban
            text1(27).Text = ""
            text1(18).Text = ""
            text1(19).Text = ""
            text1(20).Text = ""
            text1(21).Text = ""
        Else
            cta = Format(text1(18).Text, "0000") & Format(text1(19).Text, "0000") & Format(text1(20).Text, "00") & Format(text1(21).Text, "0000000000")
            If Val(ComprobarCero(cta)) = 0 Then
                CadMen = "El avnic no tiene asignada cuenta bancaria."
                MsgBox CadMen, vbExclamation
            End If
            If Not Comprueba_CC(cta) Then
                CadMen = "La cuenta bancaria del avnic no es correcta. ¿ Desea continuar ?."
                If MsgBox(CadMen, vbQuestion + vbYesNo + vbDefaultButton1) = vbYes Then
                    b = True
                Else
                    PonerFoco text1(19)
                    b = False
                End If
            Else
'                '[Monica]20/11/2013: añadimos el tema de la comprobacion del IBAN
'                If Not Comprueba_CC_IBAN(cta, Text1(42).Text) Then
'                    cadMen = "La cuenta IBAN del cliente no es correcta. ¿ Desea continuar ?."
'                    If MsgBox(cadMen, vbQuestion + vbYesNo + vbDefaultButton1) = vbYes Then
'                        b = True
'                    Else
'                        PonerFoco Text1(42)
'                        b = False
'                    End If
'                End If

'       sustituido por lo de David
                BuscaChekc = ""
                If Me.text1(27).Text <> "" Then BuscaChekc = Mid(text1(27).Text, 1, 2)
                    
                If DevuelveIBAN2(BuscaChekc, cta, cta) Then
                    If Me.text1(27).Text = "" Then
                        If MsgBox("Poner IBAN ?", vbQuestion + vbYesNo) = vbYes Then Me.text1(27).Text = BuscaChekc & cta
                    Else
                        If Mid(text1(27).Text, 3) <> cta Then
                            cta = "Calculado : " & BuscaChekc & cta
                            cta = "Introducido: " & Me.text1(27).Text & vbCrLf & cta & vbCrLf
                            cta = "Error en codigo IBAN" & vbCrLf & cta & "Continuar?"
                            If MsgBox(cta, vbQuestion + vbYesNo) = vbNo Then
                                PonerFoco text1(27)
                                b = False
                            End If
                        End If
                    End If
                End If
            End If
        End If
        
    
    
    ' ************************************************************************************
    DatosOk = b
    
EDatosOK:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Function

Private Sub PosicionarData()
Dim cad As String, Indicador As String

    ' *** canviar-ho per tota la PK de la capçalera, no llevar els () ***
    cad = "(codavnic=" & text1(0).Text & ")"
    
    ' *** gastar SituarData o SituarDataMULTI depenent de si la PK es simple o composta ***
    'If SituarDataMULTI(Data1, cad, Indicador) Then
    If SituarData(Data1, cad, Indicador) Then
        If ModoLineas <> 1 Then PonerModo 2
        lblIndicador.Caption = Indicador
    Else
       LimpiarCampos
       PonerModo 0
    End If
End Sub

Private Function eliminar() As Boolean
Dim vWhere As String

    On Error GoTo FinEliminar

    conn.BeginTrans
    ' ***** canviar el nom de la PK de la capçalera, repasar codEmpre *******
    vWhere = " WHERE codavnic=" & Data1.Recordset!codavnic
        
    'Eliminar la CAPÇALERA
    vWhere = " WHERE codavnic=" & Data1.Recordset!codavnic & " and anoejerc=" & Data1.Recordset!anoejerc
    conn.Execute "Delete from " & NombreTabla & vWhere
       
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

' *** si n'hi han combos a la capçalera ***
Private Sub Combo1_GotFocus(Index As Integer)
    If Modo = 1 Then Combo1(Index).BackColor = vbYellow
End Sub

Private Sub Combo1_LostFocus(Index As Integer)
    If Combo1(Index).BackColor = vbYellow Then Combo1(Index).BackColor = vbWhite
End Sub

Private Sub Text1_LostFocus(Index As Integer)
Dim CadMen As String
Dim Nuevo As Boolean

    If Not PerderFocoGnral(text1(Index), Modo) Then Exit Sub
    
    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub
    
    ' ***************** configurar els LostFocus dels camps de la capçalera *****************
    Select Case Index
        Case 0 'Cod.Avnic
            PonerFormatoEntero text1(0)

        Case 4, 12 'NOMBRE
            text1(Index).Text = UCase(text1(Index).Text)
        
        Case 3, 10, 11 'NIF
            text1(Index).Text = UCase(text1(Index).Text)
            ValidarNIF text1(Index).Text
                
        Case 2, 17 'Fechas
            PonerFormatoFecha text1(Index)
            
        Case 9 'cuenta contable
            If text1(Index).Text = "" Then Exit Sub
            Text2(Index).Text = PonerNombreCuenta(text1(Index), Modo, text1(0).Text, 2)
            
        Case 22, 23, 24 'IMPORTES
            CadMen = TransformaPuntosComas(text1(Index).Text)
            text1(Index).Text = Format(CadMen, "#,###,##0.00")
        
        Case 25 '% INTERES
            CadMen = TransformaPuntosComas(text1(Index).Text)
            text1(Index).Text = Format(CadMen, "##0.00")
            
        Case 27 ' codigo de iban
            text1(Index).Text = UCase(text1(Index).Text)
            
    End Select
    
    '[Monica]: calculo del iban si no lo ponen
    If Index = 18 Or Index = 19 Or Index = 20 Or Index = 21 Then
        Dim cta As String
        Dim CC As String
        If text1(18).Text <> "" And text1(19).Text <> "" And text1(20).Text <> "" And text1(21).Text <> "" Then
            
            cta = Format(text1(18).Text, "0000") & Format(text1(19).Text, "0000") & Format(text1(20).Text, "00") & Format(text1(21).Text, "0000000000")
            If Len(cta) = 20 Then
    '        Text1(42).Text = Calculo_CC_IBAN(cta, Text1(42).Text)
    
                If text1(27).Text = "" Then
                    'NO ha puesto IBAN
                    If DevuelveIBAN2("ES", cta, cta) Then text1(27).Text = "ES" & cta
                Else
                    CC = CStr(Mid(text1(27).Text, 1, 2))
                    If DevuelveIBAN2(CStr(CC), cta, cta) Then
                        If Mid(text1(27).Text, 3) <> cta Then
                            
                            MsgBox "Codigo IBAN distinto del calculado [" & CC & cta & "]", vbExclamation
                        End If
                    End If
                End If
                
                
            End If
        End If
    End If
    ' ***************************************************************************
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = teclaBuscar Then
        If Modo = 1 Or Modo = 3 Or Modo = 4 Then
            Select Case Index
                Case 2: KEYFecha KeyAscii, 15 'fecha de alta
                Case 9: KEYBusqueda KeyAscii, 2 'cuenta contable
                Case 17: KEYFecha KeyAscii, 16 'fecha de vencimiento
            End Select
        End If
    Else
        KEYpress KeyAscii
    End If
End Sub

Private Sub Text1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
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

Private Sub KEYBusqueda(KeyAscii As Integer, indice As Integer)
    KeyAscii = 0
    imgBuscar_Click (indice)
End Sub

Private Sub KEYFecha(KeyAscii As Integer, indice As Integer)
    KeyAscii = 0
    imgFec_Click (indice)
End Sub

' ********* si n'hi han combos a la capçalera ************
Private Sub CargaCombo(Index As Integer)
Dim Ini As Integer
Dim Fin As Integer
Dim i As Integer

    ' *** neteje els combos, els pose valor i seleccione el valor per defecte ***
    Combo1(0).Clear
    
    Combo1(0).AddItem "Antigua"
    Combo1(0).ItemData(Combo1(0).NewIndex) = 0
    Combo1(0).AddItem "Alta Ejercicio"
    Combo1(0).ItemData(Combo1(0).NewIndex) = 1
    Combo1(0).AddItem "Cancelada Ejercicio"
    Combo1(0).ItemData(Combo1(0).NewIndex) = 2
End Sub

Private Function SepuedeBorrar(ByRef Index As Integer) As Boolean
    SepuedeBorrar = False
    
    ' *** si cal comprovar alguna cosa abans de borrar ***
'    Select Case Index
'        Case 0 'cuentas bancarias
'            If AdoAux(Index).Recordset!ctaprpal = 1 Then
'                MsgBox "No puede borrar una Cuenta Principal. Seleccione antes otra cuenta como Principal", vbExclamation
'                Exit Function
'            End If
'    End Select
    ' ****************************************************
    
    SepuedeBorrar = True
End Function

Private Sub imgBuscar_Click(Index As Integer)
    TerminaBloquear
     Select Case Index
        Case 0 'Cuentas Contables (de contabilidad)
            If vParamAplic.NumeroConta = 0 Then Exit Sub
            
            indice = Index + 9
            Set frmCtas = New frmCtasConta
            frmCtas.NumDigit = 0
            frmCtas.DatosADevolverBusqueda = "0|1|"
            frmCtas.CodigoActual = Text2(indice).Text
            frmCtas.Facturas = 0
            frmCtas.Conexion = cConta
            frmCtas.Show vbModal
            Set frmCtas = Nothing
            PonerFoco text1(indice)
            
    End Select
    
    If Modo = 4 Then BLOQUEADesdeFormulario2 Me, Data1, 1
End Sub

Private Sub frmCtas_DatoSeleccionado(CadenaSeleccion As String)
'Cuentas contables de la Contabilidad
    text1(indice).Text = RecuperaValor(CadenaSeleccion, 1) 'codmacta
    Text2(indice).Text = RecuperaValor(CadenaSeleccion, 2) 'des macta
End Sub

' ***** si n'hi han varios nivells de tabs *****
Private Sub SituarTab(numTab As Integer)
    On Error Resume Next
    
    SSTab1.Tab = numTab
    
    If Err.Number <> 0 Then Err.Clear
End Sub
' **********************************************

'Private Sub CargaFrame(Index As Integer, enlaza As Boolean)
'Dim tip As Integer
'Dim i As Byte
'
'    AdoAux(Index).ConnectionString = conn
'    AdoAux(Index).RecordSource = MontaSQLCarga(Index, enlaza)
'    AdoAux(Index).CursorType = adOpenDynamic
'    AdoAux(Index).LockType = adLockPessimistic
'    AdoAux(Index).Refresh
'
'    If Not AdoAux(Index).Recordset.EOF Then
'        PonerCamposForma2 Me, AdoAux(Index), 2, "FrameAux" & Index
'    Else
'        ' *** si n'hi han tabs sense datagrids, li pose els valors als camps ***
'        NetejaFrameAux "FrameAux3" 'neteja només lo que te TAG
'    End If
'End Sub

' *** si n'hi han tabs sense datagrids ***
Private Sub NetejaFrameAux(nom_frame As String)
Dim Control As Object
    
    For Each Control In Me.Controls
        If (Control.Tag <> "") Then
            If (Control.Container.Name = nom_frame) Then
                If TypeOf Control Is TextBox Then
                    Control.Text = ""
                ElseIf TypeOf Control Is ComboBox Then
                    Control.ListIndex = -1
                End If
            End If
        End If
    Next Control

End Sub

Private Function ObtenerWhereCab(conW As Boolean) As String
Dim vWhere As String
    
    vWhere = ""
    If conW Then vWhere = " WHERE "
    ' *** canviar-ho per la clau primaria de la capçalera ***
    vWhere = vWhere & " codavnic=" & Val(text1(0).Text)
    
    ObtenerWhereCab = vWhere
End Function

'' *** neteja els camps dels tabs de grid que
''estan fora d'este, i els camps de descripció ***
Private Sub LimpiarCamposFrame(Index As Integer)
    On Error Resume Next
 
'    Select Case Index
'        Case 0 'Cuentas Bancarias
'            txtAux(11).Text = ""
'            txtAux(12).Text = ""
'        Case 1 'Departamentos
'            txtAux(21).Text = ""
'            txtAux(22).Text = ""
'            txtAux2(22).Text = ""
'            txtAux(23).Text = ""
'            txtAux(24).Text = ""
'        Case 2 'Tarjetas
'            txtAux(50).Text = ""
'            txtAux(51).Text = ""
'        Case 4 'comisiones
'            txtAux2(2).Text = ""
'    End Select
'
    If Err.Number <> 0 Then Err.Clear
End Sub

' ### [DavidV] 26/04/2006: Activar/desactivar la rueda del ratón.
'Private Sub DataGridAux_GotFocus(Index As Integer)
'  WheelHook DataGridAux(Index)
'End Sub
'Private Sub DataGridAux_LostFocus(Index As Integer)
'  WheelUnHook
'End Sub

