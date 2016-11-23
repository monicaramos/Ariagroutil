VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmConfParamAplic 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Parámetros de la Aplicación"
   ClientHeight    =   7350
   ClientLeft      =   45
   ClientTop       =   735
   ClientWidth     =   9045
   Icon            =   "frmConfParamAplic.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7350
   ScaleWidth      =   9045
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTab1 
      Height          =   5955
      Left            =   120
      TabIndex        =   33
      Top             =   630
      Width           =   8790
      _ExtentX        =   15505
      _ExtentY        =   10504
      _Version        =   393216
      Tabs            =   10
      Tab             =   9
      TabsPerRow      =   6
      TabHeight       =   520
      ForeColor       =   8388608
      TabCaption(0)   =   "Avnics"
      TabPicture(0)   =   "frmConfParamAplic.frx":000C
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Frame6"
      Tab(0).Control(1)=   "Frame2"
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Seguros"
      TabPicture(1)   =   "frmConfParamAplic.frx":0028
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame3"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Telefonía"
      TabPicture(2)   =   "frmConfParamAplic.frx":0044
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame8"
      Tab(2).Control(1)=   "Frame5"
      Tab(2).ControlCount=   2
      TabCaption(3)   =   "Facturas Varias"
      TabPicture(3)   =   "frmConfParamAplic.frx":0060
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Frame9"
      Tab(3).ControlCount=   1
      TabCaption(4)   =   "Internet"
      TabPicture(4)   =   "frmConfParamAplic.frx":007C
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Frame4"
      Tab(4).Control(1)=   "Frame7"
      Tab(4).ControlCount=   2
      TabCaption(5)   =   "Gasolinera"
      TabPicture(5)   =   "frmConfParamAplic.frx":0098
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "Frame11"
      Tab(5).Control(1)=   "Frame10"
      Tab(5).ControlCount=   2
      TabCaption(6)   =   "Facturas Socios"
      TabPicture(6)   =   "frmConfParamAplic.frx":00B4
      Tab(6).ControlEnabled=   0   'False
      Tab(6).Control(0)=   "Frame12"
      Tab(6).ControlCount=   1
      TabCaption(7)   =   "Fact.Coarval"
      TabPicture(7)   =   "frmConfParamAplic.frx":00D0
      Tab(7).ControlEnabled=   0   'False
      Tab(7).Control(0)=   "Frame14"
      Tab(7).ControlCount=   1
      TabCaption(8)   =   "Fact.Varias"
      TabPicture(8)   =   "frmConfParamAplic.frx":00EC
      Tab(8).ControlEnabled=   0   'False
      Tab(8).Control(0)=   "Frame13"
      Tab(8).ControlCount=   1
      TabCaption(9)   =   "Aridoc"
      TabPicture(9)   =   "frmConfParamAplic.frx":0108
      Tab(9).ControlEnabled=   -1  'True
      Tab(9).Control(0)=   "imgBuscar(37)"
      Tab(9).Control(0).Enabled=   0   'False
      Tab(9).Control(1)=   "Label1(85)"
      Tab(9).Control(1).Enabled=   0   'False
      Tab(9).Control(2)=   "imgBuscar(38)"
      Tab(9).Control(2).Enabled=   0   'False
      Tab(9).Control(3)=   "Label1(86)"
      Tab(9).Control(3).Enabled=   0   'False
      Tab(9).Control(4)=   "Frame15"
      Tab(9).Control(4).Enabled=   0   'False
      Tab(9).Control(5)=   "Text1(83)"
      Tab(9).Control(5).Enabled=   0   'False
      Tab(9).Control(6)=   "Text2(83)"
      Tab(9).Control(6).Enabled=   0   'False
      Tab(9).Control(7)=   "Text2(84)"
      Tab(9).Control(7).Enabled=   0   'False
      Tab(9).Control(8)=   "Text1(84)"
      Tab(9).Control(8).Enabled=   0   'False
      Tab(9).ControlCount=   9
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   84
         Left            =   2460
         MaxLength       =   10
         TabIndex        =   232
         Tag             =   "Extension|N|N|||sparam|codextension|000||"
         Top             =   1800
         Width           =   1215
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   84
         Left            =   3720
         TabIndex        =   237
         Top             =   1800
         Width           =   4470
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   83
         Left            =   3750
         TabIndex        =   233
         Top             =   1350
         Width           =   4470
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   83
         Left            =   2460
         MaxLength       =   10
         TabIndex        =   231
         Tag             =   "Carpeta Facturas|N|N|||sparam|codcarpetafvar|000||"
         Top             =   1350
         Width           =   1215
      End
      Begin VB.Frame Frame15 
         Caption         =   "Facturas Varias"
         ForeColor       =   &H00972E0B&
         Height          =   1050
         Left            =   270
         TabIndex        =   226
         Top             =   2460
         Width           =   8025
         Begin VB.ComboBox Combo1 
            Height          =   315
            Index           =   5
            Left            =   240
            Style           =   2  'Dropdown List
            TabIndex        =   234
            Tag             =   "C1 Factura|N|N|||sparam|c1fvararidoc||N|"
            Top             =   585
            Width           =   1710
         End
         Begin VB.ComboBox Combo1 
            Height          =   315
            Index           =   6
            Left            =   2130
            Style           =   2  'Dropdown List
            TabIndex        =   236
            Tag             =   "C2 Factura|N|N|||sparam|c2fvararidoc||N|"
            Top             =   585
            Width           =   1710
         End
         Begin VB.ComboBox Combo1 
            Height          =   315
            Index           =   7
            Left            =   4065
            Style           =   2  'Dropdown List
            TabIndex        =   238
            Tag             =   "C3 Factura|N|N|||sparam|c3fvararidoc||N|"
            Top             =   585
            Width           =   1710
         End
         Begin VB.ComboBox Combo1 
            Height          =   315
            Index           =   8
            Left            =   5955
            Style           =   2  'Dropdown List
            TabIndex        =   240
            Tag             =   "C4 Factura|N|N|||sparam|c4fvararidoc||N|"
            Top             =   585
            Width           =   1710
         End
         Begin VB.Label Label1 
            Caption         =   "Campo 4"
            Height          =   195
            Index           =   84
            Left            =   5955
            TabIndex        =   230
            Top             =   315
            Width           =   1305
         End
         Begin VB.Label Label1 
            Caption         =   "Campo 3"
            Height          =   195
            Index           =   83
            Left            =   4065
            TabIndex        =   229
            Top             =   315
            Width           =   1620
         End
         Begin VB.Label Label1 
            Caption         =   "Campo 2"
            Height          =   195
            Index           =   82
            Left            =   2130
            TabIndex        =   228
            Top             =   315
            Width           =   1755
         End
         Begin VB.Label Label1 
            Caption         =   "Campo 1"
            Height          =   195
            Index           =   81
            Left            =   240
            TabIndex        =   227
            Top             =   315
            Width           =   1620
         End
      End
      Begin VB.Frame Frame13 
         Caption         =   "Datos Contabilidad Facturas Varias"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3255
         Left            =   -74760
         TabIndex        =   211
         Top             =   735
         Width           =   8145
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   81
            Left            =   2550
            MaxLength       =   10
            TabIndex        =   213
            Tag             =   "Cta.Gastos|T|S|||sparam|ctabancocvv|||"
            Top             =   870
            Width           =   1215
         End
         Begin VB.TextBox Text2 
            BackColor       =   &H80000018&
            Enabled         =   0   'False
            Height          =   285
            Index           =   81
            Left            =   3825
            TabIndex        =   220
            Top             =   870
            Width           =   3885
         End
         Begin VB.TextBox Text2 
            BackColor       =   &H80000018&
            Enabled         =   0   'False
            Height          =   285
            Index           =   80
            Left            =   3810
            TabIndex        =   219
            Top             =   2295
            Width           =   3960
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   80
            Left            =   2535
            MaxLength       =   10
            TabIndex        =   216
            Tag             =   "F.pago Banco|N|S|||sparam|codforpaconcvv|000||"
            Top             =   2295
            Width           =   1215
         End
         Begin VB.TextBox Text2 
            BackColor       =   &H80000018&
            Enabled         =   0   'False
            Height          =   285
            Index           =   79
            Left            =   3810
            TabIndex        =   218
            Top             =   1935
            Width           =   3960
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   79
            Left            =   2550
            MaxLength       =   10
            TabIndex        =   215
            Tag             =   "F.pago Banco|N|S|||sparam|codforpabancvv|000||"
            Top             =   1935
            Width           =   1215
         End
         Begin VB.TextBox Text1 
            Height          =   285
            Index           =   78
            Left            =   2550
            MaxLength       =   3
            TabIndex        =   214
            Tag             =   "Numero Serie Facturas Int|T|S|||sparam|letrasercvv|||"
            Text            =   "3wwwwwwwwwwwwwwwwwww"
            Top             =   1335
            Width           =   585
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   77
            Left            =   2565
            MaxLength       =   2
            TabIndex        =   212
            Tag             =   "Nº Contabilidad|N|S|||sparam|numcontacvv|||"
            Text            =   "3"
            Top             =   420
            Width           =   780
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   36
            Left            =   2235
            ToolTipText     =   "Buscar Cta.Contable"
            Top             =   870
            Width           =   240
         End
         Begin VB.Label Label1 
            Caption         =   "Cta.Contable Banco"
            Height          =   195
            Index           =   79
            Left            =   195
            TabIndex        =   224
            Top             =   900
            Width           =   1920
         End
         Begin VB.Label Label1 
            Caption         =   "F.Pago Banco"
            Height          =   195
            Index           =   78
            Left            =   210
            TabIndex        =   223
            Top             =   1950
            Width           =   1845
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   35
            Left            =   2220
            ToolTipText     =   "Buscar f.pago"
            Top             =   2295
            Width           =   240
         End
         Begin VB.Label Label1 
            Caption         =   "F.Pago Contado"
            Height          =   195
            Index           =   77
            Left            =   210
            TabIndex        =   222
            Top             =   2295
            Width           =   1845
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   34
            Left            =   2220
            ToolTipText     =   "Buscar f.pago"
            Top             =   1935
            Width           =   240
         End
         Begin VB.Label Label1 
            Caption         =   "Letra Serie Facturas"
            Height          =   195
            Index           =   76
            Left            =   210
            TabIndex        =   221
            Top             =   1380
            Width           =   1815
         End
         Begin VB.Label Label1 
            Caption         =   "Nº conta"
            Height          =   195
            Index           =   75
            Left            =   210
            TabIndex        =   217
            Top             =   450
            Width           =   1800
         End
      End
      Begin VB.Frame Frame14 
         Caption         =   "Datos Contabilidad Facturas Tienda"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   5145
         Left            =   -74790
         TabIndex        =   169
         Top             =   735
         Width           =   8385
         Begin VB.TextBox Text1 
            Height          =   285
            Index           =   76
            Left            =   7500
            MaxLength       =   3
            TabIndex        =   181
            Tag             =   "Numero Serie Facturas Int|T|S|||sparam|letraserfincv|||"
            Text            =   "3wwwwwwwwwwwwwwwwwww"
            Top             =   2790
            Width           =   585
         End
         Begin VB.TextBox Text1 
            Height          =   285
            Index           =   75
            Left            =   4770
            MaxLength       =   3
            TabIndex        =   180
            Tag             =   "Numero Serie Facturas|T|S|||sparam|letraserfcv|||"
            Text            =   "3wwwwwwwwwwwwwwwwwww"
            Top             =   2790
            Width           =   555
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   74
            Left            =   2880
            MaxLength       =   10
            TabIndex        =   184
            Tag             =   "Cta.Contable Cliente ticket|T|S|||sparam|ctaclientickcv|||"
            Top             =   3870
            Width           =   1215
         End
         Begin VB.TextBox Text2 
            BackColor       =   &H80000018&
            Enabled         =   0   'False
            Height          =   285
            Index           =   74
            Left            =   4140
            TabIndex        =   207
            Top             =   3870
            Width           =   3960
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   73
            Left            =   2865
            MaxLength       =   10
            TabIndex        =   187
            Tag             =   "Cta.Contable Venta fact|T|S|||sparam|ctaventafacincv|||"
            Top             =   4800
            Width           =   1215
         End
         Begin VB.TextBox Text2 
            BackColor       =   &H80000018&
            Enabled         =   0   'False
            Height          =   285
            Index           =   73
            Left            =   4140
            TabIndex        =   205
            Top             =   4800
            Width           =   3960
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   72
            Left            =   2880
            MaxLength       =   10
            TabIndex        =   183
            Tag             =   "F.pago Banco|N|S|||sparam|codforpaconcv|000||"
            Top             =   3540
            Width           =   1215
         End
         Begin VB.TextBox Text2 
            BackColor       =   &H80000018&
            Enabled         =   0   'False
            Height          =   285
            Index           =   72
            Left            =   4140
            TabIndex        =   203
            Top             =   3540
            Width           =   3960
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   71
            Left            =   2865
            MaxLength       =   10
            TabIndex        =   182
            Tag             =   "F.pago Banco|N|S|||sparam|codforpabancv|000||"
            Top             =   3210
            Width           =   1215
         End
         Begin VB.TextBox Text2 
            BackColor       =   &H80000018&
            Enabled         =   0   'False
            Height          =   285
            Index           =   71
            Left            =   4140
            TabIndex        =   201
            Top             =   3210
            Width           =   3960
         End
         Begin VB.TextBox Text2 
            BackColor       =   &H80000018&
            Enabled         =   0   'False
            Height          =   285
            Index           =   70
            Left            =   4140
            TabIndex        =   199
            Top             =   4500
            Width           =   3960
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   70
            Left            =   2865
            MaxLength       =   10
            TabIndex        =   186
            Tag             =   "Cta.Contable Venta fact|T|S|||sparam|ctaventafaccv|||"
            Top             =   4500
            Width           =   1215
         End
         Begin VB.TextBox Text2 
            BackColor       =   &H80000018&
            Enabled         =   0   'False
            Height          =   285
            Index           =   69
            Left            =   4140
            TabIndex        =   197
            Top             =   4170
            Width           =   3960
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   69
            Left            =   2865
            MaxLength       =   10
            TabIndex        =   185
            Tag             =   "Cta.Contable Venta ticket|T|S|||sparam|ctaventatickcv|||"
            Top             =   4170
            Width           =   1215
         End
         Begin VB.TextBox Text2 
            BackColor       =   &H80000018&
            Enabled         =   0   'False
            Height          =   285
            Index           =   64
            Left            =   4140
            TabIndex        =   194
            Top             =   2040
            Width           =   3900
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   64
            Left            =   2865
            MaxLength       =   10
            TabIndex        =   177
            Tag             =   "Raiz Cta.Socio|T|S|||sparam|raizctasoccv|||"
            Top             =   2040
            Width           =   1215
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   62
            Left            =   2865
            MaxLength       =   10
            TabIndex        =   178
            Tag             =   "Cta.Contable Venta|T|S|||sparam|ctaventacv|||"
            Top             =   2370
            Width           =   1215
         End
         Begin VB.TextBox Text2 
            BackColor       =   &H80000018&
            Enabled         =   0   'False
            Height          =   285
            Index           =   62
            Left            =   4140
            TabIndex        =   193
            Top             =   2370
            Width           =   3900
         End
         Begin VB.TextBox Text1 
            Height          =   285
            Index           =   61
            Left            =   2550
            MaxLength       =   3
            TabIndex        =   179
            Tag             =   "Numero Serie Tickets|T|S|||sparam|letrasercv|||"
            Text            =   "3wwwwwwwwwwwwwwwwwww"
            Top             =   2790
            Width           =   555
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   68
            Left            =   2565
            MaxLength       =   2
            TabIndex        =   174
            Tag             =   "Nº Contabilidad|N|S|||sparam|numcontacv|||"
            Text            =   "3"
            Top             =   1170
            Width           =   780
         End
         Begin VB.TextBox Text1 
            Height          =   285
            Index           =   66
            Left            =   2565
            MaxLength       =   20
            TabIndex        =   172
            Tag             =   "Usuario Contabilidad|T|S|||sparam|usucontacv|||"
            Text            =   "3wwwwwwwwwwwwwwwwwww"
            Top             =   585
            Width           =   4560
         End
         Begin VB.TextBox Text1 
            Height          =   285
            IMEMode         =   3  'DISABLE
            Index           =   67
            Left            =   2565
            MaxLength       =   20
            PasswordChar    =   "*"
            TabIndex        =   173
            Tag             =   "Password Contabilidad|T|S|||sparam|pascontacv|||"
            Text            =   "3"
            Top             =   870
            Width           =   4560
         End
         Begin VB.TextBox Text1 
            Height          =   285
            Index           =   65
            Left            =   2565
            MaxLength       =   20
            TabIndex        =   171
            Tag             =   "Servidor Contabilidad|T|S|||sparam|sercontacv|||"
            Text            =   "3wwwwwwwwwwwwwwwwwwwwwwwwwwwww"
            Top             =   300
            Width           =   4560
         End
         Begin VB.TextBox Text2 
            BackColor       =   &H80000018&
            Enabled         =   0   'False
            Height          =   285
            Index           =   63
            Left            =   4155
            TabIndex        =   170
            Top             =   1695
            Width           =   3885
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   63
            Left            =   2880
            MaxLength       =   10
            TabIndex        =   176
            Tag             =   "Cta.Gastos|T|S|||sparam|ctabancocv|||"
            Top             =   1695
            Width           =   1215
         End
         Begin VB.Label Label1 
            Caption         =   "Letra Serie Fact.Internas"
            Height          =   195
            Index           =   74
            Left            =   5610
            TabIndex        =   210
            Top             =   2835
            Width           =   1815
         End
         Begin VB.Label Label1 
            Caption         =   "Letra Serie Facturas"
            Height          =   195
            Index           =   73
            Left            =   3270
            TabIndex        =   209
            Top             =   2835
            Width           =   1815
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   33
            Left            =   2550
            ToolTipText     =   "Buscar Cta.Contable"
            Top             =   3870
            Width           =   240
         End
         Begin VB.Label Label1 
            Caption         =   "Cta.Cble Cliente Ticket"
            Height          =   195
            Index           =   72
            Left            =   480
            TabIndex        =   208
            Top             =   3900
            Width           =   2010
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   32
            Left            =   2550
            ToolTipText     =   "Buscar Cta.Contable"
            Top             =   4800
            Width           =   240
         End
         Begin VB.Label Label1 
            Caption         =   "Cta.Cble Base Fac.In.Vta"
            Height          =   195
            Index           =   71
            Left            =   480
            TabIndex        =   206
            Top             =   4830
            Width           =   2010
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   31
            Left            =   2550
            ToolTipText     =   "Buscar f.pago"
            Top             =   3540
            Width           =   240
         End
         Begin VB.Label Label1 
            Caption         =   "F.Pago Contado"
            Height          =   195
            Index           =   70
            Left            =   480
            TabIndex        =   204
            Top             =   3570
            Width           =   1845
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   30
            Left            =   2550
            ToolTipText     =   "Buscar f.pago"
            Top             =   3210
            Width           =   240
         End
         Begin VB.Label Label1 
            Caption         =   "F.Pago Banco"
            Height          =   195
            Index           =   65
            Left            =   480
            TabIndex        =   202
            Top             =   3240
            Width           =   1845
         End
         Begin VB.Label Label1 
            Caption         =   "Cta.Cble Base Fac.Vta"
            Height          =   195
            Index           =   64
            Left            =   480
            TabIndex        =   200
            Top             =   4530
            Width           =   2010
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   29
            Left            =   2550
            ToolTipText     =   "Buscar Cta.Contable"
            Top             =   4500
            Width           =   240
         End
         Begin VB.Label Label1 
            Caption         =   "Cta.Cble Base Ticket"
            Height          =   195
            Index           =   63
            Left            =   480
            TabIndex        =   198
            Top             =   4200
            Width           =   2010
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   28
            Left            =   2550
            ToolTipText     =   "Buscar Cta.Contable"
            Top             =   4170
            Width           =   240
         End
         Begin VB.Label Label1 
            Caption         =   "Raiz Cta.Contable Socio"
            Height          =   195
            Index           =   62
            Left            =   480
            TabIndex        =   196
            Top             =   2070
            Width           =   1845
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   27
            Left            =   2550
            ToolTipText     =   "Buscar Raiz Cta.Contable"
            Top             =   2040
            Width           =   240
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   25
            Left            =   2550
            ToolTipText     =   "Buscar Cta.Contable"
            Top             =   2370
            Width           =   240
         End
         Begin VB.Label Label1 
            Caption         =   "Cta.Cble Base Fac.Compra"
            Height          =   195
            Index           =   57
            Left            =   480
            TabIndex        =   195
            Top             =   2400
            Width           =   2010
         End
         Begin VB.Label Label1 
            Caption         =   "Letra Serie Tickets"
            Height          =   195
            Index           =   56
            Left            =   480
            TabIndex        =   192
            Top             =   2835
            Width           =   2445
         End
         Begin VB.Label Label1 
            Caption         =   "Servidor"
            Height          =   195
            Index           =   69
            Left            =   510
            TabIndex        =   191
            Top             =   390
            Width           =   900
         End
         Begin VB.Label Label1 
            Caption         =   "Nº conta"
            Height          =   195
            Index           =   68
            Left            =   510
            TabIndex        =   190
            Top             =   1230
            Width           =   900
         End
         Begin VB.Label Label1 
            Caption         =   "Usuario"
            Height          =   195
            Index           =   67
            Left            =   510
            TabIndex        =   189
            Top             =   660
            Width           =   780
         End
         Begin VB.Label Label1 
            Caption         =   "Password"
            Height          =   195
            Index           =   66
            Left            =   510
            TabIndex        =   188
            Top             =   930
            Width           =   840
         End
         Begin VB.Label Label1 
            Caption         =   "Cta.Contable Banco"
            Height          =   195
            Index           =   61
            Left            =   495
            TabIndex        =   175
            Top             =   1725
            Width           =   1920
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   26
            Left            =   2565
            ToolTipText     =   "Buscar Cta.Contable"
            Top             =   1695
            Width           =   240
         End
      End
      Begin VB.Frame Frame12 
         Caption         =   "Datos Contabilidad"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3675
         Left            =   -74820
         TabIndex        =   147
         Top             =   945
         Width           =   7665
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   59
            Left            =   2895
            MaxLength       =   10
            TabIndex        =   157
            Tag             =   "Cta.Retencion|T|S|||sparam|ctaretenfacsoc|||"
            Top             =   2340
            Width           =   1215
         End
         Begin VB.TextBox Text2 
            BackColor       =   &H80000018&
            Enabled         =   0   'False
            Height          =   285
            Index           =   59
            Left            =   4155
            TabIndex        =   151
            Top             =   2340
            Width           =   2895
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   58
            Left            =   2565
            MaxLength       =   10
            TabIndex        =   155
            Tag             =   "Porcentaje Retención|N|S|||sparam|porcretenfacsoc||##0.00|"
            Top             =   1890
            Width           =   585
         End
         Begin VB.TextBox Text1 
            Height          =   285
            Index           =   57
            Left            =   2565
            MaxLength       =   20
            TabIndex        =   149
            Tag             =   "Servidor Contabilidad|T|S|||sparam|sercontafacsoc|||"
            Text            =   "3wwwwwwwwwwwwwwwwwwwwwwwwwwwww"
            Top             =   180
            Width           =   4560
         End
         Begin VB.TextBox Text1 
            Height          =   285
            IMEMode         =   3  'DISABLE
            Index           =   56
            Left            =   2565
            MaxLength       =   20
            PasswordChar    =   "*"
            TabIndex        =   152
            Tag             =   "Password Contabilidad|T|S|||sparam|pascontafacsoc|||"
            Text            =   "3"
            Top             =   855
            Width           =   4560
         End
         Begin VB.TextBox Text1 
            Height          =   285
            Index           =   55
            Left            =   2565
            MaxLength       =   20
            TabIndex        =   150
            Tag             =   "Usuario Contabilidad|T|S|||sparam|usucontafacsoc|||"
            Text            =   "3wwwwwwwwwwwwwwwwwww"
            Top             =   525
            Width           =   4560
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   54
            Left            =   2565
            MaxLength       =   2
            TabIndex        =   153
            Tag             =   "Nº Contabilidad|N|S|||sparam|numcontafacsoc|||"
            Text            =   "3"
            Top             =   1170
            Width           =   780
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   53
            Left            =   2895
            MaxLength       =   10
            TabIndex        =   159
            Tag             =   "Raiz Cta.Socio|T|S|||sparam|raizctafacsoc|||"
            Top             =   2745
            Width           =   1215
         End
         Begin VB.TextBox Text2 
            BackColor       =   &H80000018&
            Enabled         =   0   'False
            Height          =   285
            Index           =   53
            Left            =   4155
            TabIndex        =   148
            Top             =   2745
            Width           =   2895
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   24
            Left            =   2565
            ToolTipText     =   "Buscar Cta.Contable"
            Top             =   2340
            Width           =   240
         End
         Begin VB.Label Label1 
            Caption         =   "Cta.Retención"
            Height          =   195
            Index           =   59
            Left            =   495
            TabIndex        =   163
            Top             =   2370
            Width           =   1920
         End
         Begin VB.Label Label1 
            Caption         =   "Porcentaje Retención"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   58
            Left            =   495
            TabIndex        =   162
            Top             =   1935
            Width           =   2040
         End
         Begin VB.Label Label1 
            Caption         =   "Password"
            Height          =   195
            Index           =   55
            Left            =   510
            TabIndex        =   161
            Top             =   900
            Width           =   840
         End
         Begin VB.Label Label1 
            Caption         =   "Usuario"
            Height          =   195
            Index           =   54
            Left            =   510
            TabIndex        =   160
            Top             =   600
            Width           =   780
         End
         Begin VB.Label Label1 
            Caption         =   "Nº conta"
            Height          =   195
            Index           =   53
            Left            =   510
            TabIndex        =   158
            Top             =   1230
            Width           =   900
         End
         Begin VB.Label Label1 
            Caption         =   "Servidor"
            Height          =   195
            Index           =   52
            Left            =   510
            TabIndex        =   156
            Top             =   270
            Width           =   900
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   23
            Left            =   2565
            ToolTipText     =   "Buscar Raiz Cta.Contable"
            Top             =   2745
            Width           =   240
         End
         Begin VB.Label Label1 
            Caption         =   "Raiz Cta.Contable Socio"
            Height          =   195
            Index           =   51
            Left            =   495
            TabIndex        =   154
            Top             =   2775
            Width           =   1845
         End
      End
      Begin VB.Frame Frame11 
         Caption         =   "Datos Contabilidad"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4260
         Left            =   -74865
         TabIndex        =   127
         Top             =   720
         Width           =   7665
         Begin VB.TextBox Text2 
            BackColor       =   &H80000018&
            Enabled         =   0   'False
            Height          =   285
            Index           =   52
            Left            =   3600
            Locked          =   -1  'True
            TabIndex        =   145
            Top             =   3870
            Width           =   3480
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   52
            Left            =   2880
            MaxLength       =   10
            TabIndex        =   113
            Tag             =   "Código de Iva|N|S|||sparam|codivagas||000|"
            Top             =   3870
            Width           =   585
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   50
            Left            =   2565
            MaxLength       =   2
            TabIndex        =   106
            Tag             =   "Nº Contabilidad|N|S|||sparam|numcontagas|||"
            Text            =   "3"
            Top             =   1260
            Width           =   780
         End
         Begin VB.TextBox Text1 
            Height          =   285
            Index           =   49
            Left            =   2565
            MaxLength       =   20
            TabIndex        =   104
            Tag             =   "Usuario Contabilidad|T|S|||sparam|usucontagas|||"
            Text            =   "3wwwwwwwwwwwwwwwwwww"
            Top             =   615
            Width           =   4560
         End
         Begin VB.TextBox Text1 
            Height          =   285
            IMEMode         =   3  'DISABLE
            Index           =   48
            Left            =   2565
            MaxLength       =   20
            PasswordChar    =   "*"
            TabIndex        =   105
            Tag             =   "Password Contabilidad|T|S|||sparam|pascontagas|||"
            Text            =   "3"
            Top             =   930
            Width           =   4560
         End
         Begin VB.TextBox Text1 
            Height          =   285
            Index           =   47
            Left            =   2565
            MaxLength       =   20
            TabIndex        =   103
            Tag             =   "Servidor Contabilidad|T|S|||sparam|sercontagas|||"
            Text            =   "3wwwwwwwwwwwwwwwwwwwwwwwwwwwww"
            Top             =   270
            Width           =   4560
         End
         Begin VB.TextBox Text2 
            BackColor       =   &H80000018&
            Enabled         =   0   'False
            Height          =   285
            Index           =   46
            Left            =   3600
            Locked          =   -1  'True
            TabIndex        =   133
            Top             =   2775
            Width           =   3480
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   46
            Left            =   2880
            MaxLength       =   10
            TabIndex        =   110
            Tag             =   "Numero de Diario|N|S|||sparam|numdiarigas||000|"
            Top             =   2775
            Width           =   585
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   45
            Left            =   2880
            MaxLength       =   10
            TabIndex        =   109
            Tag             =   "Concepto Haber|N|S|||sparam|concehabergas|||"
            Top             =   2415
            Width           =   585
         End
         Begin VB.TextBox Text2 
            BackColor       =   &H80000018&
            Enabled         =   0   'False
            Height          =   285
            Index           =   45
            Left            =   3600
            Locked          =   -1  'True
            TabIndex        =   132
            Top             =   2400
            Width           =   3480
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   44
            Left            =   2880
            MaxLength       =   10
            TabIndex        =   108
            Tag             =   "Concepto Debe|N|S|||sparam|concedebegas|||"
            Top             =   2040
            Width           =   585
         End
         Begin VB.TextBox Text2 
            BackColor       =   &H80000018&
            Enabled         =   0   'False
            Height          =   285
            Index           =   44
            Left            =   3600
            Locked          =   -1  'True
            TabIndex        =   131
            Top             =   2040
            Width           =   3450
         End
         Begin VB.TextBox Text2 
            BackColor       =   &H80000018&
            Enabled         =   0   'False
            Height          =   285
            Index           =   43
            Left            =   4155
            TabIndex        =   130
            Top             =   3150
            Width           =   2940
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   43
            Left            =   2880
            MaxLength       =   10
            TabIndex        =   111
            Tag             =   "Raiz Cta.Socio|T|S|||sparam|raizctasocgas|||"
            Top             =   3150
            Width           =   1215
         End
         Begin VB.TextBox Text2 
            BackColor       =   &H80000018&
            Enabled         =   0   'False
            Height          =   285
            Index           =   42
            Left            =   4155
            TabIndex        =   129
            Top             =   1620
            Width           =   2895
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   42
            Left            =   2880
            MaxLength       =   10
            TabIndex        =   107
            Tag             =   "Cta.Gastos|T|S|||sparam|ctaventasgas|||"
            Top             =   1620
            Width           =   1215
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   41
            Left            =   2880
            MaxLength       =   10
            TabIndex        =   112
            Tag             =   "Cta.Contable Contrapartida|T|S|||sparam|ctacontragas|||"
            Top             =   3510
            Width           =   1215
         End
         Begin VB.TextBox Text2 
            BackColor       =   &H80000018&
            Enabled         =   0   'False
            Height          =   285
            Index           =   41
            Left            =   4155
            TabIndex        =   128
            Top             =   3510
            Width           =   2940
         End
         Begin VB.Label Label1 
            Caption         =   "Código de Iva"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   50
            Left            =   495
            TabIndex        =   146
            Top             =   3870
            Width           =   1410
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   22
            Left            =   2565
            ToolTipText     =   "Buscar Iva"
            Top             =   3870
            Width           =   240
         End
         Begin VB.Label Label1 
            Caption         =   "Servidor"
            Height          =   195
            Index           =   48
            Left            =   510
            TabIndex        =   143
            Top             =   360
            Width           =   900
         End
         Begin VB.Label Label1 
            Caption         =   "Nº conta"
            Height          =   195
            Index           =   47
            Left            =   510
            TabIndex        =   142
            Top             =   1320
            Width           =   900
         End
         Begin VB.Label Label1 
            Caption         =   "Usuario"
            Height          =   195
            Index           =   46
            Left            =   510
            TabIndex        =   141
            Top             =   690
            Width           =   780
         End
         Begin VB.Label Label1 
            Caption         =   "Password"
            Height          =   195
            Index           =   45
            Left            =   510
            TabIndex        =   140
            Top             =   990
            Width           =   840
         End
         Begin VB.Label Label1 
            Caption         =   "Numero de Diario"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   43
            Left            =   495
            TabIndex        =   139
            Top             =   2775
            Width           =   1410
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   21
            Left            =   2565
            ToolTipText     =   "Buscar Diario"
            Top             =   2775
            Width           =   240
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   20
            Left            =   2565
            ToolTipText     =   "Buscar Concepto"
            Top             =   2430
            Width           =   240
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   19
            Left            =   2565
            ToolTipText     =   "Buscar Concepto"
            Top             =   2040
            Width           =   240
         End
         Begin VB.Label Label1 
            Caption         =   "Concepto al Debe"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   40
            Left            =   495
            TabIndex        =   138
            Top             =   2025
            Width           =   1350
         End
         Begin VB.Label Label1 
            Caption         =   "Concepto al Haber"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   39
            Left            =   495
            TabIndex        =   137
            Top             =   2400
            Width           =   1410
         End
         Begin VB.Label Label1 
            Caption         =   "Raiz Cta.Contable Socio"
            Height          =   195
            Index           =   38
            Left            =   495
            TabIndex        =   136
            Top             =   3180
            Width           =   1845
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   18
            Left            =   2565
            ToolTipText     =   "Buscar Raiz Cta.Contable"
            Top             =   3150
            Width           =   240
         End
         Begin VB.Label Label1 
            Caption         =   "Cta.Contable Venta"
            Height          =   195
            Index           =   37
            Left            =   495
            TabIndex        =   135
            Top             =   1650
            Width           =   1920
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   17
            Left            =   2565
            ToolTipText     =   "Buscar Cta.Contable"
            Top             =   1620
            Width           =   240
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   16
            Left            =   2565
            ToolTipText     =   "Buscar Cta.Contable"
            Top             =   3510
            Width           =   240
         End
         Begin VB.Label Label1 
            Caption         =   "Cta.Contable Contrapartida"
            Height          =   195
            Index           =   36
            Left            =   495
            TabIndex        =   134
            Top             =   3540
            Width           =   1920
         End
      End
      Begin VB.Frame Frame10 
         Caption         =   "Otros"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   885
         Left            =   -74865
         TabIndex        =   125
         Top             =   4995
         Width           =   7665
         Begin VB.TextBox Text1 
            Height          =   285
            Index           =   82
            Left            =   3120
            MaxLength       =   3
            TabIndex        =   116
            Tag             =   "Letra Serie Gas.B|T|S|||sparam|letrasergasB|||"
            Text            =   "3wwwwwwwwwwwwwwwwwww"
            Top             =   540
            Width           =   525
         End
         Begin VB.TextBox Text1 
            Height          =   285
            Index           =   51
            Left            =   6075
            MaxLength       =   7
            TabIndex        =   115
            Tag             =   "Incremento NroFac Gasolinera|N|S|||sparam|increfacgas|0000000||"
            Text            =   "3wwwwwwwwwwwwwwwwwww"
            Top             =   210
            Width           =   960
         End
         Begin VB.TextBox Text1 
            Height          =   285
            Index           =   40
            Left            =   3105
            MaxLength       =   3
            TabIndex        =   114
            Tag             =   "Letra Serie Gasolinera|T|S|||sparam|letrasergas|||"
            Text            =   "3wwwwwwwwwwwwwwwwwww"
            Top             =   210
            Width           =   555
         End
         Begin VB.Label Label1 
            Caption         =   "Letra de Serie Facturas Gas.B"
            Height          =   195
            Index           =   80
            Left            =   510
            TabIndex        =   225
            Top             =   585
            Width           =   2445
         End
         Begin VB.Label Label1 
            Caption         =   "Increm.Nro.Factura Gasolinera"
            Height          =   195
            Index           =   49
            Left            =   3870
            TabIndex        =   144
            Top             =   255
            Width           =   2220
         End
         Begin VB.Label Label1 
            Caption         =   "Letra de Serie Facturas Gasolinera"
            Height          =   195
            Index           =   35
            Left            =   495
            TabIndex        =   126
            Top             =   255
            Width           =   2445
         End
      End
      Begin VB.Frame Frame9 
         Caption         =   "Datos Contabilidad"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1380
         Left            =   -74820
         TabIndex        =   118
         Top             =   855
         Width           =   7665
         Begin VB.TextBox Text1 
            Height          =   285
            Index           =   38
            Left            =   2565
            MaxLength       =   20
            TabIndex        =   121
            Tag             =   "Usuario Contabilidad|T|S|||sparam|usucontafac|||"
            Text            =   "3wwwwwwwwwwwwwwwwwww"
            Top             =   615
            Width           =   4560
         End
         Begin VB.TextBox Text1 
            Height          =   285
            IMEMode         =   3  'DISABLE
            Index           =   39
            Left            =   2565
            MaxLength       =   20
            PasswordChar    =   "*"
            TabIndex        =   120
            Tag             =   "Password Contabilidad|T|S|||sparam|pascontafac|||"
            Text            =   "3"
            Top             =   930
            Width           =   4560
         End
         Begin VB.TextBox Text1 
            Height          =   285
            Index           =   37
            Left            =   2565
            MaxLength       =   20
            TabIndex        =   119
            Tag             =   "Servidor Contabilidad|T|S|||sparam|sercontafac|||"
            Text            =   "3wwwwwwwwwwwwwwwwwwwwwwwwwwwww"
            Top             =   270
            Width           =   4560
         End
         Begin VB.Label Label1 
            Caption         =   "Servidor"
            Height          =   195
            Index           =   44
            Left            =   510
            TabIndex        =   124
            Top             =   360
            Width           =   900
         End
         Begin VB.Label Label1 
            Caption         =   "Usuario"
            Height          =   195
            Index           =   42
            Left            =   510
            TabIndex        =   123
            Top             =   690
            Width           =   780
         End
         Begin VB.Label Label1 
            Caption         =   "Password"
            Height          =   195
            Index           =   41
            Left            =   510
            TabIndex        =   122
            Top             =   990
            Width           =   840
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Soporte"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1035
         Left            =   -74880
         TabIndex        =   102
         Top             =   3690
         Width           =   7845
         Begin VB.TextBox Text1 
            Height          =   285
            Index           =   13
            Left            =   1290
            MaxLength       =   100
            TabIndex        =   96
            Tag             =   "Web Soporte|T|S|||sparam|websoporte|||"
            Top             =   360
            Width           =   6135
         End
         Begin VB.Label Label2 
            Caption         =   "Web soporte"
            Height          =   255
            Index           =   0
            Left            =   180
            TabIndex        =   117
            Top             =   360
            Width           =   1575
         End
      End
      Begin VB.Frame Frame7 
         Height          =   2505
         Left            =   -74910
         TabIndex        =   89
         Top             =   1020
         Width           =   7875
         Begin VB.TextBox Text1 
            Height          =   285
            IMEMode         =   3  'DISABLE
            Index           =   60
            Left            =   2730
            MaxLength       =   30
            TabIndex        =   94
            Tag             =   "LanzaMailOutlook|T|S|||sparam|arigesmail|||"
            Text            =   "3"
            Top             =   1950
            Width           =   1620
         End
         Begin VB.CheckBox chkOutlook 
            Caption         =   "Enviar desde Outlook"
            Height          =   375
            Left            =   5310
            TabIndex        =   95
            Tag             =   "Outlook|N|N|||sparam|EnvioDesdeOutlook|||"
            Top             =   1920
            Width           =   2175
         End
         Begin VB.TextBox Text1 
            Height          =   285
            IMEMode         =   3  'DISABLE
            Index           =   12
            Left            =   5250
            MaxLength       =   15
            PasswordChar    =   "*"
            TabIndex        =   93
            Tag             =   "Password SMTP|T|S|||sparam|smtpPass|||"
            Text            =   "3"
            Top             =   1440
            Width           =   2220
         End
         Begin VB.TextBox Text1 
            Height          =   285
            Index           =   11
            Left            =   1260
            MaxLength       =   50
            TabIndex        =   92
            Tag             =   "Usuario SMTP|T|S|||sparam|smtpUser|||"
            Text            =   "3"
            Top             =   1440
            Width           =   3090
         End
         Begin VB.TextBox Text1 
            Height          =   285
            Index           =   10
            Left            =   1260
            MaxLength       =   50
            TabIndex        =   91
            Tag             =   "Servidor SMTP|T|S|||sparam|smtpHost|||"
            Text            =   "3"
            Top             =   900
            Width           =   6210
         End
         Begin VB.TextBox Text1 
            Height          =   285
            Index           =   9
            Left            =   1260
            MaxLength       =   50
            TabIndex        =   90
            Tag             =   "Direccion e-mail|T|S|||sparam|diremail|||"
            Text            =   "3"
            Top             =   420
            Width           =   6210
         End
         Begin VB.Label Label1 
            Caption         =   "Lanza pantalla mail outlook"
            Height          =   195
            Index           =   60
            Left            =   120
            TabIndex        =   168
            Top             =   1980
            Width           =   2040
         End
         Begin VB.Label Label8 
            Caption         =   "Envio E-Mail"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   0
            Left            =   360
            TabIndex        =   101
            Top             =   0
            Width           =   1320
         End
         Begin VB.Label Label1 
            Caption         =   "Password"
            Height          =   195
            Index           =   23
            Left            =   4440
            TabIndex        =   100
            Top             =   1500
            Width           =   840
         End
         Begin VB.Label Label1 
            Caption         =   "Usuario"
            Height          =   195
            Index           =   22
            Left            =   120
            TabIndex        =   99
            Top             =   1500
            Width           =   1380
         End
         Begin VB.Label Label1 
            Caption         =   "Servidor SMTP"
            Height          =   195
            Index           =   21
            Left            =   120
            TabIndex        =   98
            Top             =   960
            Width           =   1380
         End
         Begin VB.Label Label1 
            Caption         =   "E-Mail"
            Height          =   195
            Index           =   20
            Left            =   120
            TabIndex        =   97
            Top             =   480
            Width           =   1380
         End
      End
      Begin VB.Frame Frame8 
         Caption         =   "Otros"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   705
         Left            =   -74820
         TabIndex        =   87
         Top             =   5040
         Width           =   7620
         Begin VB.ComboBox Combo1 
            Height          =   315
            Index           =   0
            ItemData        =   "frmConfParamAplic.frx":0124
            Left            =   5640
            List            =   "frmConfParamAplic.frx":012E
            Style           =   2  'Dropdown List
            TabIndex        =   166
            Tag             =   "Tipo Fichero Tel|N|N|0|1|sparam|tipoficherotel|||"
            Top             =   270
            Width           =   1455
         End
         Begin VB.TextBox Text1 
            Height          =   285
            Index           =   36
            Left            =   2925
            MaxLength       =   3
            TabIndex        =   28
            Tag             =   "Numero Serie Telefonia|T|N|||sparam|numserietel|||"
            Text            =   "3wwwwwwwwwwwwwwwwwww"
            Top             =   270
            Width           =   555
         End
         Begin VB.Label Label2 
            Caption         =   "Tipo Fichero"
            Height          =   255
            Index           =   1
            Left            =   4530
            TabIndex        =   167
            Top             =   315
            Width           =   900
         End
         Begin VB.Label Label1 
            Caption         =   "Letra de Serie Facturas Telefonía"
            Height          =   195
            Index           =   34
            Left            =   495
            TabIndex        =   88
            Top             =   315
            Width           =   2445
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "Datos Contabilidad"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4125
         Left            =   -74820
         TabIndex        =   70
         Top             =   810
         Width           =   7665
         Begin VB.TextBox Text2 
            BackColor       =   &H80000018&
            Enabled         =   0   'False
            Height          =   285
            Index           =   35
            Left            =   4155
            TabIndex        =   85
            Top             =   3645
            Width           =   2940
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   35
            Left            =   2880
            MaxLength       =   10
            TabIndex        =   27
            Tag             =   "Cta.Contable Venta|T|S|||sparam|ctaventatel|||"
            Top             =   3645
            Width           =   1215
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   34
            Left            =   2880
            MaxLength       =   10
            TabIndex        =   22
            Tag             =   "Cta.Gastos|T|S|||sparam|ctabancotel|||"
            Top             =   1755
            Width           =   1215
         End
         Begin VB.TextBox Text2 
            BackColor       =   &H80000018&
            Enabled         =   0   'False
            Height          =   285
            Index           =   34
            Left            =   4155
            TabIndex        =   83
            Top             =   1755
            Width           =   2895
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   33
            Left            =   2880
            MaxLength       =   10
            TabIndex        =   26
            Tag             =   "Raiz Cta.Socio|T|S|||sparam|raizctasoctel|||"
            Top             =   3285
            Width           =   1215
         End
         Begin VB.TextBox Text2 
            BackColor       =   &H80000018&
            Enabled         =   0   'False
            Height          =   285
            Index           =   33
            Left            =   4155
            TabIndex        =   81
            Top             =   3285
            Width           =   2940
         End
         Begin VB.TextBox Text2 
            BackColor       =   &H80000018&
            Enabled         =   0   'False
            Height          =   285
            Index           =   32
            Left            =   3600
            Locked          =   -1  'True
            TabIndex        =   73
            Top             =   2175
            Width           =   3450
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   32
            Left            =   2880
            MaxLength       =   10
            TabIndex        =   23
            Tag             =   "Concepto Debe|N|S|||sparam|concedebetel|||"
            Top             =   2175
            Width           =   585
         End
         Begin VB.TextBox Text2 
            BackColor       =   &H80000018&
            Enabled         =   0   'False
            Height          =   285
            Index           =   31
            Left            =   3600
            Locked          =   -1  'True
            TabIndex        =   72
            Top             =   2535
            Width           =   3480
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   31
            Left            =   2880
            MaxLength       =   10
            TabIndex        =   24
            Tag             =   "Concepto Haber|N|S|||sparam|concehabertel|||"
            Top             =   2550
            Width           =   585
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   30
            Left            =   2880
            MaxLength       =   10
            TabIndex        =   25
            Tag             =   "Numero de Diario|N|S|||sparam|numdiariotel||000|"
            Top             =   2910
            Width           =   585
         End
         Begin VB.TextBox Text2 
            BackColor       =   &H80000018&
            Enabled         =   0   'False
            Height          =   285
            Index           =   30
            Left            =   3600
            Locked          =   -1  'True
            TabIndex        =   71
            Top             =   2910
            Width           =   3480
         End
         Begin VB.TextBox Text1 
            Height          =   285
            Index           =   29
            Left            =   2565
            MaxLength       =   20
            TabIndex        =   18
            Tag             =   "Servidor Contabilidad|T|S|||sparam|sercontatel|||"
            Text            =   "3wwwwwwwwwwwwwwwwwwwwwwwwwwwww"
            Top             =   180
            Width           =   4560
         End
         Begin VB.TextBox Text1 
            Height          =   285
            IMEMode         =   3  'DISABLE
            Index           =   28
            Left            =   2565
            MaxLength       =   20
            PasswordChar    =   "*"
            TabIndex        =   20
            Tag             =   "Password Contabilidad|T|S|||sparam|pascontatel|||"
            Text            =   "3"
            Top             =   840
            Width           =   4560
         End
         Begin VB.TextBox Text1 
            Height          =   285
            Index           =   27
            Left            =   2565
            MaxLength       =   20
            TabIndex        =   19
            Tag             =   "Usuario Contabilidad|T|S|||sparam|usucontatel|||"
            Text            =   "3wwwwwwwwwwwwwwwwwww"
            Top             =   525
            Width           =   4560
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   26
            Left            =   2565
            MaxLength       =   2
            TabIndex        =   21
            Tag             =   "Nº Contabilidad|N|S|||sparam|numcontatel|||"
            Text            =   "3"
            Top             =   1170
            Width           =   780
         End
         Begin VB.Label Label1 
            Caption         =   "Cta.Contable Venta"
            Height          =   195
            Index           =   33
            Left            =   495
            TabIndex        =   86
            Top             =   3675
            Width           =   1920
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   15
            Left            =   2565
            ToolTipText     =   "Buscar Cta.Contable"
            Top             =   3645
            Width           =   240
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   14
            Left            =   2565
            ToolTipText     =   "Buscar Cta.Contable"
            Top             =   1755
            Width           =   240
         End
         Begin VB.Label Label1 
            Caption         =   "Cta.Contable Banco"
            Height          =   195
            Index           =   32
            Left            =   495
            TabIndex        =   84
            Top             =   1785
            Width           =   1920
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   13
            Left            =   2565
            ToolTipText     =   "Buscar Raiz Cta.Contable"
            Top             =   3285
            Width           =   240
         End
         Begin VB.Label Label1 
            Caption         =   "Raiz Cta.Contable Socio"
            Height          =   195
            Index           =   14
            Left            =   495
            TabIndex        =   82
            Top             =   3315
            Width           =   1845
         End
         Begin VB.Label Label1 
            Caption         =   "Concepto al Haber"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   31
            Left            =   495
            TabIndex        =   80
            Top             =   2535
            Width           =   1410
         End
         Begin VB.Label Label1 
            Caption         =   "Concepto al Debe"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   30
            Left            =   495
            TabIndex        =   79
            Top             =   2160
            Width           =   1350
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   12
            Left            =   2565
            ToolTipText     =   "Buscar Concepto"
            Top             =   2175
            Width           =   240
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   11
            Left            =   2565
            ToolTipText     =   "Buscar Concepto"
            Top             =   2565
            Width           =   240
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   10
            Left            =   2565
            ToolTipText     =   "Buscar Diario"
            Top             =   2910
            Width           =   240
         End
         Begin VB.Label Label1 
            Caption         =   "Numero de Diario"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   29
            Left            =   495
            TabIndex        =   78
            Top             =   2910
            Width           =   1410
         End
         Begin VB.Label Label1 
            Caption         =   "Password"
            Height          =   195
            Index           =   28
            Left            =   510
            TabIndex        =   77
            Top             =   900
            Width           =   840
         End
         Begin VB.Label Label1 
            Caption         =   "Usuario"
            Height          =   195
            Index           =   27
            Left            =   510
            TabIndex        =   76
            Top             =   600
            Width           =   780
         End
         Begin VB.Label Label1 
            Caption         =   "Nº conta"
            Height          =   195
            Index           =   26
            Left            =   510
            TabIndex        =   75
            Top             =   1230
            Width           =   900
         End
         Begin VB.Label Label1 
            Caption         =   "Servidor"
            Height          =   195
            Index           =   16
            Left            =   510
            TabIndex        =   74
            Top             =   270
            Width           =   900
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Datos Contabilidad"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3675
         Left            =   -74820
         TabIndex        =   55
         Top             =   900
         Width           =   7665
         Begin VB.TextBox Text2 
            BackColor       =   &H80000018&
            Enabled         =   0   'False
            Height          =   285
            Index           =   7
            Left            =   4155
            TabIndex        =   68
            Top             =   3285
            Width           =   2895
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   7
            Left            =   2895
            MaxLength       =   10
            TabIndex        =   17
            Tag             =   "Raiz Cta.Socio|T|S|||sparam|raizctasocseg|||"
            Top             =   3285
            Width           =   1215
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   25
            Left            =   2565
            MaxLength       =   2
            TabIndex        =   12
            Tag             =   "Nº Contabilidad|N|S|||sparam|numcontaseg|||"
            Text            =   "3"
            Top             =   1170
            Width           =   780
         End
         Begin VB.TextBox Text1 
            Height          =   285
            Index           =   24
            Left            =   2565
            MaxLength       =   20
            TabIndex        =   10
            Tag             =   "Usuario Contabilidad|T|S|||sparam|usucontaseg|||"
            Text            =   "3wwwwwwwwwwwwwwwwwww"
            Top             =   525
            Width           =   4560
         End
         Begin VB.TextBox Text1 
            Height          =   285
            IMEMode         =   3  'DISABLE
            Index           =   23
            Left            =   2565
            MaxLength       =   20
            PasswordChar    =   "*"
            TabIndex        =   11
            Tag             =   "Password Contabilidad|T|S|||sparam|pascontaseg|||"
            Text            =   "3"
            Top             =   855
            Width           =   4560
         End
         Begin VB.TextBox Text1 
            Height          =   285
            Index           =   22
            Left            =   2565
            MaxLength       =   20
            TabIndex        =   9
            Tag             =   "Servidor Contabilidad|T|S|||sparam|sercontaseg|||"
            Text            =   "3wwwwwwwwwwwwwwwwwwwwwwwwwwwww"
            Top             =   180
            Width           =   4560
         End
         Begin VB.TextBox Text2 
            BackColor       =   &H80000018&
            Enabled         =   0   'False
            Height          =   285
            Index           =   21
            Left            =   3600
            Locked          =   -1  'True
            TabIndex        =   59
            Top             =   2910
            Width           =   3480
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   21
            Left            =   2895
            MaxLength       =   10
            TabIndex        =   16
            Tag             =   "Numero de Diario|N|S|||sparam|numdiarioseg||000|"
            Top             =   2910
            Width           =   585
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   20
            Left            =   2895
            MaxLength       =   10
            TabIndex        =   15
            Tag             =   "Concepto Haber|N|S|||sparam|concehaberseg|||"
            Top             =   2550
            Width           =   585
         End
         Begin VB.TextBox Text2 
            BackColor       =   &H80000018&
            Enabled         =   0   'False
            Height          =   285
            Index           =   20
            Left            =   3600
            Locked          =   -1  'True
            TabIndex        =   58
            Top             =   2535
            Width           =   3480
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   19
            Left            =   2895
            MaxLength       =   10
            TabIndex        =   14
            Tag             =   "Concepto Debe|N|S|||sparam|concedebeseg|||"
            Top             =   2175
            Width           =   585
         End
         Begin VB.TextBox Text2 
            BackColor       =   &H80000018&
            Enabled         =   0   'False
            Height          =   285
            Index           =   19
            Left            =   3600
            Locked          =   -1  'True
            TabIndex        =   57
            Top             =   2175
            Width           =   3450
         End
         Begin VB.TextBox Text2 
            BackColor       =   &H80000018&
            Enabled         =   0   'False
            Height          =   285
            Index           =   8
            Left            =   4155
            TabIndex        =   56
            Top             =   1755
            Width           =   2895
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   8
            Left            =   2895
            MaxLength       =   10
            TabIndex        =   13
            Tag             =   "Cta.Gastos|T|S|||sparam|ctabancoseg|||"
            Top             =   1755
            Width           =   1215
         End
         Begin VB.Label Label1 
            Caption         =   "Raiz Cta.Contable Socio"
            Height          =   195
            Index           =   6
            Left            =   495
            TabIndex        =   69
            Top             =   3315
            Width           =   1845
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   5
            Left            =   2565
            ToolTipText     =   "Buscar Raiz Cta.Contable"
            Top             =   3285
            Width           =   240
         End
         Begin VB.Label Label1 
            Caption         =   "Servidor"
            Height          =   195
            Index           =   13
            Left            =   510
            TabIndex        =   67
            Top             =   270
            Width           =   900
         End
         Begin VB.Label Label1 
            Caption         =   "Nº conta"
            Height          =   195
            Index           =   12
            Left            =   510
            TabIndex        =   66
            Top             =   1230
            Width           =   900
         End
         Begin VB.Label Label1 
            Caption         =   "Usuario"
            Height          =   195
            Index           =   11
            Left            =   510
            TabIndex        =   65
            Top             =   600
            Width           =   780
         End
         Begin VB.Label Label1 
            Caption         =   "Password"
            Height          =   195
            Index           =   10
            Left            =   510
            TabIndex        =   64
            Top             =   900
            Width           =   840
         End
         Begin VB.Label Label1 
            Caption         =   "Numero de Diario"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   9
            Left            =   495
            TabIndex        =   63
            Top             =   2910
            Width           =   1410
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   9
            Left            =   2565
            ToolTipText     =   "Buscar Diario"
            Top             =   2910
            Width           =   240
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   8
            Left            =   2565
            ToolTipText     =   "Buscar Concepto"
            Top             =   2565
            Width           =   240
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   7
            Left            =   2565
            ToolTipText     =   "Buscar Concepto"
            Top             =   2175
            Width           =   240
         End
         Begin VB.Label Label1 
            Caption         =   "Concepto al Debe"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   8
            Left            =   495
            TabIndex        =   62
            Top             =   2160
            Width           =   1350
         End
         Begin VB.Label Label1 
            Caption         =   "Concepto al Haber"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   7
            Left            =   495
            TabIndex        =   61
            Top             =   2535
            Width           =   1410
         End
         Begin VB.Label Label1 
            Caption         =   "Cta.Contable Banco"
            Height          =   195
            Index           =   0
            Left            =   495
            TabIndex        =   60
            Top             =   1785
            Width           =   1920
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   6
            Left            =   2565
            ToolTipText     =   "Buscar Cta.Contable"
            Top             =   1755
            Width           =   240
         End
      End
      Begin VB.Frame Frame6 
         Caption         =   "Otros"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   840
         Left            =   -74820
         TabIndex        =   40
         Top             =   4725
         Width           =   7620
         Begin VB.TextBox Text1 
            Height          =   285
            Index           =   18
            Left            =   1440
            MaxLength       =   20
            TabIndex        =   42
            Tag             =   "% Interes|N|S|||sparam|porcinte|##0.00||"
            Text            =   "3wwwwwwwwwwwwwwwwwww"
            Top             =   270
            Width           =   870
         End
         Begin VB.TextBox Text1 
            Height          =   285
            Index           =   17
            Left            =   3990
            MaxLength       =   20
            TabIndex        =   41
            Tag             =   "% Interes|N|S|||sparam|porcrete|##0.00||"
            Text            =   "3wwwwwwwwwwwwwwwwwww"
            Top             =   270
            Width           =   870
         End
         Begin VB.Label Label1 
            Caption         =   "% Interés"
            Height          =   195
            Index           =   5
            Left            =   210
            TabIndex        =   44
            Top             =   330
            Width           =   1230
         End
         Begin VB.Label Label1 
            Caption         =   "%Retención"
            Height          =   195
            Index           =   4
            Left            =   2760
            TabIndex        =   43
            Top             =   360
            Width           =   1230
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Datos Contabilidad"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3630
         Left            =   -74820
         TabIndex        =   34
         Top             =   945
         Width           =   7665
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   6
            Left            =   2895
            MaxLength       =   10
            TabIndex        =   5
            Tag             =   "Cta.Retencion|T|S|||sparam|ctareten|||"
            Top             =   2115
            Width           =   1215
         End
         Begin VB.TextBox Text2 
            BackColor       =   &H80000018&
            Enabled         =   0   'False
            Height          =   285
            Index           =   6
            Left            =   4140
            TabIndex        =   49
            Top             =   2115
            Width           =   2895
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   5
            Left            =   2895
            MaxLength       =   10
            TabIndex        =   4
            Tag             =   "Cta.Gastos|T|S|||sparam|ctagasto|||"
            Top             =   1755
            Width           =   1215
         End
         Begin VB.TextBox Text2 
            BackColor       =   &H80000018&
            Enabled         =   0   'False
            Height          =   285
            Index           =   5
            Left            =   4155
            TabIndex        =   48
            Top             =   1755
            Width           =   2895
         End
         Begin VB.TextBox Text2 
            BackColor       =   &H80000018&
            Enabled         =   0   'False
            Height          =   285
            Index           =   14
            Left            =   3600
            Locked          =   -1  'True
            TabIndex        =   47
            Top             =   2490
            Width           =   3450
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   14
            Left            =   2895
            MaxLength       =   10
            TabIndex        =   6
            Tag             =   "Concepto Debe|N|S|||sparam|concedebe|||"
            Top             =   2490
            Width           =   585
         End
         Begin VB.TextBox Text2 
            BackColor       =   &H80000018&
            Enabled         =   0   'False
            Height          =   285
            Index           =   15
            Left            =   3600
            Locked          =   -1  'True
            TabIndex        =   46
            Top             =   2850
            Width           =   3480
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   15
            Left            =   2895
            MaxLength       =   10
            TabIndex        =   7
            Tag             =   "Concepto Haber|N|S|||sparam|concehaber|||"
            Top             =   2865
            Width           =   585
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   16
            Left            =   2895
            MaxLength       =   10
            TabIndex        =   8
            Tag             =   "Numero de Diario|N|S|||sparam|numdiario||000|"
            Top             =   3225
            Width           =   585
         End
         Begin VB.TextBox Text2 
            BackColor       =   &H80000018&
            Enabled         =   0   'False
            Height          =   285
            Index           =   16
            Left            =   3600
            Locked          =   -1  'True
            TabIndex        =   45
            Top             =   3225
            Width           =   3480
         End
         Begin VB.TextBox Text1 
            Height          =   285
            Index           =   1
            Left            =   2565
            MaxLength       =   20
            TabIndex        =   0
            Tag             =   "Servidor Contabilidad|T|S|||sparam|serconta|||"
            Text            =   "3wwwwwwwwwwwwwwwwwwwwwwwwwwwww"
            Top             =   180
            Width           =   4560
         End
         Begin VB.TextBox Text1 
            Height          =   285
            Index           =   0
            Left            =   4230
            MaxLength       =   15
            TabIndex        =   39
            Tag             =   "Código Parámetros Aplic|N|N|||sparam|codparam||S|"
            Text            =   "1"
            Top             =   180
            Width           =   645
         End
         Begin VB.TextBox Text1 
            Height          =   285
            IMEMode         =   3  'DISABLE
            Index           =   3
            Left            =   2565
            MaxLength       =   20
            PasswordChar    =   "*"
            TabIndex        =   2
            Tag             =   "Password Contabilidad|T|S|||sparam|pasconta|||"
            Text            =   "3"
            Top             =   840
            Width           =   4560
         End
         Begin VB.TextBox Text1 
            Height          =   285
            Index           =   2
            Left            =   2565
            MaxLength       =   20
            TabIndex        =   1
            Tag             =   "Usuario Contabilidad|T|S|||sparam|usuconta|||"
            Text            =   "3wwwwwwwwwwwwwwwwwww"
            Top             =   540
            Width           =   4560
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   4
            Left            =   2565
            MaxLength       =   2
            TabIndex        =   3
            Tag             =   "Nº Contabilidad|N|S|||sparam|numconta|||"
            Text            =   "3"
            Top             =   1170
            Width           =   780
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   1
            Left            =   2565
            ToolTipText     =   "Buscar Cta.Contable"
            Top             =   2115
            Width           =   240
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   0
            Left            =   2565
            ToolTipText     =   "Buscar Cta.Contable"
            Top             =   1755
            Width           =   240
         End
         Begin VB.Label Label1 
            Caption         =   "Cta.Contable Gastos"
            Height          =   195
            Index           =   24
            Left            =   495
            TabIndex        =   54
            Top             =   1785
            Width           =   1920
         End
         Begin VB.Label Label1 
            Caption         =   "Cta.Contable Retención"
            Height          =   195
            Index           =   25
            Left            =   495
            TabIndex        =   53
            Top             =   2115
            Width           =   1920
         End
         Begin VB.Label Label1 
            Caption         =   "Concepto al Haber"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   1
            Left            =   495
            TabIndex        =   52
            Top             =   2850
            Width           =   1410
         End
         Begin VB.Label Label1 
            Caption         =   "Concepto al Debe"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   2
            Left            =   495
            TabIndex        =   51
            Top             =   2490
            Width           =   1350
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   2
            Left            =   2565
            ToolTipText     =   "Buscar Concepto"
            Top             =   2490
            Width           =   240
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   3
            Left            =   2565
            ToolTipText     =   "Buscar Concepto"
            Top             =   2850
            Width           =   240
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   4
            Left            =   2565
            ToolTipText     =   "Buscar Diario"
            Top             =   3225
            Width           =   240
         End
         Begin VB.Label Label1 
            Caption         =   "Numero de Diario"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   3
            Left            =   495
            TabIndex        =   50
            Top             =   3225
            Width           =   1410
         End
         Begin VB.Label Label1 
            Caption         =   "Password"
            Height          =   195
            Index           =   15
            Left            =   510
            TabIndex        =   38
            Top             =   900
            Width           =   840
         End
         Begin VB.Label Label1 
            Caption         =   "Usuario"
            Height          =   195
            Index           =   17
            Left            =   510
            TabIndex        =   37
            Top             =   600
            Width           =   780
         End
         Begin VB.Label Label1 
            Caption         =   "Nº conta"
            Height          =   195
            Index           =   18
            Left            =   510
            TabIndex        =   36
            Top             =   1230
            Width           =   900
         End
         Begin VB.Label Label1 
            Caption         =   "Servidor"
            Height          =   195
            Index           =   19
            Left            =   510
            TabIndex        =   35
            Top             =   270
            Width           =   900
         End
      End
      Begin VB.Label Label1 
         Caption         =   "Extensión"
         Height          =   195
         Index           =   86
         Left            =   330
         TabIndex        =   239
         Top             =   1845
         Width           =   1605
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   38
         Left            =   2145
         ToolTipText     =   "Buscar Extensión"
         Top             =   1845
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Carpeta Facturas Varias"
         Height          =   195
         Index           =   85
         Left            =   315
         TabIndex        =   235
         Top             =   1395
         Width           =   1830
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   37
         Left            =   2160
         ToolTipText     =   "Buscar Carpeta"
         Top             =   1395
         Width           =   240
      End
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   7845
      TabIndex        =   165
      Top             =   6765
      Width           =   1035
   End
   Begin VB.Frame Frame1 
      Height          =   540
      Left            =   240
      TabIndex        =   31
      Top             =   6615
      Width           =   3000
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
         Left            =   120
         TabIndex        =   32
         Top             =   210
         Width           =   2760
      End
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   6705
      TabIndex        =   164
      Top             =   6750
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   375
      Left            =   7830
      TabIndex        =   29
      Top             =   6750
      Width           =   1035
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   30
      Top             =   0
      Width           =   9045
      _ExtentX        =   15954
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   5
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Añadir"
            Object.Tag             =   "1"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Modificar"
            Object.Tag             =   "2"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Salir"
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc Data1 
      Height          =   330
      Left            =   3630
      Top             =   5250
      Visible         =   0   'False
      Width           =   1455
      _ExtentX        =   2566
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
   Begin VB.Menu mnOpciones 
      Caption         =   "&Opciones"
      Begin VB.Menu mnAñadir 
         Caption         =   "&Añadir"
         Shortcut        =   ^A
      End
      Begin VB.Menu mnModificar 
         Caption         =   "&Modificar"
         HelpContextID   =   1
         Shortcut        =   ^M
      End
      Begin VB.Menu mnBarra1 
         Caption         =   "-"
      End
      Begin VB.Menu mnSalir 
         Caption         =   "&Salir"
         Shortcut        =   ^S
      End
   End
End
Attribute VB_Name = "frmConfParamAplic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' ### [Monica] 06/09/2006
' procedimiento nuevo introducido de la gestion

Option Explicit

'Private WithEvents frmMtoArt As frmAlmArticulos
Private WithEvents frmCtas As frmCtasConta
Attribute frmCtas.VB_VarHelpID = -1
Private WithEvents frmConce As frmConceConta
Attribute frmConce.VB_VarHelpID = -1
Private WithEvents frmTDia As frmDiaConta
Attribute frmTDia.VB_VarHelpID = -1
Private WithEvents frmTIva As frmTipIVAConta 'Tipos de Iva Contables
Attribute frmTIva.VB_VarHelpID = -1
Private WithEvents frmFPa As frmForpaConta 'formas de pago de la contabilidad
Attribute frmFPa.VB_VarHelpID = -1


Private WithEvents frmExt As frmExtAridoc1
Attribute frmExt.VB_VarHelpID = -1
Private WithEvents frmAri As frmCarpAridoc
Attribute frmAri.VB_VarHelpID = -1

Private NombreTabla As String  'Nombre de la tabla o de la
Private Ordenacion As String
Private CadenaConsulta As String

Dim indice As Byte
Dim indCodigo As Byte
Dim Encontrado As Boolean
Dim Modo As Byte
'0: Inicial
'2: Visualizacion
'3: Añadir
'4: Modificar

Private Sub chkOutlook_GotFocus()
    ConseguirfocoChk Modo
End Sub

Private Sub chkOutlook_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub cmdAceptar_Click()
Dim actualiza As Boolean
Dim kms As Currency
    
    If Modo = 3 Then
        If DatosOk Then
            'Cambiamos el path
            'CambiaPath True
            If InsertarDesdeForm(Me) Then
                PonerModo 0
'                ActualizaNombreEmpresa
                MsgBox "Debe salir de la aplicacion para que los cambios tengan efecto", vbExclamation
            End If

        End If
    End If


    If Modo = 4 Then 'MODIFICAR
        If DatosOk Then
            If Not vParamAplic Is Nothing Then
                'Datos contabilidad
                vParamAplic.ServidorConta = Text1(1).Text
                vParamAplic.UsuarioConta = Text1(2).Text
                vParamAplic.PasswordConta = Text1(3).Text
                vParamAplic.NumeroConta = ComprobarCero(Text1(4).Text)
                
                vParamAplic.CtaGasto = Text1(5).Text
                vParamAplic.CtaReten = Text1(6).Text
                vParamAplic.Porcinte = Text1(18).Text
                vParamAplic.Porcrete = Text1(17).Text
                
                vParamAplic.DireMail = Text1(9).Text
                vParamAplic.Smtphost = Text1(10).Text
                vParamAplic.SmtpUser = Text1(11).Text
                vParamAplic.Smtppass = Text1(12).Text
                vParamAplic.WebSoporte = Text1(13).Text
                
                vParamAplic.EnvioDesdeOutlook = Me.chkOutlook.Value
            
                ' Para utilizar el arigesmail
                vParamAplic.ExeEnvioMail = Trim(Text1(60).Text)
                
                
                vParamAplic.ConceDebe = Text1(14).Text
                vParamAplic.ConceHaber = Text1(15).Text
                vParamAplic.NumDiario = Text1(16).Text
                
                'parametros para seguros
                vParamAplic.ServidorContaSeg = Text1(22).Text
                vParamAplic.UsuarioContaSeg = Text1(24).Text
                vParamAplic.PasswordContaSeg = Text1(23).Text
                vParamAplic.NumeroContaSeg = ComprobarCero(Text1(25).Text)
                vParamAplic.CtaBancoSeg = Text1(8).Text
                vParamAplic.RaizCtaSocSeg = Text1(7).Text
                
                vParamAplic.ConceDebeSeg = ComprobarCero(Text1(19).Text)
                vParamAplic.ConceHaberSeg = ComprobarCero(Text1(20).Text)
                vParamAplic.NumDiarioSeg = ComprobarCero(Text1(21).Text)
                
                'parametros para telefonia
                vParamAplic.ServidorContaTel = Text1(29).Text
                vParamAplic.UsuarioContaTel = Text1(27).Text
                vParamAplic.PasswordContaTel = Text1(28).Text
                vParamAplic.NumeroContaTel = ComprobarCero(Text1(26).Text)
                
                vParamAplic.ConceDebeTel = ComprobarCero(Text1(32).Text)
                vParamAplic.ConceHaberTel = ComprobarCero(Text1(31).Text)
                vParamAplic.NumDiarioTel = ComprobarCero(Text1(30).Text)
                vParamAplic.CtaBancoTel = Text1(34).Text
                vParamAplic.RaizCtaSocTel = Text1(33).Text
                vParamAplic.CtaVentaTel = Text1(35).Text
                vParamAplic.NumSerieTel = Trim(Text1(36).Text)
                
                vParamAplic.TipoFicheroTel = ComprobarCero(Combo1(0).ListIndex)
                
                
                'parametros para facturas varias
                vParamAplic.ServidorContaFac = Text1(37).Text
                vParamAplic.UsuarioContaFac = Text1(38).Text
                vParamAplic.PasswordContaFac = Text1(39).Text
                
                
                'parametros para gasolinera
                vParamAplic.ServidorContaGas = Text1(47).Text
                vParamAplic.UsuarioContaGas = Text1(49).Text
                vParamAplic.PasswordContaGas = Text1(48).Text
                vParamAplic.NumeroContaGas = ComprobarCero(Text1(50).Text)
                vParamAplic.CtaVentasGas = Text1(42).Text
                vParamAplic.RaizCtaSocGas = Text1(43).Text
                vParamAplic.CtaContraGas = Text1(41).Text
                vParamAplic.NumSerieGas = Trim(Text1(40).Text)
                '[Monica]04/07/2013: metemos la letra de serie de gasoleo B
                vParamAplic.NumSerieGasB = Trim(Text1(82).Text)
                vParamAplic.IncreFactGas = ComprobarCero(Text1(51).Text)
                vParamAplic.CodIvaGas = ComprobarCero(Text1(52).Text)
                
                vParamAplic.ConceDebeGas = ComprobarCero(Text1(44).Text)
                vParamAplic.ConceHaberGas = ComprobarCero(Text1(45).Text)
                vParamAplic.NumDiarioGas = ComprobarCero(Text1(46).Text)
                
                'parametros para facturas socios
                vParamAplic.ServidorContaFacSoc = Text1(57).Text
                vParamAplic.UsuarioContaFacSoc = Text1(55).Text
                vParamAplic.PasswordContaFacSoc = Text1(56).Text
                vParamAplic.NumeroContaFacSoc = ComprobarCero(Text1(54).Text)
                vParamAplic.CtaRetenFacSoc = Text1(59).Text
                vParamAplic.RaizCtaFacSoc = Text1(53).Text
                vParamAplic.CtaContraGas = Text1(41).Text
                vParamAplic.PorcreteFacSoc = Text1(58).Text
                
                'parametros para facturas coarval
                vParamAplic.ServidorContaCV = Text1(65).Text
                vParamAplic.UsuarioContaCV = Text1(66).Text
                vParamAplic.PasswordContaCV = Text1(67).Text
                vParamAplic.NumeroContaCV = ComprobarCero(Text1(68).Text)
                
                vParamAplic.CtaBancoCV = Text1(63).Text
                vParamAplic.LetraSerCV = Text1(61).Text
                vParamAplic.LetraSerFCV = Text1(75).Text
                vParamAplic.LetraSerFinCV = Text1(76).Text
                vParamAplic.RaizCtaSocCV = Text1(64).Text
                vParamAplic.CtaVentaCV = Text1(62).Text
                
                vParamAplic.CtaVentaTickCV = Text1(69).Text
                vParamAplic.CtaClienTickCV = Text1(74).Text
                vParamAplic.CtaVentaFacCV = Text1(70).Text
                vParamAplic.CtaVentaFacInCV = Text1(73).Text
                vParamAplic.CodforpaBanCV = Text1(71).Text
                vParamAplic.CodforpaConCV = Text1(72).Text
                
                'varias
                vParamAplic.NumeroContaCVV = ComprobarCero(Text1(77).Text)
                
                vParamAplic.CtaBancoCVV = Text1(81).Text
                vParamAplic.LetraSerCVV = Text1(78).Text
                
                vParamAplic.CodforpaBanCVV = Text1(79).Text
                vParamAplic.CodforpaConCVV = Text1(80).Text
                
                'aridoc
                vParamAplic.CarpetaFac = Text1(83)
                vParamAplic.Extension = Text1(84)
                vParamAplic.C1Factura = Combo1(5).ListIndex
                vParamAplic.C2Factura = Combo1(6).ListIndex
                vParamAplic.C3Factura = Combo1(7).ListIndex
                vParamAplic.C4Factura = Combo1(8).ListIndex
                
                
                actualiza = vParamAplic.Modificar(Text1(0).Text)
                TerminaBloquear
    
                If actualiza Then  'Inserta o Modifica
                    'Abrir la conexion a la conta q hemos modificado
                    AccionesCerrarContabilidades
                    LeerDatosEmpresa
                    ComprobarDatos
'                    CerrarConexionConta
'                    CerrarConexionContaSeg
'                    CerrarConexionContaTel
'                    'nota:la conexion de facturas varias se abrira y cerrara en el momento
'                    If vParamAplic.Avnics = 1 Then
'                        If vParamAplic.NumeroConta <> 0 Then
'                            If Not AbrirConexionConta(vParamAplic.UsuarioConta, vParamAplic.PasswordConta) Then End
'                        End If
'                    End If
'                    If vParamAplic.Seguros = 1 Then
'                        If vParamAplic.NumeroContaSeg <> 0 Then
'                            If Not AbrirConexionContaSeg(vParamAplic.UsuarioContaSeg, vParamAplic.PasswordContaSeg) Then End
'                        End If
'                    End If
'                    If vParamAplic.Telefonia = 1 Then
'                        If vParamAplic.NumeroContaTel <> 0 Then
'                            If Not AbrirConexionContaTel(vParamAplic.UsuarioContaTel, vParamAplic.PasswordContaTel) Then End
'                        End If
'                    End If
                    PonerModo 2
                    PonerFocoBtn Me.cmdSalir
                End If
           End If
        End If
    End If
End Sub

Private Sub cmdCancelar_Click()
    TerminaBloquear
    If Data1.Recordset.EOF Then
        PonerModo 0
    Else
        PonerCampos
        PonerModo 2
    End If
End Sub

Private Sub cmdSalir_Click()
    Screen.MousePointer = vbDefault
    Unload Me
End Sub

Private Sub Form_Activate()
    Screen.MousePointer = vbDefault
    If Modo = 0 Then PonerCadenaBusqueda
    PonerFoco Text1(0)
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub Form_Load()
Dim i As Byte
    'Icono del formulario
    Me.Icon = frmPpal.Icon

    ' ICONITOS DE LA BARRA
    With Me.Toolbar1
        .HotImageList = frmPpal.imgListComun_OM
        .DisabledImageList = frmPpal.imgListComun_BN
        .ImageList = frmPpal.imgListComun
        .Buttons(1).Image = 3   'Anyadir
        .Buttons(2).Image = 4   'Modificar
        .Buttons(5).Image = 11  'Salir
    End With
    
    CargaCombo
    
    LimpiarCampos   'Limpia los campos TextBox
   
   'cargar IMAGES de busqueda
    For i = 0 To Me.imgBuscar.Count - 1
        Me.imgBuscar(i).Picture = frmPpal.imgListImages16.ListImages(1).Picture
    Next i

    SSTab1.Tab = 0

    NombreTabla = "sparam"
    Ordenacion = " ORDER BY codparam"
    'ASignamos un SQL al DATA1
    Data1.ConnectionString = conn
    CadenaConsulta = "Select * from " & NombreTabla & Ordenacion
    Data1.RecordSource = CadenaConsulta
    Data1.Refresh
    
    Encontrado = True
    If Data1.Recordset.EOF Then
        'No hay registro de datos de parametros
        'quitar###
        Encontrado = False
    End If
    
    Me.SSTab1.TabEnabled(0) = (vParamAplic.Avnics = 1)
    Me.SSTab1.TabVisible(0) = (vParamAplic.Avnics = 1)
    Me.SSTab1.TabEnabled(1) = (vParamAplic.Seguros = 1)
    Me.SSTab1.TabVisible(1) = (vParamAplic.Seguros = 1)
    Me.SSTab1.TabEnabled(2) = (vParamAplic.Telefonia = 1)
    Me.SSTab1.TabVisible(2) = (vParamAplic.Telefonia = 1)
    Me.SSTab1.TabEnabled(3) = (vParamAplic.FacturasVarias = 1)
    Me.SSTab1.TabVisible(3) = (vParamAplic.FacturasVarias = 1)
    Me.SSTab1.TabEnabled(5) = (vParamAplic.Gasolinera = 1)
    Me.SSTab1.TabVisible(5) = (vParamAplic.Gasolinera = 1)
    Me.SSTab1.TabEnabled(6) = (vParamAplic.FactSocios = 1)
    Me.SSTab1.TabVisible(6) = (vParamAplic.FactSocios = 1)
    Me.SSTab1.TabEnabled(7) = (vParamAplic.Coarval = 1)
    Me.SSTab1.TabVisible(7) = (vParamAplic.Coarval = 1)
    
    
    '[Monica]12/11/2013: si hay aridoc abrimos conexion
    Me.SSTab1.TabEnabled(9) = (vParamAplic.HayAridoc = 1)
    Me.SSTab1.TabVisible(9) = (vParamAplic.HayAridoc = 1)
    If (vParamAplic.HayAridoc = 1) Then
        AbrirConexionAridoc "root", "aritel"
    End If
    
    PonerModo 0

End Sub

Private Sub PonerCadenaBusqueda()
On Error GoTo EEPonerBusq
    Screen.MousePointer = vbHourglass

    Data1.RecordSource = CadenaConsulta
    Data1.Refresh
    If Data1.Recordset.RecordCount <= 0 Then
        LimpiarCampos
        MsgBox "No hay ningún registro en la tabla " & NombreTabla, vbInformation
        Screen.MousePointer = vbDefault
'        Me.Toolbar1.Buttons(1).Enabled = False 'Modificar
    Else
        Data1.Recordset.MoveFirst
        PonerCampos
    End If
    
    Screen.MousePointer = vbDefault
    Exit Sub
EEPonerBusq:
    MuestraError Err.Number, "PonerCadenaBusqueda"
    Screen.MousePointer = vbDefault
End Sub

Private Sub frmAri_DatoSeleccionado(CadenaSeleccion As String)
Dim Cad As String
    Cad = RecuperaValor(CadenaSeleccion, 1)
    Text1(indice).Text = Mid(Cad, 2, Len(Cad))
    Text1(indice).Text = Format(Text1(indice).Text, "000")
    Text2(indice).Text = RecuperaValor(CadenaSeleccion, 3)
End Sub

Private Sub frmConce_DatoSeleccionado(CadenaSeleccion As String)
'Form de Consulta de Clientes
    Text1(indCodigo).Text = Format(RecuperaValor(CadenaSeleccion, 1), "000")
    Text2(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmCtas_DatoSeleccionado(CadenaSeleccion As String)
'Cuentas contables de la Contabilidad
    Text1(indice).Text = RecuperaValor(CadenaSeleccion, 1) 'codmacta
    Text2(indice).Text = RecuperaValor(CadenaSeleccion, 2) 'des macta
End Sub

Private Sub frmExt_DatoSeleccionado(CadenaSeleccion As String)
'Extension de Aridoc
    Text1(indice).Text = RecuperaValor(CadenaSeleccion, 1) 'codigo
    Text2(indice).Text = RecuperaValor(CadenaSeleccion, 2) 'descripcion
End Sub

Private Sub frmFpa_DatoSeleccionado(CadenaSeleccion As String)
    Text1(indice).Text = RecuperaValor(CadenaSeleccion, 1) 'codforpa
    FormateaCampo Text1(indice)
    Text2(indice).Text = RecuperaValor(CadenaSeleccion, 2) 'nomforpa
End Sub

Private Sub frmTDia_DatoSeleccionado(CadenaSeleccion As String)
'Form de Consulta de tipos de diario
    Text1(indCodigo).Text = Format(RecuperaValor(CadenaSeleccion, 1), "000")
    Text2(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmTIva_DatoSeleccionado(CadenaSeleccion As String)
'Form de Consulta de tipos de iva
    Text1(indCodigo).Text = RecuperaValor(CadenaSeleccion, 1) 'tipos de iva contable
    Text2(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2) 'nombre tipos de iva
End Sub

Private Sub imgBuscar_Click(Index As Integer)
Dim numNivel As Byte

    If vParamAplic.NumeroConta = 0 Then Exit Sub
    
    Select Case Index
        ' **************avnics
        Case 0, 1 'Cuentas Contables (de contabilidad)
            indice = Index + 5
            Set frmCtas = New frmCtasConta
            frmCtas.NumDigit = 0
            frmCtas.DatosADevolverBusqueda = "0|1|"
            frmCtas.CodigoActual = Text1(indice).Text
            frmCtas.Facturas = 0
            frmCtas.Conexion = cConta
            frmCtas.Show vbModal
            Set frmCtas = Nothing
            PonerFoco Text1(indice)
        
        Case 2, 3 ' Conceptos Contables
            AbrirFrmConceptos Index + 12, cConta
        
        Case 4 ' Numero de Diario
            AbrirFrmDiario Index + 12, cConta
            
        '**************seguros
        Case 6 'Cuentas Contables (de contabilidad)
            indice = Index + 2
            Set frmCtas = New frmCtasConta
            frmCtas.NumDigit = 0
            frmCtas.DatosADevolverBusqueda = "0|1|"
            frmCtas.CodigoActual = Text1(indice).Text
            frmCtas.Facturas = 0
            frmCtas.Conexion = cContaSeg
            frmCtas.Show vbModal
            Set frmCtas = Nothing
            PonerFoco Text1(indice)
        
        Case 5 'raices de las cuentas contables de socio
            indice = Index + 2
            Set frmCtas = New frmCtasConta
            numNivel = DevuelveDesdeBDNew(cContaSeg, "empresa", "numnivel", "", "", "")
            frmCtas.NumDigit = vEmpresaSeg.DigitosNivelAnterior  'DevuelveDesdeBDNew(cContaSeg, "empresa", "numdigi" & numNivel - 1, "", "", "")
            frmCtas.DatosADevolverBusqueda = "0|1|"
            frmCtas.CodigoActual = Text1(indice).Text
            frmCtas.Facturas = 0
            frmCtas.Conexion = cContaSeg
            frmCtas.Show vbModal
            Set frmCtas = Nothing
            PonerFoco Text1(indice)
        
        Case 7, 8 ' Conceptos Contables
            AbrirFrmConceptos Index + 12, cContaSeg
        
        Case 9 ' Numero de Diario
            AbrirFrmDiario Index + 12, cContaSeg
        
        '**************telefonia
        Case 11, 12 ' Conceptos Contables
            AbrirFrmConceptos Index + 20, cContaSeg
        
        Case 10 ' Numero de Diario
            AbrirFrmDiario Index + 20, cContaSeg
        
        Case 13 'raices de las cuentas contables de socio
            indice = Index + 20
            Set frmCtas = New frmCtasConta
            numNivel = DevuelveDesdeBDNew(cContaTel, "empresa", "numnivel", "", "", "")
            frmCtas.NumDigit = vEmpresaTel.DigitosNivelAnterior 'DevuelveDesdeBDNew(cContaTel, "empresa", "numdigi" & numNivel - 1, "", "", "")
            frmCtas.DatosADevolverBusqueda = "0|1|"
            frmCtas.CodigoActual = Text1(indice).Text
            frmCtas.Facturas = 0
            frmCtas.Conexion = cContaTel
            frmCtas.Show vbModal
            Set frmCtas = Nothing
            PonerFoco Text1(indice)
        
        Case 14, 15 'Cuentas Contables (de contabilidad)
            indice = Index + 20
            Set frmCtas = New frmCtasConta
            frmCtas.NumDigit = 0
            frmCtas.DatosADevolverBusqueda = "0|1|"
            frmCtas.CodigoActual = Text1(indice).Text
            frmCtas.Facturas = 0
            frmCtas.Conexion = cContaTel
            frmCtas.Show vbModal
            Set frmCtas = Nothing
            PonerFoco Text1(indice)
        
        ' gasolinera
        Case 16, 17 'Cuentas Contables (de contabilidad)
            indice = Index + 25
            Set frmCtas = New frmCtasConta
            frmCtas.NumDigit = 0
            frmCtas.DatosADevolverBusqueda = "0|1|"
            frmCtas.CodigoActual = Text1(indice).Text
            frmCtas.Facturas = 0
            frmCtas.Conexion = cContaGas
            frmCtas.Show vbModal
            Set frmCtas = Nothing
            PonerFoco Text1(indice)
        
        Case 19, 20 ' Conceptos Contables
            AbrirFrmConceptos Index + 25, cContaGas
        
        Case 21 ' Numero de Diario
            AbrirFrmDiario Index + 25, cContaGas
        
        Case 18 'raices de las cuentas contables de socio
            indice = Index + 25
            Set frmCtas = New frmCtasConta
            numNivel = DevuelveDesdeBDNew(cContaGas, "empresa", "numnivel", "", "", "")
            frmCtas.NumDigit = vEmpresaGas.DigitosNivelAnterior 'DevuelveDesdeBDNew(cContaGas, "empresa", "numdigi" & numNivel - 1, "", "", "")
            frmCtas.DatosADevolverBusqueda = "0|1|"
            frmCtas.CodigoActual = Text1(indice).Text
            frmCtas.Facturas = 0
            frmCtas.Conexion = cContaGas
            frmCtas.Show vbModal
            Set frmCtas = Nothing
            PonerFoco Text1(indice)
        
        Case 22 ' codigo de iva DE GASOLINERA
            indice = Index + 30
            AbrirFrmTipIvaConta (indice)
        
        
        ' facturas socios
        Case 24 'Cuentas Contables (de contabilidad)
            indice = Index + 35
            Set frmCtas = New frmCtasConta
            frmCtas.NumDigit = 0
            frmCtas.DatosADevolverBusqueda = "0|1|"
            frmCtas.CodigoActual = Text1(indice).Text
            frmCtas.Facturas = 0
            frmCtas.Conexion = cContaFacSoc
            frmCtas.Show vbModal
            Set frmCtas = Nothing
            PonerFoco Text1(indice)
        
        Case 23 'raices de las cuentas contables de socio
            indice = Index + 30
            Set frmCtas = New frmCtasConta
            numNivel = DevuelveDesdeBDNew(cContaFacSoc, "empresa", "numnivel", "", "", "")
            frmCtas.NumDigit = vEmpresaFacSoc.DigitosNivelAnterior 'DevuelveDesdeBDNew(cContaGas, "empresa", "numdigi" & numNivel - 1, "", "", "")
            frmCtas.DatosADevolverBusqueda = "0|1|"
            frmCtas.CodigoActual = Text1(indice).Text
            frmCtas.Facturas = 0
            frmCtas.Conexion = cContaFacSoc
            frmCtas.Show vbModal
            Set frmCtas = Nothing
            PonerFoco Text1(indice)
    
    
        '**************facturas coarval
        Case 25, 26 'Cuentas Contables (de contabilidad)
            indice = Index + 37
            Set frmCtas = New frmCtasConta
            frmCtas.NumDigit = 0
            frmCtas.DatosADevolverBusqueda = "0|1|"
            frmCtas.CodigoActual = Text1(indice).Text
            frmCtas.Facturas = 0
            frmCtas.Conexion = cContaCV
            frmCtas.Show vbModal
            Set frmCtas = Nothing
            PonerFoco Text1(indice)
        
        Case 27 'raices de la cuenta contable de socio
            indice = Index + 5
            Set frmCtas = New frmCtasConta
            numNivel = DevuelveDesdeBDNew(cContaCV, "empresa", "numnivel", "", "", "")
            frmCtas.NumDigit = vEmpresaCV.DigitosNivelAnterior  'DevuelveDesdeBDNew(cContaSeg, "empresa", "numdigi" & numNivel - 1, "", "", "")
            frmCtas.DatosADevolverBusqueda = "0|1|"
            frmCtas.CodigoActual = Text1(indice).Text
            frmCtas.Facturas = 0
            frmCtas.Conexion = cContaCV
            frmCtas.Show vbModal
            Set frmCtas = Nothing
            PonerFoco Text1(indice)
    
        Case 28, 29, 32, 33 'Cuentas Contables (de contabilidad)
            indice = Index + 41
            Set frmCtas = New frmCtasConta
            frmCtas.NumDigit = 0
            frmCtas.DatosADevolverBusqueda = "0|1|"
            frmCtas.CodigoActual = Text1(indice).Text
            frmCtas.Facturas = 0
            frmCtas.Conexion = cContaCV
            frmCtas.Show vbModal
            Set frmCtas = Nothing
            PonerFoco Text1(indice)
    
        Case 30, 31 ' forma de pago
            indice = Index + 41
            Set frmFPa = New frmForpaConta
            frmFPa.DatosADevolverBusqueda = "0|1|"
            frmFPa.CodigoActual = Text1(indice)
            frmFPa.Conexion = cContaCV
            frmFPa.Show vbModal
            Set frmFPa = Nothing
            PonerFoco Text1(indice)
                
        Case 36 'Cuentas Contables (de contabilidad)
            indice = Index + 45
            Set frmCtas = New frmCtasConta
            frmCtas.NumDigit = 0
            frmCtas.DatosADevolverBusqueda = "0|1|"
            frmCtas.CodigoActual = Text1(indice).Text
            frmCtas.Facturas = 0
            frmCtas.Conexion = cContaCVV
            frmCtas.Show vbModal
            Set frmCtas = Nothing
            PonerFoco Text1(indice)
        
        Case 34, 35 ' forma de pago
            indice = Index + 45
            Set frmFPa = New frmForpaConta
            frmFPa.DatosADevolverBusqueda = "0|1|"
            frmFPa.CodigoActual = Text1(indice)
            frmFPa.Conexion = cContaCVV
            frmFPa.Show vbModal
            Set frmFPa = Nothing
            PonerFoco Text1(indice)
            
        '[Monica]12/11/2013: Aridoc
        Case 37 ' carpeta de facturas varias
            Select Case Index
                Case 37
                    indice = Index + 46
            End Select
            
            Set frmAri = New frmCarpAridoc
            frmAri.Opcion = 20
            frmAri.Show vbModal
            Set frmAri = Nothing
            PonerFoco Text1(indice)
            
        Case 38 ' extension
            indice = Index + 46
            Set frmExt = New frmExtAridoc1
            frmExt.DatosADevolverBusqueda = "0|1|"
            frmExt.CodigoActual = Text1(indice).Text
            frmExt.Show vbModal
            Set frmExt = Nothing
            PonerFoco Text1(indice)
        
    End Select
End Sub




Private Sub mnAñadir_Click()
    If BLOQUEADesdeFormulario(Me) Then BotonAnyadir
End Sub

Private Sub mnModificar_Click()
    If BLOQUEADesdeFormulario(Me) Then BotonModificar
End Sub

Private Sub mnSalir_Click()
    Screen.MousePointer = vbDefault
    Unload Me
End Sub

Private Sub Text1_GotFocus(Index As Integer)
    ConseguirFoco Text1(Index), Modo
End Sub

Private Sub Text1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
'Avanzar/Retroceder los campos con las flechas de desplazamiento del teclado.
    KEYdown KeyCode
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
'14/02/2007 antes estaba esto
'    KEYpress (KeyAscii)
' ahora
    If KeyAscii = teclaBuscar Then
        Select Case Index
            Case 0: KEYBusqueda KeyAscii, 0 'cuenta contable gastos
            Case 1: KEYBusqueda KeyAscii, 1 'cuenta contable retencion
            
            Case 2: KEYBusqueda KeyAscii, 2 'concepto al debe
            Case 3: KEYBusqueda KeyAscii, 3 'concepto al haber
            Case 4: KEYBusqueda KeyAscii, 4 'diario
            
            Case 19: KEYBusqueda KeyAscii, 19 'concepto al debe
            Case 20: KEYBusqueda KeyAscii, 20 'concepto al haber
            Case 21: KEYBusqueda KeyAscii, 21 'diario
            Case 8: KEYBusqueda KeyAscii, 8 'cuenta contable banco
            Case 7: KEYBusqueda KeyAscii, 7 'raiz cuenta contable socio
            
            Case 11: KEYBusqueda KeyAscii, 11 'concepto al debe
            Case 12: KEYBusqueda KeyAscii, 12 'concepto al haber
            Case 10: KEYBusqueda KeyAscii, 10 'diario
        
            Case 11: KEYBusqueda KeyAscii, 11 'concepto al debe
            Case 12: KEYBusqueda KeyAscii, 12 'concepto al haber
            Case 10: KEYBusqueda KeyAscii, 10 'diario
            
            Case 19: KEYBusqueda KeyAscii, 19 'concepto al debe
            Case 20: KEYBusqueda KeyAscii, 20 'concepto al haber
            Case 21: KEYBusqueda KeyAscii, 21 'diario
            Case 16: KEYBusqueda KeyAscii, 16 'cuenta contable ventas
            
            Case 17: KEYBusqueda KeyAscii, 17 'cuenta contable contrapartida
            Case 18: KEYBusqueda KeyAscii, 18 'raiz cuenta contable
            Case 22: KEYBusqueda KeyAscii, 22 'codigo de iva
        
            'facturas coarval
            Case 62: KEYBusqueda KeyAscii, 25 'cuenta contable venta
            Case 63: KEYBusqueda KeyAscii, 26 'cuenta contable banco
            Case 64: KEYBusqueda KeyAscii, 27 'raiz cuenta contable
            Case 69: KEYBusqueda KeyAscii, 28 'cuenta contable venta ticket
            Case 70: KEYBusqueda KeyAscii, 29 'cuenta contable venta factura
            Case 73: KEYBusqueda KeyAscii, 32 'cuenta contable venta factura interna
            Case 74: KEYBusqueda KeyAscii, 33 'cuenta contable cliente ticket
            Case 71: KEYBusqueda KeyAscii, 30 'forma de pago de banco
            Case 72: KEYBusqueda KeyAscii, 31 'forma de pago de contado
        
            Case 81: KEYBusqueda KeyAscii, 36 'cuenta contable banco
            Case 79: KEYBusqueda KeyAscii, 34 'forma de pago de banco
            Case 80: KEYBusqueda KeyAscii, 35 'forma de pago de contado
        
        End Select
    Else
        KEYpress KeyAscii
    End If
End Sub

Private Sub KEYBusqueda(KeyAscii As Integer, indice As Integer)
    KeyAscii = 0
    imgBuscar_Click (indice)
End Sub

Private Sub Text1_LostFocus(Index As Integer)
Dim cadMen As String
Dim Cad As String

'    If Text1(Index).Text = "" Then Exit Sub

    'Quitar espacios en blanco
    Text1(Index).Text = Trim(Text1(Index).Text)
    
    Select Case Index
        Case 4, 25, 26, 50, 54 'numero Conta
            If Not EsNumerico(Text1(Index).Text) Then
                Text1(Index).Text = ""
                PonerFoco Text1(Index)
'            Else
'                cmdAceptar_Click
            End If
        
        Case 17, 18 'porcentajes de avnics
            If Text1(Index).Text <> "" Then PonerFormatoDecimal Text1(Index), 7
            
        Case 58 'porcentajes de retencion de facturas socios
            If Text1(Index).Text <> "" Then PonerFormatoDecimal Text1(Index), 7
            
        Case 5, 6 ' cuentas contables de avnics
            If Text1(Index).Text = "" Then Exit Sub
            Text2(Index).Text = PonerNombreCuenta(Text1(Index), Modo, , cConta)
            
        Case 8 ' cuentas contables de seguros
            If Text1(Index).Text = "" Then Exit Sub
            Text2(Index).Text = PonerNombreCuenta(Text1(Index), Modo, , cContaSeg)
            
        Case 34, 35 ' cuentas contables de telefonia
            If Text1(Index).Text = "" Then Exit Sub
            Text2(Index).Text = PonerNombreCuenta(Text1(Index), Modo, , cContaTel)
        
        Case 41, 42 ' cuentas contables de gasolinera
            If Text1(Index).Text = "" Then Exit Sub
            Text2(Index).Text = PonerNombreCuenta(Text1(Index), Modo, , cContaGas)
        
        Case 59 ' cuentas contables de facturas socios
            If Text1(Index).Text = "" Then Exit Sub
            Text2(Index).Text = PonerNombreCuenta(Text1(Index), Modo, , cContaFacSoc)
        
        Case 62, 63 ' cuentas contables de facturas coarval
            If Text1(Index).Text = "" Then Exit Sub
            Text2(Index).Text = PonerNombreCuenta(Text1(Index), Modo, , cContaCV)
        
        Case 69, 70, 73, 74 ' cuentas de venta para tickets y para factura
            If Text1(Index).Text = "" Then Exit Sub
            Text2(Index).Text = PonerNombreCuenta(Text1(Index), Modo, , cContaCV)
        
        Case 81 ' cuentas contables de facturas coarval varias
            If Text1(Index).Text = "" Then Exit Sub
            Text2(Index).Text = PonerNombreCuenta(Text1(Index), Modo, , cContaCVV)
        
        Case 16 ' NUMERO DE DIARIO
            If Text1(Index).Text <> "" Then
                Text2(Index).Text = ""
                Text2(Index).Text = DevuelveDesdeBDNew(cConta, "tiposdiario", "desdiari", "numdiari", Text1(Index).Text, "N")
                If Text2(Index).Text = "" Then
                    MsgBox "Número de Diario no existe en la contabilidad. Reintroduzca.", vbExclamation
                End If
            End If
        
        Case 21 ' NUMERO DE DIARIO
            If Text1(Index).Text <> "" Then
                Text2(Index).Text = ""
                Text2(Index).Text = DevuelveDesdeBDNew(cContaSeg, "tiposdiario", "desdiari", "numdiari", Text1(Index).Text, "N")
                If Text2(Index).Text = "" Then
                    MsgBox "Número de Diario no existe en la contabilidad. Reintroduzca.", vbExclamation
                End If
            End If
        
        Case 46 ' NUMERO DE DIARIO de gasolinera
            If Text1(Index).Text <> "" Then
                Text2(Index).Text = ""
                Text2(Index).Text = DevuelveDesdeBDNew(cContaGas, "tiposdiario", "desdiari", "numdiari", Text1(Index).Text, "N")
                If Text2(Index).Text = "" Then
                    MsgBox "Número de Diario no existe en la contabilidad. Reintroduzca.", vbExclamation
                End If
            End If
            
         Case 7 ' RAIZ DE CTA CONTABLE
            If Text1(Index).Text <> "" Then
                Text2(Index).Text = ""
                Text2(Index).Text = DevuelveDesdeBDNew(cContaSeg, "cuentas", "nommacta", "codmacta", Text1(Index).Text, "T")
                If Text2(Index).Text = "" Then
                    MsgBox "Raíz de Cuenta Contable incorrecta. Reintroduzca.", vbExclamation
                End If
            End If
            
         Case 43 ' RAIZ DE CTA CONTABLE de gasolinera
            If Text1(Index).Text <> "" Then
                Text2(Index).Text = ""
                Text2(Index).Text = DevuelveDesdeBDNew(cContaGas, "cuentas", "nommacta", "codmacta", Text1(Index).Text, "T")
                If Text2(Index).Text = "" Then
                    MsgBox "Raíz de Cuenta Contable incorrecta. Reintroduzca.", vbExclamation
                End If
            End If
       
         Case 53 ' RAIZ DE CTA CONTABLE de facturas socios
            If Text1(Index).Text <> "" Then
                Text2(Index).Text = ""
                Text2(Index).Text = DevuelveDesdeBDNew(cContaFacSoc, "cuentas", "nommacta", "codmacta", Text1(Index).Text, "T")
                If Text2(Index).Text = "" Then
                    MsgBox "Raíz de Cuenta Contable incorrecta. Reintroduzca.", vbExclamation
                End If
            End If
       
         Case 64 ' RAIZ DE CTA CONTABLE de coarval
            If Text1(Index).Text <> "" Then
                Text2(Index).Text = ""
                Text2(Index).Text = DevuelveDesdeBDNew(cContaCV, "cuentas", "nommacta", "codmacta", Text1(Index).Text, "T")
                If Text2(Index).Text = "" Then
                    MsgBox "Raíz de Cuenta Contable incorrecta. Reintroduzca.", vbExclamation
                End If
            End If
       
        Case 14, 15 'CONCEPTOS de avnics
            If Text1(Index).Text <> "" Then Text2(Index).Text = PonerNombreConcepto(Text1(Index), cConta)
            If Text2(Index).Text = "" Then
                MsgBox "Número de Concepto no existe en la contabilidad. Reintroduzca.", vbExclamation
            End If
        
        Case 19, 20 'CONCEPTOS de seguros
            If Text1(Index).Text <> "" Then Text2(Index).Text = PonerNombreConcepto(Text1(Index), cContaSeg)
            If Text2(Index).Text = "" Then
                MsgBox "Número de Concepto no existe en la contabilidad. Reintroduzca.", vbExclamation
            End If
     
        Case 31, 32 'CONCEPTOS de telefonia
            If Text1(Index).Text <> "" Then Text2(Index).Text = PonerNombreConcepto(Text1(Index), cContaTel)
            If Text2(Index).Text = "" Then
                MsgBox "Número de Concepto no existe en la contabilidad. Reintroduzca.", vbExclamation
            End If
        
        Case 44, 45 'CONCEPTOS de gasolinera
            If Text1(Index).Text <> "" Then Text2(Index).Text = PonerNombreConcepto(Text1(Index), cContaGas)
            If Text2(Index).Text = "" Then
                MsgBox "Número de Concepto no existe en la contabilidad. Reintroduzca.", vbExclamation
            End If
            
        Case 36 'LETRA DE SERIE DE TELEFONIA
            If Text1(Index).Text <> "" Then Text1(Index).Text = UCase(Trim(Text1(Index).Text))
        
        Case 40 'LETRA DE SERIE DE GASOLINERA
            If Text1(Index).Text <> "" Then Text1(Index).Text = UCase(Trim(Text1(Index).Text))
            
        Case 82 'LETRA DE SERIE DE GASOLEO B
            If Text1(Index).Text <> "" Then Text1(Index).Text = UCase(Trim(Text1(Index).Text))
        
        Case 52 ' TIPO DE IVA DE GASOLINERA
            If Text1(Index).Text = "" Then Exit Sub
            Text2(Index).Text = PonerNombreTIva(Text1(Index), cContaGas)
            
        Case 61, 75, 76, 78 'LETRA DE SERIE DE FACTURAS COARVAL
            If Text1(Index).Text <> "" Then Text1(Index).Text = UCase(Text1(Index).Text)
        
        
        Case 71, 72 ' forma de pago
            If PonerFormatoEntero(Text1(Index)) Then
                If vParamAplic.ContabilidadNueva Then
                    Text2(Index) = DevuelveDesdeBDNew(cContaCV, "formapago", "nomforpa", "codforpa", Text1(Index).Text, "N")
                Else
                    Text2(Index) = DevuelveDesdeBDNew(cContaCV, "sforpa", "nomforpa", "codforpa", Text1(Index).Text, "N")
                End If
                If Text2(Index).Text = "" Then
                    MsgBox "Forma de pago no existe. Reintroduzca.", vbExclamation
                    Text1(Index).Text = ""
                    PonerFoco Text1(Index)
                End If
            End If
            
        Case 79, 80 ' forma de pago
            If PonerFormatoEntero(Text1(Index)) Then
                If vParamAplic.ContabilidadNueva Then
                    Text2(Index) = DevuelveDesdeBDNew(cContaCVV, "formapago", "nomforpa", "codforpa", Text1(Index).Text, "N")
                Else
                    Text2(Index) = DevuelveDesdeBDNew(cContaCVV, "sforpa", "nomforpa", "codforpa", Text1(Index).Text, "N")
                End If
                If Text2(Index).Text = "" Then
                    MsgBox "Forma de pago no existe. Reintroduzca.", vbExclamation
                    PonerFoco Text1(Index)
                End If
            End If
        
        '[Monica]12/11/2013: introducimos el aridoc
        Case 83 ' carpeta de aridoc de facturas varias
            If Text1(Index).Text = "" Then Exit Sub
            Text1(Index).Text = Format(Text1(Index).Text, "000")
            Cad = CargaPath(Text1(Index))
            Text2(Index).Text = Mid(Cad, 2, Len(Cad))
        
        Case 84 ' Extension del aridoc
            If Text1(Index).Text = "" Then Exit Sub
            Text2(Index).Text = PonerNombreDeCod(Text1(Index), "extension", "descripcion", "codext", "N", cAridoc)
        
            
    End Select
End Sub

'Private Sub Text1_Validate(Index As Integer, Cancel As Boolean)
'    Select Case Index
'        Case 5, 6
'            If Text1(Index).Text <> "" Then
'                If Not EsNumerico(Text1(Index).Text) Then
'                    Cancel = True
'                    ConseguirFoco Text1(Index), Modo
'                End If
'            End If
'    End Select
'
'End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
        Case 1  'Anyadir
            BotonAnyadir
        Case 2  'Modificar
            mnModificar_Click
        Case 5 'Salir
            mnSalir_Click
    End Select
End Sub

Private Sub BotonAnyadir()
    LimpiarCampos
    PonerModo 3
    Text1(0).Text = 1
    PonerFoco Text1(1)
End Sub

Private Sub BotonModificar()
    PonerModo 4
    PonerFoco Text1(1)
End Sub

Private Function DatosOk() As Boolean
Dim b As Boolean
    DatosOk = False
    b = CompForm(Me)
    DatosOk = b
End Function

Private Sub KEYpress(KeyAscii As Integer)
Dim cerrar As Boolean

    KEYpressGnral KeyAscii, Modo, cerrar
    If cerrar Then Unload Me
End Sub

Private Sub PonerBotonCabecera(b As Boolean)
    Me.cmdAceptar.visible = Not b
    Me.cmdCancelar.visible = Not b
    Me.cmdSalir.visible = b
'    If b Then Me.lblIndicador.Caption = ""
End Sub

Private Sub PonerCampos()
Dim i As Byte
Dim Cad As String


On Error GoTo EPonerCampos

    If Data1.Recordset.EOF Then Exit Sub
    PonerCamposForma Me, Data1
    
    ' ************* configurar els camps de les descripcions de les comptes *************
    
    ' avnics
    If vParamAplic.Avnics = 1 Then
     If vParamAplic.NumeroConta <> 0 Then
        For i = 5 To 6
            Text2(i).Text = PonerNombreCuenta(Text1(i), Modo, , cConta)
        Next i
        ' numero de diario
        Text2(16).Text = DevuelveDesdeBDNew(cConta, "tiposdiario", "desdiari", "numdiari", Text1(16).Text, "N")
        ' conceptos contables
        For i = 14 To 15
            Text2(i).Text = PonerNombreConcepto(Text1(i), cConta)
        Next i
     End If
    End If
    ' seguros
    If vParamAplic.Seguros = 1 Then
     If vParamAplic.NumeroContaSeg <> 0 Then
        For i = 8 To 8
            Text2(i).Text = PonerNombreCuenta(Text1(i), Modo, , cContaSeg)
        Next i
        ' raiz de la cta contable de socio
        Text2(7).Text = DevuelveDesdeBDNew(cContaSeg, "cuentas", "nommacta", "codmacta", Text1(7).Text, "N")
        ' numero de diario
        Text2(21).Text = DevuelveDesdeBDNew(cContaSeg, "tiposdiario", "desdiari", "numdiari", Text1(21).Text, "N")
        ' conceptos contables
        For i = 19 To 20
            Text2(i).Text = PonerNombreConcepto(Text1(i), cContaSeg)
        Next i
     End If
    End If
    ' telefonia
    If vParamAplic.Telefonia = 1 Then
     If vParamAplic.NumeroContaTel <> 0 Then
        ' cuentas contables
        For i = 34 To 35
            Text2(i).Text = PonerNombreCuenta(Text1(i), Modo, , cContaTel)
        Next i
        ' raiz de la cta contable de socio
        Text2(33).Text = DevuelveDesdeBDNew(cContaTel, "cuentas", "nommacta", "codmacta", Text1(33).Text, "N")
        ' numero de diario
        Text2(30).Text = DevuelveDesdeBDNew(cContaTel, "tiposdiario", "desdiari", "numdiari", Text1(30).Text, "N")
        ' conceptos contables
        For i = 31 To 32
            Text2(i).Text = PonerNombreConcepto(Text1(i), cContaTel)
        Next i
     End If
    End If

    ' gasolinera
    If vParamAplic.Gasolinera = 1 Then
     If vParamAplic.NumeroContaGas <> 0 Then
        ' cuentas contables
        For i = 41 To 42
            Text2(i).Text = PonerNombreCuenta(Text1(i), Modo, , cContaGas)
        Next i
        ' raiz de la cta contable de socio
        Text2(43).Text = DevuelveDesdeBDNew(cContaGas, "cuentas", "nommacta", "codmacta", Text1(43).Text, "T")
        ' numero de diario
        Text2(46).Text = DevuelveDesdeBDNew(cContaGas, "tiposdiario", "desdiari", "numdiari", Text1(46).Text, "N")
        ' conceptos contables
        For i = 44 To 45
            Text2(i).Text = PonerNombreConcepto(Text1(i), cContaGas)
        Next i
        'tipo de iva
        Text2(52).Text = PonerNombreTIva(Text1(52), cContaGas)
     End If
    End If
    
    ' facturas socios
    If vParamAplic.FactSocios = 1 Then
     If vParamAplic.NumeroContaFacSoc <> 0 Then
        ' cuentas contables
        For i = 59 To 59
            Text2(i).Text = PonerNombreCuenta(Text1(i), Modo, , cContaFacSoc)
        Next i
        ' raiz de la cta contable de socio
        Text2(53).Text = DevuelveDesdeBDNew(cContaFacSoc, "cuentas", "nommacta", "codmacta", Text1(53).Text, "T")
     End If
    End If
    
    ' facturas coarval
    If vParamAplic.Coarval = 1 Then
     If vParamAplic.NumeroContaCV <> 0 Then
        ' cuentas contables
        For i = 62 To 63
            Text2(i).Text = PonerNombreCuenta(Text1(i), Modo, , cContaCV)
        Next i
        ' cuentas contables
        For i = 69 To 70
            Text2(i).Text = PonerNombreCuenta(Text1(i), Modo, , cContaCV)
        Next i
        ' raiz de la cta contable de socio
        Text2(64).Text = DevuelveDesdeBDNew(cContaCV, "cuentas", "nommacta", "codmacta", Text1(64).Text, "T")
        'formas de pago
        If vParamAplic.ContabilidadNueva Then
            Text2(71) = DevuelveDesdeBDNew(cContaCV, "formapago", "nomforpa", "codforpa", Text1(71).Text, "N")
            Text2(72) = DevuelveDesdeBDNew(cContaCV, "formapago", "nomforpa", "codforpa", Text1(72).Text, "N")
        Else
            Text2(71) = DevuelveDesdeBDNew(cContaCV, "sforpa", "nomforpa", "codforpa", Text1(71).Text, "N")
            Text2(72) = DevuelveDesdeBDNew(cContaCV, "sforpa", "nomforpa", "codforpa", Text1(72).Text, "N")
        End If
        Text2(73).Text = PonerNombreCuenta(Text1(73), Modo, , cContaCV)
        Text2(74).Text = PonerNombreCuenta(Text1(74), Modo, , cContaCV)
     End If
     If vParamAplic.NumeroContaCVV <> 0 Then
        ' cuentas contables
        For i = 81 To 81
            Text2(i).Text = PonerNombreCuenta(Text1(i), Modo, , cContaCVV)
        Next i
        'formas de pago
        If vParamAplic.ContabilidadNueva Then
            Text2(79) = DevuelveDesdeBDNew(cContaCVV, "formapago", "nomforpa", "codforpa", Text1(79).Text, "N")
            Text2(80) = DevuelveDesdeBDNew(cContaCVV, "formapago", "nomforpa", "codforpa", Text1(80).Text, "N")
        Else
            Text2(79) = DevuelveDesdeBDNew(cContaCVV, "sforpa", "nomforpa", "codforpa", Text1(79).Text, "N")
            Text2(80) = DevuelveDesdeBDNew(cContaCVV, "sforpa", "nomforpa", "codforpa", Text1(80).Text, "N")
        End If
     End If
     
    End If

    'aridoc
    If vParamAplic.HayAridoc = 1 Then
         If ComprobarCero(Text1(83).Text) <> 0 Then
            Cad = CargaPath(Text1(83))
            Text2(83).Text = Mid(Cad, 2, Len(Cad))
         End If

         Text2(84).Text = DevuelveDesdeBDNew(cAridoc, "extension", "descripcion", "codext", Text1(84).Text, "N")
    End If
    


EPonerCampos:
    If Err.Number <> 0 Then MuestraError Err.Number, "Poniendo Campos", Err.Description
End Sub

Private Sub PonerOpcionesMenu()
    PonerOpcionesMenuGeneral Me
End Sub

Private Sub LimpiarCampos()
    Limpiar Me   'Metodo general: Limpia los controles TextBox
    lblIndicador.Caption = ""
    'Aqui va el especifico de cada form es
    '### a mano
End Sub

'----------------------------------------------------------------
'   En PONERMODO se habilitan, o no, los diverso campos del
'   formulario en funcion del modo en k vayamos a trabajar
Private Sub PonerModo(Kmodo As Byte)
Dim b As Boolean
Dim i As Byte

    Modo = Kmodo
    PonerIndicador lblIndicador, Modo
      
    '------------------------------------------------------
    'Modo insertar o modificar
    b = (Kmodo >= 3) '-->Luego not b sera kmodo<3
    If cmdCancelar.visible Then
        cmdCancelar.Cancel = True
    Else
        cmdCancelar.Cancel = False
    End If
    PonerBotonCabecera Not b
       
    'Bloquea los campos Text1 sino estamos modificando/Insertando Datos
    'Si estamos en Insertar además limpia los campos Text1 y bloquea la clave primaria
    BloquearText1 Me, Modo
    
    BloquearChecks Me, Modo

    BloquearCmb Combo1(0), Modo < 3
    'Bloquear imagen de Busqueda
    For i = 0 To imgBuscar.Count - 1
        Me.imgBuscar(i).Enabled = (Modo >= 3)
    Next i
   
    BloquearImgBuscar Me, Modo

    PonerModoOpcionesMenu 'Activar opciones de menu según el Modo
    PonerOpcionesMenu   'Activar opciones de menu según nivel
                        'de permisos del usuario
End Sub

Private Sub PonerModoOpcionesMenu()
Dim b As Boolean
    b = (Modo = 3) Or (Modo = 4)
    Me.Toolbar1.Buttons(1).Enabled = Not Encontrado And Not b  'Añadir
    Me.Toolbar1.Buttons(2).Enabled = Encontrado And Not b 'Modificar
    Me.mnAñadir.Enabled = Not Encontrado And Not b
    Me.mnModificar.Enabled = Encontrado And Not b
'    Me.Toolbar1.Buttons(2).Enabled = (Not b) 'Modificar
End Sub

Private Sub AbrirFrmConceptos(indice As Integer, Conexion As Byte)
    indCodigo = indice
    Set frmConce = New frmConceConta
    frmConce.DatosADevolverBusqueda = "0|1|"
    frmConce.CodigoActual = Text1(indCodigo)
    frmConce.Conexion = Conexion
    frmConce.Show vbModal
    Set frmConce = Nothing
End Sub

Private Sub AbrirFrmDiario(indice As Integer, Conexion As Byte)
    indCodigo = indice
    Set frmTDia = New frmDiaConta
    frmTDia.DatosADevolverBusqueda = "0|1|"
    frmTDia.CodigoActual = Text1(indCodigo)
    frmTDia.Conexion = Conexion
    frmTDia.Show vbModal
    Set frmTDia = Nothing
End Sub

Private Sub AbrirFrmTipIvaConta(indice As Integer)
    indCodigo = indice
    Set frmTIva = New frmTipIVAConta
    frmTIva.DatosADevolverBusqueda = "0|1|"
    frmTIva.CodigoActual = Text1(indCodigo)
    frmTIva.Conexion = cContaGas
    frmTIva.Show vbModal
    Set frmTIva = Nothing
End Sub


Private Sub ComprobarDatos()
    ' avnics
    If vParamAplic.Avnics Then
        If vParamAplic.NumeroConta <> 0 Then
            Text1_LostFocus (5)
            Text1_LostFocus (6)
            Text1_LostFocus (14)
            Text1_LostFocus (15)
            Text1_LostFocus (16)
        End If
    End If
    ' seguros
    If vParamAplic.Seguros Then
        If vParamAplic.NumeroContaSeg <> 0 Then
            Text1_LostFocus (8)
            Text1_LostFocus (19)
            Text1_LostFocus (20)
            Text1_LostFocus (21)
            Text1_LostFocus (7)
        End If
    End If
    ' telefonia
    If vParamAplic.Telefonia Then
        If vParamAplic.NumeroContaTel <> 0 Then
            Text1_LostFocus (33)
        End If
    End If
    'gasolinera
    If vParamAplic.Gasolinera Then
        If vParamAplic.NumeroContaGas <> 0 Then
            Text1_LostFocus (44)
            Text1_LostFocus (45)
            Text1_LostFocus (46)
            Text1_LostFocus (42)
            Text1_LostFocus (41)
            Text1_LostFocus (43)
            Text1_LostFocus (52)
        End If
    End If
    'facturas socios
    If vParamAplic.FactSocios Then
        If vParamAplic.NumeroContaFacSoc <> 0 Then
            Text1_LostFocus (53)
            Text1_LostFocus (54)
            Text1_LostFocus (55)
            Text1_LostFocus (56)
            Text1_LostFocus (57)
            Text1_LostFocus (58)
            Text1_LostFocus (59)
        End If
    End If


End Sub

' ********* si n'hi han combos a la capçalera ************
Private Sub CargaCombo()
Dim Ini As Integer
Dim Fin As Integer
Dim i As Integer

    ' *** neteje els combos, els pose valor i seleccione el valor per defecte ***
    ' tipo de fichero de importacion de telefonia
    Combo1(0).Clear
    
    Combo1(0).AddItem "Excel"
    Combo1(0).ItemData(Combo1(0).NewIndex) = 0
    Combo1(0).AddItem "Texto"
    Combo1(0).ItemData(Combo1(0).NewIndex) = 1
    
    '[Monica]12/11/2013: añadido la integracion al aridoc
    'combos de facturas varias
    For i = 5 To 8
        Combo1(i).AddItem "Nro.Factura"
        Combo1(i).ItemData(Combo1(i).NewIndex) = 0
        Combo1(i).AddItem "Cod.Cliente"
        Combo1(i).ItemData(Combo1(i).NewIndex) = 1
        Combo1(i).AddItem "Nom.Cliente"
        Combo1(i).ItemData(Combo1(i).NewIndex) = 2
        Combo1(i).AddItem "Procedencia"
        Combo1(i).ItemData(Combo1(i).NewIndex) = 3
    Next i
    
    
End Sub

Private Function CargaPath(Codigo As Integer) As String
Dim Nod As Node
Dim J As Integer
Dim i As Integer
Dim C As String
Dim campo1 As String
Dim padre As String
Dim A As String

    'Primero copiamos la carpeta
    C = "\" & DevuelveDesdeBDNew(cAridoc, "carpetas", "nombre", "codcarpeta", CInt(Codigo), "N")
    campo1 = "nombre"
    padre = DevuelveDesdeBDNew(cAridoc, "carpetas", "padre", "codcarpeta", CStr(Codigo), "N", campo1)
    If CInt(ComprobarCero(padre)) > 0 Then
        C = CargaPath(CInt(padre)) & C
    End If
'
'    If No.Children > 0 Then
'        J = No.Children
'        Set Nod = No.Child
'        For i = 1 To J
'           C = C & CopiaArchivosCarpetaRecursiva(Nod)
'           If i <> J Then Set Nod = Nod.Next
'        Next i
'    End If
    CargaPath = C
End Function

