VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmMensajes 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   7380
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11880
   Icon            =   "frmMensajes.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7380
   ScaleWidth      =   11880
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FrameImpPunteo 
      Height          =   3375
      Left            =   30
      TabIndex        =   176
      Top             =   0
      Width           =   5655
      Begin VB.TextBox txtImporteP 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   8
         Left            =   4200
         Locked          =   -1  'True
         TabIndex        =   192
         Text            =   "Text9"
         Top             =   2160
         Width           =   1215
      End
      Begin VB.TextBox txtImporteP 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   7
         Left            =   2520
         Locked          =   -1  'True
         TabIndex        =   191
         Text            =   "Text9"
         Top             =   2160
         Width           =   1335
      End
      Begin VB.TextBox txtImporteP 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   6
         Left            =   1080
         Locked          =   -1  'True
         TabIndex        =   190
         Text            =   "Text9"
         Top             =   2160
         Width           =   1335
      End
      Begin VB.TextBox txtImporteP 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   5
         Left            =   4200
         Locked          =   -1  'True
         TabIndex        =   187
         Text            =   "Text9"
         Top             =   1680
         Width           =   1215
      End
      Begin VB.TextBox txtImporteP 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   4
         Left            =   2520
         Locked          =   -1  'True
         TabIndex        =   186
         Text            =   "Text9"
         Top             =   1680
         Width           =   1335
      End
      Begin VB.TextBox txtImporteP 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   3
         Left            =   1080
         Locked          =   -1  'True
         TabIndex        =   185
         Text            =   "Text9"
         Top             =   1680
         Width           =   1335
      End
      Begin VB.TextBox txtImporteP 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   2
         Left            =   4200
         Locked          =   -1  'True
         TabIndex        =   184
         Text            =   "Text9"
         Top             =   1200
         Width           =   1215
      End
      Begin VB.TextBox txtImporteP 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   1
         Left            =   2520
         Locked          =   -1  'True
         TabIndex        =   183
         Text            =   "Text9"
         Top             =   1200
         Width           =   1335
      End
      Begin VB.TextBox txtImporteP 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   0
         Left            =   1080
         Locked          =   -1  'True
         TabIndex        =   182
         Text            =   "Text9"
         Top             =   1200
         Width           =   1335
      End
      Begin VB.CommandButton cmdPunteo 
         Caption         =   "Cerrar"
         Height          =   375
         Left            =   4320
         TabIndex        =   177
         Top             =   2880
         Width           =   1095
      End
      Begin VB.Line Line7 
         BorderWidth     =   3
         X1              =   120
         X2              =   5400
         Y1              =   2040
         Y2              =   2040
      End
      Begin VB.Label Label22 
         Caption         =   "Haber"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   15
         Left            =   120
         TabIndex        =   189
         Top             =   1680
         Width           =   930
      End
      Begin VB.Label Label22 
         Caption         =   "Debe"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   14
         Left            =   120
         TabIndex        =   188
         Top             =   1200
         Width           =   930
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         Caption         =   "Total"
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
         Index           =   13
         Left            =   4440
         TabIndex        =   181
         Top             =   840
         Width           =   450
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         Caption         =   "Sin puntear"
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
         Index           =   12
         Left            =   2640
         TabIndex        =   180
         Top             =   840
         Width           =   990
      End
      Begin VB.Label Label22 
         Caption         =   "Punteada"
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
         Index           =   11
         Left            =   1200
         TabIndex        =   179
         Top             =   840
         Width           =   930
      End
      Begin VB.Label Label37 
         Caption         =   "Importes punteo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   435
         Index           =   1
         Left            =   240
         TabIndex        =   178
         Top             =   240
         Width           =   3735
      End
   End
   Begin VB.Frame FrameCambioPWD 
      Height          =   3615
      Left            =   0
      TabIndex        =   148
      Top             =   0
      Width           =   4575
      Begin VB.TextBox Text7 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Index           =   3
         Left            =   3120
         PasswordChar    =   "*"
         TabIndex        =   152
         Text            =   "Text7"
         Top             =   2400
         Width           =   1215
      End
      Begin VB.TextBox Text7 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Index           =   2
         Left            =   3120
         PasswordChar    =   "*"
         TabIndex        =   151
         Text            =   "Text7"
         Top             =   1680
         Width           =   1215
      End
      Begin VB.TextBox Text7 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Index           =   1
         Left            =   240
         PasswordChar    =   "*"
         TabIndex        =   150
         Text            =   "Text7"
         Top             =   1680
         Width           =   1335
      End
      Begin VB.TextBox Text7 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   0
         Left            =   240
         TabIndex        =   149
         Text            =   "Text7"
         Top             =   960
         Width           =   4095
      End
      Begin VB.CommandButton cmdCambioPwd 
         Caption         =   "Cancelar"
         Height          =   375
         Index           =   1
         Left            =   3240
         TabIndex        =   154
         Top             =   3000
         Width           =   1095
      End
      Begin VB.CommandButton cmdCambioPwd 
         Caption         =   "Aceptar"
         Height          =   375
         Index           =   0
         Left            =   2040
         TabIndex        =   153
         Top             =   3000
         Width           =   1095
      End
      Begin VB.Label Label37 
         Alignment       =   2  'Center
         Caption         =   "Cambio clave"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   159
         Top             =   240
         Width           =   4335
      End
      Begin VB.Label Label36 
         Caption         =   "Reescribalo"
         Height          =   255
         Index           =   3
         Left            =   3120
         TabIndex        =   158
         Top             =   2160
         Width           =   975
      End
      Begin VB.Label Label36 
         Caption         =   "Nuevo password"
         Height          =   255
         Index           =   2
         Left            =   3120
         TabIndex        =   157
         Top             =   1440
         Width           =   1695
      End
      Begin VB.Label Label36 
         Caption         =   "Password actual"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   156
         Top             =   1440
         Width           =   2295
      End
      Begin VB.Label Label36 
         Caption         =   "Nombre"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   155
         Top             =   720
         Width           =   975
      End
   End
   Begin VB.Frame frameSaltos 
      Height          =   4935
      Left            =   0
      TabIndex        =   100
      Top             =   0
      Width           =   5835
      Begin VB.CommandButton cmdCabError 
         Caption         =   "&Salir"
         Height          =   375
         Index           =   1
         Left            =   4680
         TabIndex        =   107
         Top             =   4320
         Width           =   975
      End
      Begin VB.CommandButton cmdCabError 
         Caption         =   "Aceptar"
         Height          =   375
         Index           =   0
         Left            =   3480
         TabIndex        =   106
         Top             =   4320
         Width           =   975
      End
      Begin VB.TextBox Text6 
         Height          =   3255
         Left            =   3120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   103
         Text            =   "frmMensajes.frx":000C
         Top             =   900
         Width           =   2535
      End
      Begin VB.TextBox Text5 
         Height          =   3255
         Left            =   180
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   102
         Text            =   "frmMensajes.frx":0012
         Top             =   900
         Width           =   2535
      End
      Begin VB.Label Label22 
         Caption         =   "Salto"
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
         Index           =   10
         Left            =   2880
         TabIndex        =   105
         Top             =   660
         Width           =   930
      End
      Begin VB.Label Label22 
         Caption         =   "Repetidos"
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
         Index           =   9
         Left            =   180
         TabIndex        =   104
         Top             =   660
         Width           =   930
      End
      Begin VB.Label Label24 
         Caption         =   "Asientos con cabecera erronea"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   435
         Left            =   180
         TabIndex        =   101
         Top             =   240
         Width           =   4875
      End
   End
   Begin VB.Frame frameSaldosHco 
      BorderStyle     =   0  'None
      Height          =   4575
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6015
      Begin VB.Frame Frame1 
         BorderStyle     =   0  'None
         Height          =   495
         Left            =   120
         TabIndex        =   125
         Top             =   3240
         Width           =   5415
         Begin VB.TextBox txtsaldo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   9
            Left            =   3480
            Locked          =   -1  'True
            TabIndex        =   127
            Text            =   "Text1"
            Top             =   120
            Width           =   1815
         End
         Begin VB.TextBox txtsaldo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   8
            Left            =   1560
            Locked          =   -1  'True
            TabIndex        =   126
            Text            =   "Text1"
            Top             =   120
            Width           =   1815
         End
         Begin VB.Label Label28 
            Caption         =   "SALDO PERIODO"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00004000&
            Height          =   255
            Index           =   1
            Left            =   0
            TabIndex        =   128
            Top             =   120
            Width           =   1575
         End
      End
      Begin VB.TextBox txtsaldo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   7
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   123
         Text            =   "Text1"
         Top             =   2880
         Width           =   1815
      End
      Begin VB.TextBox txtsaldo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   6
         Left            =   3600
         Locked          =   -1  'True
         TabIndex        =   122
         Text            =   "Text1"
         Top             =   2880
         Width           =   1815
      End
      Begin VB.TextBox txtsaldo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   5
         Left            =   3600
         Locked          =   -1  'True
         TabIndex        =   113
         Text            =   "Text1"
         Top             =   2520
         Width           =   1815
      End
      Begin VB.TextBox txtsaldo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   4
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   112
         Text            =   "Text1"
         Top             =   2520
         Width           =   1815
      End
      Begin VB.TextBox txtsaldo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   3
         Left            =   3600
         Locked          =   -1  'True
         TabIndex        =   6
         Text            =   "Text1"
         Top             =   1560
         Width           =   1815
      End
      Begin VB.TextBox txtsaldo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   2
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   5
         Text            =   "Text1"
         Top             =   1560
         Width           =   1815
      End
      Begin VB.TextBox txtsaldo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   1
         Left            =   3600
         Locked          =   -1  'True
         TabIndex        =   4
         Text            =   "Text1"
         Top             =   2040
         Width           =   1815
      End
      Begin VB.TextBox txtsaldo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   0
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   3
         Text            =   "Text1"
         Top             =   2040
         Width           =   1815
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Salir"
         Height          =   375
         Index           =   0
         Left            =   4320
         TabIndex        =   1
         Top             =   3840
         Width           =   1095
      End
      Begin VB.Image Image6 
         Height          =   240
         Index           =   0
         Left            =   5520
         Picture         =   "frmMensajes.frx":0018
         Top             =   1600
         Width           =   240
      End
      Begin VB.Image Image6 
         Height          =   240
         Index           =   1
         Left            =   5520
         Picture         =   "frmMensajes.frx":05A2
         Top             =   2070
         Width           =   240
      End
      Begin VB.Line Line1 
         X1              =   120
         X2              =   5520
         Y1              =   2400
         Y2              =   2400
      End
      Begin VB.Label Label28 
         Caption         =   "SALDO"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   124
         Top             =   2880
         Width           =   1335
      End
      Begin VB.Label Label14 
         Caption         =   "TOTALES"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   120
         TabIndex        =   114
         Top             =   2520
         Width           =   1335
      End
      Begin VB.Label Label6 
         Caption         =   "Label6"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   720
         Width           =   6855
      End
      Begin VB.Label Label5 
         Caption         =   "PENDIENTE"
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
         TabIndex        =   10
         Top             =   2040
         Width           =   1335
      End
      Begin VB.Label Label4 
         Caption         =   "PUNTEADA"
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
         TabIndex        =   9
         Top             =   1620
         Width           =   1095
      End
      Begin VB.Label Label3 
         Caption         =   "HABER"
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
         Left            =   4080
         TabIndex        =   8
         Top             =   1200
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "DEBE"
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
         Left            =   2280
         TabIndex        =   7
         Top             =   1200
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Saldos histórico"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   495
         Left            =   120
         TabIndex        =   2
         Top             =   120
         Width           =   5295
      End
   End
   Begin VB.Frame frameAcercaDE 
      BorderStyle     =   0  'None
      Height          =   4155
      Left            =   -60
      TabIndex        =   48
      Top             =   0
      Width           =   5355
      Begin VB.Label Label27 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fax: 96 378 82 01"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   195
         Left            =   3180
         TabIndex        =   110
         Top             =   3540
         Width           =   1560
      End
      Begin VB.Label Label26 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tfno: 96 358 05 47"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   195
         Left            =   780
         TabIndex        =   109
         Top             =   3540
         Width           =   1650
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "ARICONTA"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   36
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   915
         Left            =   240
         TabIndex        =   49
         Top             =   360
         Width           =   4335
      End
      Begin VB.Label Label25 
         BackStyle       =   0  'Transparent
         Caption         =   "4"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   81.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   1725
         Left            =   3780
         TabIndex        =   108
         Top             =   60
         Width           =   1350
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "46007 - VALENCIA"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   195
         Left            =   3240
         TabIndex        =   64
         Top             =   3120
         Width           =   1620
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "C/ Franco Tormo, 3 Bajo Izda"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   195
         Left            =   300
         TabIndex        =   63
         Top             =   3120
         Width           =   2535
      End
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Label13"
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
         Left            =   0
         TabIndex        =   50
         Top             =   1200
         Width           =   3795
      End
      Begin VB.Image Image2 
         Height          =   540
         Left            =   1740
         Top             =   2460
         Width           =   2880
      End
      Begin VB.Image Image1 
         Height          =   4395
         Left            =   0
         Stretch         =   -1  'True
         Top             =   -1200
         Width           =   5355
      End
   End
   Begin VB.Frame frameCtasBalance 
      Height          =   3315
      Left            =   0
      TabIndex        =   65
      Top             =   0
      Width           =   6735
      Begin VB.CheckBox chkResta 
         Caption         =   "Se resta "
         Height          =   255
         Left            =   960
         TabIndex        =   173
         Top             =   2640
         Width           =   2535
      End
      Begin VB.CommandButton cmdCtaBalan 
         Caption         =   "&Cancelar"
         Height          =   435
         Index           =   1
         Left            =   5400
         TabIndex        =   75
         Top             =   2640
         Width           =   1095
      End
      Begin VB.CommandButton cmdCtaBalan 
         Caption         =   "Command4"
         Height          =   435
         Index           =   0
         Left            =   4140
         TabIndex        =   74
         Top             =   2640
         Width           =   1095
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Haber"
         Height          =   255
         Index           =   2
         Left            =   5220
         TabIndex        =   73
         Top             =   1920
         Width           =   1095
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Debe"
         Height          =   255
         Index           =   1
         Left            =   3960
         TabIndex        =   72
         Top             =   1920
         Width           =   855
      End
      Begin VB.OptionButton Option1 
         Caption         =   "SALDO"
         Height          =   255
         Index           =   0
         Left            =   2520
         TabIndex        =   71
         Top             =   1920
         Value           =   -1  'True
         Width           =   855
      End
      Begin VB.TextBox Text3 
         Height          =   315
         Left            =   1020
         TabIndex        =   69
         Text            =   "Text2"
         Top             =   1860
         Width           =   1275
      End
      Begin VB.TextBox Text2 
         Enabled         =   0   'False
         Height          =   315
         Left            =   960
         TabIndex        =   68
         Text            =   "Text2"
         Top             =   900
         Width           =   5535
      End
      Begin VB.Image Image3 
         Height          =   240
         Left            =   780
         Picture         =   "frmMensajes.frx":0B2C
         Top             =   1920
         Width           =   240
      End
      Begin VB.Label Label21 
         Caption         =   "Cuenta"
         Height          =   255
         Left            =   120
         TabIndex        =   70
         Top             =   1920
         Width           =   615
      End
      Begin VB.Label Label20 
         Caption         =   "Grupo"
         Height          =   255
         Left            =   120
         TabIndex        =   67
         Top             =   960
         Width           =   615
      End
      Begin VB.Label Label19 
         Caption         =   "MODIFICAR"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   435
         Left            =   180
         TabIndex        =   66
         Top             =   240
         Width           =   4875
      End
   End
   Begin VB.Frame FrameImpCta 
      Height          =   1935
      Left            =   0
      TabIndex        =   115
      Top             =   0
      Width           =   6495
      Begin VB.CheckBox chkCrear 
         Caption         =   "Crear cuentas si no existen"
         Height          =   195
         Left            =   3600
         TabIndex        =   147
         Top             =   360
         Width           =   2415
      End
      Begin VB.CommandButton cmdImpCta 
         Caption         =   "Salir"
         Height          =   375
         Index           =   1
         Left            =   4920
         TabIndex        =   118
         Top             =   1320
         Width           =   1095
      End
      Begin VB.TextBox txtImpCta 
         Height          =   285
         Left            =   360
         TabIndex        =   117
         Text            =   "Text2"
         Top             =   720
         Width           =   5655
      End
      Begin VB.CommandButton cmdImpCta 
         Caption         =   "Aceptar"
         Height          =   375
         Index           =   0
         Left            =   3720
         TabIndex        =   116
         Top             =   1320
         Width           =   1095
      End
      Begin VB.Image ImageAyudaImpcta 
         Height          =   240
         Left            =   2640
         Picture         =   "frmMensajes.frx":0C2E
         Top             =   360
         Width           =   240
      End
      Begin VB.Label lblImpCta 
         Caption         =   "Label3"
         Height          =   255
         Left            =   1080
         TabIndex        =   121
         Top             =   1320
         Width           =   1335
      End
      Begin VB.Label lblimpCta2 
         Caption         =   "Lineas"
         Height          =   255
         Left            =   360
         TabIndex        =   120
         Top             =   1320
         Width           =   615
      End
      Begin VB.Image Image4 
         Height          =   240
         Left            =   2280
         Picture         =   "frmMensajes.frx":0D30
         Top             =   360
         Width           =   240
      End
      Begin VB.Label lblDescFich 
         Caption         =   "Fichero con datos fiscales"
         Height          =   255
         Left            =   360
         TabIndex        =   119
         Top             =   360
         Width           =   2055
      End
   End
   Begin VB.Frame frameBalance 
      Height          =   5535
      Left            =   0
      TabIndex        =   51
      Top             =   0
      Width           =   5655
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   3
         Left            =   480
         MaxLength       =   10
         TabIndex        =   57
         Text            =   "Text1"
         Top             =   3420
         Width           =   1575
      End
      Begin VB.CheckBox chkPintar 
         Caption         =   "Escribir si el resultado es negativo"
         Height          =   255
         Left            =   480
         TabIndex        =   77
         Top             =   3900
         Width           =   4875
      End
      Begin VB.CheckBox chkCero 
         Caption         =   "Poner a CERO si el resultado es negativo"
         Height          =   255
         Left            =   480
         TabIndex        =   76
         Top             =   4260
         Width           =   4815
      End
      Begin VB.CheckBox chkNegrita 
         Caption         =   "Negrita"
         Height          =   255
         Left            =   480
         TabIndex        =   62
         Top             =   4620
         Width           =   1095
      End
      Begin VB.CommandButton cmdBalance 
         Caption         =   "Cancelar"
         Height          =   435
         Index           =   1
         Left            =   4260
         TabIndex        =   60
         Top             =   4980
         Width           =   1095
      End
      Begin VB.CommandButton cmdBalance 
         Caption         =   "Aceptar"
         Height          =   435
         Index           =   0
         Left            =   3000
         TabIndex        =   59
         Top             =   4980
         Width           =   1095
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   2
         Left            =   480
         MaxLength       =   60
         TabIndex        =   56
         Text            =   "Text1"
         Top             =   2700
         Width           =   4875
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   1
         Left            =   480
         MaxLength       =   60
         TabIndex        =   54
         Text            =   "Text1"
         Top             =   1980
         Width           =   4875
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   0
         Left            =   480
         MaxLength       =   120
         TabIndex        =   52
         Text            =   "Text1"
         Top             =   1200
         Width           =   4815
      End
      Begin VB.Label Label15 
         Caption         =   "Código oficial balance"
         Height          =   315
         Index           =   3
         Left            =   480
         TabIndex        =   111
         Top             =   3180
         Width           =   1575
      End
      Begin VB.Label Label16 
         Caption         =   "MODIFICAR"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   435
         Left            =   360
         TabIndex        =   61
         Top             =   300
         Width           =   4875
      End
      Begin VB.Label Label15 
         Caption         =   "Formula"
         Height          =   315
         Index           =   2
         Left            =   480
         TabIndex        =   58
         Top             =   2460
         Width           =   1575
      End
      Begin VB.Label Label15 
         Caption         =   "Texto cuentas"
         Height          =   315
         Index           =   1
         Left            =   480
         TabIndex        =   55
         Top             =   1740
         Width           =   1575
      End
      Begin VB.Label Label15 
         Caption         =   "Nombre"
         Height          =   315
         Index           =   0
         Left            =   480
         TabIndex        =   53
         Top             =   960
         Width           =   1575
      End
   End
   Begin VB.Frame FrameCarta347 
      Height          =   7095
      Left            =   0
      TabIndex        =   78
      Top             =   0
      Width           =   10215
      Begin VB.TextBox Text4 
         Height          =   1035
         Index           =   6
         Left            =   6900
         MaxLength       =   100
         MultiLine       =   -1  'True
         TabIndex        =   88
         Tag             =   "#Despedida"
         Text            =   "frmMensajes.frx":0E32
         Top             =   4860
         Width           =   3075
      End
      Begin VB.TextBox Text4 
         Height          =   315
         Index           =   8
         Left            =   7680
         MaxLength       =   100
         TabIndex        =   82
         Tag             =   "#Referencia"
         Text            =   "Text4"
         Top             =   1800
         Width           =   2175
      End
      Begin VB.TextBox Text4 
         Height          =   315
         Index           =   7
         Left            =   240
         MaxLength       =   100
         TabIndex        =   81
         Tag             =   "#Asunto"
         Text            =   "Text4"
         Top             =   1740
         Width           =   6615
      End
      Begin VB.TextBox Text4 
         Height          =   1875
         Index           =   5
         Left            =   300
         MaxLength       =   255
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   86
         Tag             =   "#Parrafo4"
         Text            =   "frmMensajes.frx":0E38
         Top             =   4860
         Width           =   3135
      End
      Begin VB.TextBox Text4 
         Height          =   1875
         Index           =   4
         Left            =   3720
         MaxLength       =   255
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   87
         Tag             =   "#Parrafo5"
         Text            =   "frmMensajes.frx":0E3E
         Top             =   4860
         Width           =   3135
      End
      Begin VB.TextBox Text4 
         Height          =   1875
         Index           =   3
         Left            =   6840
         MaxLength       =   255
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   85
         Tag             =   "#Parrafo3"
         Text            =   "frmMensajes.frx":0E44
         Top             =   2580
         Width           =   3135
      End
      Begin VB.TextBox Text4 
         Height          =   1875
         Index           =   2
         Left            =   3540
         MaxLength       =   255
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   84
         Tag             =   "#Parrafo2"
         Text            =   "frmMensajes.frx":0F41
         Top             =   2580
         Width           =   3135
      End
      Begin VB.TextBox Text4 
         Height          =   1875
         Index           =   1
         Left            =   240
         MaxLength       =   255
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   83
         Tag             =   "#Parrafo1"
         Text            =   "frmMensajes.frx":0F47
         Top             =   2580
         Width           =   3135
      End
      Begin VB.CommandButton cmd347 
         Cancel          =   -1  'True
         Caption         =   "Salir"
         Height          =   375
         Index           =   1
         Left            =   8820
         TabIndex        =   91
         Top             =   6360
         Width           =   915
      End
      Begin VB.CommandButton cmd347 
         Caption         =   "Aceptar"
         Height          =   375
         Index           =   0
         Left            =   7620
         TabIndex        =   90
         Top             =   6360
         Width           =   915
      End
      Begin VB.TextBox Text4 
         Height          =   315
         Index           =   0
         Left            =   240
         MaxLength       =   100
         TabIndex        =   80
         Tag             =   "#Saludos"
         Text            =   "Text4"
         Top             =   1080
         Width           =   6615
      End
      Begin VB.Label Label22 
         Caption         =   "Despedida"
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
         Index           =   8
         Left            =   6900
         TabIndex        =   99
         Top             =   4620
         Width           =   930
      End
      Begin VB.Label Label22 
         Caption         =   "Referencia"
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
         Index           =   7
         Left            =   7680
         TabIndex        =   98
         Top             =   1560
         Width           =   930
      End
      Begin VB.Label Label22 
         Caption         =   "Asunto"
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
         Index           =   6
         Left            =   240
         TabIndex        =   97
         Top             =   1500
         Width           =   690
      End
      Begin VB.Label Label22 
         Caption         =   "Parrafo 5"
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
         Index           =   5
         Left            =   3660
         TabIndex        =   96
         Top             =   4620
         Width           =   795
      End
      Begin VB.Label Label22 
         Caption         =   "Parrafo 4"
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
         Index           =   4
         Left            =   300
         TabIndex        =   95
         Top             =   4620
         Width           =   795
      End
      Begin VB.Label Label22 
         Caption         =   "Parrafo 3"
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
         Index           =   3
         Left            =   6900
         TabIndex        =   94
         Top             =   2340
         Width           =   795
      End
      Begin VB.Label Label22 
         Caption         =   "Parrafo 2"
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
         Index           =   2
         Left            =   3480
         TabIndex        =   93
         Top             =   2340
         Width           =   795
      End
      Begin VB.Label Label22 
         Caption         =   "Parrafo 1"
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
         Index           =   1
         Left            =   240
         TabIndex        =   92
         Top             =   2340
         Width           =   795
      End
      Begin VB.Label Label23 
         Caption         =   "Datos carta modelo 347"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   435
         Left            =   300
         TabIndex        =   89
         Top             =   240
         Width           =   4875
      End
      Begin VB.Label Label22 
         Caption         =   "Saludos"
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
         Left            =   240
         TabIndex        =   79
         Top             =   840
         Width           =   690
      End
   End
   Begin VB.Timer tCuadre 
      Enabled         =   0   'False
      Interval        =   1500
      Left            =   6420
      Top             =   5700
   End
   Begin VB.Frame frameamort 
      Height          =   6255
      Left            =   0
      TabIndex        =   18
      Top             =   0
      Width           =   7095
      Begin VB.CommandButton Command3 
         Caption         =   "Salir"
         Default         =   -1  'True
         Height          =   375
         Left            =   5400
         TabIndex        =   42
         Top             =   5760
         Width           =   1455
      End
      Begin VB.Label Label10 
         Caption         =   "porcentaje"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   11
         Left            =   4080
         TabIndex        =   41
         Top             =   4920
         Width           =   1065
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Coefi. maximo"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   10
         Left            =   5760
         TabIndex        =   40
         Top             =   3720
         Width           =   1200
      End
      Begin VB.Label Label11 
         Caption         =   "="
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
         Index           =   4
         Left            =   2040
         TabIndex        =   39
         Top             =   5040
         Width           =   120
      End
      Begin VB.Line Line4 
         Index           =   4
         X1              =   4080
         X2              =   5280
         Y1              =   5160
         Y2              =   5160
      End
      Begin VB.Label Label10 
         Caption         =   "100"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   9
         Left            =   4440
         TabIndex        =   38
         Top             =   5280
         Width           =   465
      End
      Begin VB.Label Label10 
         Caption         =   "Valor adquisición   x  "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   8
         Left            =   2280
         TabIndex        =   37
         Top             =   5040
         Width           =   1755
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "PORCENTAJE"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   285
         Index           =   4
         Left            =   240
         TabIndex        =   36
         Top             =   4440
         Width           =   1635
      End
      Begin VB.Label Label11 
         Caption         =   "="
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
         Index           =   3
         Left            =   2040
         TabIndex        =   35
         Top             =   3840
         Width           =   120
      End
      Begin VB.Line Line4 
         Index           =   3
         X1              =   5880
         X2              =   6840
         Y1              =   3960
         Y2              =   3960
      End
      Begin VB.Label Label10 
         Caption         =   "100"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   7
         Left            =   6120
         TabIndex        =   34
         Top             =   4080
         Width           =   345
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "(Valor adquisición -amort. acumulada)  x"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   6
         Left            =   2280
         TabIndex        =   33
         Top             =   3840
         Width           =   3435
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "DEGRESIVO"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   285
         Index           =   3
         Left            =   240
         TabIndex        =   32
         Top             =   3240
         Width           =   1440
      End
      Begin VB.Label Label11 
         Caption         =   "="
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
         Index           =   1
         Left            =   2040
         TabIndex        =   27
         Top             =   2760
         Width           =   120
      End
      Begin VB.Label Label11 
         Caption         =   "="
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
         Left            =   2040
         TabIndex        =   26
         Top             =   1560
         Width           =   120
      End
      Begin VB.Line Line4 
         Index           =   1
         X1              =   2280
         X2              =   5280
         Y1              =   2880
         Y2              =   2880
      End
      Begin VB.Label Label10 
         Caption         =   "años de vida"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   3
         Left            =   2760
         TabIndex        =   25
         Top             =   3000
         Width           =   1065
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Valor adquisición - valor residual"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   2
         Left            =   2280
         TabIndex        =   24
         Top             =   2640
         Width           =   2745
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "LINEAL"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   285
         Index           =   1
         Left            =   240
         TabIndex        =   23
         Top             =   2160
         Width           =   975
      End
      Begin VB.Line Line4 
         Index           =   0
         X1              =   2280
         X2              =   4200
         Y1              =   1680
         Y2              =   1680
      End
      Begin VB.Label Label10 
         Caption         =   "años de vida"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   1
         Left            =   2760
         TabIndex        =   22
         Top             =   1800
         Width           =   1065
      End
      Begin VB.Label Label10 
         Caption         =   "Valor adquisición"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   2640
         TabIndex        =   21
         Top             =   1440
         Width           =   1425
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "TABLAS"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   285
         Index           =   0
         Left            =   240
         TabIndex        =   20
         Top             =   960
         Width           =   975
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00000080&
         X1              =   5280
         X2              =   2040
         Y1              =   720
         Y2              =   720
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00000080&
         X1              =   1200
         X2              =   1800
         Y1              =   720
         Y2              =   720
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Tipos de amortización"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   375
         Left            =   1560
         TabIndex        =   19
         Top             =   360
         Width           =   3450
      End
   End
   Begin VB.Frame FrameeMPRESAS 
      BorderStyle     =   0  'None
      Height          =   5415
      Left            =   0
      TabIndex        =   43
      Top             =   0
      Width           =   5535
      Begin VB.CommandButton cmdEmpresa 
         Caption         =   "Salir"
         Height          =   375
         Index           =   1
         Left            =   4320
         TabIndex        =   47
         Top             =   4800
         Width           =   975
      End
      Begin VB.CommandButton cmdEmpresa 
         Caption         =   "Regresar"
         Height          =   375
         Index           =   0
         Left            =   3240
         TabIndex        =   46
         Top             =   4800
         Width           =   975
      End
      Begin MSComctlLib.ListView lwE 
         Height          =   3615
         Left            =   120
         TabIndex        =   45
         Top             =   840
         Width           =   5295
         _ExtentX        =   9340
         _ExtentY        =   6376
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         HideColumnHeaders=   -1  'True
         Checkboxes      =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "dsdsd"
            Object.Width           =   7056
         EndProperty
      End
      Begin VB.Label Label7 
         Caption         =   "Empresas en el sistema"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   495
         Index           =   1
         Left            =   120
         TabIndex        =   44
         Top             =   240
         Width           =   5295
      End
   End
   Begin MSComDlg.CommonDialog cd1 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame frameCalculoSaldos 
      BorderStyle     =   0  'None
      Height          =   5535
      Left            =   0
      TabIndex        =   12
      Top             =   0
      Width           =   6975
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   1320
         Top             =   5760
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   3
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMensajes.frx":0F4D
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMensajes.frx":673F
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMensajes.frx":7151
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   375
         Left            =   120
         TabIndex        =   17
         Top             =   4800
         Visible         =   0   'False
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   661
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Salir"
         Height          =   375
         Index           =   1
         Left            =   5640
         TabIndex        =   16
         Top             =   4800
         Width           =   1095
      End
      Begin VB.CommandButton Command2 
         Caption         =   "&Iniciar"
         Height          =   375
         Left            =   4200
         TabIndex        =   15
         Top             =   4800
         Width           =   1215
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   3735
         Left            =   120
         TabIndex        =   14
         Top             =   840
         Width           =   6735
         _ExtentX        =   11880
         _ExtentY        =   6588
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         _Version        =   393217
         SmallIcons      =   "ImageList1"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Nivel"
            Object.Width           =   2699
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   1
            Text            =   "Debe"
            Object.Width           =   2999
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   2
            Text            =   "Haber"
            Object.Width           =   2999
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   3
            Text            =   "Saldo"
            Object.Width           =   2999
         EndProperty
      End
      Begin VB.Label Label7 
         Caption         =   "Cálculo de saldos"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   495
         Index           =   0
         Left            =   120
         TabIndex        =   13
         Top             =   240
         Width           =   5295
      End
   End
   Begin VB.Frame frameMultibase 
      BorderStyle     =   0  'None
      Height          =   5775
      Left            =   0
      TabIndex        =   129
      Top             =   0
      Width           =   6015
      Begin VB.CheckBox chkMultibase 
         Caption         =   "Conceptos"
         Height          =   255
         Index           =   5
         Left            =   3000
         TabIndex        =   175
         Top             =   2640
         Value           =   1  'Checked
         Width           =   2055
      End
      Begin VB.TextBox txtFecha 
         Height          =   285
         Index           =   1
         Left            =   4320
         TabIndex        =   145
         Text            =   "Text7"
         Top             =   4200
         Width           =   1095
      End
      Begin VB.TextBox txtFecha 
         Height          =   285
         Index           =   0
         Left            =   1320
         TabIndex        =   143
         Text            =   "Text7"
         Top             =   4200
         Width           =   1095
      End
      Begin VB.CheckBox chkMultibase 
         Caption         =   "Facturas proveedores"
         Height          =   255
         Index           =   4
         Left            =   3000
         TabIndex        =   137
         Top             =   3600
         Width           =   2295
      End
      Begin VB.CheckBox chkMultibase 
         Caption         =   "Facturas clientes"
         Height          =   255
         Index           =   3
         Left            =   720
         TabIndex        =   136
         Top             =   3600
         Width           =   2055
      End
      Begin VB.CheckBox chkMultibase 
         Caption         =   "Histórico apuntes cerrados"
         Height          =   255
         Index           =   2
         Left            =   3000
         TabIndex        =   135
         Top             =   3180
         Width           =   2655
      End
      Begin VB.CheckBox chkMultibase 
         Caption         =   "Histórico apuntes"
         Height          =   255
         Index           =   1
         Left            =   720
         TabIndex        =   134
         Top             =   3180
         Value           =   1  'Checked
         Width           =   3735
      End
      Begin VB.CheckBox chkMultibase 
         Caption         =   "Cuentas"
         Height          =   255
         Index           =   0
         Left            =   720
         TabIndex        =   133
         Top             =   2640
         Value           =   1  'Checked
         Width           =   2055
      End
      Begin VB.CommandButton cmdMultiBase 
         Caption         =   "Salir"
         Height          =   375
         Index           =   1
         Left            =   4680
         TabIndex        =   132
         Top             =   5160
         Width           =   1095
      End
      Begin VB.CommandButton cmdMultiBase 
         Caption         =   "Aceptar"
         Height          =   375
         Index           =   0
         Left            =   3480
         TabIndex        =   131
         Top             =   5160
         Width           =   1095
      End
      Begin VB.Label Label35 
         Alignment       =   1  'Right Justify
         Caption         =   "Fecha Fin"
         Height          =   255
         Index           =   1
         Left            =   3120
         TabIndex        =   146
         Top             =   4200
         Width           =   975
      End
      Begin VB.Label Label35 
         Caption         =   "Fecha Inicio"
         Height          =   255
         Index           =   0
         Left            =   360
         TabIndex        =   144
         Top             =   4200
         Width           =   1095
      End
      Begin VB.Line Line5 
         X1              =   240
         X2              =   5640
         Y1              =   3000
         Y2              =   3000
      End
      Begin VB.Label Label34 
         Caption         =   "Label34"
         Height          =   255
         Left            =   240
         TabIndex        =   142
         Top             =   4800
         Width           =   5535
      End
      Begin VB.Label Label33 
         Caption         =   "Datos a revisar:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   240
         TabIndex        =   141
         Top             =   2280
         Width           =   4815
      End
      Begin VB.Label Label32 
         Caption         =   "A este proceso le puede costar mucho tiempo."
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
         Left            =   360
         TabIndex        =   140
         Top             =   1680
         Width           =   4815
      End
      Begin VB.Label Label31 
         Caption         =   "No debe trabajar nadie en esta empresa"
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
         Left            =   360
         TabIndex        =   139
         Top             =   1320
         Width           =   4815
      End
      Begin VB.Label Label30 
         Caption         =   "Utlidad para revisar los caracteres especiales que puedan quedar al realizar integraciones. "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   360
         TabIndex        =   138
         Top             =   720
         Width           =   5775
      End
      Begin VB.Label Label29 
         Caption         =   "Revisión caracteres multibase"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   375
         Left            =   720
         TabIndex        =   130
         Top             =   120
         Width           =   4935
      End
   End
   Begin VB.Frame framaLlevarFacturas 
      Height          =   4455
      Left            =   0
      TabIndex        =   160
      Top             =   0
      Width           =   5775
      Begin VB.Frame FrameImportarFechas 
         Height          =   1455
         Left            =   120
         TabIndex        =   165
         Top             =   1560
         Width           =   5535
         Begin VB.TextBox Text8 
            Height          =   285
            Left            =   120
            TabIndex        =   170
            Text            =   "Text8"
            Top             =   600
            Width           =   5295
         End
         Begin VB.Image Image5 
            Height          =   240
            Left            =   840
            Picture         =   "frmMensajes.frx":75A3
            Top             =   360
            Width           =   240
         End
         Begin VB.Label Label39 
            Caption         =   "Fichero"
            Height          =   255
            Left            =   120
            TabIndex        =   171
            Top             =   360
            Width           =   735
         End
      End
      Begin VB.TextBox txtFecha 
         Height          =   285
         Index           =   3
         Left            =   4200
         TabIndex        =   168
         Text            =   "Text7"
         Top             =   2160
         Width           =   1095
      End
      Begin VB.TextBox txtFecha 
         Height          =   285
         Index           =   2
         Left            =   1320
         TabIndex        =   166
         Text            =   "Text7"
         Top             =   2160
         Width           =   1095
      End
      Begin VB.CheckBox chkImportarFacturas 
         Caption         =   "Eliminar ficheros"
         Height          =   255
         Left            =   240
         TabIndex        =   164
         Top             =   3960
         Width           =   2535
      End
      Begin VB.CommandButton cmdImportarFacuras 
         Caption         =   "Salir"
         Height          =   375
         Index           =   1
         Left            =   4320
         TabIndex        =   163
         Top             =   3840
         Width           =   1095
      End
      Begin VB.CommandButton cmdImportarFacuras 
         Caption         =   "Aceptar"
         Height          =   375
         Index           =   0
         Left            =   3000
         TabIndex        =   162
         Top             =   3840
         Width           =   1095
      End
      Begin VB.Line Line6 
         BorderWidth     =   3
         X1              =   120
         X2              =   5640
         Y1              =   720
         Y2              =   720
      End
      Begin VB.Label Label38 
         Caption         =   "Label38"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   1
         Left            =   240
         TabIndex        =   174
         Top             =   840
         Width           =   5415
      End
      Begin VB.Label Label40 
         Caption         =   "Label40"
         Height          =   255
         Left            =   120
         TabIndex        =   172
         Top             =   3240
         Width           =   5295
      End
      Begin VB.Label Label35 
         Caption         =   "Fecha Inicio"
         Height          =   255
         Index           =   3
         Left            =   3240
         TabIndex        =   169
         Top             =   2160
         Width           =   1095
      End
      Begin VB.Label Label35 
         Caption         =   "Fecha Inicio"
         Height          =   255
         Index           =   2
         Left            =   360
         TabIndex        =   167
         Top             =   2160
         Width           =   1095
      End
      Begin VB.Label Label38 
         Alignment       =   2  'Center
         Caption         =   "Label38"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   0
         Left            =   240
         TabIndex        =   161
         Top             =   240
         Width           =   5295
      End
   End
   Begin VB.Label Label11 
      Caption         =   "="
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
      Index           =   2
      Left            =   1800
      TabIndex        =   31
      Top             =   600
      Width           =   120
   End
   Begin VB.Line Line4 
      Index           =   2
      X1              =   2040
      X2              =   3960
      Y1              =   720
      Y2              =   720
   End
   Begin VB.Label Label10 
      Caption         =   "años de vida"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   5
      Left            =   2520
      TabIndex        =   30
      Top             =   840
      Width           =   1065
   End
   Begin VB.Label Label10 
      Caption         =   "Valor adquisición"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   4
      Left            =   2400
      TabIndex        =   29
      Top             =   480
      Width           =   1425
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "TABLAS"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   285
      Index           =   2
      Left            =   0
      TabIndex        =   28
      Top             =   0
      Width           =   975
   End
End
Attribute VB_Name = "frmMensajes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Opcion As Byte
    '1.- Saldos historico
    '2.- Comprobar saldos
    '3.- Mostrar tipos de amortizacion
    '4.- Seleccionar empresas
    
    '5.- Es, como si fuera comprobar saldos , pero se lanza y se cierra autmaticamente
    '6.- El acerca DE
    
    '7.- Nueva linea en configuracion balances
    '8.- Modificar linea balances
    
    
    '9.- Nueva CTA de configuracion balances
    '10- MODIFICAR  "   "             "
        
    '11- Carta modelo 347
    '12- Asientos con saltos y/o repetidos
    
    '13- Importar datos fiscales de las cuentas
    
    '14-  Cambios caracteres multibase
    
    '15- Cambio Password
    
        
    '16- Traspaso de facturas entre PC's. EXP
    '17-   "                  "           IMPORTAR
    
    
     '18- Importes punteo
    
     '19- Copiar de un balance a OTRO
    
    
Public Parametros As String
    '1.- Vendran empipados: Cuenta, PunteadoD, punteadoH, pdteD,PdteH

Private WithEvents frmC As frmColCtas
Attribute frmC.VB_VarHelpID = -1
Private PrimeraVez As Boolean

Dim I As Integer
Dim SQL As String
Dim RS As Recordset
Dim ItmX As ListItem
Dim Errores As String
Dim NE As Integer
Dim OK As Integer

Private Sub cmd347_Click(Index As Integer)
    If Index = 0 Then
        If Not GuardarDatosCarta Then Exit Sub
    End If
    Unload Me
End Sub

Private Sub cmdBalance_Click(Index As Integer)
    If Index = 1 Then
        CadenaDesdeOtroForm = ""
        Unload Me
    Else
        If Text1(0).Text = "" Then
            MsgBox "Primer campo obligatorio", vbExclamation
            Exit Sub
        End If
        If InsertarModificar Then Unload Me
    End If
End Sub

Private Sub cmdCabError_Click(Index As Integer)
Dim RS As ADODB.Recordset
Dim J As Long
Dim ii As Long
Dim Anyo As Integer

    If Index = 1 Then
        Unload Me
    Else
        Screen.MousePointer = vbHourglass
        Anyo = 0
        I = 0
        Do
          
            SQL = "select numasien,fechaent from hcabapu where fechaent >= '"
            SQL = SQL & Format(DateAdd("yyyy", Anyo, vParam.fechaini), FormatoFecha)
            SQL = SQL & "' AND fechaent <= '" & Format(DateAdd("yyyy", Anyo, vParam.fechafin), FormatoFecha) & "' ORDER By NumAsien"
            Set RS = New ADODB.Recordset
            RS.Open SQL, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
            ii = 0
            While Not RS.EOF
               J = RS.Fields(0)
               'Igual
                If J - ii = 0 Then
                    
                    SQL = Format(J, "00000")
                    SQL = SQL & "  -  " & Format(RS!fechaent, "dd/mm/yyyy")
                    Text5.Text = Text5.Text & SQL & vbCrLf
                    I = I + 1
                Else
                    If J - ii > 1 Then
                        If J - ii = 2 Then
                            SQL = Format(J - 1, "00000")
                        Else
                            SQL = "Entre " & Format(ii, "00000") & "  y  " & Format(J, "00000")
                        End If
                        SQL = SQL & " (" & CStr(Year(vParam.fechaini) + Anyo) & ")"
                        Text6.Text = Text6.Text & SQL & vbCrLf
                        I = I + 1
                    End If
                End If
                ii = J
                'Refrescamos
                If I > 50 Then
                    Text5.Refresh
                    Text6.Refresh
                    I = 0
                End If
                
                '
                RS.MoveNext
            Wend
            RS.Close
            Anyo = Anyo + 1
        Loop Until Anyo > 1
        Me.Refresh
        Screen.MousePointer = vbDefault
        cmdCabError(0).Enabled = False
    End If
End Sub

Private Sub cmdCambioPwd_Click(Index As Integer)
    If Index = 1 Then
        Unload Me
        Exit Sub
    End If
    
    For I = 1 To Text7.Count - 1
        Text7(I).Text = Trim(Text7(I).Text)
        If Text7(I).Text = "" Then
            MsgBox "Hay que rellenar todos los campos", vbExclamation
            Exit Sub
        End If
    Next I
    
    
    'Todos rellenados
    'Ha puesto la clave actual real
    If Text7(1).Text <> vUsu.PasswdPROPIO Then
        MsgBox "Clave actual incorrecta", vbExclamation
        Exit Sub
    End If
    
    If Text7(2).Text <> Text7(3).Text Then
        MsgBox "Mal reescrita la clave nueva", vbExclamation
        Exit Sub
    End If
    
    
    If InStr(1, Text7(2).Text, "'") > 0 Then
        MsgBox "Clave nueva contiene caracter no permitido", vbExclamation
        Exit Sub
    End If
    
    
    'UPDATEAMOS
    On Error Resume Next
    SQL = "UPDATE Usuarios.Usuarios Set passwordpropio='" & Text7(2).Text & "' WHERE codusu = " & vUsu.Codigo
    conn.Execute SQL
    If Err.Number <> 0 Then
        MuestraError Err.Number, "Cambio clave"
    Else
        vUsu.PasswdPROPIO = Text7(2).Text
        MsgBox "Cambio de clave realizado con éxito", vbInformation
        Unload Me
    End If
End Sub

Private Sub cmdCtaBalan_Click(Index As Integer)
    If Index = 1 Then
        CadenaDesdeOtroForm = ""
    Else
        If Text3.Text = "" Then
            MsgBox "la cuenta no puede estar en blanco", vbExclamation
            Exit Sub
        End If
        If Not IsNumeric(Text3.Text) Then
            MsgBox "La cuenta debe ser numérica", vbExclamation
            Exit Sub
        End If
        'Esto es el OPTION
        SQL = ""
        For I = 0 To 2
            If Option1(I).Value Then SQL = SQL & Mid(Option1(I).Caption, 1, 1)
        Next I
        If SQL = "" Then
            MsgBox "Seleccione una opción de la cuenta (Saldo - Debe - Haber )"
            Exit Sub
        End If
        
        'RESTA y la resta
        SQL = SQL & "|" & Abs(Me.chkResta.Value)
        CadenaDesdeOtroForm = Text3.Text & "|" & SQL & "|"
    End If
    Unload Me
End Sub

Private Sub cmdEmpresa_Click(Index As Integer)
    CadenaDesdeOtroForm = ""
    If Index = 0 Then
        SQL = ""
        Parametros = ""
        For I = 1 To lwE.ListItems.Count
            If Me.lwE.ListItems(I).Checked Then
                SQL = SQL & Me.lwE.ListItems(I).Text & "|"
                Parametros = Parametros & "1" 'Contador
            End If
        Next I
        CadenaDesdeOtroForm = Len(Parametros) & "|" & SQL
        'Vemos las conta
        SQL = ""
        For I = 1 To lwE.ListItems.Count
            If Me.lwE.ListItems(I).Checked Then
                SQL = SQL & Me.lwE.ListItems(I).Tag & "|"
            End If
        Next I
        CadenaDesdeOtroForm = CadenaDesdeOtroForm & SQL
    End If
    Unload Me
End Sub

Private Sub cmdImpCta_Click(Index As Integer)
Dim Cad As String

    If Index = 1 Then
        Unload Me
        Exit Sub
    End If

    txtImpCta.Text = Trim(txtImpCta.Text)
    
    If txtImpCta.Text = "" Then Exit Sub
    
    If Dir(txtImpCta.Text) = "" Then
        MsgBox "El fichero: " & txtImpCta.Text & " NO existe.", vbExclamation
        Exit Sub
    End If
    
    Cad = "Seguro que desa continuar con la importación de los datos fiscales?"
    If MsgBox(Cad, vbQuestion + vbYesNoCancel + vbDefaultButton2) <> vbYes Then Exit Sub
    
    Screen.MousePointer = vbHourglass
    'Ya esta preparado comprobemos k podemos abrir la conexion

        lblimpCta2.Caption = "Lineas"
        lblImpCta.Caption = "0"
        Errores = ""
        NE = 0
        OK = 0
        Me.Refresh
        CadenaDesdeOtroForm = ""
        'La contabilidad existe
        HacerImportacion
        'Si hay errores
        If NE > 0 Then
            Errores = OK & " lineas pasadas con exito." & vbCrLf & vbCrLf & Errores
            ImprimeFichero
        Else
            MsgBox OK & " lineas pasadas con exito", vbInformation
        End If
        CadenaDesdeOtroForm = ""
    Screen.MousePointer = vbDefault
    lblimpCta2.Caption = ""
    lblImpCta.Caption = ""
End Sub

Private Sub cmdImportarFacuras_Click(Index As Integer)
    If Index = 1 Then
        Unload Me
        Exit Sub
    End If
    
    'COMPROBACIONES
    If Opcion = 17 Then
        If Text8.Text = "" Then
            MsgBox "Fichero en blanco", vbExclamation
            Exit Sub
        End If
        
        If Dir(Text8.Text, vbArchive) = "" Then
            MsgBox "Fichero no existe.", vbExclamation
            Exit Sub
        End If
        
        SQL = "Va a realizar la importación de datos  de facturas en la empresa: " & vbCrLf & vbCrLf & vEmpresa.nomEmpre
        SQL = SQL & "(" & vEmpresa.nomresum & ") - Conta: " & vEmpresa.codEmpre & vbCrLf & vbCrLf & "¿Desea continuar?"
        If MsgBox(SQL, vbQuestion + vbYesNoCancel + vbDefaultButton2) <> vbYes Then Exit Sub
    End If
    
    
    If Opcion = 16 Then
        ExportarFactur
    Else
        ImportarFicheroFac
    End If
End Sub


Private Sub ExportarFactur()
    On Error GoTo EExportarDatosF
        'Primero borramos el temporal facturas
        Errores = App.path & "\tmpexpdatos.tmp"
        If Dir(Errores, vbArchive) <> "" Then Kill Errores
        NE = FreeFile
        Open Errores For Output As NE
        'Primero proveedores
        ExportarDatosFacturas True
        'Clientes
        ExportarDatosFacturas False
        Close NE
        
        
        
        Text8.Text = ""
        Image4_Click
        If Text8.Text <> "" Then
        'Hay k copiar el archivo
            Errores = App.path & "\tmpexpdatos.tmp"
            CopiarArchivo
        Else
           ' MsgBox "Opcion cancelada", vbExclamation
        End If
        
        
        
        
        Exit Sub



EExportarDatosF:
    MuestraError Err.Number, Err.Description
    On Error Resume Next
    Close NE
End Sub


Private Sub cmdMultiBase_Click(Index As Integer)
Dim I As Integer
    If Index = 1 Then
        Unload Me
        Exit Sub
    End If
    
    
    'Comprobamos k ha selecionado algun nivel
    NE = 0
    For I = 0 To Me.chkMultibase.Count - 1
        If Me.chkMultibase(I).Value = 1 Then NE = NE + 1
    Next I
    If NE = 0 Then
        MsgBox "Seleccione donde se van a realizar lso cambios", vbExclamation
        Exit Sub
    End If
    
    'Comprobacion si hay alguien trabajando
    If UsuariosConectados Then Exit Sub
    
    SQL = "Seguro que desea continuar con el proceso"
    If MsgBox(SQL, vbCritical + vbYesNoCancel) <> vbYes Then Exit Sub
    
   'BLOQUEAMOS LA BD
   If Not Bloquear_DesbloquearBD(True) Then
        MsgBox "No se ha podido bloquea a nivel de BD.", vbExclamation
        Exit Sub
    End If
    
    
    Screen.MousePointer = vbHourglass
    NumRegElim = 0
    For I = 0 To Me.chkMultibase.Count - 1
        If Me.chkMultibase(I).Value = 1 Then
            'Hacemos los cambios para ese valor
            HacerCambios I
        End If
    Next I
    Bloquear_DesbloquearBD False
    Screen.MousePointer = vbDefault
    Label34.Caption = ""
    SQL = "Proceso finalizado" & vbCrLf & "Se han realizado: " & NumRegElim & " cambio(s)."
    MsgBox SQL, vbInformation
End Sub

Private Sub cmdPunteo_Click()
    Unload Me
End Sub

Private Sub Command1_Click(Index As Integer)
    Unload Me
End Sub

Private Sub Command2_Click()
Dim Digitos As Integer
    ListView1.ListItems.Clear
    Me.ProgressBar1.Value = 0
    Me.ProgressBar1.Max = vEmpresa.numNivel + 1
    Me.ProgressBar1.visible = True
    Screen.MousePointer = vbHourglass
    'Calculamos en historico
    CalculaSaldosNivel True, Digitos
    'Iniciamos el calculo de saldos para cada nivel
    For I = 1 To vEmpresa.numNivel
        Digitos = DigitosNivel(I)
        CalculaSaldosNivel False, Digitos
    Next I
    Me.ProgressBar1.visible = False
    Screen.MousePointer = vbDefault
End Sub

Private Sub Command3_Click()
    Unload Me
End Sub



Private Sub Form_Activate()
    If PrimeraVez Then
        PrimeraVez = False
        If Opcion = 5 Then
            Screen.MousePointer = vbHourglass
            Me.tCuadre.Enabled = True
        End If
    Else
        Screen.MousePointer = vbDefault
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
Dim W, H
    Me.tCuadre.Enabled = False
    PrimeraVez = True
    Me.frameSaldosHco.visible = False
    Me.frameCalculoSaldos.visible = False
    Me.frameamort.visible = False
    Me.FrameeMPRESAS.visible = False
    Me.frameAcercaDE.visible = False
    Me.frameCtasBalance.visible = False
    Me.FrameCarta347.visible = False
    Me.frameSaltos.visible = False
    frameBalance.visible = False
    FrameImpCta.visible = False
    Me.frameMultibase.visible = False
    Me.FrameCambioPWD.visible = False
    framaLlevarFacturas.visible = False
    Me.FrameImpPunteo.visible = False
    Select Case Opcion
    Case 1
        Me.Caption = "Cálculo de saldo"
        W = frameSaldosHco.Width
        H = Me.frameSaldosHco.Height
        Me.frameSaldosHco.visible = True
        
        CargaValoresHco
        Command1(0).Cancel = True
    Case 2
        Me.Caption = "Comprobacion saldos"
        W = Me.frameCalculoSaldos.Width
        H = Me.frameCalculoSaldos.Height + 150
        Me.frameCalculoSaldos.visible = True
        Command1(1).Enabled = True
        Command2.Enabled = True
    Case 3
        Me.Caption = "Información tipo amortización"
        W = Me.frameamort.Width
        H = Me.frameamort.Height + 200
        Me.frameamort.visible = True
    Case 4
        Me.Caption = "Seleccion"
        W = Me.FrameeMPRESAS.Width
        H = Me.FrameeMPRESAS.Height + 200
        Me.FrameeMPRESAS.visible = True
        cargaempresas
    Case 5
        'Lanzar automaticamente la comprobación de saldo
        Me.Caption = "Comprobacion saldos"
        W = Me.frameCalculoSaldos.Width
        H = Me.frameCalculoSaldos.Height
        Me.frameCalculoSaldos.visible = True
        Command1(1).Enabled = False
        Command2.Enabled = False
    Case 6
        CargaImagen
        Me.Caption = "Acerca de ....."
        W = Me.frameAcercaDE.Width
        H = Me.frameAcercaDE.Height + 200
        Me.frameAcercaDE.visible = True
        Label13.Caption = "Versión:  " & App.Major & "." & App.Minor & "." & App.Revision & " "
    Case 7, 8
        Me.Caption = "Lineas configuracion balance"
        W = Me.frameBalance.Width
        H = Me.frameBalance.Height + 300
        Me.frameBalance.visible = True
        PonerCamposBalance
    Case 9, 10
        If Opcion = 9 Then
            Me.cmdCtaBalan(0).Caption = "Insertar"
        Else
            Me.cmdCtaBalan(0).Caption = "Modificar"
        End If
        Me.Caption = "Cuentas configuracion balances"
        W = Me.frameCtasBalance.Width
        H = Me.frameCtasBalance.Height + 300
        frameCtasBalance.visible = True
        PonerCamposCtaBalance
    Case 11
        'Carta modelo 347
        Me.Caption = "Datos carta modelo 347"
        W = Me.FrameCarta347.Width
        H = Me.FrameCarta347.Height + 300
        Me.FrameCarta347.visible = True
        CargarDatosCarta
    Case 12
        'Saltos y repedtidos
        Me.Caption = "Búsqueda cabeceras asientos incorrectos"
        W = Me.frameSaltos.Width
        H = Me.frameSaltos.Height + 300
        Me.frameSaltos.visible = True
        Me.cmdCabError(0).Enabled = True
        Text5.Text = ""
        Text6.Text = ""
        cmdCabError(1).Cancel = True
    Case 13
        Me.Caption = "Importar datos fiscales de las cuentas"
        W = Me.FrameImpCta.Width
        H = Me.FrameImpCta.Height + 450
        Me.FrameImpCta.visible = True
        cmdImpCta(1).Cancel = True
        txtImpCta.Text = ""
        Me.lblImpCta.Caption = ""
        Me.lblimpCta2.Caption = ""
    Case 14
        'MULTIBASE
        Me.Caption = "Sustitución caracteres multibase"
        W = Me.frameMultibase.Width
        H = Me.frameMultibase.Height + 300
        Me.frameMultibase.visible = True
        Label34.Caption = ""
        txtFecha(0).Text = Format(vParam.fechaini, "dd/mm/yyyy")
        txtFecha(1).Text = Format(vParam.fechafin, "dd/mm/yyyy")
        cmdMultiBase(1).Cancel = True
    Case 15
        'Cambio password usuario
        Me.Caption = "Cambio password"
        W = Me.FrameCambioPWD.Width
        H = Me.FrameCambioPWD.Height + 300
        Me.FrameCambioPWD.visible = True
        Text7(0).Text = vUsu.Nombre
        For I = 1 To 3
            Text7(I).Text = ""
        Next I
        cmdCambioPwd(1).Cancel = True
    Case 16, 17
        Text8.Text = ""
        Caption = "UTIL. FACTURAS"
        W = Me.framaLlevarFacturas.Width
        H = Me.framaLlevarFacturas.Height + 300
        Me.framaLlevarFacturas.visible = True
        chkImportarFacturas.visible = Opcion = 17
        FrameImportarFechas.visible = Opcion = 17
'        optTraerFacturas(0).Visible = Opcion = 16
'        optTraerFacturas(1).Visible = Opcion = 16
        If Opcion = 16 Then
            Label38(0).Caption = "EXPORTAR"
        Else
            Label38(0).Caption = "IMPORTAR"
        End If
        Label38(1).Caption = vEmpresa.nomEmpre & "   (" & vEmpresa.nomresum & ")"
        Me.txtFecha(2).Text = Format(vParam.fechaini, "dd/mm/yyyy")
        Me.txtFecha(3).Text = Format(Now, "dd/mm/yyyy")
        Label40.Caption = ""
        cmdImportarFacuras(1).Cancel = True
        
    Case 18
        Me.FrameImpPunteo.visible = True
        Caption = "Importes"
        For I = 0 To 8
            Me.txtImporteP(I).Text = RecuperaValor(Parametros, I + 1)
        Next I
        W = Me.FrameImpPunteo.Width
        H = Me.FrameImpPunteo.Height + 300
        cmdPunteo.Cancel = True
    End Select
    Me.Width = W + 120
    Me.Height = H + 120
End Sub




Private Sub CargaValoresHco()
'Lo que hace es dado el parametro scamos la cuenta, nomcuenta, saldos
'1.- Vendran empipados: Cuenta, PunteadoD, punteadoH, pdteD,PdteH

Label6.Caption = RecuperaValor(Parametros, 1)
For I = 0 To 3
    Me.txtsaldo(I).Text = RecuperaValor(Parametros, I + 2)
Next I
CalculaSaldosFinales
End Sub



Private Sub CalculaSaldosFinales()
Dim Importe As Currency
    For I = 0 To 3
        Importe = ImporteFormateado(txtsaldo(I).Text)
        txtsaldo(I).Tag = Importe
    Next I
    txtsaldo(4).Text = ""
    txtsaldo(5).Text = ""
    
    Importe = CCur(txtsaldo(1).Tag) + CCur(txtsaldo(3).Tag)
    txtsaldo(5).Tag = Importe
    txtsaldo(5).Text = Format(Importe, FormatoImporte)
    Importe = CCur(txtsaldo(0).Tag) + CCur(txtsaldo(2).Tag)
    txtsaldo(4).Tag = Importe
    txtsaldo(4).Text = Format(Importe, FormatoImporte)
    Importe = CCur(txtsaldo(5).Tag) - CCur(txtsaldo(4).Tag)
    txtsaldo(6).Text = ""
    txtsaldo(7).Text = ""
    If Importe <> 0 Then
        If Importe > 0 Then
            txtsaldo(6).Text = Format(Importe, FormatoImporte)
        Else
            txtsaldo(7).Text = Format(Abs(Importe), FormatoImporte)
        End If
    End If
    
    
    'Ahora veremos si tiene del periodo
    txtsaldo(8).Text = ""
    txtsaldo(9).Text = ""
    SQL = RecuperaValor(Parametros, 6)
    If SQL = "" Then
        NE = 0
        
    Else
        NE = 1
        Importe = CCur(SQL)
        If Importe >= 0 Then
            txtsaldo(8).Text = Format(Importe, FormatoImporte)
        Else
            txtsaldo(9).Text = Format(Abs(Importe), FormatoImporte)
        End If
    End If
    
    Label28(1).visible = (NE = 1)
    txtsaldo(9).visible = (NE = 1)
    txtsaldo(8).visible = (NE = 1)
    
    'Descripcion cuenta
    SQL = Trim(RecuperaValor(Parametros, 7))   'Descripcion cuenta
    If SQL <> "" Then SQL = " - " & SQL
    Label6.Caption = Label6.Caption & SQL
    
    
    
    'NUEVO 14 Febrero... San valentin
    Importe = CCur(txtsaldo(2).Tag) - CCur(txtsaldo(3).Tag)
    Image6(0).ToolTipText = "Saldo punteado: " & Format(Importe, FormatoImporte)
    Importe = CCur(txtsaldo(0).Tag) - CCur(txtsaldo(1).Tag)
    Image6(1).ToolTipText = "Saldo pendiente: " & Format(Importe, FormatoImporte)
    
    
    
End Sub

Private Sub CalculaSaldosNivel(EnHistorico As Boolean, ByRef digit As Integer)
Dim Debe As Double
Dim Haber As Double
Dim TieneDatos As Boolean
Dim SubCad As String

Debe = 0: Haber = 0
If Not EnHistorico Then
    SQL = "SELECT sum(impmesde) as sumad,sum(impmesha)as sumah  from hsaldos where codmacta like '"
    SQL = SQL & Mid("__________", 1, digit) & "'"

    Set RS = New ADODB.Recordset
    RS.Open SQL, conn, adOpenKeyset, adLockOptimistic, adCmdText
    If RS.EOF Then
        TieneDatos = False
    Else
        TieneDatos = True
        Debe = 0: Haber = 0
        If Not IsNull(RS.Fields(0)) Then Debe = RS.Fields(0)
        If Not IsNull(RS.Fields(1)) Then Haber = RS.Fields(1)
    End If
    SubCad = "Digitos: " & digit
Else
    'Calculamos directamente sobre las lineas de hcoapuntes
    SubCad = "Apuntes"
    SQL = "Select SUM(timporteD) as impd,sum(timporteh) as imph from hlinapu"
    Set RS = New ADODB.Recordset
    RS.Open SQL, conn, adOpenKeyset, adLockOptimistic, adCmdText
    If RS.EOF Then
        TieneDatos = False
    Else
        TieneDatos = True
        If Not IsNull(RS.Fields(0)) Then Debe = RS.Fields(0)
        If Not IsNull(RS.Fields(1)) Then Haber = RS.Fields(1)
    End If
End If
RS.Close
Set RS = Nothing
Set ItmX = ListView1.ListItems.Add(, , SubCad)
If TieneDatos Then
    ItmX.SubItems(1) = Format(Debe, FormatoImporte)
    ItmX.SubItems(2) = Format(Haber, FormatoImporte)
    Debe = Debe - Haber
    ItmX.SubItems(3) = Format(Debe, FormatoImporte)
End If
'Comprobamos los importes
ItmX.SmallIcon = 1
If Debe = 0 Then
    'Comprobamos con el de arriba
    If ItmX.ListSubItems(1) = ListView1.ListItems(1).SubItems(1) Then
        If ItmX.ListSubItems(2) = ListView1.ListItems(1).SubItems(2) Then
            ItmX.SmallIcon = 2
        End If
    End If
    Else
        'Por si quiero poner otro ciocno
End If
If ItmX.SmallIcon <> 2 Then CadenaDesdeOtroForm = CadenaDesdeOtroForm & "M"
Me.ProgressBar1.Value = Me.ProgressBar1.Value + 1
End Sub


Private Sub cargaempresas()
Dim Prohibidas As String
On Error GoTo Ecargaempresas

    VerEmresasProhibidas Prohibidas
    
    SQL = "Select * from Usuarios.Empresas order by codempre"
    Set lwE.SmallIcons = Me.ImageList1
    lwE.ListItems.Clear
    Set RS = New ADODB.Recordset
    I = -1
    RS.Open SQL, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    While Not RS.EOF
        SQL = "|" & RS!codEmpre & "|"
        If InStr(1, Prohibidas, SQL) = 0 Then
            Set ItmX = lwE.ListItems.Add(, , RS!nomEmpre, , 3)
            ItmX.Tag = RS!codEmpre
            If ItmX.Tag = vEmpresa.codEmpre Then
                ItmX.Checked = True
                I = ItmX.Index
            End If
            ItmX.ToolTipText = RS!CONTA
        End If
        RS.MoveNext
    Wend
    RS.Close
    If I > 0 Then Set lwE.SelectedItem = lwE.ListItems(I)
Ecargaempresas:
    If Err.Number <> 0 Then MuestraError Err.Number, "Cargando datos empresas"
    Set RS = Nothing
End Sub

Private Sub VerEmresasProhibidas(ByRef VarProhibidas As String)

On Error GoTo EVerEmresasProhibidas
    VarProhibidas = "|"
    SQL = "Select codempre from Usuarios.usuarioempresa WHERE codusu = " & (vUsu.Codigo Mod 1000)
    SQL = SQL & " order by codempre"
    Set RS = New ADODB.Recordset
    RS.Open SQL, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    While Not RS.EOF
          VarProhibidas = VarProhibidas & RS!codEmpre & "|"
          RS.MoveNext
    Wend
    RS.Close
    Exit Sub
EVerEmresasProhibidas:
    MuestraError Err.Number, Err.Description & vbCrLf & " Consulte soporte técnico"
    Set RS = Nothing
End Sub




Private Sub frmC_DatoSeleccionado(CadenaSeleccion As String)
    Text3.Text = RecuperaValor(CadenaSeleccion, 1)
End Sub

Private Sub Image3_Click()
    Set frmC = New frmColCtas
    frmC.ConfigurarBalances = 1
    frmC.DatosADevolverBusqueda = "0|1"
    frmC.Show vbModal
    Set frmC = Nothing
End Sub



Private Sub Image4_Click()
   On Error GoTo ELee
    
    With cd1
        .CancelError = True
        If Opcion = 16 Then
            .DialogTitle = "DESTINO. Nuevo nombre de fichero."
        Else
            .DialogTitle = "Seleccione archivo importación"
        End If
        .InitDir = "C:\"
        .ShowOpen
        Select Case Opcion
        Case 16
            'A mano. Es para el fichero de gaurdar datos
            Text8.Text = .FileName
        Case 17
            Text8.Text = .FileName
        Case Else
            txtImpCta.Text = .FileName
        End Select
    End With
    
    Exit Sub
ELee:
    Err.Clear
End Sub

Private Sub Image5_Click()
    Image4_Click
End Sub

Private Sub Image6_Click(Index As Integer)
    MsgBox Image6(Index).ToolTipText, vbInformation
End Sub

Private Sub ImageAyudaImpcta_Click()
    'Ejemplo
    '43000001|SECUVE, S.L.|RIU VERT  N§ 7|46600|ALZIRA|VALENCIA|B97301808|
    SQL = "Formato para la importación de datos fiscales. " & vbCrLf & vbCrLf & vbCrLf
    SQL = SQL & "El fichero vendrá con cada campo separados por PIPES." & vbCrLf
    SQL = SQL & "Codigo cta contable |" & vbCrLf
    SQL = SQL & "Descripcion |" & vbCrLf
    SQL = SQL & "Direccion |" & vbCrLf
    SQL = SQL & "Cod. Postal |" & vbCrLf
    SQL = SQL & "Poblacion |" & vbCrLf
    SQL = SQL & "Provincia |" & vbCrLf
    SQL = SQL & "NIF|" & vbCrLf
    SQL = SQL & "Cta bancaria:   ENTIDAD|" & vbCrLf
    SQL = SQL & "Cta bancaria:   OFICINA|" & vbCrLf
    SQL = SQL & "Cta bancaria:   CC|" & vbCrLf
    SQL = SQL & "Cta bancaria:   CUENTA|" & vbCrLf
    SQL = SQL & "347:    0.- No    1.- Si|" & vbCrLf
    MsgBox SQL, vbInformation
End Sub

Private Sub Label13_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        If Shift And vbCtrlMask > 0 Then
            MsgBox "HOLITA VECINO. Has encontrado el huevo de pascua...., a curraaaaaarrr", vbExclamation
        End If
    End If
End Sub

Private Sub Option1_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
   ' If KeyAscii = 13 Then
   '     KeyAscii = 0
   '     SendKeys "{tab}"
   ' End If
End Sub

Private Sub tCuadre_Timer()
    tCuadre.Enabled = False
    Screen.MousePointer = vbHourglass
    Command2_Click
    Me.ListView1.Refresh
    Screen.MousePointer = vbHourglass
    espera 2
    Unload Me
End Sub


Private Sub CargaImagen()
On Error Resume Next
    Image2.Picture = LoadPicture(App.path & "\minilogo.bmp")
    'Image1.Picture = LoadPicture(App.path & "\fondon.gif")
    Err.Clear
End Sub


Private Sub PonerCamposBalance()
    If Opcion = 7 Then
        Label16.Caption = "NUEVO"
        Label16.ForeColor = &H800000
        Me.chkPintar.Value = 1
        For I = 0 To 3
            Text1(I).Text = ""
        Next I

    Else
        'NumBalan|Pasivo|codigo|padre|Orden|tipo|deslinea|texlinea|formula|TienenCtas|Negrita|LibroCD|
        Text1(0).Text = RecuperaValor(Parametros, 7)
        Text1(1).Text = RecuperaValor(Parametros, 8)
        I = Val(RecuperaValor(Parametros, 10))
        If I = 1 Then
            'Tiene cuentas
            Text1(2).Text = ""
            Text1(2).Enabled = False
        Else
            Text1(2).Text = RecuperaValor(Parametros, 9)
        End If
        I = Val(RecuperaValor(Parametros, 11))
        chkNegrita.Value = I
        I = Val(RecuperaValor(Parametros, 12))
        chkCero.Value = I
        I = Val(RecuperaValor(Parametros, 13))
        chkPintar.Value = I
        Text1(3).Text = RecuperaValor(Parametros, 14)
    End If
End Sub



Private Sub PonerCamposCtaBalance()
    'EL grupo se le pasa siempre
    Text2.Text = RecuperaValor(Parametros, 1)
    
    
    If Opcion = 9 Then
        Label19.Caption = "NUEVO"
        Label19.ForeColor = &H800000
        Text3.Text = ""
        Text3.Enabled = True
        chkResta.Value = 0
    Else
        Text3.Enabled = False
        Text3.Text = RecuperaValor(Parametros, 2)
        I = Val(RecuperaValor(Parametros, 3))
        Option1(I).Value = True
        I = Val(RecuperaValor(Parametros, 4))
        chkResta.Value = I
    End If
End Sub




Private Function InsertarModificar() As Boolean
Dim Aux As String

On Error GoTo EInse

    InsertarModificar = False
    
    'Comprobamos el concpeto del libro a CD
     Text1(3).Text = UCase(Trim(Text1(3).Text))
    If Text1(3).Text <> "" Then
        If Not IsNumeric(Text1(3).Text) Then
            MsgBox "El campo 'Concepto Libro CD' debe ser numérico", vbExclamation
            Exit Function
        End If
    End If
    
    'Hay k comprobar, si tiene formula k sea correcta
    Text1(2).Text = UCase(Trim(Text1(2).Text))
    If Text1(2).Text <> "" Then
        SQL = CompruebaFormulaConfigBalan(CInt(RecuperaValor(Parametros, 1)), Text1(2).Text)
        If SQL <> "" Then
            MsgBox SQL, vbExclamation
            Exit Function
        End If
    End If
    If Opcion = 7 Then
        SQL = "INSERT INTO sperdid (NumBalan, Pasivo, codigo, padre, "
        SQL = SQL & "Orden, tipo, deslinea, texlinea, formula, TienenCtas, Negrita,A_Cero,Pintar,LibroCD) VALUES ("
        SQL = SQL & RecuperaValor(Parametros, 1)  'Numero
        SQL = SQL & ",'" & RecuperaValor(Parametros, 2) 'pasivo
        SQL = SQL & "'," & RecuperaValor(Parametros, 3)  'Codigo
        Aux = RecuperaValor(Parametros, 4) 'padre
        If Aux = "" Then
            Aux = ",NULL,"
        Else
            Aux = ",'" & Aux & "',"
        End If
        SQL = SQL & Aux
        SQL = SQL & RecuperaValor(Parametros, 5)
        If Text1(2).Text = "" Then
            Aux = "0"
        Else
            Aux = "1"
        End If
        SQL = SQL & "," & Aux
        SQL = SQL & ",'" & Text1(0).Text 'Text linea
        SQL = SQL & "','" & Text1(1).Text 'Desc linea
        SQL = SQL & "','" & Text1(2).Text 'Formula
        SQL = SQL & "',0," & chkNegrita.Value
        SQL = SQL & "," & Me.chkCero.Value
        SQL = SQL & "," & Me.chkPintar.Value
        SQL = SQL & ",'" & Text1(3).Text 'Libro CD
        SQL = SQL & "')"
    Else
        'Modificar
        'NumBalan|Pasivo|codigo|padre|Orden|tipo|deslinea|texlinea|formula|TienenCtas|Negrita|
        SQL = "UPDATE sperdid SET "
        SQL = SQL & "deslinea='" & Text1(0).Text & "',"
        SQL = SQL & "texlinea='" & Text1(1).Text & "',"
        SQL = SQL & "formula='" & Text1(2).Text & "',"
        If Text1(2).Text = "" Then
            Aux = "0"
        Else
            Aux = "1"
        End If
        SQL = SQL & "Tipo =" & Aux & ","
        SQL = SQL & "Negrita = " & chkNegrita.Value
        SQL = SQL & ", A_Cero = " & Me.chkCero.Value
        SQL = SQL & ", Pintar = " & Me.chkPintar.Value
        SQL = SQL & ", LibroCD = '" & Text1(3).Text & "'"
        SQL = SQL & " WHERE numbalan =" & RecuperaValor(Parametros, 1)
        SQL = SQL & " AND Pasivo = '" & RecuperaValor(Parametros, 2)
        SQL = SQL & "' AND codigo = " & RecuperaValor(Parametros, 3)
        
    End If
    conn.Execute SQL
    InsertarModificar = True
    'Ha insertado
    'Devuelve el texto, el texto auxiliar, y si es formula o no, descripcion cta y concepto oficial
    CadenaDesdeOtroForm = Text1(0).Text & "|" & Text1(1).Text & "|" & Aux & "|"
    CadenaDesdeOtroForm = CadenaDesdeOtroForm & Text1(3).Text & "|"
    Exit Function
EInse:
    MuestraError Err.Number
End Function





Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
KEYpress KeyAscii
'    If KeyAscii = 13 Then
'        KeyAscii = 0
'        SendKeys "{tab}"
'    End If
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
KEYpress KeyAscii
'    If KeyAscii = 13 Then
'        KeyAscii = 0
'        SendKeys "{tab}"
'    End If
End Sub

'--------------------------------------------------------------------
'
'       Carta IVA Clientes
'
Private Sub CargarDatosCarta()
Dim Limpiar As Boolean

    On Error GoTo ECargarDatos
    Limpiar = True
    SQL = App.path & "\txt347.dat"
    If Dir(SQL) <> "" Then
        'Vamos a ir leyendo , y devoviendo cadena
        I = FreeFile
        Open SQL For Input As #I
        For NumRegElim = 0 To Text4.Count - 1
            'Obtenemos la cadena
           LeerCadenaFicheroTexto    'lo guarda en SQL
           Text4(NumRegElim) = SQL
        Next NumRegElim
        Close #I
        Limpiar = False
    End If
ECargarDatos:
    If Err.Number <> 0 Then MuestraError Err.Number, "Carga fichero. " & Err.Description
    If Limpiar Then
        'No existe el fihero de configuracion
        For I = 0 To Text4.Count - 1
            Text4(I).Text = ""
        Next I
    End If
End Sub


Private Sub LeerCadenaFicheroTexto()
On Error GoTo ELeerCadenaFicheroTexto
    'Son dos lineas. La primaera indica k campo y la segunda el valor
    Line Input #I, SQL
    Line Input #I, SQL
    Exit Sub
ELeerCadenaFicheroTexto:
    SQL = ""
    Err.Clear
End Sub


Private Function GuardarDatosCarta()
    On Error GoTo Eguardardatoscarta
    SQL = App.path & "\txt347.dat"
    I = FreeFile
    Open SQL For Output As #I
    For NumRegElim = 0 To Text4.Count - 1
        Print #I, Text4(NumRegElim).Tag
        Print #I, Text4(NumRegElim).Text
    Next NumRegElim
    Close #I
    Exit Function
Eguardardatoscarta:
    MuestraError Err.Number, "guardar datos carta"
End Function

Private Sub KEYpress(ByRef KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        SendKeys "{tab}"
    End If
End Sub




'------------- IMportar datos fiscales


Private Sub HacerImportacion()
Dim NF As Integer
Dim LINEA As String

    On Error GoTo EHacere
    


    'Abrimos el fichero
    NF = FreeFile
    Open txtImpCta.Text For Input As #NF
    
  
    
    'Vamos linea a linea
    While Not EOF(NF)
        Line Input #NF, LINEA
        LINEA = Trim(LINEA)
        If LINEA <> "" Then
            If ProcesarLinea(LINEA) Then OK = OK + 1
        End If
        lblImpCta.Caption = CStr(NE + OK)
        lblImpCta.Refresh
    Wend
    
    'Cerramos
    Close (NF)
    
    Exit Sub
EHacere:
    MsgBox Err.Description, vbExclamation
End Sub





Private Function ProcesarLinea(LINEA As String) As Boolean
'Dim Valores(6) As String
Dim Valores(11) As String
Dim I As Integer
Dim Cad As String
Dim Crear As Boolean

    On Error GoTo EProcesarLinea
    ProcesarLinea = False
    
    
    'Orden en el k llegan
    For I = 0 To 11
        Valores(I) = RecuperaValor(LINEA, I + 1)
    Next I
    
    'TRIM
    For I = 0 To 11
        Valores(I) = Trim(Valores(I))
    Next I
    
    'Comprobaciones
    '-----------------
    If Valores(0) = "" Or Valores(1) = "" Or Valores(6) = "" Then
        'Ni cta, ni nombre cta, ni NIF pueden ser nulos
        AnyadeErrores "Valores nulos ", LINEA
        Exit Function
    End If
    
    'Cuenta NO puede ser numerica
    If Not IsNumeric(Valores(0)) Then
        AnyadeErrores "Cuenta: " & Valores(0), "No Numerica"
        Exit Function
    End If
    
    
    
    For I = 7 To 10
        If Valores(I) <> "" Then
            If Not IsNumeric(Valores(I)) Then
                AnyadeErrores "Cuenta bancaria", "CCC(" & I & "):   " & Valores(I)
                Exit Function
            End If
        End If
    Next I
    
    'Vemos si existe
    'Vemos si existe
    Crear = False
    If Not ExisteCuenta(Valores(0)) Then
        If Me.chkCrear.Value = 0 Then
            AnyadeErrores "Cuenta: " & Valores(0), "No existe"
            Exit Function
        Else
            Crear = True
        End If
    End If
    
    'Controlamos valores de Multibase para los textos, y las ' para la insercion
    For I = 1 To 5 'Sin NIF ni codmacta, 6 y 0 respectivamente
        If I <> 3 Then
            Cad = RevisaCaracterMultibase(Valores(I))
            NombreSQL Cad
            Valores(I) = Cad
        End If
    Next I
    
    
    
    
    '
    
    
     If Crear Then
        I = DigitosNivel(vEmpresa.numNivel - 1)
        Cad = Mid(Valores(0), 1, I)
        If Cad <> CadenaDesdeOtroForm Then
            If Not CreaSubcuentas(Valores(0), I, "IMPORTACION AUTOMATICA") Then
                AnyadeErrores "Cuenta: " & Valores(0), "GENERANDO SUBNIVELES"
                Exit Function
            End If
            CadenaDesdeOtroForm = Cad
        End If
    End If
    
'        ALTER TABLE `cuentas` ADD `entidad` VARCHAR(4) ;
'        ALTER TABLE `cuentas` ADD `oficina` VARCHAR(4) ;
'        ALTER TABLE `cuentas` ADD `CC` VARCHAR(4);
'        ALTER TABLE `cuentas` ADD `cuentaba` VARCHAR(10);
'
    'Montamos el SQL
        'Montamos el SQL
    If Crear Then
        Cad = "INSERT INTO Cuentas (codmacta,nommacta,razosoci,dirdatos,codposta,despobla,desprovi,nifdatos,"
        'NUEVO
        Cad = Cad & "entidad,oficina,CC,cuentaba,"
        Cad = Cad & "model347,apudirec) VALUES ("
        Cad = Cad & "'" & Valores(0) & "',"
        Cad = Cad & "'" & Valores(1) & "',"
        Cad = Cad & "'" & Valores(1) & "',"
        Cad = Cad & "'" & Valores(2) & "',"
        Cad = Cad & "'" & Valores(3) & "',"
        Cad = Cad & "'" & Valores(4) & "',"
        Cad = Cad & "'" & Valores(5) & "',"
        Cad = Cad & "'" & Valores(6) & "',"
        For I = 7 To 10
            If Valores(I) = "" Then
                Cad = Cad & "NULL,"
            Else
                Cad = Cad & "'" & Valores(I) & "',"
            End If
        Next I
        If Valores(11) = "1" Then
            Cad = Cad & "1"
        Else
            Cad = Cad & "0"
        End If
        
        Cad = Cad & ",'S')"
    
    Else
        Cad = "UPDATE Cuentas SET "
        Cad = Cad & " nommacta = '" & Valores(1) & "',"
        Cad = Cad & " razosoci = '" & Valores(1) & "',"
        Cad = Cad & " dirdatos = '" & Valores(2) & "',"
        Cad = Cad & " codposta = '" & Valores(3) & "',"
        Cad = Cad & " despobla = '" & Valores(4) & "',"
        Cad = Cad & " desprovi = '" & Valores(5) & "',"
        Cad = Cad & " nifdatos = '" & Valores(6) & "',"
        'model347
        Cad = Cad & " model347 = "
        If Valores(11) = "1" Then
            Cad = Cad & "1"
        Else
            Cad = Cad & "0"
        End If
        
        'CCC
        Cad = Cad & ", entidad =" & ValorSQL(Valores(7))
        Cad = Cad & ", oficina =" & ValorSQL(Valores(8))
        Cad = Cad & ", CC =" & ValorSQL(Valores(9))
        Cad = Cad & ", cuentaba =" & ValorSQL(Valores(10))
            
        
        Cad = Cad & " WHERE codmacta ='" & Valores(0) & "'"
    End If
   
    If Not EjecutaSQL2(Cad) Then Exit Function
    ProcesarLinea = True
    Exit Function
EProcesarLinea:
    AnyadeErrores "Linea: " & LINEA, Err.Description
    Err.Clear
    
End Function

Private Function ValorSQL(ByRef C As String) As String
    If C = "" Then
        ValorSQL = "NULL"
    Else
        ValorSQL = "'" & C & "'"
    End If
End Function
Private Function EjecutaSQL2(SQL As String) As Boolean
    EjecutaSQL2 = False
    On Error Resume Next
    conn.Execute SQL
    If Err.Number <> 0 Then
        AnyadeErrores "SQL: " & SQL, Err.Description
        Err.Clear
    Else
        EjecutaSQL2 = True
    End If
End Function


Private Sub AnyadeErrores(L1 As String, L2 As String)
    NE = NE + 1
    Errores = Errores & "-----------------------------" & vbCrLf
    Errores = Errores & L1 & vbCrLf
    Errores = Errores & L2 & vbCrLf


End Sub




Private Sub ImprimeFichero()
Dim NF As Integer
    On Error GoTo EImprimeFichero
    NF = FreeFile
    Open App.path & "\errimpdat.txt" For Output As #NF
    Print #NF, Errores
    Close (NF)
    Shell "notepad.exe " & App.path & "\errimpdat.txt", vbMaximizedFocus
    Exit Sub
EImprimeFichero:
    MsgBox Err.Description & vbCrLf, vbCritical
    Err.Clear
End Sub


Private Function ExisteCuenta(Cta As String) As Boolean

    
    ExisteCuenta = False
    SQL = DevuelveDesdeBD("codmacta", "cuentas", "codmacta", Cta, "T")
    If SQL <> "" Then ExisteCuenta = True
    
End Function










Private Sub Text7_GotFocus(Index As Integer)
    Text7(Index).SelStart = 0
    Text7(Index).SelLength = Len(Text7(Index).Text)
End Sub

Private Sub Text7_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        SendKeys "{tab}"
    End If
End Sub





Private Sub txtFecha_GotFocus(Index As Integer)
    txtFecha(Index).SelStart = 0
    txtFecha(Index).SelLength = Len(txtFecha(Index).Text)
End Sub

Private Sub txtfecha_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        SendKeys "{tab}"
    End If
End Sub

Private Sub txtfecha_LostFocus(Index As Integer)
    txtFecha(Index).Text = Trim(txtFecha(Index))
    If txtFecha(Index) = "" Then Exit Sub
    If Not EsFechaOK(txtFecha(Index)) Then
        MsgBox "Fecha incorrecta: " & txtFecha(Index), vbExclamation
        txtFecha(Index).Text = ""
        txtFecha(Index).SetFocus
    End If
End Sub




Private Sub HacerCambios(ByVal Tabla As Integer)
Dim Cambio As String
Dim Inicio As Integer
Dim Fin As Integer
Dim Cad As String

    'RevisaCaracterMultibase
    Select Case Tabla
    Case 0
        'Cuentas
        SQL = "Select codmacta,nommacta, razosoci, dirdatos,  despobla, desprovi,pais"
        SQL = SQL & " FROM CUentas"
        Inicio = 1 'k es dos
        Fin = 6
    Case 1, 2
        'HCO apuntes
        SQL = "Select fechaent,numasien,numdiari,numdocum,ampconce,linliapu from hlinapu"
        If Tabla = 2 Then SQL = SQL & "1"
        Cad = ""
        If txtFecha(0).Text <> "" Then Cad = "fechaent >='" & Format(txtFecha(0).Text, FormatoFecha) & "'"
        If txtFecha(1).Text <> "" Then
            If Cad <> "" Then Cad = Cad & " AND "
            Cad = Cad & "fechaent <='" & Format(txtFecha(1).Text, FormatoFecha) & "'"
        End If
        If Cad <> "" Then SQL = SQL & " WHERE " & Cad
        Inicio = 4
        Fin = 4
    Case 3
        'Facturas clientes
        SQL = "Select anofaccl,codfaccl,numserie,confaccl FROM cabfact "
        Cad = ""
        If txtFecha(0).Text <> "" Then Cad = "fecfaccl >='" & Format(txtFecha(0).Text, FormatoFecha) & "'"
        If txtFecha(1).Text <> "" Then
            If Cad <> "" Then Cad = Cad & " AND "
            Cad = Cad & "fecfaccl <='" & Format(txtFecha(1).Text, FormatoFecha) & "'"
        End If
        If Cad <> "" Then SQL = SQL & " WHERE " & Cad
        Inicio = 3
        Fin = 3
    Case 4
        'Facturas proveedores
        SQL = "Select anofacpr,numregis,confacpr FROM cabfactprov "
        Cad = ""
        If txtFecha(0).Text <> "" Then Cad = "fecrecpr >='" & Format(txtFecha(0).Text, FormatoFecha) & "'"
        If txtFecha(1).Text <> "" Then
            If Cad <> "" Then Cad = Cad & " AND "
            Cad = Cad & "fecrecpr <='" & Format(txtFecha(1).Text, FormatoFecha) & "'"
        End If
        If Cad <> "" Then SQL = SQL & " WHERE " & Cad
        Inicio = 2
        Fin = 2
        
    Case 5
        SQL = "Select * from conceptos"
        Inicio = 1
        Fin = 1
    End Select
    
    Set RS = New ADODB.Recordset
    RS.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not RS.EOF Then
        
        While Not RS.EOF
            Label34.Caption = RS.Fields(0) & " - " & RS.Fields(1)
            Label34.Refresh
            Cambio = ""
            For I = Inicio To Fin
                'Campo no nulo
                If Not IsNull(RS.Fields(I)) Then
                    SQL = RS.Fields(I)
                    Cad = RevisaCaracterMultibase(SQL)
                    If SQL <> Cad Then
                        'Han habido cambios
                        If Cambio <> "" Then Cambio = Cambio & ","
                        SQL = DevNombreSQL(Cad)
                        NumRegElim = NumRegElim + 1
                        Cambio = Cambio & RS.Fields(I).Name & " = '" & SQL & "'"
                    End If
                End If
            Next I
            If Cambio <> "" Then
                'OK HAY K CAMBIAR, k updatear
                Select Case Tabla
                Case 0
                    SQL = "UPDATE Cuentas SET " & Cambio & " WHERE codmacta ='" & RS.Fields(0) & "'"
            
                Case 1, 2
                    SQL = "UPDATE Hlinapu"
                    If Tabla = 2 Then SQL = SQL & "1"
                    SQL = SQL & " SET " & Cambio & " WHERE numdiari =" & RS!NumDiari
                    SQL = SQL & " AND numasien =" & RS!Numasien
                    SQL = SQL & " AND fechaent = '" & Format(RS!fechaent, FormatoFecha) & "'"
                    SQL = SQL & " AND linliapu =" & RS!Linliapu
                Case 3
                    SQL = "UPDATE cabfact SET " & Cambio & " WHERE numserie ='" & RS!numserie
                    SQL = SQL & "' AND anofaccl =" & RS!anofaccl & " AND codfaccl=" & RS!codfaccl
                Case 4
                    SQL = "UPDATE cabfactprov SET " & Cambio & " WHERE "
                    SQL = SQL & " anofacpr =" & RS!anofacpr & " AND numregis=" & RS!numRegis
                Case 5
                    SQL = "UPDATE conceptos SET " & Cambio & " WHERE codconce=" & RS!codConce
                End Select
                
                'Ejecutamos
                conn.Execute SQL
            End If
            RS.MoveNext
        Wend
    End If
    RS.Close
    Set RS = Nothing
            
End Sub





'-----------------------------------------------------------------------------------
'
'
Private Function ExportarDatosFacturas(Proveedores As Boolean) As Boolean
Dim vOpc As String
Dim Aux As String

    'Comprobamos el RS
    ExportarDatosFacturas = False
    If Proveedores Then
        SQL = "prov"
        Parametros = "fecrecpr"
    Else
        Parametros = "fecfaccl"
        SQL = ""
    End If
    SQL = SQL & " where " & Parametros & " >= '" & Format(CDate(txtFecha(2).Text), FormatoFecha) & "'"
    SQL = SQL & " AND " & Parametros & " <= '" & Format(CDate(txtFecha(3).Text), FormatoFecha) & "'"
    Set RS = New ADODB.Recordset
    Errores = "select count(*) from cabfact" & SQL
    RS.Open Errores, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    OK = 0
    If Not RS.EOF Then
        If Not IsNull(RS.Fields(0)) Then
            If RS.Fields(0) > 0 Then OK = 1
        End If
    End If
    RS.Close
    
    If OK = 0 Then
        SQL = "Ningun dato a traspasar de facturas "
        If Proveedores Then
            SQL = SQL & "proveedores"
        Else
            SQL = SQL & "clientes"
        End If
        MsgBox SQL, vbExclamation
        Exit Function
    End If
    

    

    
    '----------------------------------------------------------------------
    'OPCION
    vOpc = "OPCION"
    EncabezadoPieFact False, vOpc, 0
    If Proveedores Then
        Print #NE, 0
        Print #NE, "Proveedores"
    Else
        Print #NE, 1
        Print #NE, "Clientes"
    End If
    
    'Ultimo nivel de las cuentas contables
    Print #NE, vEmpresa.DigitosUltimoNivel
    EncabezadoPieFact True, vOpc, 1
    
    '----------------------------------------------------------------------
    'CUENTAS
    vOpc = "CUENTAS"
    Label40.Caption = "Cuentas"
    Label40.Refresh
    
    EncabezadoPieFact False, vOpc, 0
    Parametros = "DELETE FROM tmpcierre1 where codusu =" & vUsu.Codigo
    conn.Execute Parametros
    
    
    'Cuentas que necesito
    
    Parametros = "Select distinct(codmacta) from cabfact" & SQL
    RS.Open Parametros, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    While Not RS.EOF
        InsertaEnTmpCta
        RS.MoveNext
        
    Wend
    RS.Close
    
    Parametros = "Select distinct(Cuereten) from cabfact" & SQL & " and not (cuereten is null)"
    RS.Open Parametros, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    While Not RS.EOF
        InsertaEnTmpCta
        RS.MoveNext
    Wend
    RS.Close
    
    
    'la cuentas de las lineas de factura
    Parametros = "Select codtbase from linfact"
    If Proveedores Then
        Parametros = Parametros & "prov where anofacpr >=" & Year(CDate(txtFecha(2).Text))
        Parametros = Parametros & " and anofacpr <=" & Year(CDate(txtFecha(3).Text))
    Else
        Parametros = Parametros & " where anofaccl >=" & Year(CDate(txtFecha(2).Text))
        Parametros = Parametros & " and anofaccl <=" & Year(CDate(txtFecha(3).Text))
    End If
    Parametros = Parametros & " GROUP BY codtbase"
    RS.Open Parametros, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    While Not RS.EOF
        InsertaEnTmpCta
        RS.MoveNext
    Wend
    RS.Close
    
    
    'Ahora cojo todos los datos de tmpcierr1 y creo los inserts de las cuentas
    Parametros = "Select cuentas.* from cuentas,tmpcierre1 where cuentas.codmacta=tmpcierre1.cta "
    Parametros = Parametros & " and codusu =" & vUsu.Codigo
    RS.Open Parametros, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    BACKUP_TablaIzquierda RS, CadenaDesdeOtroForm
    OK = 0
    While Not RS.EOF
        Label40.Caption = RS!Codmacta
        Label40.Refresh
        OK = OK + 1
        BACKUP_Tabla RS, Parametros
        Parametros = "INSERT INTO Cuentas " & CadenaDesdeOtroForm & " VALUES " & Parametros & ";"
        Print #NE, Parametros
        RS.MoveNext
    Wend
    RS.Close
    
    EncabezadoPieFact True, vOpc, OK


    '----------------------------------------------------------------------
    'OPCION
    vOpc = "CC"
    Label40.Caption = "C.C."
    Label40.Refresh
    EncabezadoPieFact False, vOpc, 0
    'Volvemos a utlizar la misma tabla
    Parametros = "DELETE FROM tmpcierre1 where codusu =" & vUsu.Codigo
    conn.Execute Parametros

    Parametros = "Select codccost from linfact"
    If Proveedores Then
        Parametros = Parametros & "prov where anofacpr >=" & Year(CDate(txtFecha(2).Text))
        Parametros = Parametros & " and anofacpr <=" & Year(CDate(txtFecha(3).Text))
    Else
        Parametros = Parametros & " where anofaccl >=" & Year(CDate(txtFecha(2).Text))
        Parametros = Parametros & " and anofaccl <=" & Year(CDate(txtFecha(3).Text))
    End If
    Parametros = Parametros & " AND not (codccost is null)"
    Parametros = Parametros & " GROUP BY codccost"
    RS.Open Parametros, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    OK = 0
    While Not RS.EOF
        OK = OK + 1
        InsertaEnTmpCta
        RS.MoveNext
        
    Wend
    RS.Close
    
    
    
    
    If OK > 0 Then
        'Ahora cojo todos los datos de tmpcierr1 y creo los inserts de las cuentas
        Parametros = "Select cabccost.* from cabccost,tmpcierre1 where cabccost.codccost=tmpcierre1.cta "
        Parametros = Parametros & " and codusu =" & vUsu.Codigo
        RS.Open Parametros, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        
        BACKUP_TablaIzquierda RS, CadenaDesdeOtroForm

        'Si k hay CC
        
        While Not RS.EOF
            OK = OK + 1
            BACKUP_Tabla RS, Parametros
            Parametros = "INSERT INTO cabccost " & CadenaDesdeOtroForm & " VALUES " & Parametros & ";"
            Print #NE, Parametros
            RS.MoveNext
        Wend
        RS.Close
            
    End If
    
    
    EncabezadoPieFact True, vOpc, 0
    
    
    
    
    '----------------------------------------------------------------------
    'OPCION
    vOpc = "IVA"
    EncabezadoPieFact False, vOpc, 0
    Parametros = "Select * from tiposiva"
    RS.Open Parametros, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    BACKUP_TablaIzquierda RS, CadenaDesdeOtroForm


    While Not RS.EOF
        OK = OK + 1
        BACKUP_Tabla RS, Parametros
        Parametros = "INSERT INTO tiposiva " & CadenaDesdeOtroForm & " VALUES " & Parametros & ";"
        Print #NE, Parametros
        RS.MoveNext
    Wend
    RS.Close

    
    EncabezadoPieFact True, vOpc, 1
    
    
    '------------------------------------
    'Para las facturas de clientes necesitare tb las series de factura
    If Not Proveedores Then
        vOpc = "CONTADORES"
        EncabezadoPieFact False, vOpc, 0
        Parametros = "Select * from contadores"
        RS.Open Parametros, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
        BACKUP_TablaIzquierda RS, CadenaDesdeOtroForm


        While Not RS.EOF
            OK = OK + 1
            BACKUP_Tabla RS, Parametros
            Parametros = "INSERT INTO contadores " & CadenaDesdeOtroForm & " VALUES " & Parametros & ";"
            Print #NE, Parametros
            RS.MoveNext
        Wend
        RS.Close
    
        
        EncabezadoPieFact True, vOpc, 1
    End If
    
    
    
    '----------------------------------------------------------------------
    'FACTURAS
    'Grabaremos en cada linea
    '
    '  codigo |INSERT |UPDATE |base1|base2....
    '   Codigo: Para clientes será: numserie, codfacl, anofaccl
    
    vOpc = "FACTURAS"
    EncabezadoPieFact False, vOpc, 0
    
    
    If Not Proveedores Then
        SQL = "numserie,codfaccl,anofaccl,fecfaccl,codmacta,confaccl,ba1faccl,ba2faccl,ba3faccl,pi1faccl,pi2faccl,pi3faccl,pr1faccl,pr2faccl,pr3faccl,ti1faccl,ti2faccl,ti3faccl,tr1faccl,tr2faccl,tr3faccl,totfaccl,tp1faccl,tp2faccl,tp3faccl,intracom,retfaccl,trefaccl,cuereten,fecliqcl"
        Parametros = "Select " & SQL & " from cabfact"
        Parametros = Parametros & " where fecfaccl >= '" & Format(CDate(txtFecha(2).Text), FormatoFecha) & "'"
        Parametros = Parametros & " and fecfaccl <= '" & Format(CDate(txtFecha(3).Text), FormatoFecha) & "'"
        OK = 3
    Else
        SQL = "numregis,anofacpr,fecfacpr,fecrecpr,numfacpr,codmacta,confacpr,ba1facpr,ba2facpr,ba3facpr,pi1facpr,pi2facpr,pi3facpr,pr1facpr,pr2facpr,pr3facpr,ti1facpr,ti2facpr,ti3facpr,tr1facpr,tr2facpr,tr3facpr,totfacpr,tp1facpr,tp2facpr,tp3facpr,extranje,retfacpr,trefacpr,cuereten,fecliqpr,nodeducible"
        Parametros = "Select " & SQL & " from cabfactprov"
        Parametros = Parametros & " where fecrecpr >= '" & Format(CDate(txtFecha(2).Text), FormatoFecha) & "'"
        Parametros = Parametros & " and fecrecpr <= '" & Format(CDate(txtFecha(3).Text), FormatoFecha) & "'"
        OK = 2
    End If
    
    RS.Open Parametros, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    BACKUP_TablaIzquierda RS, CadenaDesdeOtroForm
    If Not Proveedores Then
        CadenaDesdeOtroForm = "INSERT INTO cabfact " & CadenaDesdeOtroForm & " VALUES "
    Else
        CadenaDesdeOtroForm = "INSERT INTO cabfactprov " & CadenaDesdeOtroForm & " VALUES "
    End If
        
        
    Set miRsAux = New ADODB.Recordset
    Errores = ""
    NumRegElim = 0
    While Not RS.EOF
        
        NumRegElim = NumRegElim + 1
        SQL = RS.Fields(0) & "|" & RS.Fields(1) & "|"
        If Not Proveedores Then SQL = SQL & RS.Fields(2) & "|"
        Label40.Caption = SQL
        Label40.Refresh
        'Cadena insert
        BACKUP_Tabla RS, Parametros
        Parametros = SQL & CadenaDesdeOtroForm & Parametros & ";|"
        
        
        'El UPDATE
        SQL = ""
        
        For I = OK To RS.Fields.Count - 1
            If SQL <> "" Then SQL = SQL & ","
            SQL = SQL & RS.Fields(I).Name & " = "
            If IsNull(RS.Fields(I)) Then
                SQL = SQL & "NULL"
            Else
                Select Case RS.Fields(I).Type
                Case 133
                    SQL = SQL & "'" & Format(RS.Fields(I), FormatoFecha) & "'"
                
                Case 17
                    'numero
                    SQL = SQL & RS.Fields(I)
                    
                Case 131
                    SQL = SQL & TransformaComasPuntos(CStr(RS.Fields(I)))
                Case Else
                    SQL = SQL & "'" & DevNombreSQL(RS.Fields(I)) & "'"
                End Select
                
            End If
        Next I
      
        SQL = SQL & " WHERE "
        Aux = ""
        For I = 0 To OK - 1
            Aux = Aux & RS.Fields(I).Name & " = '" & RS.Fields(I) & "' and "
        Next
        Aux = Mid(Aux, 1, Len(Aux) - 4)
        SQL = SQL & Aux
        If Not Proveedores Then
            SQL = "UPDATE cabfact SET " & SQL
        Else
            SQL = "UPDATE cabfactprov SET " & SQL
        End If
        Parametros = Parametros & SQL & "|"
        
        
        'Metemos una marca para separar las lineas
        Parametros = Parametros & "<>"
        
        'Las lineas
        '----------------------------
        
        SQL = "Select * from linfact"
        If Proveedores Then SQL = SQL & "prov"
        SQL = SQL & " WHERE " & Aux
        miRsAux.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        If Errores = "" Then
            BACKUP_TablaIzquierda miRsAux, Errores
            SQL = "INSERT INTO linfact"
            If Proveedores Then SQL = SQL & "prov"
            Errores = SQL & "  " & Errores & " VALUES "
        End If
        While Not miRsAux.EOF
            BACKUP_Tabla miRsAux, SQL
            SQL = Errores & SQL & ";"
            Parametros = Parametros & SQL & "|"
            miRsAux.MoveNext
        Wend
        miRsAux.Close
        Print #NE, Parametros
        
        RS.MoveNext
    Wend
    RS.Close
    Set miRsAux = Nothing
    
    EncabezadoPieFact True, vOpc, CInt(NumRegElim)
    
    
    
    
    'Y dejo limpio el tajo
    Parametros = "DELETE FROM tmpcierre1 where codusu =" & vUsu.Codigo
    conn.Execute Parametros
    Label40.Caption = ""
    CadenaDesdeOtroForm = "2"
    Set RS = Nothing
    
    ExportarDatosFacturas = True
End Function


Private Sub CopiarArchivo()
On Error GoTo ECopiarArchivo

    If Dir(Text8.Text, vbArchive) <> "" Then Kill Text8.Text
    FileCopy Errores, Text8.Text
    
    Errores = "El fichero: " & Text8.Text & " se ha generado con éxito"
    MsgBox Errores, vbInformation
    Exit Sub
ECopiarArchivo:
    MuestraError Err.Number, "Copiar archivo"
End Sub



Private Sub EncabezadoPieFact(Pie As Boolean, ByVal Text As String, REG As Integer)
    If Pie Then
        Text = "[/" & Text & "]" & REG
    Else
        Text = "[" & Text & "]"
    End If
    Print #NE, Text
End Sub


Private Sub InsertaEnTmpCta()
On Error Resume Next
    
    conn.Execute "INSERT INTO tmpcierre1 (codusu, cta) VALUES (" & vUsu.Codigo & ",'" & RS.Fields(0) & "')"
    If Err.Number <> 0 Then Err.Clear
End Sub

Private Sub ImportarFicheroFac()
    NE = FreeFile
    Screen.MousePointer = vbHourglass
    Open Text8.Text For Input As #NE
    'Importamos el primer trozo. PROVEEDORES
    If ImportarDatosFacturas Then
        'CLIENTES
        If ImportarDatosFacturas Then
            Close #NE
            MsgBox "Proceso finalizado", vbExclamation
            If chkImportarFacturas.Value Then
                If Dir(Text8.Text, vbArchive) <> "" Then Kill Text8.Text
            End If
        End If
    End If
    Label40.Caption = ""
    Screen.MousePointer = vbDefault
End Sub

'---------------------------------------------------------------------------------------
Private Function ImportarDatosFacturas() As Boolean
Dim Fin As Boolean
Dim Clientes As Boolean

    On Error GoTo EIM
    
    CadenaDesdeOtroForm = "Abriendo fichero. Datos basicos"
    Label40.Caption = CadenaDesdeOtroForm
    Label40.Refresh
    
    
    Line Input #NE, SQL   'OPCION
    If SQL <> "[OPCION]" Then
        MsgBox "Formato fichero incorrecto", vbExclamation
        Close NE
        Exit Function
    End If
    Line Input #NE, SQL   ' PRoveedores o clientes
    I = Val(SQL)
    If I = 1 Then
        'CLIENTES
        Clientes = True
    Else
        'PROVEEDORES
        Clientes = False
    End If
    Line Input #NE, SQL   ' Datos vacios
    Line Input #NE, SQL   ' digitos ultimo nivel
    I = Val(SQL)
    If I <> vEmpresa.DigitosUltimoNivel Then
        MsgBox "Ultimo nivel disitinto:" & I, vbExclamation
        Close NE
        Exit Function
    End If
    Line Input #NE, SQL   'FIN OPCION
    
    'CUENTAS
    CadenaDesdeOtroForm = "Cuentas"
    Label40.Caption = CadenaDesdeOtroForm
    Label40.Refresh
    Line Input #NE, SQL   'CUENTAS
    I = 0
    Fin = False
    Do
        Line Input #NE, SQL   'FIN OPCION
        If InStr(1, SQL, "[/CUENTAS]") > 0 Then
            'Fin
            Fin = True
            'Ver numero registros
            OK = InStr(1, SQL, "]")
            SQL = Mid(SQL, OK + 1)
            OK = Val(SQL)
            If I <> OK Then
            
            End If
        Else
            'Mandamos la linea a ejecutar
            Label40.Caption = "Cta: " & Mid(SQL, 155 + vEmpresa.DigitosUltimoNivel, 30)
            Label40.Refresh
            EjecutarSQL
            I = I + 1
            
        End If
    Loop Until Fin
    
    'CC
    CadenaDesdeOtroForm = "CC"
    Label40.Caption = CadenaDesdeOtroForm
    Label40.Refresh
    Line Input #NE, SQL   'CC
    I = 0
    Fin = False
    Do
        Line Input #NE, SQL   'FIN OPCION
        If InStr(1, SQL, "[/CC]") > 0 Then
            'Fin
            Fin = True
            'Ver numero registros
        Else
            'Mandamos la linea a ejecutar
            EjecutarSQL
            I = I + 1
            
        End If
    Loop Until Fin
    
    
    
    'IVA
    CadenaDesdeOtroForm = "IVA"
    Label40.Caption = CadenaDesdeOtroForm
    Label40.Refresh
    Line Input #NE, SQL   'IVA
    I = 0
    Fin = False
    Do
        Line Input #NE, SQL   'FIN OPCION
        If InStr(1, SQL, "[/IVA]") > 0 Then
            'Fin
            Fin = True
            'Ver numero registros
        Else
            'Mandamos la linea a ejecutar
            EjecutarSQL
            I = I + 1
            
        End If
    Loop Until Fin
    
    
    If Clientes Then
        'SOLO CLIENTES LLEVA CONTADORES
        CadenaDesdeOtroForm = "CONTADORES"
        Label40.Caption = CadenaDesdeOtroForm
        Label40.Refresh
        Line Input #NE, SQL   'CONTADPORES
        I = 0
        Fin = False
        Do
            Line Input #NE, SQL   'FIN OPCION
            If InStr(1, SQL, "[/CONTADORES]") > 0 Then
                'Fin
                Fin = True
                'Ver numero registros
            Else
                'Mandamos la linea a ejecutar
                EjecutarSQL
                I = I + 1
                
            End If
        Loop Until Fin
    End If
    
    
    
    
    'FACTURAS
    Set RS = New ADODB.Recordset
    CadenaDesdeOtroForm = "FACTURAS"
    Label40.Caption = CadenaDesdeOtroForm
    Label40.Refresh
    Line Input #NE, SQL   'FACTS
    I = 0
    Fin = False
    Do
        Line Input #NE, SQL   'FIN OPCION
        If InStr(1, SQL, "[/FACT") > 0 Then
            'Fin
            Fin = True
            'Ver numero registros
            OK = InStr(1, SQL, "]")
            OK = Val(Mid(SQL, OK + 1))
            If OK <> I Then MsgBox "Diferencia entre facturas procesadas. Fichero: " & OK & " -> " & I, vbExclamation
        Else
            'Mandamos la linea a ejecutar
            ProcesarLineaFactura Clientes
            I = I + 1
        End If
    Loop Until Fin
    Label40.Caption = "Actualizando datos"
    Me.Refresh
    espera 1
    ImportarDatosFacturas = True
    Exit Function
    
EIM:
    
    If Err.Number <> 0 Then
        MuestraError Err.Number, CadenaDesdeOtroForm
        I = 1
    Else
        I = 0
    End If
    Close NE
    If I = 0 Then
        If Me.chkImportarFacturas.Value = 1 Then Kill Text8.Text
    End If
    Set RS = Nothing
End Function


Private Sub EjecutarSQL()
    On Error Resume Next
    
    conn.Execute SQL
    If Err.Number <> 0 Then
        If conn.Errors(0).Number = 1062 Then
            Err.Clear
        Else
            'MuestraError Err.Number, Err.Description
        End If
        Err.Clear
    End If
End Sub


Private Sub ProcesarLineaFactura(Clientes As Boolean)
Dim Año As Integer
Dim numero As Long
Dim Serie As String
Dim J As Long
Dim Aux As String

    If Clientes Then
        Serie = RecuperaValor(SQL, 1)
        numero = RecuperaValor(SQL, 2)
        Año = RecuperaValor(SQL, 3)
    Else
        Serie = ""
        numero = RecuperaValor(SQL, 1)
        Año = RecuperaValor(SQL, 2)
    End If
    
        
    Label40.Caption = Serie & " " & numero & " / " & Año
    Label40.Refresh
    DoEvents
    'Quitamos el La cadaena
    J = InStr(4, SQL, "|" & Año & "|")
    If J = 0 Then
        MsgBox "Error en año factura", vbExclamation
        Exit Sub
    End If
    
    J = J + 6
    SQL = Mid(SQL, J)
    
    If Clientes Then
        Aux = "Select * from cabfact WHERE numserie = '" & Serie & "'"
        Aux = Aux & " and anofaccl = " & Año & " and codfaccl =" & numero
    Else
        Aux = "Select * from cabfactprov WHERE "
        Aux = Aux & " anofacpr = " & Año & " and numregis =" & numero
    End If
    RS.Open Aux, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    J = 1
    If Not RS.EOF Then
        J = 2
        'Borro las lineas
        If Clientes Then
            Aux = "DELETE from linfact WHERE numserie = '" & Serie & "'"
            Aux = Aux & " and anofaccl = " & Año & " and codfaccl =" & numero
        Else
            Aux = "DELETE from linfactprov WHERE "
            Aux = Aux & " anofacpr = " & Año & " and numregis =" & numero
        End If
        conn.Execute Aux
    End If
    RS.Close
    
    Aux = RecuperaValor(SQL, CInt(J))
    conn.Execute Aux
    
    '---------------------------------------------------
    J = InStr(1, SQL, "<>")
    If J = 0 Then
        MsgBox "Error lineas: " & Aux, vbExclamation
        Exit Sub
    End If
        
    SQL = Mid(SQL, J + 2)
    Do
        J = InStr(1, SQL, "|")
        If J > 0 Then
            Aux = Mid(SQL, 1, J - 1)
            SQL = Mid(SQL, J + 1)
            conn.Execute Aux
        End If
    Loop Until J = 0
End Sub
