VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmCargaFactVar 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Carga Masiva de Facturas Varias"
   ClientHeight    =   7485
   ClientLeft      =   45
   ClientTop       =   2430
   ClientWidth     =   6645
   Icon            =   "frmCargaFactVar.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7485
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
      Height          =   7365
      Left            =   90
      TabIndex        =   13
      Top             =   60
      Width           =   6375
      Begin VB.TextBox txtcodigo 
         Height          =   285
         Index           =   10
         Left            =   1560
         MaxLength       =   50
         TabIndex        =   7
         Tag             =   "Código Propio|N|N|1|99|sbanco|codbanpr|000|S|"
         Top             =   4560
         Width           =   3990
      End
      Begin VB.TextBox txtcodigo 
         Height          =   795
         Index           =   9
         Left            =   360
         MaxLength       =   250
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   5
         Tag             =   "Observaciones|T|S|||cabfact|observac|||"
         Top             =   3150
         Width           =   5235
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   3
         Left            =   3030
         MaxLength       =   10
         TabIndex        =   10
         Tag             =   "Código Propio|N|N|1|99|sbanco|codbanpr|00|S|"
         Top             =   5250
         Width           =   1320
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   2
         Left            =   1620
         MaxLength       =   10
         TabIndex        =   9
         Tag             =   "Código Propio|N|N|1|99|sbanco|codbanpr|00|S|"
         Top             =   5250
         Width           =   1050
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   1
         Left            =   360
         MaxLength       =   10
         TabIndex        =   8
         Tag             =   "Código Propio|N|N|1|99|sbanco|codbanpr|00|S|"
         Top             =   5250
         Width           =   1050
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   0
         Left            =   2655
         Locked          =   -1  'True
         TabIndex        =   27
         Text            =   "Text5"
         Top             =   4200
         Width           =   2910
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   0
         Left            =   1560
         MaxLength       =   3
         TabIndex        =   6
         Tag             =   "Código Propio|N|N|1|99|sbanco|codbanpr|000|S|"
         Top             =   4200
         Width           =   1050
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   5
         Left            =   2730
         Locked          =   -1  'True
         TabIndex        =   26
         Text            =   "Text5"
         Top             =   2460
         Width           =   2910
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   4
         Left            =   2730
         Locked          =   -1  'True
         TabIndex        =   25
         Text            =   "Text5"
         Top             =   2100
         Width           =   2910
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   6
         Left            =   1620
         MaxLength       =   6
         TabIndex        =   0
         Top             =   495
         Width           =   830
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   6
         Left            =   2490
         Locked          =   -1  'True
         TabIndex        =   23
         Text            =   "Text5"
         Top             =   495
         Width           =   3135
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   8
         Left            =   1590
         MaxLength       =   10
         TabIndex        =   2
         Tag             =   "Código Propio|N|N|1|99|sbanco|codbanpr|00|S|"
         Top             =   1440
         Width           =   1050
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   8
         Left            =   2715
         Locked          =   -1  'True
         TabIndex        =   18
         Text            =   "Text5"
         Top             =   1440
         Width           =   2910
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   7
         Left            =   1605
         MaxLength       =   10
         TabIndex        =   1
         Tag             =   "Código Postal|T|S|||clientes|codposta|||"
         Top             =   960
         Width           =   1050
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   5
         Left            =   1590
         MaxLength       =   10
         TabIndex        =   4
         Top             =   2445
         Width           =   1035
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   4
         Left            =   1605
         MaxLength       =   10
         TabIndex        =   3
         Top             =   2100
         Width           =   1035
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Left            =   4755
         TabIndex        =   12
         Top             =   6720
         Width           =   975
      End
      Begin VB.CommandButton cmdAceptar 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   3570
         TabIndex        =   11
         Top             =   6720
         Width           =   975
      End
      Begin MSComctlLib.ProgressBar Pb1 
         Height          =   285
         Left            =   300
         TabIndex        =   17
         Top             =   6420
         Width           =   5370
         _ExtentX        =   9472
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Ampliación"
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
         Index           =   10
         Left            =   360
         TabIndex        =   33
         Top             =   4560
         Width           =   750
      End
      Begin VB.Label Label29 
         Caption         =   "Observaciones"
         ForeColor       =   &H00972E0B&
         Height          =   255
         Left            =   360
         TabIndex        =   32
         Top             =   2880
         Width           =   1125
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Importe"
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
         Index           =   9
         Left            =   3060
         TabIndex        =   31
         Top             =   5010
         Width           =   570
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Precio"
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
         Index           =   8
         Left            =   1650
         TabIndex        =   30
         Top             =   5010
         Width           =   435
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   2
         Left            =   1290
         MouseIcon       =   "frmCargaFactVar.frx":000C
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar cuenta"
         Top             =   2460
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   1
         Left            =   1290
         MouseIcon       =   "frmCargaFactVar.frx":015E
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar cuenta"
         Top             =   2100
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Cantidad"
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
         TabIndex        =   29
         Top             =   5010
         Width           =   645
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   0
         Left            =   1230
         MouseIcon       =   "frmCargaFactVar.frx":02B0
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar concepto"
         Top             =   4200
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Concepto "
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
         TabIndex        =   28
         Top             =   4200
         Width           =   735
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
         Left            =   375
         TabIndex        =   24
         Top             =   480
         Width           =   540
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   6
         Left            =   1290
         MouseIcon       =   "frmCargaFactVar.frx":0402
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar sección"
         Top             =   495
         Width           =   240
      End
      Begin VB.Label lblProgres 
         Height          =   285
         Index           =   1
         Left            =   330
         TabIndex        =   22
         Top             =   6120
         Width           =   5295
      End
      Begin VB.Label lblProgres 
         Height          =   285
         Index           =   0
         Left            =   330
         TabIndex        =   21
         Top             =   5760
         Width           =   5265
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
         Index           =   5
         Left            =   360
         TabIndex        =   20
         Top             =   1440
         Width           =   855
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   8
         Left            =   1290
         MouseIcon       =   "frmCargaFactVar.frx":0554
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar f.pago"
         Top             =   1440
         Width           =   240
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   7
         Left            =   1290
         Picture         =   "frmCargaFactVar.frx":06A6
         ToolTipText     =   "Buscar fecha"
         Top             =   960
         Width           =   240
      End
      Begin VB.Label Label4 
         Caption         =   "F.Factura"
         ForeColor       =   &H00972E0B&
         Height          =   255
         Index           =   4
         Left            =   360
         TabIndex        =   19
         Top             =   960
         Width           =   1815
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
         Index           =   2
         Left            =   360
         TabIndex        =   16
         Top             =   1860
         Width           =   1215
      End
      Begin VB.Label Label4 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   1
         Left            =   720
         TabIndex        =   15
         Top             =   2475
         Width           =   420
      End
      Begin VB.Label Label4 
         Caption         =   "Desde"
         Height          =   195
         Index           =   0
         Left            =   720
         TabIndex        =   14
         Top             =   2130
         Width           =   465
      End
   End
End
Attribute VB_Name = "frmCargaFactVar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' +-+-+-+-+-+-+-+-+-+-+-
' +-+- Autor MONICA +-+-
' +-+-+-+-+-+-+-+-+-+-+-

Option Explicit

Public NumCod As String 'Para indicar cod. Traspaso,Movimiento, etc. que llama
                        'Para indicar nº oferta a imprimir

Public CadTag As String 'Cadena con el Tag del campo que se va a poner en D/H en los listados
                        'Se necesita si el tipo de codigo es texto

    
Private Conexion As Byte
'1.- Conexión a BD Ariges  2.- Conexión a BD Conta

Private HaDevueltoDatos As Boolean

Private WithEvents frmSec As frmManSecciones 'Secciones
Attribute frmSec.VB_VarHelpID = -1
Private WithEvents frmCtas As frmCtasConta 'Cuentas contables
Attribute frmCtas.VB_VarHelpID = -1
Private WithEvents frmC As frmCal 'calendario fechas
Attribute frmC.VB_VarHelpID = -1
Private WithEvents frmFPa As frmForpaConta 'formas de pago
Attribute frmFPa.VB_VarHelpID = -1
Private WithEvents frmCon As frmManConceptos ' conceptos
Attribute frmCon.VB_VarHelpID = -1
Private WithEvents frmMens As frmAyuda ' ayuda de cuentas contables
Attribute frmMens.VB_VarHelpID = -1


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

Dim BdConta As Integer

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
Dim cWhere As String

    If Not DatosOk Then Exit Sub
            
    If AbrirConexionContaFac(vParamAplic.UsuarioContaFac, vParamAplic.PasswordContaFac, BdConta) Then
        Set vEmpresaFac = New CempresaFac
        If vEmpresaFac.LeerNiveles Then
        
            cWhere = " codmacta >= '" & Trim(txtCodigo(4).Text) & "' and codmacta <= '" & Trim(txtCodigo(5).Text) & "'"
        
            Set frmMens = New frmAyuda
        
            frmMens.OpcionMensaje = 21
            frmMens.Label5 = "Cuentas Contables"
            frmMens.cadwhere = " and  " & cWhere
            frmMens.Show vbModal
        
            Set frmMens = Nothing
            
            If cadSelect <> "" Then
                GenerarFacturas cadTABLA, cadSelect, NumError, MensError
                'Eliminar la tabla TMP
                BorrarTMP
            End If
            'Desbloqueamos ya no estamos contabilizando facturas
            DesBloqueoManual ("VENCON") 'VENtas CONtabilizar
        
        End If
        Set vEmpresaFac = Nothing
        CerrarConexionContaFac
        
        If cadSelect = "" Then
            MsgBox "No se ha realizado el proceso, no se han seleccionado cuentas.", vbExclamation
            Pb1.visible = False
            lblProgres(0).Caption = ""
            lblProgres(1).Caption = ""
            Exit Sub
        End If
    End If
    
eError:
    If Err.Number <> 0 Or NumError <> 0 Then
        MuestraError Err.Number, "No se ha realizado el proceso de generación. Llame a soporte." & vbCrLf & vbCrLf & MensError
    Else
        MsgBox "Proceso realizado correctamente.", vbExclamation
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
        PonerFoco txtCodigo(6)
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
Dim h As Integer, w As Integer
Dim List As Collection

    PrimeraVez = True
    Limpiar Me

    'IMAGES para busqueda
     Me.imgBuscar(0).Picture = frmPpal.imgListImages16.ListImages(1).Picture
     Me.imgBuscar(1).Picture = frmPpal.imgListImages16.ListImages(1).Picture
     Me.imgBuscar(2).Picture = frmPpal.imgListImages16.ListImages(1).Picture
     Me.imgBuscar(8).Picture = frmPpal.imgListImages16.ListImages(1).Picture
     Me.imgBuscar(6).Picture = frmPpal.imgListImages16.ListImages(1).Picture
     
    '###Descomentar
'    CommitConexion
         
    FrameCobrosVisible True, h, w
    indFrame = 5
    tabla = "cabfact"
    
    Pb1.visible = False
    
    'Esto se consigue poneinedo el cancel en el opcion k corresponda
    Me.cmdCancel.Cancel = True
    Me.Width = w + 70
    Me.Height = h + 350
    

End Sub

Private Sub frmC_Selec(vFecha As Date)
 'Fecha
    txtCodigo(CByte(imgFec(7).Tag)).Text = Format(vFecha, "dd/MM/yyyy")
End Sub

Private Sub frmCon_DatoSeleccionado(CadenaSeleccion As String)
' concepto de factura
    txtCodigo(indCodigo).Text = RecuperaValor(CadenaSeleccion, 1) 'codigo de concepto
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2) 'nombre de concepto
End Sub

Private Sub frmCtas_DatoSeleccionado(CadenaSeleccion As String)
' cta de banco
    txtCodigo(indCodigo).Text = RecuperaValor(CadenaSeleccion, 1) 'codmacta
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2) 'nommacta
End Sub

Private Sub frmFpa_DatoSeleccionado(CadenaSeleccion As String)
' forma de pago
    txtCodigo(indCodigo).Text = RecuperaValor(CadenaSeleccion, 1) 'codforpa
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2) 'nomforpa
End Sub

Private Sub frmMens_DatoSeleccionado(CadenaSeleccion As String)
Dim Sql As String
Dim sql2 As String

    If CadenaSeleccion <> "" Then
        Sql = " cuentas.codmacta in (" & CadenaSeleccion & ")"
        sql2 = " cuentas.codmacta in [" & CadenaSeleccion & "]"
    Else
        Sql = ""
    End If
    If Not AnyadirAFormula(cadSelect, Sql) Then Exit Sub
    If Not AnyadirAFormula(cadFormula, sql2) Then Exit Sub

End Sub

Private Sub frmSec_DatoSeleccionado(CadenaSeleccion As String)
'Form de Consulta de Secciones
Dim Cad As String

    txtCodigo(indCodigo).Text = Format(RecuperaValor(CadenaSeleccion, 1), "00")
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
    
    Cad = RecuperaValor(CadenaSeleccion, 5)  'numconta
    If Cad <> "" Then BdConta = CInt(Cad)  'numero de conta

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
    imgFec(7).Tag = Index 'independentment de les dates que tinga, sempre pose l'index en la 27
    If txtCodigo(Index).Text <> "" Then frmC.NovaData = txtCodigo(Index).Text

    frmC.Show vbModal
    Set frmC = Nothing
    PonerFoco txtCodigo(CByte(imgFec(7).Tag) + 2)
    ' ***************************
End Sub

Private Sub ListView1_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub imgBuscar_Click(Index As Integer)
   Select Case Index
        Case 0 ' Concepto
            indCodigo = 0
            Set frmCon = New frmManConceptos
            frmCon.DatosADevolverBusqueda = "0|1|2|4|"
            frmCon.CodigoActual = txtCodigo(0).Text
            frmCon.Show vbModal
            Set frmCon = Nothing
        
        Case 1, 2 ' Ctas Contables de Socio
            If BdConta = 0 Then Exit Sub
            
            AbrirFrmCuentas (Index + 3)
            
        Case 6 ' Seccion
            AbrirFrmSeccion (Index)
        
        Case 8 ' Forma de Pago
            If BdConta = 0 Then Exit Sub
            
            AbrirFrmForpaConta (Index)
    
    End Select
    PonerFoco txtCodigo(indCodigo)
End Sub

Private Sub Optcodigo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then 'ENTER
        PonerFocoBtn Me.cmdAceptar
    ElseIf KeyAscii = 27 Then Unload Me  'ESC
    End If
End Sub

Private Sub OptNombre_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then 'ENTER
        PonerFocoBtn Me.cmdAceptar
    ElseIf KeyAscii = 27 Then Unload Me  'ESC
    End If
End Sub

Private Sub Option1_Click(Index As Integer)
    Select Case Index
        Case 0
            Me.Caption = "Facturas por Cliente"
        Case 1
            Me.Caption = "Facturas por Tarjeta"
        Case 2
            Me.Caption = "Facturas por Cliente y por Tarjeta"
    End Select
    
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
            Case 6: KEYBusqueda KeyAscii, 6 'seccion
            Case 8: KEYBusqueda KeyAscii, 8 'forma de pago
            Case 4: KEYBusqueda KeyAscii, 1 'cta contable desde
            Case 5: KEYBusqueda KeyAscii, 2 'cta contable hasta
            Case 0: KEYBusqueda KeyAscii, 0 'concepto
            Case 7: KEYFecha KeyAscii, 7 'fecha factura
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
Dim cadMen As String
Dim BdConta1 As Integer
    'Quitar espacios en blanco por los lados
    txtCodigo(Index).Text = Trim(txtCodigo(Index).Text)
    
    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub

    
    Select Case Index
        Case 6 ' Seccion
            BdConta = 0
            txtNombre(Index).Text = PonerNombreDeCod(txtCodigo(Index), "seccion", "nomsecci", "codsecci", "N")
            If txtCodigo(Index).Text <> "" Then
                txtCodigo(Index).Text = Format(txtCodigo(Index).Text, "00")
            
                Cad = DevuelveDesdeBDNew(cPTours, "seccion", "numconta", "codsecci", txtCodigo(6).Text, "N") 'numconta
                If Cad <> "" Then BdConta = CInt(Cad)  'numero de conta
            End If
        
        Case 4, 5 ' Cuenta de contables
            If BdConta = 0 Then Exit Sub
            
            If AbrirConexionContaFac(vParamAplic.UsuarioContaFac, vParamAplic.PasswordContaFac, BdConta) Then
                Set vEmpresaFac = New CempresaFac
                If vEmpresaFac.LeerNiveles Then
                    txtNombre(Index).Text = PonerNombreCuenta(txtCodigo(Index), 1, , BdConta, True) 'DevuelveDesdeBDNewFac("cuentas", "nommacta", "codmacta", txtCodigo(Index), "N")
                End If
                Set vEmpresaFac = Nothing
                CerrarConexionContaFac
            End If
            
        Case 8 ' Forma de pago
            If BdConta = 0 Then Exit Sub
            
            If AbrirConexionContaFac(vParamAplic.UsuarioContaFac, vParamAplic.PasswordContaFac, BdConta) Then
                Set vEmpresaFac = New CempresaFac
                If vEmpresaFac.LeerNiveles Then
                    If vParamAplic.ContabilidadNueva Then
                        txtNombre(8).Text = DevuelveDesdeBDNewFac("formapago", "nomforpa", "codforpa", txtCodigo(8).Text, "N")
                    Else
                        txtNombre(8).Text = DevuelveDesdeBDNewFac("sforpa", "nomforpa", "codforpa", txtCodigo(8).Text, "N")
                    End If
                    If txtNombre(8).Text = "" Then
                        MsgBox "No existe la Forma de Pago. Reintroduzca.", vbExclamation
                        PonerFoco txtCodigo(Index)
                    End If
                End If
                Set vEmpresaFac = Nothing
                CerrarConexionContaFac
            End If
        
        Case 7  'FECHA FACTURA
            If txtCodigo(Index).Text <> "" Then PonerFormatoFecha txtCodigo(Index)
        
        Case 0 ' Concepto
            If PonerFormatoEntero(txtCodigo(Index)) Then
                txtNombre(0).Text = PonerNombreDeCod(txtCodigo(Index), "concefact", "nomconce", "codconce", "N")
                If txtNombre(0).Text = "" Then
                    cadMen = "No existe el Concepto: " & txtCodigo(Index).Text & vbCrLf
                    cadMen = cadMen & "¿Desea crearlo?" & vbCrLf
                    If MsgBox(cadMen, vbQuestion + vbYesNo) = vbYes Then
                        Set frmCon = New frmManConceptos
                        frmCon.DatosADevolverBusqueda = "0|1|"
                        frmCon.NuevoCodigo = txtCodigo(Index).Text
                        txtCodigo(Index).Text = ""
                        TerminaBloquear
                        frmCon.Show vbModal
                        Set frmCon = Nothing
                    Else
                        txtCodigo(Index).Text = ""
                    End If
                    PonerFoco txtCodigo(Index)
                Else
                    BdConta1 = PonerNombreDeCod(txtCodigo(Index), "concefact", "numconta", "codconce", "N")
                    
                    If BdConta1 <> BdConta Then
                        MsgBox "La Conta de este concepto ha de ser la misma que la de la sección. Reintroduzca.", vbExclamation
                        txtCodigo(Index).Text = ""
                        PonerFoco txtCodigo(0)
                    End If
                End If
            Else
                txtNombre(0).Text = ""
            End If
        
        Case 1 ' cantidad
            PonerFormatoDecimal txtCodigo(Index), 3
            txtCodigo(3).Text = Round2(CCur(ComprobarCero(txtCodigo(1).Text)) * CCur(ComprobarCero(txtCodigo(2).Text)), 2)
            PonerFormatoDecimal txtCodigo(3), 3
        
        Case 2 ' precio
            PonerFormatoDecimal txtCodigo(Index), 8
            txtCodigo(3).Text = Round2(CCur(ComprobarCero(txtCodigo(1).Text)) * CCur(txtCodigo(2).Text), 2)
            PonerFormatoDecimal txtCodigo(3), 3
            
        Case 3 ' importe
            PonerFormatoDecimal txtCodigo(Index), 3
        
        
    End Select
End Sub

Private Sub FrameCobrosVisible(visible As Boolean, ByRef h As Integer, ByRef w As Integer)
    Me.FrameCobros.visible = visible
    If visible = True Then
        Me.FrameCobros.Top = -90
        Me.FrameCobros.Left = 0
        Me.FrameCobros.Height = 7365
        Me.FrameCobros.Width = 6375
        w = Me.FrameCobros.Width
        h = Me.FrameCobros.Height
    End If
End Sub

Private Sub ValoresPorDefecto()
    txtCodigo(7).Text = Format(Now, "dd/mm/yyyy")
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

Private Sub AbrirFrmCuentas(indice As Integer)
    If BdConta = 0 Then Exit Sub
    
    If AbrirConexionContaFac(vParamAplic.UsuarioContaFac, vParamAplic.PasswordContaFac, BdConta) Then
        Set vEmpresaFac = New CempresaFac
        If vEmpresaFac.LeerNiveles Then
            indCodigo = indice
            Set frmCtas = New frmCtasConta
            frmCtas.Conexion = BdConta
            frmCtas.Facturas = True
            frmCtas.CadBusqueda = DevuelveDesdeBDNew(cPTours, "seccion", "raizcta", "codsecci", txtCodigo(6).Text, "N")
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

Private Sub AbrirFrmForpaConta(indice As Integer)
    If BdConta = 0 Then Exit Sub
    
    If AbrirConexionContaFac(vParamAplic.UsuarioContaFac, vParamAplic.PasswordContaFac, BdConta) Then
        Set vEmpresaFac = New CempresaFac
        If vEmpresaFac.LeerNiveles Then
'                    txtAux(6) = PonerNombreCuenta(txtAux(5), Modo, cContaFac)
            indCodigo = 8
            Set frmFPa = New frmForpaConta
            frmFPa.Conexion = BdConta
            frmFPa.Facturas = True
            frmFPa.DatosADevolverBusqueda = "0|1|"
            frmFPa.CodigoActual = txtCodigo(8).Text
            frmFPa.Show vbModal
            Set frmFPa = Nothing
            PonerFoco txtCodigo(8)
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

Private Function DatosOk() As Boolean
Dim b As Boolean
Dim Cad As String
Dim UltNiv As Integer
Dim PorcIva As String


    DatosOk = False

    If txtCodigo(6).Text = "" Then
        MsgBox "Debe introducir obligatoriamente una Sección.", vbExclamation
        PonerFoco txtCodigo(6)
        Exit Function
    Else
        Cad = DevuelveDesdeBDNew(cPTours, "seccion", "numconta", "codsecci", txtCodigo(6).Text, "N") 'numconta
        If Cad <> "" Then BdConta = CInt(Cad)  'numero de conta
    End If


    If txtCodigo(7).Text = "" Then
        MsgBox "Debe introducir obligatoriamente una Fecha de Factura.", vbExclamation
        PonerFoco txtCodigo(7)
        Exit Function
    End If
    
    
    'cuenta contable desde
    If txtCodigo(4).Text <> "" Then
        If BdConta = 0 Then
            MsgBox "No hay conexion a la contabilidad de la seccion. Revise", vbExclamation
            Exit Function
        Else
            b = True
            If AbrirConexionContaFac(vParamAplic.UsuarioContaFac, vParamAplic.PasswordContaFac, CByte(BdConta)) Then
                Set vEmpresaFac = New CempresaFac
                If vEmpresaFac.LeerNiveles Then
                    txtNombre(4) = PonerNombreCuenta(txtCodigo(4), 1, , BdConta, True)
                    If txtNombre(4) = "" Then
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
                        If Mid(txtCodigo(4), 1, UltNiv) <> DevuelveDesdeBDNew(cPTours, "seccion", "raizcta", "codsecci", txtCodigo(6), "N") Then
                            If MsgBox("La Cuenta Contable no tiene la misma raiz que la sección." & vbCrLf & "          ¿ Desea Continuar ?", vbQuestion + vbYesNo + vbDefaultButton1) = vbNo Then
                                b = False
                            End If
                        End If
                    End If
                End If
                Set vEmpresaFac = Nothing
                CerrarConexionContaFac
                If Not b Then Exit Function
            End If
        End If
    End If
    
    ' cuenta contable hasta
    If txtCodigo(5).Text <> "" Then
        If BdConta = 0 Then
            MsgBox "No hay conexion a la contabilidad de la seccion. Revise", vbExclamation
            Exit Function
        Else
            b = True
            If AbrirConexionContaFac(vParamAplic.UsuarioContaFac, vParamAplic.PasswordContaFac, CByte(BdConta)) Then
                Set vEmpresaFac = New CempresaFac
                If vEmpresaFac.LeerNiveles Then
                    txtNombre(5) = PonerNombreCuenta(txtCodigo(5), 1, , BdConta, True)
                    If txtNombre(5) = "" Then
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
                        If Mid(txtCodigo(5), 1, UltNiv) <> DevuelveDesdeBDNew(cPTours, "seccion", "raizcta", "codsecci", txtCodigo(6), "N") Then
                            If MsgBox("La Cuenta Contable no tiene la misma raiz que la sección." & vbCrLf & "          ¿ Desea Continuar ?", vbQuestion + vbYesNo + vbDefaultButton1) = vbNo Then
                                b = False
                            End If
                        End If
                    End If
                End If
                Set vEmpresaFac = Nothing
                CerrarConexionContaFac
                
                If Not b Then Exit Function
            End If
        End If
    End If
    
    If txtCodigo(4).Text = "" Or txtCodigo(5).Text = "" Then
        MsgBox "Debe introducir un valor en las cuentas contables", vbExclamation
        PonerFoco txtCodigo(4)
        Exit Function
    Else
        If Len(txtCodigo(4).Text) = Len(txtCodigo(5).Text) Then
            If CDbl(txtCodigo(4).Text) > CDbl(txtCodigo(5).Text) Then
                MsgBox "La cuenta desde ha de ser inferior a la cuenta hasta. Revise.", vbExclamation
                PonerFoco txtCodigo(4)
                Exit Function
            End If
        Else
            MsgBox "Las cuentas contables deben de tener el mismo nivel. Revise.", vbExclamation
            PonerFoco txtCodigo(4)
            Exit Function
        End If
    End If
    
    If txtCodigo(8).Text = "" Then
        MsgBox "Debe introducir obligatoriamente una forma de pago.", vbExclamation
        PonerFoco txtCodigo(8)
        Exit Function
    Else
        If BdConta = 0 Then Exit Function
        b = True
        If AbrirConexionContaFac(vParamAplic.UsuarioContaFac, vParamAplic.PasswordContaFac, BdConta) Then
            Set vEmpresaFac = New CempresaFac
            If vEmpresaFac.LeerNiveles Then
                If vParamAplic.ContabilidadNueva Then
                    txtNombre(8).Text = DevuelveDesdeBDNewFac("formapago", "nomforpa", "codforpa", txtCodigo(8).Text, "N")
                Else
                    txtNombre(8).Text = DevuelveDesdeBDNewFac("sforpa", "nomforpa", "codforpa", txtCodigo(8).Text, "N")
                End If
                If txtNombre(8).Text = "" Then
                    MsgBox "No existe la Forma de Pago. Reintroduzca.", vbExclamation
                    PonerFoco txtCodigo(8)
                    b = False
                End If
            End If
            Set vEmpresaFac = Nothing
            CerrarConexionContaFac
            
            If Not b Then Exit Function
        End If
    End If
        
    If txtCodigo(0).Text = "" Then
        MsgBox "Debe introducir obligatoriamente un concepto.", vbExclamation
        PonerFoco txtCodigo(0)
        Exit Function
    Else
        Cad = ""
        Cad = DevuelveDesdeBDNew(cPTours, "concefact", "tipoiva", "codconce", txtCodigo(0).Text, "N")
        If Cad = "" Then
            MsgBox "El concepto no tiene asociado un tipo de iva. Revise.", vbExclamation
            PonerFoco txtCodigo(0)
            Exit Function
        Else
            ' comprobamos que existe el tipo de iva en contabilidad
            If BdConta = 0 Then Exit Function
            b = True
            If AbrirConexionContaFac(vParamAplic.UsuarioContaFac, vParamAplic.PasswordContaFac, BdConta) Then
                Set vEmpresaFac = New CempresaFac
                If vEmpresaFac.LeerNiveles Then
                    PorcIva = DevuelveDesdeBDNewFac("tiposiva", "porceiva", "codigiva", Cad, "N")
                    If PorcIva = "" Then
                        MsgBox "No existe el tipo de Iva del concepto. Revise.", vbExclamation
                        PonerFoco txtCodigo(0)
                        b = False
                    End If
                End If
                Set vEmpresaFac = Nothing
                CerrarConexionContaFac
            End If
            If Not b Then Exit Function
        End If
    End If
    
    DatosOk = True
End Function

Private Sub GenerarFacturas(cadTABLA As String, cadwhere As String, NumError As Long, MensError As String)
Dim Sql As String
Dim b As Boolean
Dim tmpErrores As Boolean 'Indica si se creo correctamente la tabla de errores
Dim CCoste As String
Dim Cad As String
Dim NumF As Long
Dim CabSql As String
Dim LinSql As String

Dim vSec As CSeccion
Dim NumFact As Long

Dim TipoIva As String
Dim PorIva As String
Dim Impoiva As Currency
Dim TotalFact As Currency

Dim Rs As ADODB.Recordset
Dim NomCuenta As String
Dim Existe As Boolean

    On Error GoTo EContab


    Sql = "GENFAC" 'generar facturas de venta

    'Bloquear para que nadie mas pueda contabilizar
    DesBloqueoManual (Sql)
    If Not BloqueoManual(Sql, "1") Then
        MsgBox "No se pueden Generar Facturas. Hay otro usuario realizando el proceso.", vbExclamation
        Screen.MousePointer = vbDefault
        Exit Sub
    End If

    conn.BeginTrans

    BorrarTMP
    'Cargar tabla TEMP con las Facturas que vamos a Trabajar
    b = CrearTMP("cuentas", cadwhere, True)
    If Not b Then Exit Sub
            
    'Visualizar la barra de Progreso
    Me.Pb1.visible = True
    
    NumF = DevuelveValor("select count(*) from ariagroutil.tmpfactu")
    
    Me.lblProgres(0).Caption = "Comprobaciones: "
    CargarProgresNew Me.Pb1, CInt(NumF)
        
    Sql = "select ctaclien from ariagroutil.tmpfactu order by ctaclien"
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockReadOnly, adCmdText
        
    While Not Rs.EOF
    
        IncrementarProgresNew Me.Pb1, 1
        Me.lblProgres(1).Caption = "Procesando Cuenta Contable ..."
        Me.Refresh
        
        Set vSec = New CSeccion
        If vSec.leer(txtCodigo(6).Text) Then
            NumFact = vSec.ConseguirContador(txtCodigo(6).Text)
            
            Existe = False
            Do
                Sql = "select count(*) from cabfact where codsecci = " & DBSet(txtCodigo(6).Text, "N")
                Sql = Sql & " and letraser = " & DBSet(vSec.LetraSerie, "T")
                Sql = Sql & " and numfactu = " & DBSet(NumFact, "N")
                Sql = Sql & " and fecfactu = " & DBSet(txtCodigo(7).Text, "F")
                
                If TotalRegistros(Sql) > 0 Then
                    vSec.IncrementarContador txtCodigo(6).Text
                    
                    NumFact = NumFact + 1
                    Existe = True
                Else
                    Existe = False
                End If
            Loop Until Not Existe
            
            TipoIva = ""
            PorIva = ""
            Impoiva = 0
            TotalFact = 0
            
            TipoIva = DevuelveDesdeBDNew(cPTours, "concefact", "tipoiva", "codconce", txtCodigo(0).Text, "N")
            PorIva = DevuelveDesdeBDNewFac("tiposiva", "porceiva", "codigiva", TipoIva, "N")
            Impoiva = Round2(CCur(ImporteSinFormato(txtCodigo(3).Text)) * ComprobarCero(PorIva) / 100, 2)
            TotalFact = CCur(ImporteSinFormato(txtCodigo(3).Text)) + Impoiva
            
            ' Insertamos en la cabecera de factura
            CabSql = "insert into cabfact ("
            CabSql = CabSql & "codsecci,letraser,numfactu,fecfactu,ctaclien,observac,intconta,baseiva1,baseiva2,baseiva3,"
            CabSql = CabSql & "impoiva1,impoiva2,impoiva3,imporec1,imporec2,imporec3,totalfac,tipoiva1,tipoiva2,tipoiva3,"
            CabSql = CabSql & "porciva1 , porciva2, porciva3, codforpa, porcrec1, porcrec2, porcrec3, retfaccl, trefaccl, cuereten)  values  "
            
            CabSql = CabSql & "(" & DBSet(txtCodigo(6).Text, "N")
            CabSql = CabSql & "," & DBSet(vSec.LetraSerie, "T")
            CabSql = CabSql & "," & DBSet(NumFact, "N")
            CabSql = CabSql & "," & DBSet(txtCodigo(7).Text, "F")
            CabSql = CabSql & "," & DBSet(Rs!ctaclien, "T")
            CabSql = CabSql & "," & DBSet(txtCodigo(9).Text, "T", "S")
            CabSql = CabSql & ",0"
            CabSql = CabSql & "," & DBSet(txtCodigo(3).Text, "N")
            CabSql = CabSql & "," & ValorNulo
            CabSql = CabSql & "," & ValorNulo
            CabSql = CabSql & "," & DBSet(Impoiva, "N")
            CabSql = CabSql & "," & ValorNulo
            CabSql = CabSql & "," & ValorNulo
            CabSql = CabSql & "," & ValorNulo
            CabSql = CabSql & "," & ValorNulo
            CabSql = CabSql & "," & ValorNulo
            
            CabSql = CabSql & "," & DBSet(TotalFact, "N")
            CabSql = CabSql & "," & DBSet(TipoIva, "N")
            CabSql = CabSql & "," & ValorNulo
            CabSql = CabSql & "," & ValorNulo
            CabSql = CabSql & "," & DBSet(PorIva, "N")
            CabSql = CabSql & "," & ValorNulo
            CabSql = CabSql & "," & ValorNulo
            CabSql = CabSql & "," & DBSet(txtCodigo(8).Text, "N") ' forma de pago
            CabSql = CabSql & "," & ValorNulo
            CabSql = CabSql & "," & ValorNulo
            CabSql = CabSql & "," & ValorNulo
            CabSql = CabSql & "," & ValorNulo
            CabSql = CabSql & "," & ValorNulo
            CabSql = CabSql & "," & ValorNulo
            CabSql = CabSql & ")"
            
            conn.Execute CabSql
            
            
            ' insertamos en la linea de factura
            LinSql = "insert into linfact (codsecci , letraser, numfactu, fecfactu, NumLinea, codConce, ampliaci, precio, cantidad, Importe, TipoIva) values "
            LinSql = LinSql & "(" & DBSet(txtCodigo(6).Text, "N")
            LinSql = LinSql & "," & DBSet(vSec.LetraSerie, "T")
            LinSql = LinSql & "," & DBSet(NumFact, "N")
            LinSql = LinSql & "," & DBSet(txtCodigo(7).Text, "F")
            LinSql = LinSql & ",1"
            LinSql = LinSql & "," & DBSet(txtCodigo(0).Text, "N")
            LinSql = LinSql & "," & DBSet(txtCodigo(10).Text, "T")
            LinSql = LinSql & "," & DBSet(txtCodigo(2).Text, "N")
            LinSql = LinSql & "," & DBSet(txtCodigo(1).Text, "N")
            LinSql = LinSql & "," & DBSet(txtCodigo(3).Text, "N")
            LinSql = LinSql & "," & DBSet(TipoIva, "N")
            LinSql = LinSql & ")"
            
            conn.Execute LinSql
            
            
            vSec.IncrementarContador (txtCodigo(6).Text)
            Set vSec = Nothing
                    
        End If
        
        Rs.MoveNext
    Wend
    
EContab:
    If Err.Number <> 0 Then
        NumError = Err.Number
        MensError = "Generar Facturas " '& Err.Description
        conn.RollbackTrans
    Else
        conn.CommitTrans
        
    End If
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
    codigo1 = "letraser"
    Sql = Sql & " ON " & cadTABLA & "." & codigo1 & "=tmpfactu.numserie"
    Sql = Sql & " AND " & cadTABLA & ".codsecci=tmpfactu.codsecci"
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
        CargarProgres Me.Pb1, numfactu
        
         Set cContaFra = New cContabilizarFacturas
        
        If Not cContaFra.EstablecerValoresInciales(ConnContaFac) Then
            'NO ha establcedio los valores de la conta.  Le dejaremos seguir, avisando que
            ' obviamente, no va a contabilizar las FRAS
            Sql = "Si continua, las facturas se insertaran en el registro, pero no serán contabilizadas" & vbCrLf
            Sql = Sql & "en este momento. Deberán ser contabilizadas desde el ARICONTA" & vbCrLf & vbCrLf
            Sql = Sql & Space(50) & "¿Continuar?"
            If MsgBox(Sql, vbQuestion + vbYesNoCancel) <> vbYes Then Exit Function
        End If
               
        
        
        Sql = "SELECT * "
        Sql = Sql & " FROM tmpfactu "
            
        Set Rs = New ADODB.Recordset
        Rs.Open Sql, conn, adOpenStatic, adLockPessimistic, adCmdText
        i = 1

        b = True
        'contabilizar cada una de las facturas seleccionadas
        While Not Rs.EOF
            Sql = cadTABLA & ".codsecci = " & DBSet(Rs.Fields(0), "N") & " and "
            Sql = Sql & cadTABLA & "." & codigo1 & "=" & DBSet(Rs.Fields(1), "T") & " and numfactu=" & DBLet(Rs!numfactu, "N")
            Sql = Sql & " and fecfactu=" & DBSet(Rs!fecfactu, "F")
            If PasarFacturaFac(Sql, FecVenci, Banpr, CCoste, cContaFra) = False And b Then b = False
            
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


Private Function ComprobarFechasConta(ind As Integer) As Boolean
'comprobar que el periodo de fechas a contabilizar esta dentro del
'periodo de fechas del ejercicio de la contabilidad
Dim FechaIni As String, FechaFin As String
Dim Cad As String
Dim Rs As ADODB.Recordset
    
On Error GoTo EComprobar

    ComprobarFechasConta = False
    
    If txtCodigo(ind).Text <> "" Then
        FechaIni = "Select fechaini,fechafin From parametros"
        Set Rs = New ADODB.Recordset
        Rs.Open FechaIni, ConnContaFac, adOpenForwardOnly, adLockPessimistic, adCmdText
    
        If Not Rs.EOF Then
            FechaIni = DBLet(Rs!FechaIni, "F")
            FechaFin = DateAdd("yyyy", 1, CDate(DBLet(Rs!FechaFin, "F"))) ' + 365
            'nos guardamos los valores
            Orden1 = FechaIni
            Orden2 = FechaFin
        
            If Not EntreFechas(FechaIni, txtCodigo(ind).Text, FechaFin) Then
                 Cad = "El período de contabilización debe estar dentro del ejercicio:" & vbCrLf & vbCrLf
                 Cad = Cad & "    Desde: " & FechaIni & vbCrLf
                 Cad = Cad & "    Hasta: " & FechaFin
                 MsgBox Cad, vbExclamation
                 txtCodigo(ind).Text = ""
            Else
                ComprobarFechasConta = True
            End If
        End If
        Rs.Close
        Set Rs = Nothing
    Else
        ComprobarFechasConta = True
    End If
    
EComprobar:
    If Err.Number <> 0 Then MuestraError Err.Number, "Comprobar Fechas", Err.Description
End Function



Private Sub BorrarTMP()
On Error Resume Next

    conn.Execute " DROP TABLE IF EXISTS ariagroutil.tmpfactu;"
    If Err.Number <> 0 Then Err.Clear
End Sub

Private Function CrearTMP(cadTABLA As String, cadwhere As String, Optional Facturas As Boolean, Optional Telefono As Boolean) As Boolean
'Crea una temporal donde inserta la clave primaria de las
'facturas seleccionadas para facturar y trabaja siempre con ellas
' facturas indica si viene de facturas varias o de telefonia
Dim Sql As String
    
    On Error GoTo ECrear
    
    CrearTMP = False
    
    Sql = "CREATE TABLE ariagroutil.tmpfactu ( "
    Sql = Sql & "ctaclien varchar(10) NOT NULL default '')"
    conn.Execute Sql
     
    Sql = "SELECT codmacta "
    Sql = Sql & " FROM " & cadTABLA
    Sql = Sql & " WHERE " & cadwhere
    Sql = " INSERT INTO ariagroutil.tmpfactu " & Sql
    ConnContaFac.Execute Sql

    CrearTMP = True
    
ECrear:
     If Err.Number <> 0 Then
        CrearTMP = False
        'Borrar la tabla temporal
        Sql = " DROP TABLE IF EXISTS ariagroutil.tmpfactu;"
        conn.Execute Sql
    End If
End Function


