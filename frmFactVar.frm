VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmFactVar 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Reimpresion de Facturas"
   ClientHeight    =   5175
   ClientLeft      =   45
   ClientTop       =   2430
   ClientWidth     =   7185
   Icon            =   "frmFactVar.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5175
   ScaleWidth      =   7185
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
      Height          =   4995
      Left            =   120
      TabIndex        =   11
      Top             =   120
      Width           =   6915
      Begin VB.CheckBox Check1 
         Caption         =   "Duplicado"
         Height          =   240
         Index           =   0
         Left            =   675
         TabIndex        =   28
         Top             =   4320
         Width           =   1590
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   8
         Left            =   2760
         Locked          =   -1  'True
         TabIndex        =   26
         Text            =   "Text5"
         Top             =   600
         Width           =   3135
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   8
         Left            =   1845
         MaxLength       =   6
         TabIndex        =   0
         Top             =   600
         Width           =   830
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   7
         Left            =   4245
         MaxLength       =   7
         TabIndex        =   8
         Top             =   3720
         Width           =   930
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   6
         Left            =   4245
         MaxLength       =   7
         TabIndex        =   7
         Top             =   3360
         Width           =   930
      End
      Begin VB.TextBox txtCodigo 
         Height          =   285
         Index           =   5
         Left            =   1845
         MaxLength       =   3
         TabIndex        =   6
         Top             =   3735
         Width           =   570
      End
      Begin VB.TextBox txtCodigo 
         Height          =   285
         Index           =   4
         Left            =   1845
         MaxLength       =   3
         TabIndex        =   5
         Top             =   3360
         Width           =   570
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   3
         Left            =   1845
         MaxLength       =   10
         TabIndex        =   4
         Tag             =   "C�digo Postal|T|S|||clientes|codposta|||"
         Top             =   2640
         Width           =   1050
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   2
         Left            =   1845
         MaxLength       =   10
         TabIndex        =   3
         Tag             =   "C�digo Postal|T|S|||clientes|codposta|||"
         Top             =   2310
         Width           =   1050
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Left            =   4905
         TabIndex        =   10
         Top             =   4380
         Width           =   975
      End
      Begin VB.CommandButton cmdAceptar 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   3720
         TabIndex        =   9
         Top             =   4380
         Width           =   975
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   0
         Left            =   1845
         MaxLength       =   10
         TabIndex        =   1
         Top             =   1200
         Width           =   1230
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   1
         Left            =   1845
         MaxLength       =   10
         TabIndex        =   2
         Top             =   1575
         Width           =   1230
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   0
         Left            =   3150
         Locked          =   -1  'True
         TabIndex        =   21
         Text            =   "Text5"
         Top             =   1215
         Width           =   3135
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   1
         Left            =   3150
         Locked          =   -1  'True
         TabIndex        =   22
         Text            =   "Text5"
         Top             =   1575
         Width           =   3135
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   8
         Left            =   1530
         MouseIcon       =   "frmFactVar.frx":000C
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar Secci�n"
         Top             =   600
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Secci�n"
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
         Left            =   600
         TabIndex        =   27
         Top             =   570
         Width           =   540
      End
      Begin VB.Label Label4 
         Caption         =   "Desde"
         Height          =   195
         Index           =   5
         Left            =   3360
         TabIndex        =   25
         Top             =   3360
         Width           =   465
      End
      Begin VB.Label Label4 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   4
         Left            =   3360
         TabIndex        =   24
         Top             =   3735
         Width           =   420
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Factura"
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
         Left            =   3000
         TabIndex        =   23
         Top             =   3120
         Width           =   555
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Serie"
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
         Left            =   600
         TabIndex        =   20
         Top             =   3120
         Width           =   360
      End
      Begin VB.Label Label4 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   1
         Left            =   960
         TabIndex        =   19
         Top             =   3735
         Width           =   420
      End
      Begin VB.Label Label4 
         Caption         =   "Desde"
         Height          =   195
         Index           =   0
         Left            =   960
         TabIndex        =   18
         Top             =   3360
         Width           =   465
      End
      Begin VB.Label Label4 
         Caption         =   "Fecha Factura"
         ForeColor       =   &H00972E0B&
         Height          =   255
         Index           =   16
         Left            =   600
         TabIndex        =   17
         Top             =   2040
         Width           =   1815
      End
      Begin VB.Label Label4 
         Caption         =   "Desde"
         Height          =   195
         Index           =   15
         Left            =   960
         TabIndex        =   16
         Top             =   2280
         Width           =   465
      End
      Begin VB.Label Label4 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   14
         Left            =   960
         TabIndex        =   15
         Top             =   2640
         Width           =   420
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   2
         Left            =   1530
         Picture         =   "frmFactVar.frx":015E
         ToolTipText     =   "Buscar fecha"
         Top             =   2280
         Width           =   240
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   3
         Left            =   1530
         Picture         =   "frmFactVar.frx":01E9
         ToolTipText     =   "Buscar fecha"
         Top             =   2640
         Width           =   240
      End
      Begin VB.Label Label4 
         Caption         =   "Desde"
         Height          =   195
         Index           =   13
         Left            =   960
         TabIndex        =   14
         Top             =   1200
         Width           =   465
      End
      Begin VB.Label Label4 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   12
         Left            =   960
         TabIndex        =   13
         Top             =   1575
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
         Left            =   600
         TabIndex        =   12
         Top             =   960
         Width           =   1215
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   0
         Left            =   1530
         MouseIcon       =   "frmFactVar.frx":0274
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar Cuenta Contable"
         Top             =   1200
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   1
         Left            =   1530
         MouseIcon       =   "frmFactVar.frx":03C6
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar Cuenta Contable"
         Top             =   1575
         Width           =   240
      End
   End
End
Attribute VB_Name = "frmFactVar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' +-+-+-+-+-+-+-+-+-+-+-
' +-+- Autor MANOLO +-+-
' +-+-+-+-+-+-+-+-+-+-+-

Option Explicit

Public NumCod As String ' 0 factura normal
                        ' 1 ajena
                        

Public CadTag As String 'Cadena con el Tag del campo que se va a poner en D/H en los listados
                        'Se necesita si el tipo de codigo es texto

    
Private Conexion As Byte
'1.- Conexi�n a BD Ariges  2.- Conexi�n a BD Conta

Private HaDevueltoDatos As Boolean

'Private WithEvents frmcli As frmManClien 'Clientes
'Private WithEvents frmCol As frmManCoope 'Colectivo
Private WithEvents frmC As frmCal 'calendario fechas
Attribute frmC.VB_VarHelpID = -1
Private WithEvents frmCtas As frmCtasConta
Attribute frmCtas.VB_VarHelpID = -1
Private WithEvents frmSec As frmManSecciones
Attribute frmSec.VB_VarHelpID = -1

'GENERALES PARA PASARLE A CRYSTAL REPORT
Private cadFormula As String 'Cadena con la FormulaSelection para Crystal Report
Private cadParam As String 'Cadena con los parametros para Crystal Report
Private numParam As Byte 'Numero de parametros que se pasan a Crystal Report
Private cadSelect As String 'Cadena para comprobar si hay datos antes de abrir Informe
Private cadTitulo As String 'Titulo para la ventana frmImprimir
Private cadNombreRPT As String 'Nombre del informe

Dim indCodigo As Integer 'indice para txtCodigo
Dim indFrame As Single 'n� de frame en el que estamos

Dim BdConta As Integer ' numero de la contabilidad donde se hace conexion

'Se inicializan para cada Informe (tabla de BD a la que hace referencia
Dim tabla As String
Dim Codigo As String 'C�digo para FormulaSelection de Crystal Report
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


Private Sub cmdAceptar_Click()
Dim cDesde As String, cHasta As String 'cadena codigo Desde/Hasta
Dim nDesde As String, nHasta As String 'cadena Descripcion Desde/Hasta
Dim cadTABLA As String, cOrden As String
Dim i As Byte
Dim b As Boolean
    InicializarVbles
    
    If txtCodigo(8).Text = "" Then
        MsgBox "Introduzca la Secci�n.", vbExclamation
        Exit Sub
    End If
    
    b = AnyadirAFormula(cadFormula, "{" & tabla & ".codsecci} = " & txtCodigo(8).Text)
    b = AnyadirAFormula(cadSelect, tabla & ".codsecci = " & DBSet(txtCodigo(8).Text, "N"))
    
    '========= PARAMETROS  =============================
    'A�adir el parametro de Empresa
    cadParam = cadParam & "|pEmpresa=""" & vEmpresa.nomEmpre & """|"
    numParam = numParam + 1
    
    'D/H Cuenta contable
    cDesde = Trim(txtCodigo(0).Text)
    cHasta = Trim(txtCodigo(1).Text)
    nDesde = txtNombre(0).Text
    nHasta = txtNombre(1).Text
    If Not (cDesde = "" And cHasta = "") Then
        'Cadena para seleccion Desde y Hasta
        Codigo = "{" & tabla & ".ctaclien}"
        TipCod = "T"
        If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHcliente= """) Then Exit Sub
    End If
    
    'D/H Fecha factura
    cDesde = Trim(txtCodigo(2).Text)
    cHasta = Trim(txtCodigo(3).Text)
    If Not (cDesde = "" And cHasta = "") Then
        'Cadena para seleccion Desde y Hasta
        Codigo = "{" & tabla & ".fecfactu}"
        TipCod = "F"
        If Not PonerDesdeHasta(cDesde, cHasta, "", "", "pDHfechaFac= """) Then Exit Sub
    End If
    
    'D/H Serie
    cDesde = Trim(txtCodigo(4).Text)
    cHasta = Trim(txtCodigo(5).Text)
    If Not (cDesde = "" And cHasta = "") Then
        'Cadena para seleccion Desde y Hasta
        Codigo = "{" & tabla & ".letraser}"
        TipCod = "T"
        If Not PonerDesdeHasta(cDesde, cHasta, "", "", "pDHSerie= """) Then Exit Sub
    End If
    
    'Factura
    cDesde = Trim(txtCodigo(6).Text)
    cHasta = Trim(txtCodigo(7).Text)
    If Not (cDesde = "" And cHasta = "") Then
        'Cadena para seleccion Desde y Hasta
        Codigo = "{" & tabla & ".numfactu}"
        TipCod = "N"
        If Not PonerDesdeHasta(cDesde, cHasta, "", "", "pDHFact= """) Then Exit Sub
    End If
    
    
    
    If CBool(Me.Check1(0).Value) Then
        cadParam = cadParam & "pDuplicado=1|"
    Else
        cadParam = cadParam & "pDuplicado=0|"
    End If
    numParam = numParam + 1
    

    cadTABLA = "cabfact"
    If HayRegParaInforme(cadTABLA, cadSelect) Then
        cadTitulo = "Reimpresion de Facturas"
       
        Dim indRPT As Byte 'Indica el tipo de Documento en la tabla "scryst"
        Dim nomDocu As String 'Nombre de Informe rpt de crystal
        
        indRPT = 1 'Facturas Varias
        
        '[Monica]26/05/2016: otro report para materna
        If EsSeccionMaterna(txtCodigo(8).Text) Then indRPT = 4
        
        
       If Not PonerParamRPT(indRPT, cadParam, numParam, nomDocu) Then Exit Sub
       ' he a�adido estas dos lineas para que llame al rpt correspondiente
       frmImprimir.NombreRPT = nomDocu
       cadNombreRPT = nomDocu
       LlamarImprimir
    End If
    
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
    If PrimeraVez Then
        PrimeraVez = False
        BdConta = 0
'        If NumCod = 1 Then
            PonerFoco txtCodigo(8)
'            Option1(1).Value = True
'        Else
'            PonerFoco txtCodigo(0)
'        End If
        
      ' PonerFoco txtCodigo(0)
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
Dim h As Integer, w As Integer
Dim List As Collection

    PrimeraVez = True
    Limpiar Me

'    Label4(6).visible = (NumCod = 1)
'    txtCodigo(8).visible = (NumCod = 1)
'    txtCodigo(8).Enabled = (NumCod = 1)
'    txtNombre(8).visible = (NumCod = 1)
'    txtNombre(8).Enabled = (NumCod = 1)
'    Me.imgBuscar(8).visible = (NumCod = 1)
'    Me.imgBuscar(8).Enabled = (NumCod = 1)
'    ChkTipoDocu(0).Value = 0

    'IMAGES para busqueda
     Me.imgBuscar(0).Picture = frmPpal.imgListImages16.ListImages(1).Picture
     Me.imgBuscar(1).Picture = frmPpal.imgListImages16.ListImages(1).Picture
     Me.imgBuscar(8).Picture = frmPpal.imgListImages16.ListImages(1).Picture

    '###Descomentar
'    CommitConexion
         
    FrameCobrosVisible True, h, w
    indFrame = 5
'    If NumCod = 0 Then
    tabla = "cabfact"
'    Else
'        Me.Caption = Me.Caption & " Ajenas"
'        tabla = "schfacr" ' historico del Regaixo
'    End If
            
    'Esto se consigue poneinedo el cancel en el opcion k corresponda
    Me.cmdCancel.Cancel = True



End Sub

Private Sub frmC_Selec(vFecha As Date)
 'Fecha
    txtCodigo(CByte(imgFec(2).Tag)).Text = Format(vFecha, "dd/MM/yyyy")
End Sub

Private Sub frmCli_DatoSeleccionado(CadenaSeleccion As String)
'Form de Consulta de Clientes
    txtCodigo(indCodigo).Text = Format(RecuperaValor(CadenaSeleccion, 1), "000000")
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
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
            If BdConta = 0 Then Exit Sub
            
            AbrirFrmCuentas (Index)
        
        Case 8 ' seccion
            AbrirFrmSeccion (Index)
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
            Case 6: KEYBusqueda KeyAscii, 6 'numero de factura desde
            Case 7: KEYBusqueda KeyAscii, 7 'numero de factura hasta
            Case 2: KEYFecha KeyAscii, 3 'fecha desde
            Case 3: KEYFecha KeyAscii, 3 'fecha hasta
            Case 8: KEYBusqueda KeyAscii, 8 'colectivo
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
            
        Case 0, 1 'Cuenta Cliente
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
            
        Case 2, 3 'FECHAS
            If txtCodigo(Index).Text <> "" Then PonerFormatoFecha txtCodigo(Index)
            
        Case 4, 5 'SERIE
            txtCodigo(Index).Text = UCase(txtCodigo(Index).Text)
        
        Case 6, 7 'FACTURAS
            txtCodigo(Index).Text = Format(txtCodigo(Index).Text, "0000000")
            
        Case 8 'Seccion
            BdConta = 0
            txtNombre(Index).Text = PonerNombreDeCod(txtCodigo(Index), "seccion", "nomsecci", "codsecci", "N")
            If txtCodigo(Index).Text <> "" Then
                txtCodigo(Index).Text = Format(txtCodigo(Index).Text, "000")
            
                Cad = DevuelveDesdeBDNew(cPTours, "seccion", "numconta", "codsecci", txtCodigo(8).Text, "N") 'numconta
                If Cad <> "" Then BdConta = CByte(Cad)  'numero de conta
            Else
                MsgBox "Debe introducir un c�digo existente en la secci�n.", vbExclamation
            End If
        
    End Select
End Sub

Private Sub FrameCobrosVisible(visible As Boolean, ByRef h As Integer, ByRef w As Integer)
    Me.FrameCobros.visible = visible
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
'A�ade a cadFormula y cadSelect la cadena de seleccion:
'       "(codigo>=codD AND codigo<=codH)"
' y a�ade a cadParam la cadena para mostrar en la cabecera informe:
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
        .Facturas = True
        .Contabilidad = BdConta
        .EnvioEMail = False
        .Show vbModal
    End With
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

Public Function InsertarNrosFacturaEnTemporal(cTabla As String, cWhere As String) As Boolean
'Comprobar si hay registros a Mostrar antes de abrir el Informe
Dim Sql As String
Dim sql2 As String

    On Error GoTo eInsertarNrosFacturaEnTemporal

    InsertarNrosFacturaEnTemporal = False

    sql2 = "delete from tmpfacturas where codusu = " & vSesion.Codigo
    conn.Execute sql2

    Sql = "Select distinct " & vSesion.Codigo & "," & tabla & ".letraser," & tabla & ".numfactu" & "," & tabla & ".fecfactu " & " FROM " & QuitarCaracterACadena(cTabla, "_1")
    If cWhere <> "" Then
        cWhere = QuitarCaracterACadena(cWhere, "{")
        cWhere = QuitarCaracterACadena(cWhere, "}")
        cWhere = QuitarCaracterACadena(cWhere, "_1")
        Sql = Sql & " WHERE " & cWhere
    End If
    
    sql2 = "insert into tmpfacturas (codusu, letraser, numfactu, fecfactu) " & Sql
    conn.Execute sql2
    
eInsertarNrosFacturaEnTemporal:
    If Err.Number = 0 Then
        InsertarNrosFacturaEnTemporal = True
    Else
        MsgBox "Se ha producido un error en la carga de la tabla intermedia", vbExclamation
    End If
End Function

