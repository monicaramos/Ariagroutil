VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmContabFact 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Contabilización de Facturas "
   ClientHeight    =   5955
   ClientLeft      =   45
   ClientTop       =   2430
   ClientWidth     =   6645
   Icon            =   "frmContabFact.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5955
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
      Height          =   5535
      Left            =   90
      TabIndex        =   10
      Top             =   120
      Width           =   6375
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   6
         Left            =   1575
         MaxLength       =   6
         TabIndex        =   27
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
         TabIndex        =   26
         Text            =   "Text5"
         Top             =   495
         Width           =   3135
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   8
         Left            =   1590
         MaxLength       =   10
         TabIndex        =   1
         Tag             =   "Código Propio|N|N|1|99|sbanco|codbanpr|00|S|"
         Top             =   1530
         Width           =   1050
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   8
         Left            =   2715
         Locked          =   -1  'True
         TabIndex        =   21
         Text            =   "Text5"
         Top             =   1530
         Width           =   2910
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   7
         Left            =   1605
         MaxLength       =   10
         TabIndex        =   0
         Tag             =   "Código Postal|T|S|||clientes|codposta|||"
         Top             =   1020
         Width           =   1050
      End
      Begin VB.TextBox txtCodigo 
         Height          =   285
         Index           =   5
         Left            =   3720
         MaxLength       =   3
         TabIndex        =   3
         Top             =   2115
         Width           =   405
      End
      Begin VB.TextBox txtCodigo 
         Height          =   285
         Index           =   4
         Left            =   1605
         MaxLength       =   3
         TabIndex        =   2
         Top             =   2100
         Width           =   405
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   3
         Left            =   3735
         MaxLength       =   10
         TabIndex        =   7
         Tag             =   "Código Postal|T|S|||clientes|codposta|||"
         Top             =   3330
         Width           =   1050
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   2
         Left            =   1605
         MaxLength       =   10
         TabIndex        =   6
         Tag             =   "Código Postal|T|S|||clientes|codposta|||"
         Top             =   3300
         Width           =   1050
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Left            =   4755
         TabIndex        =   9
         Top             =   4980
         Width           =   975
      End
      Begin VB.CommandButton cmdAceptar 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   3570
         TabIndex        =   8
         Top             =   4980
         Width           =   975
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   0
         Left            =   1605
         MaxLength       =   7
         TabIndex        =   4
         Tag             =   "Nº de Factura|N|N|0|9999999|schfac|numfactu|0000000|S|"
         Top             =   2700
         Width           =   830
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   1
         Left            =   3720
         MaxLength       =   7
         TabIndex        =   5
         Tag             =   "Nº de Factura|N|N|0|9999999|schfac|numfactu|0000000|S|"
         Top             =   2730
         Width           =   830
      End
      Begin MSComctlLib.ProgressBar Pb1 
         Height          =   345
         Left            =   330
         TabIndex        =   20
         Top             =   3825
         Width           =   5370
         _ExtentX        =   9472
         _ExtentY        =   609
         _Version        =   393216
         Appearance      =   1
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
         Left            =   405
         TabIndex        =   28
         Top             =   360
         Width           =   540
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   6
         Left            =   1290
         MouseIcon       =   "frmContabFact.frx":000C
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar sección"
         Top             =   495
         Width           =   240
      End
      Begin VB.Label lblProgres 
         Height          =   285
         Index           =   1
         Left            =   330
         TabIndex        =   25
         Top             =   4620
         Width           =   5295
      End
      Begin VB.Label lblProgres 
         Height          =   285
         Index           =   0
         Left            =   360
         TabIndex        =   24
         Top             =   4260
         Width           =   5265
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Cta Banco"
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
         TabIndex        =   23
         Top             =   1320
         Width           =   735
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   8
         Left            =   1290
         MouseIcon       =   "frmContabFact.frx":015E
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar cta.banco"
         Top             =   1530
         Width           =   240
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   7
         Left            =   1290
         Picture         =   "frmContabFact.frx":02B0
         ToolTipText     =   "Buscar fecha"
         Top             =   1020
         Width           =   240
      End
      Begin VB.Label Label4 
         Caption         =   "Fecha Vto"
         ForeColor       =   &H00972E0B&
         Height          =   255
         Index           =   4
         Left            =   360
         TabIndex        =   22
         Top             =   840
         Width           =   1815
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Letra de Serie"
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
         TabIndex        =   19
         Top             =   1860
         Width           =   1005
      End
      Begin VB.Label Label4 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   1
         Left            =   2880
         TabIndex        =   18
         Top             =   2115
         Width           =   420
      End
      Begin VB.Label Label4 
         Caption         =   "Desde"
         Height          =   195
         Index           =   0
         Left            =   720
         TabIndex        =   17
         Top             =   2100
         Width           =   465
      End
      Begin VB.Label Label4 
         Caption         =   "Fecha Factura"
         ForeColor       =   &H00972E0B&
         Height          =   255
         Index           =   16
         Left            =   330
         TabIndex        =   16
         Top             =   3060
         Width           =   1815
      End
      Begin VB.Label Label4 
         Caption         =   "Desde"
         Height          =   195
         Index           =   15
         Left            =   720
         TabIndex        =   15
         Top             =   3300
         Width           =   465
      End
      Begin VB.Label Label4 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   14
         Left            =   2880
         TabIndex        =   14
         Top             =   3330
         Width           =   420
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   2
         Left            =   1290
         Picture         =   "frmContabFact.frx":033B
         ToolTipText     =   "Buscar fecha"
         Top             =   3300
         Width           =   240
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   3
         Left            =   3420
         Picture         =   "frmContabFact.frx":03C6
         ToolTipText     =   "Buscar fecha"
         Top             =   3330
         Width           =   240
      End
      Begin VB.Label Label4 
         Caption         =   "Desde"
         Height          =   195
         Index           =   13
         Left            =   720
         TabIndex        =   13
         Top             =   2700
         Width           =   465
      End
      Begin VB.Label Label4 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   12
         Left            =   2880
         TabIndex        =   12
         Top             =   2745
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
         Index           =   11
         Left            =   360
         TabIndex        =   11
         Top             =   2460
         Width           =   555
      End
   End
End
Attribute VB_Name = "frmContabFact"
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

    If Not DatosOk Then Exit Sub
    
    cadSelect = tabla & ".intconta=0 "
    
    'Seccion
    AnyadirAFormula cadFormula, tabla & ".codsecci = " & DBSet(txtCodigo(6).Text, "N")
    AnyadirAFormula cadSelect, tabla & ".codsecci = " & DBSet(txtCodigo(6).Text, "N")
    
    
    'D/H Fecha factura
    cDesde = Trim(txtCodigo(2).Text)
    cHasta = Trim(txtCodigo(3).Text)
    If Not (cDesde = "" And cHasta = "") Then
        'Cadena para seleccion Desde y Hasta
        Codigo = "{" & tabla & ".fecfactu}"
        TipCod = "F"
        If Not PonerDesdeHasta(cDesde, cHasta, "", "", "pDHfechaFac= """) Then Exit Sub
    End If
    
    'D/H letra de serie
    cDesde = Trim(txtCodigo(4).Text)
    cHasta = Trim(txtCodigo(5).Text)
    If Not (cDesde = "" And cHasta = "") Then
        'Cadena para seleccion Desde y Hasta
        Codigo = "{" & tabla & ".letraser}"
        TipCod = "T"
        If Not PonerDesdeHasta(cDesde, cHasta, "", "", "pDHColec= """) Then Exit Sub
    End If
    
    'D/H numero de factura
    cDesde = Trim(txtCodigo(0).Text)
    cHasta = Trim(txtCodigo(1).Text)
    If Not (cDesde = "" And cHasta = "") Then
        'Cadena para seleccion Desde y Hasta
        Codigo = "{" & tabla & ".numfactu}"
        TipCod = "N"
        If Not PonerDesdeHasta(cDesde, cHasta, "", "", "pDHColec= """) Then Exit Sub
    End If
    
    If Not HayRegParaInforme(tabla, cadSelect) Then Exit Sub
    
    ' abrimos la conexion de la contabilidad de la seccion correspondiente
'    If BdConta = 0 Then Exit Sub
            
    If AbrirConexionContaFac(vParamAplic.UsuarioContaFac, vParamAplic.PasswordContaFac, BdConta) Then
        Set vEmpresaFac = New CempresaFac
        If vEmpresaFac.LeerNiveles Then
        
            ContabilizarFacturas tabla, cadSelect
             'Eliminar la tabla TMP
            BorrarTMPFacturas
            'Desbloqueamos ya no estamos contabilizando facturas
            DesBloqueoManual ("VENCON") 'VENtas CONtabilizar
        
        End If
        Set vEmpresaFac = Nothing
        CerrarConexionContaFac
    End If
    
eError:
    If Err.Number <> 0 Or NumError <> 0 Then
        MuestraError Err.Number, "No se ha realizado el proceso de contabilización. Llame a soporte."
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
    
    txtCodigo(2).Text = Format(Now, "dd/mm/yyyy")
    txtCodigo(3).Text = Format(Now, "dd/mm/yyyy")
    
    
'14/02/2007 lo dejo donde estaba
'
'     '07022007   lo he quitado de contbilizar facturas y lo he puesto aqui para acotar registros
'     'comprobar que se han rellenado los dos campos de fecha
'     'sino rellenar con fechaini o fechafin del ejercicio
'     'que guardamos en vbles Orden1,Orden2
'
'     Orden1 = ""
'     Orden1 = DevuelveDesdeBDNew(cConta, "parametros", "fechaini", "", "", "", "", "", "", "", "", "", "")
'
'     Orden2 = ""
'     Orden2 = DevuelveDesdeBDNew(cConta, "parametros", "fechafin", "", "", "", "", "", "", "", "", "", "")

'--monica
'     If txtCodigo(2).Text = "" Then
'        txtCodigo(2).Text = Orden1 'fechaini del ejercicio de la conta
'     End If
'
'     If txtCodigo(3).Text = "" Then
'        txtCodigo(3).Text = Orden2 'fecha fin del ejercicio de la conta
'     End If
     

End Sub

Private Sub frmC_Selec(vFecha As Date)
 'Fecha
    txtCodigo(CByte(imgFec(2).Tag)).Text = Format(vFecha, "dd/MM/yyyy")
End Sub

Private Sub frmCtas_DatoSeleccionado(CadenaSeleccion As String)
' cta de banco
    txtCodigo(indCodigo).Text = RecuperaValor(CadenaSeleccion, 1) 'codmacta
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2) 'nommacta
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
    imgFec(2).Tag = Index 'independentment de les dates que tinga, sempre pose l'index en la 27
    If txtCodigo(Index).Text <> "" Then frmC.NovaData = txtCodigo(Index).Text

    frmC.Show vbModal
    Set frmC = Nothing
    PonerFoco txtCodigo(CByte(imgFec(2).Tag) + 2)
    ' ***************************
End Sub

Private Sub ListView1_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub imgBuscar_Click(Index As Integer)
   Select Case Index
        Case 6 ' Seccion
            AbrirFrmSeccion (Index)
        
        Case 8 ' Cta Contable de Banco
            If BdConta = 0 Then Exit Sub
            
            AbrirFrmCuentas (Index)
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
            Case 8: KEYBusqueda KeyAscii, 8 'cta banco
            Case 2: KEYFecha KeyAscii, 2 'fecha desde
            Case 3: KEYFecha KeyAscii, 3 'fecha hasta
            Case 7: KEYFecha KeyAscii, 7 'fecha de vencimiento
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
        Case 6 ' Seccion
            BdConta = 0
            txtNombre(Index).Text = PonerNombreDeCod(txtCodigo(Index), "seccion", "nomsecci", "codsecci", "N")
            If txtCodigo(Index).Text <> "" Then
                txtCodigo(Index).Text = Format(txtCodigo(Index).Text, "00")
            
                Cad = DevuelveDesdeBDNew(cPTours, "seccion", "numconta", "codsecci", txtCodigo(6).Text, "N") 'numconta
                If Cad <> "" Then BdConta = CInt(Cad)  'numero de conta
'            Else
'                MsgBox "Debe introducir un código existente en la sección.", vbExclamation
            End If
        
        Case 8 ' Cuenta de Banco
            If BdConta = 0 Then Exit Sub
            
            If AbrirConexionContaFac(vParamAplic.UsuarioContaFac, vParamAplic.PasswordContaFac, BdConta) Then
                Set vEmpresaFac = New CempresaFac
                If vEmpresaFac.LeerNiveles Then
                    txtNombre(Index).Text = PonerNombreCuenta(txtCodigo(8), 1, , BdConta, True) 'DevuelveDesdeBDNewFac("cuentas", "nommacta", "codmacta", txtCodigo(Index), "N")
                End If
                Set vEmpresaFac = Nothing
                CerrarConexionContaFac
            End If
            
        Case 2, 3, 7  'FECHAS
            If txtCodigo(Index).Text <> "" Then PonerFormatoFecha txtCodigo(Index)
'            If Not ConnContaFac Is Nothing Then
'                If (Index = 2 Or Index = 3) Then
'                    If txtCodigo(Index).Text <> "" Then
'                        'Contabilizar facturas
'                        If Not ComprobarFechasConta(Index) Then PonerFoco txtCodigo(Index)
'                    End If
'                End If
'            End If
        Case 0, 1 ' NUMERO DE FACTURA
            If txtCodigo(Index).Text <> "" Then PonerFormatoEntero txtCodigo(Index)
        
        Case 4, 5 ' LETRA DE SERIE
            If txtCodigo(Index).Text <> "" Then txtCodigo(Index).Text = UCase(txtCodigo(Index).Text)
        
        
    End Select
End Sub

Private Sub FrameCobrosVisible(visible As Boolean, ByRef h As Integer, ByRef w As Integer)
    Me.FrameCobros.visible = visible
    If visible = True Then
        Me.FrameCobros.Top = -90
        Me.FrameCobros.Left = 0
        Me.FrameCobros.Height = 6015
        Me.FrameCobros.Width = 6555
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


Private Function DatosOk() As Boolean
Dim b As Boolean
Dim Cad As String

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
        MsgBox "Debe introducir obligatoriamente una Fecha de Vencimiento.", vbExclamation
        PonerFoco txtCodigo(7)
        Exit Function
    End If
    
    If txtCodigo(8).Text = "" Then
        MsgBox "Debe introducir obligatoriamente una Cta.Banco para realizar el cobro.", vbExclamation
        PonerFoco txtCodigo(8)
        Exit Function
    Else
        If BdConta = 0 Then Exit Function
        
        If AbrirConexionContaFac(vParamAplic.UsuarioContaFac, vParamAplic.PasswordContaFac, BdConta) Then
            Set vEmpresaFac = New CempresaFac
            If vEmpresaFac.LeerNiveles Then
                txtNombre(8).Text = PonerNombreCuenta(txtCodigo(8), 1, , BdConta, True) 'DevuelveDesdeBDNewFac("cuentas", "nommacta", "codmacta", txtCodigo(Index), "N")
                If txtNombre(8).Text = "" Then
                    PonerFoco txtCodigo(8)
                    Exit Function
                End If
                Orden1 = ""
                Orden1 = DevuelveDesdeBDNewFac("parametros", "fechaini", "", "", "", "", "", "", "", "", "", "")

                Orden2 = ""
                Orden2 = DevuelveDesdeBDNewFac("parametros", "fechafin", "", "", "", "", "", "", "", "", "", "")
                'comprobar que se han rellenado los dos campos de fecha
                'sino rellenar con fechaini o fechafin del ejercicio
                'que guardamos en vbles Orden1,Orden2
                If txtCodigo(2).Text = "" Then
                    txtCodigo(2).Text = Orden1 'fechaini del ejercicio de la conta
                End If
            
                If txtCodigo(3).Text = "" Then
                    txtCodigo(3).Text = DateAdd("yyyy", 1, CDate(Orden2))  'fecha fin del ejercicio de la conta
                End If
                If Not ComprobarFechasConta(2) Then Exit Function
                If Not ComprobarFechasConta(3) Then Exit Function
                
            End If
            Set vEmpresaFac = Nothing
            CerrarConexionContaFac
        End If
    End If
    
    DatosOk = True
End Function

' copiado del ariges
Private Sub ContabilizarFacturas(cadTABLA As String, cadwhere As String)
'Contabiliza Facturas de Clientes o de Proveedores
Dim Sql As String
Dim b As Boolean
Dim tmpErrores As Boolean 'Indica si se creo correctamente la tabla de errores
Dim CCoste As String
Dim Cad As String

    Sql = "VENCON" 'contabilizar facturas de venta

    'Bloquear para que nadie mas pueda contabilizar
    DesBloqueoManual (Sql)
    If Not BloqueoManual(Sql, "1") Then
        MsgBox "No se pueden Contabilizar Facturas. Hay otro usuario contabilizando.", vbExclamation
        Screen.MousePointer = vbDefault
        Exit Sub
    End If


     'comprobar que se han rellenado los dos campos de fecha
     'sino rellenar con fechaini o fechafin del ejercicio
     'que guardamos en vbles Orden1,Orden2
     If txtCodigo(2).Text = "" Then
        txtCodigo(2).Text = Orden1 'vEmpresa.FechaIni  'fechaini del ejercicio de la conta
     End If

     If txtCodigo(3).Text = "" Then
        txtCodigo(3).Text = Orden2 'vEmpresa.FechaFin  'fecha fin del ejercicio de la conta
     End If

     'Comprobar que el intervalo de fechas D/H esta dentro del ejercicio de la
     'contabilidad par ello mirar en la BD de la Conta los parámetros
     If Not ComprobarFechasConta(3) Then Exit Sub
     
     

    'La comprobacion solo lo hago para facturas nuestras, ya que mas adelante
    'el programa hara cdate(text1(31) cuando contabilice las facturas y dara error de tipos
    If Me.txtCodigo(2).Text = "" Then
        MsgBox "Fecha inicio incorrecta", vbExclamation
        Exit Sub
    End If



    'comprobar si existen en Ariagroutil facturas anteriores al periodo solicitado
    'sin contabilizar
    If Me.txtCodigo(2).Text <> "" Then
        Sql = "SELECT COUNT(*) FROM " & cadTABLA
        Sql = Sql & " WHERE fecfactu <"
        Sql = Sql & DBSet(txtCodigo(2), "F") & " AND intconta=0 "
'        SQL = SQL & " and codsecci = " & DBSet(txtCodigo(6).Text, "N")
        If RegistrosAListar(Sql) > 0 Then
            '[Monica]11/10/2011: indico si es de esta seccion o de otra seccion
            Sql = "select count(*) from " & cadTABLA
            Sql = Sql & " WHERE fecfactu <"
            Sql = Sql & DBSet(txtCodigo(2), "F") & " AND intconta=0 and codsecci = " & DBSet(txtCodigo(6).Text, "N")
            If RegistrosAListar(Sql) > 0 Then
                Cad = "Hay Facturas anteriores sin contabilizar de esta sección." & vbCrLf
            Else
                Cad = "Hay Facturas anteriores sin contabilizar de otra sección." & vbCrLf
            End If
'            cad = "Hay Facturas anteriores sin contabilizar." & vbCrLf
            Cad = Cad & "            ¿ Desea continuar ? "
            If MsgBox(Cad, vbQuestion + vbYesNo + vbDefaultButton1) = vbNo Then
                Exit Sub
            End If
        End If
    End If
    
    'Visualizar la barra de Progreso
    Me.Pb1.visible = True
'    Me.Pb1.Top = 3350
    
    
    '==========================================================
    'REALIZAR COMPROBACIONES ANTES DE CONTABILIZAR FACTURAS
    '==========================================================
    
    Me.lblProgres(0).Caption = "Comprobaciones: "
    CargarProgres Me.Pb1, 100
        
    BorrarTMPFacturas
    'Cargar tabla TEMP con las Facturas que vamos a Trabajar
    b = CrearTMPFacturas(cadTABLA, cadwhere, True)
    If Not b Then Exit Sub
            
    BorrarTMPErrComprob
    ' nuevo
    b = CrearTMPErrComprob()
    If Not b Then Exit Sub
    
    'comprobar que todas las LETRAS SERIE existen en la contabilidad y en Ariagroutilgasol
    '-----------------------------------------------------------------------------
    IncrementarProgres Me.Pb1, 10
    Me.lblProgres(1).Caption = "Comprobando letras de serie ..."
    b = ComprobarLetraSerieFac()
    IncrementarProgres Me.Pb1, 10
    Me.Refresh
    If Not b Then
        frmMensaje.OpcionMensaje = 1
        frmMensaje.Show vbModal
        Exit Sub
    End If
    
    'comprobar que no haya Nº FACTURAS en la contabilidad para esa fecha
    'que ya existan
    '-----------------------------------------------------------------------
    Me.lblProgres(1).Caption = "Comprobando Nº Facturas en contabilidad ..."
    If vParamAplic.ContabilidadNueva Then
        Sql = "anofactu>=" & Year(txtCodigo(2).Text) & " AND anofactu<= " & Year(txtCodigo(3).Text)
        b = ComprobarNumFacturasFacContaNueva(Sql)
    Else
        Sql = "anofaccl>=" & Year(txtCodigo(2).Text) & " AND anofaccl<= " & Year(txtCodigo(3).Text)
        b = ComprobarNumFacturasFac(Sql)
    End If
    IncrementarProgres Me.Pb1, 20
    Me.Refresh
    If Not b Then
        frmMensaje.OpcionMensaje = 1
        frmMensaje.Show vbModal
        Exit Sub
    End If
    
    'comprobar que todas las CUENTAS de los distintos clientes que vamos a
    'contabilizar existen en la Conta: sclien.codmacta IN (conta.cuentas.codmacta)
    '-----------------------------------------------------------------------------
    Me.lblProgres(1).Caption = "Comprobando Cuentas Contables en contabilidad ..."
    b = ComprobarCtaContableFac(1, cadSelect)
    IncrementarProgres Me.Pb1, 20
    Me.Refresh
    If Not b Then
        frmMensaje.OpcionMensaje = 2
        frmMensaje.Show vbModal
        Exit Sub
    End If
    
    'comprobar que todas las CUENTAS de retencion de las distintas facturas que vamos a
    'contabilizar existen en la Conta: cabfact.cuereten IN (conta.cuentas.codmacta)
    '-----------------------------------------------------------------------------
    Me.lblProgres(1).Caption = "Comprobando Cuentas Contables en contabilidad ..."
    b = ComprobarCtaContableFac(8, cadSelect)
    IncrementarProgres Me.Pb1, 20
    Me.Refresh
    If Not b Then
        frmMensaje.OpcionMensaje = 2
        frmMensaje.Show vbModal
        Exit Sub
    End If
    
    
    
    'comprobar que todas las CUENTAS de venta de la familia de los articulos que vamos a
    'contabilizar existen en la Conta: sfamia.ctaventa IN (conta.cuentas.codmacta)
    '-----------------------------------------------------------------------------
    Me.lblProgres(1).Caption = "Comprobando Cuentas Ctbles Ventas en contabilidad ..."
    b = ComprobarCtaContableFac(2, cadSelect)
    IncrementarProgres Me.Pb1, 20
    Me.Refresh
    If Not b Then
        frmMensaje.OpcionMensaje = 2
        frmMensaje.Show vbModal
        Exit Sub
    End If
    
    'comprobar que todas las CUENTAS de venta de los conceptos que vamos a
    'contabilizar son de grupo de ventas: empiezan por conta.parametros.grupovtas
    '-----------------------------------------------------------------------------
    Me.lblProgres(1).Caption = "Comprobando Cuentas Ctbles Ventas en contabilidad ..."
    b = ComprobarCtaContableFac(3)
    IncrementarProgres Me.Pb1, 20
    Me.Refresh
    If Not b Then
        frmMensaje.OpcionMensaje = 2
        frmMensaje.Show vbModal
        Exit Sub
    End If
    
    'comprobar que todas la CUENTA del banco propio donde contabilizar el cobro
    'que existen en la Conta: sbanpr.codmacta IN (conta.cuentas.codmacta)
    '-----------------------------------------------------------------------------
    Me.lblProgres(1).Caption = "Comprobando Cuentas Contables del Banco en contabilidad ..."
    
    b = ComprobarCtaContableFac(4, CStr(txtCodigo(8).Text))
    IncrementarProgres Me.Pb1, 10
    Me.Refresh
    If Not b Then
        frmMensaje.OpcionMensaje = 2
        frmMensaje.Show vbModal
        Exit Sub
    End If
    
    
    'comprobar que todos las TIPO IVA de las distintas facturas que vamos a
    'contabilizar existen en la Conta: schfac.codigiv1,codigiv2,codigiv3 IN (conta.tiposiva.codigiva)
    '--------------------------------------------------------------------------
    Me.lblProgres(1).Caption = "Comprobando Tipos de IVA en contabilidad ..."
    b = ComprobarTiposIVA(txtCodigo(6))
    IncrementarProgres Me.Pb1, 10
    Me.Refresh
    If Not b Then
        frmMensaje.OpcionMensaje = 3
        frmMensaje.Show vbModal
        Exit Sub
    End If
    
    
    'comprobar que todos los CENTRO DE COSTE de las distintas facturas que vamos a
    'contabilizar existen en la Conta: codccost in conta.cabccost
    '--------------------------------------------------------------------------
    If vEmpresaFac.TieneAnalitica Then
        Me.lblProgres(1).Caption = "Comprobando Centros Coste en contabilidad ..."
        b = ComprobarCCoste()
        IncrementarProgres Me.Pb1, 10
        Me.Refresh
        If Not b Then
            frmMensaje.OpcionMensaje = 0
            frmMensaje.Show vbModal
            Exit Sub
        End If
    Else
        IncrementarProgres Me.Pb1, 10
        Me.Refresh
    End If
    
    
    '===========================================================================
    'CONTABILIZAR FACTURAS
    '===========================================================================
    Me.lblProgres(0).Caption = "Contabilizar Facturas: "
    CargarProgres Me.Pb1, 10
    Me.lblProgres(1).Caption = "Insertando Facturas en Contabilidad..."
       
    
    'Crear tabla TEMP para los posible errores de facturas
    tmpErrores = CrearTMPErrFact(cadTABLA)
    
    
    b = PasarFacturasAContab(cadTABLA, txtCodigo(7).Text, txtCodigo(8).Text, CCoste)
    
    If Not b Then
        If tmpErrores Then
            'Cargar un listview con la tabla TEMP de Errores y mostrar
            'las facturas que fallaron
            frmMensaje.OpcionMensaje = 10
            frmMensaje.Show vbModal
        Else
            MsgBox "No pueden mostrarse los errores.", vbInformation
        End If
    Else
        MsgBox "El proceso ha finalizado correctamente.", vbInformation
    End If
    
    'Eliminar tabla TEMP de Errores
    BorrarTMPErrFact
    BorrarTMPErrComprob
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

