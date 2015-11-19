VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmListFactCV 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Informes"
   ClientHeight    =   4860
   ClientLeft      =   45
   ClientTop       =   2430
   ClientWidth     =   6720
   Icon            =   "frmListFactCV.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4860
   ScaleWidth      =   6720
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
      Height          =   4740
      Left            =   0
      TabIndex        =   6
      Top             =   45
      Width           =   6690
      Begin VB.Frame Frame2 
         BorderStyle     =   0  'None
         Height          =   705
         Left            =   3870
         TabIndex        =   18
         Top             =   660
         Width           =   1950
         Begin VB.TextBox txtCodigo 
            Height          =   285
            Index           =   4
            Left            =   1290
            MaxLength       =   10
            TabIndex        =   19
            Top             =   210
            Width           =   405
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Letra Serie"
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
            Left            =   150
            TabIndex        =   20
            Top             =   255
            Width           =   780
         End
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Index           =   1
         ItemData        =   "frmListFactCV.frx":000C
         Left            =   1530
         List            =   "frmListFactCV.frx":0016
         Style           =   2  'Dropdown List
         TabIndex        =   17
         Tag             =   "Tipo Factura|N|N|0|5|factsocio|tipofact|||"
         Top             =   900
         Width           =   1665
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   3
         Left            =   1530
         MaxLength       =   10
         TabIndex        =   3
         Tag             =   "Código Postal|T|S|||clientes|codposta|||"
         Top             =   3030
         Width           =   1050
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   2
         Left            =   1530
         MaxLength       =   10
         TabIndex        =   2
         Tag             =   "Código Postal|T|S|||clientes|codposta|||"
         Top             =   2700
         Width           =   1050
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Left            =   5220
         TabIndex        =   5
         Top             =   4005
         Width           =   975
      End
      Begin VB.CommandButton cmdAceptar 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   4005
         TabIndex        =   4
         Top             =   4005
         Width           =   975
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   0
         Left            =   1530
         MaxLength       =   10
         TabIndex        =   0
         Top             =   1590
         Width           =   1230
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   1
         Left            =   1530
         MaxLength       =   10
         TabIndex        =   1
         Top             =   1965
         Width           =   1230
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   0
         Left            =   2835
         Locked          =   -1  'True
         TabIndex        =   13
         Text            =   "Text5"
         Top             =   1605
         Width           =   3360
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   1
         Left            =   2835
         Locked          =   -1  'True
         TabIndex        =   14
         Text            =   "Text5"
         Top             =   1965
         Width           =   3360
      End
      Begin VB.Label Label1 
         Caption         =   "Listado Facturas Varias / Coarval"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   345
         Index           =   0
         Left            =   315
         TabIndex        =   16
         Top             =   360
         Width           =   5160
      End
      Begin VB.Label Label4 
         Caption         =   "Tipo de Factura"
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
         Height          =   255
         Index           =   0
         Left            =   300
         TabIndex        =   15
         Top             =   900
         Width           =   1815
      End
      Begin VB.Label Label4 
         Caption         =   "Fecha"
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
         Height          =   255
         Index           =   16
         Left            =   285
         TabIndex        =   12
         Top             =   2430
         Width           =   1815
      End
      Begin VB.Label Label4 
         Caption         =   "Desde"
         Height          =   195
         Index           =   15
         Left            =   645
         TabIndex        =   11
         Top             =   2700
         Width           =   465
      End
      Begin VB.Label Label4 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   14
         Left            =   645
         TabIndex        =   10
         Top             =   3030
         Width           =   420
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   2
         Left            =   1215
         Picture         =   "frmListFactCV.frx":002C
         ToolTipText     =   "Buscar fecha"
         Top             =   2670
         Width           =   240
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   3
         Left            =   1215
         Picture         =   "frmListFactCV.frx":00B7
         ToolTipText     =   "Buscar fecha"
         Top             =   3030
         Width           =   240
      End
      Begin VB.Label Label4 
         Caption         =   "Desde"
         Height          =   195
         Index           =   13
         Left            =   645
         TabIndex        =   9
         Top             =   1590
         Width           =   465
      End
      Begin VB.Label Label4 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   12
         Left            =   645
         TabIndex        =   8
         Top             =   1965
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
         Left            =   285
         TabIndex        =   7
         Top             =   1335
         Width           =   1215
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   0
         Left            =   1245
         MouseIcon       =   "frmListFactCV.frx":0142
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar Cuenta Contable"
         Top             =   1590
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   1
         Left            =   1245
         MouseIcon       =   "frmListFactCV.frx":0294
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar Cuenta Contable"
         Top             =   1965
         Width           =   240
      End
   End
End
Attribute VB_Name = "frmListFactCV"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' +-+-+-+-+-+-+-+-+-+-+-
' +-+- Autor MANOLO +-+-
' +-+-+-+-+-+-+-+-+-+-+-

Option Explicit

Public OpcionListado As Byte
' =====FACTURAS SOCIOS====
' 0 = Listado de Retenciones

Public CadTag As String 'Cadena con el Tag del campo que se va a poner en D/H en los listados
                        'Se necesita si el tipo de codigo es texto

    
Private Conexion As Byte
'1.- Conexión a BD Ariges  2.- Conexión a BD Conta

Private HaDevueltoDatos As Boolean

'Private WithEvents frmcli As frmManClien 'Clientes
'Private WithEvents frmCol As frmManCoope 'Colectivo
Private WithEvents frmC As frmCal 'calendario fechas
Attribute frmC.VB_VarHelpID = -1
Private WithEvents frmCtas As frmCtasConta
Attribute frmCtas.VB_VarHelpID = -1

'GENERALES PARA PASARLE A CRYSTAL REPORT
Private cadFormula As String 'Cadena con la FormulaSelection para Crystal Report
Private cadParam As String 'Cadena con los parametros para Crystal Report
Private numParam As Byte 'Numero de parametros que se pasan a Crystal Report
Private cadSelect As String 'Cadena para comprobar si hay datos antes de abrir Informe
Private cadTitulo As String 'Titulo para la ventana frmImprimir
Private cadNombreRPT As String 'Nombre del informe

Dim indCodigo As Integer 'indice para txtCodigo
Dim indFrame As Single 'nº de frame en el que estamos

Dim BdConta As Integer ' numero de la contabilidad donde se hace conexion

'Se inicializan para cada Informe (tabla de BD a la que hace referencia
Dim tabla As String
Dim Codigo As String 'Código para FormulaSelection de Crystal Report
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


Private Sub Check1_Click(Index As Integer)
'    If Index = 1 Then
'        If Check1(Index).Value = 0 Then
'            Check1(0).Value = 0
'            Check1(0).Enabled = False
'        Else
'            Check1(0).Enabled = True
'        End If
'    End If
End Sub


Private Sub cmdAceptar_Click()
Dim cDesde As String, cHasta As String 'cadena codigo Desde/Hasta
Dim nDesde As String, nHasta As String 'cadena Descripcion Desde/Hasta
Dim cadTABLA As String, cOrden As String
Dim i As Byte
Dim b As Boolean
Dim Tipos As String

    InicializarVbles
    
    
    '========= PARAMETROS  =============================
    'Añadir el parametro de Empresa
    cadParam = cadParam & "|pEmpresa=""" & vEmpresa.nomEmpre & """|"
    numParam = numParam + 1
    
    'D/H Cuenta contable
    cDesde = Trim(txtCodigo(0).Text)
    cHasta = Trim(txtCodigo(1).Text)
    nDesde = txtNombre(0).Text
    nHasta = txtNombre(1).Text
    If Not (cDesde = "" And cHasta = "") Then
        'Cadena para seleccion Desde y Hasta
        Codigo = "{" & tabla & ".codmactasoc}"
        TipCod = "T"
        If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHcliente= """) Then Exit Sub
    End If
    
    'D/H Fecha
    cDesde = Trim(txtCodigo(2).Text)
    cHasta = Trim(txtCodigo(3).Text)
    If Not (cDesde = "" And cHasta = "") Then
        'Cadena para seleccion Desde y Hasta
        Codigo = "{" & tabla & ".fecfactu}"
        TipCod = "F"
        If Not PonerDesdeHasta(cDesde, cHasta, "", "", "pDHfechaFac= """) Then Exit Sub
    End If
    
    'cargamos en la formula que tipo de factura vamos a seleccionar
    Select Case Combo1(1).ListIndex
        Case 0
            Tipos = "0,"
            
            If txtCodigo(4).Text <> "" Then
                If Not AnyadirAFormula(cadFormula, "{cvfacturas.letraser} = '" & Trim(txtCodigo(4).Text) & "'") Then Exit Sub
                If Not AnyadirAFormula(cadSelect, "{cvfacturas.letraser} = '" & Trim(txtCodigo(4).Text) & "'") Then Exit Sub
            End If
        Case 1
            Tipos = "1,"
        Case 2
            Tipos = "2,"
    End Select
    
    
    If Tipos = "" Then
        MsgBox "Debe seleccionar al menos un tipo de factura.", vbExclamation
        Exit Sub
    Else
        ' quitamos la ultima coma
        Tipos = "{cvfacturas.tipofactu} in (" & Mid(Tipos, 1, Len(Tipos) - 1) & ")"
        If Not AnyadirAFormula(cadSelect, Tipos) Then Exit Sub
        Tipos = Replace(Replace(Tipos, "(", "["), ")", "]")
        If Not AnyadirAFormula(cadFormula, Tipos) Then Exit Sub
    End If
        
    
    cadTABLA = "cvfacturas"
    If HayRegParaInforme(cadTABLA, cadSelect) Then
        If CargarTemporal(cadTABLA, cadSelect) Then
            cadTitulo = "Listado de Facturas Varias / Coarval"
            
            cadParam = cadParam & "pUsu=" & vSesion.Codigo & "|"
            numParam = numParam + 1
            
            cadNombreRPT = "rManCV.rpt"
            LlamarImprimir
        End If
    End If
    
End Sub

Private Function CargarTemporal(cTabla As String, cWhere As String) As Boolean
Dim RS As ADODB.Recordset
Dim Sql As String
Dim Sql1 As String
Dim Sql2 As String

Dim cad As String
Dim HayReg As Boolean
Dim Nregs As Long
Dim NumeroConta As Integer

    On Error GoTo eCargarTemporal
    
    CargarTemporal = False


    Sql2 = "delete from tmpinformes where codusu = " & vSesion.Codigo
    conn.Execute Sql2

    cTabla = QuitarCaracterACadena(cTabla, "{")
    cTabla = QuitarCaracterACadena(cTabla, "}")
    
    If cWhere <> "" Then
        cWhere = QuitarCaracterACadena(cWhere, "{")
        cWhere = QuitarCaracterACadena(cWhere, "}")
        cWhere = QuitarCaracterACadena(cWhere, "_1")
    End If
    
    If Combo1(1).ListIndex = 0 Then
        NumeroConta = vParamAplic.NumeroContaCVV
    Else
        NumeroConta = vParamAplic.NumeroContaCV
    End If
        
    ' insertamos en la temporal con la suma de superficie a cero
    '                                       codforpa,nomforpa
    Sql = "insert into tmpinformes (codusu, codigo1, nombre1, importe1)    "
    Sql = Sql & "select " & DBSet(vSesion.Codigo, "N") & ",cvfacturas.codforpa, sforpa.nomforpa, sum(totalfac) from (" & cTabla & ") inner join conta" & NumeroConta & ".sforpa  On cvfacturas.codforpa = conta" & NumeroConta & ".sforpa.codforpa "
    Sql = Sql & " where " & cWhere
    Sql = Sql & " group by 1,2,3 "
    Sql = Sql & " order by 1,2,3 "
    
    conn.Execute Sql
    
    CargarTemporal = True
    Exit Function
    
eCargarTemporal:
    CargarTemporal = False
    MuestraError "Cargando temporal Forma de pago", Err.Description
End Function



Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub CmdCancelModelo_Click()
    Unload Me
End Sub

Private Sub Combo1_Change(Index As Integer)
    Frame2.Enabled = (Combo1(1).ListIndex = 0)
    Frame2.visible = (Combo1(1).ListIndex = 0)
    txtCodigo(4).Text = ""
End Sub

Private Sub Combo1_LostFocus(Index As Integer)
    Frame2.Enabled = (Combo1(1).ListIndex = 0)
    Frame2.visible = (Combo1(1).ListIndex = 0)
    txtCodigo(4).Text = ""
End Sub

Private Sub Form_Activate()
    If PrimeraVez Then
        PrimeraVez = False
        BdConta = 0
        
        Select Case OpcionListado
            Case 0
                Combo1(1).ListIndex = 0
                PonerFoco txtCodigo(0)
        
        End Select
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
Dim h As Integer, w As Integer
Dim List As Collection
Dim i As Integer

    PrimeraVez = True
    Limpiar Me

    'IMAGES para busqueda
    Me.imgBuscar(0).Picture = frmPpal.imgListImages16.ListImages(1).Picture
    Me.imgBuscar(1).Picture = frmPpal.imgListImages16.ListImages(1).Picture

    Me.FrameCobros.visible = False
         
         
    Select Case OpcionListado
        Case 0
            CargaCombo
                 
            FrameCobrosVisible True, h, w
            
    End Select
    
    indFrame = 5
    tabla = "cvfacturas"
            
            
    'Esto se consigue poneinedo el cancel en el opcion k corresponda
    Me.cmdCancel.Cancel = True

End Sub

Private Sub frmC_Selec(vFecha As Date)
 'Fecha
    txtCodigo(CByte(imgFec(2).Tag)).Text = Format(vFecha, "dd/MM/yyyy")
End Sub

Private Sub frmCtas_DatoSeleccionado(CadenaSeleccion As String)
    txtCodigo(indCodigo).Text = RecuperaValor(CadenaSeleccion, 1) 'codmacta
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2) 'nommacta
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
            AbrirFrmCuentas (Index)
        Case 2, 3 'cuentas de modelo 190
            AbrirFrmCuentas (Index + 5)
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
            Case 7: KEYBusqueda KeyAscii, 2 'cuenta desde
            Case 8: KEYBusqueda KeyAscii, 3 'cuenta hasta
            Case 2: KEYFecha KeyAscii, 3 'fecha desde
            Case 3: KEYFecha KeyAscii, 3 'fecha hasta
            Case 5: KEYFecha KeyAscii, 5 'fecha desde
            Case 6: KEYFecha KeyAscii, 6 'fecha hasta
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
Dim cad As String, cadTipo As String 'tipo cliente

    'Quitar espacios en blanco por los lados
    txtCodigo(Index).Text = Trim(txtCodigo(Index).Text)
    
    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub

    
    Select Case Index
        Case 0, 1 'Cuenta Cliente
            If txtCodigo(Index).Text = "" Then Exit Sub
            If Combo1(1).ListIndex = 0 Then
                txtNombre(Index).Text = PonerNombreCuenta(txtCodigo(Index), 0, , cContaCVV, False)
            Else
                txtNombre(Index).Text = PonerNombreCuenta(txtCodigo(Index), 0, , cContaCV, False)
            End If
            'DevuelveDesdeBDNew(cContaFacSoc, "cuentas", "nommacta", "codmacta", txtCodigo(Index).Text, "T")
            
        Case 2, 3 'FECHAS
            If txtCodigo(Index).Text <> "" Then PonerFormatoFecha txtCodigo(Index)
            
        Case 4 ' letra de serie
            If txtCodigo(Index).Text <> "" Then txtCodigo(Index).Text = UCase(txtCodigo(Index))
            
    End Select
End Sub

Private Sub FrameCobrosVisible(visible As Boolean, ByRef h As Integer, ByRef w As Integer)
    Me.FrameCobros.visible = visible
    If visible = True Then
        Me.FrameCobros.Top = -90
        Me.FrameCobros.Left = 0
        Me.FrameCobros.Height = 5595
        Me.FrameCobros.Width = 6810
        w = Me.FrameCobros.Width
        h = Me.FrameCobros.Height
    End If
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

Private Sub LlamarImprimir()
    With frmImprimir
        .FormulaSeleccion = cadFormula
        .OtrosParametros = cadParam
        .NumeroParametros = numParam
        .SoloImprimir = False
        .Titulo = cadTitulo
        .NombreRPT = cadNombreRPT
        .Opcion = 1
        .Facturas = False
        .EnvioEMail = False
        .ConSubInforme = True
        If Combo1(1).ListIndex = 0 Then
            .Contabilidad = cContaCVV
        Else
            .Contabilidad = cContaCV
        End If
        .Show vbModal
    End With
End Sub

Private Sub AbrirFrmCuentas(indice As Integer)
            
    indCodigo = indice
    Set frmCtas = New frmCtasConta
    If Combo1(1).ListIndex = 0 Then
        frmCtas.Conexion = cContaCVV
        frmCtas.Facturas = False
'        frmCtas.CadBusqueda = vParamAplic.RaizCtaFacSoc
        frmCtas.NumDigit = vEmpresaCVV.DigitosUltimoNivel
    Else
        frmCtas.Conexion = cContaCV
        frmCtas.Facturas = False
'        frmCtas.CadBusqueda = vParamAplic.RaizCtaFacSoc
        frmCtas.NumDigit = vEmpresaCV.DigitosUltimoNivel
    End If
    
    frmCtas.DatosADevolverBusqueda = "0|1|"
    frmCtas.CodigoActual = txtCodigo(indice).Text
    frmCtas.Show vbModal
    Set frmCtas = Nothing
    PonerFoco txtCodigo(indice)
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



Private Function GeneraFicheroModelo(tipo As Byte, pTabla As String, pWhere As String) As Boolean
Dim NFic As Integer
Dim Regs As Integer
Dim CodigoOrdenante As String
Dim RS As ADODB.Recordset
Dim Rs4 As ADODB.Recordset
Dim Aux As String
Dim Aux2 As String
Dim cad As String
Dim Pagos As Boolean
Dim Concepto As Byte
Dim b As Boolean
Dim Nregs As Long
Dim total As Variant
Dim Sql4 As String
Dim cTabla As String
Dim vWhere As String


    On Error GoTo EGen
    GeneraFicheroModelo = False
    
    cTabla = pTabla
    vWhere = pWhere
    
    cTabla = QuitarCaracterACadena(cTabla, "{")
    cTabla = QuitarCaracterACadena(cTabla, "}")
    cTabla = QuitarCaracterACadena(cTabla, "_1")
    If vWhere <> "" Then
        vWhere = QuitarCaracterACadena(vWhere, "{")
        vWhere = QuitarCaracterACadena(vWhere, "}")
        vWhere = QuitarCaracterACadena(vWhere, "_1")
    End If
    
    NFic = FreeFile
    
    Open App.path & "\modelo.txt" For Output As #NFic
    
    Select Case tipo
        Case 1 ' MODELO 190
            Aux = "select count(*), sum(factsocio.basereten), sum(factsocio.impreten) "
            Aux = Aux & " from " & cTabla
            Aux = Aux & " where " & vWhere
                
            Set RS = New ADODB.Recordset
            RS.Open Aux, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                
            'CABECERA
' [Monica] 14/01/2010 : no sé de donde he copiado que habían dos cabeceras
'            Cabecera190a NFic, CLng(DBLet(Rs.Fields(0).Value, "N"))
            Cabecera190b NFic, CLng(DBLet(RS.Fields(0).Value, "N")), CCur(DBLet(RS.Fields(1).Value, "N")), CCur(DBLet(RS.Fields(2).Value, "N"))
            
            Set RS = Nothing
            
            'Imprimimos las lineas
            Aux = "select factsocio.codmacta, sum(factsocio.basereten), sum(factsocio.impreten) "
            Aux = Aux & " from " & cTabla
            Aux = Aux & " where " & vWhere
            Aux = Aux & " group by 1 "
            Aux = Aux & " having sum(factsocio.basereten) <> 0 "
            Aux = Aux & " order by 1 "
            
            Set RS = New ADODB.Recordset
            RS.Open Aux, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            
            If RS.EOF Then
                'No hayningun registro
            Else
                b = True
                Regs = 0
                While Not RS.EOF And b
                    Regs = Regs + 1
                    
                    Sql4 = "select nifdatos, nommacta, codposta from cuentas where codmacta = " & DBLet(RS!Codmacta, "T")
                    Set Rs4 = New ADODB.Recordset
                    
                    Rs4.Open Sql4, ConnContaFacSoc, adOpenForwardOnly, adLockPessimistic, adCmdText
                    
                    If Not Rs4.EOF Then
                        Linea190 NFic, Rs4, RS
                    Else
                        b = False
                    End If
                    Set Rs4 = Nothing
                    
                    RS.MoveNext
                Wend
            End If
            RS.Close
            Set RS = Nothing
            
'        Case 2 ' MODELO 346
''            cTabla = "(" & cTabla & ") INNER JOIN variedades ON rfactsoc_variedad.codvarie = variedades.codvarie "
''            cTabla = "(" & cTabla & ") INNER JOIN productos ON variedades.codprodu = productos.codprodu "
''            cTabla = "(" & cTabla & ") INNER JOIN grupopro ON productos.codgrupo = grupopro.codgrupo "
''
''            Aux = "select rfactsoc.codsocio, grupopro.codgrupo, sum(rfactsoc_variedad.imporvar) "
''            Aux = Aux & " from " & cTabla
''            Aux = Aux & " where " & vWhere & " and grupopro.codgrupo in (4,5) " ' algarrobos y olivos
''            Aux = Aux & " group by rfactsoc.codsocio, grupopro.codgrupo "
''            Aux = Aux & "  union "
''            Aux = Aux & " select rfactsoc.codsocio, 0, sum(rfactsoc_variedad.imporvar) "
''            Aux = Aux & " from " & cTabla
''            Aux = Aux & " where " & vWhere & " and not grupopro.codgrupo in (4,5) " ' el resto
''            Aux = Aux & " group by rfactsoc.codsocio, grupopro.codgrupo "
''            Aux = Aux & " order by 1,2"
'
'            Aux = "select tmp346.codsocio, tmp346.codgrupo, sum(tmp346.importe) "
'            Aux = Aux & " from tmp346 "
'            Aux = Aux & " where " & vWhere & " and tmp346.codgrupo in (4,5) " ' algarrobos y olivos
'            Aux = Aux & " group by tmp346.codsocio, tmp346.codgrupo "
'            Aux = Aux & "  union "
'            Aux = Aux & " select tmp346.codsocio, 0, sum(tmp346.importe) "
'            Aux = Aux & " from tmp346 "
'            Aux = Aux & " where " & vWhere & " and not tmp346.codgrupo in (4,5) " ' el resto
'            Aux = Aux & " group by tmp346.codsocio, tmp346.codgrupo "
'            Aux = Aux & " order by 1,2"
'
'
'
'            Nregs = TotalRegistrosConsulta(Aux)
'
'            If Nregs <> 0 Then
'                Aux2 = "select sum(tmp346.importe) from tmp346 "
'                Aux2 = Aux2 & " where " & vWhere
'
'                total = DevuelveValor(Aux2)
'
'                Cabecera346 NFic, Nregs, CCur(total)
'
'                Set RS = New ADODB.Recordset
'                RS.Open Aux, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
'
'                If RS.EOF Then
'                    'No hayningun registro
'                Else
'                    b = True
'                    Regs = 0
'                    While Not RS.EOF And b
'                        Regs = Regs + 1
'                        Set vSocio = New CSocio
'
'                        If vSocio.LeerDatos(DBLet(RS!Codsocio, "N")) Then
'                            Linea346 NFic, vSocio, RS
'                        Else
'                            b = False
'                        End If
'
'                        Set vSocio = Nothing
'                        RS.MoveNext
'                    Wend
'                End If
'                RS.Close
'                Set RS = Nothing
'
'            End If
    End Select
    Close (NFic)
    
    If Regs > 0 Then GeneraFicheroModelo = True
    Exit Function
    
EGen:
    Set RS = Nothing
    Close (NFic)
    MuestraError Err.Number, Err.Description
End Function


Private Sub Cabecera190b(NFich As Integer, Nregs As Currency, ImpReten As Currency, BaseReten As Currency)
Dim cad As String

'TIPO DE REGISTRO 1:REGISTRO DEL RETENEDOR}
    
    cad = "1190"                                                  'p.1
    cad = cad & Format(txtCodigo(30).Text, "0000")                'p.5 año de ejercicio
    cad = cad & RellenaABlancos(vEmpresa.CifEmpresa, True, 9)        'p.9 cif empresa
    cad = cad & RellenaABlancos(vEmpresa.nomEmpre, True, 40)   'p.18 nombre de empresa
    cad = cad & "D"                                               'p.58
    cad = cad & RellenaAceros(txtCodigo(10).Text, True, 9)        'p.59 telefono
    cad = cad & RellenaABlancos(txtCodigo(9).Text, True, 40)     'p.68 persona de contacto
    cad = cad & RellenaAceros(txtCodigo(4).Text, True, 13)       'p.108 nro de justificante
    cad = cad & Space(2)                                          'p.121 ni es complementaria ni sustitutiva
    cad = cad & RellenaAceros("0", True, 13)                      'p.123 13 ceros (justificante de la complementaria o sustitutiva)
    cad = cad & Format(Nregs, "000000000")                        'p.136 nro de registros

    If BaseReten < 0 Then
        cad = cad & "N"                                           'p.145 signo de retenciones
        cad = cad & RellenaAceros(ImporteSinFormato(CStr(BaseReten * (-1) * 100)), False, 15)    'p.146
    Else
        cad = cad & " "                                           'p.145
        cad = cad & RellenaAceros(ImporteSinFormato(CStr(BaseReten * 100)), False, 15)           'p.146
    End If
              
    If ImpReten < 0 Then                                          'p.161
        cad = cad & RellenaAceros(ImporteSinFormato(CStr(ImpReten * (-1) * 100)), False, 15)
    Else
        cad = cad & RellenaAceros(ImporteSinFormato(CStr(ImpReten * 100)), False, 15)
    End If
    cad = cad & Space(322) 'p.176 a 487                    'antes:  Space(62)                             'p.176
    cad = cad & Space(3)   'p.488 a 500 firma digital      'antes:  Space(13)                                         'p.238

    Print #NFich, cad

End Sub


Private Sub Linea190(NFich As Integer, ByRef Rs4 As ADODB.Recordset, ByRef RS As ADODB.Recordset)
Dim cad As String

    cad = "2190"                                                'p.1
    cad = cad & Format(txtCodigo(30).Text, "0000")              'p.5 año ejercicio
    cad = cad & RellenaABlancos(vEmpresa.CifEmpresa, True, 9)     'p.9 cif empresa
    cad = cad & RellenaABlancos(Rs4!nifdatos, True, 9)            'p.18 nifsocio
    cad = cad & Space(9)                                        'p.27 nif del representante legal
    cad = cad & RellenaABlancos(Rs4!Nommacta, True, 40)        'p.36 nombre socio
    cad = cad & RellenaABlancos(Mid(Rs4!Codposta, 1, 2), True, 2) 'p.76 codpobla[1,2] codigo de provincia
    cad = cad & "H"                                             'p.78 clave de percepcion H=actividades agrícolas, ganaderas y forestales
    cad = cad & "01"                                            'p.79 subclave:
'                                                                       01 =  Se consignará esta subclave cuando se trate de percepciones
'                                                                        a las que resulte aplicable el tipo de retención establecido
'                                                                        con carácter general en el artículo 95.4.2º del Reglamento
'                                                                        del Impuesto.
   
'[Monica]: 14/01/2010
' antes no estaba en el if de abajo siempre era un blanco lo he cambiado según el signo.
'    cad = cad & " "                                             'p.81
    
    If DBLet(RS.Fields(1).Value, "N") < 0 Then                  'p.82 base de retencion
        cad = cad & "N"                                             'p.81
        cad = cad & RellenaAceros(ImporteSinFormato(CStr(DBLet(RS.Fields(1).Value, "N") * (-1) * 100)), False, 13)
    Else
        cad = cad & " "                                             'p.81
        cad = cad & RellenaAceros(ImporteSinFormato(CStr(DBLet(RS.Fields(1).Value, "N") * 100)), False, 13)
    End If
    
    If DBLet(RS.Fields(2).Value, "N") < 0 Then                  'p.95 importe de retencion
        cad = cad & RellenaAceros(ImporteSinFormato(CStr(DBLet(RS.Fields(2).Value, "N") * (-1) * 100)), False, 13)
    Else
        cad = cad & RellenaAceros(ImporteSinFormato(CStr(DBLet(RS.Fields(2).Value, "N") * 100)), False, 13)
    End If
    
    cad = cad & " "                                             'p.108
    cad = cad & RellenaAceros("0", True, 13)                    'p.109
    cad = cad & RellenaAceros("0", True, 13)                    'p.122
    cad = cad & RellenaAceros("0", True, 13)                    'p.135
    cad = cad & RellenaAceros("0", True, 4)                     'p.148
    cad = cad & "0"                                             'p.152
    cad = cad & RellenaAceros("0", True, 5)                     'p.153
    cad = cad & RellenaABlancos(" ", True, 9)                   'p.158
    cad = cad & String(88, "0")                                 'p.167  antes eran 84 ceros
    cad = cad & Space(246)                                      'p.255 - 500 se rellenan a blancos
    
    Print #NFich, cad
End Sub



Private Function DatosOk() As Boolean
Dim b As Boolean
Dim Sql As String
Dim Sql2 As String
Dim vClien As CSocio
' añadido
Dim Mens As String
Dim numfactu As String
Dim numser As String
Dim Fecha As Date
Dim tipoMov As String

    b = True
    Select Case OpcionListado
        Case 1
            '1 - grabacion de modelo 190
            If txtCodigo(5).Text = "" Or txtCodigo(6) = "" Then
                MsgBox "Debe introducir obligatoriamente el rango de fechas.", vbExclamation
                b = False
                PonerFoco txtCodigo(5)
            Else
                If Mid(txtCodigo(5).Text, 7, 4) = Mid(txtCodigo(6).Text, 7, 4) Then
                    If Not (Mid(txtCodigo(5).Text, 1, 5) = "01/01" And Mid(txtCodigo(6).Text, 1, 5) = "31/12") Then
                        MsgBox "El rango de fechas debe de corresponderse con un año natural. Revise.", vbExclamation
                        b = False
                        PonerFoco txtCodigo(5)
                    End If
                Else
                    MsgBox "El rango de fechas debe ser el de año natural. Revise.", vbExclamation
                    b = False
                    PonerFoco txtCodigo(5)
                End If
            End If
    End Select
    DatosOk = b

End Function

Private Sub CargaCombo()
Dim Ini As Integer
Dim Fin As Integer
Dim i As Integer

    ' *** neteje els combos, els pose valor i seleccione el valor per defecte ***
    'tipo de factura
    Combo1(1).Clear
    
    Combo1(1).AddItem "Varias"
    Combo1(1).ItemData(Combo1(1).NewIndex) = 0
    Combo1(1).AddItem "Ventas Tienda"
    Combo1(1).ItemData(Combo1(1).NewIndex) = 1
    Combo1(1).AddItem "Compras Tienda"
    Combo1(1).ItemData(Combo1(1).NewIndex) = 2

End Sub

