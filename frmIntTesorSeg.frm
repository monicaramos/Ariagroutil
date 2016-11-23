VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmIntTesorSeg 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Integración de Cobros en Tesorería"
   ClientHeight    =   5640
   ClientLeft      =   45
   ClientTop       =   2430
   ClientWidth     =   6750
   Icon            =   "frmIntTesorSeg.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5640
   ScaleWidth      =   6750
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FrameCobros 
      Height          =   5460
      Left            =   135
      TabIndex        =   7
      Top             =   120
      Width           =   6555
      Begin VB.Frame Frame2 
         Caption         =   "Datos para la contabilización"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   1710
         Left            =   90
         TabIndex        =   9
         Top             =   1890
         Width           =   6075
         Begin VB.TextBox txtCodigo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   4
            Left            =   1980
            MaxLength       =   10
            TabIndex        =   4
            Tag             =   "Cta.Contable|T|S|||sparam|ctaconta|||"
            Top             =   1260
            Width           =   1125
         End
         Begin VB.TextBox txtNombre 
            BackColor       =   &H80000018&
            Enabled         =   0   'False
            Height          =   285
            Index           =   4
            Left            =   3195
            Locked          =   -1  'True
            TabIndex        =   19
            Top             =   1260
            Width           =   2685
         End
         Begin VB.TextBox txtCodigo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   3
            Left            =   1980
            MaxLength       =   10
            TabIndex        =   3
            Tag             =   "Cta.Contable|T|S|||sparam|ctaconta|||"
            Top             =   855
            Width           =   1125
         End
         Begin VB.TextBox txtNombre 
            BackColor       =   &H80000018&
            Enabled         =   0   'False
            Height          =   285
            Index           =   3
            Left            =   3195
            Locked          =   -1  'True
            TabIndex        =   14
            Top             =   855
            Width           =   2685
         End
         Begin VB.TextBox txtCodigo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   2
            Left            =   1980
            MaxLength       =   10
            TabIndex        =   2
            Tag             =   "Código Postal|T|S|||clientes|codposta|||"
            Top             =   405
            Width           =   1080
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   4
            Left            =   1710
            ToolTipText     =   "Buscar Concepto"
            Top             =   1260
            Width           =   240
         End
         Begin VB.Label Label1 
            Caption         =   "Cta.Banco Prevista"
            ForeColor       =   &H00972E0B&
            Height          =   195
            Index           =   24
            Left            =   180
            TabIndex        =   20
            Top             =   1305
            Width           =   1395
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   3
            Left            =   1710
            ToolTipText     =   "Buscar Concepto"
            Top             =   855
            Width           =   240
         End
         Begin VB.Label Label1 
            Caption         =   "Forma de Pago"
            ForeColor       =   &H00972E0B&
            Height          =   195
            Index           =   0
            Left            =   180
            TabIndex        =   15
            Top             =   900
            Width           =   1395
         End
         Begin VB.Label Label4 
            Caption         =   "Fecha Vencimiento"
            ForeColor       =   &H00972E0B&
            Height          =   255
            Index           =   0
            Left            =   180
            TabIndex        =   13
            Top             =   450
            Width           =   1425
         End
         Begin VB.Image imgFec 
            Height          =   240
            Index           =   2
            Left            =   1710
            Picture         =   "frmIntTesorSeg.frx":000C
            ToolTipText     =   "Buscar fecha"
            Top             =   405
            Width           =   240
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Datos para Selección"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   1590
         Left            =   90
         TabIndex        =   8
         Top             =   225
         Width           =   6090
         Begin VB.TextBox txtCodigo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   0
            Left            =   2025
            MaxLength       =   10
            TabIndex        =   0
            Top             =   690
            Width           =   1050
         End
         Begin VB.TextBox txtCodigo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   1
            Left            =   2025
            MaxLength       =   10
            TabIndex        =   1
            Top             =   1050
            Width           =   1050
         End
         Begin VB.Image imgFec 
            Height          =   240
            Index           =   1
            Left            =   1755
            Picture         =   "frmIntTesorSeg.frx":0097
            ToolTipText     =   "Buscar fecha"
            Top             =   1035
            Width           =   240
         End
         Begin VB.Image imgFec 
            Height          =   240
            Index           =   0
            Left            =   1755
            Picture         =   "frmIntTesorSeg.frx":0122
            ToolTipText     =   "Buscar fecha"
            Top             =   690
            Width           =   240
         End
         Begin VB.Label Label4 
            Caption         =   "Hasta"
            Height          =   195
            Index           =   14
            Left            =   1170
            TabIndex        =   18
            Top             =   1050
            Width           =   420
         End
         Begin VB.Label Label4 
            Caption         =   "Desde"
            Height          =   195
            Index           =   15
            Left            =   1170
            TabIndex        =   17
            Top             =   690
            Width           =   465
         End
         Begin VB.Label Label4 
            Caption         =   "Fecha Póliza"
            ForeColor       =   &H00972E0B&
            Height          =   255
            Index           =   1
            Left            =   225
            TabIndex        =   16
            Top             =   450
            Width           =   1815
         End
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Left            =   4995
         TabIndex        =   6
         Top             =   4815
         Width           =   975
      End
      Begin VB.CommandButton CmdAceptar 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   3810
         TabIndex        =   5
         Top             =   4815
         Width           =   975
      End
      Begin MSComctlLib.ProgressBar Pb1 
         Height          =   345
         Left            =   180
         TabIndex        =   10
         Top             =   3690
         Width           =   5955
         _ExtentX        =   10504
         _ExtentY        =   609
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.Label lblProgres 
         Height          =   285
         Index           =   0
         Left            =   495
         TabIndex        =   12
         Top             =   4095
         Width           =   5265
      End
      Begin VB.Label lblProgres 
         Height          =   285
         Index           =   1
         Left            =   495
         TabIndex        =   11
         Top             =   4410
         Width           =   5295
      End
   End
   Begin MSComDlg.CommonDialog cd1 
      Left            =   7680
      Top             =   5160
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "frmIntTesorSeg"
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

Private WithEvents frmC As frmCal 'calendario fechas
Attribute frmC.VB_VarHelpID = -1
Private WithEvents frmFPa As frmForpaConta 'formas de pago de la contabilidad
Attribute frmFPa.VB_VarHelpID = -1
Private WithEvents frmCtas As frmCtasConta 'cuentas contables de contabilidad
Attribute frmCtas.VB_VarHelpID = -1

'GENERALES PARA PASARLE A CRYSTAL REPORT
Private cadFormula As String 'Cadena con la FormulaSelection para Crystal Report
Private cadParam As String 'Cadena con los parametros para Crystal Report
Private numParam As Byte 'Numero de parametros que se pasan a Crystal Report
Private cadSelect As String 'Cadena para comprobar si hay datos antes de abrir Informe
Private cadTitulo As String 'Titulo para la ventana frmImprimir
Private cadNombreRPT As String 'Nombre del informe

Dim indCodigo As Integer 'indice para txtCodigo
 
'Se inicializan para cada Informe (tabla de BD a la que hace referencia
Dim Codigo As String 'Código para FormulaSelection de Crystal Report
Dim TipCod As String

Dim PrimeraVez As Boolean

Private Sub KEYpress(KeyAscii As Integer)
    If KeyAscii = 13 Then 'ENTER
        KeyAscii = 0
        SendKeys "{tab}"
    ElseIf KeyAscii = 27 Then Unload Me  'ESC
    End If
End Sub

Private Sub cmdAceptar_Click()
Dim SQL As String
Dim i As Byte
Dim cadwhere As String

    If Not DatosOk Then Exit Sub
             
    SQL = "SELECT count(*)" & _
          " FROM segpoliza " & _
          "WHERE "
          
    cadwhere = "inttesor = 0"
          
    If txtCodigo(0).Text <> "" Then cadwhere = cadwhere & " and fechaenv >= " & DBSet(txtCodigo(0).Text, "F")
    If txtCodigo(1).Text <> "" Then cadwhere = cadwhere & " and fechaenv <= " & DBSet(txtCodigo(1).Text, "F")
             
    SQL = SQL & cadwhere
             
    If RegistrosAListar(SQL) = 0 Then
        MsgBox "No existen datos a contabilizar entre esas fechas.", vbExclamation
        Exit Sub
    End If
    
    ContabilizarCobros (cadwhere)
    BorrarTMPErrComprob
    DesBloqueoManual ("CONTES") 'CONtabilizacion a TESoreria
    
    Pb1.visible = False
    lblProgres(0).Caption = ""
    lblProgres(1).Caption = ""
    cmdCancel_Click
    
eError:
    If Err.Number <> 0 Then
        MuestraError Err.Number, "No se ha realizado el proceso de contabilización a tesoreria. Llame a soporte."
    End If
    
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
    If PrimeraVez Then
        PrimeraVez = False
        PonerFoco txtCodigo(0)
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
Dim h As Integer, w As Integer
Dim List As Collection

    PrimeraVez = True
    Limpiar Me

    'IMAGES para busqueda
     Me.imgBuscar(3).Picture = frmPpal.imgListImages16.ListImages(1).Picture
     Me.imgBuscar(4).Picture = frmPpal.imgListImages16.ListImages(1).Picture
     txtCodigo(2).Text = Format(Now, "dd/mm/yyyy") ' fecha de vencimiento

    '###Descomentar
'    CommitConexion
         
         
    FrameCobrosVisible True, h, w
    Pb1.visible = False
    
    'Esto se consigue poneinedo el cancel en el opcion k corresponda
    Me.cmdCancel.Cancel = True
'   Me.Width = w + 70
'   Me.Height = h + 350
End Sub

Private Sub frmC_Selec(vFecha As Date)
 'Fecha
    txtCodigo(CByte(imgFec(0).Tag)).Text = Format(vFecha, "dd/MM/yyyy")
End Sub

Private Sub frmCtas_DatoSeleccionado(CadenaSeleccion As String)
'Form de Consulta de Cuentas contables
    txtCodigo(indCodigo).Text = RecuperaValor(CadenaSeleccion, 1)
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmFpa_DatoSeleccionado(CadenaSeleccion As String)
'Form de Consulta de formas de pago de contabilidad
    txtCodigo(indCodigo).Text = RecuperaValor(CadenaSeleccion, 1)
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
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
    imgFec(0).Tag = Index 'independentment de les dates que tinga, sempre pose l'index en la 27
    If txtCodigo(Index).Text <> "" Then frmC.NovaData = txtCodigo(Index).Text

    frmC.Show vbModal
    Set frmC = Nothing
    PonerFoco txtCodigo(CByte(imgFec(0).Tag) + 1)
    ' ***************************
End Sub

Private Sub imgBuscar_Click(Index As Integer)
   Select Case Index
        Case 3 ' forma de pago de la tesoreria
            AbrirFrmForpaConta (Index)
        Case 4 'cuenta contable
            AbrirFrmCuentas (Index)
        
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
'14/02/2007
'    KEYpress KeyAscii
' ahora
    If KeyAscii = teclaBuscar Then
        Select Case Index
            Case 3: KEYBusqueda KeyAscii, 3 'forma de pago
            Case 0: KEYFecha KeyAscii, 0 'fecha desde
            Case 1: KEYFecha KeyAscii, 1 'fecha hasta
            Case 2: KEYFecha KeyAscii, 2 'fecha de vencimiento
            Case 4: KEYBusqueda KeyAscii, 4 'cuenta banco
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
        Case 3 ' FORMA DE PAGO DE LA CONTABILIDAD
            If vParamAplic.ContabilidadNueva Then
                If txtCodigo(Index).Text <> "" Then txtNombre(Index).Text = DevuelveDesdeBDNew(cContaSeg, "formapago", "nomforpa", "codforpa", txtCodigo(3).Text, "N")
            Else
                If txtCodigo(Index).Text <> "" Then txtNombre(Index).Text = DevuelveDesdeBDNew(cContaSeg, "sforpa", "nomforpa", "codforpa", txtCodigo(3).Text, "N")
            End If
            If txtNombre(Index).Text = "" Then
                MsgBox "Forma de Pago  no existe en la contabilidad. Reintroduzca.", vbExclamation
            End If
            
        Case 0, 1, 2 'FECHAS
            If txtCodigo(Index).Text <> "" Then PonerFormatoFecha txtCodigo(Index)
            
        Case 4 ' CUENTA CONTABLE
            If txtCodigo(Index).Text <> "" Then txtNombre(Index).Text = PonerNombreCuenta(txtCodigo(Index), 2, , cContaSeg)
            If txtNombre(Index).Text = "" Then
                MsgBox "Número de Cuenta contable no existe en la contabilidad. Reintroduzca.", vbExclamation
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

Private Sub AbrirFrmForpaConta(indice As Integer)
    indCodigo = indice
    Set frmFPa = New frmForpaConta
    frmFPa.DatosADevolverBusqueda = "0|1|"
    frmFPa.CodigoActual = txtCodigo(indCodigo)
    frmFPa.Conexion = cContaSeg
    frmFPa.Show vbModal
    Set frmFPa = Nothing
End Sub
 
Private Sub AbrirFrmCuentas(indice As Integer)
    indCodigo = indice
    Set frmCtas = New frmCtasConta
    frmCtas.DatosADevolverBusqueda = "0|1|"
    frmCtas.CodigoActual = txtCodigo(indCodigo)
    frmCtas.Conexion = cContaSeg
    frmCtas.Show vbModal
    Set frmCtas = Nothing
End Sub

Private Sub ContabilizarCobros(cadwhere As String)
'Contabiliza Facturas de Clientes o de Proveedores
Dim SQL As String
Dim b As Boolean
Dim tmpErrores As Boolean 'Indica si se creo correctamente la tabla de errores
Dim CCoste As String
Dim cadTABLA As String

    SQL = "CONTES" 'contabilizar tesoreria

    'Bloquear para que nadie mas pueda contabilizar
    DesBloqueoManual (SQL)
    If Not BloqueoManual(SQL, "1") Then
        MsgBox "No se pueden Contabilizar Cobros. Hay otro usuario contabilizándolo.", vbExclamation
        Screen.MousePointer = vbDefault
        Exit Sub
    End If

    
    'Visualizar la barra de Progreso
    Me.Pb1.visible = True
'    Me.Pb1.Top = 3350
    
    
    '==========================================================
    'REALIZAR COMPROBACIONES ANTES DE CONTABILIZAR FACTURAS
    '==========================================================
    
    Me.lblProgres(0).Caption = "Comprobaciones: "
    CargarProgres Me.Pb1, 100
        
    BorrarTMPErrComprob
    ' nuevo
    b = CrearTMPErrComprob()
    If Not b Then Exit Sub
    
    
    'comprobar que todas las CUENTAS de codigos avnic existen
    'en la Conta: savnic.codmacta IN (conta.cuentas.codmacta)
    '-----------------------------------------------------------------------------
    Me.lblProgres(1).Caption = "Comprobando Cuenta Ctble de Pólizas ..."
    b = ComprobarCtaContable(cadTABLA, 2, cadwhere, cContaSeg)
    IncrementarProgres Me.Pb1, 100
    Me.Refresh
    If Not b Then
        frmMensaje.OpcionMensaje = 2
        frmMensaje.Show vbModal
        Exit Sub
    End If

    
    '===========================================================================
    'CONTABILIZAR CIERRE
    '===========================================================================
    Me.lblProgres(0).Caption = "Contabilizar a Tesorería: "
    CargarProgres Me.Pb1, 10
    Me.lblProgres(1).Caption = "Insertando Registro en Tesorería..."
    
    
    b = PasarCalculoAContab(cadwhere)
    
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
    
End Sub

Private Function DatosOk() As Boolean
Dim b As Boolean
Dim Orden1 As String
Dim Orden2 As String
Dim FFin As Date

   b = True


'    ' comprobamos que la contabilizacion se encuentre en los ejercicios contables
'     Orden1 = ""
'     Orden1 = DevuelveDesdeBDNew(cContaSeg, "parametros", "fechaini", "", "", "", "", "", "", "", "", "", "")
'
'     Orden2 = ""
'     Orden2 = DevuelveDesdeBDNew(cContaSeg, "parametros", "fechafin", "", "", "", "", "", "", "", "", "", "")
'     FFinSeg = CDate(Orden2)
'     If Not (CDate(Orden1) <= CDate(txtCodigo(2).Text) And CDate(txtCodigo(2).Text) < CDate(Day(FIniSeg) & "/" & Month(FIniSeg) & "/" & Year(FIniSeg) + 2)) Then
'        MsgBox "La Fecha de la contabilización no es del ejercicio actual ni del siguiente. Reintroduzca.", vbExclamation
'        b = False
'        PonerFoco txtCodigo(2)
'     End If
    
   If txtCodigo(2).Text = "" And b Then
        MsgBox "Introduzca la Fecha de Vencimiento a contabilizar.", vbExclamation
        b = False
        PonerFoco txtCodigo(2)
   End If
    
   If txtCodigo(3).Text = "" And b Then
        MsgBox "Introduzca la Forma de Pago para contabilizar.", vbExclamation
        b = False
        PonerFoco txtCodigo(3)
   End If
   
   If txtCodigo(4).Text = "" And b Then
        MsgBox "Introduzca la Cta.Contable de Banco para contabilizar.", vbExclamation
        b = False
        PonerFoco txtCodigo(4)
   End If
   
   DatosOk = b
   
End Function

Private Function PasarCalculoAContab(cadwhere As String) As Boolean
Dim SQL As String
Dim Rs As ADODB.Recordset
Dim b As Boolean
Dim i As Integer
Dim NumLinea As Integer
Dim numdocum As String
Dim ampliacion As String
Dim ampliaciond As String
Dim ampliacionh As String
Dim ImporteD As Currency
Dim ImporteH As Currency
Dim Diferencia As Currency
Dim Obs As String
Dim cadMen As String
Dim Cad As String
Dim CtaDifer As String
Dim Codmacta As String

    On Error GoTo EPasarCal

    PasarCalculoAContab = False
    
    'Total de lineas de asiento a Insertar en la contabilidad
    SQL = "SELECT count(*)" & _
          " FROM segpoliza " & _
          "WHERE " & cadwhere
             
    NumLinea = TotalRegistros(SQL)
    
    If NumLinea = 0 Then Exit Function
    
    
    If NumLinea > 0 Then
        NumLinea = NumLinea
        
        CargarProgres Me.Pb1, NumLinea
        
        ConnConta.BeginTrans
        conn.BeginTrans
        
        Obs = "Contabilización de Cobro de Pólizas de fecha " & Format(txtCodigo(0).Text, "dd/mm/yyyy")

        SQL = "select * from segpoliza where " & cadwhere
        Set Rs = New ADODB.Recordset
        Rs.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText


        b = True
        i = 1
        While Not Rs.EOF And b
                IncrementarProgres Me.Pb1, 1
                Me.lblProgres(1).Caption = "Insertando registro en Tesorería...   (" & i & " de " & NumLinea & ")"
                Me.Refresh
                
                i = i + 1
                b = InsertarEnTesoreriaNew2(Rs, txtCodigo(2).Text, txtCodigo(3).Text, txtCodigo(4).Text, cadMen)
                cadMen = "Insertando en Tesoreria: "
               
                Rs.MoveNext
        Wend
        Rs.Close
            
' de momento comentado para hacer pruebas
        If b Then
            'Poner intconta=1 en ariagroutil.movim
            b = ActualizarCobros(cadwhere, cadMen)
            cadMen = "Actualizando Movimientos: " & cadMen
        End If
            
   End If
   
EPasarCal:
    If Err.Number <> 0 Then
        b = False
        MuestraError Err.Number, "Integrando Asiento a Contabilidad", Err.Description
    End If
    If b Then
        ConnConta.CommitTrans
        conn.CommitTrans
        PasarCalculoAContab = True
    Else
        ConnConta.RollbackTrans
        conn.RollbackTrans
        PasarCalculoAContab = False
    End If
End Function


Private Function ActualizarCobros(cadwhere As String, caderr As String) As Boolean
'Poner el movimiento como contabilizada
Dim SQL As String

    On Error GoTo EActualizar
    
    SQL = "UPDATE segpoliza SET inttesor=1 "
    SQL = SQL & " WHERE " & cadwhere

    conn.Execute SQL
    
EActualizar:
    If Err.Number <> 0 Then
        ActualizarCobros = False
        caderr = Err.Description
    Else
        ActualizarCobros = True
    End If
End Function



