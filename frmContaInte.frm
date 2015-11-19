VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmContaInte 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Integración Contable de Intereses"
   ClientHeight    =   5700
   ClientLeft      =   45
   ClientTop       =   2430
   ClientWidth     =   6825
   Icon            =   "frmContaInte.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5700
   ScaleWidth      =   6825
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
      Height          =   5460
      Left            =   150
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
         Height          =   2115
         Left            =   90
         TabIndex        =   10
         Top             =   1395
         Width           =   6075
         Begin VB.TextBox txtCodigo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   3
            Left            =   1980
            MaxLength       =   10
            TabIndex        =   4
            Tag             =   "Cta.Contable|T|S|||sparam|ctaconta|||"
            Top             =   1575
            Width           =   1125
         End
         Begin VB.TextBox txtNombre 
            BackColor       =   &H80000018&
            Enabled         =   0   'False
            Height          =   285
            Index           =   3
            Left            =   3195
            Locked          =   -1  'True
            TabIndex        =   18
            Top             =   1575
            Width           =   2685
         End
         Begin VB.TextBox txtCodigo 
            Height          =   285
            Index           =   2
            Left            =   1980
            MaxLength       =   30
            TabIndex        =   3
            Top             =   1170
            Width           =   3870
         End
         Begin VB.TextBox txtCodigo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   1
            Left            =   1980
            MaxLength       =   10
            TabIndex        =   2
            Tag             =   "Código Postal|T|S|||clientes|codposta|||"
            Top             =   810
            Width           =   1080
         End
         Begin VB.TextBox txtNombre 
            BackColor       =   &H80000018&
            Enabled         =   0   'False
            Height          =   285
            Index           =   4
            Left            =   3195
            Locked          =   -1  'True
            TabIndex        =   11
            Top             =   450
            Width           =   2685
         End
         Begin VB.TextBox txtCodigo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   4
            Left            =   1980
            MaxLength       =   10
            TabIndex        =   1
            Tag             =   "Cta.Contable|T|S|||sparam|ctaconta|||"
            Top             =   450
            Width           =   1125
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   3
            Left            =   1710
            ToolTipText     =   "Buscar Concepto"
            Top             =   1575
            Width           =   240
         End
         Begin VB.Label Label1 
            Caption         =   "Forma de Pago"
            ForeColor       =   &H00972E0B&
            Height          =   195
            Index           =   0
            Left            =   180
            TabIndex        =   19
            Top             =   1620
            Width           =   1395
         End
         Begin VB.Label Label4 
            Caption         =   "Concepto "
            ForeColor       =   &H00972E0B&
            Height          =   255
            Index           =   1
            Left            =   180
            TabIndex        =   17
            Top             =   1215
            Width           =   1050
         End
         Begin VB.Label Label4 
            Caption         =   "Fecha Vencimiento"
            ForeColor       =   &H00972E0B&
            Height          =   255
            Index           =   0
            Left            =   180
            TabIndex        =   16
            Top             =   855
            Width           =   1425
         End
         Begin VB.Image imgFec 
            Height          =   240
            Index           =   1
            Left            =   1710
            Picture         =   "frmContaInte.frx":000C
            ToolTipText     =   "Buscar fecha"
            Top             =   810
            Width           =   240
         End
         Begin VB.Label Label1 
            Caption         =   "Cta.Banco Prevista"
            ForeColor       =   &H00972E0B&
            Height          =   195
            Index           =   24
            Left            =   180
            TabIndex        =   12
            Top             =   495
            Width           =   1395
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   4
            Left            =   1710
            ToolTipText     =   "Buscar Concepto"
            Top             =   450
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
         Height          =   1095
         Left            =   90
         TabIndex        =   8
         Top             =   225
         Width           =   6090
         Begin VB.TextBox txtCodigo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   0
            Left            =   1980
            MaxLength       =   10
            TabIndex        =   0
            Tag             =   "Código Postal|T|S|||clientes|codposta|||"
            Top             =   570
            Width           =   1080
         End
         Begin VB.Image imgFec 
            Height          =   240
            Index           =   0
            Left            =   1680
            Picture         =   "frmContaInte.frx":0097
            ToolTipText     =   "Buscar fecha"
            Top             =   570
            Width           =   240
         End
         Begin VB.Label Label4 
            Caption         =   "Fecha"
            ForeColor       =   &H00972E0B&
            Height          =   255
            Index           =   16
            Left            =   300
            TabIndex        =   9
            Top             =   570
            Width           =   1425
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
      Begin VB.CommandButton cmdAceptar 
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
         TabIndex        =   13
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
         TabIndex        =   15
         Top             =   4095
         Width           =   5265
      End
      Begin VB.Label lblProgres 
         Height          =   285
         Index           =   1
         Left            =   495
         TabIndex        =   14
         Top             =   4410
         Width           =   5295
      End
   End
End
Attribute VB_Name = "frmContaInte"
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
Private WithEvents frmCtas As frmCtasConta 'cuentas de contabilidad
Attribute frmCtas.VB_VarHelpID = -1
Private WithEvents frmFpa As frmForpaConta 'formas de pago de la contabilidad
Attribute frmFpa.VB_VarHelpID = -1

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
          " FROM movim " & _
          "WHERE fechamov = " & DBSet(txtCodigo(0).Text, "F") & " and " & _
               " intconta = 0"
             
    If RegistrosAListar(SQL) = 0 Then
        MsgBox "No existen datos a contabilizar a esa fecha.", vbExclamation
        Exit Sub
    End If
    
    cadwhere = " fechamov = " & DBSet(txtCodigo(0).Text, "F") & " and " & _
               " intconta = 0"
               
    
    ContabilizarIntereses (cadwhere)
     'Eliminar la tabla TMP
    BorrarTMPErrComprob

    DesBloqueoManual ("CONINT") 'CONtabilizacion de CALculo
    
    
    
eError:
    If Err.Number <> 0 Then
        MuestraError Err.Number, "No se ha realizado el proceso de contabilización de cierre de turno. Llame a soporte."
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
     Me.imgBuscar(4).Picture = frmPpal.imgListImages16.ListImages(1).Picture
     Me.imgBuscar(3).Picture = frmPpal.imgListImages16.ListImages(1).Picture
     txtCodigo(0).Text = Format(Now, "dd/mm/yyyy") ' fecha de movimiento
     txtCodigo(1).Text = Format(Now, "dd/mm/yyyy") ' fecha de vencimiento

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

Private Sub frmTDia_DatoSeleccionado(CadenaSeleccion As String)
'Form de Consulta de Clientes
    txtCodigo(indCodigo).Text = Format(RecuperaValor(CadenaSeleccion, 1), "00")
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmCtas_DatoSeleccionado(CadenaSeleccion As String)
'Form de Consulta de Cuentas contables
    txtCodigo(indCodigo).Text = RecuperaValor(CadenaSeleccion, 1)
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmFPa_DatoSeleccionado(CadenaSeleccion As String)
'Form de Consulta de formas de pago de contabilidad
    txtCodigo(indCodigo).Text = RecuperaValor(CadenaSeleccion, 1)
    txtCodigo(indCodigo).Text = Format(txtCodigo(indCodigo).Text, "000")
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

Private Sub ListView1_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
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
            Case 4: KEYBusqueda KeyAscii, 4 'concepto al debe
            Case 0: KEYFecha KeyAscii, 0 'fecha de movimiento
            Case 1: KEYFecha KeyAscii, 1 'fecha de vencimiento
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
        Case 3 ' FORMA DE PAGO DE LA CONTABILIDAD
            If txtCodigo(Index).Text <> "" Then txtNombre(Index).Text = DevuelveDesdeBDNew(cConta, "sforpa", "nomforpa", "codforpa", txtCodigo(3).Text, "N")
            If txtNombre(Index).Text = "" Then
                MsgBox "Forma de Pago  no existe en la contabilidad. Reintroduzca.", vbExclamation
            End If
            
        Case 4 ' CUENTA CONTABLE
            If txtCodigo(Index).Text <> "" Then txtNombre(Index).Text = PonerNombreCuenta(txtCodigo(Index), 2, , 2)
            If txtNombre(Index).Text = "" Then
                MsgBox "Número de Cuenta contable no existe en la contabilidad. Reintroduzca.", vbExclamation
            End If

        Case 0, 1 'FECHAS
            If txtCodigo(Index).Text <> "" Then PonerFormatoFecha txtCodigo(Index)
            If Index = 0 Then
                txtCodigo(1).Text = txtCodigo(0).Text
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

Private Sub AbrirFrmCuentas(indice As Integer)
    indCodigo = indice
    Set frmCtas = New frmCtasConta
    frmCtas.DatosADevolverBusqueda = "0|1|"
    frmCtas.CodigoActual = txtCodigo(indCodigo)
    frmCtas.Conexion = cConta
    frmCtas.Show vbModal
    Set frmCtas = Nothing
End Sub

Private Sub AbrirFrmForpaConta(indice As Integer)
    indCodigo = indice
    Set frmFpa = New frmForpaConta
    frmFpa.DatosADevolverBusqueda = "0|1|"
    frmFpa.CodigoActual = txtCodigo(indCodigo)
    frmFpa.Conexion = cConta
    frmFpa.Show vbModal
    Set frmFpa = Nothing
End Sub
 

Private Sub ContabilizarIntereses(cadwhere As String)
'Contabiliza Facturas de Clientes o de Proveedores
Dim SQL As String
Dim b As Boolean
Dim tmpErrores As Boolean 'Indica si se creo correctamente la tabla de errores
Dim CCoste As String
Dim cadTABLA As String

    SQL = "CONINT" 'contabilizar CALCULO DE INTERESES

    'Bloquear para que nadie mas pueda contabilizar
    DesBloqueoManual (SQL)
    If Not BloqueoManual(SQL, "1") Then
        MsgBox "No se pueden Contabilizar Cálculo de Intereses. Hay otro usuario contabilizándolo.", vbExclamation
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
        
    ' nuevo
    b = CrearTMPErrComprob()
    If Not b Then Exit Sub
    
    
    'comprobar que todas las CUENTAS de codigos avnic existen
    'en la Conta: savnic.codmacta IN (conta.cuentas.codmacta)
    '-----------------------------------------------------------------------------
    Me.lblProgres(1).Caption = "Comprobando Cuenta Ctble Retención ..."
    b = ComprobarCtaContable(cadTABLA, 1, "movim.fechamov = " & DBSet(txtCodigo(0).Text, "F") & " and intconta = 0", cConta)
    IncrementarProgres Me.Pb1, 33
    Me.Refresh
    If Not b Then
        frmMensaje.OpcionMensaje = 2
        frmMensaje.Show vbModal
        Exit Sub
    End If

    
    'comprobar que todas las CUENTAS de gasto existen
    'en la Conta: sparam.ctagasto IN (conta.cuentas.codmacta)
    '-----------------------------------------------------------------------------
    Me.lblProgres(1).Caption = "Comprobando Cuenta Ctble Gasto ..."
    b = ComprobarCtaContable(cadTABLA, 6, , cConta)
    IncrementarProgres Me.Pb1, 33
    Me.Refresh
    If Not b Then
        frmMensaje.OpcionMensaje = 2
        frmMensaje.Show vbModal
        Exit Sub
    End If
    'comprobar que todas las CUENTAS de retencion existen
    'en la Conta: sparam.ctareten IN (conta.cuentas.codmacta)
    '-----------------------------------------------------------------------------
    Me.lblProgres(1).Caption = "Comprobando Cuenta Ctble Retención ..."
    b = ComprobarCtaContable(cadTABLA, 7, , cConta)
    IncrementarProgres Me.Pb1, 33
    Me.Refresh
    If Not b Then
        frmMensaje.OpcionMensaje = 2
        frmMensaje.Show vbModal
        Exit Sub
    End If
     
     
     
    
    '===========================================================================
    'CONTABILIZAR CIERRE
    '===========================================================================
    Me.lblProgres(0).Caption = "Contabilizar Cierre: "
    CargarProgres Me.Pb1, 10
    Me.lblProgres(1).Caption = "Insertando Asiento en Contabilidad..."
    
    
    cadwhere = "fechamov = " & DBSet(txtCodigo(0).Text, "F") & " and intconta = 0"
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

   

   If txtCodigo(0).Text = "" And b Then
        MsgBox "Introduzca la Fecha de movimiento a contabilizar.", vbExclamation
        b = False
        PonerFoco txtCodigo(0)
   Else
        ' comprobamos que la contabilizacion se encuentre en los ejercicios contables
         Orden1 = ""
         Orden1 = DevuelveDesdeBDNew(cConta, "parametros", "fechaini", "", "", "", "", "", "", "", "", "", "")
    
         Orden2 = ""
         Orden2 = DevuelveDesdeBDNew(cConta, "parametros", "fechafin", "", "", "", "", "", "", "", "", "", "")
         FFin = CDate(Orden2)
         If Not (CDate(Orden1) <= CDate(txtCodigo(0).Text) And CDate(txtCodigo(0).Text) < CDate(Day(FIni) & "/" & Month(FIni) & "/" & Year(FIni) + 2)) Then
            MsgBox "La Fecha de la contabilización no es del ejercicio actual ni del siguiente. Reintroduzca.", vbExclamation
            b = False
            PonerFoco txtCodigo(0)
         End If
   End If
    
   If txtCodigo(1).Text = "" And b Then
        MsgBox "Introduzca la Fecha de Vencimiento a contabilizar.", vbExclamation
        b = False
        PonerFoco txtCodigo(1)
   End If
    
   If txtCodigo(3).Text = "" And b Then
        MsgBox "Introduzca la Forma de Pago para contabilizar.", vbExclamation
        b = False
        PonerFoco txtCodigo(3)
   End If
   
   DatosOk = b
   
End Function

Private Function PasarCalculoAContab(cadwhere As String) As Boolean
Dim SQL As String
Dim RS As ADODB.Recordset
Dim b As Boolean
Dim i As Integer
Dim NumLinea As Integer
Dim Mc As CContadorContab
Dim numdocum As String
Dim ampliacion As String
Dim ampliaciond As String
Dim ampliacionh As String
Dim ImporteD As Currency
Dim ImporteH As Currency
Dim Diferencia As Currency
Dim Obs As String
Dim cadMen As String
Dim cad As String
Dim CtaDifer As String
Dim Codmacta As String

    On Error GoTo EPasarCal

    PasarCalculoAContab = False
    
    'Total de lineas de asiento a Insertar en la contabilidad
    SQL = "SELECT count(*)" & _
          " FROM movim " & _
          "WHERE " & cadwhere
             
    NumLinea = TotalRegistros(SQL)
    
    If NumLinea = 0 Then Exit Function
    
    NumLinea = NumLinea * 3
    
    If NumLinea > 0 Then
        NumLinea = NumLinea + 1
        
        CargarProgres Me.Pb1, NumLinea
        
        ConnConta.BeginTrans
        conn.BeginTrans
        
        Set Mc = New CContadorContab
        
        If Mc.ConseguirContador("0", (CDate(txtCodigo(0).Text) <= CDate(FFin)), True, cConta) = 0 Then
        
        Obs = "Contabilizacion de Cálculo de Intereses AVNICS de fecha " & Format(txtCodigo(0).Text, "dd/mm/yyyy")

    
        'Insertar en la conta Cabecera Asiento
        b = InsertarCabAsientoDia("1", Mc.Contador, txtCodigo(0).Text, Obs, cadMen, cConta)
        cadMen = "Insertando Cab. Asiento: " & cadMen
        
        If b Then
            SQL = "SELECT codavnic, timporte, timport1, timport2 " & _
                  " FROM movim " & _
                  " WHERE " & cadwhere
            
            Set RS = New ADODB.Recordset
            
            RS.Open SQL, conn, adOpenDynamic, adLockOptimistic, adCmdText
            
            i = 0
            ImporteD = 0
            ImporteH = 0
            
            ampliacion = "Int.AVNICS AriagroUtil"
            ampliaciond = Trim(DevuelveDesdeBDNew(cConta, "conceptos", "nomconce", "codconce", vParamAplic.ConceDebe, "N")) & " " & ampliacion
            ampliacionh = Trim(DevuelveDesdeBDNew(cConta, "conceptos", "nomconce", "codconce", vParamAplic.ConceHaber, "N")) & " " & ampliacion
            
            
            If Not RS.EOF Then RS.MoveFirst
            While Not RS.EOF And b
                Codmacta = ""
                Codmacta = DevuelveDesdeBDNew(cPTours, "avnic", "codmacta", "codavnic", RS.Fields(0).Value, "N", , "anoejerc", Year(CDate(txtCodigo(0).Text)), "N")
                
                numdocum = "Av-" & Format(DBLet(RS!codavnic, "N"), "000000")
                ' ******************IMPORTE BRUTO
                i = i + 1
                
                cad = "1," & DBSet(txtCodigo(0).Text, "F") & "," & DBSet(Mc.Contador, "N") & ","
                cad = cad & DBSet(i, "N") & "," & DBSet(vParamAplic.CtaGasto, "T") & "," & DBSet(numdocum, "T") & ","
                
                ' COMPROBAMOS EL SIGNO DEL IMPORTE PQ NO PERMITIMOS INTRODUCIR APUNTES CON IMPORTES NEGATIVOS
                If RS.Fields(2).Value > 0 Then
                    ' importe al debe en positivo
                    cad = cad & DBSet(vParamAplic.ConceDebe, "N") & "," & DBSet(ampliaciond, "T") & "," & DBSet(RS.Fields(2).Value, "N") & ","
                    cad = cad & ValorNulo & "," & ValorNulo & "," & DBSet(Codmacta, "T") & "," & ValorNulo & ",0"
                
                    ImporteD = ImporteD + CCur(RS.Fields(2).Value)
                Else
                    ' importe al haber en positivo, cambiamos el signo
                    cad = cad & DBSet(vParamAplic.ConceHaber, "N") & "," & DBSet(ampliacionh, "T") & "," & ValorNulo & ","
                    cad = cad & DBSet((RS.Fields(2).Value * -1), "N") & "," & ValorNulo & "," & DBSet(Codmacta, "T") & "," & ValorNulo & ",0"
                
                    ImporteH = ImporteH + (CCur(RS.Fields(2).Value) * (-1))
                End If
                
                cad = "(" & cad & ")"
                
                b = InsertarLinAsientoDia(cad, cadMen, cConta)
                cadMen = "Insertando Lin. Asiento: " & i
            
                IncrementarProgres Me.Pb1, 1
                Me.lblProgres(1).Caption = "Insertando línea de Asiento en Contabilidad...   (" & i & " de " & NumLinea & ")"
                Me.Refresh
                
                ' ******************RETENCION
                i = i + 1
                
                cad = "1," & DBSet(txtCodigo(0).Text, "F") & "," & DBSet(Mc.Contador, "N") & ","
                cad = cad & DBSet(i, "N") & "," & DBSet(vParamAplic.CtaReten, "T") & "," & DBSet(numdocum, "T") & ","
                
                ' COMPROBAMOS EL SIGNO DEL IMPORTE PQ NO PERMITIMOS INTRODUCIR APUNTES CON IMPORTES NEGATIVOS
                If RS.Fields(3).Value > 0 Then
                    ' importe al haber en positivo
                    cad = cad & DBSet(vParamAplic.ConceHaber, "N") & "," & DBSet(ampliacionh, "T") & "," & ValorNulo & ","
                    cad = cad & DBSet(RS.Fields(3).Value, "N") & "," & ValorNulo & "," & DBSet(Codmacta, "T") & "," & ValorNulo & ",0"
                
                    ImporteH = ImporteH + CCur(RS.Fields(3).Value)
                Else
                    ' importe al debe en positivo, cambiamos el signo
                    cad = cad & DBSet(vParamAplic.ConceDebe, "N") & "," & DBSet(ampliaciond, "T") & "," & DBSet((RS.Fields(3).Value * -1), "N") & ","
                    cad = cad & ValorNulo & "," & ValorNulo & "," & DBSet(Codmacta, "T") & "," & ValorNulo & ",0"
                
                    ImporteH = ImporteH + (CCur(RS.Fields(3).Value) * (-1))
                End If
                
                cad = "(" & cad & ")"
                
                b = InsertarLinAsientoDia(cad, cadMen, cConta)
                cadMen = "Insertando Lin. Asiento: " & i
            
                IncrementarProgres Me.Pb1, 1
                Me.lblProgres(1).Caption = "Insertando línea de Asiento en Contabilidad...   (" & i & " de " & NumLinea & ")"
                Me.Refresh
                
                ' ******************IMPORTE NETO
                i = i + 1
                
                cad = "1," & DBSet(txtCodigo(0).Text, "F") & "," & DBSet(Mc.Contador, "N") & ","
                cad = cad & DBSet(i, "N") & "," & DBSet(Codmacta, "T") & "," & DBSet(numdocum, "T") & ","
                
                ' COMPROBAMOS EL SIGNO DEL IMPORTE PQ NO PERMITIMOS INTRODUCIR APUNTES CON IMPORTES NEGATIVOS
                If RS.Fields(1).Value > 0 Then
                    ' importe al haber en positivo
                    cad = cad & DBSet(vParamAplic.ConceHaber, "N") & "," & DBSet(ampliacionh, "T") & "," & ValorNulo & ","
                    cad = cad & DBSet(RS.Fields(1).Value, "N") & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & ",0"
                
                    ImporteH = ImporteH + CCur(RS.Fields(1).Value)
                Else
                    ' importe al debe en positivo, cambiamos el signo
                    cad = cad & DBSet(vParamAplic.ConceDebe, "N") & "," & DBSet(ampliaciond, "T") & "," & DBSet((RS.Fields(1).Value * -1), "N") & ","
                    cad = cad & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & ",0"
                
                    ImporteH = ImporteH + (CCur(RS.Fields(1).Value) * (-1))
                End If
                
                cad = "(" & cad & ")"
                
                b = InsertarLinAsientoDia(cad, cadMen, cConta)
                cadMen = "Insertando Lin. Asiento: " & i
            
                b = InsertarEnTesoreriaNew(txtCodigo(0).Text, txtCodigo(1).Text, RS.Fields(0).Value, Year(CDate(txtCodigo(0).Text)), txtCodigo(4).Text, txtCodigo(2).Text, txtCodigo(3).Text, cadMen)
                cadMen = "Insertando en Tesoreria: "
               
            
                IncrementarProgres Me.Pb1, 1
                Me.lblProgres(1).Caption = "Insertando línea de Asiento en Contabilidad...   (" & i & " de " & NumLinea & ")"
                Me.Refresh
                
            
                RS.MoveNext
            Wend
            RS.Close
            
            If b Then
                'Poner intconta=1 en ariagroutil.movim
                b = ActualizarMovimientos(cadwhere, cadMen)
                cadMen = "Actualizando Movimientos: " & cadMen
            End If
        End If
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
