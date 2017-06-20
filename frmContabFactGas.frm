VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmContabFactGas 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Contabilización de Facturas Gasolinera"
   ClientHeight    =   3075
   ClientLeft      =   45
   ClientTop       =   2430
   ClientWidth     =   6645
   Icon            =   "frmContabFactGas.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3075
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
      Height          =   2880
      Left            =   90
      TabIndex        =   3
      Top             =   120
      Width           =   6375
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   7
         Left            =   1845
         MaxLength       =   10
         TabIndex        =   0
         Tag             =   "Código Postal|T|S|||clientes|codposta|||"
         Top             =   810
         Width           =   1050
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Left            =   4710
         TabIndex        =   2
         Top             =   2415
         Width           =   975
      End
      Begin VB.CommandButton cmdAceptar 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   3525
         TabIndex        =   1
         Top             =   2415
         Width           =   975
      End
      Begin MSComctlLib.ProgressBar Pb1 
         Height          =   285
         Left            =   360
         TabIndex        =   4
         Top             =   1980
         Width           =   5265
         _ExtentX        =   9287
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.Label lblProgres 
         Height          =   285
         Index           =   1
         Left            =   360
         TabIndex        =   7
         Top             =   1620
         Width           =   5265
      End
      Begin VB.Label lblProgres 
         Height          =   285
         Index           =   0
         Left            =   360
         TabIndex        =   6
         Top             =   1350
         Width           =   5265
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   7
         Left            =   1530
         Picture         =   "frmContabFactGas.frx":000C
         ToolTipText     =   "Buscar fecha"
         Top             =   810
         Width           =   240
      End
      Begin VB.Label Label4 
         Caption         =   "Fecha Factura"
         ForeColor       =   &H00972E0B&
         Height          =   255
         Index           =   4
         Left            =   360
         TabIndex        =   5
         Top             =   840
         Width           =   1815
      End
   End
End
Attribute VB_Name = "frmContabFactGas"
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
    InicializarVbles
    
    If Not DatosOk Then Exit Sub
    
    cadSelect = tabla & ".intconta=0 "
  
    'Fecha de Factura
    AnyadirAFormula cadFormula, tabla & ".fecfactu = " & DBSet(txtCodigo(7).Text, "F")
    AnyadirAFormula cadSelect, tabla & ".fecfactu = " & DBSet(txtCodigo(7).Text, "F")
    
    
    If Not HayRegParaInforme(tabla, cadSelect) Then Exit Sub
    
    ' abrimos la conexion de la contabilidad de la seccion correspondiente
'    If BdConta = 0 Then Exit Sub
            
    ContabilizarFacturas tabla, cadSelect
     'Eliminar la tabla TMP
    BorrarTMPFacturas
    'Desbloqueamos ya no estamos contabilizando facturas
    DesBloqueoManual ("GASCON") 'GASolinera CONtabilizar
    
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
        PonerFoco txtCodigo(7)
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
Dim h As Integer, w As Integer
Dim List As Collection

    PrimeraVez = True
    Limpiar Me

    '###Descomentar
'    CommitConexion
         
    FrameCobrosVisible True, h, w
    indFrame = 5
    tabla = "gascabfac"
    
    Pb1.visible = False
    
    'Esto se consigue poneinedo el cancel en el opcion k corresponda
    Me.cmdCancel.Cancel = True
    Me.Width = w + 70
    Me.Height = h + 70
    
    txtCodigo(7).Text = Format(Now, "dd/mm/yyyy")
    
 End Sub

Private Sub frmC_Selec(vFecha As Date)
 'Fecha
    txtCodigo(CByte(imgFec(7).Tag)).Text = Format(vFecha, "dd/MM/yyyy")
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
    PonerFoco txtCodigo(CByte(imgFec(7).Tag))
    ' ***************************
End Sub

Private Sub ListView1_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
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
            Case 7: KEYFecha KeyAscii, 7 'fecha de factura
        End Select
    Else
        KEYpress KeyAscii
    End If

End Sub

Private Sub KEYBusqueda(KeyAscii As Integer, indice As Integer)
    KeyAscii = 0
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
            
        Case 7   'FECHAS
            If txtCodigo(Index).Text <> "" Then PonerFormatoFecha txtCodigo(Index)
    
    End Select
End Sub

Private Sub FrameCobrosVisible(visible As Boolean, ByRef h As Integer, ByRef w As Integer)
    Me.FrameCobros.visible = visible
    If visible = True Then
        Me.FrameCobros.Top = -90
        Me.FrameCobros.Left = 0
        Me.FrameCobros.Height = 3750
        Me.FrameCobros.Width = 6735
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

Private Function DatosOk() As Boolean
Dim b As Boolean
Dim Cad As String

    DatosOk = False

    If txtCodigo(7).Text = "" Then
        MsgBox "Debe introducir obligatoriamente una Fecha de Factura.", vbExclamation
        PonerFoco txtCodigo(7)
        Exit Function
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

    Sql = "GASCON" 'contabilizar facturas de venta

    'Bloquear para que nadie mas pueda contabilizar
    DesBloqueoManual (Sql)
    If Not BloqueoManual(Sql, "1") Then
        MsgBox "No se pueden Contabilizar Facturas de Gasolinera. Hay otro usuario contabilizando.", vbExclamation
        Screen.MousePointer = vbDefault
        Exit Sub
    End If


     'comprobar que se han rellenado los dos campos de fecha
     'sino rellenar con fechaini o fechafin del ejercicio
     'que guardamos en vbles Orden1,Orden2
     If txtCodigo(7).Text = "" Then
        txtCodigo(7).Text = Orden1 'vEmpresa.FechaIni  'fechaini del ejercicio de la conta
     End If

'     If txtCodigo(3).Text = "" Then
'        txtCodigo(3).Text = Orden2 'vEmpresa.FechaFin  'fecha fin del ejercicio de la conta
'     End If

     'Comprobar que el intervalo de fechas D/H esta dentro del ejercicio de la
     'contabilidad par ello mirar en la BD de la Conta los parámetros
     If Not ComprobarFechasConta(7) Then Exit Sub
     
     

    'La comprobacion solo lo hago para facturas nuestras, ya que mas adelante
    'el programa hara cdate(text1(31) cuando contabilice las facturas y dara error de tipos
    If Me.txtCodigo(7).Text = "" Then
        MsgBox "Fecha inicio incorrecta", vbExclamation
        Exit Sub
    End If



    'comprobar si existen en Ariagroutil facturas anteriores al periodo solicitado
    'sin contabilizar
    If Me.txtCodigo(7).Text <> "" Then
        Sql = "SELECT COUNT(*) FROM " & cadTABLA
        Sql = Sql & " WHERE fecfactu <"
        Sql = Sql & DBSet(txtCodigo(7), "F") & " AND intconta=0 "
        If RegistrosAListar(Sql) > 0 Then
            MsgBox "Hay Facturas anteriores sin contabilizar.", vbExclamation
            Exit Sub
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
    b = CrearTMPFacturas(cadTABLA, cadwhere, False, False)
    If Not b Then Exit Sub
            
    BorrarTMPErrComprob
    ' nuevo
    b = CrearTMPErrComprob()
    If Not b Then Exit Sub
    
    'comprobar que todas las LETRAS SERIE existen en la contabilidad y en Ariagroutilgasol
    '-----------------------------------------------------------------------------
    IncrementarProgres Me.Pb1, 10
    Me.lblProgres(1).Caption = "Comprobando letras de serie ..."
    b = ComprobarLetraSerieGas()
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
        Sql = "anofactu>=" & Year(txtCodigo(7).Text) & " AND anofactu<= " & Year(txtCodigo(7).Text)
        b = ComprobarNumFacturasContaNueva(cContaGas, Sql)
    Else
        Sql = "anofaccl>=" & Year(txtCodigo(7).Text) & " AND anofaccl<= " & Year(txtCodigo(7).Text)
        b = ComprobarNumFacturas(cContaGas, Sql)
    End If
    IncrementarProgres Me.Pb1, 20
    Me.Refresh
    If Not b Then
        frmMensaje.OpcionMensaje = 1
        frmMensaje.Show vbModal
        Exit Sub
    End If
    
    'comprobar que todas las CUENTAS de los distintos socios que vamos a
    'contabilizar existen en la Conta: vparamaplic.raizctasocgas & codsocio IN (conta.cuentas.codmacta)
    '-----------------------------------------------------------------------------
    Me.lblProgres(1).Caption = "Comprobando Cuentas Contables en contabilidad ..."
    b = ComprobarCtaContableGas(1, cadSelect)
    IncrementarProgres Me.Pb1, 20
    Me.Refresh
    If Not b Then
        frmMensaje.OpcionMensaje = 2
        frmMensaje.Show vbModal
        Exit Sub
    End If
    
    'comprobar que todas las CUENTAS de venta de la parametros
    '-----------------------------------------------------------------------------
    Me.lblProgres(1).Caption = "Comprobando Cuenta Ctble Ventas de parámetros en contabilidad ..."
    b = ComprobarCtaContableGas(2, cadSelect)
    IncrementarProgres Me.Pb1, 20
    Me.Refresh
    If Not b Then
        frmMensaje.OpcionMensaje = 2
        frmMensaje.Show vbModal
        Exit Sub
    End If
    
    'comprobar que todas las CUENTAS de venta de parametros
    'es del grupo de ventas: empiezan por conta.parametros.grupovtas
    '-----------------------------------------------------------------------------
    Me.lblProgres(1).Caption = "Comprobando Cuentas Ctbles Ventas en contabilidad ..."
    b = ComprobarCtaContableGas(3)
    IncrementarProgres Me.Pb1, 20
    Me.Refresh
    If Not b Then
        frmMensaje.OpcionMensaje = 2
        frmMensaje.Show vbModal
        Exit Sub
    End If
    
    
    'comprobar que todas la CUENTA de contrapartida
    'existe en la Conta: vparamaplic.ctacontra  IN (conta.cuentas.codmacta)
    '-----------------------------------------------------------------------------
    Me.lblProgres(1).Caption = "Comprobando Cuentas Contables de Contrapartida en contabilidad ..."
    
    b = ComprobarCtaContableGas(4)
    IncrementarProgres Me.Pb1, 20
    Me.Refresh
    If Not b Then
        frmMensaje.OpcionMensaje = 2
        frmMensaje.Show vbModal
        Exit Sub
    End If
    
    
    'comprobar que todos el TIPO IVA
    'existe en la Conta: vparamaplic.codivagas IN (conta.tiposiva.codigiva)
    '--------------------------------------------------------------------------
    Me.lblProgres(1).Caption = "Comprobando Tipos de IVA en contabilidad ..."
    b = ComprobarTiposIVAGas()
    IncrementarProgres Me.Pb1, 10
    Me.Refresh
    If Not b Then
        frmMensaje.OpcionMensaje = 3
        frmMensaje.Show vbModal
        Exit Sub
    End If
    
    
    '===========================================================================
    'CONTABILIZAR FACTURAS
    '===========================================================================
    Me.lblProgres(0).Caption = "Contabilizar Facturas: "
    CargarProgres Me.Pb1, 10
    Me.lblProgres(1).Caption = "Insertando Facturas en Contabilidad..."
       
    
    'Crear tabla TEMP para los posible errores de facturas
    tmpErrores = CrearTMPErrFact(cadTABLA)
    
    
    b = PasarFacturasAContab(cadTABLA, CCoste)
    
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


Private Function PasarFacturasAContab(cadTABLA As String, CCoste As String) As Boolean
Dim Sql As String
Dim Rs As ADODB.Recordset
Dim b As Boolean
Dim i As Integer
Dim numfactu As Integer
Dim codigo1 As String
Dim Mc As CContadorContab
Dim numlinea As Currency
Dim Obs As String
Dim cadMen As String
Dim TotalFacturas As Currency
Dim Cad As String
Dim ampliacion As String
Dim ampliaciond As String
Dim ampliacionh As String

    On Error GoTo EPasarFac

    PasarFacturasAContab = False
    
    ConnContaGas.BeginTrans
    conn.BeginTrans
     
    'Total de Facturas a Insertar en la contabilidad
    Sql = "SELECT count(*) "
    Sql = Sql & " FROM " & cadTABLA & " INNER JOIN tmpfactu "
    codigo1 = "letraser"
    Sql = Sql & " ON " & cadTABLA & "." & codigo1 & "=tmpfactu.numserie"
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
     CargarProgresNew Me.Pb1, numfactu
        
     Sql = "SELECT * "
     Sql = Sql & " FROM tmpfactu "
            
     Set Rs = New ADODB.Recordset
     Rs.Open Sql, conn, adOpenStatic, adLockPessimistic, adCmdText
     i = 1
        
     Set Mc = New CContadorContab
        
     If Mc.ConseguirContador("0", (CDate(txtCodigo(7).Text) <= CDate(FFinGas)), True, cContaGas) = 0 Then
        
        Obs = "Asiento de Contabilizacion de Facturas de Gasolinera de fecha " & Format(txtCodigo(7).Text, "dd/mm/yyyy")
        
        b = True
        
        'Insertar en la conta Cabecera Asiento
        b = InsertarCabAsientoDia(vParamAplic.NumDiarioGas, Mc.Contador, txtCodigo(7).Text, Obs, cadMen, cContaGas)
        cadMen = "Insertando Cab. Asiento: " & cadMen


        Set cContaFra = New cContabilizarFacturas
        
        If Not cContaFra.EstablecerValoresInciales(ConnContaGas) Then
            'NO ha establcedio los valores de la conta.  Le dejaremos seguir, avisando que
            ' obviamente, no va a contabilizar las FRAS
            Sql = "Si continua, las facturas se insertaran en el registro, pero no serán contabilizadas" & vbCrLf
            Sql = Sql & "en este momento. Deberán ser contabilizadas desde el ARICONTA" & vbCrLf & vbCrLf
            Sql = Sql & Space(50) & "¿Continuar?"
            If MsgBox(Sql, vbQuestion + vbYesNoCancel) <> vbYes Then Exit Function
        End If
        




        If b Then
            numlinea = 0
            TotalFacturas = 0
            
            'contabilizar cada una de las facturas seleccionadas
            While Not Rs.EOF And b
                Sql = cadTABLA & "." & codigo1 & "=" & DBSet(Trim(Rs.Fields(0)), "T") & " and numfactu=" & DBLet(Rs!numfactu, "N")
                Sql = Sql & " and fecfactu=" & DBSet(Rs!fecfactu, "F")
                
                numlinea = numlinea + 1
                
                b = PasarFacturaGas(Sql, Rs!fecfactu, CStr(Mc.Contador), CStr(numlinea), CCoste, TotalFacturas, cContaFra)
                
                IncrementarProgresNew Me.Pb1, 1
                Me.lblProgres(1).Caption = "Insertando Facturas en Contabilidad...   (" & i & " de " & numfactu & ")"
                Me.Refresh
                i = i + 1
                Rs.MoveNext
            Wend
            Rs.Close
            Set Rs = Nothing
            
            'insertamos la contrapartida
            
            ampliacion = "Facturas Gasolinera"
            ampliaciond = Trim(DevuelveDesdeBDNew(cContaGas, "conceptos", "nomconce", "codconce", vParamAplic.ConceDebeGas, "N")) & " " & ampliacion
            ampliacionh = Trim(DevuelveDesdeBDNew(cContaGas, "conceptos", "nomconce", "codconce", vParamAplic.ConceHaberGas, "N")) & " " & ampliacion
            
            numlinea = numlinea + 1
            
            Cad = vParamAplic.NumDiarioGas & "," & DBSet(txtCodigo(7).Text, "F") & "," & DBSet(Mc.Contador, "N") & ","
                                                                                            '++monica:05/11/2008 antes valor nulo en el numero de documento
            Cad = Cad & DBSet(numlinea, "N") & "," & DBSet(vParamAplic.CtaContraGas, "T") & ",'" & Format(txtCodigo(7).Text, "ddmmyyyy") & "',"
            
            ' COMPROBAMOS EL SIGNO DEL IMPORTE PQ NO PERMITIMOS INTRODUCIR APUNTES CON IMPORTES NEGATIVOS
            If TotalFacturas < 0 Then
                ' importe al haber en positivo, cambiamos el signo
                Cad = Cad & DBSet(vParamAplic.ConceHaberGas, "N") & "," & DBSet(ampliacionh, "T") & "," & ValorNulo & "," & DBSet(TotalFacturas * (-1), "N") & ","
                Cad = Cad & ValorNulo & "," & ValorNulo & "," & ValorNulo & ",0"
            Else
                ' importe al debe en positivo
                Cad = Cad & DBSet(vParamAplic.ConceDebeGas, "N") & "," & DBSet(ampliaciond, "T") & "," & DBSet(TotalFacturas, "N") & "," & ValorNulo & ","
                Cad = Cad & ValorNulo & "," & ValorNulo & "," & ValorNulo & ",0"
            End If
            
            Cad = "(" & Cad & ")"
            
            b = InsertarLinAsientoDia(Cad, cadMen, cContaGas)
            cadMen = "Insertando Lin. Asiento: " & numlinea
        End If
        
     End If
    End If
    
EPasarFac:
    If Err.Number <> 0 Then
        b = False
        MuestraError Err.Number, "Contabilizando Factura", Err.Description
    End If
    If b Then
        ConnContaGas.CommitTrans
        conn.CommitTrans
        PasarFacturasAContab = True
    Else
        ConnContaGas.RollbackTrans
        conn.RollbackTrans
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
        Rs.Open FechaIni, ConnContaGas, adOpenForwardOnly, adLockPessimistic, adCmdText
    
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

