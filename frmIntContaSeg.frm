VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmIntContaSeg 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Integración del Asiento Contable "
   ClientHeight    =   5700
   ClientLeft      =   45
   ClientTop       =   2430
   ClientWidth     =   8160
   Icon            =   "frmIntContaSeg.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5700
   ScaleWidth      =   8160
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FrameCobros 
      Height          =   5460
      Left            =   150
      TabIndex        =   3
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
         Height          =   1080
         Left            =   135
         TabIndex        =   5
         Top             =   2295
         Width           =   6075
         Begin VB.TextBox txtCodigo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   2
            Left            =   1980
            MaxLength       =   10
            TabIndex        =   0
            Tag             =   "Código Postal|T|S|||clientes|codposta|||"
            Top             =   450
            Width           =   1080
         End
         Begin VB.Label Label4 
            Caption         =   "Fecha Asiento"
            ForeColor       =   &H00972E0B&
            Height          =   255
            Index           =   0
            Left            =   180
            TabIndex        =   9
            Top             =   495
            Width           =   1425
         End
         Begin VB.Image imgFec 
            Height          =   240
            Index           =   2
            Left            =   1710
            Picture         =   "frmIntContaSeg.frx":000C
            ToolTipText     =   "Buscar fecha"
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
         Height          =   1725
         Left            =   135
         TabIndex        =   4
         Top             =   315
         Width           =   6090
         Begin VB.TextBox txtCodigo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   1
            Left            =   1980
            MaxLength       =   10
            TabIndex        =   11
            Top             =   1035
            Width           =   1050
         End
         Begin VB.TextBox txtCodigo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   0
            Left            =   1980
            MaxLength       =   10
            TabIndex        =   10
            Top             =   675
            Width           =   1050
         End
         Begin VB.Image imgFec 
            Height          =   240
            Index           =   1
            Left            =   1710
            Picture         =   "frmIntContaSeg.frx":0097
            ToolTipText     =   "Buscar fecha"
            Top             =   1035
            Width           =   240
         End
         Begin VB.Image imgFec 
            Height          =   240
            Index           =   0
            Left            =   1710
            Picture         =   "frmIntContaSeg.frx":0122
            ToolTipText     =   "Buscar fecha"
            Top             =   690
            Width           =   240
         End
         Begin VB.Label Label4 
            Caption         =   "Hasta"
            Height          =   195
            Index           =   14
            Left            =   1125
            TabIndex        =   14
            Top             =   1050
            Width           =   420
         End
         Begin VB.Label Label4 
            Caption         =   "Desde"
            Height          =   195
            Index           =   15
            Left            =   1125
            TabIndex        =   13
            Top             =   690
            Width           =   465
         End
         Begin VB.Label Label4 
            Caption         =   "Fecha Póliza"
            ForeColor       =   &H00972E0B&
            Height          =   255
            Index           =   1
            Left            =   180
            TabIndex        =   12
            Top             =   450
            Width           =   1815
         End
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   0
         Left            =   4995
         TabIndex        =   2
         Top             =   4815
         Width           =   975
      End
      Begin VB.CommandButton CmdAceptar 
         Caption         =   "&Aceptar"
         Height          =   375
         Index           =   0
         Left            =   3810
         TabIndex        =   1
         Top             =   4815
         Width           =   975
      End
      Begin MSComctlLib.ProgressBar Pb1 
         Height          =   345
         Left            =   180
         TabIndex        =   6
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
         TabIndex        =   8
         Top             =   4095
         Width           =   5265
      End
      Begin VB.Label lblProgres 
         Height          =   285
         Index           =   1
         Left            =   495
         TabIndex        =   7
         Top             =   4410
         Width           =   5295
      End
   End
   Begin VB.Frame FrameSeleccion 
      Height          =   5505
      Left            =   0
      TabIndex        =   15
      Top             =   0
      Width           =   8055
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFC0C0&
         Height          =   315
         Index           =   1
         Left            =   6495
         TabIndex        =   18
         Text            =   "Text2"
         Top             =   4980
         Width           =   1305
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   1
         Left            =   1500
         TabIndex        =   17
         Top             =   4980
         Width           =   1095
      End
      Begin VB.CommandButton CmdAceptar 
         Caption         =   "Aceptar"
         Height          =   375
         Index           =   1
         Left            =   240
         TabIndex        =   16
         Top             =   4980
         Width           =   1185
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   4425
         Left            =   90
         TabIndex        =   19
         Top             =   420
         Width           =   7815
         _ExtentX        =   13785
         _ExtentY        =   7805
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         Checkboxes      =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   16777215
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin VB.Image Image1 
         Height          =   240
         Index           =   1
         Left            =   390
         Picture         =   "frmIntContaSeg.frx":01AD
         ToolTipText     =   "Desmarcar todos"
         Top             =   120
         Width           =   240
      End
      Begin VB.Image Image1 
         Height          =   240
         Index           =   0
         Left            =   150
         Picture         =   "frmIntContaSeg.frx":0BAF
         ToolTipText     =   "Marcar todos"
         Top             =   120
         Width           =   240
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00C00000&
         X1              =   6300
         X2              =   7785
         Y1              =   4905
         Y2              =   4905
      End
      Begin VB.Label Label1 
         Caption         =   "Totales"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   1
         Left            =   5265
         TabIndex        =   20
         Top             =   5010
         Width           =   795
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
Attribute VB_Name = "frmIntContaSeg"
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
Private WithEvents frmFPa As frmForpaConta 'formas de pago de la contabilidad
Attribute frmFPa.VB_VarHelpID = -1

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
Dim NRegSelec As Integer
Dim cadwhere As String

Private Sub KEYpress(KeyAscii As Integer)
    If KeyAscii = 13 Then 'ENTER
        KeyAscii = 0
        SendKeys "{tab}"
    ElseIf KeyAscii = 27 Then Unload Me  'ESC
    End If
End Sub

Private Sub cmdAceptar_Click(Index As Integer)
Dim SQL As String
Dim i As Byte
'Dim cadwhere As String

    
    Select Case Index
        Case 0
            If Not DatosOk Then Exit Sub
                     
            cadwhere = "intconta = 0 "
            If txtCodigo(0).Text <> "" Then cadwhere = cadwhere & " and fechaenv >= " & DBSet(txtCodigo(0).Text, "F")
            If txtCodigo(1).Text <> "" Then cadwhere = cadwhere & " and fechaenv <= " & DBSet(txtCodigo(1).Text, "F")
            
            If CargarListView(cadwhere) = 0 Then
                VisualizarListview False
                MsgBox "No existen datos entre esos límites.", vbExclamation
            Else
                VisualizarListview True
            End If
        Case 1
            If NRegSelec = 0 Then
                MsgBox "No ha seleccionado ninguna poliza para realizar cobro.", vbExclamation
                Exit Sub
            Else
                If MsgBox("Desea continuar con el proceso de contabilizacion de pólizas de cliente.", vbQuestion + vbYesNo + vbDefaultButton2) = vbYes Then
                    ContabilizarAsiento (cadwhere)
                     'Eliminar la tabla TMP
                    BorrarTMPErrComprob
                    DesBloqueoManual ("CONASI") 'CONtabilizacion de asiento
                    Pb1.visible = False
                    lblProgres(0).Caption = ""
                    lblProgres(1).Caption = ""
                    cmdCancel_Click (1)
                End If
            End If
    End Select
    
    
eError:
    If Err.Number <> 0 Then
        MuestraError Err.Number, "No se ha realizado el proceso de contabilización de pólizas. Llame a soporte."
    End If
    
End Sub

Private Sub cmdCancel_Click(Index As Integer)
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
     txtCodigo(2).Text = Format(Now, "dd/mm/yyyy") ' fecha del asiento
     
    '###Descomentar
'    CommitConexion
    VisualizarListview False
    NRegSelec = 0
         
    FrameCobrosVisible True, h, w
    Pb1.visible = False
    
    'Esto se consigue poneinedo el cancel en el opcion k corresponda
    Me.cmdCancel(1).Cancel = True
'   Me.Width = w + 70
'   Me.Height = h + 350
End Sub

Private Sub frmC_Selec(vFecha As Date)
 'Fecha
    txtCodigo(CByte(imgFec(0).Tag)).Text = Format(vFecha, "dd/MM/yyyy")
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
            Case 0: KEYFecha KeyAscii, 0 'fecha desde
            Case 1: KEYFecha KeyAscii, 1 'fecha hasta
            Case 2: KEYFecha KeyAscii, 2 'fecha asiento
        End Select
    Else
        KEYpress KeyAscii
    End If

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
        Case 0, 1, 2 'FECHAS
            If txtCodigo(Index).Text <> "" Then PonerFormatoFecha txtCodigo(Index)
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

Private Sub ContabilizarAsiento(cadwhere As String)
'Contabiliza Facturas de Clientes o de Proveedores
Dim SQL As String
Dim b As Boolean
Dim tmpErrores As Boolean 'Indica si se creo correctamente la tabla de errores
Dim CCoste As String
Dim cadTABLA As String
Dim cadWhere1 As String

    SQL = "CONASI" 'contabilizar Poliza de seguros

    'Bloquear para que nadie mas pueda contabilizar
    DesBloqueoManual (SQL)
    If Not BloqueoManual(SQL, "1") Then
        MsgBox "No se pueden Contabilizar las Pólizas. Hay otro usuario contabilizándolo.", vbExclamation
        Screen.MousePointer = vbDefault
        Exit Sub
    End If

    VisualizarListview False
    
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
    
    
    'comprobar que todas las CUENTAS de referencias de polizas existan
    'en la Conta: segpoliza.codmacta IN (conta.cuentas.codmacta)
    '-----------------------------------------------------------------------------
    Me.lblProgres(1).Caption = "Comprobando Cuenta Ctble de Pólizas ..."
    
    cadWhere1 = "codmacta in (" & CargarCtas(cadwhere) & ")"
    
    b = ComprobarCtaContable(cadTABLA, 2, cadWhere1, cContaSeg)
    IncrementarProgres Me.Pb1, 50
    Me.Refresh
    If Not b Then
        frmMensaje.OpcionMensaje = 2
        frmMensaje.Show vbModal
        Exit Sub
    End If

    
    'comprobar que todas las CUENTAS de banco  existe
    'en la Conta: sparam.ctabancoseg IN (conta.cuentas.codmacta)
    '-----------------------------------------------------------------------------
    Me.lblProgres(1).Caption = "Comprobando Cuenta Ctble de Banco ..."
    b = ComprobarCtaContable(cadTABLA, 5, , cContaSeg)
    IncrementarProgres Me.Pb1, 40
    Me.Refresh
    If Not b Then
        frmMensaje.OpcionMensaje = 2
        frmMensaje.Show vbModal
        Exit Sub
    End If
    
    '===========================================================================
    'CONTABILIZAR CIERRE
    '===========================================================================
    Me.lblProgres(0).Caption = "Contabilizar Asiento: "
    CargarProgres Me.Pb1, 10
    Me.lblProgres(1).Caption = "Insertando Asiento en Contabilidad..."
    
    
    b = PasarAContabNew(cadwhere)
    
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
Dim FFinSeg As Date

   b = True

   If txtCodigo(2).Text = "" And b Then
        MsgBox "Introduzca la Fecha de Asiento.", vbExclamation
        b = False
        PonerFoco txtCodigo(2)
   Else
        ' comprobamos que la contabilizacion se encuentre en los ejercicios contables
         Orden1 = ""
         Orden1 = DevuelveDesdeBDNew(cContaSeg, "parametros", "fechaini", "", "", "", "", "", "", "", "", "", "")
    
         Orden2 = ""
         Orden2 = DevuelveDesdeBDNew(cContaSeg, "parametros", "fechafin", "", "", "", "", "", "", "", "", "", "")
         FFinSeg = CDate(Orden2)
         If Not (CDate(Orden1) <= CDate(txtCodigo(2).Text) And CDate(txtCodigo(2).Text) < CDate(Day(FIniSeg) & "/" & Month(FIniSeg) & "/" & Year(FIniSeg) + 2)) Then
            MsgBox "La Fecha de la contabilización no es del ejercicio actual ni del siguiente. Reintroduzca.", vbExclamation
            b = False
            PonerFoco txtCodigo(2)
         End If
   End If
   
   DatosOk = b
   
End Function

Private Function PasarAContab(cadwhere As String) As Boolean
Dim SQL As String
Dim RS As ADODB.Recordset
Dim b As Boolean
Dim i As Long
Dim J As Long
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
Dim Importe As Currency

    On Error GoTo EPasarCal

    PasarAContab = False
    
    'Total de lineas de asiento a Insertar en la contabilidad
'    SQL = "SELECT count(*)" & _
'          " FROM segpoliza " & _
'          "WHERE " & cadwhere
             
    NumLinea = Me.ListView1.ListItems.Count
    
    If NumLinea = 0 Then Exit Function
    
    If NumLinea > 0 Then
        NumLinea = NumLinea + 1
        
        CargarProgres Me.Pb1, NumLinea
        
        ConnContaSeg.BeginTrans
        conn.BeginTrans
        
        Set Mc = New CContadorContab
        
        If Mc.ConseguirContador("0", (CDate(txtCodigo(2).Text) <= CDate(FFinSeg)), True, cContaSeg) = 0 Then
        
        Obs = "Asiento de Contabilizacion de Pólizas de Seguros Agrarios de fecha " & Format(txtCodigo(2).Text, "dd/mm/yyyy")

    
        'Insertar en la conta Cabecera Asiento
        b = InsertarCabAsientoDia(vParamAplic.NumDiarioSeg, Mc.Contador, txtCodigo(2).Text, Obs, cadMen, cContaSeg)
        cadMen = "Insertando Cab. Asiento: " & cadMen
        
        If b Then
            i = 0
            ImporteD = 0
            ImporteH = 0
            
            ampliacion = "Pólizas Seguros"
            ampliaciond = Trim(DevuelveDesdeBDNew(cConta, "conceptos", "nomconce", "codconce", vParamAplic.ConceDebeSeg, "N")) & " " & ampliacion
            ampliacionh = Trim(DevuelveDesdeBDNew(cConta, "conceptos", "nomconce", "codconce", vParamAplic.ConceHaberSeg, "N")) & " " & ampliacion
            
            For J = 1 To Me.ListView1.ListItems.Count
                If Me.ListView1.ListItems(J).Checked Then
        
                    numdocum = DBLet(Me.ListView1.ListItems(J).Text, "T")
                    ' ******************IMPORTE de la poliza
                    i = i + 1
                    
                    cad = DBSet(vParamAplic.NumDiarioSeg, "N") & "," & DBSet(txtCodigo(2).Text, "F") & "," & DBSet(Mc.Contador, "N") & ","
                    cad = cad & DBSet(i, "N") & "," & DBSet(Me.ListView1.ListItems(J).SubItems(7), "T") & "," & DBSet(numdocum, "T") & ","
                    
                    ' COMPROBAMOS EL SIGNO DEL IMPORTE PQ NO PERMITIMOS INTRODUCIR APUNTES CON IMPORTES NEGATIVOS
                    If Me.ListView1.ListItems(J).SubItems(6) > 0 Then
                        ' importe al debe en positivo
                        cad = cad & DBSet(vParamAplic.ConceDebeSeg, "N") & "," & DBSet(ampliaciond, "T") & "," & DBSet(Me.ListView1.ListItems(J).SubItems(6), "N") & "," & ValorNulo & ","
                        cad = cad & ValorNulo & "," & DBSet(vParamAplic.CtaBancoSeg, "T") & "," & ValorNulo & ",0"
                    
                        ImporteD = ImporteD + CCur(Me.ListView1.ListItems(J).SubItems(6))
                    Else
                        ' importe al haber en positivo, cambiamos el signo
                        cad = cad & DBSet(vParamAplic.ConceHaberSeg, "N") & "," & DBSet(ampliacionh, "T") & "," & ValorNulo & "," & DBSet(Me.ListView1.ListItems(J).SubItems(7) * (-1), "N") & ","
                        cad = cad & ValorNulo & "," & DBSet(vParamAplic.CtaBancoSeg, "T") & "," & ValorNulo & ",0"
                    
                        ImporteH = ImporteH + (CCur(Me.ListView1.ListItems(J).SubItems(7)) * (-1))
                    End If
                    
                    cad = "(" & cad & ")"
                    
                    b = InsertarLinAsientoDia(cad, cadMen, cContaSeg)
                    cadMen = "Insertando Lin. Asiento: " & i
                
                    IncrementarProgres Me.Pb1, 1
                    Me.lblProgres(1).Caption = "Insertando línea de Asiento en Contabilidad...   (" & i & " de " & NumLinea & ")"
                    Me.Refresh
                End If
            Next J
            ' insertamos la contrapartida con la diferencia entre imported y importeh al debe
            Importe = ImporteD - ImporteH
            
            i = i + 1
            
            cad = vParamAplic.NumDiario & "," & DBSet(txtCodigo(2).Text, "F") & "," & DBSet(Mc.Contador, "N") & ","
            cad = cad & DBSet(i, "N") & "," & DBSet(vParamAplic.CtaBancoSeg, "T") & "," & DBSet(numdocum, "T") & ","
            
            ' COMPROBAMOS EL SIGNO DEL IMPORTE PQ NO PERMITIMOS INTRODUCIR APUNTES CON IMPORTES NEGATIVOS
            If Importe > 0 Then
                ' importe al haber en positivo
                cad = cad & DBSet(vParamAplic.ConceHaberSeg, "N") & "," & DBSet(ampliacionh, "T") & "," & ValorNulo & "," & DBSet(Importe, "N") & ","
                cad = cad & ValorNulo & "," & ValorNulo & "," & ValorNulo & ",0"
            Else
                ' importe al debe en positivo, cambiamos el signo
                cad = cad & DBSet(vParamAplic.ConceDebeSeg, "N") & "," & DBSet(ampliaciond, "T") & "," & DBSet(Importe * (-1), "N") & "," & ValorNulo & ","
                cad = cad & ValorNulo & "," & ValorNulo & "," & ValorNulo & ",0"
            End If
            
            cad = "(" & cad & ")"
            
            b = InsertarLinAsientoDia(cad, cadMen, cContaSeg)
            cadMen = "Insertando Lin. Asiento: " & i
        
            IncrementarProgres Me.Pb1, 1
            Me.lblProgres(1).Caption = "Insertando línea de Asiento en Contabilidad...   (" & i & " de " & NumLinea & ")"
            Me.Refresh
        
            
' de momento comentado para hacer pruebas
            If b Then
                 'Poner intconta=1 en ariagroutil.segpoliz
                 b = ActualizarIntPolizas(cadwhere, cadMen)
                 cadMen = "Actualizando Pólizas: " & cadMen
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
        ConnContaSeg.CommitTrans
        conn.CommitTrans
        PasarAContab = True
    Else
        ConnContaSeg.RollbackTrans
        conn.RollbackTrans
        PasarAContab = False
    End If
End Function

Private Sub ListView1_ItemCheck(ByVal Item As MSComctlLib.ListItem)
Dim i As Integer
Dim TotalCant As Currency
Dim TotalImporte As Currency
    
    Screen.MousePointer = vbHourglass
    
    TotalImporte = 0
    NRegSelec = 0
    
    ' vemos si lo podemos seleccionar
    For i = 1 To ListView1.ListItems.Count
        If ListView1.ListItems(i).Checked Then
            TotalImporte = TotalImporte + CCur(ListView1.ListItems(i).SubItems(6))
            NRegSelec = NRegSelec + 1
        End If
    Next i
    
    Screen.MousePointer = vbDefault

    
    Text1(1).Text = TotalImporte
    
End Sub

Private Function CargarCtas(cadwhere) As String
Dim SQL As String
Dim i As Long
Dim Codmacta As String

    CargarCtas = ""
    ' vemos si lo podemos seleccionar
    For i = 1 To ListView1.ListItems.Count
        If ListView1.ListItems(i).Checked Then
            Codmacta = ""
            Codmacta = DevuelveDesdeBDNew(cPTours, "segpoliza", "codmacta", "codiplan", ListView1.ListItems(i).SubItems(1), "N", , "codlinea", ListView1.ListItems(i).SubItems(2), "N", "codrefer", ListView1.ListItems(i).Text, "T")
            
            CargarCtas = CargarCtas & DBSet(Codmacta, "T") & ","
        End If
    Next i
    ' quitamos la ultima coma de la cadena
    CargarCtas = Mid(CargarCtas, 1, Len(CargarCtas) - 1)
End Function

Private Function CargarListView(cadwhere As String) As Integer
'Muestra la lista Detallada de Facturas que dieron error al contabilizar
'en un ListView
Dim RS As ADODB.Recordset
Dim ItmX As ListItem
Dim SQL As String
Dim HayReg As Integer

    On Error GoTo ECargarList
    
    HayReg = 0
    CargarListView = 0
    If ListView1.ListItems.Count <> 0 Then Exit Function

    Screen.MousePointer = vbHourglass

    Me.FrameSeleccion.Height = 5415
    Me.FrameSeleccion.Width = 8055
    Me.Height = 6120
    Me.Width = 8370
    
    SQL = " SELECT  codrefer, codiplan, codlinea, colectiv, nifasegu, nomasegu, imppoliz, codmacta "
    SQL = SQL & " FROM segpoliza where " & cadwhere
    SQL = SQL & " order by codrefer, codiplan, codlinea "
    
    Set RS = New ADODB.Recordset
    RS.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not RS.EOF Then
        VisualizarListview True
    
        'Los encabezados
        ListView1.ColumnHeaders.Clear

        ListView1.ColumnHeaders.Add , , "Referencia", 1100
        ListView1.ColumnHeaders.Add , , "Plan", 600, 1
        ListView1.ColumnHeaders.Add , , "Li.", 500
        ListView1.ColumnHeaders.Add , , "Colectivo", 950, 1
        ListView1.ColumnHeaders.Add , , "NIF", 1050, 1
        ListView1.ColumnHeaders.Add , , "Asegurado", 2300, 0
        ListView1.ColumnHeaders.Add , , "Importe", 1200, 1
        ListView1.ColumnHeaders.Add , , "", 0, 0
       
        ListView1.ListItems.Clear
        
        While Not RS.EOF
            Set ItmX = ListView1.ListItems.Add
            'El primer campo será codtipom si llamamos desde Ventas
            ' y será codprove si llamamos desde Compras
            ItmX.Text = RS!Codrefer
            ItmX.SubItems(1) = DBLet(RS!CodiPlan, "N")
            ItmX.SubItems(2) = DBLet(RS!CodLinea, "N")
            ItmX.SubItems(3) = DBLet(RS!Colectiv, "N")
            ItmX.SubItems(4) = DBLet(RS!nifasegu, "T")
            ItmX.SubItems(5) = DBLet(RS!NomAsegu, "T")
            ItmX.SubItems(6) = DBLet(RS!imppoliz, "N")
            ItmX.SubItems(7) = DBLet(RS!Codmacta, "T")
            HayReg = 1
            RS.MoveNext
        Wend
    End If
    RS.Close
    Set RS = Nothing
    CargarListView = HayReg
    
    Screen.MousePointer = vbDefault
    Exit Function
ECargarList:
    If Err.Number <> 0 Then
        Err.Clear
    End If
End Function

Private Sub VisualizarListview(Modo As Boolean)
    If Modo = False Then
        Me.Width = 6570
    Else
        Me.Width = 8250
    End If
    FrameSeleccion.visible = Modo
    FrameCobros.visible = Not Modo
End Sub

Private Sub Image1_Click(Index As Integer)
Dim i As Integer
Dim TotalCant As Currency
Dim TotalImporte As Currency

    Screen.MousePointer = vbHourglass
    
    TotalCant = 0
    TotalImporte = 0
    NRegSelec = 0
    
    Select Case Index
        Case 0
            For i = 1 To ListView1.ListItems.Count
                ListView1.ListItems(i).Checked = True
                TotalImporte = TotalImporte + CCur(ListView1.ListItems(i).SubItems(6))
                NRegSelec = NRegSelec + 1
            Next i
        Case 1
            For i = 1 To ListView1.ListItems.Count
                ListView1.ListItems(i).Checked = False
            Next i
    End Select
    Screen.MousePointer = vbDefault

    Text1(1).Text = TotalImporte
End Sub

Private Function ActualizarIntPolizas(cadwhere As String, caderr As String) As Boolean
'Poner el movimiento como contabilizada
Dim SQL As String
Dim J As Integer

    On Error GoTo EActualizar
    
    
    For J = 1 To Me.ListView1.ListItems.Count
        If Me.ListView1.ListItems(J).Checked Then
            SQL = "UPDATE segpoliza SET intconta=1 "
            SQL = SQL & " WHERE codrefer = " & DBSet(Me.ListView1.ListItems(J).Text, "T") & " and "
            SQL = SQL & " codiplan = " & DBSet(Me.ListView1.ListItems(J).SubItems(1), "N") & " and "
            SQL = SQL & " codlinea = " & DBSet(Me.ListView1.ListItems(J).SubItems(2), "N")
    
            conn.Execute SQL
        End If
    Next J
    
EActualizar:
    If Err.Number <> 0 Then
        ActualizarIntPolizas = False
        caderr = Err.Description
    Else
        ActualizarIntPolizas = True
    End If
End Function

'se crea un asiento por cada cargo que se hace a la cuenta de banco
Private Function PasarAContabNew(cadwhere As String) As Boolean
Dim SQL As String
Dim RS As ADODB.Recordset
Dim b As Boolean
Dim i As Long
Dim J As Long
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
Dim Importe As Currency

    On Error GoTo EPasarCal

    PasarAContabNew = False
    
    'Total de lineas de asiento a Insertar en la contabilidad
'    SQL = "SELECT count(*)" & _
'          " FROM segpoliza " & _
'          "WHERE " & cadwhere
             
    NumLinea = Me.ListView1.ListItems.Count
    
    If NumLinea = 0 Then Exit Function
    
    If NumLinea > 0 Then
        CargarProgres Me.Pb1, NumLinea
        
        ConnContaSeg.BeginTrans
        conn.BeginTrans
        
        Obs = "Asiento de Contabilizacion de Pólizas de Seguros Agrarios de fecha " & Format(txtCodigo(2).Text, "dd/mm/yyyy")
        
        For J = 1 To Me.ListView1.ListItems.Count
            If Me.ListView1.ListItems(J).Checked Then
        
                Set Mc = New CContadorContab
                
                If Mc.ConseguirContador("0", (CDate(txtCodigo(2).Text) <= CDate(FFinSeg)), True, cContaSeg) = 0 Then
            
                'Insertar en la conta Cabecera Asiento
                b = InsertarCabAsientoDia(vParamAplic.NumDiarioSeg, Mc.Contador, txtCodigo(2).Text, Obs, cadMen, cContaSeg)
                cadMen = "Insertando Cab. Asiento: " & cadMen
        
                If b Then
                    i = 0
                    ImporteD = 0
                    ImporteH = 0
                    
                    ampliacion = "Pólizas Seguros"
                    ampliaciond = Trim(DevuelveDesdeBDNew(cConta, "conceptos", "nomconce", "codconce", vParamAplic.ConceDebeSeg, "N")) & " " & ampliacion
                    ampliacionh = Trim(DevuelveDesdeBDNew(cConta, "conceptos", "nomconce", "codconce", vParamAplic.ConceHaberSeg, "N")) & " " & ampliacion
            
                    numdocum = DBLet(Me.ListView1.ListItems(J).Text, "T")
                    ' ******************IMPORTE de la poliza
                    i = i + 1
                    
                    cad = DBSet(vParamAplic.NumDiarioSeg, "N") & "," & DBSet(txtCodigo(2).Text, "F") & "," & DBSet(Mc.Contador, "N") & ","
                    cad = cad & DBSet(i, "N") & "," & DBSet(Me.ListView1.ListItems(J).SubItems(7), "T") & "," & DBSet(numdocum, "T") & ","
                    
                    ' COMPROBAMOS EL SIGNO DEL IMPORTE PQ NO PERMITIMOS INTRODUCIR APUNTES CON IMPORTES NEGATIVOS
                    If Me.ListView1.ListItems(J).SubItems(6) > 0 Then
                        ' importe al debe en positivo
                        cad = cad & DBSet(vParamAplic.ConceDebeSeg, "N") & "," & DBSet(ampliaciond, "T") & "," & DBSet(Me.ListView1.ListItems(J).SubItems(6), "N") & "," & ValorNulo & ","
                        cad = cad & ValorNulo & "," & DBSet(vParamAplic.CtaBancoSeg, "T") & "," & ValorNulo & ",0"
                    
                        ImporteD = ImporteD + CCur(Me.ListView1.ListItems(J).SubItems(6))
                    Else
                        ' importe al haber en positivo, cambiamos el signo
                        cad = cad & DBSet(vParamAplic.ConceHaberSeg, "N") & "," & DBSet(ampliacionh, "T") & "," & ValorNulo & "," & DBSet(Me.ListView1.ListItems(J).SubItems(7) * (-1), "N") & ","
                        cad = cad & ValorNulo & "," & DBSet(vParamAplic.CtaBancoSeg, "T") & "," & ValorNulo & ",0"
                    
                        ImporteH = ImporteH + (CCur(Me.ListView1.ListItems(J).SubItems(7)) * (-1))
                    End If
                    
                    cad = "(" & cad & ")"
                    
                    b = InsertarLinAsientoDia(cad, cadMen, cContaSeg)
                    cadMen = "Insertando Lin. Asiento: " & i
                
                    ' CONTRAPARTIDA
                    ' insertamos la contrapartida con la diferencia entre imported y importeh al debe
                    Importe = ImporteD - ImporteH
                    
                    i = i + 1
                    
                    cad = vParamAplic.NumDiario & "," & DBSet(txtCodigo(2).Text, "F") & "," & DBSet(Mc.Contador, "N") & ","
                    cad = cad & DBSet(i, "N") & "," & DBSet(vParamAplic.CtaBancoSeg, "T") & "," & DBSet(numdocum, "T") & ","
                    
                    ' COMPROBAMOS EL SIGNO DEL IMPORTE PQ NO PERMITIMOS INTRODUCIR APUNTES CON IMPORTES NEGATIVOS
                    If Importe > 0 Then
                        ' importe al haber en positivo
                        cad = cad & DBSet(vParamAplic.ConceHaberSeg, "N") & "," & DBSet(ampliacionh, "T") & "," & ValorNulo & "," & DBSet(Importe, "N") & ","
                        cad = cad & ValorNulo & "," & DBSet(Me.ListView1.ListItems(J).SubItems(7), "T") & "," & ValorNulo & ",0"
                    Else
                        ' importe al debe en positivo, cambiamos el signo
                        cad = cad & DBSet(vParamAplic.ConceDebeSeg, "N") & "," & DBSet(ampliaciond, "T") & "," & DBSet(Importe * (-1), "N") & "," & ValorNulo & ","
                        cad = cad & ValorNulo & "," & DBSet(Me.ListView1.ListItems(J).SubItems(7), "T") & "," & ValorNulo & ",0"
                    End If
                    
                    cad = "(" & cad & ")"
                    
                    b = InsertarLinAsientoDia(cad, cadMen, cContaSeg)
                    cadMen = "Insertando Lin. Asiento: " & i
                
                    IncrementarProgres Me.Pb1, 1
                    Me.lblProgres(1).Caption = "Insertando Asiento en Contabilidad...   (" & J & " de " & NumLinea & ")"
                    Me.Refresh
                
                
                    If b Then
                         'Poner intconta=1 en ariagroutil.segpoliz
                         b = ActualizarIntPolizas(cadwhere, cadMen)
                         cadMen = "Actualizando Pólizas: " & cadMen
                    End If
                
                End If
            End If
            End If
        Next J
            
   End If
   
EPasarCal:
    If Err.Number <> 0 Then
        b = False
        MuestraError Err.Number, "Integrando Asiento a Contabilidad", Err.Description
    End If
    If b Then
        ConnContaSeg.CommitTrans
        conn.CommitTrans
        PasarAContabNew = True
    Else
        ConnContaSeg.RollbackTrans
        conn.RollbackTrans
        PasarAContabNew = False
    End If
End Function

