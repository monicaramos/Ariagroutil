VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmTrasFacAce 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Importación Facturas Aceite"
   ClientHeight    =   4980
   ClientLeft      =   45
   ClientTop       =   2430
   ClientWidth     =   6825
   Icon            =   "frmTrasFacAce.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4980
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
      Height          =   4800
      Left            =   180
      TabIndex        =   5
      Top             =   90
      Width           =   6555
      Begin VB.Frame Frame2 
         Height          =   1815
         Left            =   225
         TabIndex        =   11
         Top             =   225
         Width           =   6180
         Begin VB.TextBox txtCodigo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   2
            Left            =   1350
            MaxLength       =   10
            TabIndex        =   2
            Top             =   1170
            Width           =   1050
         End
         Begin VB.TextBox txtNombre 
            BackColor       =   &H80000018&
            Enabled         =   0   'False
            Height          =   285
            Index           =   2
            Left            =   2430
            Locked          =   -1  'True
            TabIndex        =   16
            Text            =   "Text5"
            Top             =   1170
            Width           =   3180
         End
         Begin VB.TextBox txtCodigo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   0
            Left            =   1350
            MaxLength       =   10
            TabIndex        =   1
            Top             =   720
            Width           =   1050
         End
         Begin VB.TextBox txtNombre 
            BackColor       =   &H80000018&
            Enabled         =   0   'False
            Height          =   285
            Index           =   0
            Left            =   2430
            Locked          =   -1  'True
            TabIndex        =   14
            Text            =   "Text5"
            Top             =   720
            Width           =   3180
         End
         Begin VB.TextBox txtNombre 
            BackColor       =   &H80000018&
            Enabled         =   0   'False
            Height          =   285
            Index           =   1
            Left            =   2430
            Locked          =   -1  'True
            TabIndex        =   12
            Text            =   "Text5"
            Top             =   315
            Width           =   3180
         End
         Begin VB.TextBox txtCodigo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   1
            Left            =   1350
            MaxLength       =   10
            TabIndex        =   0
            Top             =   315
            Width           =   1050
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "F.Pago"
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
            Left            =   270
            TabIndex        =   17
            Top             =   1215
            Width           =   510
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   2
            Left            =   1035
            MouseIcon       =   "frmTrasFacAce.frx":000C
            MousePointer    =   4  'Icon
            ToolTipText     =   "Buscar forma pago"
            Top             =   1215
            Width           =   240
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Concepto"
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
            Index           =   1
            Left            =   270
            TabIndex        =   15
            Top             =   765
            Width           =   690
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   0
            Left            =   1035
            MouseIcon       =   "frmTrasFacAce.frx":015E
            MousePointer    =   4  'Icon
            ToolTipText     =   "Buscar concepto"
            Top             =   765
            Width           =   240
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   1
            Left            =   1035
            MouseIcon       =   "frmTrasFacAce.frx":02B0
            MousePointer    =   4  'Icon
            ToolTipText     =   "Buscar Sección"
            Top             =   360
            Width           =   240
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
            Index           =   0
            Left            =   270
            TabIndex        =   13
            Top             =   360
            Width           =   540
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Fichero"
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
         Height          =   870
         Left            =   225
         TabIndex        =   9
         Top             =   2160
         Width           =   6180
         Begin VB.Label Label1 
            Caption         =   "Fichero: "
            Height          =   420
            Left            =   180
            TabIndex        =   10
            Top             =   315
            Width           =   5595
         End
      End
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   570
         Top             =   3390
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
         DefaultExt      =   "doc"
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Left            =   4905
         TabIndex        =   4
         Top             =   4320
         Width           =   975
      End
      Begin VB.CommandButton cmdAceptar 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   3720
         TabIndex        =   3
         Top             =   4320
         Width           =   975
      End
      Begin MSComctlLib.ProgressBar Pb1 
         Height          =   255
         Left            =   225
         TabIndex        =   6
         Top             =   3105
         Width           =   6165
         _ExtentX        =   10874
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.Label lblProgres 
         Height          =   375
         Index           =   0
         Left            =   225
         TabIndex        =   8
         Top             =   3480
         Width           =   6195
      End
      Begin VB.Label lblProgres 
         Height          =   285
         Index           =   1
         Left            =   225
         TabIndex        =   7
         Top             =   3960
         Width           =   6195
      End
   End
End
Attribute VB_Name = "frmTrasFacAce"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' PROGRAMA DE TRASPASO DE FACTURAS DE GASOLINERA
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
Private WithEvents frmSec As frmManSecciones 'secciones
Attribute frmSec.VB_VarHelpID = -1
Private WithEvents frmCon As frmManConceptos 'conceptos de las secciones
Attribute frmCon.VB_VarHelpID = -1
Private WithEvents frmFPag As frmForpaConta 'formas de pago de contabilidad
Attribute frmFPag.VB_VarHelpID = -1


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
Dim Cad As String
Dim cadTABLA As String

Dim vContad As Long

Dim PrimeraVez As Boolean
Dim BdConta As Integer

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
Dim b As Boolean
Dim NomFic As String
Dim Cadena As String
Dim cadena1 As String
Dim Acabar As Boolean

On Error GoTo eError
    
    If DatosOk Then
        Me.CommonDialog1.InitDir = App.path
        Me.CommonDialog1.DefaultExt = "txt"
        CommonDialog1.Filter = "Archivos TXT|aceite.txt|"
        CommonDialog1.FilterIndex = 1

        CommonDialog1.CancelError = True
        Me.CommonDialog1.ShowOpen

        Label1.Caption = Me.CommonDialog1.FileName

        Me.Frame1.visible = True
    End If


    Acabar = False
    If Me.CommonDialog1.FileName <> "" Then
        InicializarTabla
    
        If Not Acabar Then
            InicializarVbles
                '========= PARAMETROS  =============================
            'Añadir el parametro de Empresa
            cadParam = cadParam & "|pEmpresa=""" & vEmpresa.nomEmpre & """|"
            numParam = numParam + 1
            
            If ProcesarFichero2(Me.CommonDialog1.FileName) Then
                    cadTABLA = "tmpinformes"
                    cadFormula = "{tmpinformes.codusu} = " & vSesion.Codigo
                    
                    SQL = "select count(*) from tmpinformes where codusu = " & vSesion.Codigo
                    
                    If TotalRegistros(SQL) <> 0 Then
                        MsgBox "Hay errores en la Importación de Facturas de Aceite. Debe corregirlos previamente.", vbExclamation
                        cadTitulo = "Errores de Importación"
                        cadNombreRPT = "rErroresImpGas.rpt"
                        LlamarImprimir
                        Exit Sub
                    Else
                        conn.BeginTrans
                        b = ProcesarFichero(Me.CommonDialog1.FileName)
                    End If
            
            Else
                Exit Sub
            End If
        End If
    Else
        MsgBox "No ha seleccionado ningún fichero", vbExclamation
        Exit Sub
    End If
             
eError:
    If Err.Number = cdlCancel Then
        Unload Me
    Else
        If Err.Number <> 0 Or Not b Then
            conn.RollbackTrans
            MsgBox "No se ha podido realizar el proceso. LLame a Ariadna.", vbExclamation
        Else
            conn.CommitTrans
            MsgBox "Proceso realizado correctamente.", vbExclamation
            Pb1.visible = False
            lblProgres(0).Caption = ""
            lblProgres(1).Caption = ""
            BorrarArchivo Me.CommonDialog1.FileName
            cmdCancel_Click
        End If
    End If
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
    
    On Error GoTo eError

    If PrimeraVez Then
        PrimeraVez = False
        
'        Me.CommonDialog1.InitDir = App.path
'        Me.CommonDialog1.DefaultExt = "txt"
'        CommonDialog1.Filter = "Archivos TXT|cabecera.txt|"
'        CommonDialog1.FilterIndex = 1
'
'        CommonDialog1.CancelError = True
'        Me.CommonDialog1.ShowOpen
'
'        Label1.Caption = Me.CommonDialog1.FileName
'
'        Me.Frame1.visible = True
    End If
    
    Screen.MousePointer = vbDefault
    
eError:
    If Err.Number <> 0 Then Unload Me
End Sub

Private Sub Form_Load()
Dim h As Integer, w As Integer
Dim List As Collection


    PrimeraVez = True
    Limpiar Me


    '###Descomentar
'    CommitConexion
         
    FrameCobrosVisible True, h, w
    Pb1.visible = False
    
    'Esto se consigue poneinedo el cancel en el opcion k corresponda
    Me.cmdCancel.Cancel = True
'   Me.Width = w + 70
'   Me.Height = h + 350

    For h = 0 To 2
        Me.imgBuscar(h).Picture = frmPpal.imgListImages16.ListImages(1).Picture
    Next h
    
    PonerFoco txtCodigo(1)
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    DesBloqueoManual ("TRASGAS")
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

Private Function RecuperaFichero() As Boolean
Dim NF As Integer

    RecuperaFichero = False
    NF = FreeFile
    Open App.path For Input As #NF ' & "\BV" & Format(CDate(txtcodigo(0).Text), "ddmmyy") & "." & Format(txtcodigo(1).Text, "000") For Input As #NF
    Line Input #NF, Cad
    Close #NF
    If Cad <> "" Then RecuperaFichero = True
    
End Function

Private Function ProcesarFichero(nomFich As String) As Boolean
Dim NF As Long
Dim Cad As String
Dim i As Integer
Dim longitud As Long
Dim Rs As ADODB.Recordset
Dim Rs1 As ADODB.Recordset
Dim NumReg As Long
Dim SQL As String
Dim SQL1 As String
Dim SQL11 As String
Dim sql2 As String
Dim Sql3 As String
Dim total As Long
Dim b As Boolean
Dim Linea As Long
Dim letraser As String
Dim TipoIva As String



    On Error GoTo eProcesarFichero

    ProcesarFichero = False
    
    ' procesamos cabecera
    NF = FreeFile
    
    Open nomFich For Input As #NF
    
    lblProgres(0).Caption = "Procesando Fichero: cabecera.txt"
    
    longitud = FileLen(nomFich)
    
    Pb1.visible = True
    Me.Pb1.Max = longitud + 1
    Me.Refresh
    Me.Pb1.Value = 0
    b = True
    SQL1 = ""
    SQL11 = ""
    
    Linea = 1
    
    BdConta = 0
    BdConta = DevuelveDesdeBDNew(cPTours, "seccion", "numconta", "codsecci", txtCodigo(1).Text, "N")
    letraser = ""
    letraser = DevuelveDesdeBDNew(cPTours, "seccion", "letraser", "codsecci", txtCodigo(1).Text, "N")
    TipoIva = ""
    TipoIva = DevuelveDesdeBDNew(cPTours, "concefact", "tipoiva", "codconce", txtCodigo(0).Text, "N")
    
    If AbrirConexionContaFac(vParamAplic.UsuarioContaFac, vParamAplic.PasswordContaFac, CByte(BdConta)) Then
        Set vEmpresaFac = New CempresaFac
        If vEmpresaFac.LeerNiveles Then
            
    
            ' PROCESO DEL FICHERO
            While Not EOF(NF) And b
                Line Input #NF, Cad
                Me.Pb1.Value = Me.Pb1.Value + Len(Cad)
                
                lblProgres(1).Caption = "Linea " & Linea
                Me.Refresh
                sql2 = ""
                Sql3 = ""
                If Cad <> "" Then
                    b = InsertarLinea(Cad, sql2, Sql3)
                    If b Then
                        SQL1 = SQL1 & sql2
                        SQL11 = SQL11 & Sql3
                    End If
                End If
                
            Wend
            Close #NF
        
        
            If b Then
                ' quitamos la ultima coma para hacer el insert
                SQL1 = Mid(SQL1, 1, Len(SQL1) - 1)
                
                SQL = "insert into cabfact (codsecci,letraser,numfactu,fecfactu,ctaclien,observac,intconta,baseiva1,baseiva2,baseiva3,"
                SQL = SQL & "impoiva1,impoiva2,impoiva3,imporec1,imporec2,imporec3,totalfac,tipoiva1,tipoiva2,tipoiva3,"
                SQL = SQL & "porciva1,porciva2,porciva3,codforpa,porcrec1,porcrec2,porcrec3,retfaccl,trefaccl,cuereten) values "
                SQL = SQL & SQL1
                
                conn.Execute SQL
                
                ' quitamos la ultima coma para hacer el insert
                SQL11 = Mid(SQL11, 1, Len(SQL11) - 1)
                
                SQL = "insert into linfact (codsecci,letraser,numfactu,fecfactu,numlinea,codconce,ampliaci,importe,tipoiva) values "
                SQL = SQL & SQL11
                
                conn.Execute SQL
            End If
        End If
        
        Set vEmpresaFac = Nothing
        CerrarConexionContaFac
        
    End If
    ProcesarFichero = b
    
    Pb1.visible = False
    lblProgres(0).Caption = ""
    lblProgres(1).Caption = ""

eProcesarFichero:
    If Err.Number <> 0 Then
        MuestraError Err.Number, "Error en Procesar fichero", Err.Description
    Else
        If Not b Then
            MuestraError 1, "Error en Procesar fichero: ", sql2
        End If
    End If

End Function
                
Private Function ProcesarFichero2(nomFich As String) As Boolean
Dim NF As Long
Dim Cad As String
Dim i As Integer
Dim longitud As Long
Dim Rs As ADODB.Recordset
Dim SQL As String
Dim sql2 As String
Dim Nifsocio As String
Dim CtaSocio As String
Dim Mens As String
Dim Mens1 As String
Dim Linea As Integer

Dim tipo As String
Dim numfactu As String
Dim fecfactu As String
Dim Socio As String
Dim Cuenta As String
Dim FormatSocio As String
Dim Base As String
Dim Iva As String
Dim total As String
Dim PorcIva As Currency
Dim vIva As Currency
Dim vBase As Currency
Dim vNumFac As Currency

Dim Codmacta As String
Dim CadenaCeros As String

Dim IvaCalc As Currency

Dim vCodsocio As String
Dim Codsocio As Long
Dim RaizCta As String

Dim Reten As String

Dim vFecfactu As String
Dim vEmpresaFac As CempresaFac
Dim BdConta As String
Dim ampliacion As String
Dim letraser As String
Dim TipoIva As String


    On Error GoTo eProcesarFichero2
    
    ProcesarFichero2 = False
    NF = FreeFile
    Open nomFich For Input As #NF
    
    i = 0
    
    lblProgres(0).Caption = "Comprobando Facturas: " & nomFich
    longitud = FileLen(nomFich)
    
    Pb1.visible = True
    Me.Pb1.Max = longitud + 1
    Me.Refresh
    Me.Pb1.Value = 0
    
    Linea = 0
    
    BdConta = 0
    BdConta = DevuelveDesdeBDNew(cPTours, "seccion", "numconta", "codsecci", txtCodigo(1).Text, "N")
    letraser = ""
    letraser = DevuelveDesdeBDNew(cPTours, "seccion", "letraser", "codsecci", txtCodigo(1).Text, "N")
    TipoIva = ""
    TipoIva = DevuelveDesdeBDNew(cPTours, "concefact", "tipoiva", "codconce", txtCodigo(0).Text, "N")
    
    If AbrirConexionContaFac(vParamAplic.UsuarioContaFac, vParamAplic.PasswordContaFac, CByte(BdConta)) Then
        Set vEmpresaFac = New CempresaFac
        If vEmpresaFac.LeerNiveles Then
            
            ' PROCESO DEL FICHERO
            While Not EOF(NF)
                Line Input #NF, Cad
                Me.Pb1.Value = Me.Pb1.Value + Len(Cad)
                
                Linea = Linea + 1
                
                If Cad <> "" Then
                    lblProgres(1).Caption = "Linea " & Linea
                    Me.Refresh
                    
                    numfactu = Mid(Cad, 4, 7)
                    vFecfactu = Mid(Cad, 12, 8)
                    fecfactu = CDate(Mid(vFecfactu, 1, 2) & "/" & Mid(vFecfactu, 3, 2) & "/" & Mid(vFecfactu, 5, 4))
                    
                    If Mid(Cad, 22, 1) = "0" Then
                        vCodsocio = Mid(Cad, 24, 3)
                    Else
                        vCodsocio = Mid(Cad, 23, 4)
                    End If
                    Codsocio = CCur(vCodsocio)
'                    If Codsocio >= 1000 Then Codsocio = Codsocio - 1000
            
                    RaizCta = DevuelveDesdeBDNew(cPTours, "seccion", "raizcta", "codsecci", txtCodigo(1).Text, "N")
                    
                    CadenaCeros = Repeat("0", vEmpresaFac.DigitosUltimoNivel - vEmpresaFac.DigitosNivelAnterior)
                    Codmacta = RaizCta & Format(Codsocio, CadenaCeros)
                    
                    ampliacion = Mid(Cad, 34, 15) & " " & Mid(Cad, 83, 9) & " litros"
                    
                    Base = Mid(Cad, 50, 10)
                    Iva = Mid(Cad, 61, 10)
                    total = Mid(Cad, 72, 10)
                    
                    'comprobamos que la cuenta contable del socio sea correcta
                    sql2 = "select codmacta from cuentas where codmacta = " & DBSet(Codmacta, "T")
                    
                    Set Rs = New ADODB.Recordset
                    Rs.Open sql2, ConnContaFac, adOpenForwardOnly, adLockPessimistic, adCmdText
                    
                    If Not Rs.EOF Then
                        Codmacta = DBLet(Rs!Codmacta, "T")
                    Else
                        Mens = "Cuenta contable del socio inexistente "
                        SQL = "insert into tmpinformes (codusu, nombre1, nombre2) values (" & _
                              vSesion.Codigo & "," & DBSet(Codmacta, "T") & "," & DBSet(Mens, "T") & ")"
                              
                        conn.Execute SQL
                    End If
                   
                    Set Rs = Nothing
                    
                    
                    ' comprobamos que la factura no exista ya en la tabla
                    SQL = ""
                    SQL = "select count(*) from cabfact where codsecci = " & DBSet(txtCodigo(1).Text, "N") & " and "
                    SQL = SQL & " letraser = " & DBSet(letraser, "T") & " and numfactu = " & DBSet(numfactu, "N") & " and "
                    SQL = SQL & " fecfactu = " & DBSet(fecfactu, "F")
                    
                    
                    If TotalRegistros(SQL) <> 0 Then
                        Mens = "Ya existe esta factura "
                        SQL = "insert into tmpinformes (codusu, nombre1, nombre2) values (" & _
                              vSesion.Codigo & "," & DBSet(letraser & " " & numfactu & " " & fecfactu & " " & Codmacta, "T") & "," & DBSet(Mens, "T") & ")"
                              
                        conn.Execute SQL
                    End If
                    
                    ' comprobamos que los valores se correspondan con el iva introducido
                    '[Monica]11/07/2012: antes cContaFacSoc
                    PorcIva = DevuelveDesdeBDNewFac("tiposiva", "porceiva", "codigiva", TipoIva, "N")
                    IvaCalc = Round2(CCur(Base) * PorcIva / 100, 2)
                    
                    If Round2((CCur(Base) + CCur(Iva)), 2) <> Round2((CCur(Base) + IvaCalc), 2) Then
                        Mens = "Iva distinto al introducido"
                        Mens1 = Base & " + " & Iva & " <> " & Base & " + " & IvaCalc
                        SQL = "insert into tmpinformes (codusu, nombre1, nombre2) values (" & _
                              vSesion.Codigo & "," & DBSet(Mens1, "T") & "," & DBSet(Mens, "T") & ")"
                              
                        conn.Execute SQL
                    End If
                    
                    ' comprobamos que los importes introducidos sean correctos
                    If (CCur(Base) + CCur(Iva)) <> CCur(total) Then
                        Mens = "Total distinto a Base + Iva "
                        Mens1 = DBSet(Base, "N") & " + " & DBSet(Iva, "N") & " = " & DBSet(total, "N")
                        SQL = "insert into tmpinformes (codusu, nombre1, nombre2) values (" & _
                              vSesion.Codigo & "," & DBSet(Mens1, "T") & "," & DBSet(Mens, "T") & ")"
                              
                        conn.Execute SQL
                    End If
                    
                End If
                
            Wend

        End If
        Set vEmpresaFac = Nothing
        CerrarConexionContaFac
    End If
    
    
    
    Close #NF
    
    If Linea > 0 Then ProcesarFichero2 = True
    
    Pb1.visible = False
    lblProgres(0).Caption = ""
    lblProgres(1).Caption = ""

eProcesarFichero2:
    If Err.Number <> 0 Then
        Set vEmpresaFac = Nothing
        CerrarConexionContaFac
        MuestraError Err.Number, "Error en el proceso de comprobación. Llame a Ariadna.", vbExclamation
    End If
End Function
            
Private Function InsertarLinea(Cad As String, ByRef Result As String, ByRef Result2 As String) As Boolean
Dim b As Boolean
Dim SQL As String
Dim sql2 As String
Dim registro As String

Dim numfactu As String
Dim fecfactu As String
Dim Nifsocio As String
Dim Base As String
Dim Iva As String
Dim total As String
Dim Codmacta As String
Dim CodArtic As String
Dim NomArtic As String
Dim NomSocio As String
Dim cantidad As String
Dim PrecioVe As String
Dim ImpLinea As String
Dim NumLinea As Long
Dim vNumFac As Currency
Dim PorcIva As String

Dim Rs As ADODB.Recordset
Dim tipo As String
Dim vCodsocio As String
Dim Codsocio As Long
Dim CadenaCeros As String
Dim vProd As String
Dim vVar As String
Dim Variedad As String
Dim BaseReten As Currency
Dim ImpReten As String
Dim vFecfactu As String
Dim letraser As String
Dim TipoIva As String
Dim RaizCta As String
Dim ampliacion As String



    On Error GoTo eInsertarLinea

    letraser = ""
    letraser = DevuelveDesdeBDNew(cPTours, "seccion", "letraser", "codsecci", txtCodigo(1).Text, "N")
    TipoIva = ""
    TipoIva = DevuelveDesdeBDNew(cPTours, "concefact", "tipoiva", "codconce", txtCodigo(0).Text, "N")
    PorcIva = ""
    PorcIva = DevuelveDesdeBDNewFac("tiposiva", "porceiva", "codigiva", TipoIva, "N")
    
    numfactu = Mid(Cad, 4, 7)
    vFecfactu = Mid(Cad, 12, 8)
    fecfactu = CDate(Mid(vFecfactu, 1, 2) & "/" & Mid(vFecfactu, 3, 2) & "/" & Mid(vFecfactu, 5, 4))
    
    If Mid(Cad, 22, 1) = "0" Then
        vCodsocio = Mid(Cad, 24, 3)
    Else
        vCodsocio = Mid(Cad, 23, 4)
    End If
    Codsocio = CCur(vCodsocio)
'    If Codsocio >= 1000 Then Codsocio = Codsocio - 1000

    RaizCta = DevuelveDesdeBDNew(cPTours, "seccion", "raizcta", "codsecci", txtCodigo(1).Text, "N")
    
    CadenaCeros = Repeat("0", vEmpresaFac.DigitosUltimoNivel - vEmpresaFac.DigitosNivelAnterior)
    Codmacta = RaizCta & Format(Codsocio, CadenaCeros)
    
    ampliacion = Mid(Cad, 34, 15) & " " & Mid(Cad, 83, 9) & " litros"
    
    Base = Mid(Cad, 50, 10)
    Iva = Mid(Cad, 61, 10)
    total = Mid(Cad, 72, 10)
    
'   +++++CABFACT+++++
'   (codsecci,letraser,numfactu,fecfactu,ctaclien,observac,intconta,baseiva1,baseiva2,baseiva3,
'   impoiva1,impoiva2,impoiva3,impoiva3,imporec1,imporec2,imporec3,totalfac,tipoiva1,tipoiva2,tipoiva3,
'   porciva1,porciva2,porciva3,codforpa,porcrec1,porcrec2,porcrec3,retfaccl,trefaccl,cuereten)
    SQL = DBSet(txtCodigo(1).Text, "N") & ","
    SQL = SQL & DBSet(letraser, "T") & ","
    SQL = SQL & DBSet(numfactu, "N") & ","
    SQL = SQL & DBSet(fecfactu, "F") & ","
    SQL = SQL & DBSet(Codmacta, "T") & ","
    SQL = SQL & ValorNulo & "," 'observaciones
    SQL = SQL & "0," ' intconta
    SQL = SQL & DBSet(ImporteSinFormato(Base), "N") & "," ' baseiva1
    SQL = SQL & ValorNulo & "," ' baseiva2
    SQL = SQL & ValorNulo & "," ' baseiva3
    SQL = SQL & DBSet(ImporteSinFormato(Iva), "N") & "," 'impoiva1
    SQL = SQL & ValorNulo & "," ' impoiva2
    SQL = SQL & ValorNulo & "," ' impoiva3
    SQL = SQL & ValorNulo & "," ' imporec1
    SQL = SQL & ValorNulo & "," ' imporec2
    SQL = SQL & ValorNulo & "," ' imporec3
    SQL = SQL & DBSet(ImporteSinFormato(total), "N") & "," 'totalfac
    SQL = SQL & DBSet(TipoIva, "N") & "," ' tipoiva1
    SQL = SQL & ValorNulo & "," 'tipoiva2
    SQL = SQL & ValorNulo & "," 'tipoiva3
    SQL = SQL & DBSet(ImporteSinFormato(PorcIva), "N") & "," 'porceiva1
    SQL = SQL & ValorNulo & "," 'porceiva2
    SQL = SQL & ValorNulo & "," 'porceiva3
    SQL = SQL & DBSet(txtCodigo(2).Text, "N") & "," ' forma de pago
    SQL = SQL & ValorNulo & "," ' porcrec1
    SQL = SQL & ValorNulo & "," ' porcrec2
    SQL = SQL & ValorNulo & "," ' porcrec3
    SQL = SQL & ValorNulo & "," ' retfaccl
    SQL = SQL & ValorNulo & "," ' trefaccl
    SQL = SQL & ValorNulo  ' cuereten
   
    Result = "(" & SQL & "),"
     
    
'   +++++LINFACT+++++
'(codsecci,letraser,numfactu,fecfactu,numlinea,codconce,ampliaci,importe,tipoiva)
    SQL = DBSet(txtCodigo(1).Text, "N") & ","
    SQL = SQL & DBSet(letraser, "T") & ","
    SQL = SQL & DBSet(numfactu, "N") & ","
    SQL = SQL & DBSet(fecfactu, "F") & ","
    SQL = SQL & "1,"
    SQL = SQL & DBSet(txtCodigo(0), "N") & ","
    SQL = SQL & DBSet(ampliacion, "T") & ","
    SQL = SQL & DBSet(ImporteSinFormato(Base), "N") & ","
    SQL = SQL & DBSet(TipoIva, "N")
    
    Result2 = "(" & SQL & "),"
    InsertarLinea = True
    Exit Function
    
eInsertarLinea:
    If Err.Number <> 0 Then
        InsertarLinea = False
        Result = Err.Description
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
        .Opcion = 0
        .Show vbModal
    End With
End Sub

Private Sub InicializarTabla()
Dim SQL As String
    SQL = "delete from tmpinformes where codusu = " & vSesion.Codigo
    
    conn.Execute SQL
End Sub

Private Sub frmCon_DatoSeleccionado(CadenaSeleccion As String)
'Form de Consulta de Conceptos de la seccion
    txtCodigo(indCodigo).Text = RecuperaValor(CadenaSeleccion, 1)
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmFPag_DatoSeleccionado(CadenaSeleccion As String)
    txtCodigo(2).Text = RecuperaValor(CadenaSeleccion, 1) 'codforpa
    txtNombre(2).Text = RecuperaValor(CadenaSeleccion, 2) 'nomforpa
End Sub

Private Sub frmSec_DatoSeleccionado(CadenaSeleccion As String)
'Form de Consulta de Secciones
    txtCodigo(indCodigo).Text = RecuperaValor(CadenaSeleccion, 1)
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub imgBuscar_Click(Index As Integer)
   Select Case Index
        Case 1 ' Seccion
            AbrirFrmSeccion (Index)
        Case 0 ' Concepto Seccion
            AbrirFrmConceptos (Index)
        Case 2 ' Forma de pago
            AbrirFrmFormaPagos (Index)
    End Select
    PonerFoco txtCodigo(indCodigo)

End Sub


Private Sub AbrirFrmSeccion(indice As Integer)
    indCodigo = indice
    Set frmSec = New frmManSecciones
    frmSec.DatosADevolverBusqueda = "0|1|"
    frmSec.CodigoActual = txtCodigo(indCodigo)
    frmSec.Show vbModal
    Set frmSec = Nothing
End Sub

Private Sub AbrirFrmConceptos(indice As Integer)
    indCodigo = indice
    Set frmCon = New frmManConceptos
    frmCon.DatosADevolverBusqueda = "0|1|"
    frmCon.CodigoActual = txtCodigo(indCodigo)
    frmCon.Show vbModal
    Set frmCon = Nothing
End Sub

Private Sub AbrirFrmFormaPagos(indice As Integer)
    indCodigo = indice
    Set frmFPag = New frmForpaConta
    frmFPag.DatosADevolverBusqueda = "0|1|"
    frmFPag.Facturas = True
    frmFPag.Conexion = BdConta
    frmFPag.CodigoActual = txtCodigo(indCodigo)
    frmFPag.Show vbModal
    Set frmFPag = Nothing
End Sub


Private Function DatosOk() As Boolean
Dim b As Boolean
Dim Orden1 As String
Dim Orden2 As String
Dim FFin As Date
Dim Conta1 As String
Dim Conta2 As String
Dim SQL As String

   b = True

   If txtCodigo(1).Text = "" And b Then
        MsgBox "Introduzca la Sección.", vbExclamation
        b = False
        PonerFoco txtCodigo(1)
   End If
   
   'comprobamos el concepto
   If txtCodigo(0).Text = "" And b Then
        MsgBox "Introduzca el concepto de la Sección.", vbExclamation
        b = False
        PonerFoco txtCodigo(0)
   Else
        ' comprobamos que el concepto introducido pertenece a la misma conta que la seccion
        Conta1 = ""
        Conta1 = DevuelveDesdeBDNew(cPTours, "concefact", "numconta", "codconce", txtCodigo(0).Text, "N")
        Conta2 = ""
        Conta2 = DevuelveDesdeBDNew(cPTours, "seccion", "numconta", "codsecci", txtCodigo(1).Text, "N")
        
        If CCur(Conta1) <> CCur(Conta2) Then
            MsgBox "El Concepto no pertenece a la misma contabilidad que la Sección. Reintroduzca.", vbExclamation
            b = False
            PonerFoco txtCodigo(1)
        End If
   End If
   
   'comprobamos la forma de pago
   If BdConta = 0 Then
        MsgBox "No hay ninguna sección introducida.", vbExclamation
        b = False
   Else
        If AbrirConexionContaFac(vParamAplic.UsuarioContaFac, vParamAplic.PasswordContaFac, CByte(BdConta)) Then
            Set vEmpresaFac = New CempresaFac
            If vEmpresaFac.LeerNiveles Then
                SQL = ""
                If vParamAplic.ContabilidadNueva Then
                    SQL = DevuelveDesdeBDNewFac("formapago", "nomforpa", "codforpa", txtCodigo(2).Text, "N")
                Else
                    SQL = DevuelveDesdeBDNewFac("sforpa", "nomforpa", "codforpa", txtCodigo(2).Text, "N")
                End If
                If SQL = "" Then
                    MsgBox "No existe la Forma de Pago. Reintroduzca.", vbExclamation
                    b = False
                    PonerFoco txtCodigo(2)
                End If
            End If
            Set vEmpresaFac = Nothing
            CerrarConexionContaFac
        Else
            b = False
        End If
   End If
   
   
   DatosOk = b
   
End Function

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
            Case 1: KEYBusqueda KeyAscii, 1 'tipo de iva
        End Select
    Else
        KEYpress KeyAscii
    End If

End Sub

Private Sub KEYBusqueda(KeyAscii As Integer, indice As Integer)
    KeyAscii = 0
    imgBuscar_Click (indice)
End Sub


Private Sub txtCodigo_LostFocus(Index As Integer)
Dim Cad As String, cadTipo As String 'tipo cliente
Dim Conta1 As String
Dim Conta2 As String

    'Quitar espacios en blanco por los lados
    txtCodigo(Index).Text = Trim(txtCodigo(Index).Text)
    
    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub

    Select Case Index
        Case 1 'Seccion
            If txtCodigo(Index).Text <> "" Then
                txtNombre(Index).Text = DevuelveDesdeBDNew(cPTours, "seccion", "nomsecci", "codsecci", txtCodigo(1).Text, "N")
                If txtNombre(Index).Text = "" Then
                    MsgBox "El código de Sección no existe. Reintroduzca.", vbExclamation
                Else
                    BdConta = DevuelveDesdeBDNew(cPTours, "seccion", "numconta", "codsecci", txtCodigo(1).Text, "N")
                    If DBLet(BdConta, "N") = 0 Then
                        MsgBox "Esta seccion no está asociada a ninguna contabilidad. Revise.", vbExclamation
                        txtCodigo(Index).Text = ""
                        PonerFoco txtCodigo(Index)
                    End If
                End If
            End If
        
        Case 0 'Concepto
            If txtCodigo(Index).Text <> "" Then
                txtNombre(Index).Text = DevuelveDesdeBDNew(cPTours, "concefact", "nomconce", "codconce", txtCodigo(0).Text, "N")
                If txtNombre(Index).Text = "" Then
                    MsgBox "El código de Concepto no existe. Reintroduzca.", vbExclamation
                Else
                    If txtCodigo(1).Text <> "" Then
                        ' comprobamos que el concepto introducido pertenece a la misma conta que la seccion
                        Conta1 = ""
                        Conta1 = DevuelveDesdeBDNew(cPTours, "concefact", "numconta", "codconce", txtCodigo(0).Text, "N")
                        
                        If CCur(BdConta) <> CCur(Conta1) Then
                            MsgBox "El Concepto no pertenece a la misma contabilidad que la Sección. Reintroduzca.", vbExclamation
                            PonerFoco txtCodigo(0)
                        End If
                    End If
                End If
            End If
        
       Case 2 ' forma de pago
            If txtCodigo(Index).Text = "" Then Exit Sub
            
            If BdConta = 0 Then Exit Sub
            
            If AbrirConexionContaFac(vParamAplic.UsuarioContaFac, vParamAplic.PasswordContaFac, CByte(BdConta)) Then
                Set vEmpresaFac = New CempresaFac
                If vEmpresaFac.LeerNiveles Then
                    If vParamAplic.ContabilidadNueva Then
                        txtNombre(Index).Text = DevuelveDesdeBDNewFac("formapago", "nomforpa", "codforpa", txtCodigo(Index).Text, "N")
                    Else
                        txtNombre(Index).Text = DevuelveDesdeBDNewFac("sforpa", "nomforpa", "codforpa", txtCodigo(Index).Text, "N")
                    End If
                    If txtNombre(Index).Text = "" Then
                        MsgBox "No existe la Forma de Pago. Reintroduzca.", vbExclamation
                        PonerFoco txtCodigo(Index)
                    End If
                End If
                Set vEmpresaFac = Nothing
                CerrarConexionContaFac
            End If
       
       
    End Select
End Sub


