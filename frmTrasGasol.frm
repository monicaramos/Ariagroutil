VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmTrasGasol 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Importación de Datos de Gasolinera"
   ClientHeight    =   4890
   ClientLeft      =   45
   ClientTop       =   2430
   ClientWidth     =   6825
   Icon            =   "frmTrasGasol.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4890
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
      Height          =   4665
      Left            =   150
      TabIndex        =   2
      Top             =   135
      Width           =   6555
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
         Left            =   270
         TabIndex        =   6
         Top             =   1170
         Width           =   5910
         Begin VB.Label Label1 
            Caption         =   "Fichero: "
            Height          =   420
            Left            =   180
            TabIndex        =   7
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
         TabIndex        =   1
         Top             =   3960
         Width           =   975
      End
      Begin VB.CommandButton cmdAceptar 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   3720
         TabIndex        =   0
         Top             =   3960
         Width           =   975
      End
      Begin MSComctlLib.ProgressBar Pb1 
         Height          =   255
         Left            =   210
         TabIndex        =   3
         Top             =   2730
         Width           =   6030
         _ExtentX        =   10636
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.Label lblProgres 
         Height          =   375
         Index           =   0
         Left            =   180
         TabIndex        =   5
         Top             =   3120
         Width           =   6195
      End
      Begin VB.Label lblProgres 
         Height          =   285
         Index           =   1
         Left            =   180
         TabIndex        =   4
         Top             =   3600
         Width           =   6195
      End
   End
End
Attribute VB_Name = "frmTrasGasol"
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
Private WithEvents frmConce As frmConceConta 'conceptos de contabilidad
Attribute frmConce.VB_VarHelpID = -1
Private WithEvents frmTDia As frmDiaConta 'diarios de contabilidad
Attribute frmTDia.VB_VarHelpID = -1


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
Dim cad As String
Dim cadTABLA As String

Dim vContad As Long

Dim PrimeraVez As Boolean

Private Sub KEYpress(KeyAscii As Integer)
    If KeyAscii = 13 Then 'ENTER
        KeyAscii = 0
        SendKeys "{tab}"
    ElseIf KeyAscii = 27 Then Unload Me  'ESC
    End If
End Sub

Private Sub cmdAceptar_Click()
Dim sql As String
Dim i As Byte
Dim cadwhere As String
Dim b As Boolean
Dim NomFic As String
Dim Cadena As String
Dim cadena1 As String
Dim Acabar As Boolean

On Error GoTo eError
    

    Acabar = False
    If Me.CommonDialog1.FileName <> "" Then
        InicializarTabla
        If Dir(Replace(Me.CommonDialog1.FileName, "cabecera", "lineas")) = "" Then
            cad = "No existe el fichero de lineas." & vbCrLf & vbCrLf & "¿ Desea Continuar ?"
            If MsgBox(cad, vbQuestion + vbYesNo + vbDefaultButton1) = vbNo Then
                Acabar = True
            End If
        End If
    
        If Not Acabar Then
            InicializarVbles
                '========= PARAMETROS  =============================
            'Añadir el parametro de Empresa
            cadParam = cadParam & "|pEmpresa=""" & vEmpresa.nomEmpre & """|"
            numParam = numParam + 1
    
              
            If ProcesarFichero2(Me.CommonDialog1.FileName) Then
                    cadTABLA = "tmpinformes"
                    cadFormula = "{tmpinformes.codusu} = " & vSesion.Codigo
                    
                    sql = "select count(*) from tmpinformes where codusu = " & vSesion.Codigo
                    
                    If TotalRegistros(sql) <> 0 Then
                        MsgBox "Hay errores en la Importación de datos de Gasolinera. Debe corregirlos previamente.", vbExclamation
                        cadTitulo = "Errores de Importación"
                        cadNombreRPT = "rErroresImpGas.rpt"
                        LlamarImprimir
                        Exit Sub
                    Else
                        conn.BeginTrans
                        b = ProcesarFichero(Me.CommonDialog1.FileName, 0)
                        If b Then
                            If Dir(Replace(Me.CommonDialog1.FileName, "cabecera", "lineas")) <> "" Then
                                ' si existe el fichero de lineas lo procesamos
                                b = ProcesarFichero(Replace(Me.CommonDialog1.FileName, "cabecera", "lineas"), 1)
                            End If
                        End If
                    End If
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
            BorrarArchivo Replace(LCase(Me.CommonDialog1.FileName), "cabecera", "lineas")
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
        
        Me.CommonDialog1.InitDir = App.path
        Me.CommonDialog1.DefaultExt = "txt"
        CommonDialog1.Filter = "Archivos TXT|cabecera.txt|"
        CommonDialog1.FilterIndex = 1
    
        CommonDialog1.CancelError = True
        Me.CommonDialog1.ShowOpen
        
        Label1.Caption = Me.CommonDialog1.FileName
        
        Me.Frame1.visible = True
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
    Line Input #NF, cad
    Close #NF
    If cad <> "" Then RecuperaFichero = True
    
End Function

Private Function ProcesarFichero(nomFich As String, tipo As Byte) As Boolean
' tipo = 0 cabecera
' tipo = 1 lineas
Dim NF As Long
Dim cad As String
Dim i As Integer
Dim longitud As Long
Dim Rs As ADODB.Recordset
Dim Rs1 As ADODB.Recordset
Dim NumReg As Long
Dim sql As String
Dim SQL1 As String
Dim Total As Long
Dim b As Boolean
Dim Linea As Long

    ProcesarFichero = False
    
    ' procesamos cabecera
    NF = FreeFile
    
    Open nomFich For Input As #NF
    
    If tipo = 0 Then
        lblProgres(0).Caption = "Procesando Fichero: cabecera.txt"
    Else
        lblProgres(0).Caption = "Procesando Fichero: lineas.txt"
    End If
    longitud = FileLen(nomFich)
    
    Pb1.visible = True
    Me.Pb1.Max = longitud + 1
    Me.Refresh
    Me.Pb1.Value = 0
    b = True
    
    Linea = 1
    
    ' PROCESO DEL FICHERO
    While Not EOF(NF) And b
        Line Input #NF, cad
        Me.Pb1.Value = Me.Pb1.Value + Len(cad)
        
        lblProgres(1).Caption = "Linea " & Linea
        Me.Refresh
        If cad <> "" Then
            b = InsertarLinea(cad, tipo)
        End If
        
    Wend
    Close #NF
    
    ProcesarFichero = b
    
    Pb1.visible = False
    lblProgres(0).Caption = ""
    lblProgres(1).Caption = ""

End Function
                
Private Function ProcesarFichero2(nomFich As String) As Boolean
Dim NF As Long
Dim cad As String
Dim i As Integer
Dim longitud As Long
Dim Rs As ADODB.Recordset
Dim sql As String
Dim NifSocio As String
Dim CtaSocio As String
Dim Mens As String
Dim Linea As Integer

Dim letraser As String
Dim numfactu As String
Dim fecfactu As String
Dim socio As String
Dim cuenta As String
Dim FormatSocio As String
Dim Base As Currency
Dim iva As Currency
Dim Total As Currency
Dim PorcIva As Currency
Dim vIva As Currency
Dim vBase As Currency
Dim vNumFac As Currency

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

    ' PROCESO DEL FICHERO
    While Not EOF(NF)
        Line Input #NF, cad
        Me.Pb1.Value = Me.Pb1.Value + Len(cad)
        
        Linea = Linea + 1
        
        If cad <> "" Then
            lblProgres(1).Caption = "Linea " & Linea
            Me.Refresh
            
            '[Monica]02/07/2013: la letra de serie la cogemos del fichero
            letraser = RecuperaValor(cad, 1, "|") 'vParamAplic.NumSerieGas
            
            If letraser = "CB" Then
                letraser = vParamAplic.NumSerieGasB
            Else
                letraser = vParamAplic.NumSerieGas
            End If
            
            numfactu = RecuperaValor(cad, 2, "|")
            fecfactu = RecuperaValor(cad, 3, "|")
            
            vNumFac = CCur(numfactu) + vParamAplic.IncreFactGas
            
            ' comprobamos que la factura no exista ya en la tabla
            sql = ""
            sql = DevuelveDesdeBDNew(cPTours, "gascabfac", "numfactu", "letraser", letraser, "T", , "numfactu", CStr(vNumFac), "N", "fecfactu", fecfactu, "F")
            If sql <> "" Then
                Mens = "Ya existe esta factura "
                sql = "insert into tmpinformes (codusu, nombre1, nombre2) values (" & _
                      vSesion.Codigo & "," & DBSet(letraser & " " & numfactu & " " & fecfactu, "T") & "," & DBSet(Mens, "T") & ")"
                      
                conn.Execute sql
            End If
            
            ' comprobamos que la cuenta contable exista en la contabilidad
            socio = RecuperaValor(cad, 4, "|")
            FormatSocio = Repeat("0", (vEmpresaGas.DigitosUltimoNivel - vEmpresaGas.DigitosNivelAnterior))
            cuenta = Trim(vParamAplic.RaizCtaSocGas & Format(socio, FormatSocio))
            
            sql = ""
            sql = DevuelveDesdeBDNew(cContaGas, "cuentas", "nommacta", "codmacta", cuenta, "T")
            If sql = "" Then
                Mens = "La Cta.Contable del Socio no existe "
                sql = "insert into tmpinformes (codusu, nombre1, nombre2) values (" & _
                      vSesion.Codigo & "," & DBSet(cuenta, "T") & "," & DBSet(Mens, "T") & ")"
                      
                conn.Execute sql
            End If
            
            ' comprobamos que el iva que tenemos en parametros se corresponde
            Base = DBLet(RecuperaValor(cad, 5, "|"), "N")
            iva = DBLet(RecuperaValor(cad, 6, "|"), "N")
            Total = DBLet(RecuperaValor(cad, 7, "|"), "N")
            
            PorcIva = DevuelveDesdeBDNew(cContaGas, "tiposiva", "porceiva", "codigiva", vParamAplic.CodIvaGas, "N")
            vBase = Round2(Total / (1 + (PorcIva / 100)), 2)
            vIva = Round2(Total - vBase, 2)
            If vBase <> Base Or vIva <> iva Then
                Mens = "Los Importes de la Factura son incorrectos "
                sql = "insert into tmpinformes (codusu, nombre1, nombre2) values (" & _
                      vSesion.Codigo & "," & DBSet(numfactu, "T") & "," & DBSet(Mens, "T") & ")"
                      
                conn.Execute sql
            End If
        End If
        
    Wend
    Close #NF
    
    If cad <> "" Then ProcesarFichero2 = True
    
    Pb1.visible = False
    lblProgres(0).Caption = ""
    lblProgres(1).Caption = ""

eProcesarFichero2:
    If Err.Number <> 0 Then
        MsgBox "Se ha producido un error en el proceso de comprobación", vbExclamation
    End If
End Function
            
Private Function InsertarLinea(cad As String, tipo As Byte) As Boolean
Dim b As Boolean
Dim sql As String
Dim Sql2 As String
Dim registro As String

Dim numfactu As String
Dim fecfactu As String
Dim Codsocio As String
Dim Base As String
Dim iva As String
Dim Total As String
Dim FecAlbar As String
Dim codartic As String
Dim NomArtic As String
Dim NomSocio As String
Dim cantidad As String
Dim PrecioVe As String
Dim Implinea As String
Dim NumLinea As Long
Dim vNumFac As Currency
Dim PorcIva As String

Dim letraser As String


    On Error GoTo eInsertarLinea

    InsertarLinea = False
    
    
    If tipo = 0 Then 'cabecera
        '[Monica]02/07/2013: la letra de serie ya no puede ser de parametros
        letraser = RecuperaValor(cad, 1, "|")
        If letraser = "CB" Then
            letraser = vParamAplic.NumSerieGasB
        Else
            letraser = vParamAplic.NumSerieGas
        End If
        
        numfactu = RecuperaValor(cad, 2, "|")
        fecfactu = RecuperaValor(cad, 3, "|")
        Codsocio = RecuperaValor(cad, 4, "|")
        Base = RecuperaValor(cad, 5, "|")
        iva = RecuperaValor(cad, 6, "|")
        Total = RecuperaValor(cad, 7, "|")
        NomSocio = ""
    
        vNumFac = CCur(numfactu) + vParamAplic.IncreFactGas
        
        PorcIva = ""
        PorcIva = DevuelveDesdeBDNew(cContaGas, "tiposiva", "porceiva", "codigiva", vParamAplic.CodIvaGas, "N")
        
        sql = "insert into gascabfac (letraser, numfactu, fecfactu, codsocio, nomsocio, base, iva, total, codiva, porciva, intconta) values ("
        
        '[Monica]02/07/2013: la letra de serie ya no puede ser de parametros (por las de gasoleo bonificado)
        'Sql = Sql & DBSet(Trim(vParamAplic.NumSerieGas), "T") & ","
        sql = sql & DBSet(letraser, "T") & ","
        
        sql = sql & DBSet(vNumFac, "N") & ","
        sql = sql & DBSet(fecfactu, "F") & ","
        sql = sql & DBSet(Codsocio, "N") & ","
        sql = sql & DBSet(NomSocio, "T", "N") & ","
        sql = sql & DBSet(ImporteSinFormato(Base), "N") & ","
        sql = sql & DBSet(ImporteSinFormato(iva), "N") & ","
        sql = sql & DBSet(ImporteSinFormato(Total), "N") & ","
        sql = sql & DBSet(ImporteSinFormato(vParamAplic.CodIvaGas), "N") & ","
        sql = sql & DBSet(ImporteSinFormato(PorcIva), "N") & ","
        sql = sql & "0)"
        
        conn.Execute sql
        
    Else
        '[Monica]02/07/2013: hemos modificado las posiciones pq se ha añadido la letra de serie
        letraser = RecuperaValor(cad, 3, "|")
        If letraser = "CB" Then
            letraser = vParamAplic.NumSerieGasB
        Else
            letraser = vParamAplic.NumSerieGas
        End If
    
        numfactu = RecuperaValor(cad, 4, "|")
        fecfactu = RecuperaValor(cad, 5, "|")
        Codsocio = RecuperaValor(cad, 1, "|")
        NomSocio = RecuperaValor(cad, 2, "|")
        FecAlbar = RecuperaValor(cad, 6, "|")
        codartic = RecuperaValor(cad, 7, "|")
        NomArtic = RecuperaValor(cad, 8, "|")
        cantidad = RecuperaValor(cad, 9, "|")
        PrecioVe = RecuperaValor(cad, 10, "|")
        Implinea = RecuperaValor(cad, 11, "|")
        
        vNumFac = CCur(numfactu) + vParamAplic.IncreFactGas
        
        NumLinea = 1
        '[Monica]02/07/2013: ahora la letra de serie no viene de parametros, viene del fichero
        sql = "select max(numlinea) from gaslinfac where letraser = " & DBSet(letraser, "T")       '& DBSet(Trim(vParamAplic.NumSerieGas), "T")
        sql = sql & " and numfactu = " & DBSet(vNumFac, "N") & " and fecfactu = " & DBSet(fecfactu, "F")
        sql = DevuelveValor(sql)
        If sql <> "" Then NumLinea = CInt(sql) + 1
        
        sql = "insert into gaslinfac (letraser, numfactu, fecfactu, numlinea, codsocio, nomsocio, "
        sql = sql & "fecalbar, codartic, nomartic, cantidad, preciove, implinea) values ("
        '[Monica]02/07/2013: ahora la letra de serie no viene de parametros, viene del fichero
        'Sql = Sql & DBSet(Trim(vParamAplic.NumSerieGas), "T") & ","
        sql = sql & DBSet(Trim(letraser), "T") & ","
        sql = sql & DBSet(vNumFac, "N") & ","
        sql = sql & DBSet(fecfactu, "F") & ","
        sql = sql & DBSet(NumLinea, "N") & ","
        sql = sql & DBSet(Codsocio, "N") & ","
        sql = sql & DBSet(NomSocio, "T") & ","
        sql = sql & DBSet(FecAlbar, "F") & ","
        sql = sql & DBSet(codartic, "N") & ","
        sql = sql & DBSet(NomArtic, "T") & ","
        
        sql = sql & DBSet(ImporteSinFormato(cantidad), "N") & ","
        sql = sql & DBSet(ImporteSinFormato(PrecioVe), "N") & ","
        sql = sql & DBSet(ImporteSinFormato(Implinea), "N") & ")"
        
        conn.Execute sql
        
        sql = "update gascabfac set nomsocio = " & DBSet(NomSocio, "T")
        sql = sql & " where letraser = " & DBSet(Trim(letraser), "T") ' DBSet(Trim(vParamAplic.NumSerieGas), "T")
        sql = sql & " and numfactu = " & DBSet(vNumFac, "N")
        sql = sql & " and fecfactu = " & DBSet(fecfactu, "F")
        
        conn.Execute sql
        
    End If
    
    
    
    InsertarLinea = True

eInsertarLinea:
    If Err.Number <> 0 Then
        MsgBox "Error en Insertar Linea " & Err.Description, vbExclamation
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
        .EnvioEMail = False
        .Show vbModal
    End With
End Sub

Private Sub InicializarTabla()
Dim sql As String
    sql = "delete from tmpinformes where codusu = " & vSesion.Codigo
    
    conn.Execute sql
End Sub

