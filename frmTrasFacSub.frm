VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmTrasFacSub 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Importación de Datos de Subvenciones FEDEPROL"
   ClientHeight    =   4890
   ClientLeft      =   45
   ClientTop       =   2430
   ClientWidth     =   6825
   Icon            =   "frmTrasFacSub.frx":0000
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
      TabIndex        =   5
      Top             =   135
      Width           =   6555
      Begin VB.Frame Frame2 
         Height          =   1455
         Left            =   225
         TabIndex        =   11
         Top             =   225
         Width           =   6180
         Begin VB.TextBox txtNombre 
            BackColor       =   &H80000018&
            Enabled         =   0   'False
            Height          =   285
            Index           =   1
            Left            =   2880
            Locked          =   -1  'True
            TabIndex        =   15
            Text            =   "Text5"
            Top             =   945
            Width           =   3180
         End
         Begin VB.TextBox txtCodigo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   1
            Left            =   1800
            MaxLength       =   10
            TabIndex        =   2
            Top             =   945
            Width           =   1050
         End
         Begin VB.TextBox txtNombre 
            BackColor       =   &H80000018&
            Enabled         =   0   'False
            Height          =   285
            Index           =   0
            Left            =   2880
            Locked          =   -1  'True
            TabIndex        =   13
            Text            =   "Text5"
            Top             =   585
            Width           =   3180
         End
         Begin VB.TextBox txtCodigo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   0
            Left            =   1800
            MaxLength       =   10
            TabIndex        =   1
            Top             =   585
            Width           =   1050
         End
         Begin VB.TextBox txtCodigo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   2
            Left            =   1800
            MaxLength       =   10
            TabIndex        =   0
            Tag             =   "Código Postal|T|S|||clientes|codposta|||"
            Top             =   270
            Width           =   1050
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   1
            Left            =   1485
            MouseIcon       =   "frmTrasFacSub.frx":000C
            MousePointer    =   4  'Icon
            ToolTipText     =   "Buscar Tipo Iva"
            Top             =   990
            Width           =   240
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Tipo Iva"
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
            TabIndex        =   16
            Top             =   990
            Width           =   585
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   0
            Left            =   1485
            MouseIcon       =   "frmTrasFacSub.frx":015E
            MousePointer    =   4  'Icon
            ToolTipText     =   "Buscar Variedad"
            Top             =   630
            Width           =   240
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Variedad"
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
            Left            =   270
            TabIndex        =   14
            Top             =   630
            Width           =   630
         End
         Begin VB.Label Label4 
            Caption         =   "Fecha Factura"
            ForeColor       =   &H00972E0B&
            Height          =   255
            Index           =   2
            Left            =   270
            TabIndex        =   12
            Top             =   270
            Width           =   1050
         End
         Begin VB.Image imgFec 
            Height          =   240
            Index           =   2
            Left            =   1485
            Picture         =   "frmTrasFacSub.frx":02B0
            ToolTipText     =   "Buscar fecha"
            Top             =   270
            Width           =   240
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
         Top             =   1845
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
         Top             =   3960
         Width           =   975
      End
      Begin VB.CommandButton cmdAceptar 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   3720
         TabIndex        =   3
         Top             =   3960
         Width           =   975
      End
      Begin MSComctlLib.ProgressBar Pb1 
         Height          =   255
         Left            =   210
         TabIndex        =   6
         Top             =   2790
         Width           =   6165
         _ExtentX        =   10874
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.Label lblProgres 
         Height          =   375
         Index           =   0
         Left            =   180
         TabIndex        =   8
         Top             =   3120
         Width           =   6195
      End
      Begin VB.Label lblProgres 
         Height          =   285
         Index           =   1
         Left            =   180
         TabIndex        =   7
         Top             =   3600
         Width           =   6195
      End
   End
End
Attribute VB_Name = "frmTrasFacSub"
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
Private WithEvents frmVar As frmManVariedad 'variedades
Attribute frmVar.VB_VarHelpID = -1
Private WithEvents frmIva As frmTipIVAConta 'tipos de iva de la contabilidad
Attribute frmIva.VB_VarHelpID = -1


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
Dim SQL As String
Dim i As Byte
Dim cadwhere As String
Dim b As Boolean
Dim nomfic As String
Dim CADENA As String
Dim cadena1 As String
Dim Acabar As Boolean

On Error GoTo eError
    
    If DatosOk Then
        Me.CommonDialog1.InitDir = App.path
        Me.CommonDialog1.DefaultExt = "txt"
        CommonDialog1.Filter = "Archivos TXT|*.txt|"
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
                        MsgBox "Hay errores en la Importación de Facturas de Subvenciones. Debe corregirlos previamente.", vbExclamation
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

     Me.imgBuscar(0).Picture = frmPpal.imgListImages16.ListImages(1).Picture
     Me.imgBuscar(1).Picture = frmPpal.imgListImages16.ListImages(1).Picture


    txtCodigo(2).Text = Format(Now, "dd/mm/yyyy")
    PonerFoco txtCodigo(2)
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

Private Function ProcesarFichero(nomFich As String) As Boolean
Dim NF As Long
Dim cad As String
Dim i As Integer
Dim longitud As Long
Dim Rs As ADODB.Recordset
Dim Rs1 As ADODB.Recordset
Dim NumReg As Long
Dim SQL As String
Dim SQL1 As String
Dim Sql2 As String
Dim Total As Long
Dim b As Boolean
Dim Linea As Long


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
    
    Linea = 1
    
    ' PROCESO DEL FICHERO
    While Not EOF(NF) And b
        Line Input #NF, cad
        Me.Pb1.Value = Me.Pb1.Value + Len(cad)
        
        lblProgres(1).Caption = "Linea " & Linea
        Me.Refresh
        Sql2 = ""
        If cad <> "" Then
            b = InsertarLinea(cad, Sql2)
            If b Then SQL1 = SQL1 & Sql2
        End If
        
    Wend
    Close #NF
    
    If b Then
        ' quitamos la ultima coma para hacer el insert
        SQL1 = Mid(SQL1, 1, Len(SQL1) - 1)
        
        SQL = "insert into factsocio (numfactu,fecfactu,codmacta,codvarie,tiposoci,tipofact,"
        SQL = SQL & "kilosfac,baseimpo,porciva,cuotaiva,tipoiva,basereten,porcreten,impreten,"
        SQL = SQL & "totalfac,intconta) values "
        SQL = SQL & SQL1
        
        conn.Execute SQL
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
            MuestraError 1, "Error en Procesar fichero: ", Sql2
        End If
    End If

End Function
                
Private Function ProcesarFichero2(nomFich As String) As Boolean
Dim NF As Long
Dim cad As String
Dim i As Integer
Dim longitud As Long
Dim Rs As ADODB.Recordset
Dim SQL As String
Dim Sql2 As String
Dim NifSocio As String
Dim CtaSocio As String
Dim Mens As String
Dim Linea As Integer

Dim LetraSer As String
Dim NumFactu As String
Dim FecFactu As String
Dim Socio As String
Dim Cuenta As String
Dim FormatSocio As String
Dim Base As Currency
Dim Iva As Currency
Dim Total As Currency
Dim PorcIva As Currency
Dim vIva As Currency
Dim vBase As Currency
Dim vNumFac As Currency
Dim Codmacta As String


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
    
        If Mid(cad, 1, 1) <> "2" Then Exit Function
        
        Linea = Linea + 1
        
        If cad <> "" Then
            lblProgres(1).Caption = "Linea " & Linea
            Me.Refresh
            
            NumFactu = Mid(cad, 5, 4)
            NifSocio = Mid(cad, 18, 9)
            
            'comprobamos que la cuenta contable del socio sea correcta
            Sql2 = "select codmacta from cuentas where nifdatos = " & DBSet(NifSocio, "T")
            Sql2 = Sql2 & " and mid(codmacta,1," & vEmpresaFacSoc.DigitosNivelAnterior & ") = " & DBSet(vParamAplic.RaizCtaFacSoc, "T")
            
            Set Rs = New ADODB.Recordset
            Rs.Open Sql2, ConnContaFacSoc, adOpenForwardOnly, adLockPessimistic, adCmdText
            
            If Not Rs.EOF Then
                Codmacta = DBLet(Rs!Codmacta, "T")
            Else
                Codmacta = ""
                Mens = "Cuenta contable del socio incorrecta "
                SQL = "insert into tmpinformes (codusu, nombre1, nombre2) values (" & _
                      vSesion.Codigo & "," & DBSet("Fact:" & NumFactu, "T") & "," & DBSet(Mens, "T") & ")"
                      
                conn.Execute SQL
            End If
           
            Set Rs = Nothing
            
            ' comprobamos que la factura no exista ya en la tabla
            SQL = ""
            SQL = DevuelveDesdeBDNew(cPTours, "factsocio", "numfactu", "numfactu", NumFactu, "N", , "fecfactu", txtCodigo(2).Text, "F", "codmacta", Codmacta, "T")
            If SQL <> "" Then
                Mens = "Ya existe esta factura "
                SQL = "insert into tmpinformes (codusu, nombre1, nombre2) values (" & _
                      vSesion.Codigo & "," & DBSet(NumFactu & " " & txtCodigo(2).Text & " " & Codmacta, "T") & "," & DBSet(Mens, "T") & ")"
                      
                conn.Execute SQL
            End If
            
        End If
        
    Wend
    Close #NF
    
    If Linea > 0 Then ProcesarFichero2 = True
    
    Pb1.visible = False
    lblProgres(0).Caption = ""
    lblProgres(1).Caption = ""

eProcesarFichero2:
    If Err.Number <> 0 Then
        MuestraError Err.Number, "Error en el proceso de comprobación. Llame a Ariadna.", vbExclamation
    End If
End Function
            
Private Function InsertarLinea(cad As String, ByRef Result As String) As Boolean
Dim b As Boolean
Dim SQL As String
Dim Sql2 As String
Dim registro As String

Dim NumFactu As String
Dim FecFactu As String
Dim NifSocio As String
Dim Base As String
Dim Iva As String
Dim Total As String
Dim Codmacta As String
Dim CodArtic As String
Dim NomArtic As String
Dim NomSocio As String
Dim cantidad As String
Dim PrecioVe As String
Dim ImpLinea As String
Dim NumLinea As Long
Dim vNumFac As Currency
Dim PorcIva As Currency
Dim vPorcIva As String

Dim Rs As ADODB.Recordset


    On Error GoTo eInsertarLinea

    
    If Mid(cad, 1, 1) <> "2" Then Exit Function
    
    NumFactu = Mid(cad, 5, 4) ' numero de factura  'RecuperaValor(cad, 2, "|")
    NifSocio = Mid(cad, 18, 9) ' nif del socio
    Total = Mid(cad, 81, 14)
    
    
    Sql2 = "select codmacta from cuentas where nifdatos = " & DBSet(NifSocio, "T")
    Sql2 = Sql2 & " and mid(codmacta,1," & vEmpresaFacSoc.DigitosNivelAnterior & ") = " & vParamAplic.RaizCtaFacSoc
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql2, ConnContaFacSoc, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    If Not Rs.EOF Then
        Codmacta = DBLet(Rs!Codmacta, "T")
    Else
        Codmacta = ""
    End If
    
    'Codmacta = DevuelveDesdeBDNew(cContaFacSoc, "cuentas", "codmacta", "nifdatos", NifSocio, "T")
    
    vPorcIva = ""
    vPorcIva = DevuelveDesdeBDNew(cContaFacSoc, "tiposiva", "porceiva", "codigiva", txtCodigo(1).Text, "N")
    PorcIva = CCur(vPorcIva)
    
    SQL = SQL & DBSet(NumFactu, "N") & ","
    SQL = SQL & DBSet(txtCodigo(2).Text, "F") & ","
    SQL = SQL & DBSet(Codmacta, "T") & ","
    SQL = SQL & DBSet(txtCodigo(0).Text, "N")
    SQL = SQL & ",0,4,0,"
    
    Total = CCur(Mid(cad, 81, 12)) + CCur("0," & Mid(cad, 93, 2))
    Base = Round2(Total / (1 + (PorcIva / 100)), 2)
    Iva = Total - Base
    
    
    SQL = SQL & DBSet(ImporteSinFormato(Base), "N") & ","
    SQL = SQL & DBSet(ImporteSinFormato(vPorcIva), "N") & ","
    SQL = SQL & DBSet(ImporteSinFormato(Iva), "N") & ","
    SQL = SQL & DBSet(txtCodigo(1).Text, "N") & ","
    SQL = SQL & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," ' retenciones
    SQL = SQL & DBSet(ImporteSinFormato(Total), "N")
    SQL = SQL & ",0" ' intconta = 0
    
    
    Result = "(" & SQL & "),"
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
        .EnvioEMail = False
        .Show vbModal
    End With
End Sub

Private Sub InicializarTabla()
Dim SQL As String
    SQL = "delete from tmpinformes where codusu = " & vSesion.Codigo
    
    conn.Execute SQL
End Sub


Private Sub frmC_Selec(vFecha As Date)
 'Fecha
    txtCodigo(CByte(imgFec(2).Tag)).Text = Format(vFecha, "dd/MM/yyyy")
End Sub

Private Sub frmIva_DatoSeleccionado(CadenaSeleccion As String)
'Form de Consulta de Cuentas contables
    txtCodigo(indCodigo).Text = RecuperaValor(CadenaSeleccion, 1)
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmVar_DatoSeleccionado(CadenaSeleccion As String)
'Form de Consulta de Cuentas contables
    txtCodigo(indCodigo).Text = RecuperaValor(CadenaSeleccion, 1)
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub imgBuscar_Click(Index As Integer)
   Select Case Index
        Case 0 ' Variedad
            AbrirFrmVariedad (Index)
        Case 1 ' Tipo de Iva
            AbrirFrmIvaConta (Index)
    End Select
    PonerFoco txtCodigo(indCodigo)

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
    PonerFoco txtCodigo(CByte(imgFec(2).Tag))
    ' ***************************
End Sub

Private Sub AbrirFrmIvaConta(indice As Integer)
    indCodigo = indice
    Set frmIva = New frmTipIVAConta
    frmIva.DatosADevolverBusqueda = "0|1|"
    frmIva.CodigoActual = txtCodigo(indCodigo)
    frmIva.Conexion = cContaFacSoc
    frmIva.Show vbModal
    Set frmIva = Nothing
End Sub

Private Sub AbrirFrmVariedad(indice As Integer)
    indCodigo = indice
    
    Set frmVar = New frmManVariedad
    frmVar.DatosADevolverBusqueda = "0|1|"
    frmVar.CodigoActual = txtCodigo(indCodigo)
    frmVar.Show vbModal
    Set frmVar = Nothing

End Sub


Private Function DatosOk() As Boolean
Dim b As Boolean
Dim Orden1 As String
Dim Orden2 As String
Dim FFin As Date
Dim Cta As String

   b = True

   If txtCodigo(2).Text = "" Then
        MsgBox "Introduzca la Fecha de Factura.", vbExclamation
        b = False
        PonerFoco txtCodigo(2)
   End If
    
   If txtCodigo(0).Text = "" And b Then
        MsgBox "Introduzca la Variedad.", vbExclamation
        b = False
        PonerFoco txtCodigo(0)
   End If
    
   If txtCodigo(1).Text = "" And b Then
        MsgBox "Introduzca el Tipo de Iva.", vbExclamation
        b = False
        PonerFoco txtCodigo(1)
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
            Case 0: KEYBusqueda KeyAscii, 0 'variedad
            Case 1: KEYBusqueda KeyAscii, 1 'tipo de iva
            Case 2: KEYFecha KeyAscii, 2 'fecha factura
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
        Case 2 'FECHAS
            If txtCodigo(Index).Text <> "" Then PonerFormatoFecha txtCodigo(Index)
            
        Case 0 'variedad
            If txtCodigo(Index).Text <> "" Then txtNombre(Index).Text = PonerNombreDeCod(txtCodigo(Index), "variedad", "nomvarie", "codvarie", "N")
        
        Case 1 'tipo de iva
            If txtCodigo(Index).Text <> "" Then txtNombre(Index).Text = DevuelveDesdeBDNew(cContaFacSoc, "tiposiva", "nombriva", "codigiva", txtCodigo(1).Text, "N")
            If txtNombre(Index).Text = "" Then
                MsgBox "Tipo de Iva  no existe en la contabilidad. Reintroduzca.", vbExclamation
            End If
        
    End Select
End Sub

