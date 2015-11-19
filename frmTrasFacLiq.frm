VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmTrasFacLiq 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Importación Liquidaciones"
   ClientHeight    =   4470
   ClientLeft      =   45
   ClientTop       =   2430
   ClientWidth     =   6825
   Icon            =   "frmTrasFacLiq.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4470
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
      Height          =   4305
      Left            =   180
      TabIndex        =   3
      Top             =   90
      Width           =   6555
      Begin VB.Frame Frame2 
         Height          =   1005
         Left            =   225
         TabIndex        =   9
         Top             =   225
         Width           =   6180
         Begin VB.TextBox txtNombre 
            BackColor       =   &H80000018&
            Enabled         =   0   'False
            Height          =   285
            Index           =   1
            Left            =   2430
            Locked          =   -1  'True
            TabIndex        =   10
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
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   1
            Left            =   1035
            MouseIcon       =   "frmTrasFacLiq.frx":000C
            MousePointer    =   4  'Icon
            ToolTipText     =   "Buscar Tipo Iva"
            Top             =   360
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
            TabIndex        =   11
            Top             =   360
            Width           =   585
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
         TabIndex        =   7
         Top             =   1395
         Width           =   6180
         Begin VB.Label Label1 
            Caption         =   "Fichero: "
            Height          =   420
            Left            =   180
            TabIndex        =   8
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
         TabIndex        =   2
         Top             =   3690
         Width           =   975
      End
      Begin VB.CommandButton cmdAceptar 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   3720
         TabIndex        =   1
         Top             =   3690
         Width           =   975
      End
      Begin MSComctlLib.ProgressBar Pb1 
         Height          =   255
         Left            =   225
         TabIndex        =   4
         Top             =   2520
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
         TabIndex        =   6
         Top             =   2895
         Width           =   6195
      End
      Begin VB.Label lblProgres 
         Height          =   285
         Index           =   1
         Left            =   225
         TabIndex        =   5
         Top             =   3375
         Width           =   6195
      End
   End
End
Attribute VB_Name = "frmTrasFacLiq"
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
Dim Sql As String
Dim i As Byte
Dim cadwhere As String
Dim b As Boolean
Dim nomfic As String
Dim CADENA As String
Dim cadena1 As String
Dim Acabar As Boolean

On Error GoTo eError
    
    If DatosOk Then
'        Me.CommonDialog1.InitDir = App.path
'        Me.CommonDialog1.DefaultExt = "txt"
'        CommonDialog1.Filter = "Archivos TXT|*.txt|"
'        CommonDialog1.FilterIndex = 1

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
                    
                    Sql = "select count(*) from tmpinformes where codusu = " & vSesion.Codigo
                    
                    If TotalRegistros(Sql) <> 0 Then
                        MsgBox "Hay errores en la Importación de Facturas de Liquidaciones. Debe corregirlos previamente.", vbExclamation
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
'            BorrarArchivo Me.CommonDialog1.FileName
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

    Me.imgBuscar(1).Picture = frmPpal.imgListImages16.ListImages(1).Picture

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
    Line Input #NF, cad
    Close #NF
    If cad <> "" Then RecuperaFichero = True
    
End Function

Private Function ProcesarFichero(nomFich As String) As Boolean
Dim NF As Long
Dim cad As String
Dim i As Integer
Dim longitud As Long
Dim RS As ADODB.Recordset
Dim Rs1 As ADODB.Recordset
Dim NumReg As Long
Dim Sql As String
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
        
        Sql = "insert into factsocio (numfactu,fecfactu,codmacta,codvarie,tiposoci,tipofact,"
        Sql = Sql & "kilosfac,baseimpo,porciva,cuotaiva,tipoiva,basereten,porcreten,impreten,"
        Sql = Sql & "totalfac,intconta) values "
        Sql = Sql & SQL1
        
        conn.Execute Sql
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
Dim RS As ADODB.Recordset
Dim Sql As String
Dim Sql2 As String
Dim NifSocio As String
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
Dim Total As String
Dim PorcIva As Currency
Dim vIva As Currency
Dim vBase As Currency
Dim vNumFac As Currency

Dim Codmacta As String
Dim CadenaCeros As String

Dim vVar As String
Dim vProd As String
Dim Variedad As String

Dim IvaCalc As Currency

Dim vCodsocio As String
Dim Codsocio As Long
Dim Reten As String

Dim vFecfactu As String
'++monica
Dim vNifSocio As String

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
    
        
        If cad <> "" Then
            If Not (Mid(cad, 1, 1) = "0" Or Mid(cad, 1, 1) = "1") Then Exit Function
            
            Linea = Linea + 1
            
            lblProgres(1).Caption = "Linea " & Linea
            Me.Refresh
            
            tipo = Mid(cad, 1, 1)
            numfactu = RecuperaValor(cad, 2, "|")
            vFecfactu = RecuperaValor(cad, 3, "|")
            fecfactu = CDate(Mid(vFecfactu, 5, 2) & "/" & Mid(vFecfactu, 3, 2) & "/" & Mid(vFecfactu, 1, 2))
'--monica: antes
            vCodsocio = RecuperaValor(cad, 4, "|")
            Codsocio = CCur(vCodsocio) - 1000
'
'            CadenaCeros = Repeat("0", vEmpresaFacSoc.DigitosUltimoNivel - vEmpresaFacSoc.DigitosNivelAnterior)
'            Codmacta = vParamAplic.RaizCtaFacSoc & Format(Codsocio, CadenaCeros)
'++monica: ahora sacamos la cta contable por el nif del socio
            vNifSocio = RecuperaValor(cad, 13, "|") ' sacamos el nif de la 12 posicion
            

            vProd = RecuperaValor(cad, 5, "|")
            vVar = RecuperaValor(cad, 6, "|")
            
            Variedad = Format(CCur(vProd), "00") & Format(CCur(vVar), "00")
            
            Base = RecuperaValor(cad, 8, "|")
            Iva = RecuperaValor(cad, 9, "|")
            Total = RecuperaValor(cad, 10, "|")
            Reten = RecuperaValor(cad, 11, "|")
            
            'comprobamos que la cuenta contable del socio sea correcta
'--monica: antes
'            sql2 = "select codmacta from cuentas where codmacta = " & DBSet(Codmacta, "T")
'++monica: ahora
            Sql2 = "select codmacta from cuentas where nifdatos = " & DBSet(vNifSocio, "T")
            Sql2 = Sql2 & " and codmacta like '" & vParamAplic.RaizCtaFacSoc & "%'"
            
            Set RS = New ADODB.Recordset
            RS.Open Sql2, ConnContaFacSoc, adOpenForwardOnly, adLockPessimistic, adCmdText
            
            If Not RS.EOF Then
                Codmacta = DBLet(RS!Codmacta, "T")
            Else
                Mens = "Cta cble inexistente o raiz incor."
                Sql = "insert into tmpinformes (codusu, nombre1, nombre2) values (" & _
                      vSesion.Codigo & "," & DBSet(vNifSocio, "T") & "," & DBSet(Mens, "T") & ")"
                      
                conn.Execute Sql
            End If
           
            Set RS = Nothing
            
            
            ' comprobamos que la factura no exista ya en la tabla
            If Codmacta <> "" Then '++monica: añadida la condicion
                Sql = ""
                Sql = DevuelveDesdeBDNew(cPTours, "factsocio", "numfactu", "numfactu", numfactu, "N", , "fecfactu", fecfactu, "F", "codmacta", Codmacta, "T")
                If Sql <> "" Then
                    Mens = "Ya existe esta factura "
                    Sql = "insert into tmpinformes (codusu, nombre1, nombre2) values (" & _
                          vSesion.Codigo & "," & DBSet(numfactu & " " & fecfactu & " " & Codmacta, "T") & "," & DBSet(Mens, "T") & ")"
                          
                    conn.Execute Sql
                End If
            End If
            
            ' comprobamos que exista la variedad
            Sql = ""
            Sql = DevuelveDesdeBDNew(cPTours, "variedad", "nomvarie", "codvarie", Variedad, "N")
            If Sql = "" Then
                Mens = "No existe la variedad "
                Sql = "insert into tmpinformes (codusu, nombre1, nombre2) values (" & _
                      vSesion.Codigo & "," & Variedad & "," & DBSet(Mens, "T") & ")"
                      
                conn.Execute Sql
            End If
            
            ' comprobamos que los valores se correspondan con el iva introducido
            PorcIva = DevuelveDesdeBDNew(cContaFacSoc, "tiposiva", "porceiva", "codigiva", txtCodigo(1).Text, "N")
            IvaCalc = Round2(CCur(ImporteSinFormato(Base)) * CCur(PorcIva) / 100, 2)
            
            If (CCur(ImporteSinFormato(Base)) + CCur(ImporteSinFormato(Iva))) <> (CCur(ImporteSinFormato(Base)) + IvaCalc) Then
                Mens = "Iva distinto al introducido"
                Mens1 = DBSet(Base, "N") & " + " & DBSet(Iva, "N") & " <> " & DBSet(Base, "N") & " + " & DBSet(IvaCalc, "N")
                Sql = "insert into tmpinformes (codusu, nombre1, nombre2) values (" & _
                      vSesion.Codigo & "," & DBSet(Mens1, "T") & "," & DBSet(Mens, "T") & ")"
                      
                conn.Execute Sql
            End If
            
            ' comprobamos que los importes introducidos sean correctos
            If (CCur(ImporteSinFormato(Base)) + CCur(ImporteSinFormato(Iva)) - CCur(ImporteSinFormato(Reten))) <> CCur(ImporteSinFormato(Total)) Then
                Mens = "Total distinto a Base + Iva - Reten"
                Mens1 = DBSet(Base, "N") & " + " & DBSet(Iva, "N") & " - " & DBSet(Reten, "N") & " <> " & DBSet(Total, "N")
                Sql = "insert into tmpinformes (codusu, nombre1, nombre2) values (" & _
                      vSesion.Codigo & "," & DBSet(Mens1, "T") & "," & DBSet(Mens, "T") & ")"
                      
                conn.Execute Sql
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
Dim Sql As String
Dim Sql2 As String
Dim registro As String

Dim numfactu As String
Dim fecfactu As String
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

Dim RS As ADODB.Recordset
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
Dim PorcReten As Currency

Dim vNifSocio As String
Dim Sql3 As String

Dim vKilos As String

    On Error GoTo eInsertarLinea

    
    If Not (Mid(cad, 1, 1) = "0" Or Mid(cad, 1, 1) = "1") Then Exit Function
    
        
    tipo = Mid(cad, 1, 1)
    numfactu = RecuperaValor(cad, 2, "|")
    vFecfactu = RecuperaValor(cad, 3, "|")
    fecfactu = CDate(Mid(vFecfactu, 5, 2) & "/" & Mid(vFecfactu, 3, 2) & "/" & Mid(vFecfactu, 1, 2))
    
    
    vCodsocio = RecuperaValor(cad, 4, "|")
    Codsocio = CCur(vCodsocio) - 1000
'--monica: antes
'    CadenaCeros = Repeat("0", vEmpresaFacSoc.DigitosUltimoNivel - vEmpresaFacSoc.DigitosNivelAnterior)
'    Codmacta = vParamAplic.RaizCtaFacSoc & Format(Codsocio, CadenaCeros)
'++monica: ahora
    vNifSocio = RecuperaValor(cad, 13, "|") ' sacamos el nif de la 12 posicion
    
    Sql3 = "select codmacta from cuentas where nifdatos = " & DBSet(vNifSocio, "T")
    Sql3 = Sql3 & " and codmacta like '" & vParamAplic.RaizCtaFacSoc & "%'"
    
    Set RS = New ADODB.Recordset
    RS.Open Sql3, ConnContaFacSoc, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    If Not RS.EOF Then
        Codmacta = DBLet(RS!Codmacta, "T")
    End If
'++monica

    vKilos = RecuperaValor(cad, 14, "|")

    vProd = RecuperaValor(cad, 5, "|")
    vVar = RecuperaValor(cad, 6, "|")
    
    Variedad = Format(CCur(vProd), "00") & Format(CCur(vVar), "00")
    
    Base = RecuperaValor(cad, 8, "|")
    Iva = RecuperaValor(cad, 9, "|")
    Total = RecuperaValor(cad, 10, "|")
    
    
    vPorcIva = ""
    vPorcIva = DevuelveDesdeBDNew(cContaFacSoc, "tiposiva", "porceiva", "codigiva", txtCodigo(1).Text, "N")
    PorcIva = CCur(vPorcIva)
    
    Sql = Sql & DBSet(numfactu, "N") & ","
    Sql = Sql & DBSet(fecfactu, "F") & ","
    Sql = Sql & DBSet(Codmacta, "T") & ","
    Sql = Sql & DBSet(Variedad, "N")
    Sql = Sql & ",0,"                  ' tipo socio
    Sql = Sql & DBSet(tipo, "N") & "," ' tipo factura
    Sql = Sql & DBSet(vKilos, "N") & "," ' kilos
    
    Sql = Sql & DBSet(ImporteSinFormato(Base), "N") & ","
    Sql = Sql & DBSet(ImporteSinFormato(vPorcIva), "N") & ","
    Sql = Sql & DBSet(ImporteSinFormato(Iva), "N") & ","
    Sql = Sql & DBSet(txtCodigo(1).Text, "N") & ","
    
    'calculamos las retenciones
    BaseReten = CCur(ImporteSinFormato(Base)) + CCur(ImporteSinFormato(Iva))
    ImpReten = RecuperaValor(cad, 11, "|")
    
    PorcReten = 0
    If BaseReten <> 0 Then
        PorcReten = Round2(CCur(ImporteSinFormato(ImpReten)) / BaseReten * 100, 2)
    End If
    
    Sql = Sql & DBSet(BaseReten, "N") & ","
    Sql = Sql & DBSet(PorcReten, "N") & ","
    Sql = Sql & DBSet(ImpReten, "N") & ","
    
    Sql = Sql & DBSet(ImporteSinFormato(Total), "N")
    Sql = Sql & ",0" ' intconta = 0
    
    
    Result = "(" & Sql & "),"
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
Dim Sql As String
    Sql = "delete from tmpinformes where codusu = " & vSesion.Codigo
    
    conn.Execute Sql
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
        Case 1 ' Tipo de Iva
            AbrirFrmIvaConta (Index)
    End Select
    PonerFoco txtCodigo(indCodigo)

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


Private Function DatosOk() As Boolean
Dim b As Boolean
Dim Orden1 As String
Dim Orden2 As String
Dim FFin As Date
Dim Cta As String

   b = True

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
Dim cad As String, cadTipo As String 'tipo cliente

    'Quitar espacios en blanco por los lados
    txtCodigo(Index).Text = Trim(txtCodigo(Index).Text)
    
    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub

    Select Case Index
        Case 1 'tipo de iva
            If txtCodigo(Index).Text <> "" Then txtNombre(Index).Text = DevuelveDesdeBDNew(cContaFacSoc, "tiposiva", "nombriva", "codigiva", txtCodigo(1).Text, "N")
            If txtNombre(Index).Text = "" Then
                MsgBox "Tipo de Iva  no existe en la contabilidad. Reintroduzca.", vbExclamation
            End If
        
    End Select
End Sub


