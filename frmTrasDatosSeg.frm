VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmTrasDatosSeg 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Importación de Datos de Agroweb"
   ClientHeight    =   4890
   ClientLeft      =   45
   ClientTop       =   2430
   ClientWidth     =   6825
   Icon            =   "frmTrasDatosSeg.frx":0000
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
      TabIndex        =   3
      Top             =   135
      Width           =   6555
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   0
         Left            =   3645
         MaxLength       =   10
         TabIndex        =   0
         Top             =   1035
         Width           =   1050
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
         Left            =   270
         TabIndex        =   7
         Top             =   1665
         Width           =   5910
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
         Top             =   3960
         Width           =   975
      End
      Begin VB.CommandButton cmdAceptar 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   3690
         TabIndex        =   1
         Top             =   3960
         Width           =   975
      End
      Begin MSComctlLib.ProgressBar Pb1 
         Height          =   255
         Left            =   210
         TabIndex        =   4
         Top             =   2730
         Width           =   6030
         _ExtentX        =   10636
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.Label Label4 
         Caption         =   "Porcentaje Aumento sobre Precio"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00972E0B&
         Height          =   255
         Index           =   1
         Left            =   405
         TabIndex        =   9
         Top             =   1035
         Width           =   3435
      End
      Begin VB.Label lblProgres 
         Height          =   375
         Index           =   0
         Left            =   180
         TabIndex        =   6
         Top             =   3120
         Width           =   6195
      End
      Begin VB.Label lblProgres 
         Height          =   285
         Index           =   1
         Left            =   180
         TabIndex        =   5
         Top             =   3600
         Width           =   6195
      End
   End
End
Attribute VB_Name = "frmTrasDatosSeg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' PROGRAMA DE TRASPASO DE POSTE PARA ALZICOOP
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
Dim Sql As String
Dim i As Byte
Dim cadwhere As String
Dim b As Boolean
Dim nomfic As String
Dim CADENA As String
Dim cadena1 As String

On Error GoTo eError
    
'    Me.CommonDialog1.InitDir = App.path
'    Me.CommonDialog1.DefaultExt = "csv"
'    CommonDialog1.Filter = "Archivos CSV|seguros.csv|"
'    CommonDialog1.FilterIndex = 1
'
'    Me.CommonDialog1.ShowOpen
    
    
    If txtCodigo(0).Text = "" Then
        MsgBox "Debe introducir un valor en campo Porcentaje de Aumento sobre Precio", vbExclamation
        PonerFoco txtCodigo(0)
        Exit Sub
    End If
    
    Me.CommonDialog1.InitDir = App.path
    Me.CommonDialog1.DefaultExt = "csv"
    CommonDialog1.Filter = "Archivos CSV|*.csv|"
    CommonDialog1.FilterIndex = 1

    CommonDialog1.CancelError = True

    Me.CommonDialog1.ShowOpen

    Label1.Caption = Me.CommonDialog1.FileName
    
    If vParamAplic.RaizCtaSocSeg = "" Then
        MsgBox "Debe introducir la raíz de la Cta.Contable de Socio en parámetros", vbExclamation
        cmdCancel_Click
        Exit Sub
    End If
    
    Me.Frame1.visible = True
    
    If Me.CommonDialog1.FileName <> "" Then
        InicializarVbles
        InicializarTabla
            '========= PARAMETROS  =============================
        'Añadir el parametro de Empresa
        cadParam = cadParam & "|pEmpresa=""" & vEmpresa.nomEmpre & """|"
        numParam = numParam + 1

          
        If ProcesarFichero2(Me.CommonDialog1.FileName) Then
                cadTABLA = "tmpinformes"
                cadFormula = "{tmpinformes.codusu} = " & vSesion.Codigo
                
                Sql = "select count(*) from tmpinformes where codusu = " & vSesion.Codigo
                
                If TotalRegistros(Sql) <> 0 Then
                    MsgBox "Hay errores en la Importación de datos de Agroweb. Debe corregirlos previamente.", vbExclamation
                    cadTitulo = "Errores de Importación"
                    cadNombreRPT = "rErroresImporte.rpt"
                    LlamarImprimir
                    Exit Sub
                Else
                    conn.BeginTrans
                    b = ProcesarFichero(Me.CommonDialog1.FileName)
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
            cmdCancel_Click
        End If
    End If
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
    If PrimeraVez Then
        PrimeraVez = False
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
    Pb1.visible = False
    
    'Esto se consigue poneinedo el cancel en el opcion k corresponda
    Me.cmdCancel.Cancel = True
'   Me.Width = w + 70
'   Me.Height = h + 350

'    Me.CommonDialog1.InitDir = App.path
'    Me.CommonDialog1.DefaultExt = "csv"
'    CommonDialog1.Filter = "Archivos CSV|*.csv|"
'    CommonDialog1.FilterIndex = 1
'
'    Me.CommonDialog1.ShowOpen
'
'    Label1.Caption = Me.CommonDialog1.FileName

    Frame1.visible = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
    DesBloqueoManual ("TRASSEG")
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


Private Function ProcesarFichero(nomfich As String) As Boolean
Dim NF As Long
Dim cad As String
Dim i As Integer
Dim longitud As Long
Dim RS As ADODB.Recordset
Dim Rs1 As ADODB.Recordset
Dim NumReg As Long
Dim Sql As String
Dim SQL1 As String
Dim Total As Long
Dim b As Boolean

    ProcesarFichero = False
    NF = FreeFile
    
    Open nomfich For Input As #NF
    
    lblProgres(0).Caption = "Procesando Fichero: " & nomfich
    longitud = FileLen(nomfich)
    
    Pb1.visible = True
    Me.Pb1.Max = longitud + 1
    Me.Refresh
    Me.Pb1.Value = 0
    b = True
    
    For i = 1 To 2 '[Monica]30/01/2012: antes saltabamos 6 lineas, ahora 1
        Line Input #NF, cad
        Me.Pb1.Value = Me.Pb1.Value + Len(cad)
        Me.Refresh
    Next i
    
    i = 0
    ' PROCESO DEL FICHERO
    While Not EOF(NF) And b
        cad = cad & ";"
        
        i = i + 1
       
        lblProgres(1).Caption = "Linea " & i '[Monica]30/01/2012: antes: & RecuperaValor(cad, 1, ";")
        Me.Refresh
        If Mid(cad, 1, 1) <> ";" Then '[Monica]30/01/2012:antes:  =";"
            b = InsertarLinea(cad)
        End If
        
        Line Input #NF, cad
        Me.Pb1.Value = Me.Pb1.Value + Len(cad) - 1
    Wend
    Close #NF
    
    ProcesarFichero = b
    
    Pb1.visible = False
    lblProgres(0).Caption = ""
    lblProgres(1).Caption = ""

End Function
                
Private Function ProcesarFichero2(nomfich As String) As Boolean
Dim NF As Long
Dim cad As String
Dim i As Integer
Dim longitud As Long
Dim RS As ADODB.Recordset
Dim Sql As String
Dim NifSocio As String
Dim CtaSocio As String
Dim Mens As String

Dim Referencia As String
Dim Plan As String
Dim Linea As String
Dim Sql2 As String

    On Error GoTo eProcesarFichero2
    
    ProcesarFichero2 = False
    NF = FreeFile
    Open nomfich For Input As #NF
    
    i = 0
    
    lblProgres(0).Caption = "Comprobando NIF Socios: " & nomfich
    longitud = FileLen(nomfich)
    
    Pb1.visible = True
    Me.Pb1.Max = longitud + 1
    Me.Refresh
    Me.Pb1.Value = 0
    
    ' saltamos las 6 primeras lineas
    For i = 1 To 2 '[Monica]30/01/2012: antes saltabamos 6 lineas
        Line Input #NF, cad
        Me.Pb1.Value = Me.Pb1.Value + Len(cad)
    Next i
    
    ' PROCESO DEL FICHERO
    i = 0
    While Not EOF(NF)
        i = i + 1
       
        lblProgres(1).Caption = "Linea " & i '[Monica]30/01/2012: antes: & RecuperaValor(cad, 2, ";")
        Me.Refresh
        
        If Mid(cad, 1, 1) <> ";" Then
            '30/01/2012: Cambiamos posiciones
            'NifSocio = RecuperaValor(cad, 6, ";")
            '[Monica]08/01/2013: han cambiado de aseguradora
            'NifSocio = RecuperaValor(cad, 4, ";")
            '[Monica]15/05/2014: han cambiado de aseguradora
            'NifSocio = RecuperaValor(cad, 7, ";")
            '[Monica]25/01/2016: han vuelto a cambiar de aseguradora
            'NifSocio = RecuperaValor(cad, 2, ";")
            NifSocio = RecuperaValor(cad, 7, ";")
            
            'MIRAMOS SI EXISTE EL NIF
            CtaSocio = CtaContableSocio(NifSocio, cContaSeg)

            If CtaSocio = "" Then
                Mens = "Cta.Contable para NIF socio no existe"
                Sql = "insert into tmpinformes (codusu, nombre1, importe1, nombre2) values (" & _
                      vSesion.Codigo & "," & DBSet(NifSocio, "T") & "," & RecuperaValor(cad, 6, ";") & "," & DBSet(Mens, "T") & ")"

                conn.Execute Sql
            End If
            
'            '[Monica]15/05/2014: han cambiado de empresa
'            Referencia = RecuperaValor(cad, 3, ";") ' antes 2
'            Referencia = Replace(Referencia, ".", "") ' me viene ahora con un punto
'            Plan = RecuperaValor(cad, 4, ";")       ' antes 5
'            Linea = RecuperaValor(cad, 5, ";")      ' antes 6
            
            '[Monica]25/01/2016: han cambiado de empresa
            Referencia = RecuperaValor(cad, 2, ";")
            Plan = RecuperaValor(cad, 5, ";")
            Linea = RecuperaValor(cad, 6, ";")
            
            If Trim(Referencia) = "" Or Len(Referencia) = 1 Then
                Mens = "No hay Referencia"
                Sql = "insert into tmpinformes (codusu, nombre1, importe1, nombre2) values (" & _
                      vSesion.Codigo & "," & DBSet(NifSocio, "T") & "," & RecuperaValor(cad, 6, ";") & "," & DBSet(Mens, "T") & ")"

                conn.Execute Sql
            End If
            
            Sql2 = "select count(*) from segpoliza where codrefer = " & DBSet(Referencia, "T") & " and codiplan = " & DBSet(Plan, "N") & " and codlinea = " & DBSet(Linea, "N")
            If TotalRegistros(Sql2) <> 0 Then
                Mens = "Referencia ya existe"
                Sql = "insert into tmpinformes (codusu, nombre1, importe1, nombre2) values (" & _
                      vSesion.Codigo & "," & DBSet(Referencia, "T") & "," & RecuperaValor(cad, 6, ";") & "," & DBSet(Mens, "T") & ")"

                conn.Execute Sql
            End If
        End If
        
        Line Input #NF, cad
        Me.Pb1.Value = Me.Pb1.Value + Len(cad)
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
            
Private Function InsertarLinea(cad As String) As Boolean
Dim b As Boolean
Dim Sql As String
Dim Sql2 As String
Dim registro As String
Dim CtaSocio As String
Dim Plan As String
Dim Linea As String
Dim Colectivo As String
Dim nif As String
Dim Nombre As String
Dim Apellido1 As String
Dim Apellido2 As String
Dim Estado As String
Dim FechaEnvio As String
Dim CCC As String
Dim ImpNeto As String
Dim Referencia As String
Dim Usuario As String
Dim Organizacion As String
Dim NomAsegu As String
Dim Existe As String

Dim NuevoImporte As Currency

    On Error GoTo eInsertarLinea

    InsertarLinea = False
    
'[Monica]30/01/2012: cambiamos posiciones
'    registro = RecuperaValor(cad, 2, ";")
'    Plan = RecuperaValor(cad, 3, ";")
'    Linea = RecuperaValor(cad, 4, ";")
'    Colectivo = RecuperaValor(cad, 5, ";")
'    nif = RecuperaValor(cad, 6, ";")
'    Nombre = RecuperaValor(cad, 7, ";")
'    Apellido1 = RecuperaValor(cad, 8, ";")
'    Apellido2 = RecuperaValor(cad, 9, ";")
'    Estado = RecuperaValor(cad, 10, ";")
'    FechaEnvio = RecuperaValor(cad, 11, ";")
'    CCC = RecuperaValor(cad, 12, ";")
'    ImpNeto = RecuperaValor(cad, 13, ";")
'    Referencia = RecuperaValor(cad, 14, ";")
'    Usuario = RecuperaValor(cad, 15, ";")
'    Organizacion = RecuperaValor(cad, 16, ";")
    
'[Monica]08/01/2013: han cambiado de aseguradora
'    Plan = RecuperaValor(cad, 1, ";")
'    Linea = RecuperaValor(cad, 2, ";")
'    Colectivo = RecuperaValor(cad, 3, ";")
'    nif = RecuperaValor(cad, 4, ";")
'    Nombre = RecuperaValor(cad, 5, ";")
'    Apellido1 = RecuperaValor(cad, 6, ";")
'    Apellido2 = RecuperaValor(cad, 7, ";")
'    Estado = RecuperaValor(cad, 9, ";")
'    FechaEnvio = RecuperaValor(cad, 10, ";")
'    CCC = RecuperaValor(cad, 11, ";")
'    ImpNeto = RecuperaValor(cad, 13, ";")
'    Referencia = RecuperaValor(cad, 14, ";")
'[Monica]08/01/2013: han cambiado de aseguradora
'[Monica]15/05/2014: han cambiado de aseguradora
'    Plan = RecuperaValor(cad, 5, ";")
'    Linea = RecuperaValor(cad, 6, ";")
'    Colectivo = RecuperaValor(cad, 3, ";")
'    nif = RecuperaValor(cad, 7, ";")
'    Nombre = RecuperaValor(cad, 8, ";")
'    Estado = RecuperaValor(cad, 9, ";")
'    FechaEnvio = RecuperaValor(cad, 16, ";")
'    ImpNeto = RecuperaValor(cad, 10, ";")
'    Referencia = RecuperaValor(cad, 2, ";")
'[Monica]15/05/2014: han cambiado de aseguradora
'    Plan = RecuperaValor(cad, 4, ";")
'    Linea = RecuperaValor(cad, 5, ";")
'    Colectivo = RecuperaValor(cad, 8, ";")
'    Colectivo = Replace(Colectivo, ".", "")
'    nif = RecuperaValor(cad, 2, ";")
'    Nombre = RecuperaValor(cad, 1, ";")
'    Estado = RecuperaValor(cad, 12, ";")
'    FechaEnvio = RecuperaValor(cad, 7, ";") ' cogemos la fecha de pago del fichero
'    ImpNeto = RecuperaValor(cad, 21, ";")
'    Referencia = RecuperaValor(cad, 3, ";")
'    Referencia = Replace(Referencia, ".", "")

'[Monica]25/01/2016: han cambiado de aseguradora
    Plan = RecuperaValor(cad, 5, ";")
    Linea = RecuperaValor(cad, 6, ";")
    Colectivo = RecuperaValor(cad, 3, ";")
    nif = RecuperaValor(cad, 7, ";")
    Nombre = RecuperaValor(cad, 8, ";")
    Estado = RecuperaValor(cad, 9, ";")
    FechaEnvio = RecuperaValor(cad, 12, ";") ' cogemos la fecha de pago del fichero
    ImpNeto = RecuperaValor(cad, 10, ";")
    Referencia = RecuperaValor(cad, 2, ";")
    Referencia = Replace(Referencia, ".", "")



    
    'MIRAMOS SI EXISTE EL NIF
    CtaSocio = ""
    CtaSocio = CtaContableSocio(nif, cContaSeg)
    
    Existe = ""
    Existe = DevuelveDesdeBDNew(cPTours, "segpoliza", "codrefer", "codrefer", Referencia, "T", , "codiplan", Plan, "N", "codlinea", Linea, "N")
    If Existe = "" Then
        '[Monica]08/01/2013: han cambiado de aseguradora
        'NomAsegu = Trim(Apellido1) & " " & Trim(Apellido2) & "," & Trim(Nombre)
        NomAsegu = Trim(Nombre)
    
        ' SI NO existe la linea de seguro insertamos en la tabla maestra
        Existe = ""
        Existe = DevuelveDesdeBDNew(cPTours, "seglinea", "nomlinea", "codlinea", Linea, "N")
        If Existe = "" Then
            Sql2 = "insert into seglinea (codlinea, nomlinea) values (" & DBSet(Linea, "N") & ",'AUTOMATICA')"
            conn.Execute Sql2
        End If
    
        NuevoImporte = Round2(ImpNeto * ((ImporteFormateado(txtCodigo(0).Text / 100))), 2)
    
        Sql = "insert into segpoliza(codrefer, codiplan, codlinea, colectiv, codmacta, nifasegu, nomasegu, "
        Sql = Sql & "fechaenv, imppoliz, impinter, impampli, impreduc, intconta, inttesor) values ( "
        Sql = Sql & DBSet(Referencia, "T") & ","
        Sql = Sql & DBSet(Plan, "N") & ","
        Sql = Sql & DBSet(Linea, "N") & ","
        Sql = Sql & DBSet(Colectivo, "N") & ","
        Sql = Sql & DBSet(CtaSocio, "T") & ","
        Sql = Sql & DBSet(nif, "T") & ","
        Sql = Sql & DBSet(NomAsegu, "T") & ","
        Sql = Sql & DBSet(FechaEnvio, "F") & ","
        Sql = Sql & DBSet(ImpNeto, "N") & ","
        Sql = Sql & DBSet(NuevoImporte, "N") & ",0,0,0,0)"
'        Sql = Sql & DBSet(ImpNeto, "N") & ",0,0,0,0)"
        
        conn.Execute Sql
    
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
        .Show vbModal
    End With
End Sub

Private Sub InicializarTabla()
Dim Sql As String
    Sql = "delete from tmpinformes where codusu = " & vSesion.Codigo
    
    conn.Execute Sql
End Sub

Private Sub Txtcodigo_GotFocus(Index As Integer)
    ConseguirFoco txtCodigo(Index), 0
End Sub

Private Sub txtCodigo_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
'Alvançar/Retrocedir els camps en les fleches de desplaçament del teclat.
    KEYdown KeyCode
End Sub

Private Sub txtCodigo_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub txtCodigo_LostFocus(Index As Integer)
Dim cadmen As String

        
    Select Case 0
        Case 0 'Porcentajes
           
            cadmen = TransformaPuntosComas(txtCodigo(Index).Text)
            txtCodigo(Index).Text = Format(cadmen, "##0.00")
    End Select
End Sub

