VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmTrasFacCV 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Traspaso Facturas de Coarval/Varias"
   ClientHeight    =   4890
   ClientLeft      =   45
   ClientTop       =   2430
   ClientWidth     =   6825
   Icon            =   "frmTrasFactCV.frx":0000
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
      TabIndex        =   0
      Top             =   120
      Width           =   6555
      Begin VB.Frame Frame2 
         BorderStyle     =   0  'None
         Height          =   945
         Left            =   120
         TabIndex        =   7
         Top             =   1560
         Width           =   6180
         Begin VB.TextBox txtCodigo 
            Height          =   285
            Index           =   2
            Left            =   1410
            MaxLength       =   10
            TabIndex        =   6
            Top             =   360
            Width           =   405
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   0
            Left            =   1140
            Tag             =   "-1"
            ToolTipText     =   "Buscar Contadores"
            Top             =   390
            Width           =   240
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
            Left            =   270
            TabIndex        =   9
            Top             =   405
            Width           =   780
         End
      End
      Begin VB.ComboBox Combo1 
         Appearance      =   0  'Flat
         Height          =   315
         ItemData        =   "frmTrasFactCV.frx":000C
         Left            =   1560
         List            =   "frmTrasFactCV.frx":000E
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Tag             =   "Tipo Forpa|N|N|0|9|sforpa|tipforpa||N|"
         Top             =   1200
         Width           =   1485
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
         TabIndex        =   10
         Top             =   3780
         Width           =   975
      End
      Begin VB.CommandButton cmdAceptar 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   3690
         TabIndex        =   8
         Top             =   3780
         Width           =   975
      End
      Begin MSComctlLib.ProgressBar Pb1 
         Height          =   255
         Left            =   210
         TabIndex        =   1
         Top             =   2730
         Width           =   6030
         _ExtentX        =   10636
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.Label Label1 
         Caption         =   "Tipo de Factura"
         ForeColor       =   &H00972E0B&
         Height          =   195
         Index           =   67
         Left            =   360
         TabIndex        =   4
         Top             =   1245
         Width           =   1230
      End
      Begin VB.Label lblProgres 
         Height          =   375
         Index           =   0
         Left            =   180
         TabIndex        =   3
         Top             =   3120
         Width           =   6195
      End
      Begin VB.Label lblProgres 
         Height          =   285
         Index           =   1
         Left            =   180
         TabIndex        =   2
         Top             =   3480
         Width           =   6195
      End
   End
End
Attribute VB_Name = "frmTrasFacCV"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' PROGRAMA DE TRASPASO FACTURAS DE TELEFONIA PARA VALSUR
' basado en frmTrasPoste de gasolinera
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
Private WithEvents frmCont As frmContConta 'Contadores de Contabilidad
Attribute frmCont.VB_VarHelpID = -1


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
Dim NomFic As String
Dim CADENA As String
Dim cadena1 As String
Dim SqlError As String
Dim Salir As Boolean

On Error GoTo eError


    If Not DatosOk Then Exit Sub
    
    
    Select Case Combo1.ListIndex
        Case 0 ' facturas varias
            Me.CommonDialog1.DefaultExt = "CSV"
            CommonDialog1.FilterIndex = 1
            Me.CommonDialog1.FileName = "kilo.csv"
        
        Case 1 ' facturas de venta en tienda
            Me.CommonDialog1.DefaultExt = "TXT"
            CommonDialog1.FilterIndex = 1
            Me.CommonDialog1.FileName = "cab.txt"
        
        Case 2 ' facturas de compra
            Me.CommonDialog1.DefaultExt = "TXT"
            CommonDialog1.FilterIndex = 1
            Me.CommonDialog1.FileName = "cabc.txt"
        
    End Select
    
    Me.CommonDialog1.CancelError = True
    Salir = True
    Me.CommonDialog1.ShowOpen
    Salir = False
    
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
                
            SQL = "select count(*) from tmpinformes where codusu = " & vSesion.Codigo
                
            If TotalRegistros(SQL) <> 0 Then
                SqlError = "select distinct importe2 from tmpinformes where codusu = " & DBSet(vSesion.Codigo, "N") & " order by 1"
                MensError = "Hay errores en el Traspaso de Facturas. "
                If DevuelveValor(SqlError) <> 1 Then
                    MensError = MensError & "Debe corregirlos previamente."
                End If
                
                MsgBox MensError, vbExclamation
                
                cadTitulo = "Errores de Traspaso de Facturas"
                cadNombreRPT = "rErroresTrasCV.rpt"
                LlamarImprimir
                SqlError = "select distinct importe2 from tmpinformes where codusu = " & DBSet(vSesion.Codigo, "N") & " order by 1"
                If DevuelveValor(SqlError) = 1 Then
                    If MsgBox("¿ Desea continuar con el proceso de carga ?", vbQuestion + vbYesNo + vbDefaultButton1) = vbYes Then
                        conn.BeginTrans
                        
                        b = ProcesarFichero(Me.CommonDialog1.FileName)
                    Else
                        Exit Sub
                    End If
                Else
                    Exit Sub
                End If
            Else
                conn.BeginTrans
                b = ProcesarFichero(Me.CommonDialog1.FileName)
            End If
        Else
            MsgBox "No se ha procesado ningún fichero. Revise.", vbExclamation
            Exit Sub
        End If
    Else
        MsgBox "No ha seleccionado ningún fichero", vbExclamation
        Exit Sub
    End If
             
             
eError:
    If Salir Then Exit Sub
    If Err.Number <> 0 Or Not b Then
        If Combo1.ListIndex = 1 Then ConnContaCV.RollbackTrans
        conn.RollbackTrans
        MsgBox "No se ha podido realizar el proceso. LLame a Ariadna.", vbExclamation
    Else
        conn.CommitTrans
        If Combo1.ListIndex = 1 Then ConnContaCV.CommitTrans
        MsgBox "Proceso realizado correctamente.", vbExclamation
        Pb1.visible = False
        lblProgres(0).Caption = ""
        lblProgres(1).Caption = ""
'        BorrarArchivo Me.CommonDialog1.FileName
        cmdCancel_Click
    End If
    
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub Combo1_Change()
    Frame2.Enabled = (Combo1.ListIndex = 0)
    Frame2.visible = (Combo1.ListIndex = 0)
    txtCodigo(2).Text = ""
End Sub

Private Sub Combo1_Click()
    Frame2.Enabled = (Combo1.ListIndex = 0)
    Frame2.visible = (Combo1.ListIndex = 0)
    txtCodigo(2).Text = ""
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

    'IMAGES para busqueda
    '###Descomentar
'    CommitConexion
   'cargar IMAGES de busqueda
         
    FrameCobrosVisible True, h, w
    Pb1.visible = False
    
    Me.imgBuscar(0).Picture = frmPpal.imgListImages16.ListImages(1).Picture
    'Esto se consigue poneinedo el cancel en el opcion k corresponda
    
    CargaCombo
    Combo1.ListIndex = 0
    
    Frame2.Enabled = True
    Frame2.visible = True
    txtCodigo(2).Text = vParamAplic.LetraSerCVV
    
    Me.cmdCancel.Cancel = True
'   Me.Width = w + 70
'   Me.Height = h + 350
End Sub

Private Sub Form_Unload(Cancel As Integer)
    DesBloqueoManual ("TRASPOST")
End Sub



Private Sub ListView1_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub


Private Sub Text1_LostFocus(Index As Integer)
Dim cadMen As String
Dim Nuevo As Boolean

    
'    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
'    'mostrar mensajes ni hacer nada
'    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub
'
'    ' ***************** configurar els LostFocus dels camps de la capçalera *****************
'    Select Case Index
'        Case 17 'Letra de serie
'            If Text1(Index).Text <> "" Then Text1(Index).Text = UCase(Text1(Index).Text)
'
'    End Select
        ' ***************************************************************************
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = teclaBuscar Then
        Select Case Index
'            Case 17: KEYBusqueda KeyAscii, 3 'forma pago
        End Select
    Else
        KEYpress KeyAscii
    End If
End Sub

Private Sub Text1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
'Alvançar/Retrocedir els camps en les fleches de desplaçament del teclat.
    KEYdown KeyCode
End Sub

Private Sub KEYBusqueda(KeyAscii As Integer, indice As Integer)
    KeyAscii = 0
'    imgBuscar_Click (indice)
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

 

Private Function DatosOk() As Boolean
Dim b As Boolean
Dim SQL As String
Dim Rs As ADODB.Recordset

   b = True

   Select Case Combo1.ListIndex
        Case 0
            If txtCodigo(2).Text = "" Then
                MsgBox "Debe introducir obligatoriamente una letra de serie. Revise.", vbExclamation
                PonerFoco txtCodigo(2)
                b = False
            Else
                SQL = "select * from contadores where tiporegi = " & DBSet(Trim(txtCodigo(2).Text), "T") & " and  tiporegi >='A' and tiporegi < 'Z' "
                Set Rs = New ADODB.Recordset
                Rs.Open SQL, ConnContaCVV, adOpenForwardOnly, adLockPessimistic, adCmdText
    
                If Rs.EOF Then
'                cad = DevuelveDesdeBDNew(cContaCVV, "contadores", "nomregis", "tiporegi", Trim(txtCodigo(2).Text), "T")
'                If cad = "" Then
                    MsgBox "La letra de serie no existe en contabilidad o no está permitida. Revise.", vbExclamation
                    PonerFoco txtCodigo(2)
                    b = False
                End If
            End If
   End Select
   
   DatosOk = b
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
Dim total As Long
Dim v_cant As Currency
Dim v_impo As Currency
Dim v_prec As Currency
Dim b As Boolean
Dim NomFic As String

    ProcesarFichero = False
    NF = FreeFile
    
    Open nomFich For Input As #NF ' & "\BV" & Format(CDate(txtcodigo(0).Text), "ddmmyy") & "." & Format(txtcodigo(1).Text, "000") For Input As #NF
    
    Line Input #NF, cad
    i = 0
    
    lblProgres(0).Caption = "Procesando Fichero: " & nomFich
    longitud = FileLen(nomFich)
    
    Pb1.visible = True
    Me.Pb1.Max = longitud
    Me.Refresh
    Me.Pb1.Value = 0
        
    b = True
    
    If Combo1.ListIndex = 1 Then b = InsertarLineaTickets()
    
    While Not EOF(NF) And b
        i = i + 1
        
        Me.Pb1.Value = Me.Pb1.Value + Len(cad)
        lblProgres(1).Caption = "Linea " & i
        Me.Refresh
        
        Select Case Combo1.ListIndex
            Case 0
                cad = Replace(cad, Chr(9), "|")
                cad = Replace(cad, ";", "|") & "|"
            
                b = InsertarLineaVarias(cad)
            
            Case 1
                b = InsertarLineaVentas(cad)
            
            Case 2
                b = InsertarLineaCompras(cad)
            
            
        End Select
        
        If b = False Then
            ProcesarFichero = False
            Exit Function
        End If
        
        Line Input #NF, cad
    Wend
    Close #NF
    
    If cad <> "" Then
        Select Case Combo1.ListIndex
            Case 0
                cad = Replace(cad, Chr(9), "|")
                cad = Replace(cad, ";", "|") & "|"
            
                b = InsertarLineaVarias(cad)
            Case 1
                b = InsertarLineaVentas(cad)
            
            Case 2
                b = InsertarLineaCompras(cad)
        End Select

        If b = False Then
            ProcesarFichero = False
            Exit Function
        End If
    End If
    
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
Dim Rs1 As ADODB.Recordset
Dim NumReg As Long
Dim SQL As String
Dim SQL1 As String
Dim total As Long
Dim v_cant As Currency
Dim v_impo As Currency
Dim v_prec As Currency
Dim b As Boolean

    On Error GoTo eProcesarFichero2
    
    ProcesarFichero2 = False
    
    NF = FreeFile
    Open nomFich For Input As #NF ' & "\BV" & Format(CDate(txtcodigo(0).Text), "ddmmyy") & "." & Format(txtcodigo(1).Text, "000") For Input As #NF
    
    Line Input #NF, cad
    i = 0
    
    lblProgres(0).Caption = "Insertando en Tabla temporal: " & nomFich
    longitud = FileLen(nomFich)
    
    Pb1.visible = True
    Me.Pb1.Max = longitud
    Me.Refresh
    Me.Pb1.Value = 0
    ' PROCESO DEL FICHERO VENTAS.TXT

    b = True

    While Not EOF(NF) And b
        i = i + 1
        
        Me.Pb1.Value = Me.Pb1.Value + Len(cad)
        lblProgres(1).Caption = "Linea " & i
        Me.Refresh
'        Cad = Replace(Cad, Chr(9), "|")
        
        Select Case Combo1.ListIndex
            Case 0
                cad = Replace(cad, ";", "|")
                cad = Trim(cad) & "|"
                b = ComprobarRegistroVarias(cad)
            Case 1
                b = ComprobarRegistroVentas(cad)
            Case 2
                b = ComprobarRegistroCompras(cad)
        End Select
        
        Line Input #NF, cad
    Wend
    Close #NF
    
    If cad <> "" And b Then
        i = i + 1
        
        Me.Pb1.Value = Me.Pb1.Value + Len(cad)
        lblProgres(1).Caption = "Linea " & i
        Me.Refresh
        
        Select Case Combo1.ListIndex
            Case 0
                cad = Replace(cad, ";", "|")
                cad = Trim(cad) & "|"
                b = ComprobarRegistroVarias(cad)
            Case 1
                b = ComprobarRegistroVentas(cad)
            Case 2
                b = ComprobarRegistroCompras(cad)
        End Select
        
    End If
    
    Pb1.visible = False
    lblProgres(0).Caption = ""
    lblProgres(1).Caption = ""

    ProcesarFichero2 = b
    Exit Function

eProcesarFichero2:
    ProcesarFichero2 = False
End Function
                

Private Function ComprobarRegistroVarias(cad As String) As Boolean
Dim SQL As String

Dim c_BaseImpo As Currency
Dim c_CuotaIva As Currency
Dim c_TotalFac As Currency

Dim Mens As String

Dim Fecha As String
Dim codsoc As String
Dim numfactu As String
Dim BaseImpo As String
Dim CuotaIva As String
Dim TotalFac As String
Dim Digito As String
Dim CodmactaSoc As String
Dim CodmactaVta As String
Dim longitud As Integer
Dim vCodSoc As String
Dim porceiva As Currency
Dim letraser As String
Dim TipForpa As String
Dim CodForpa As Long
Dim CodIVA As String

    On Error GoTo eComprobarRegistroVarias

    
    ComprobarRegistroVarias = True

    numfactu = RecuperaValor(cad, 1)
    
    If numfactu = "CodiFactura" Then Exit Function
    
    Fecha = RecuperaValor(cad, 2)
    codsoc = RecuperaValor(cad, 3)
    porceiva = Mid(RecuperaValor(cad, 5), 1, Len(RecuperaValor(cad, 5)) - 1)
    CuotaIva = RecuperaValor(cad, 6)
    BaseImpo = RecuperaValor(cad, 7)
    TotalFac = RecuperaValor(cad, 8)
    CodmactaSoc = RecuperaValor(cad, 9)
    CodmactaVta = RecuperaValor(cad, 10)
    letraser = Trim(txtCodigo(2).Text) 'vParamAplic.LetraSerCVV 'Mid(RecuperaValor(cad, 11), 1, 1)
    TipForpa = RecuperaValor(cad, 12)
    If TipForpa = "Remesa" Then
        CodForpa = vParamAplic.CodforpaBanCVV
    Else
        CodForpa = vParamAplic.CodforpaConCVV
    End If
    
    
    c_BaseImpo = CCur(TransformaPuntosComas(ImporteSinFormato(BaseImpo)))
    c_CuotaIva = CCur(TransformaPuntosComas(ImporteSinFormato(CuotaIva)))
    c_TotalFac = CCur(TransformaPuntosComas(ImporteSinFormato(TotalFac)))
    
    'Comprobamos fechas
    If Not EsFechaOK(Fecha) Then
        Mens = "Fecha incorrecta"
        SQL = "insert into tmpinformes (codusu, fecha1, importe1, nombre2, importe3, " & _
              "importe4, importe5, nombre1, importe2) values (" & _
              vSesion.Codigo & "," & DBSet(Fecha, "F") & _
              "," & DBSet(codsoc, "N") & "," & _
              DBSet(numfactu, "N") & "," & _
              DBSet(c_BaseImpo, "N") & "," & _
              DBSet(c_CuotaIva, "N") & "," & _
              DBSet(c_TotalFac, "N") & "," & DBSet(Mens, "T") & ",0)"
        
        conn.Execute SQL
    End If
    
    'Comprobamos la cuenta contable socio
    If CodmactaSoc <> "" Then
        If Not ComprobarCtaContaV(CodmactaSoc) Then
            Mens = "Cta.Cble Socio " & Trim(CodmactaSoc) & " no existe"
            SQL = "insert into tmpinformes (codusu, fecha1, importe1, nombre2, importe3, " & _
                  "importe4, importe5, nombre1, importe2) values (" & _
                  vSesion.Codigo & "," & DBSet(Fecha, "F") & _
                  "," & DBSet(codsoc, "N") & "," & _
                  DBSet(numfactu, "N") & "," & _
                  DBSet(c_BaseImpo, "N") & "," & _
                  DBSet(c_CuotaIva, "N") & "," & _
                  DBSet(c_TotalFac, "N") & "," & DBSet(Mens, "T") & ",0)"

            conn.Execute SQL
        End If
    End If

    'Comprobamos la cuenta contable venta
    If CodmactaVta <> "" Then
        If Not ComprobarCtaContaV(CodmactaVta) Then
            Mens = "Cta.Cble Venta " & Trim(CodmactaVta) & " no existe"
            SQL = "insert into tmpinformes (codusu, fecha1, importe1, nombre2, importe3, " & _
                  "importe4, importe5, nombre1, importe2) values (" & _
                  vSesion.Codigo & "," & DBSet(Fecha, "F") & _
                  "," & DBSet(codsoc, "N") & "," & _
                  DBSet(numfactu, "N") & "," & _
                  DBSet(c_BaseImpo, "N") & "," & _
                  DBSet(c_CuotaIva, "N") & "," & _
                  DBSet(c_TotalFac, "N") & "," & DBSet(Mens, "T") & ",0)"

            conn.Execute SQL
        End If
    End If
    
    
    'Comprobamos que el tipo de iva existe en la contabilidad
    CodIVA = DevuelveDesdeBDNew(cContaCVV, "tiposiva", "codigiva", "porceiva", CStr(porceiva), "N")
    If CodIVA = "" Then
        Mens = "El codigo de iva " & Trim(porceiva) & " no existe"
        SQL = "insert into tmpinformes (codusu, fecha1, importe1, nombre2, importe3, " & _
              "importe4, importe5, nombre1, importe2) values (" & _
              vSesion.Codigo & "," & DBSet(Fecha, "F") & _
              "," & DBSet(codsoc, "N") & "," & _
              DBSet(numfactu, "T") & "," & _
              DBSet(c_BaseImpo, "N") & "," & _
              DBSet(c_CuotaIva, "N") & "," & _
              DBSet(c_TotalFac, "N") & "," & DBSet(Mens, "T") & ",0)"

        conn.Execute SQL
    End If
    
    ' Comprobamos que la base + iva dan el total factura
    If (c_BaseImpo + c_CuotaIva) <> c_TotalFac Then
        Mens = "Base más Iva distinto de Total"
        SQL = "insert into tmpinformes (codusu, fecha1, importe1, nombre2, importe3, " & _
              "importe4, importe5, nombre1, importe2) values (" & _
              vSesion.Codigo & "," & DBSet(Fecha, "F") & _
              "," & DBSet(codsoc, "N") & "," & _
              DBSet(numfactu, "N") & "," & _
              DBSet(c_BaseImpo, "N") & "," & _
              DBSet(c_CuotaIva, "N") & "," & _
              DBSet(c_TotalFac, "N") & "," & DBSet(Mens, "T") & ",1)"

        conn.Execute SQL
    End If
    
    
    'Comprobamos que la factura no existe
    SQL = "select count(*) from cvfacturas where tipofactu = " & DBSet(Combo1.ListIndex, "N")
    SQL = SQL & " and letraser = " & DBSet(letraser, "T")
    SQL = SQL & " and numfactu = " & DBSet(numfactu, "N")
    SQL = SQL & " and fecfactu = " & DBSet(Fecha, "F")
    
    If TotalRegistros(SQL) > 0 Then
        Mens = "Existe la factura"
        SQL = "insert into tmpinformes (codusu, fecha1, importe1, nombre2, importe3, " & _
              "importe4, importe5, nombre1, importe2) values (" & _
              vSesion.Codigo & "," & DBSet(Fecha, "F") & _
              "," & DBSet(codsoc, "N") & "," & _
              DBSet(numfactu, "N") & "," & _
              DBSet(c_BaseImpo, "N") & "," & _
              DBSet(c_CuotaIva, "N") & "," & _
              DBSet(c_TotalFac, "N") & "," & DBSet(Mens, "T") & ",0)"
        
        conn.Execute SQL
    End If
    
    
eComprobarRegistroVarias:
    If Err.Number <> 0 Then
        ComprobarRegistroVarias = False
    End If
End Function
            
            
Private Function ComprobarRegistroCompras(cad As String) As Boolean
Dim SQL As String

Dim c_Base1 As Currency
Dim c_Base2 As Currency
Dim c_Base3 As Currency
Dim c_Impiv1 As Currency
Dim c_Impiv2 As Currency
Dim c_Impiv3 As Currency
Dim c_TotalFac As Currency

Dim Mens As String

Dim Fecha As String
Dim codsoc As String
Dim numfactu As String
Dim BaseImpo As String
Dim CuotaIva As String
Dim TotalFac As String
Dim Digito As String
Dim CodmactaSoc As String
Dim CodmactaVta As String
Dim longitud As Integer
Dim vCodSoc As String
Dim porceiva As Currency
Dim letraser As String

Dim Nifsocio As String
Dim CodIva1 As String
Dim CodIva2 As String
Dim CodIva3 As String

Dim Base1 As String
Dim Base2 As String
Dim Base3 As String
Dim Porc1 As String
Dim Porc2 As String
Dim Porc3 As String
Dim Impiv1 As String
Dim Impiv2 As String
Dim Impiv3 As String
Dim TipForpa As String
Dim CodForpa As Long

    On Error GoTo eComprobarRegistroCompras

    ComprobarRegistroCompras = True

    numfactu = Trim(Mid(cad, 19, 10))
    
    Fecha = Mid(cad, 11, 8)
    Fecha = Mid(Fecha, 1, 2) & "/" & Mid(Fecha, 3, 2) & "/" & Mid(Fecha, 6, 4)
    
    codsoc = Mid(cad, 37, 13)
    CodmactaSoc = Mid(cad, 97, 10)
    Nifsocio = Mid(cad, 109, 9)
    
    CodmactaVta = DevuelveDesdeBDNew(cContaCV, "cuentas", "webdatos", "codmacta", Trim(CodmactaSoc), "T")
    If CodmactaVta = "" Then CodmactaVta = vParamAplic.CtaVentaCV ' codmacta generica de parametros
    
    
    TipForpa = Mid(cad, 29, 3)
    If TipForpa = "BCO" Then
        CodForpa = vParamAplic.CodforpaBanCV
    Else
        CodForpa = vParamAplic.CodforpaConCV
    End If
    
    Base1 = Mid(cad, 124, 20)
    Base2 = Mid(cad, 144, 20)
    Base3 = Mid(cad, 164, 20)
    Porc1 = Mid(cad, 184, 3)
    Porc2 = Mid(cad, 189, 3)
    Porc3 = Mid(cad, 194, 3)
    Impiv1 = Mid(cad, 214, 20)
    Impiv2 = Mid(cad, 234, 20)
    Impiv3 = Mid(cad, 254, 20)
    TotalFac = Mid(cad, 334, 20)
    
    
    c_Base1 = CCur(Mid(Base1, 1, 14) & "," & Mid(Base1, 15, 6))
    c_Base2 = CCur(Mid(Base2, 1, 14) & "," & Mid(Base2, 15, 6))
    c_Base3 = CCur(Mid(Base3, 1, 14) & "," & Mid(Base3, 15, 6))
    c_Impiv1 = CCur(Mid(Impiv1, 1, 14) & "," & Mid(Impiv1, 15, 6))
    c_Impiv2 = CCur(Mid(Impiv2, 1, 14) & "," & Mid(Impiv2, 15, 6))
    c_Impiv3 = CCur(Mid(Impiv3, 1, 14) & "," & Mid(Impiv3, 15, 6))
    c_TotalFac = CCur(Mid(TotalFac, 1, 14) & "," & Mid(TotalFac, 15, 6))
    
    
    'Comprobamos numero de factura
    If numfactu = "" Then
        Mens = "Factura sin número"
        SQL = "insert into tmpinformes (codusu, fecha1, importe1, nombre2, importe3, " & _
              "importe4, importe5, nombre1, importe2) values (" & _
              vSesion.Codigo & "," & DBSet(Fecha, "F") & _
              "," & DBSet(codsoc, "N") & "," & _
              DBSet(numfactu, "T") & "," & _
              DBSet(c_Base1 + c_Base2 + c_Base3, "N") & "," & _
              DBSet(c_Impiv1 + c_Impiv2 + c_Impiv3, "N") & "," & _
              DBSet(c_TotalFac, "N") & "," & DBSet(Mens, "T") & ",1)"
        
        conn.Execute SQL
    End If
    
    'Comprobamos fechas
    If Not EsFechaOK(Fecha) Then
        Mens = "Fecha incorrecta"
        SQL = "insert into tmpinformes (codusu, fecha1, importe1, nombre2, importe3, " & _
              "importe4, importe5, nombre1, importe2) values (" & _
              vSesion.Codigo & "," & DBSet(Fecha, "F") & _
              "," & DBSet(codsoc, "N") & "," & _
              DBSet(numfactu, "T") & "," & _
              DBSet(c_Base1 + c_Base2 + c_Base3, "N") & "," & _
              DBSet(c_Impiv1 + c_Impiv2 + c_Impiv3, "N") & "," & _
              DBSet(c_TotalFac, "N") & "," & DBSet(Mens, "T") & ",0)"
        
        conn.Execute SQL
    End If
    
    'Comprobamos la cuenta contable socio
    If CodmactaSoc <> "" Then
        If Not ComprobarCtaConta(CodmactaSoc) Then
            Mens = "Cta.Cble Socio " & Trim(CodmactaSoc) & " no existe"
            SQL = "insert into tmpinformes (codusu, fecha1, importe1, nombre2, importe3, " & _
                  "importe4, importe5, nombre1, importe2) values (" & _
                  vSesion.Codigo & "," & DBSet(Fecha, "F") & _
                  "," & DBSet(codsoc, "N") & "," & _
                  DBSet(numfactu, "T") & "," & _
                  DBSet(c_Base1 + c_Base2 + c_Base3, "N") & "," & _
                  DBSet(c_Impiv1 + c_Impiv2 + c_Impiv3, "N") & "," & _
                  DBSet(c_TotalFac, "N") & "," & DBSet(Mens, "T") & ",0)"

            conn.Execute SQL
        End If
    End If

    'Comprobamos la cuenta contable venta
    If CodmactaVta <> "" Then
        If Not ComprobarCtaConta(CodmactaVta) Then
            Mens = "Cta.Cble Venta " & Trim(CodmactaVta) & " no existe"
            SQL = "insert into tmpinformes (codusu, fecha1, importe1, nombre2, importe3, " & _
                  "importe4, importe5, nombre1, importe2) values (" & _
                  vSesion.Codigo & "," & DBSet(Fecha, "F") & _
                  "," & DBSet(codsoc, "N") & "," & _
                  DBSet(numfactu, "T") & "," & _
                  DBSet(c_Base1 + c_Base2 + c_Base3, "N") & "," & _
                  DBSet(c_Impiv1 + c_Impiv2 + c_Impiv3, "N") & "," & _
                  DBSet(c_TotalFac, "N") & "," & DBSet(Mens, "T") & ",0)"

            conn.Execute SQL
        End If
    End If
    
    ' Comprobamos que la base + iva dan el total factura
    If (c_Base1 + c_Base2 + c_Base3 + c_Impiv1 + c_Impiv2 + c_Impiv3) <> c_TotalFac Then
        Mens = "Base más Iva distinto de Total"
        SQL = "insert into tmpinformes (codusu, fecha1, importe1, nombre2, importe3, " & _
              "importe4, importe5, nombre1, importe2) values (" & _
              vSesion.Codigo & "," & DBSet(Fecha, "F") & _
              "," & DBSet(codsoc, "N") & "," & _
              DBSet(numfactu, "T") & "," & _
              DBSet(c_Base1 + c_Base2 + c_Base3, "N") & "," & _
              DBSet(c_Impiv1 + c_Impiv2 + c_Impiv3, "N") & "," & _
              DBSet(c_TotalFac, "N") & "," & DBSet(Mens, "T") & ",1)"

        conn.Execute SQL
    End If
    
    
    Porc1 = Mid(cad, 184, 3)
    Porc2 = Mid(cad, 189, 3)
    Porc3 = Mid(cad, 194, 3)
    CodIva1 = ""
    CodIva2 = ""
    CodIva3 = ""
    'Comprobamos que los tipos de iva existen
    If c_Base1 <> 0 Then
        CodIva1 = DevuelveDesdeBDNew(cContaCV, "tiposiva", "codigiva", "porceiva", CStr(Porc1), "N")
        If CodIva1 = "" Then
            Mens = "El codigo de iva " & Trim(Porc1) & " no existe"
            SQL = "insert into tmpinformes (codusu, fecha1, importe1, nombre2, importe3, " & _
                  "importe4, importe5, nombre1, importe2) values (" & _
                  vSesion.Codigo & "," & DBSet(Fecha, "F") & _
                  "," & DBSet(codsoc, "N") & "," & _
                  DBSet(numfactu, "T") & "," & _
                  DBSet(c_Base1 + c_Base2 + c_Base3, "N") & "," & _
                  DBSet(c_Impiv1 + c_Impiv2 + c_Impiv3, "N") & "," & _
                  DBSet(c_TotalFac, "N") & "," & DBSet(Mens, "T") & ",0)"

            conn.Execute SQL
        End If
    End If
    If c_Base2 <> 0 Then
        CodIva2 = DevuelveDesdeBDNew(cContaCV, "tiposiva", "codigiva", "porceiva", CStr(Porc2), "N")
        If CodIva2 = "" Then
            Mens = "El codigo de iva " & Trim(Porc2) & " no existe"
            SQL = "insert into tmpinformes (codusu, fecha1, importe1, nombre2, importe3, " & _
                  "importe4, importe5, nombre1, importe2) values (" & _
                  vSesion.Codigo & "," & DBSet(Fecha, "F") & _
                  "," & DBSet(codsoc, "N") & "," & _
                  DBSet(numfactu, "T") & "," & _
                  DBSet(c_Base1 + c_Base2 + c_Base3, "N") & "," & _
                  DBSet(c_Impiv1 + c_Impiv2 + c_Impiv3, "N") & "," & _
                  DBSet(c_TotalFac, "N") & "," & DBSet(Mens, "T") & ",0)"

            conn.Execute SQL
        End If
    End If
    If c_Base3 <> 0 Then
        CodIva3 = DevuelveDesdeBDNew(cContaCV, "tiposiva", "codigiva", "porceiva", CStr(Porc3), "N")
        If CodIva3 = "" Then
            Mens = "El codigo de iva " & Trim(Porc3) & " no existe"
            SQL = "insert into tmpinformes (codusu, fecha1, importe1, nombre2, importe3, " & _
                  "importe4, importe5, nombre1, importe2) values (" & _
                  vSesion.Codigo & "," & DBSet(Fecha, "F") & _
                  "," & DBSet(codsoc, "N") & "," & _
                  DBSet(numfactu, "T") & "," & _
                  DBSet(c_Base1 + c_Base2 + c_Base3, "N") & "," & _
                  DBSet(c_Impiv1 + c_Impiv2 + c_Impiv3, "N") & "," & _
                  DBSet(c_TotalFac, "N") & "," & DBSet(Mens, "T") & ",0)"

            conn.Execute SQL
        End If
    End If

    
    'Comprobamos que la factura no existe
    SQL = "select count(*) from cvfacturas where tipofactu = " & DBSet(Combo1.ListIndex, "N")
    SQL = SQL & " and letraser = 'C' "
    SQL = SQL & " and numfactu = " & DBSet(numfactu, "T")
    SQL = SQL & " and fecfactu = " & DBSet(Fecha, "F")
    
    If TotalRegistros(SQL) > 0 Then
        Mens = "Existe la factura"
        SQL = "insert into tmpinformes (codusu, fecha1, importe1, nombre2, importe3, " & _
              "importe4, importe5, nombre1, importe2) values (" & _
              vSesion.Codigo & "," & DBSet(Fecha, "F") & _
              "," & DBSet(codsoc, "N") & "," & _
              DBSet(numfactu, "T") & "," & _
              DBSet(c_Base1 + c_Base2 + c_Base3, "N") & "," & _
              DBSet(c_Impiv1 + c_Impiv2 + c_Impiv3, "N") & "," & _
              DBSet(c_TotalFac, "N") & "," & DBSet(Mens, "T") & ",0)"
        
        conn.Execute SQL
    End If
    
    
eComprobarRegistroCompras:
    If Err.Number <> 0 Then
        ComprobarRegistroCompras = False
    End If
End Function
            
            
Private Function ComprobarRegistroVentas(cad As String) As Boolean
Dim SQL As String

Dim c_Base1 As Currency
Dim c_Base2 As Currency
Dim c_Base3 As Currency
Dim c_Impiv1 As Currency
Dim c_Impiv2 As Currency
Dim c_Impiv3 As Currency
Dim c_TotalFac As Currency

Dim Mens As String

Dim Fecha As String
Dim codsoc As String
Dim numfactu As String
Dim BaseImpo As String
Dim CuotaIva As String
Dim TotalFac As String
Dim Digito As String
Dim CodmactaSoc As String
Dim CodmactaVta As String
Dim longitud As Integer
Dim vCodSoc As String
Dim porceiva As Currency
Dim letraser As String

Dim Nifsocio As String
Dim CodIva1 As String
Dim CodIva2 As String
Dim CodIva3 As String

Dim Base1 As String
Dim Base2 As String
Dim Base3 As String
Dim Porc1 As String
Dim Porc2 As String
Dim Porc3 As String
Dim Impiv1 As String
Dim Impiv2 As String
Dim Impiv3 As String
Dim TipForpa As String
Dim CodForpa As Long
Dim EsTicket As Boolean

Dim sql2 As String
Dim Rs2 As ADODB.Recordset

    On Error GoTo eComprobarRegistroVentas

    ComprobarRegistroVentas = True

    EsTicket = (UCase(Mid(cad, 19, 1)) = "T")

    If Not EsTicket Then
        numfactu = Mid(cad, 19, 10)
'        codsoc = Mid(Cad, 37, 13)
    '    CodmactaSoc = Mid(Cad, 97, 10)
        codsoc = 0
        Nifsocio = Mid(cad, 109, 9)
        
        sql2 = "select codmacta from cuentas where nifdatos = " & DBSet(Nifsocio, "T") & " and mid(codmacta,1," & vEmpresaCV.DigitosNivelAnterior - 1 & ") = " & DBSet(vParamAplic.RaizCtaSocCV, "T")
        Set Rs2 = New ADODB.Recordset
        Rs2.Open sql2, ConnContaCV, adOpenForwardOnly, adLockReadOnly, adCmdText
        
        CodmactaSoc = ""
        If Not Rs2.EOF Then
            CodmactaSoc = DBLet(Rs2!Codmacta)
        End If
        Set Rs2 = Nothing
    
        If Trim(Nifsocio) <> vEmpresa.CifEmpresa Then
            CodmactaVta = vParamAplic.CtaVentaFacCV
            letraser = vParamAplic.LetraSerFCV
        Else
            CodmactaVta = vParamAplic.CtaVentaFacInCV
            letraser = vParamAplic.LetraSerFinCV
        End If
    End If
    
    
    Fecha = Mid(cad, 11, 8)
    Fecha = Mid(Fecha, 1, 2) & "/" & Mid(Fecha, 3, 2) & "/" & Mid(Fecha, 6, 4)
    
    TipForpa = Mid(cad, 29, 3)
    If TipForpa = "BCO" Then
        CodForpa = vParamAplic.CodforpaBanCV
    Else
        CodForpa = vParamAplic.CodforpaConCV
    End If
    
    Base1 = Mid(cad, 124, 20)
    Base2 = Mid(cad, 144, 20)
    Base3 = Mid(cad, 164, 20)
    Porc1 = Mid(cad, 184, 3)
    Porc2 = Mid(cad, 189, 3)
    Porc3 = Mid(cad, 194, 3)
    Impiv1 = Mid(cad, 214, 20)
    Impiv2 = Mid(cad, 234, 20)
    Impiv3 = Mid(cad, 254, 20)
    TotalFac = Mid(cad, 334, 20)
    
    
    c_Base1 = CCur(Mid(Base1, 1, 14) & "," & Mid(Base1, 15, 6))
    c_Base2 = CCur(Mid(Base2, 1, 14) & "," & Mid(Base2, 15, 6))
    c_Base3 = CCur(Mid(Base3, 1, 14) & "," & Mid(Base3, 15, 6))
    c_Impiv1 = CCur(Mid(Impiv1, 1, 14) & "," & Mid(Impiv1, 15, 6))
    c_Impiv2 = CCur(Mid(Impiv2, 1, 14) & "," & Mid(Impiv2, 15, 6))
    c_Impiv3 = CCur(Mid(Impiv3, 1, 14) & "," & Mid(Impiv3, 15, 6))
    c_TotalFac = CCur(Mid(TotalFac, 1, 14) & "," & Mid(TotalFac, 15, 6))
    
    ' si es una factura interna
    If Trim(Nifsocio) = vEmpresa.CifEmpresa And c_Base1 = 0 Then
        c_Base1 = c_TotalFac
        Porc1 = 0
        c_Impiv1 = 0
    End If
    
    'Comprobamos fechas
    If Not EsFechaOK(Fecha) Then
        Mens = "Fecha incorrecta"
        SQL = "insert into tmpinformes (codusu, fecha1, importe1, nombre2, importe3, " & _
              "importe4, importe5, nombre1, importe2) values (" & _
              vSesion.Codigo & "," & DBSet(Fecha, "F") & _
              "," & DBSet(codsoc, "N") & "," & _
              DBSet(numfactu, "T") & "," & _
              DBSet(c_Base1 + c_Base2 + c_Base3, "N") & "," & _
              DBSet(c_Impiv1 + c_Impiv2 + c_Impiv3, "N") & "," & _
              DBSet(c_TotalFac, "N") & "," & DBSet(Mens, "T") & ",0)"
        
        conn.Execute SQL
    End If
    
    If Not EsTicket Then
        'Comprobamos numero de factura
        If numfactu = "" Or Not IsNumeric(Trim(Right(numfactu, 7))) Then
            Mens = "Factura sin número o incorrecto"
            SQL = "insert into tmpinformes (codusu, fecha1, importe1, nombre2, importe3, " & _
                  "importe4, importe5, nombre1, importe2) values (" & _
                  vSesion.Codigo & "," & DBSet(Fecha, "F") & _
                  "," & DBSet(codsoc, "N") & "," & _
                  DBSet(numfactu, "T") & "," & _
                  DBSet(c_Base1 + c_Base2 + c_Base3, "N") & "," & _
                  DBSet(c_Impiv1 + c_Impiv2 + c_Impiv3, "N") & "," & _
                  DBSet(c_TotalFac, "N") & "," & DBSet(Mens, "T") & ",0)"
            
            conn.Execute SQL
        End If
    
    
        'Comprobamos la cuenta contable socio
        If CodmactaSoc <> "" Then
            If Not ComprobarCtaConta(CodmactaSoc) Then
                Mens = "Cta.Cble Socio " & Trim(CodmactaSoc) & " no existe"
                SQL = "insert into tmpinformes (codusu, fecha1, importe1, nombre2, importe3, " & _
                      "importe4, importe5, nombre1, importe2) values (" & _
                      vSesion.Codigo & "," & DBSet(Fecha, "F") & _
                      "," & DBSet(codsoc, "N") & "," & _
                      DBSet(numfactu, "T") & "," & _
                      DBSet(c_Base1 + c_Base2 + c_Base3, "N") & "," & _
                      DBSet(c_Impiv1 + c_Impiv2 + c_Impiv3, "N") & "," & _
                      DBSet(c_TotalFac, "N") & "," & DBSet(Mens, "T") & ",0)"
    
                conn.Execute SQL
            End If
        Else
            Mens = "Cta.Cble Socio " & Trim(CodmactaSoc) & " no existe"
            SQL = "insert into tmpinformes (codusu, fecha1, importe1, nombre2, importe3, " & _
                  "importe4, importe5, nombre1, importe2) values (" & _
                  vSesion.Codigo & "," & DBSet(Fecha, "F") & _
                  "," & DBSet(codsoc, "N") & "," & _
                  DBSet(numfactu, "T") & "," & _
                  DBSet(c_Base1 + c_Base2 + c_Base3, "N") & "," & _
                  DBSet(c_Impiv1 + c_Impiv2 + c_Impiv3, "N") & "," & _
                  DBSet(c_TotalFac, "N") & "," & DBSet(Mens, "T") & ",0)"

            conn.Execute SQL
        End If
    
        'Comprobamos la cuenta contable venta
        If CodmactaVta <> "" Then
            If Not ComprobarCtaConta(CodmactaVta) Then
                Mens = "Cta.Cble Venta " & Trim(CodmactaVta) & " no existe"
                SQL = "insert into tmpinformes (codusu, fecha1, importe1, nombre2, importe3, " & _
                      "importe4, importe5, nombre1, importe2) values (" & _
                      vSesion.Codigo & "," & DBSet(Fecha, "F") & _
                      "," & DBSet(codsoc, "N") & "," & _
                      DBSet(numfactu, "T") & "," & _
                      DBSet(c_Base1 + c_Base2 + c_Base3, "N") & "," & _
                      DBSet(c_Impiv1 + c_Impiv2 + c_Impiv3, "N") & "," & _
                      DBSet(c_TotalFac, "N") & "," & DBSet(Mens, "T") & ",0)"
    
                conn.Execute SQL
            End If
        Else
            Mens = "Cta.Cble Venta " & Trim(CodmactaVta) & " no existe"
            SQL = "insert into tmpinformes (codusu, fecha1, importe1, nombre2, importe3, " & _
                  "importe4, importe5, nombre1, importe2) values (" & _
                  vSesion.Codigo & "," & DBSet(Fecha, "F") & _
                  "," & DBSet(codsoc, "N") & "," & _
                  DBSet(numfactu, "T") & "," & _
                  DBSet(c_Base1 + c_Base2 + c_Base3, "N") & "," & _
                  DBSet(c_Impiv1 + c_Impiv2 + c_Impiv3, "N") & "," & _
                  DBSet(c_TotalFac, "N") & "," & DBSet(Mens, "T") & ",0)"

            conn.Execute SQL
        End If
    
    End If
    
    
    
    ' Comprobamos que la base + iva dan el total factura
    If (c_Base1 + c_Base2 + c_Base3 + c_Impiv1 + c_Impiv2 + c_Impiv3) <> c_TotalFac Then
        Mens = "Base más Iva distinto de Total"
        SQL = "insert into tmpinformes (codusu, fecha1, importe1, nombre2, importe3, " & _
              "importe4, importe5, nombre1, importe2) values (" & _
              vSesion.Codigo & "," & DBSet(Fecha, "F") & _
              "," & DBSet(codsoc, "N") & "," & _
              DBSet(numfactu, "T") & "," & _
              DBSet(c_Base1 + c_Base2 + c_Base3, "N") & "," & _
              DBSet(c_Impiv1 + c_Impiv2 + c_Impiv3, "N") & "," & _
              DBSet(c_TotalFac, "N") & "," & DBSet(Mens, "T") & ",1)"

        conn.Execute SQL
    End If
    
    
    CodIva1 = ""
    CodIva2 = ""
    CodIva3 = ""
    'Comprobamos que los tipos de iva existen
    If c_Base1 <> 0 Then
        CodIva1 = DevuelveDesdeBDNew(cContaCV, "tiposiva", "codigiva", "porceiva", CStr(Porc1), "N")
        If CodIva1 = "" Then
            Mens = "El codigo de iva " & Trim(Porc1) & " no existe"
            SQL = "insert into tmpinformes (codusu, fecha1, importe1, nombre2, importe3, " & _
                  "importe4, importe5, nombre1, importe2) values (" & _
                  vSesion.Codigo & "," & DBSet(Fecha, "F") & _
                  "," & DBSet(codsoc, "N") & "," & _
                  DBSet(numfactu, "T") & "," & _
                  DBSet(c_Base1 + c_Base2 + c_Base3, "N") & "," & _
                  DBSet(c_Impiv1 + c_Impiv2 + c_Impiv3, "N") & "," & _
                  DBSet(c_TotalFac, "N") & "," & DBSet(Mens, "T") & ",0)"

            conn.Execute SQL
        End If
    End If
    If c_Base2 <> 0 Then
        CodIva2 = DevuelveDesdeBDNew(cContaCV, "tiposiva", "codigiva", "porceiva", CStr(Porc2), "N")
        If CodIva2 = "" Then
            Mens = "El codigo de iva " & Trim(Porc2) & " no existe"
            SQL = "insert into tmpinformes (codusu, fecha1, importe1, nombre2, importe3, " & _
                  "importe4, importe5, nombre1, importe2) values (" & _
                  vSesion.Codigo & "," & DBSet(Fecha, "F") & _
                  "," & DBSet(codsoc, "N") & "," & _
                  DBSet(numfactu, "T") & "," & _
                  DBSet(c_Base1 + c_Base2 + c_Base3, "N") & "," & _
                  DBSet(c_Impiv1 + c_Impiv2 + c_Impiv3, "N") & "," & _
                  DBSet(c_TotalFac, "N") & "," & DBSet(Mens, "T") & ",0)"

            conn.Execute SQL
        End If
    End If
    If c_Base3 <> 0 Then
        CodIva3 = DevuelveDesdeBDNew(cContaCV, "tiposiva", "codigiva", "porceiva", CStr(Porc3), "N")
        If CodIva3 = "" Then
            Mens = "El codigo de iva " & Trim(Porc3) & " no existe"
            SQL = "insert into tmpinformes (codusu, fecha1, importe1, nombre2, importe3, " & _
                  "importe4, importe5, nombre1, importe2) values (" & _
                  vSesion.Codigo & "," & DBSet(Fecha, "F") & _
                  "," & DBSet(codsoc, "N") & "," & _
                  DBSet(numfactu, "T") & "," & _
                  DBSet(c_Base1 + c_Base2 + c_Base3, "N") & "," & _
                  DBSet(c_Impiv1 + c_Impiv2 + c_Impiv3, "N") & "," & _
                  DBSet(c_TotalFac, "N") & "," & DBSet(Mens, "T") & ",0)"

            conn.Execute SQL
        End If
    End If

    If Not EsTicket Then
        'Comprobamos que la factura no existe
        SQL = "select count(*) from cvfacturas where tipofactu = " & DBSet(Combo1.ListIndex, "N")
        SQL = SQL & " and letraser = " & DBSet(letraser, "T")
        SQL = SQL & " and numfactu = " & DBSet(numfactu, "T")
        SQL = SQL & " and fecfactu = " & DBSet(Fecha, "F")
        
        If TotalRegistros(SQL) > 0 Then
            Mens = "Existe la factura"
            SQL = "insert into tmpinformes (codusu, fecha1, importe1, nombre2, importe3, " & _
                  "importe4, importe5, nombre1, importe2) values (" & _
                  vSesion.Codigo & "," & DBSet(Fecha, "F") & _
                  "," & DBSet(codsoc, "N") & "," & _
                  DBSet(numfactu, "T") & "," & _
                  DBSet(c_Base1 + c_Base2 + c_Base3, "N") & "," & _
                  DBSet(c_Impiv1 + c_Impiv2 + c_Impiv3, "N") & "," & _
                  DBSet(c_TotalFac, "N") & "," & DBSet(Mens, "T") & ",0)"
            
            conn.Execute SQL
        End If
    
    Else
        ' insertamos en la tabla temporal de tickets
        SQL = "insert into tmptickets (codusu,fecfactu,baseimpo,porciva,codiva,cuotaiva,baseimpo2,porciva2,codiva2,cuotaiva2,"
        SQL = SQL & "baseimpo3,porciva3,codiva3,cuotaiva3,totalfac,codforpa) values (" & vSesion.Codigo & "," & DBSet(Fecha, "F") & ","
        SQL = SQL & DBSet(c_Base1, "N") & "," & DBSet(Porc1, "N") & "," & DBSet(CodIva1, "N") & "," & DBSet(c_Impiv1, "N") & ","
        SQL = SQL & DBSet(c_Base2, "N") & "," & DBSet(Porc2, "N") & "," & DBSet(CodIva2, "N") & "," & DBSet(c_Impiv2, "N") & ","
        SQL = SQL & DBSet(c_Base3, "N") & "," & DBSet(Porc3, "N") & "," & DBSet(CodIva3, "N") & "," & DBSet(c_Impiv3, "N") & ","
        SQL = SQL & DBSet(c_TotalFac, "N") & "," & DBSet(CodForpa, "N") & ")"
        
        conn.Execute SQL
    
    End If
eComprobarRegistroVentas:
    If Err.Number <> 0 Then
        ComprobarRegistroVentas = False
    End If
End Function
            
            
            
            
            
            
            
Private Function ComprobarRegistro(cad As String) As Boolean
Dim SQL As String

Dim c_BaseImpo As Currency
Dim c_CuotaIva As Currency
Dim c_TotalFac As Currency

Dim Mens As String

Dim Fecha As String
Dim codsoc As String
Dim numfactu As String
Dim BaseImpo As String
Dim CuotaIva As String
Dim TotalFac As String
Dim Digito As String
Dim Codmacta As String
Dim longitud As Integer
Dim vCodSoc As String

    On Error GoTo eComprobarRegistro

'    ComprobarRegistro = True
'
'    fecha = RecuperaValor(Cad, 1)
'    codsoc = RecuperaValor(Cad, 4)
'    vCodSoc = ""
'    If EsEntero(codsoc) Then
''[Monica]28/10/2010: ya no se resta 3000
''12/11/2010: de momento sí
'        If CLng(codsoc) > 3000 Then
'            codsoc = CStr(CLng(codsoc) - 3000)
''        Else
''            vCodSoc = "0"
'        End If
'        vCodSoc = codsoc
'    Else
'        vCodSoc = "0"
'    End If
'
'    numfactu = RecuperaValor(Cad, 2)
'    numfactu = Replace(numfactu, "-", "|") & "|"
'    Digito = RecuperaValor(numfactu, 1)
'    numfactu = RecuperaValor(numfactu, 4)
''    NumFactu = Format((CInt(Digito) * 1000000) + CLng(NumFactu), "0000000")
'
'    longitud = vEmpresaTel.DigitosUltimoNivel - vEmpresaTel.DigitosNivelAnterior
'    Codmacta = vParamAplic.RaizCtaSocTel & Right("0000000000" & codsoc, longitud)
''    Codmacta = Trim(vParamAplic.RaizCtaSocTel & Format(CodSoc, "00000"))
'
'    BaseImpo = RecuperaValor(Cad, 6)
'    CuotaIva = RecuperaValor(Cad, 7)
'    TotalFac = RecuperaValor(Cad, 8)
'
'    c_BaseImpo = CCur(TransformaPuntosComas(BaseImpo))
'    c_CuotaIva = CCur(TransformaPuntosComas(CuotaIva))
'    c_TotalFac = CCur(TransformaPuntosComas(TotalFac))
'
'    'Codigo de socio incorrecto
'    If vCodSoc = "0" Then
'        Mens = "Socio incorrecto"
'        sql = "insert into tmpinformes (codusu, fecha1, importe1, importe2, importe3, " & _
'              "importe4, importe5, nombre1) values (" & _
'              vSesion.Codigo & "," & DBSet(fecha, "F") & _
'              "," & DBSet(codsoc, "N") & "," & _
'              DBSet(numfactu, "N") & "," & _
'              DBSet(c_BaseImpo, "N") & "," & _
'              DBSet(c_CuotaIva, "N") & "," & _
'              DBSet(c_TotalFac, "N") & "," & DBSet(Mens, "T") & ")"
'
'        conn.Execute sql
'    End If
'
'    'Comprobamos fechas
'    If Not EsFechaOK(fecha) Then
'        Mens = "Fecha incorrecta"
'        sql = "insert into tmpinformes (codusu, fecha1, importe1, importe2, importe3, " & _
'              "importe4, importe5, nombre1) values (" & _
'              vSesion.Codigo & "," & DBSet(fecha, "F") & _
'              "," & DBSet(codsoc, "N") & "," & _
'              DBSet(numfactu, "N") & "," & _
'              DBSet(c_BaseImpo, "N") & "," & _
'              DBSet(c_CuotaIva, "N") & "," & _
'              DBSet(c_TotalFac, "N") & "," & DBSet(Mens, "T") & ")"
'
'        conn.Execute sql
'    End If
'
'    'Comprobamos la cuenta contable
'    If Codmacta <> "" Then
'        If Not ComprobarCtaConta(Codmacta) Then
'            Mens = "Cta.Contable " & Trim(Codmacta) & " no existe"
'            sql = "insert into tmpinformes (codusu, fecha1, importe1, importe2, importe3, " & _
'                  "importe4, importe5, nombre1) values (" & _
'                  vSesion.Codigo & "," & DBSet(fecha, "F") & _
'                  "," & DBSet(codsoc, "N") & "," & _
'                  DBSet(numfactu, "N") & "," & _
'                  DBSet(c_BaseImpo, "N") & "," & _
'                  DBSet(c_CuotaIva, "N") & "," & _
'                  DBSet(c_TotalFac, "N") & "," & DBSet(Mens, "T") & ")"
'
'            conn.Execute sql
'        End If
'    End If
'
'    ' Comprobamos que la base + iva dan el total factura
'    If (c_BaseImpo + c_CuotaIva) <> c_TotalFac Then
'        Mens = "Base más Iva distinto de Total"
'        sql = "insert into tmpinformes (codusu, fecha1, importe1, importe2, importe3, " & _
'              "importe4, importe5, nombre1) values (" & _
'              vSesion.Codigo & "," & DBSet(fecha, "F") & _
'              "," & DBSet(codsoc, "N") & "," & _
'              DBSet(numfactu, "N") & "," & _
'              DBSet(c_BaseImpo, "N") & "," & _
'              DBSet(c_CuotaIva, "N") & "," & _
'              DBSet(c_TotalFac, "N") & "," & DBSet(Mens, "T") & ")"
'
'        conn.Execute sql
'    End If
'
'
'    'Comprobamos que la factura no existe
'    sql = "select count(*) from telmovil where numserie = " & DBSet(Text1(17).Text, "T")
'    sql = sql & " and numfactu = " & DBSet(numfactu, "N")
'    sql = sql & " and fecfactu = " & DBSet(fecha, "F")
'
'    If TotalRegistros(sql) > 0 Then
'        Mens = "Existe la factura"
'        sql = "insert into tmpinformes (codusu, fecha1, importe1, importe2, importe3, " & _
'              "importe4, importe5, nombre1) values (" & _
'              vSesion.Codigo & "," & DBSet(fecha, "F") & _
'              "," & DBSet(codsoc, "N") & "," & _
'              DBSet(numfactu, "N") & "," & _
'              DBSet(c_BaseImpo, "N") & "," & _
'              DBSet(c_CuotaIva, "N") & "," & _
'              DBSet(c_TotalFac, "N") & "," & DBSet(Mens, "T") & ")"
'
'        conn.Execute sql
'    End If
'
    
    
eComprobarRegistro:
    If Err.Number <> 0 Then
        ComprobarRegistro = False
    End If
End Function
            
Private Function InsertarLineaVarias(cad As String) As Boolean
Dim c_BaseImpo As Currency
Dim c_CuotaIva As Currency
Dim c_TotalFac As Currency

Dim Mens As String

Dim Fecha As String
Dim codsoc As String
Dim numfactu As String
Dim BaseImpo As String
Dim CuotaIva As String
Dim TotalFac As String
Dim Digito As String
Dim SQL As String
Dim CodmactaSoc As String
Dim CodmactaVta As String
Dim longitud As Integer
Dim porceiva As Currency
Dim letraser As String
Dim CodIVA As Integer
Dim TipForpa As String
Dim CodForpa As Long
Dim Erronea As Byte

    On Error GoTo eInsertarLinea

    InsertarLineaVarias = True

    numfactu = RecuperaValor(cad, 1)
    
    If numfactu = "CodiFactura" Then Exit Function
    
    
    Fecha = RecuperaValor(cad, 2)
    codsoc = RecuperaValor(cad, 3)
    porceiva = Mid(RecuperaValor(cad, 5), 1, Len(RecuperaValor(cad, 5)) - 1)
    CuotaIva = RecuperaValor(cad, 6)
    BaseImpo = RecuperaValor(cad, 7)
    TotalFac = RecuperaValor(cad, 8)
    
    
    CodmactaSoc = RecuperaValor(cad, 9)
    CodmactaVta = RecuperaValor(cad, 10)
    letraser = Trim(txtCodigo(2).Text) 'vParamAplic.LetraSerCVV 'Mid(RecuperaValor(cad, 11), 1, 1)
    
    TipForpa = RecuperaValor(cad, 12)
    If TipForpa = "Remesa" Then
        CodForpa = vParamAplic.CodforpaBanCVV
    Else
        CodForpa = vParamAplic.CodforpaConCVV
    End If
    
    
    
    c_BaseImpo = CCur(TransformaPuntosComas(ImporteSinFormato(BaseImpo)))
    c_CuotaIva = CCur(TransformaPuntosComas(ImporteSinFormato(CuotaIva)))
    c_TotalFac = CCur(TransformaPuntosComas(ImporteSinFormato(TotalFac)))
    
    Erronea = 0
    If c_BaseImpo + c_CuotaIva <> c_TotalFac Then Erronea = 1
    
    '[Monica]26/10/2012: tipo de iva tiene que ser 0
    CodIVA = DevuelveDesdeBDNew(cContaCVV, "tiposiva", "codigiva", "porceiva", CStr(porceiva), "N", , "tipodiva", "0", "N")
    
    ' insertamos en la tabla de facturas
    SQL = "INSERT INTO cvfacturas (tipofactu,letraser,numfactu,fecfactu,codsocio,codmactasoc,codmactavta,"
    SQL = SQL & "baseimpo,codiva,porciva,cuotaiva,totalfac,intconta,codforpa, erronea) VALUES (0," & DBSet(letraser, "T") & ","
    SQL = SQL & DBSet(numfactu, "N") & "," & DBSet(Fecha, "F") & "," & DBSet(codsoc, "N") & "," & DBSet(CodmactaSoc, "T") & ","
    SQL = SQL & DBSet(CodmactaVta, "T") & "," & DBSet(c_BaseImpo, "N") & "," & DBSet(CodIVA, "N") & ","
    SQL = SQL & DBSet(porceiva, "N") & ","
    SQL = SQL & DBSet(c_CuotaIva, "N") & "," & DBSet(c_TotalFac, "N") & ",0," & DBSet(CodForpa, "N") & "," & DBSet(Erronea, "N") & ")"
    
    conn.Execute SQL
 
eInsertarLinea:
    If Err.Number <> 0 Then
        InsertarLineaVarias = False
        MsgBox "Error en Insertar Linea Varias" & Err.Description, vbExclamation
    End If
End Function
            
            
            
Private Function InsertarLineaCompras(cad As String) As Boolean

Dim c_Base1 As Currency
Dim c_Base2 As Currency
Dim c_Base3 As Currency
Dim c_Impiv1 As Currency
Dim c_Impiv2 As Currency
Dim c_Impiv3 As Currency
Dim c_TotalFac As Currency


Dim Mens As String

Dim Fecha As String
Dim codsoc As String
Dim numfactu As String
Dim BaseImpo As String
Dim CuotaIva As String
Dim TotalFac As String
Dim Digito As String
Dim SQL As String
Dim CodmactaSoc As String
Dim CodmactaVta As String
Dim longitud As Integer
Dim porceiva As Currency
Dim letraser As String
Dim CodIVA As Integer

Dim Nifsocio As String
Dim Base1 As String
Dim Base2 As String
Dim Base3 As String
Dim Porc1 As String
Dim Porc2 As String
Dim Porc3 As String
Dim Impiv1 As String
Dim Impiv2 As String
Dim Impiv3 As String
Dim CodIva1 As String
Dim CodIva2 As String
Dim CodIva3 As String

Dim TipForpa As String
Dim CodForpa As Long

Dim CtaVentas As String
Dim Erronea As Byte

    On Error GoTo eInsertarLinea

    InsertarLineaCompras = True

    numfactu = Trim(Mid(cad, 19, 10))
    If numfactu = "" Then numfactu = "SIN NRO"
    Fecha = Mid(cad, 11, 8)
    Fecha = Mid(Fecha, 1, 2) & "/" & Mid(Fecha, 3, 2) & "/" & Mid(Fecha, 6, 4)
'    FPago = Mid(Cad, 29, 3)
    codsoc = Mid(cad, 37, 13)
    CodmactaSoc = Mid(cad, 97, 10)
    Nifsocio = Mid(cad, 109, 9)
    
    CodmactaVta = DevuelveDesdeBDNew(cContaCV, "cuentas", "webdatos", "codmacta", Trim(CodmactaSoc), "T")
    If CodmactaVta = "" Then CodmactaVta = vParamAplic.CtaVentaCV ' codmacta generica de parametros
    
    TipForpa = Mid(cad, 29, 3)
    If TipForpa = "BCO" Then
        CodForpa = vParamAplic.CodforpaBanCV
    Else
        CodForpa = vParamAplic.CodforpaConCV
    End If
    
    Base1 = Mid(cad, 124, 20)
    Base2 = Mid(cad, 144, 20)
    Base3 = Mid(cad, 164, 20)
    Porc1 = Mid(cad, 184, 3)
    Porc2 = Mid(cad, 189, 3)
    Porc3 = Mid(cad, 194, 3)
    Impiv1 = Mid(cad, 214, 20)
    Impiv2 = Mid(cad, 234, 20)
    Impiv3 = Mid(cad, 254, 20)
    TotalFac = Mid(cad, 334, 20)
    
    '[Monica]26/10/2012: tipo de iva tiene que ser 0
    CodIva1 = DevuelveDesdeBDNew(cContaCV, "tiposiva", "codigiva", "porceiva", CStr(Porc1), "N", , "tipodiva", "0", "N")
    CodIva2 = DevuelveDesdeBDNew(cContaCV, "tiposiva", "codigiva", "porceiva", CStr(Porc2), "N", , "tipodiva", "0", "N")
    CodIva3 = DevuelveDesdeBDNew(cContaCV, "tiposiva", "codigiva", "porceiva", CStr(Porc3), "N", , "tipodiva", "0", "N")
    
    
    c_Base1 = CCur(Mid(Base1, 1, 14) & "," & Mid(Base1, 15, 6))
    c_Base2 = CCur(Mid(Base2, 1, 14) & "," & Mid(Base2, 15, 6))
    c_Base3 = CCur(Mid(Base3, 1, 14) & "," & Mid(Base3, 15, 6))
    c_Impiv1 = CCur(Mid(Impiv1, 1, 14) & "," & Mid(Impiv1, 15, 6))
    c_Impiv2 = CCur(Mid(Impiv2, 1, 14) & "," & Mid(Impiv2, 15, 6))
    c_Impiv3 = CCur(Mid(Impiv3, 1, 14) & "," & Mid(Impiv3, 15, 6))
    c_TotalFac = CCur(Mid(TotalFac, 1, 14) & "," & Mid(TotalFac, 15, 6))

    Erronea = 0
    If c_Base1 + c_Base2 + c_Base3 + c_Impiv1 + c_Impiv2 + c_Impiv3 <> c_TotalFac Then Erronea = 1
    
    If c_Base1 = 0 Then
        If c_Base2 = 0 Then
            c_Base1 = c_Base3
            c_Impiv1 = c_Impiv3
            Porc1 = Porc3
            CodIva1 = CodIva3
            
            c_Base3 = 0
            c_Impiv3 = 0
            Porc3 = 0
            CodIva3 = ""
        Else
            If c_Base3 = 0 Then
                c_Base1 = c_Base2
                c_Impiv1 = c_Impiv2
                Porc1 = Porc2
                CodIva1 = CodIva2
                
                c_Base2 = 0
                c_Impiv2 = 0
                Porc2 = 0
                CodIva2 = ""
            Else
                c_Base1 = c_Base3
                c_Impiv1 = c_Impiv3
                Porc1 = Porc3
                CodIva1 = CodIva3
                
                c_Base3 = 0
                c_Impiv3 = 0
                Porc3 = 0
                CodIva3 = ""
            End If
        End If
    End If
    If c_Base2 = 0 Then
        If c_Base3 = 0 Then
            ' No hacemos nada
        Else
            c_Base2 = c_Base3
            c_Impiv2 = c_Impiv3
            Porc2 = Porc3
            CodIva2 = CodIva3
            
            c_Base3 = 0
            c_Impiv3 = 0
            Porc3 = 0
            CodIva3 = ""
        End If
    End If
    
    If c_Base2 = 0 Then
         Porc2 = 0
         CodIva2 = 0
    End If
    If c_Base3 = 0 Then
         Porc3 = 0
         CodIva3 = 0
    End If
    
    
    ' insertamos en la tabla de facturas
    SQL = "INSERT INTO cvfacturas (tipofactu,letraser,numfactu,fecfactu,codsocio,nifsocio,codmactasoc,codmactavta,"
    SQL = SQL & "baseimpo,codiva,porciva,cuotaiva,baseimpo2,codiva2,porciva2,cuotaiva2,baseimpo3,codiva3,porciva3,cuotaiva3,totalfac,intconta, codforpa, erronea) VALUES (2,'C',"
    SQL = SQL & DBSet(numfactu, "T") & "," & DBSet(Fecha, "F") & "," & DBSet(codsoc, "N") & "," & DBSet(Nifsocio, "T") & "," & DBSet(CodmactaSoc, "T") & ","
    SQL = SQL & DBSet(CodmactaVta, "T") & ","
    SQL = SQL & DBSet(c_Base1, "N") & "," & DBSet(CodIva1, "N") & "," & DBSet(Porc1, "N") & "," & DBSet(c_Impiv1, "N") & ","
    SQL = SQL & DBSet(c_Base2, "N", "S") & "," & DBSet(CodIva2, "N", "S") & "," & DBSet(Porc2, "N", "S") & "," & DBSet(c_Impiv2, "N", "S") & ","
    SQL = SQL & DBSet(c_Base3, "N", "S") & "," & DBSet(CodIva3, "N", "S") & "," & DBSet(Porc3, "N", "S") & "," & DBSet(c_Impiv3, "N", "S") & ","
    SQL = SQL & DBSet(c_TotalFac, "N") & ",0," & DBSet(CodForpa, "N") & "," & DBSet(Erronea, "N") & ")"
    
    conn.Execute SQL
 
eInsertarLinea:
    If Err.Number <> 0 Then
        InsertarLineaCompras = False
        MsgBox "Error en Insertar Linea Compras" & Err.Description, vbExclamation
    End If
End Function
            
            
Private Function InsertarLineaTickets() As Boolean
Dim c_Base1 As Currency
Dim c_Base2 As Currency
Dim c_Base3 As Currency
Dim c_Impiv1 As Currency
Dim c_Impiv2 As Currency
Dim c_Impiv3 As Currency
Dim c_TotalFac As Currency


Dim Mens As String

Dim Fecha As String
Dim codsoc As String
Dim numfactu As String
Dim BaseImpo As String
Dim CuotaIva As String
Dim TotalFac As String
Dim Digito As String
Dim SQL As String
Dim sql2 As String
Dim Sql3 As String
Dim CodmactaSoc As String
Dim CodmactaVta As String
Dim longitud As Integer
Dim porceiva As Currency
Dim letraser As String
Dim CodIVA As Integer

Dim Nifsocio As String
Dim Base1 As String
Dim Base2 As String
Dim Base3 As String
Dim Porc1 As String
Dim Porc2 As String
Dim Porc3 As String
Dim Impiv1 As String
Dim Impiv2 As String
Dim Impiv3 As String
Dim CodIva1 As String
Dim CodIva2 As String
Dim CodIva3 As String

Dim TipForpa As String
Dim CodForpa As Long

Dim CtaVentas As String
Dim EsTicket As Boolean

Dim Rs As ADODB.Recordset
Dim Rs3 As ADODB.Recordset
Dim NumLinea As Long

Dim FIni As String
Dim FFin As String
Dim Existe As Boolean
Dim Erronea As Byte

Dim OK As Boolean

    On Error GoTo eInsertarLinea

    InsertarLineaTickets = False


    ConnContaCV.BeginTrans


    ' insertamos previamente los tickets
    SQL = "select fecfactu, codforpa, porciva, porciva2, porciva3, "
    SQL = SQL & " sum(baseimpo) baseimpo, sum(baseimpo2) baseimpo2, sum(baseimpo3) baseimpo3, "
    SQL = SQL & " sum(cuotaiva) cuotaiva, sum(cuotaiva2) cuotaiva2, sum(cuotaiva3) cuotaiva3, sum(totalfac) totalfac "
    SQL = SQL & " from tmptickets where codusu = " & vSesion.Codigo
    SQL = SQL & " group by 1,2,3,4,5 "
    SQL = SQL & " order by 1,2,3,4,5 "
    
    FIni = ""
    FFin = ""
'    FechasEjercicioContaCV FIni, FFin
    
    
    Set Rs = New ADODB.Recordset
    Rs.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    '[Monica]05/11/2012: cambio de fecha inicio y fin de ejercicio
    If Not Rs.EOF Then
        FIni = "01/01/" & Format(Year(Rs!fecfactu), "0000")
        FFin = "31/12/" & Format(Year(Rs!fecfactu), "0000")
    End If
    
    While Not Rs.EOF
        
        Sql3 = "select if(" & DBSet(Rs!fecfactu, "F") & " <= " & DBSet(FFin, "F") & ",contado1,contado2) numlinea from contadores where tiporegi = " & DBSet(vParamAplic.LetraSerCV, "T")
        Set Rs3 = New ADODB.Recordset
        Rs3.Open Sql3, ConnContaCV, adOpenForwardOnly, adLockReadOnly, adCmdText
        If Not Rs3.EOF Then
            NumLinea = DBLet(Rs3!NumLinea)
            
        End If
        
        Set Rs3 = Nothing
        
        '[Monica]05/04/2013: Comprobamos que el contador sea correcto
        OK = (NumLinea = DevuelveValor("select max(numfactu) from cvfacturas where tipofactu = 1 and letraser = 'TK' and fecfactu between " & DBSet(FIni, "F") & " and " & DBSet(FFin, "F")))
        If NumLinea = 0 Or Not OK Then
            MsgBox "El contador de tickets correspondiente en contabilidad no está correcto. Revise.", vbExclamation
            Exit Function
        End If
        
        
        Existe = False
        Do
            NumLinea = NumLinea + 1
            Sql3 = "select count(*) from cvfacturas where letraser = " & DBSet(vParamAplic.LetraSerCV, "T") & " and numfactu = " & DBSet(NumLinea, "N")
            Set Rs3 = New ADODB.Recordset
            Rs3.Open Sql3, conn, adOpenForwardOnly, adLockReadOnly, adCmdText
            If Not Rs3.EOF Then
                Existe = (DBLet(Rs3.Fields(0).Value, "N") >= 1)
            End If
            Set Rs3 = Nothing
        Loop Until Not Existe
        
        
        c_Base1 = DBLet(Rs!BaseImpo)
        c_Base2 = DBLet(Rs!BaseImpo2)
        c_Base3 = DBLet(Rs!BaseImpo3)
        c_Impiv1 = DBLet(Rs!CuotaIva)
        c_Impiv2 = DBLet(Rs!cuotaiva2)
        c_Impiv3 = DBLet(Rs!cuotaiva3)
        c_TotalFac = DBLet(Rs!TotalFac)
    
        Porc1 = DBLet(Rs!PorcIva)
        Porc2 = DBLet(Rs!PorcIva2)
        Porc3 = DBLet(Rs!PorcIva3)
        
       '[Monica]26/10/2012: tipo de iva tiene que ser 0
        CodIva1 = DevuelveDesdeBDNew(cContaCV, "tiposiva", "codigiva", "porceiva", CStr(Porc1), "N", "tipodiva", "0", "N")
        CodIva2 = DevuelveDesdeBDNew(cContaCV, "tiposiva", "codigiva", "porceiva", CStr(Porc2), "N", "tipodiva", "0", "N")
        CodIva3 = DevuelveDesdeBDNew(cContaCV, "tiposiva", "codigiva", "porceiva", CStr(Porc3), "N", "tipodiva", "0", "N")
        
        Erronea = 0
        If c_Base1 + c_Base2 + c_Base3 + c_Impiv1 + c_Impiv2 + c_Impiv3 <> c_TotalFac Then Erronea = 1
        
        
        If c_Base1 = 0 Then
            If c_Base2 = 0 Then
                c_Base1 = c_Base3
                c_Impiv1 = c_Impiv3
                Porc1 = Porc3
                CodIva1 = CodIva3
                
                c_Base3 = 0
                c_Impiv3 = 0
                Porc3 = 0
                CodIva3 = ""
            Else
                If c_Base3 = 0 Then
                    c_Base1 = c_Base2
                    c_Impiv1 = c_Impiv2
                    Porc1 = Porc2
                    CodIva1 = CodIva2
                    
                    c_Base2 = 0
                    c_Impiv2 = 0
                    Porc2 = 0
                    CodIva2 = ""
                Else
                    c_Base1 = c_Base3
                    c_Impiv1 = c_Impiv3
                    Porc1 = Porc3
                    CodIva1 = CodIva3
                    
                    c_Base3 = 0
                    c_Impiv3 = 0
                    Porc3 = 0
                    CodIva3 = ""
                End If
            End If
        End If
        If c_Base2 = 0 Then
            If c_Base3 = 0 Then
                ' No hacemos nada
            Else
                c_Base2 = c_Base3
                c_Impiv2 = c_Impiv3
                Porc2 = Porc3
                CodIva2 = CodIva3
                
                c_Base3 = 0
                c_Impiv3 = 0
                Porc3 = 0
                CodIva3 = ""
            End If
        End If
    
        ' insertamos en la tabla de facturas
        SQL = "INSERT INTO cvfacturas (tipofactu,letraser,numfactu,fecfactu,codsocio,nifsocio,codmactasoc,codmactavta,"
        SQL = SQL & "baseimpo,codiva,porciva,cuotaiva,baseimpo2,codiva2,porciva2,cuotaiva2,baseimpo3,codiva3,porciva3,cuotaiva3,totalfac,intconta, codforpa, erronea) VALUES (1," & DBSet(vParamAplic.LetraSerCV, "T") & ","
        SQL = SQL & DBSet(NumLinea, "T") & "," & DBSet(Rs!fecfactu, "F") & "," & DBSet(0, "N") & "," & DBSet("", "T") & "," & DBSet(vParamAplic.CtaClienTickCV, "T") & ","
        SQL = SQL & DBSet(vParamAplic.CtaVentaTickCV, "T") & ","
        SQL = SQL & DBSet(c_Base1, "N") & "," & DBSet(CodIva1, "N") & "," & DBSet(Porc1, "N") & "," & DBSet(c_Impiv1, "N") & ","
        SQL = SQL & DBSet(c_Base2, "N") & "," & DBSet(CodIva2, "N") & "," & DBSet(Porc2, "N") & "," & DBSet(c_Impiv2, "N") & ","
        SQL = SQL & DBSet(c_Base3, "N") & "," & DBSet(CodIva3, "N") & "," & DBSet(Porc3, "N") & "," & DBSet(c_Impiv3, "N") & ","
        SQL = SQL & DBSet(c_TotalFac, "N") & ",0," & DBSet(Rs!CodForpa, "N") & "," & DBSet(Erronea, "N") & ")"
        
        conn.Execute SQL
    
        If CDate(DBLet(Rs!fecfactu)) <= CDate(FFin) Then
            sql2 = "update contadores set contado1 = " & DBSet(NumLinea, "N")
            sql2 = sql2 & " where tiporegi = " & DBSet(vParamAplic.LetraSerCV, "T")
            ConnContaCV.Execute sql2
        Else
            sql2 = "update contadores set contado2 = " & DBSet(NumLinea, "N")
            sql2 = sql2 & " where tiporegi = " & DBSet(vParamAplic.LetraSerCV, "T")
            ConnContaCV.Execute sql2
        End If
        
        Rs.MoveNext
    
    Wend
    Set Rs = Nothing
'    ConnContaCV.CommitTrans
    InsertarLineaTickets = True
    
eInsertarLinea:
    If Err.Number <> 0 Then
'        ConnContaCV.RollbackTrans
        
        InsertarLineaTickets = False
        MsgBox "Error en Insertar Linea Tickets" & Err.Description, vbExclamation
    End If


End Function
            
            
            
Private Function InsertarLineaVentas(cad As String) As Boolean

Dim c_Base1 As Currency
Dim c_Base2 As Currency
Dim c_Base3 As Currency
Dim c_Impiv1 As Currency
Dim c_Impiv2 As Currency
Dim c_Impiv3 As Currency
Dim c_TotalFac As Currency


Dim Mens As String

Dim Fecha As String
Dim codsoc As String
Dim numfactu As String
Dim BaseImpo As String
Dim CuotaIva As String
Dim TotalFac As String
Dim Digito As String
Dim SQL As String
Dim sql2 As String
Dim Sql3 As String
Dim CodmactaSoc As String
Dim CodmactaVta As String
Dim longitud As Integer
Dim porceiva As Currency
Dim letraser As String
Dim CodIVA As Integer

Dim Nifsocio As String
Dim Base1 As String
Dim Base2 As String
Dim Base3 As String
Dim Porc1 As String
Dim Porc2 As String
Dim Porc3 As String
Dim Impiv1 As String
Dim Impiv2 As String
Dim Impiv3 As String
Dim CodIva1 As String
Dim CodIva2 As String
Dim CodIva3 As String

Dim TipForpa As String
Dim CodForpa As Long

Dim CtaVentas As String
Dim EsTicket As Boolean

Dim Rs As ADODB.Recordset
Dim Rs3 As ADODB.Recordset
Dim NumLinea As Long

Dim FIni As String
Dim FFin As String
Dim Existe As Boolean
Dim Erronea As Byte
Dim Rs2 As ADODB.Recordset

    On Error GoTo eInsertarLinea

    InsertarLineaVentas = True

    EsTicket = (UCase(Mid(cad, 19, 1)) = "T")

    If Not EsTicket Then
        numfactu = Mid(cad, 19, 10)
'        codsoc = Mid(cad, 37, 13)
        codsoc = 0
        Nifsocio = Mid(cad, 109, 9)
        
        sql2 = "select codmacta from cuentas where nifdatos = " & DBSet(Nifsocio, "T") & " and mid(codmacta,1," & vEmpresaCV.DigitosNivelAnterior - 1 & ") = " & DBSet(vParamAplic.RaizCtaSocCV, "T")
        Set Rs2 = New ADODB.Recordset
        Rs2.Open sql2, ConnContaCV, adOpenForwardOnly, adLockReadOnly, adCmdText
        
        CodmactaSoc = "11"
        If Not Rs2.EOF Then
            CodmactaSoc = DBLet(Rs2!Codmacta)
        End If
        Set Rs2 = Nothing
    
        If Trim(Nifsocio) <> vEmpresa.CifEmpresa Then
            CodmactaVta = vParamAplic.CtaVentaFacCV
            letraser = vParamAplic.LetraSerFCV
        Else
            CodmactaVta = vParamAplic.CtaVentaFacInCV
            letraser = vParamAplic.LetraSerFinCV
        End If
    
    
       Fecha = Mid(cad, 11, 8)
       Fecha = Mid(Fecha, 1, 2) & "/" & Mid(Fecha, 3, 2) & "/" & Mid(Fecha, 6, 4)
       
       TipForpa = Mid(cad, 29, 3)
       If TipForpa = "BCO" Then
           CodForpa = vParamAplic.CodforpaBanCV
       Else
           CodForpa = vParamAplic.CodforpaConCV
       End If
       
       
       Base1 = Mid(cad, 124, 20)
       Base2 = Mid(cad, 144, 20)
       Base3 = Mid(cad, 164, 20)
       Porc1 = Mid(cad, 184, 3)
       Porc2 = Mid(cad, 189, 3)
       Porc3 = Mid(cad, 194, 3)
       Impiv1 = Mid(cad, 214, 20)
       Impiv2 = Mid(cad, 234, 20)
       Impiv3 = Mid(cad, 254, 20)
       TotalFac = Mid(cad, 334, 20)
       
       '[Monica]26/10/2012: tipo de iva tiene que ser 0
       CodIva1 = DevuelveDesdeBDNew(cContaCV, "tiposiva", "codigiva", "porceiva", CStr(Porc1), "N", , "tipodiva", "0", "N")
       CodIva2 = DevuelveDesdeBDNew(cContaCV, "tiposiva", "codigiva", "porceiva", CStr(Porc2), "N", , "tipodiva", "0", "N")
       CodIva3 = DevuelveDesdeBDNew(cContaCV, "tiposiva", "codigiva", "porceiva", CStr(Porc3), "N", , "tipodiva", "0", "N")
       
       c_Base1 = CCur(Mid(Base1, 1, 14) & "," & Mid(Base1, 15, 6))
       c_Base2 = CCur(Mid(Base2, 1, 14) & "," & Mid(Base2, 15, 6))
       c_Base3 = CCur(Mid(Base3, 1, 14) & "," & Mid(Base3, 15, 6))
       c_Impiv1 = CCur(Mid(Impiv1, 1, 14) & "," & Mid(Impiv1, 15, 6))
       c_Impiv2 = CCur(Mid(Impiv2, 1, 14) & "," & Mid(Impiv2, 15, 6))
       c_Impiv3 = CCur(Mid(Impiv3, 1, 14) & "," & Mid(Impiv3, 15, 6))
       c_TotalFac = CCur(Mid(TotalFac, 1, 14) & "," & Mid(TotalFac, 15, 6))
        
        If Trim(Nifsocio) = vEmpresa.CifEmpresa And c_Base1 = 0 Then
            c_Base1 = c_TotalFac
            c_Impiv1 = 0
            Porc1 = 0
           '[Monica]26/10/2012: tipo de iva tiene que ser 0
            CodIva1 = DevuelveDesdeBDNew(cContaCV, "tiposiva", "codigiva", "porceiva", CStr(Porc1), "N", "tipodiva", "0", "N")
        End If
       
       
        Erronea = 0
        If c_Base1 + c_Base2 + c_Base3 + c_Impiv1 + c_Impiv2 + c_Impiv3 <> c_TotalFac Then Erronea = 1
        
        If Erronea = 0 Then
            If (numfactu = "" Or Not IsNumeric(Trim(Right(numfactu, 7)))) Then Erronea = 1
        End If
       
       
       If c_Base1 = 0 Then
           If c_Base2 = 0 Then
               c_Base1 = c_Base3
               c_Impiv1 = c_Impiv3
               Porc1 = Porc3
               CodIva1 = CodIva3
               
               c_Base3 = 0
               c_Impiv3 = 0
               Porc3 = 0
               CodIva3 = ""
           Else
               If c_Base3 = 0 Then
                   c_Base1 = c_Base2
                   c_Impiv1 = c_Impiv2
                   Porc1 = Porc2
                   CodIva1 = CodIva2
                   
                   c_Base2 = 0
                   c_Impiv2 = 0
                   Porc2 = 0
                   CodIva2 = ""
               Else
                   c_Base1 = c_Base3
                   c_Impiv1 = c_Impiv3
                   Porc1 = Porc3
                   CodIva1 = CodIva3
                   
                   c_Base3 = 0
                   c_Impiv3 = 0
                   Porc3 = 0
                   CodIva3 = ""
               End If
           End If
       End If
       If c_Base2 = 0 Then
           If c_Base3 = 0 Then
               ' No hacemos nada
           Else
               c_Base2 = c_Base3
               c_Impiv2 = c_Impiv3
               Porc2 = Porc3
               CodIva2 = CodIva3
               
               c_Base3 = 0
               c_Impiv3 = 0
               Porc3 = 0
               CodIva3 = ""
           End If
       End If
       If c_Base2 = 0 Then
            Porc2 = 0
            CodIva2 = 0
       End If
       If c_Base3 = 0 Then
            Porc3 = 0
            CodIva3 = 0
       End If
       
       
       
       ' insertamos en la tabla de facturas
       SQL = "INSERT INTO cvfacturas (tipofactu,letraser,numfactu,fecfactu,codsocio,nifsocio,codmactasoc,codmactavta,"
       SQL = SQL & "baseimpo,codiva,porciva,cuotaiva,baseimpo2,codiva2,porciva2,cuotaiva2,baseimpo3,codiva3,porciva3,cuotaiva3,totalfac,intconta, codforpa, erronea) VALUES (1," & DBSet(letraser, "T") & ","
       SQL = SQL & DBSet(numfactu, "T") & "," & DBSet(Fecha, "F") & "," & DBSet(codsoc, "N") & "," & DBSet(Nifsocio, "T") & "," & DBSet(CodmactaSoc, "T") & ","
       SQL = SQL & DBSet(CodmactaVta, "T") & ","
       SQL = SQL & DBSet(c_Base1, "N") & "," & DBSet(CodIva1, "N") & "," & DBSet(Porc1, "N") & "," & DBSet(c_Impiv1, "N") & ","
       SQL = SQL & DBSet(c_Base2, "N", "S") & "," & DBSet(CodIva2, "N", "S") & "," & DBSet(Porc2, "N", "S") & "," & DBSet(c_Impiv2, "N", "S") & ","
       SQL = SQL & DBSet(c_Base3, "N", "S") & "," & DBSet(CodIva3, "N", "S") & "," & DBSet(Porc3, "N", "S") & "," & DBSet(c_Impiv3, "N", "S") & ","
       SQL = SQL & DBSet(c_TotalFac, "N") & ",0," & DBSet(CodForpa, "N") & "," & DBSet(Erronea, "N") & ")"
       
       conn.Execute SQL
 
    End If
 
eInsertarLinea:
    If Err.Number <> 0 Then
        InsertarLineaVentas = False
        MsgBox "Error en Insertar Linea Ventas" & Err.Description, vbExclamation
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
    SQL = "delete from tmptickets where codusu = " & vSesion.Codigo
    conn.Execute SQL
End Sub

Private Function ComprobarCtaConta(C As String) As Boolean
    If vParamAplic.NumeroContaCV <> 0 Then
        ComprobarCtaConta = (DevuelveDesdeBDNew(cContaCV, "cuentas", "codmacta", "codmacta", C, "T") <> "")
    End If
End Function

Private Function ComprobarCtaContaV(C As String) As Boolean
    If vParamAplic.NumeroContaCVV <> 0 Then
        ComprobarCtaContaV = (DevuelveDesdeBDNew(cContaCVV, "cuentas", "codmacta", "codmacta", C, "T") <> "")
    End If
End Function

Private Function ComprobarTipoIvaConta(C As String) As Boolean
    If vParamAplic.NumeroContaCV <> 0 Then
        ComprobarTipoIvaConta = (DevuelveDesdeBDNew(cContaCVV, "tiposiva", "codigiva", "porceiva", C, "T") <> "")
    End If
End Function

Private Function ComprobarTipoIvaContaV(C As String) As Boolean
    If vParamAplic.NumeroContaCVV <> 0 Then
        ComprobarTipoIvaContaV = (DevuelveDesdeBDNew(cContaCVV, "tiposiva", "codigiva", "porceiva", C, "T") <> "")
    End If
End Function



' ********* si n'hi han combos a la capçalera ************
Private Sub CargaCombo()
Dim Ini As Integer
Dim Fin As Integer
Dim i As Integer

    ' *** neteje els combos, els pose valor i seleccione el valor per defecte ***
    ' tipo de fichero de importacion de telefonia
    Combo1.Clear
    
    Combo1.AddItem "Varias"
    Combo1.ItemData(Combo1.NewIndex) = 0
    Combo1.AddItem "Venta Tienda"
    Combo1.ItemData(Combo1.NewIndex) = 1
    Combo1.AddItem "Compra Tienda"
    Combo1.ItemData(Combo1.NewIndex) = 2
    
End Sub

Private Sub frmCont_DatoSeleccionado(CadenaSeleccion As String)
    txtCodigo(2).Text = RecuperaValor(CadenaSeleccion, 1)
End Sub

Private Sub imgBuscar_Click(Index As Integer)
    TerminaBloquear
    
    Select Case Index
        Case 0 'Contadores de contabilidad
            If vParamAplic.NumeroContaCVV = 0 Then Exit Sub
            
            Set frmCont = New frmContConta
            frmCont.Facturas = False
            frmCont.DatosADevolverBusqueda = "0|"
            frmCont.CodigoActual = txtCodigo(2).Text
            frmCont.Conexion = cContaCVV
            frmCont.Show vbModal
            Set frmCont = Nothing
            PonerFoco txtCodigo(2)
    End Select

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
            Case 1: KEYBusqueda KeyAscii, 1 'tipo de iva
        End Select
    Else
        KEYpress KeyAscii
    End If

End Sub


Private Sub txtCodigo_LostFocus(Index As Integer)
Dim cad As String, cadTipo As String 'tipo cliente
Dim Conta1 As String
Dim Conta2 As String

    'Quitar espacios en blanco por los lados
    txtCodigo(Index).Text = Trim(txtCodigo(Index).Text)
    
    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub

    Select Case Index
        Case 2 'Letra Contador de la contabilidad
            If txtCodigo(Index).Text <> "" Then
                cad = DevuelveDesdeBDNew(cContaCVV, "contadores", "nomregis", "tiporegi", Trim(txtCodigo(Index).Text), "T")
                If cad = "" Then
                    MsgBox "La letra de serie no existe en contabilidad. Revise.", vbExclamation
                    PonerFoco txtCodigo(Index)
                End If
            End If
        
       
    End Select
End Sub





