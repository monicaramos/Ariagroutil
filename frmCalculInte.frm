VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmCalculInte 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cálculo de Intereses"
   ClientHeight    =   6045
   ClientLeft      =   45
   ClientTop       =   2430
   ClientWidth     =   6990
   Icon            =   "frmCalculInte.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6045
   ScaleWidth      =   6990
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
      Height          =   6015
      Left            =   45
      TabIndex        =   9
      Top             =   0
      Width           =   6915
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   6
         Left            =   2730
         Locked          =   -1  'True
         TabIndex        =   20
         Text            =   "Text5"
         Top             =   1230
         Width           =   3135
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   5
         Left            =   2730
         Locked          =   -1  'True
         TabIndex        =   19
         Text            =   "Text5"
         Top             =   855
         Width           =   3135
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   0
         Left            =   1845
         MaxLength       =   10
         TabIndex        =   4
         Top             =   3285
         Width           =   1050
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   6
         Left            =   1845
         MaxLength       =   6
         TabIndex        =   1
         Top             =   1215
         Width           =   830
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   5
         Left            =   1845
         MaxLength       =   6
         TabIndex        =   0
         Top             =   840
         Width           =   830
      End
      Begin VB.TextBox txtCodigo 
         Height          =   285
         Index           =   4
         Left            =   1845
         MaxLength       =   30
         TabIndex        =   5
         Top             =   3780
         Width           =   4410
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   3
         Left            =   1845
         MaxLength       =   10
         TabIndex        =   3
         Top             =   2400
         Width           =   1050
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   2
         Left            =   1845
         MaxLength       =   10
         TabIndex        =   2
         Top             =   2040
         Width           =   1050
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Left            =   4905
         TabIndex        =   8
         Top             =   5160
         Width           =   975
      End
      Begin VB.CommandButton cmdAceptar 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   3720
         TabIndex        =   7
         Top             =   5160
         Width           =   975
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   1
         Left            =   1845
         MaxLength       =   6
         TabIndex        =   6
         Top             =   4320
         Width           =   830
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   6
         Left            =   1530
         MouseIcon       =   "frmCalculInte.frx":000C
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar cliente"
         Top             =   1230
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   5
         Left            =   1545
         MouseIcon       =   "frmCalculInte.frx":015E
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar cliente"
         Top             =   855
         Width           =   240
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   0
         Left            =   1530
         Picture         =   "frmCalculInte.frx":02B0
         ToolTipText     =   "Buscar fecha"
         Top             =   3285
         Width           =   240
      End
      Begin VB.Label Label4 
         Caption         =   "Fecha de Movimiento"
         ForeColor       =   &H00972E0B&
         Height          =   255
         Index           =   5
         Left            =   585
         TabIndex        =   18
         Top             =   3015
         Width           =   1815
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Codigo Avnics"
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
         Index           =   4
         Left            =   585
         TabIndex        =   17
         Top             =   600
         Width           =   1005
      End
      Begin VB.Label Label4 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   3
         Left            =   960
         TabIndex        =   16
         Top             =   1215
         Width           =   420
      End
      Begin VB.Label Label4 
         Caption         =   "Desde"
         Height          =   195
         Index           =   2
         Left            =   960
         TabIndex        =   15
         Top             =   840
         Width           =   465
      End
      Begin VB.Label Label4 
         Caption         =   "Concepto "
         ForeColor       =   &H00972E0B&
         Height          =   255
         Index           =   0
         Left            =   585
         TabIndex        =   14
         Top             =   3780
         Width           =   1815
      End
      Begin VB.Label Label4 
         Caption         =   "Fecha Liquidación"
         ForeColor       =   &H00972E0B&
         Height          =   255
         Index           =   16
         Left            =   585
         TabIndex        =   13
         Top             =   1800
         Width           =   1815
      End
      Begin VB.Label Label4 
         Caption         =   "Desde"
         Height          =   195
         Index           =   15
         Left            =   960
         TabIndex        =   12
         Top             =   2040
         Width           =   465
      End
      Begin VB.Label Label4 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   14
         Left            =   960
         TabIndex        =   11
         Top             =   2400
         Width           =   420
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   2
         Left            =   1530
         Picture         =   "frmCalculInte.frx":033B
         ToolTipText     =   "Buscar fecha"
         Top             =   2040
         Width           =   240
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   3
         Left            =   1530
         Picture         =   "frmCalculInte.frx":03C6
         ToolTipText     =   "Buscar fecha"
         Top             =   2400
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Porc. Retención"
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
         Left            =   585
         TabIndex        =   10
         Top             =   4320
         Width           =   1140
      End
   End
End
Attribute VB_Name = "frmCalculInte"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' +-+-+-+-+-+-+-+-+-+-+-
' +-+- Autor MANOLO +-+-
' +-+-+-+-+-+-+-+-+-+-+-

Option Explicit

Public NumCod As String 'Para indicar cod. Traspaso,Movimiento, etc. que llama
                        'Para indicar nº oferta a imprimir

Public CadTag As String 'Cadena con el Tag del campo que se va a poner en D/H en los listados
                        'Se necesita si el tipo de codigo es texto

    
Private Conexion As Byte
'1.- Conexión a BD Ariges  2.- Conexión a BD Conta

Private HaDevueltoDatos As Boolean


Private WithEvents frmavn As frmAvnics 'Avnics
Attribute frmavn.VB_VarHelpID = -1
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
Dim Tabla As String
Dim Codigo As String 'Código para FormulaSelection de Crystal Report
Dim TipCod As String
Dim Orden1 As String 'Campo de Ordenacion (por codigo) para Cristal Report
Dim Orden2 As String 'Campo de Ordenacion (por nombre) para Cristal Report

Dim PrimeraVez As Boolean

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
Dim i As Byte
Dim NReg As Long
Dim cadwhere As String
Dim cad As String


InicializarVbles
    
    If Not DatosOk Then Exit Sub
    
    '========= PARAMETROS  =============================
    'Añadir el parametro de Empresa
    cadParam = cadParam & "|pEmpresa=""" & vEmpresa.nomEmpre & """|"
    numParam = numParam + 1
    
    
    'D/H codigo Avnics
    cDesde = Trim(txtCodigo(5).Text)
    cHasta = Trim(txtCodigo(6).Text)
    nDesde = txtNombre(5).Text
    nHasta = txtNombre(6).Text
    If Not (cDesde = "" And cHasta = "") Then
        'Cadena para seleccion Desde y Hasta
        Codigo = "{" & Tabla & ".codavnic}"
        TipCod = "N"
        If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHavnics= """) Then Exit Sub
    End If
    
    'D/H Fecha Alta
    cDesde = Trim(txtCodigo(2).Text)
    cHasta = Trim(txtCodigo(3).Text)
    If Not (cDesde = "" And cHasta = "") Then
        'Cadena para seleccion Desde y Hasta
        Codigo = "{" & Tabla & ".fechalta}"
        TipCod = "F"
        If Not PonerDesdeHasta(cDesde, cHasta, "", "", "pDHfechaLiq= """) Then Exit Sub
    End If
    
'    AnyadirAFormula cadFormula, cDesde
'    AnyadirAFormula cadSelect, cDesde
    cadFormula = ""
    cadSelect = ""
  
  
    cadParam = cadParam & "pFechaMov= """ & txtCodigo(0).Text & """|"
    numParam = numParam + 1
    cadParam = cadParam & "pReten= " & DBSet(txtCodigo(1).Text, "N") & "|"
    numParam = numParam + 1
    
    AnyadirAFormula cadFormula, "{tmpinformes.codusu} = " & vSesion.Codigo
    
    ' el avnic no tiene que haber sido cancelado en el ejercicio
    cadwhere = "where anoejerc = year(" & DBSet(txtCodigo(0).Text, "F") & ") and codialta <> 2 "
    
    If txtCodigo(5).Text <> "" Then cadwhere = cadwhere & " and avnic.codavnic >= " & DBSet(txtCodigo(5).Text, "N")
    If txtCodigo(6).Text <> "" Then cadwhere = cadwhere & " and avnic.codavnic <= " & DBSet(txtCodigo(6).Text, "N")
    cadwhere = cadwhere & " and (( (1 = 1) "
    If txtCodigo(2).Text <> "" Then cadwhere = cadwhere & " and avnic.fechalta >= " & DBSet(txtCodigo(2).Text, "F")
    If txtCodigo(3).Text <> "" Then cadwhere = cadwhere & " and avnic.fechalta <= " & DBSet(txtCodigo(3).Text, "F")
    cadwhere = cadwhere & " ) or ( (1 = 1) "
    If txtCodigo(2).Text <> "" Then cadwhere = cadwhere & " and avnic.fechalta <= " & DBSet(txtCodigo(2).Text, "F")
    If txtCodigo(2).Text <> "" Then cadwhere = cadwhere & " and avnic.fechavto >= " & DBSet(txtCodigo(2).Text, "F")
    If txtCodigo(3).Text <> "" Then cadwhere = cadwhere & " and avnic.fechavto <= " & DBSet(txtCodigo(3).Text, "F")
    cadwhere = cadwhere & " ) or ( (1 = 1) "
    If txtCodigo(2).Text <> "" Then cadwhere = cadwhere & " and avnic.fechalta <= " & DBSet(txtCodigo(2).Text, "F")
    If txtCodigo(3).Text <> "" Then cadwhere = cadwhere & " and avnic.fechavto > " & DBSet(txtCodigo(3).Text, "F")
    cadwhere = cadwhere & "))"
    
  
    cad = "select count(*) from avnic  " & cadwhere
    NReg = TotalRegistros(cad)
    If NReg <> 0 Then
       If CargarTablaIntermedia(cadwhere) Then
            cadTitulo = "Intereses Avnics"
            cadNombreRPT = "rCalculInt.rpt"
            LlamarImprimir
            cad = "Impresión correcta para actualizar"
            If MsgBox(cad, vbYesNo + vbDefaultButton2 + vbQuestion) = vbYes Then
                If ActualizarTablas Then
                    MsgBox "Proceso realizado correctamente", vbExclamation
                    cmdCancel_Click
                End If
            End If
       End If
    Else
        MsgBox "No existen datos entre esos límites. Reintroduzca.", vbExclamation
    End If
    
    
    
'    'Comprobar si hay registros a Mostrar antes de abrir el Informe
'    cadTABLA = Tabla '& " INNER JOIN ssocio ON " & Tabla & ".codsocio=ssocio.codsocio "
'
'    If HayRegParaInforme(cadTABLA, cadSelect) Then
'       cadTitulo = "Informe de Avnics"
'       cadNombreRPT = "rInfAvnics.rpt"
'       LlamarImprimir
'       'AbrirVisReport
'    End If
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
    If PrimeraVez Then
        PrimeraVez = False
        PonerFoco txtCodigo(5)
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
Dim h As Integer, w As Integer
Dim List As Collection

    PrimeraVez = True
    Limpiar Me

    'IMAGES para busqueda
     Me.imgBuscar(5).Picture = frmPpal.imgListImages16.ListImages(1).Picture
     Me.imgBuscar(6).Picture = frmPpal.imgListImages16.ListImages(1).Picture

    '###Descomentar
'    CommitConexion
         
    FrameCobrosVisible True, h, w
    indFrame = 5
    Tabla = "avnic"
            
    txtCodigo(0).Text = Format(Now, "dd/mm/yyyy")
    txtCodigo(1).Text = Format(vParamAplic.Porcrete, "##0.00")
            
            
    'Esto se consigue poneinedo el cancel en el opcion k corresponda
    Me.cmdCancel.Cancel = True
    Me.Width = w + 70
    Me.Height = h + 350
End Sub

Private Sub frmC_Selec(vFecha As Date)
 'Fecha
    txtCodigo(CByte(imgFec(2).Tag)).Text = Format(vFecha, "dd/MM/yyyy")
End Sub

Private Sub frmAvn_DatoSeleccionado(CadenaSeleccion As String)
'Form de Consulta de Clientes
    txtCodigo(indCodigo).Text = Format(RecuperaValor(CadenaSeleccion, 1), "000000")
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
    imgFec(2).Tag = Index 'independentment de les dates que tinga, sempre pose l'index en la 27
    If txtCodigo(Index).Text <> "" Then frmC.NovaData = txtCodigo(Index).Text

    frmC.Show vbModal
    Set frmC = Nothing
    PonerFoco txtCodigo(CByte(imgFec(2).Tag) + 2)
    ' ***************************
End Sub

Private Sub ListView1_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub imgBuscar_Click(Index As Integer)
   Select Case Index
        Case 5, 6 'AVNICS
            AbrirFrmAvnics (Index)
        
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
'15/02/2007
'    KEYpress KeyAscii
' ahora
    If KeyAscii = teclaBuscar Then
        Select Case Index
            Case 5: KEYBusqueda KeyAscii, 5 'codigo desde
            Case 6: KEYBusqueda KeyAscii, 6 'codigo hasta
            Case 2: KEYFecha KeyAscii, 2    'fecha liquidacion desde
            Case 3: KEYFecha KeyAscii, 3    'fecha liquidacion hasta
            Case 0: KEYFecha KeyAscii, 0    'fecha de movimiento
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
            
        Case 5, 6 'codigo
            txtNombre(Index).Text = PonerNombreDeCod(txtCodigo(Index), "avnic", "nombrper", "codavnic", "N")
            If txtCodigo(Index).Text <> "" Then txtCodigo(Index).Text = Format(txtCodigo(Index).Text, "000000")
        
        Case 0, 2, 3 'FECHAS
            If txtCodigo(Index).Text <> "" Then PonerFormatoFecha txtCodigo(Index)
            
        Case 1 'porcentaje de retencion
            If txtCodigo(Index).Text <> "" Then PonerFormatoDecimal txtCodigo(Index), 7
              
  End Select
End Sub

Private Sub FrameCobrosVisible(visible As Boolean, ByRef h As Integer, ByRef w As Integer)
'Frame para los cobros a clientes por fecha vencimiento
    Me.FrameCobros.visible = visible
    If visible = True Then
        Me.FrameCobros.Top = -90
        Me.FrameCobros.Left = 0
        Me.FrameCobros.Height = 6015
        Me.FrameCobros.Width = 6555
        w = Me.FrameCobros.Width
        h = Me.FrameCobros.Height
    End If
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

Private Sub AbrirFrmAvnics(indice As Integer)
    indCodigo = indice
    Set frmavn = New frmAvnics
    frmavn.DatosADevolverBusqueda = "0|4|"
    frmavn.DeConsulta = True
    frmavn.CodigoActual = txtCodigo(indCodigo)
    frmavn.Show vbModal
    Set frmavn = Nothing
End Sub
 
Private Sub AbrirVisReport()
    Screen.MousePointer = vbHourglass
    CadenaDesdeOtroForm = ""
    With frmVisReport
        .FormulaSeleccion = cadFormula
'        .SoloImprimir = (Me.OptVisualizar(indFrame).Value = 1)
        .OtrosParametros = cadParam
        .NumeroParametros = numParam
        '##descomen
'        .MostrarTree = MostrarTree
'        .Informe = MIPATH & Nombre
'        .InfConta = InfConta
        '##
        
'        If NombreSubRptConta <> "" Then
'            .SubInformeConta = NombreSubRptConta
'        Else
'            .SubInformeConta = ""
'        End If
        '##descomen
'        .ConSubInforme = ConSubInforme
        '##
        .Opcion = ""
'        .ExportarPDF = (chkEMAIL.Value = 1)
        .Show vbModal
    End With
    
'    If Me.chkEMAIL.Value = 1 Then
'    '####Descomentar
'        If CadenaDesdeOtroForm <> "" Then frmEMail.Show vbModal
'    End If
    Unload Me
End Sub

Private Sub AbrirEMail()
    If CadenaDesdeOtroForm <> "" Then frmEMail.Show vbModal
End Sub


Private Function DatosOk() As Boolean
Dim b As Boolean
    b = True
    ' la fecha de vencimiento no debe de ser nula
    If txtCodigo(0).Text = "" Then
        MsgBox "La fecha de movimiento debe de tener un valor. Reintroduzca.", vbExclamation
        b = False
        PonerFoco txtCodigo(0)
    End If
    DatosOk = b
    
End Function

Private Function CargarTablaIntermedia(cadwhere As String) As Boolean
Dim sql As String
Dim Rs As ADODB.Recordset
Dim Rs2 As ADODB.Recordset
Dim Rs3 As ADODB.Recordset
Dim AntCodavnic As Long
Dim ActCodavnic As Long
Dim v_fecha As String
Dim DesFec As String
Dim HasFec As String
Dim DifDia As Long
Dim Pasado As Boolean
Dim cad As String
Dim Importe As Currency
Dim Sql2 As String
Dim Sql3 As String

    On Error GoTo eCargarTablaIntermedia

    CargarTablaIntermedia = False
    
    If Not BorrarTablaIntermedia Then Exit Function

    Set Rs = New ADODB.Recordset
    
    sql = "select codavnic, importes, porcinte, anoejerc, fechalta, fechavto from avnic " & cadwhere
    sql = sql & " order by codavnic"
    
    Rs.Open sql, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    AntCodavnic = DBLet(Rs!codavnic, "N")
    ActCodavnic = AntCodavnic
    Pasado = False
    While Not Rs.EOF
        AntCodavnic = ActCodavnic
        ActCodavnic = Rs!codavnic
        
        If ActCodavnic <> AntCodavnic Then Pasado = False
        
        If Not Pasado Then
                ' obtenemos la maxima fehca de movimiento
                Set Rs3 = New ADODB.Recordset
                
                cad = "select max(fechamov) from movim where codavnic = " & DBSet(Rs!codavnic, "N")
                Rs3.Open cad, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
                
                v_fecha = ""
                If Not Rs3.EOF Then v_fecha = DBLet(Rs3.Fields(0).Value, "F")
                Set Rs3 = Nothing
                
                If v_fecha = "" Then DesFec = DBLet(Rs!fechalta, "F")
                If v_fecha <> "" Then DesFec = v_fecha
    
                HasFec = DBLet(Rs!fechavto, "F")
                If txtCodigo(3).Text <> "" Then
                    If DBLet(Rs!fechavto, "F") < CDate(txtCodigo(3).Text) Then
                        HasFec = DBLet(Rs!fechavto, "F")
                    Else
                        HasFec = txtCodigo(3).Text
                    End If
                End If
                DifDia = CDate(HasFec) - CDate(DesFec)
                If DifDia < 0 Then DifDia = 0
                
                If DBLet(Rs!importes, "N") <> 0 Then
                    Importe = Round2(DBLet(Rs!importes, "N") * DBLet(Rs!Porcinte, "N") * 0.01 * DifDia / 365, 2)
                    Sql2 = "insert into tmpinformes (codusu, codigo1, campo1, fecha1, fecha2, importe1, importe2, nombre1) values (" & DBSet(vSesion.Codigo, "N") & ","
                    Sql2 = Sql2 & DBSet(Rs!codavnic, "N") & "," & DBSet(Rs!anoejerc, "N") & "," & DBSet(DesFec, "F") & "," & DBSet(HasFec, "F") & ","
                    Sql2 = Sql2 & DBSet(Importe, "N") & ","
                    
                    Set Rs2 = New ADODB.Recordset
                    Sql3 = "select nombrper, importes from avnic where codavnic = " & DBSet(Rs!codavnic, "N") & " and anoejerc = " & Year(CDate(txtCodigo(0).Text))
                    Rs2.Open Sql3, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
                    
                    If Rs2.EOF Then
                        Sql2 = Sql2 & "0, null)"
                    Else
                        Sql2 = Sql2 & DBSet(Rs2!importes, "N") & "," & DBSet(Rs2!nombrper, "T") & ")"
                    End If
                
                    conn.Execute Sql2
                End If
                
                Pasado = True
        End If
        
        Rs.MoveNext
    Wend
    CargarTablaIntermedia = True
    Exit Function
    
eCargarTablaIntermedia:
    MuestraError Err.Number, "Error cargando la tabla intermedia. Llame a soporte.", Err.Description
End Function


Private Function ActualizarTablas() As Boolean
Dim sql As String
Dim Rs3 As ADODB.Recordset
Dim cad As String
Dim v_import As Currency
Dim t_import As Currency

    On Error GoTo eActualizarTablas

    ActualizarTablas = False

    conn.BeginTrans


    Set Rs3 = New ADODB.Recordset
    cad = "select codigo1, campo1, importe1 from tmpinformes where codusu = " & vSesion.Codigo
    Rs3.Open cad, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    While Not Rs3.EOF
        v_import = 0
        t_import = 0
        v_import = Round2(Rs3!importe1 * CCur(txtCodigo(1).Text) * 0.01, 2)
        t_import = Rs3!importe1 - v_import
        
        If t_import <> 0 Then
            sql = ""
            sql = DevuelveDesdeBDNew(cPTours, "movim", "codavnic", "codavnic", Rs3!codigo1, "N", , "fechamov", txtCodigo(0), "F", "anoejerc", Year(CDate(txtCodigo(0).Text)), "N")
            If sql = "" Then
                sql = "insert into movim (codavnic, fechamov, concepto, timporte, intconta, anoejerc, timport1, timport2) "
                sql = sql & "values (" & DBSet(Rs3!codigo1, "N") & "," & DBSet(txtCodigo(0).Text, "F") & "," & DBSet(txtCodigo(4).Text, "T") & ","
                sql = sql & DBSet(t_import, "N") & ",0," & DBSet(Year(CDate(txtCodigo(0).Text)), "N") & ","
                sql = sql & DBSet(Rs3!importe1, "N") & "," & DBSet(v_import, "N") & ")"
                
                conn.Execute sql
            End If
            sql = "update avnic set imporper = imporper + " & DBSet(Rs3!importe1, "N") & ","
            sql = sql & " imporret = imporret + " & DBSet(v_import, "N")
            sql = sql & " where codavnic = " & DBSet(Rs3!codigo1, "N") & " and anoejerc = " & DBSet(Year(CDate(txtCodigo(0).Text)), "N")
            
            conn.Execute sql
                
        End If
    
        Rs3.MoveNext
    Wend
    Set Rs3 = Nothing
    
    conn.CommitTrans
    ActualizarTablas = True
    Exit Function
    
eActualizarTablas:
    If Err.Number <> 0 Then
        MuestraError Err.Number, "Error en la actualizacion de datos: " & Err.Description
        conn.RollbackTrans
    End If
End Function
