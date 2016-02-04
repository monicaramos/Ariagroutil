VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmCancelAvnic 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cancelación de Avnics"
   ClientHeight    =   4755
   ClientLeft      =   45
   ClientTop       =   2430
   ClientWidth     =   6990
   Icon            =   "frmCancelAvnic.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4755
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
      Height          =   4710
      Left            =   45
      TabIndex        =   9
      Top             =   0
      Width           =   6915
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   9
         Left            =   5400
         MaxLength       =   6
         TabIndex        =   6
         Top             =   3375
         Width           =   830
      End
      Begin VB.Frame FrameResultado 
         Caption         =   "Cálculo a Ingresar"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   1680
         Left            =   3285
         TabIndex        =   19
         Top             =   1170
         Width           =   2895
         Begin VB.TextBox txtCodigo 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            Height          =   285
            Index           =   8
            Left            =   1170
            MaxLength       =   6
            TabIndex        =   25
            Top             =   1215
            Width           =   1320
         End
         Begin VB.TextBox txtCodigo 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            Height          =   285
            Index           =   7
            Left            =   1170
            MaxLength       =   6
            TabIndex        =   24
            Top             =   810
            Width           =   1320
         End
         Begin VB.TextBox txtCodigo 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            Height          =   285
            Index           =   1
            Left            =   1170
            MaxLength       =   6
            TabIndex        =   23
            Top             =   405
            Width           =   1320
         End
         Begin VB.Label Label3 
            Caption         =   "Neto"
            Height          =   285
            Left            =   315
            TabIndex        =   22
            Top             =   1215
            Width           =   915
         End
         Begin VB.Label Label2 
            Caption         =   "Retención"
            Height          =   285
            Left            =   315
            TabIndex        =   21
            Top             =   810
            Width           =   915
         End
         Begin VB.Label Label1 
            Caption         =   "Bruto"
            Height          =   285
            Left            =   315
            TabIndex        =   20
            Top             =   405
            Width           =   915
         End
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   6
         Left            =   2385
         MaxLength       =   6
         TabIndex        =   5
         Top             =   3375
         Width           =   830
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   5
         Left            =   2730
         Locked          =   -1  'True
         TabIndex        =   17
         Text            =   "Text5"
         Top             =   675
         Width           =   3405
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   0
         Left            =   1800
         MaxLength       =   10
         TabIndex        =   3
         Top             =   2385
         Width           =   1050
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   5
         Left            =   1845
         MaxLength       =   6
         TabIndex        =   0
         Top             =   660
         Width           =   830
      End
      Begin VB.TextBox txtCodigo 
         Height          =   285
         Index           =   4
         Left            =   1800
         MaxLength       =   30
         TabIndex        =   4
         Top             =   2925
         Width           =   4410
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   3
         Left            =   1800
         MaxLength       =   10
         TabIndex        =   2
         Top             =   1725
         Width           =   1050
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   2
         Left            =   1800
         MaxLength       =   10
         TabIndex        =   1
         Top             =   1365
         Width           =   1050
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Left            =   5355
         TabIndex        =   8
         Top             =   4005
         Width           =   975
      End
      Begin VB.CommandButton cmdAceptar 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   4170
         TabIndex        =   7
         Top             =   4005
         Width           =   975
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Porcentaje Retención"
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
         Index           =   3
         Left            =   3645
         TabIndex        =   26
         Top             =   3420
         Width           =   1545
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Porcentaje Penalización"
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
         Left            =   540
         TabIndex        =   18
         Top             =   3420
         Width           =   1695
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   5
         Left            =   1545
         MouseIcon       =   "frmCancelAvnic.frx":000C
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar cliente"
         Top             =   675
         Width           =   240
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   0
         Left            =   1485
         Picture         =   "frmCancelAvnic.frx":015E
         ToolTipText     =   "Buscar fecha"
         Top             =   2385
         Width           =   240
      End
      Begin VB.Label Label4 
         Caption         =   "Fecha de Cancelación"
         ForeColor       =   &H00972E0B&
         Height          =   255
         Index           =   5
         Left            =   540
         TabIndex        =   16
         Top             =   2115
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
         TabIndex        =   15
         Top             =   420
         Width           =   1005
      End
      Begin VB.Label Label4 
         Caption         =   "Desde"
         Height          =   195
         Index           =   2
         Left            =   960
         TabIndex        =   14
         Top             =   660
         Width           =   465
      End
      Begin VB.Label Label4 
         Caption         =   "Concepto "
         ForeColor       =   &H00972E0B&
         Height          =   255
         Index           =   0
         Left            =   540
         TabIndex        =   13
         Top             =   2970
         Width           =   1815
      End
      Begin VB.Label Label4 
         Caption         =   "Período Liquidación"
         ForeColor       =   &H00972E0B&
         Height          =   255
         Index           =   16
         Left            =   540
         TabIndex        =   12
         Top             =   1125
         Width           =   1815
      End
      Begin VB.Label Label4 
         Caption         =   "Desde"
         Height          =   195
         Index           =   15
         Left            =   915
         TabIndex        =   11
         Top             =   1365
         Width           =   465
      End
      Begin VB.Label Label4 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   14
         Left            =   915
         TabIndex        =   10
         Top             =   1725
         Width           =   420
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   2
         Left            =   1485
         Picture         =   "frmCancelAvnic.frx":01E9
         ToolTipText     =   "Buscar fecha"
         Top             =   1365
         Width           =   240
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   3
         Left            =   1485
         Picture         =   "frmCancelAvnic.frx":0274
         ToolTipText     =   "Buscar fecha"
         Top             =   1725
         Width           =   240
      End
   End
End
Attribute VB_Name = "frmCancelAvnic"
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
Dim cerrar As Boolean

    KEYpressGnral KeyAscii, 2, cerrar
    If cerrar Then Unload Me
End Sub

Private Sub cmdAceptar_Click()
Dim cDesde As String, cHasta As String 'cadena codigo Desde/Hasta
Dim nDesde As String, nHasta As String 'cadena Descripcion Desde/Hasta
Dim cadTABLA As String, cOrden As String
Dim i As Byte
Dim NReg As Long
Dim cadWhere As String
Dim Cad As String

    If Not DatosOk Then Exit Sub
    
    ' el avnic no tiene que haber sido cancelado en el ejercicio
    cadWhere = "anoejerc = year(" & DBSet(txtCodigo(0).Text, "F") & ") and codialta <> 2 "
    cadWhere = cadWhere & " and codavnic = " & DBSet(txtCodigo(5).Text, "N")
    
    Cad = "select count(*) from avnic where " & cadWhere
    NReg = TotalRegistros(Cad)
    
    If NReg <> 0 Then
        If CalculoPenalizacion(cadWhere) Then
'            MsgBox "Proceso realizado correctamente", vbExclamation
            VisualizarFrameResultado (True)
'            cmdCancel_Click
        End If
    Else
        MsgBox "No existen datos entre esos límites. Reintroduzca.", vbExclamation
    End If
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
     
    '###Descomentar
'    CommitConexion
         
    FrameCobrosVisible True, h, w
    indFrame = 5
    Tabla = "avnic"
            
    'fecha de cancelacion
    txtCodigo(0).Text = Format(Now, "dd/mm/yyyy")
    txtCodigo(9).Text = Format(vParamAplic.Porcrete, "##0.00")
    
    'periodo de liquidacion
    Select Case Month(CDate(txtCodigo(0).Text))
        Case 1 To 3
            txtCodigo(2).Text = "01/01/" & Format(Year(CDate(txtCodigo(0).Text)), "0000")
            txtCodigo(3).Text = "31/03/" & Format(Year(CDate(txtCodigo(0).Text)), "0000")
        Case 4 To 6
            txtCodigo(2).Text = "01/04/" & Format(Year(CDate(txtCodigo(0).Text)), "0000")
            txtCodigo(3).Text = "30/06/" & Format(Year(CDate(txtCodigo(0).Text)), "0000")
        Case 7 To 9
            txtCodigo(2).Text = "01/07/" & Format(Year(CDate(txtCodigo(0).Text)), "0000")
            txtCodigo(3).Text = "30/09/" & Format(Year(CDate(txtCodigo(0).Text)), "0000")
        Case 10 To 12
            txtCodigo(2).Text = "01/10/" & Format(Year(CDate(txtCodigo(0).Text)), "0000")
            txtCodigo(3).Text = "31/12/" & Format(Year(CDate(txtCodigo(0).Text)), "0000")
    End Select
                
    VisualizarFrameResultado False
                
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
        Case 5 'AVNICS
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
            Case 2: KEYFecha KeyAscii, 2    'fecha liquidacion desde
            Case 3: KEYFecha KeyAscii, 3    'fecha liquidacion hasta
            Case 0: KEYFecha KeyAscii, 0    'fecha de cancelacion
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
Dim Cad As String, cadTipo As String 'tipo cliente

    'Quitar espacios en blanco por los lados
    txtCodigo(Index).Text = Trim(txtCodigo(Index).Text)
    
    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub

    
    Select Case Index
            
        Case 5 'codigo
            txtNombre(Index).Text = PonerNombreDeCod(txtCodigo(Index), "avnic", "nombrper", "codavnic", "N")
            If txtCodigo(Index).Text <> "" Then txtCodigo(Index).Text = Format(txtCodigo(Index).Text, "0000")
        
        Case 0, 2, 3 'FECHAS
            If txtCodigo(Index).Text <> "" Then PonerFormatoFecha txtCodigo(Index)
            
        Case 6, 9 'porcentaje de penalizacion
            If txtCodigo(Index).Text <> "" Then PonerFormatoDecimal txtCodigo(Index), 7
              
  End Select
End Sub

Private Sub FrameCobrosVisible(visible As Boolean, ByRef h As Integer, ByRef w As Integer)
'Frame para los cobros a clientes por fecha vencimiento
    Me.FrameCobros.visible = visible
    If visible = True Then
        Me.FrameCobros.Top = -90
        Me.FrameCobros.Left = 0
        Me.FrameCobros.Height = 4710
        Me.FrameCobros.Width = 6915
        w = Me.FrameCobros.Width
        h = Me.FrameCobros.Height
    End If
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
    ' el periodo de liquidacion debe de tener un valor
    If txtCodigo(2).Text = "" Or txtCodigo(3).Text = "" Then
        MsgBox "El período de liquidación debe de tener valores. Reintroduzca.", vbExclamation
        b = False
        PonerFoco txtCodigo(2)
    End If
    
    ' la fecha de cancelacion no debe de ser nula
    If txtCodigo(0).Text = "" Then
        MsgBox "La fecha de cancelación debe de tener un valor. Reintroduzca.", vbExclamation
        b = False
        PonerFoco txtCodigo(0)
    Else
        ' la fecha de cancelacion ha de estar comprendida dentro del periodo de liquidacion
        If Not (CDate(txtCodigo(0).Text) >= CDate(txtCodigo(2).Text) And CDate(txtCodigo(0).Text) <= CDate(txtCodigo(3).Text)) Then
            MsgBox "La fecha de cancelación del avnic debe de estar comprendida dentro del período de liquidación", vbExclamation
            b = False
            PonerFoco txtCodigo(0)
        End If
    End If
    
    DatosOk = b
    
End Function

Private Function CalculoPenalizacion(cadWhere As String) As Boolean
Dim Sql As String
Dim DiasPen As Integer
Dim DiasInt As Integer
Dim Intereses As Currency
Dim Penalizacion As Currency
Dim bruto As Currency
Dim neto As Currency
Dim retencion As Currency

Dim RS As ADODB.Recordset
Dim b As Boolean
Dim Mens As String

    On Error GoTo eCalculoPenalizacion

    Set RS = New ADODB.Recordset

    CalculoPenalizacion = False
    
    conn.BeginTrans
    b = True
    Sql = "select * from avnic where " & cadWhere
    RS.Open Sql, conn, adOpenForwardOnly, adLockOptimistic, adCmdText

    txtCodigo(1).Text = Format(0, "###,###,##0.00")
    txtCodigo(7).Text = Format(0, "###,###,##0.00")
    txtCodigo(8).Text = Format(0, "###,###,##0.00")

    If Not RS.EOF Then
        DiasInt = CDate(txtCodigo(0).Text) - CDate(txtCodigo(2).Text) + 1
        DiasPen = CDate(txtCodigo(3).Text) - CDate(txtCodigo(0).Text)
            
        Intereses = Round2(DBLet(RS!importes, "N") * DBLet(RS!Porcinte, "N") / 100 * DiasInt / 365, 2)
        Penalizacion = Round2(DBLet(RS!importes, "N") * ImporteSinFormato(txtCodigo(6).Text) / 100 * DiasPen / 365, 2)
            
        bruto = Intereses - Penalizacion
        If bruto > 0 Then
            retencion = Round2(bruto * ImporteSinFormato(txtCodigo(9).Text) / 100, 2)
            neto = bruto - retencion
            
            'cargamos las variables para visualizarlas posteriormente
            txtCodigo(1).Text = Format(bruto, "###,###,##0.00")
            txtCodigo(7).Text = Format(retencion, "###,###,##0.00")
            txtCodigo(8).Text = Format(neto, "###,###,##0.00")
            
            Mens = "Insertar movimiento y actualizar avnic "
            b = InsertarMovimiento(RS!codavnic, RS!anoejerc, bruto, retencion, neto, Mens)
        End If
        If b Then
            Sql = "update avnic set codialta = 2 "
            Sql = Sql & "where " & cadWhere
            conn.Execute Sql
        End If
    End If
    If b Then
        CalculoPenalizacion = b
        conn.CommitTrans
        Exit Function
    End If
eCalculoPenalizacion:
    If Err.Number <> 0 Or Not b Then
        MuestraError Err.Number, Mens
        conn.RollbackTrans
    End If
End Function

Private Function InsertarMovimiento(codavnic As Long, anoejerc As Integer, bruto As Currency, retencion As Currency, neto As Currency, ByRef Mens As String) As Boolean
Dim Sql As String

    On Error GoTo eInsertarMovimiento
    
    InsertarMovimiento = False
    
    Sql = ""
    Sql = DevuelveDesdeBDNew(cPTours, "movim", "codavnic", "codavnic", CStr(codavnic), "N", , "fechamov", txtCodigo(0).Text, "F", "anoejerc", CStr(anoejerc), "N")
    If Sql = "" Then
        Sql = "insert into movim (codavnic, fechamov, concepto, timporte, intconta, anoejerc, timport1, timport2) "
        Sql = Sql & "values (" & DBSet(codavnic, "N") & "," & DBSet(txtCodigo(0).Text, "F") & "," & DBSet(txtCodigo(4).Text, "T") & ","
        Sql = Sql & DBSet(neto, "N") & ",0," & DBSet(anoejerc, "N") & ","
        Sql = Sql & DBSet(bruto, "N") & "," & DBSet(retencion, "N") & ")"
        
        conn.Execute Sql
    End If
    Sql = "update avnic set imporper = imporper + " & DBSet(neto, "N") & ","
    Sql = Sql & " imporret = imporret + " & DBSet(retencion, "N")
    Sql = Sql & " where codavnic = " & DBSet(codavnic, "N") & " and anoejerc = " & DBSet(anoejerc, "N")
    
    conn.Execute Sql
    
    InsertarMovimiento = True
    Exit Function
    
eInsertarMovimiento:
    If Err.Number <> 0 Then
        Mens = Mens & vbCrLf & Err.Description
    End If
End Function


Private Sub VisualizarFrameResultado(b As Boolean)
Dim i As Integer

    FrameResultado.visible = b
    FrameResultado.Enabled = b
    cmdAceptar.visible = Not b
    cmdAceptar.Enabled = Not b
    If b Then
        cmdCancel.Caption = "Salir"
        For i = 0 To txtCodigo.Count - 1
            txtCodigo(i).Enabled = False
        Next i
        imgBuscar(5).Enabled = False
        imgFec(0).Enabled = False
        imgFec(2).Enabled = False
        imgFec(3).Enabled = False
    Else
        cmdCancel.Caption = "Cancelar"
    End If
    
End Sub
