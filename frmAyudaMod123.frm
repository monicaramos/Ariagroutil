VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmAyudaMod123 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Ayuda Modelo 123"
   ClientHeight    =   4035
   ClientLeft      =   45
   ClientTop       =   2430
   ClientWidth     =   7185
   Icon            =   "frmAyudaMod123.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4035
   ScaleWidth      =   7185
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
      Height          =   3795
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   6915
      Begin VB.Frame FrameResultados 
         Caption         =   "Resultados"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   1185
         Left            =   480
         TabIndex        =   8
         Top             =   1680
         Width           =   6195
         Begin VB.TextBox txtCodigo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   4
            Left            =   4080
            Locked          =   -1  'True
            MaxLength       =   10
            TabIndex        =   13
            Tag             =   "Código Postal|T|S|||clientes|codposta|||"
            Top             =   600
            Width           =   1500
         End
         Begin VB.TextBox txtCodigo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   2
            Left            =   420
            Locked          =   -1  'True
            MaxLength       =   10
            TabIndex        =   10
            Tag             =   "Código Postal|T|S|||clientes|codposta|||"
            Top             =   600
            Width           =   1200
         End
         Begin VB.TextBox txtCodigo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   3
            Left            =   2040
            Locked          =   -1  'True
            MaxLength       =   10
            TabIndex        =   9
            Tag             =   "Código Postal|T|S|||clientes|codposta|||"
            Top             =   600
            Width           =   1500
         End
         Begin VB.Label Label4 
            Caption         =   "Retenciones"
            ForeColor       =   &H00972E0B&
            Height          =   255
            Index           =   1
            Left            =   4080
            TabIndex        =   14
            Top             =   360
            Width           =   1515
         End
         Begin VB.Label Label4 
            Caption         =   "Nro.Perceptores"
            ForeColor       =   &H00972E0B&
            Height          =   255
            Index           =   5
            Left            =   420
            TabIndex        =   12
            Top             =   360
            Width           =   1215
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            Caption         =   "Base de Retenciones"
            ForeColor       =   &H00972E0B&
            Height          =   255
            Index           =   4
            Left            =   1920
            TabIndex        =   11
            Top             =   360
            Width           =   1635
         End
      End
      Begin VB.CommandButton cmdAceptar 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   4470
         TabIndex        =   2
         Top             =   3180
         Width           =   975
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   0
         Left            =   1920
         MaxLength       =   10
         TabIndex        =   0
         Tag             =   "Código Postal|T|S|||clientes|codposta|||"
         Top             =   780
         Width           =   1050
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   1
         Left            =   1920
         MaxLength       =   10
         TabIndex        =   1
         Tag             =   "Código Postal|T|S|||clientes|codposta|||"
         Top             =   1140
         Width           =   1050
      End
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   480
         Top             =   2730
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Left            =   5655
         TabIndex        =   3
         Top             =   3180
         Width           =   975
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   1
         Left            =   1590
         Picture         =   "frmAyudaMod123.frx":000C
         ToolTipText     =   "Buscar fecha"
         Top             =   1140
         Width           =   240
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   0
         Left            =   1590
         Picture         =   "frmAyudaMod123.frx":0097
         ToolTipText     =   "Buscar fecha"
         Top             =   780
         Width           =   240
      End
      Begin VB.Label Label4 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   14
         Left            =   870
         TabIndex        =   7
         Top             =   1140
         Width           =   420
      End
      Begin VB.Label Label4 
         Caption         =   "Desde"
         Height          =   195
         Index           =   15
         Left            =   870
         TabIndex        =   6
         Top             =   780
         Width           =   465
      End
      Begin VB.Label Label4 
         Caption         =   "Fecha"
         ForeColor       =   &H00972E0B&
         Height          =   255
         Index           =   0
         Left            =   510
         TabIndex        =   5
         Top             =   510
         Width           =   1815
      End
   End
End
Attribute VB_Name = "frmAyudaMod123"
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
Private WithEvents frmC As frmCal 'calendario fechas
Attribute frmC.VB_VarHelpID = -1

Dim indCodigo As Integer 'indice para txtCodigo
Dim indFrame As Single 'nº de frame en el que estamos

'Se inicializan para cada Informe (tabla de BD a la que hace referencia
Dim tabla As String
Dim Codigo As String 'Código para FormulaSelection de Crystal Report
Dim TipCod As String
Dim Orden1 As String 'Campo de Ordenacion (por codigo) para Cristal Report
Dim Orden2 As String 'Campo de Ordenacion (por nombre) para Cristal Report

Dim PrimeraVez As Boolean

Dim Socios As Currency
Dim Total As Currency
Dim Socios1 As Currency
Dim Total1 As Currency
Dim Socios2 As Currency
Dim Total2 As Currency


Private Sub KEYpress(KeyAscii As Integer)
    If KeyAscii = 13 Then 'ENTER
        KeyAscii = 0
        SendKeys "{tab}"
    ElseIf KeyAscii = 27 Then Unload Me  'ESC
    End If
End Sub

Private Sub cmdAceptar_Click()
Dim i As Byte
Dim SQl As String
Dim cadWhere As String
Dim nRegs As Long
Dim c As Object
Dim Rs As ADODB.Recordset
Dim nReceptores As Long
Dim Mens As String

        
    cadWhere = " (1 = 1) "
    If txtCodigo(0).Text <> "" Then cadWhere = cadWhere & " and movim.fechamov >= " & DBSet(txtCodigo(0).Text, "F")
    If txtCodigo(1).Text <> "" Then cadWhere = cadWhere & " and movim.fechamov <= " & DBSet(txtCodigo(1).Text, "F")

    SQl = "SELECT  count(*) from movim where " & cadWhere

    nRegs = TotalRegistros(SQl)

    If nRegs <> 0 Then
        Screen.MousePointer = vbHourglass
        
        Mens = "Cálculo del Total de perceptores: "
        nReceptores = TotalReceptores(cadWhere, Mens)
        
        If nReceptores <> 0 Then
            Set Rs = New ADODB.Recordset
            SQl = "select sum(timport1), sum(timport2) from movim where " & cadWhere
            Rs.Open SQl, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    
            If Not Rs.EOF Then
                txtCodigo(2).Text = Format(nReceptores, "###,###,##0")
                txtCodigo(3).Text = Format(Rs.Fields(0).Value, "###,###,##0.00")
                txtCodigo(4).Text = Format(Rs.Fields(1).Value, "###,###,##0.00")
                
                FrameResultados.visible = True
                imgFec(0).Enabled = False
                imgFec(1).Enabled = False
                cmdAceptar.visible = False
                cmdAceptar.Enabled = False
                cmdCancel.Caption = "Salir"
                
                Screen.MousePointer = vbDefault
            End If
        End If
    Else
        MsgBox "No hay registros para generar el fichero", vbExclamation
    End If

End Sub

Private Sub cmdCancel_Click()
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

    ValoresporDefecto

    'IMAGES para busqueda
'     Me.imgBuscar(0).Picture = frmPpal.imgListImages16.ListImages(1).Picture
'     Me.imgBuscar(1).Picture = frmPpal.imgListImages16.ListImages(1).Picture
'     Me.imgBuscar(6).Picture = frmPpal.imgListImages16.ListImages(1).Picture

    '###Descomentar
'    CommitConexion

    FrameCobrosVisible True, h, w
    indFrame = 5
    tabla = "slhfac"
    FrameResultados.visible = False
    'Esto se consigue poneinedo el cancel en el opcion k corresponda
    Me.cmdCancel.Cancel = True
    Me.Width = w + 70
    Me.Height = h + 350
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

Private Sub Optcodigo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then 'ENTER
    ElseIf KeyAscii = 27 Then Unload Me  'ESC
    End If
End Sub

Private Sub OptNombre_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then 'ENTER
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
            Case 0: KEYFecha KeyAscii, 0 'FECHA desde
            Case 1: KEYFecha KeyAscii, 1 'FECHA hasta
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
        Case 0, 1 'FECHA
            If txtCodigo(Index).Text <> "" Then PonerFormatoFecha txtCodigo(Index)

        Case 2, 3 ' ENTIDADES
            If txtCodigo(Index).Text <> "" Then txtCodigo(Index).Text = Format(txtCodigo(Index).Text, "0000")
    End Select
End Sub

Private Sub FrameCobrosVisible(visible As Boolean, ByRef h As Integer, ByRef w As Integer)
    Me.FrameCobros.visible = visible
    If visible = True Then
        Me.FrameCobros.Top = -90
        Me.FrameCobros.Left = 0
'        Me.FrameCobros.Height = 6015
'        Me.FrameCobros.Width = 6555
        w = Me.FrameCobros.Width
        h = Me.FrameCobros.Height
    End If
End Sub


Private Function TotalReceptores(cadWhere As String, Mens As String) As Long
Dim Rs As ADODB.Recordset
Dim SQl As String
Dim Sql2 As String

    On Error GoTo eTotalReceptores
    
    
    TotalReceptores = 0
    BorrarTMPavnics
    
    If CrearTMPavnics Then
        Set Rs = New ADODB.Recordset
        SQl = "select distinct nifperso, nifrepre from movim, avnic where " & cadWhere
        SQl = SQl & " and movim.codavnic = avnic.codavnic and movim.anoejerc = avnic.anoejerc "
        
        Rs.Open SQl, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
        While Not Rs.EOF
            If DBLet(Rs.Fields(0).Value, "T") <> "" Then
                Sql2 = "insert into tmpavnics (nif) values (" & DBSet(Rs.Fields(0).Value, "T") & ")"
                conn.Execute Sql2
            End If
            If DBLet(Rs.Fields(1).Value, "T") <> "" Then
                Sql2 = "insert into tmpavnics (nif) values (" & DBSet(Rs.Fields(1).Value, "T") & ")"
                conn.Execute Sql2
            End If
            
            Rs.MoveNext
        Wend
        Rs.Close
        SQl = "select count(distinct nif) from tmpavnics"
        Rs.Open SQl, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
        
        If Not Rs.EOF Then TotalReceptores = DBLet(Rs.Fields(0).Value, "N")
        
        Set Rs = Nothing
    End If

eTotalReceptores:
    If Err.Number <> 0 Then
        Mens = "Error en el calculo de Total de Receptores " & Err.Description
    End If
End Function
        
Public Sub BorrarTMPavnics()
On Error Resume Next
    conn.Execute " DROP TABLE IF EXISTS tmpavnics;"
    If Err.Number <> 0 Then Err.Clear
End Sub

Private Function CrearTMPavnics() As Boolean
'Crea una temporal donde insertara los nifs tanto de personas como de representantes
Dim SQl As String
    
    On Error GoTo ECrear
    
    CrearTMPavnics = False
    
    SQl = "CREATE TEMPORARY TABLE tmpavnics ( nif varchar(9) )"
    conn.Execute SQl
     
    CrearTMPavnics = True
    
ECrear:
     If Err.Number <> 0 Then
        CrearTMPavnics = False
        'Borrar la tabla temporal
        SQl = " DROP TABLE IF EXISTS tmpavnics;"
        conn.Execute SQl
    End If
End Function

Private Sub ValoresporDefecto()
Dim MesAnt As Byte
Dim AnoAnt As Integer
Dim fec As Date

    AnoAnt = Year(Now)
    MesAnt = Month(Now) - 1
    If MesAnt = 0 Then
        MesAnt = 12
        AnoAnt = AnoAnt - 1
    End If
    
    txtCodigo(0).Text = "01/" & Format(MesAnt, "00") & "/" & Format(AnoAnt, "0000")
    ' calculo para el mes siguiente
    MesAnt = MesAnt + 1
    If MesAnt = 13 Then
        MesAnt = 1
        AnoAnt = AnoAnt + 1
    End If
    fec = CDate("01/" & Format(MesAnt, "00") & "/" & Format(AnoAnt, "0000"))
    ' al primer dia de este mes le quitamos 1 para que nos de el ultimo dia del mes anterior
    fec = fec - 1
    
    txtCodigo(1).Text = Format(fec, "dd/mm/yyyy")
    
End Sub
