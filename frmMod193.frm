VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMod193 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Grabación de AVNICS Modelo 193"
   ClientHeight    =   5835
   ClientLeft      =   45
   ClientTop       =   2430
   ClientWidth     =   6900
   Icon            =   "frmMod193.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5835
   ScaleWidth      =   6900
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
      Height          =   5625
      Left            =   90
      TabIndex        =   13
      Top             =   90
      Width           =   6675
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   7
         Left            =   4275
         MaxLength       =   10
         TabIndex        =   7
         Tag             =   "Código Postal|T|S|||clientes|codposta|||"
         Top             =   2700
         Width           =   1005
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   3
         Left            =   2475
         MaxLength       =   10
         TabIndex        =   1
         Tag             =   "Código Postal|T|S|||clientes|codposta|||"
         Top             =   765
         Width           =   1005
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   5
         Left            =   1575
         MaxLength       =   6
         TabIndex        =   8
         Top             =   3300
         Width           =   830
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   6
         Left            =   1575
         MaxLength       =   6
         TabIndex        =   9
         Top             =   3675
         Width           =   830
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   5
         Left            =   2460
         Locked          =   -1  'True
         TabIndex        =   24
         Text            =   "Text5"
         Top             =   3315
         Width           =   3135
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   6
         Left            =   2460
         Locked          =   -1  'True
         TabIndex        =   23
         Text            =   "Text5"
         Top             =   3690
         Width           =   3135
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   2
         Left            =   1215
         MaxLength       =   10
         TabIndex        =   10
         Tag             =   "Código Postal|T|S|||clientes|codposta|||"
         Top             =   4140
         Width           =   1185
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   285
         Index           =   4
         Left            =   3645
         Locked          =   -1  'True
         MaxLength       =   3
         TabIndex        =   21
         Tag             =   "Código Postal|T|S|||clientes|codposta|||"
         Top             =   2700
         Width           =   555
      End
      Begin VB.Frame Frame2 
         Caption         =   "Tipo de Declaración"
         ForeColor       =   &H00972E0B&
         Height          =   690
         Left            =   225
         TabIndex        =   19
         Top             =   1890
         Width           =   6045
         Begin VB.OptionButton Option2 
            Caption         =   "Sustitutiva"
            ForeColor       =   &H00000000&
            Height          =   240
            Index           =   2
            Left            =   3915
            TabIndex        =   6
            Top             =   315
            Width           =   1545
         End
         Begin VB.OptionButton Option2 
            Caption         =   "Primera declaración"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   0
            Left            =   225
            TabIndex        =   4
            Top             =   315
            Width           =   1770
         End
         Begin VB.OptionButton Option2 
            Caption         =   "Complementaria"
            ForeColor       =   &H00000000&
            Height          =   240
            Index           =   1
            Left            =   2160
            TabIndex        =   5
            Top             =   315
            Width           =   1545
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Tipo de Soporte"
         ForeColor       =   &H00972E0B&
         Height          =   690
         Left            =   225
         TabIndex        =   18
         Top             =   1125
         Width           =   6045
         Begin VB.OptionButton Option1 
            Caption         =   "Presentación Telemática"
            ForeColor       =   &H00000000&
            Height          =   240
            Index           =   1
            Left            =   2925
            TabIndex        =   3
            Top             =   315
            Width           =   2625
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Presentación en Diskette"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   0
            Left            =   225
            TabIndex        =   2
            Top             =   315
            Width           =   2805
         End
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   285
         Index           =   1
         Left            =   1845
         Locked          =   -1  'True
         MaxLength       =   3
         TabIndex        =   17
         Tag             =   "Código Postal|T|S|||clientes|codposta|||"
         Top             =   765
         Width           =   555
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   0
         Left            =   1845
         MaxLength       =   4
         TabIndex        =   0
         Tag             =   "Código Postal|T|S|||clientes|codposta|||"
         Top             =   360
         Width           =   600
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Left            =   5205
         TabIndex        =   12
         Top             =   4980
         Width           =   975
      End
      Begin VB.CommandButton cmdAceptar 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   4050
         TabIndex        =   11
         Top             =   4980
         Width           =   975
      End
      Begin MSComctlLib.ProgressBar Pb1 
         Height          =   255
         Left            =   360
         TabIndex        =   14
         Top             =   4545
         Width           =   5850
         _ExtentX        =   10319
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   1
      End
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   5985
         Top             =   3420
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.Label lblProgres 
         Height          =   285
         Index           =   1
         Left            =   360
         TabIndex        =   28
         Top             =   4995
         Width           =   3540
      End
      Begin VB.Label Label4 
         Caption         =   "Desde"
         Height          =   195
         Index           =   6
         Left            =   690
         TabIndex        =   27
         Top             =   3300
         Width           =   465
      End
      Begin VB.Label Label4 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   5
         Left            =   690
         TabIndex        =   26
         Top             =   3675
         Width           =   420
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
         Left            =   315
         TabIndex        =   25
         Top             =   3060
         Width           =   1005
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   5
         Left            =   1275
         MouseIcon       =   "frmMod193.frx":000C
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar cliente"
         Top             =   3315
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   6
         Left            =   1260
         MouseIcon       =   "frmMod193.frx":015E
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar cliente"
         Top             =   3690
         Width           =   240
      End
      Begin VB.Label Label4 
         Caption         =   "Teléfono"
         ForeColor       =   &H00972E0B&
         Height          =   255
         Index           =   3
         Left            =   315
         TabIndex        =   22
         Top             =   4140
         Width           =   1815
      End
      Begin VB.Label Label4 
         Caption         =   "Número Justificación Declaración Sustitutiva"
         ForeColor       =   &H00972E0B&
         Height          =   255
         Index           =   2
         Left            =   315
         TabIndex        =   20
         Top             =   2745
         Width           =   3165
      End
      Begin VB.Label Label4 
         Caption         =   "Número Justificante"
         ForeColor       =   &H00972E0B&
         Height          =   255
         Index           =   1
         Left            =   360
         TabIndex        =   16
         Top             =   810
         Width           =   1815
      End
      Begin VB.Label Label4 
         Caption         =   "Ejercicio"
         ForeColor       =   &H00972E0B&
         Height          =   255
         Index           =   0
         Left            =   360
         TabIndex        =   15
         Top             =   405
         Width           =   1815
      End
   End
End
Attribute VB_Name = "frmMod193"
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

'Private WithEvents frmFam As frmManFamia 'Familias
'Private WithEvents frmCol As frmManCoope 'Colectivos
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
Dim tabla As String
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
Dim i As Byte
Dim SQL As String
Dim nRegs As Long
Dim cadWhere As String

    If Not DatosOk Then Exit Sub
    
    SQL = "SELECT  count(*) "
    SQL = SQL & " from avnic where "
    cadWhere = "anoejerc = " & DBSet(txtCodigo(0).Text, "N")
    If txtCodigo(5).Text <> "" Then cadWhere = cadWhere & " and avnic.codavnic >= " & DBSet(txtCodigo(5).Text, "N")
    If txtCodigo(6).Text <> "" Then cadWhere = cadWhere & " and avnic.codavnic <= " & DBSet(txtCodigo(6).Text, "N")
    SQL = SQL & cadWhere
    
    nRegs = TotalRegistros(SQL)

    If nRegs <> 0 Then
        If GeneraFichero(cadWhere) Then
            If CopiarFichero Then
                MsgBox "Proceso realizado correctamente", vbExclamation
                cmdCancel_Click
            End If
        End If
    Else
        MsgBox "No hay datos entre esos límites.", vbExclamation
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

    'IMAGES para busqueda
     Me.imgBuscar(5).Picture = frmPpal.imgListImages16.ListImages(1).Picture
     Me.imgBuscar(6).Picture = frmPpal.imgListImages16.ListImages(1).Picture

    '###Descomentar
'    CommitConexion
         
    FrameCobrosVisible True, h, w
    indFrame = 5
    tabla = "slhfac"
    
    Me.Pb1.visible = False
    InicializarValores
    
    'Esto se consigue poneinedo el cancel en el opcion k corresponda
    Me.cmdCancel.Cancel = True
End Sub


Private Sub frmFam_DatoSeleccionado(CadenaSeleccion As String)
'Form de Consulta de Familias
    txtCodigo(indCodigo).Text = Format(RecuperaValor(CadenaSeleccion, 1), "00")
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmCol_DatoSeleccionado(CadenaSeleccion As String)
'Form de Consulta de Colectivos
    txtCodigo(indCodigo).Text = Format(RecuperaValor(CadenaSeleccion, 1), "00")
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmAvn_DatoSeleccionado(CadenaSeleccion As String)
'Form de Consulta de Clientes
    txtCodigo(indCodigo).Text = Format(RecuperaValor(CadenaSeleccion, 1), "000000")
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub imgBuscar_Click(Index As Integer)
   Select Case Index
        Case 5, 6 'codigo avnics
            AbrirFrmAvnics (Index)
        
    End Select
    PonerFoco txtCodigo(indCodigo)
End Sub

Private Sub Option2_Click(Index As Integer)
    Label4(2).visible = Option2(2).Value
    txtCodigo(4).visible = Option2(2).Value
    txtCodigo(7).visible = Option2(2).Value
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
        Case 0
            If txtCodigo(Index).Text <> "" Then txtCodigo(Index).Text = Format(txtCodigo(Index).Text, "0000")
        Case 2
            If txtCodigo(Index).Text <> "" Then txtCodigo(Index).Text = Format(txtCodigo(Index).Text, "000000000")
        Case 3, 7
            If txtCodigo(Index).Text <> "" Then txtCodigo(Index).Text = Format(txtCodigo(Index).Text, "0000000000")
        Case 1, 4
            If txtCodigo(Index).Text <> "" Then txtCodigo(Index).Text = Format(txtCodigo(Index).Text, "000")
        Case 5, 6 'codigo de avnics
            txtNombre(Index).Text = PonerNombreDeCod(txtCodigo(Index), "avnic", "nombrper", "codavnic", "N")
            If txtCodigo(Index).Text <> "" Then txtCodigo(Index).Text = Format(txtCodigo(Index).Text, "000000")
    End Select
End Sub

Private Sub FrameCobrosVisible(visible As Boolean, ByRef h As Integer, ByRef w As Integer)
    Me.FrameCobros.visible = visible
End Sub

Private Function GeneraFichero(cadWhere As String) As Boolean
Dim NFich1 As Integer
Dim RS As ADODB.Recordset
Dim cad As String
Dim SQL As String
Dim v_Hayreg As Integer
Dim nRegs As Long
Dim b As Boolean
Dim Mens As String

Dim impbase As Currency
Dim ImpReten As Currency
Dim TotalReg As Currency

Dim v_import As String
Dim v_impret As String
Dim t_import As Currency
Dim t_impret As Currency

    On Error GoTo EGen
    
    GeneraFichero = False

    Mens = "Cargando la tabla temporal."
    b = CrearTMPavnicsNew(cadWhere, Mens)
    
    If b Then
        Mens = "Calculando totales."
        b = CalcularTotalesNew(impbase, ImpReten, TotalReg, Mens)
    End If
    
    If b Then
        Me.lblProgres(1).Caption = "Cargando fichero..."
        nRegs = TotalRegistros("select count(*) from tmptempo")
        
        Pb1.visible = True
        Pb1.Max = nRegs
        Pb1.Value = 0
            
        
        NFich1 = FreeFile
        Open App.path & "\fichero.txt" For Output As #NFich1
    
        cad = "1193"
        cad = cad & txtCodigo(0).Text ' anoejerc
        cad = cad & RellenaABlancos(vEmpresa.CifEmpresa, True, 9)
        cad = cad & RellenaABlancos(vEmpresa.nomEmpre, True, 40)
        
        If Option1(0).Value Then 'tipo de documento
            cad = cad & "D"
        Else
            cad = cad & "T"
        End If
        cad = cad & Format(CCur(txtCodigo(2).Text), "000000000") ' telefono
        '[Monica]25/01/2012: cambiado antes el nompresi era de 39, ahora de 40
        cad = cad & RellenaABlancos(vEmpresa.PerEmpresa, True, 40) ' nompresi
        
        cad = cad & Format(txtCodigo(1).Text, "000") 'numero
        
        cad = cad & Format(ComprobarCero(txtCodigo(3).Text), "0000000000") 'justificante
        
        'tipo de declaracion
        If Option2(0).Value Then
            cad = cad & "  " & Repeat("0", 13)
        End If
        If Option2(1).Value Then
            cad = cad & "C " & Repeat("0", 13)
        End If
        If Option2(2).Value Then
            cad = cad & " S" & Format(CCur(txtCodigo(1).Text), "000") & Format(CCur(txtCodigo(7).Text), "0000000000")
        End If
            
        cad = cad & Format(TotalReg, "000000000")
        cad = cad & Format(Round2(impbase * 100, 0), "000000000000000")
        cad = cad & Format(Round2(ImpReten * 100, 0), "000000000000000")
        cad = cad & Format(Round2(ImpReten * 100, 0), "000000000000000")
        
        '[Monica]25/01/2012: ahora son blancos
        cad = cad & Space(30) ' 30 blancos
        cad = cad & Repeat("0", 15) ' gastos
        cad = cad & " " ' naturaleza del declarante
        cad = cad & Space(265)
        
'antes de 25/01/2012
'        cad = cad & Repeat("0", 15)
'        cad = cad & Repeat("0", 15)
'        cad = cad & Repeat("0", 15)
'        cad = cad & Space(1)
'        cad = cad & Space(2)
'        cad = cad & Space(13)
'        cad = cad & "             "
        
        Print #NFich1, cad
    
    
        Set RS = New ADODB.Recordset
        
        'partimos de la tabla de historico de facturas
        SQL = "SELECT * from tmptempo order by nifperso " ' codavnic, tipocodi"
        RS.Open SQL, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
        
        b = True
        v_Hayreg = 0
        While Not RS.EOF And b
            v_Hayreg = 1
            
            Pb1.Value = Pb1.Value + 1
            
            
            cad = "2193"
            cad = cad & txtCodigo(0).Text
            cad = cad & RellenaABlancos(vEmpresa.CifEmpresa, True, 9)
            
            If Trim(DBLet(RS!nifrepre, "T")) <> "" And Not IsNull(RS!nifrepre) Then
                cad = cad & Space(9)
                cad = cad & RellenaABlancos(DBLet(RS!nifrepre, "T"), True, 9)
            Else
                cad = cad & RellenaABlancos(DBLet(RS!nifperso, "T"), True, 9)
                cad = cad & Space(9)
            End If
                
            cad = cad & RellenaABlancos(DBLet(RS!nombrper, "T"), True, 40)
            cad = cad & " "
            cad = cad & Mid(RellenaABlancos(DBLet(RS!codPobla, "T"), True, 6), 1, 2)
            cad = cad & "1"
            cad = cad & RellenaABlancos(vEmpresa.CifEmpresa, True, 12)
            cad = cad & "B"
            cad = cad & "06"
            cad = cad & "1"
            cad = cad & "O"
            cad = cad & Space(20)
            cad = cad & " "
            cad = cad & "0000" '[Monica]25/01/2012: antes esto:"0" & Mid(txtCodigo(0).Text, 2, 3)
            cad = cad & "1"
            cad = cad & Format(Round2(DBLet(RS!ImporPer, "N") * 100, 0), "0000000000000")
            cad = cad & Space(3) '[Monica]25/01/2012: antes esto:"000"
            cad = cad & Repeat("0", 13)
            cad = cad & Format(Round2(DBLet(RS!ImporPer, "N") * 100, 0), "0000000000000")
            cad = cad & Format(Round2(vParamAplic.Porcrete * 100, 0), "0000")
            cad = cad & Format(Round2(DBLet(RS!ImporRet, "N") * 100, 0), "0000000000000")
            cad = cad & Space(13) '[Monica]25/01/2012: antes esto: Repeat("0", 13)
            
            '++monica:24/01/2008 añadido por un tema de Alzira
            cad = cad & Space(14)
            cad = cad & Repeat("0", 40)
            cad = cad & Space(2)
            '++
            
            '[Monica]25/01/2012: ahora la longitud del registro es hasta la 500 (antes 250)
            cad = cad & Space(250)
            'fin
            
            Print #NFich1, cad
                
            RS.MoveNext
        Wend
    End If
    
EGen:
    Close (NFich1)
    Set RS = Nothing
    If Err.Number <> 0 Or Not b Then
        MuestraError Err.Number, Err.Description & vbCrLf & Mens
    Else
        GeneraFichero = True
    End If

End Function


Public Function CopiarFichero() As Boolean
Dim nomFich As String
Dim CADENA As String
On Error GoTo ecopiarfichero

    CopiarFichero = True
    ' abrimos el commondialog para indicar donde guardarlo
'    Me.CommonDialog1.InitDir = App.path

    Me.CommonDialog1.DefaultExt = "txt"
'    CADENA = Format(txtCodigo(2).Text, FormatoFecha)
    CommonDialog1.Filter = "Archivos txt|txt|"
    CommonDialog1.FilterIndex = 1
    CommonDialog1.CancelError = True
    ' copiamos el primer fichero
    CommonDialog1.FileName = "fichero.txt"
    
    Me.CommonDialog1.ShowSave
    If CommonDialog1.FileName <> "" Then
        FileCopy App.path & "\fichero.txt", CommonDialog1.FileName
    End If
    


ecopiarfichero:
    If Err.Number <> 0 And Err.Number <> cdlCancel Then
        MuestraError Err.Number, Err.Description
        CopiarFichero = False
    End If
    Err.Clear

End Function

Private Function DatosOk() As Boolean

    DatosOk = True
    
    If txtCodigo(0).Text = "" Then
        MsgBox "Debe introducir un valor en el Ejercicio de la declaración.", vbExclamation
        PonerFoco txtCodigo(0)
        DatosOk = False
        Exit Function
    End If
    If txtCodigo(2).Text = "" Then
        MsgBox "Debe introducir un valor en el campo Teléfono.", vbExclamation
        PonerFoco txtCodigo(2)
        DatosOk = False
        Exit Function
    End If
    
    
    
End Function

Private Function InsertarEnFichero1(NFich1 As Integer, LetraSer As String, numfactu As Long, vsocio As String, ByRef Mens As String) As Boolean
Dim SQL As String
Dim RS As ADODB.Recordset
Dim vBase As Currency
Dim vIva As Currency
Dim cad As String


    On Error GoTo eInsertarEnFichero1

    InsertarEnFichero1 = False

    SQL = "select schfacr.baseimp1, schfacr.baseimp2, schfacr.baseimp3, schfacr.impoiva1, "
    SQL = SQL & " schfacr.impoiva2, schfacr.impoiva3, schfacr.totalfac, schfacr.fecfactu "
    SQL = SQL & " from schfacr "
    SQL = SQL & " where letraser = " & DBSet(LetraSer, "T") & " and numfactu = " & DBSet(numfactu, "N")
    
    Set RS = New ADODB.Recordset
    RS.Open SQL, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    
    If Not RS.EOF Then
        vBase = DBLet(RS.Fields(0).Value, "N") + DBLet(RS.Fields(1).Value, "N") + DBLet(RS.Fields(2).Value, "N")
        vIva = DBLet(RS.Fields(3).Value, "N") + DBLet(RS.Fields(4).Value, "N") + DBLet(RS.Fields(5).Value, "N")
        
        ' cargamos el fichero fichero1
        cad = LetraSer & "|"
        cad = cad & Format(numfactu, "0000000") & "|"
        cad = cad & Format(DBLet(RS.Fields(7).Value, "F"), FormatoFecha) & "|"
        cad = cad & vsocio & "|"
        cad = cad & Format(vBase, "##,###,##0.00") & "|"
        cad = cad & Format(vIva, "##,###,##0.00") & "|"
        cad = cad & Format(DBLet(RS.Fields(6).Value, "N"), "##,###,##0.00") & "|"
        Print #NFich1, cad
    End If
    InsertarEnFichero1 = True
eInsertarEnFichero1:
    If Err.Number <> 0 Then
        Mens = "Error en la Insercion en el fichero1 " & Err.Description
    End If
End Function

Private Sub InicializarValores()
    txtCodigo(0).Text = Format(Year(Now), "0000")
    Option2(0).Value = True
    txtCodigo(4).visible = False
    txtCodigo(7).visible = False
    Me.Label4(2).visible = False
    txtCodigo(1).Text = 173
    txtCodigo(4).Text = 173
    Option1(1).Value = True
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

Private Function CrearTMPavnics(cadWhere As String, ByRef Mens As String) As Boolean
Dim RS As ADODB.Recordset
Dim SQL As String
Dim ImporPer As Currency
Dim ImporRet As Currency
Dim Existe As String


    On Error GoTo ECrear
    
    CrearTMPavnics = False
        
    Me.lblProgres(1).Caption = "Cargando la tabla temporal..."
        
    'Borrar la tabla temporal
    SQL = " DROP TABLE IF EXISTS tmptempo;"
    conn.Execute SQL
    
    SQL = "CREATE TEMPORARY TABLE tmptempo ( "
'    SQL = "CREATE TABLE tmptempo ( "
    SQL = SQL & "nombrper char(40), "
    SQL = SQL & "nifperso char(9) ,"
    SQL = SQL & "nifrepre char(9) ,"
    SQL = SQL & "codpobla char(6) ,"
    SQL = SQL & "imporper decimal(12,2),"
    SQL = SQL & "imporret decimal(12,2),"
    SQL = SQL & "codavnic int(6),"
    SQL = SQL & "nifpers1 char(9),"
    SQL = SQL & "nombper1 char(40),"
    SQL = SQL & "codpobl1 char(6),"
    SQL = SQL & "tipocodi tinyint(1))"
    
    conn.Execute SQL
    
    SQL = "select nombrper, nifperso, nifrepre, codposta, imporper, imporret, codavnic, "
    SQL = SQL & "nifpers1, nombper1, codposta from avnic where " & cadWhere
    
    
    Set RS = New ADODB.Recordset
    RS.Open SQL, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    While Not RS.EOF
'        If Trim(DBLet(RS!nifrepre, "T")) <> "" Then
        Debug.Print RS!nifrepre & "-" & RS!nifperso & "-" & RS!nifpers1
        If Not IsNull(RS!nifrepre) And Trim(DBLet(RS!nifrepre, "T")) <> "" Then
            '++monica: añadida la condicion de no añadir nifs duplicados
            Existe = ""
            Existe = DevuelveDesdeBDNew(cPTours, "tmptempo", "nifrepre", "nifrepre", RS.Fields(2).Value, "T")
        
            If Existe = "" Then
                SQL = "insert into tmptempo ( nombrper, nifperso, nifrepre, codpobla, imporper,"
                SQL = SQL & " imporret, codavnic, nifpers1, nombper1, codpobl1, tipocodi) values ("
                SQL = SQL & DBSet(RS.Fields(0).Value, "T") & "," 'nombrper
                SQL = SQL & ValorNulo & "," 'nifperso
                SQL = SQL & DBSet(RS.Fields(2).Value, "T") & "," 'nifrepre
                SQL = SQL & DBSet(RS.Fields(3).Value, "T") & "," 'codpobla
                SQL = SQL & DBSet(RS.Fields(4).Value, "N") & "," 'imporper
                SQL = SQL & DBSet(RS.Fields(5).Value, "N") & "," 'imporret
                SQL = SQL & DBSet(RS.Fields(6).Value, "N") & "," 'codavnic
                SQL = SQL & ValorNulo & "," 'nifpers1
                SQL = SQL & ValorNulo & "," 'nombper1
                SQL = SQL & ValorNulo & "," 'codpobl1
                SQL = SQL & "0)"
            Else
                SQL = "update tmptempo set imporper = imporper + " & DBSet(RS.Fields(4).Value, "N")
                SQL = SQL & ", imporret = imporret + " & DBSet(RS.Fields(5).Value, "N")
                SQL = SQL & " where nifrepre = " & DBSet(RS.Fields(2).Value, "T")
            End If
            
            conn.Execute SQL
'       ElseIf Trim(DBLet(RS!nifpers1, "T")) = "" Then
       ElseIf (IsNull(RS!nifpers1) Or Trim(DBLet(RS!nifpers1, "T")) = "") Then
            '++monica: añadida la condicion de no añadir nifs duplicados
            Existe = ""
            Existe = DevuelveDesdeBDNew(cPTours, "tmptempo", "nifperso", "nifperso", RS.Fields(1).Value, "T")
        
            If Existe = "" Then
                SQL = "insert into tmptempo ( nombrper, nifperso, nifrepre, codpobla, imporper,"
                SQL = SQL & " imporret, codavnic, nifpers1, nombper1, codpobl1, tipocodi) values ("
                SQL = SQL & DBSet(RS.Fields(0).Value, "T") & "," 'nombrper
                SQL = SQL & DBSet(RS.Fields(1).Value, "T") & "," 'nifperso
                SQL = SQL & ValorNulo & "," 'nifrepre
                SQL = SQL & DBSet(RS.Fields(3).Value, "T") & "," 'codpobla
                SQL = SQL & DBSet(RS.Fields(4).Value, "N") & "," 'imporper
                SQL = SQL & DBSet(RS.Fields(5).Value, "N") & "," 'imporret
                SQL = SQL & DBSet(RS.Fields(6).Value, "N") & "," 'codavnic
                SQL = SQL & ValorNulo & "," 'nifpers1
                SQL = SQL & ValorNulo & "," 'nombper1
                SQL = SQL & ValorNulo & "," 'codpobl1
                SQL = SQL & "1)"
            Else
                SQL = "update tmptempo set imporper = imporper + " & DBSet(RS.Fields(4).Value, "N")
                SQL = SQL & ", imporret = imporret + " & DBSet(RS.Fields(5).Value, "N")
                SQL = SQL & " where nifperso = " & DBSet(RS.Fields(1).Value, "T")
            End If
            
            conn.Execute SQL
       Else
            ImporPer = Round2(DBLet(RS!ImporPer, "N") * 0.5, 2)
            ImporRet = Round2(DBLet(RS!ImporRet, "N") * 0.5, 2)
            
            '++monica: añadida la condicion de no añadir nifs duplicados
            Existe = ""
            Existe = DevuelveDesdeBDNew(cPTours, "tmptempo", "nifpers1", "nifpers1", RS.Fields(7).Value, "T", , "nifperso", RS.Fields(1).Value, "T")
        
            If Existe = "" Then
                SQL = "insert into tmptempo ( nombrper, nifperso, nifrepre, codpobla, imporper,"
                SQL = SQL & " imporret, codavnic, nifpers1, nombper1, codpobl1, tipocodi) values ("
                SQL = SQL & DBSet(RS.Fields(0).Value, "T") & "," 'nombrper
                SQL = SQL & DBSet(RS.Fields(1).Value, "T") & "," 'nifperso
                SQL = SQL & ValorNulo & "," 'nifrepre
                SQL = SQL & DBSet(RS.Fields(3).Value, "T") & "," 'codpobla
                SQL = SQL & DBSet(ImporPer, "N") & "," 'imporper
                SQL = SQL & DBSet(ImporRet, "N") & "," 'imporret
                SQL = SQL & DBSet(RS.Fields(6).Value, "N") & "," 'codavnic
                SQL = SQL & DBSet(RS.Fields(7).Value, "T") & "," 'nifpers1
                SQL = SQL & DBSet(RS.Fields(8).Value, "T") & "," 'nombper1
                SQL = SQL & DBSet(RS.Fields(9).Value, "T") & "," 'codpobl1
                SQL = SQL & "2)"
            Else
                SQL = "update tmptempo set imporper = imporper + " & DBSet(ImporPer, "N")
                SQL = SQL & ", imporret = imporret + " & DBSet(ImporRet, "N")
                SQL = SQL & " where nifpers1 = " & DBSet(RS.Fields(7).Value, "T")
                SQL = SQL & " and nifperso = " & DBSet(RS.Fields(1).Value, "T")
            End If
            
            conn.Execute SQL
       End If
       RS.MoveNext
    Wend

    CrearTMPavnics = True
    Exit Function
ECrear:
     If Err.Number <> 0 Then
        Mens = Err.Description
        'Borrar la tabla temporal
        SQL = " DROP TABLE IF EXISTS tmptempo;"
        conn.Execute SQL
    End If
End Function


Private Function CrearTMPavnicsNew(cadWhere As String, ByRef Mens As String) As Boolean
Dim RS As ADODB.Recordset
Dim SQL As String
Dim ImporPer As Currency
Dim ImporRet As Currency
Dim Existe As String


    On Error GoTo ECrear
    
    CrearTMPavnicsNew = False
        
    Me.lblProgres(1).Caption = "Cargando la tabla temporal..."
        
    'Borrar la tabla temporal
    SQL = " DROP TABLE IF EXISTS tmptempo;"
    conn.Execute SQL
    
    SQL = "CREATE TEMPORARY TABLE tmptempo ( "
'    SQL = "CREATE TABLE tmptempo ( "
    SQL = SQL & "nombrper char(40), "
    SQL = SQL & "nifperso char(9) ,"
    SQL = SQL & "nifrepre char(9) ,"
    SQL = SQL & "codpobla char(6) ,"
    SQL = SQL & "imporper decimal(12,2),"
    SQL = SQL & "imporret decimal(12,2))"
    
    conn.Execute SQL
    
    SQL = "select nombrper, nifperso, nifrepre, codposta, imporper, imporret, codavnic, "
    SQL = SQL & "nifpers1, nombper1, codposta from avnic where " & cadWhere
    
    
    Set RS = New ADODB.Recordset
    RS.Open SQL, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    While Not RS.EOF
'        If Trim(DBLet(RS!nifrepre, "T")) <> "" Then
        If Not IsNull(RS!nifrepre) And Trim(DBLet(RS!nifrepre, "T")) <> "" Then
            '++monica: añadida la condicion de no añadir nifs duplicados
            Existe = ""
            Existe = DevuelveDesdeBDNew(cPTours, "tmptempo", "nifperso", "nifperso", RS.Fields(1).Value, "T", , "nifrepre", RS.Fields(2).Value, "T")
        
            If Existe = "" Then
                SQL = "insert into tmptempo ( nombrper, nifperso, codpobla, imporper,"
                SQL = SQL & " imporret) values ("
                SQL = SQL & DBSet(RS.Fields(0).Value, "T") & "," 'nombrper
                SQL = SQL & DBSet(RS.Fields(1).Value, "T") & "," 'nifperso
                SQL = SQL & DBSet(RS.Fields(2).Value, "T") & "," 'nifrepre
                SQL = SQL & DBSet(RS.Fields(3).Value, "T") & "," 'codpobla
                SQL = SQL & DBSet(RS.Fields(4).Value, "N") & "," 'imporper
                SQL = SQL & DBSet(RS.Fields(5).Value, "N") & ")" 'imporret
            Else
                SQL = "update tmptempo set imporper = imporper + " & DBSet(RS.Fields(4).Value, "N")
                SQL = SQL & ", imporret = imporret + " & DBSet(RS.Fields(5).Value, "N")
                SQL = SQL & " where nifperso = " & DBSet(RS.Fields(1).Value, "T")
                SQL = SQL & " and nifrepre = " & DBSet(RS.Fields(2).Value, "T")
            End If
            
            conn.Execute SQL
'       ElseIf Trim(DBLet(RS!nifpers1, "T")) = "" Then
       ElseIf (IsNull(RS!nifpers1) Or Trim(DBLet(RS!nifpers1, "T")) = "") Then
            '++monica: añadida la condicion de no añadir nifs duplicados
            Existe = ""
            Existe = DevuelveDesdeBDNew(cPTours, "tmptempo", "nifperso", "nifperso", RS.Fields(1).Value, "T")
        
            If Existe = "" Then
                SQL = "insert into tmptempo ( nombrper, nifperso, nifrepre, codpobla, imporper,"
                SQL = SQL & " imporret) values ("
                SQL = SQL & DBSet(RS.Fields(0).Value, "T") & "," 'nombrper
                SQL = SQL & DBSet(RS.Fields(1).Value, "T") & "," 'nifperso
                SQL = SQL & ValorNulo & "," 'nifrepre
                SQL = SQL & DBSet(RS.Fields(3).Value, "T") & "," 'codpobla
                SQL = SQL & DBSet(RS.Fields(4).Value, "N") & "," 'imporper
                SQL = SQL & DBSet(RS.Fields(5).Value, "N") & ")" 'imporret
            Else
                SQL = "update tmptempo set imporper = imporper + " & DBSet(RS.Fields(4).Value, "N")
                SQL = SQL & ", imporret = imporret + " & DBSet(RS.Fields(5).Value, "N")
                SQL = SQL & " where nifperso = " & DBSet(RS.Fields(1).Value, "T")
            End If
            
            conn.Execute SQL
       Else
            ImporPer = Round2(DBLet(RS!ImporPer, "N") * 0.5, 2)
            ImporRet = Round2(DBLet(RS!ImporRet, "N") * 0.5, 2)
        
        Debug.Print RS!nifrepre & "-" & RS!nifperso & "-" & RS!nifpers1
            
            '++monica: añadida la condicion de no añadir nifs duplicados
            Existe = ""
            Existe = DevuelveDesdeBDNew(cPTours, "tmptempo", "nifperso", "nifperso", RS.Fields(1).Value, "T")
        
            If Existe = "" Then
                SQL = "insert into tmptempo ( nombrper, nifperso, nifrepre, codpobla, imporper,"
                SQL = SQL & " imporret) values ("
                SQL = SQL & DBSet(RS.Fields(0).Value, "T") & "," 'nombrper
                SQL = SQL & DBSet(RS.Fields(1).Value, "T") & "," 'nifperso
                SQL = SQL & ValorNulo & "," 'nifrepre
                SQL = SQL & DBSet(RS.Fields(3).Value, "T") & "," 'codpobla
                SQL = SQL & DBSet(ImporPer, "N") & "," 'imporper
                SQL = SQL & DBSet(ImporRet, "N") & ")" 'imporret
            Else
                SQL = "update tmptempo set imporper = imporper + " & DBSet(ImporPer, "N")
                SQL = SQL & ", imporret = imporret + " & DBSet(ImporRet, "N")
                SQL = SQL & " where nifperso = " & DBSet(RS.Fields(1).Value, "T")
            End If
            
            conn.Execute SQL
            
            Existe = ""
            Existe = DevuelveDesdeBDNew(cPTours, "tmptempo", "nifperso", "nifperso", RS.Fields(7).Value, "T")
       
            If Existe = "" Then
                SQL = "insert into tmptempo ( nombrper, nifperso, nifrepre, codpobla, imporper,"
                SQL = SQL & " imporret) values ("
                SQL = SQL & DBSet(RS.Fields(8).Value, "T") & "," 'nombrper
                SQL = SQL & DBSet(RS.Fields(7).Value, "T") & "," 'nifpers1
                SQL = SQL & ValorNulo & "," 'nifrepre
                SQL = SQL & DBSet(RS.Fields(3).Value, "T") & "," 'codpobla
                SQL = SQL & DBSet(ImporPer, "N") & "," 'imporper
                SQL = SQL & DBSet(ImporRet, "N") & ")" 'imporret
            Else
                SQL = "update tmptempo set imporper = imporper + " & DBSet(ImporPer, "N")
                SQL = SQL & ", imporret = imporret + " & DBSet(ImporRet, "N")
                SQL = SQL & " where nifperso = " & DBSet(RS.Fields(7).Value, "T")
            End If
            
            conn.Execute SQL
       
       
       End If
       RS.MoveNext
    Wend

    CrearTMPavnicsNew = True
    Exit Function
ECrear:
     If Err.Number <> 0 Then
        Mens = Err.Description
        'Borrar la tabla temporal
        SQL = " DROP TABLE IF EXISTS tmptempo;"
        conn.Execute SQL
    End If
End Function



Private Function CalcularTotales(ByRef impbase As Currency, ByRef ImpReten As Currency, ByRef TotalReg As Currency, ByRef Mens As String) As Boolean
Dim SQL As String
Dim RS As ADODB.Recordset
Dim v_import As String
Dim v_impret As String


    On Error GoTo eCalcularTotales

    CalcularTotales = False
    
    impbase = 0
    ImpReten = 0
    TotalReg = 0
    
    Me.lblProgres(1).Caption = "Calculando totales..."
    
    SQL = "select * from tmptempo"
    
    Set RS = New ADODB.Recordset
    RS.Open SQL, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    While Not RS.EOF
        TotalReg = TotalReg + 1
        impbase = impbase + DBLet(RS!ImporPer, "N")
        ImpReten = ImpReten + DBLet(RS!ImporRet, "N")
        
        If DBLet(RS!tipocodi, "N") = 2 Then
            TotalReg = TotalReg + 1
            v_import = ""
            v_impret = "imporret"
            v_import = DevuelveDesdeBDNew(cPTours, "avnic", "imporper", "codavnic", RS!codavnic, "N", v_impret, "anoejerc", txtCodigo(0).Text, "N")
            
            impbase = impbase + CCur(v_import) - DBLet(RS!ImporPer, "N")
            ImpReten = ImpReten + CCur(v_impret) - DBLet(RS!ImporRet, "N")
        End If
        RS.MoveNext
    Wend
    Set RS = Nothing
    CalcularTotales = True
    Exit Function
    
eCalcularTotales:
    If Err.Number <> 0 Then
        Mens = Err.Description
    End If
End Function

Private Function CalcularTotalesNew(ByRef impbase As Currency, ByRef ImpReten As Currency, ByRef TotalReg As Currency, ByRef Mens As String) As Boolean
Dim SQL As String
Dim RS As ADODB.Recordset
Dim v_import As String
Dim v_impret As String


    On Error GoTo eCalcularTotales

    CalcularTotalesNew = False
    
    impbase = 0
    ImpReten = 0
    TotalReg = 0
    
    Me.lblProgres(1).Caption = "Calculando totales..."
    
    SQL = "select * from tmptempo"
    
    Set RS = New ADODB.Recordset
    RS.Open SQL, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    While Not RS.EOF
        TotalReg = TotalReg + 1
        impbase = impbase + DBLet(RS!ImporPer, "N")
        ImpReten = ImpReten + DBLet(RS!ImporRet, "N")
        
        RS.MoveNext
    Wend
    Set RS = Nothing
    CalcularTotalesNew = True
    Exit Function
    
eCalcularTotales:
    If Err.Number <> 0 Then
        Mens = Err.Description
    End If
End Function


