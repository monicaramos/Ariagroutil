VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmConfParamGral 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Datos de Empresa"
   ClientHeight    =   6375
   ClientLeft      =   45
   ClientTop       =   735
   ClientWidth     =   6720
   Icon            =   "frmConfParamGral.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6375
   ScaleWidth      =   6720
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   11
      Left            =   1665
      MaxLength       =   40
      TabIndex        =   11
      Tag             =   "Gerente|T|S|||sempre|perempre|||"
      Top             =   4680
      Width           =   4125
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   4755
      TabIndex        =   29
      Top             =   5880
      Width           =   1035
   End
   Begin VB.Frame Frame1 
      Height          =   540
      Left            =   720
      TabIndex        =   27
      Top             =   5760
      Width           =   2355
      Begin VB.Label lblIndicador 
         Alignment       =   2  'Center
         Caption         =   "Label2"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   28
         Top             =   240
         Width           =   1920
      End
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   10
      Left            =   1665
      MaxLength       =   100
      TabIndex        =   10
      Tag             =   "eMail|T|S|||sempre|maiempre|||"
      Top             =   4245
      Width           =   4125
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   9
      Left            =   1665
      MaxLength       =   100
      TabIndex        =   9
      Tag             =   "Web|T|S|||sempre|wwwempre|||"
      Top             =   3840
      Width           =   4125
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   8
      Left            =   4050
      MaxLength       =   10
      TabIndex        =   8
      Tag             =   "Fax|T|S|||sempre|faxempre|||"
      Top             =   3375
      Width           =   1725
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   7
      Left            =   1680
      MaxLength       =   10
      TabIndex        =   7
      Tag             =   "Tel�fono|T|S|||sempre|telempre|||"
      Top             =   3375
      Width           =   1725
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   6
      Left            =   1665
      MaxLength       =   9
      TabIndex        =   6
      Tag             =   "C.I.F.|T|N|||sempre|nifempre|||"
      Top             =   2940
      Width           =   1605
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   5
      Left            =   1665
      MaxLength       =   35
      TabIndex        =   5
      Tag             =   "Provincia|T|N|||sempre|proempre|||"
      Top             =   2505
      Width           =   4110
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   4
      Left            =   3465
      MaxLength       =   35
      TabIndex        =   4
      Tag             =   "Poblaci�n|T|N|||sempre|pobempre|||"
      Top             =   2070
      Width           =   2325
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   3
      Left            =   1665
      MaxLength       =   6
      TabIndex        =   3
      Tag             =   "CPostal|T|N|||sempre|codposta|||"
      Top             =   2070
      Width           =   765
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   2
      Left            =   1665
      MaxLength       =   40
      TabIndex        =   2
      Tag             =   "Domicilio de la Empresa|T|N|||sempre|domempre|||"
      Top             =   1635
      Width           =   4125
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   3600
      TabIndex        =   12
      Top             =   5880
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   375
      Left            =   4755
      TabIndex        =   13
      Top             =   5880
      Width           =   1035
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   1
      Left            =   1680
      MaxLength       =   40
      TabIndex        =   1
      Tag             =   "Nombre de la Empresa|T|N|||sempre|nomempre|||"
      Top             =   1200
      Width           =   4125
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   0
      Left            =   3360
      MaxLength       =   15
      TabIndex        =   0
      Tag             =   "C�digo Par�metros Generales|N|N|||sempre|codempre||S|"
      Text            =   "1"
      Top             =   1200
      Width           =   645
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   16
      Top             =   0
      Width           =   6720
      _ExtentX        =   11853
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   4
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "A�adir"
            Object.Tag             =   "1"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Modificar"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Salir"
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc Data1 
      Height          =   330
      Left            =   4440
      Top             =   600
      Visible         =   0   'False
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.Label Label1 
      Caption         =   "Gerente"
      Height          =   255
      Index           =   12
      Left            =   780
      TabIndex        =   30
      Top             =   4680
      Width           =   735
   End
   Begin VB.Image ImgMail 
      Height          =   240
      Index           =   0
      Left            =   1395
      Picture         =   "frmConfParamGral.frx":000C
      Tag             =   "-1"
      ToolTipText     =   "Enviar e-mail"
      Top             =   4320
      Width           =   240
   End
   Begin VB.Image imgWeb 
      Height          =   255
      Left            =   1395
      Picture         =   "frmConfParamGral.frx":0596
      Stretch         =   -1  'True
      Tag             =   "-1"
      ToolTipText     =   "Abrir web"
      Top             =   3880
      Width           =   255
   End
   Begin VB.Label Label1 
      Caption         =   "Datos de la Empresa"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Index           =   11
      Left            =   720
      TabIndex        =   26
      Top             =   600
      Width           =   2535
   End
   Begin VB.Label Label1 
      Caption         =   "E-mail"
      Height          =   255
      Index           =   10
      Left            =   780
      TabIndex        =   25
      Top             =   4320
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "Web"
      Height          =   255
      Index           =   9
      Left            =   780
      TabIndex        =   24
      Top             =   3870
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "Fax"
      Height          =   255
      Index           =   8
      Left            =   3700
      TabIndex        =   23
      Top             =   3420
      Width           =   255
   End
   Begin VB.Label Label1 
      Caption         =   "Tel�fono"
      Height          =   255
      Index           =   7
      Left            =   780
      TabIndex        =   22
      Top             =   3420
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "C.I.F."
      Height          =   255
      Index           =   6
      Left            =   780
      TabIndex        =   21
      Top             =   2985
      Width           =   495
   End
   Begin VB.Label Label1 
      Caption         =   "Provincia"
      Height          =   255
      Index           =   5
      Left            =   780
      TabIndex        =   20
      Top             =   2535
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "Poblaci�n"
      Height          =   255
      Index           =   4
      Left            =   2660
      TabIndex        =   19
      Top             =   2070
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "CPostal"
      Height          =   255
      Index           =   3
      Left            =   780
      TabIndex        =   18
      Top             =   2070
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "Domicilio"
      Height          =   255
      Index           =   2
      Left            =   780
      TabIndex        =   17
      Top             =   1650
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "C�digo"
      Height          =   255
      Index           =   0
      Left            =   2460
      TabIndex        =   15
      Top             =   1200
      Width           =   495
   End
   Begin VB.Label Label1 
      Caption         =   "Nombre"
      Height          =   255
      Index           =   1
      Left            =   780
      TabIndex        =   14
      Top             =   1200
      Width           =   615
   End
   Begin VB.Menu mnOpciones 
      Caption         =   "&Opciones"
      Begin VB.Menu mnA�adir 
         Caption         =   "&A�adir"
      End
      Begin VB.Menu mnModificar 
         Caption         =   "&Modificar"
         HelpContextID   =   1
         Shortcut        =   ^M
      End
      Begin VB.Menu mnBarra1 
         Caption         =   "-"
      End
      Begin VB.Menu mnSalir 
         Caption         =   "&Salir"
         Shortcut        =   ^S
      End
   End
End
Attribute VB_Name = "frmConfParamGral"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private NombreTabla As String  'Nombre de la tabla o de la
Private Ordenacion As String
Private CadenaConsulta As String

Private VieneDeBuscar As Boolean
'Para cuando devuelve dos poblaciones con el mismo codigo Postal. Si viene de pulsar prismatico
'de busqueda poner el valor de poblacion seleccionado y no volver a recuperar de la Base de Datos

Dim Modo As Byte
'Solo hay Modo=0 Visualizacion y Modo=4 para Modificar datos
Dim Encontrado As Boolean

Private Sub cmdAceptar_Click()
    If Modo = 3 Then
        If DatosOk Then
            'Cambiamos el path
            'CambiaPath True
            If InsertarDesdeForm(Me) Then
                PonerModo 0
'                ActualizaNombreEmpresa
                MsgBox "Debe salir de la aplicacion para que los cambios tengan efecto", vbExclamation
            End If

        End If
    End If
    
    
    If Modo = 4 Then
        If DatosOk Then
            'Modifica datos en la Tabla: sparam
            If Not ModificaDesdeFormulario(Me) Then Exit Sub
            
            'Actualizar campos de la clase
            vEmpresa.nomEmpre = Text1(1).Text
            vEmpresa.ModificarDatos
    '
    '        vParam.NombreEmpresa = Text1(1).Text
    '        vParam.DomicilioEmpresa = Text1(2).Text
    '        vParam.CPostal = Text1(3).Text
    '        vParam.Poblacion = Text1(4).Text
    '        vParam.Provincia = Text1(5).Text
    '        vParam.CifEmpresa = Text1(6).Text
    '        vParam.Telefono = Text1(7).Text
    '        vParam.Fax = Text1(8).Text
    '        vParam.WebEmpresa = Text1(9).Text
    '        vParam.MailEmpresa = Text1(10).Text
    '        vParam.PerEmpresa = Text1(11).Text
    '        vParam.Modificar
            TerminaBloquear
            
            PonerModo 0
            PonerFocoBtn Me.cmdSalir
        End If
    End If
End Sub

Private Sub cmdCancelar_Click()
    TerminaBloquear
    If Data1.Recordset.EOF Then
        PonerModo 0
    Else
        PonerCampos
        PonerModo 2
    End If
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
    Screen.MousePointer = vbDefault
    If Modo <> 4 Then PonerCadenaBusqueda 'Modo 4: MOdificar
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub Form_Load()
    'Icono del formulario
'    Me.Icon = frmPpal.Icon

    ' ICONITOS DE LA BARRA
    With Me.Toolbar1
        .ImageList = frmPpal.imgListComun
        .Buttons(1).Image = 3   'Anyadir
        .Buttons(2).Image = 4   'Modificar
        .Buttons(4).Image = 11  'Salir
    End With
    
    VieneDeBuscar = False
    
    '## A mano
    NombreTabla = "sempre"
    Ordenacion = " ORDER BY codempre"
    'ASignamos un SQL al DATA1
    Data1.ConnectionString = conn
    CadenaConsulta = "Select * from " & NombreTabla & Ordenacion
    Data1.RecordSource = CadenaConsulta
    Data1.Refresh
    
    Encontrado = True
    If Data1.Recordset.EOF Then
        'No hay registro de datos de parametros
        'quitar###
        Encontrado = False
    End If
    
    PonerModo 0
'    PonerCadenaBusqueda
End Sub

Private Sub PonerCadenaBusqueda()
On Error GoTo EEPonerBusq
    Screen.MousePointer = vbHourglass
    Data1.RecordSource = CadenaConsulta
    Data1.Refresh
    If Data1.Recordset.RecordCount <= 0 Then
        MsgBox "No hay ning�n registro en la tabla " & NombreTabla, vbInformation
        Screen.MousePointer = vbDefault
        Exit Sub
    Else
        Data1.Recordset.MoveFirst
        PonerCampos
        'Bloquea los campos Text1 sino estamos modificando/Insertando Datos
        'Si estamos en Insertar adem�s limpia los campos Text1
        BloquearText1 Me, Modo
    End If

    Screen.MousePointer = vbDefault
    Exit Sub
EEPonerBusq:
    MuestraError Err.Number, "PonerCadenaBusqueda"
    Screen.MousePointer = vbDefault
End Sub

Private Sub PonerCampos()
    If Data1.Recordset.EOF Then Exit Sub
    PonerCamposForma Me, Data1
    If Trim(Text1(3).Text) = "0" Then Text1(3).Text = ""
    If Trim(Text1(6).Text) = "0" Then Text1(6).Text = ""
End Sub

Private Sub imgMail_Click(Index As Integer)
'Abrir Outlook para enviar e-mail
Dim dirMail As String

'    If Modo = 0 Then Exit Sub
    Screen.MousePointer = vbHourglass
    
    Select Case Index
        Case 0: dirMail = Text1(10).Text
    End Select

    If LanzaMailGnral(dirMail) Then espera 2
    Screen.MousePointer = vbDefault
End Sub

Private Sub imgWeb_Click()
'Abrimos el explorador de windows con la pagina Web del cliente

'    If Modo = 0 Then Exit Sub
    Screen.MousePointer = vbHourglass

    If LanzaHomeGnral(Text1(9).Text) Then espera 2
    Screen.MousePointer = vbDefault
End Sub

Private Sub mnModificar_Click()
    If BLOQUEADesdeFormulario(Me) Then BotonModificar
End Sub

Private Sub mnSalir_Click()
    Unload Me
End Sub

Private Sub Text1_GotFocus(Index As Integer)
    ConseguirFoco Text1(Index), 3
End Sub

Private Sub Text1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
'Avanzar/Retroceder los campos con las flechas de desplazamiento del teclado.
    KEYdown KeyCode
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress (KeyAscii)
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
        Case 1 ' a�adir
            mnA�adir_Click
        Case 2  'Modificar
            mnModificar_Click
        Case 4 'Salir
            mnSalir_Click
    End Select
End Sub

Private Sub BotonModificar()
    'A�adiremos el boton de aceptar y demas objetos para insertar
'    Me.lblIndicador.Caption = "MODIFICAR"
    PonerModo 4
    'Me.imgBuscar.Enabled = True
    PonerFoco Text1(1)
End Sub

Private Function DatosOk() As Boolean
Dim b As Boolean
    DatosOk = False
    b = CompForm(Me)
    DatosOk = b
End Function
'
'Private Sub KEYpress(KeyAscii As Integer)
'Dim cerrar As Boolean
'
'    KEYpressGnral KeyAscii, Modo, cerrar
'    If cerrar Then Unload Me
'End Sub

Private Sub PonerBotonCabecera(b As Boolean)
    Me.cmdAceptar.visible = Not b
    Me.cmdCancelar.visible = Not b
    Me.cmdSalir.visible = b
    If b Then Me.lblIndicador.Caption = ""
End Sub

Private Sub PonerModo(vModo As Byte)
Dim b As Boolean

    Modo = vModo
    b = (Modo = 0)
    PonerIndicador Me.lblIndicador, Modo
'    If b Then Me.lblIndicador.Caption = ""
    
' ### [Monica] 13/11/2006
    b = (Modo >= 3) '-->Luego not b sera kmodo<3
    If cmdCancelar.visible Then
        cmdCancelar.Cancel = True
    Else
        cmdCancelar.Cancel = False
    End If
    PonerBotonCabecera Not b

    
    'Bloquea los campos Text1 sino estamos modificando/Insertando Datos
    'Si estamos en Insertar adem�s limpia los campos Text1
    BloquearText1 Me, Modo
    
' ### [Monica] 13/11/2006
    'Poner Botones Aceptar/Cancelar si estamos Modificando datos
'    PonerBotonCabecera b
    
    'Solo si es root o administrador puede modificar el registro
    'cmdAceptar.Enabled = (vUsu.Nivel <= 1)
    
    'Modificar
    Toolbar1.Buttons(1).Enabled = b
    Me.mnModificar.Enabled = b

' ### [Monica] 13/11/2006
    PonerModoOpcionesMenu 'Activar opciones de menu seg�n el Modo


    PonerOpcionesMenu   'Activar opciones de menu seg�n nivel
                        'de permisos del usuario
End Sub

Private Sub PonerOpcionesMenu()
    PonerOpcionesMenuGeneral Me
End Sub



Private Sub KEYpress(KeyAscii As Integer)
Dim cerrar As Boolean

    KEYpressGnral KeyAscii, Modo, cerrar
    If cerrar Then Unload Me

End Sub

' ### [Monica] 13/11/2006
' a�adida la opcion de a�adir cuando no hay registro en la tabla

Private Sub BotonAnyadir()
    'LimpiarCampos
    PonerModo 3
    Text1(0).Text = 1
    PonerFoco Text1(1)
End Sub

Private Sub PonerModoOpcionesMenu()
Dim b As Boolean
    b = (Modo = 3) Or (Modo = 4)
    Me.Toolbar1.Buttons(1).Enabled = Not Encontrado And Not b  'A�adir
    Me.Toolbar1.Buttons(2).Enabled = Encontrado And Not b 'Modificar
    Me.mnA�adir.Enabled = Not Encontrado And Not b
    Me.mnModificar.Enabled = Encontrado And Not b
'    Me.Toolbar1.Buttons(2).Enabled = (Not b) 'Modificar
End Sub

Private Sub mnA�adir_Click()
    BotonAnyadir
End Sub

