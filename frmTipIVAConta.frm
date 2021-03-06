VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmTipIVAConta 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Tipos de IVA"
   ClientHeight    =   5880
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10470
   Icon            =   "frmTipIVAConta.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5880
   ScaleWidth      =   10470
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtAux 
      Alignment       =   2  'Center
      Height          =   315
      Index           =   8
      Left            =   2835
      MaxLength       =   30
      TabIndex        =   29
      Tag             =   "% Recargo de IVA|N|N|0|100|tiposiva|porcerec|#0.00||"
      Text            =   "%Rec"
      Top             =   4905
      Width           =   645
   End
   Begin VB.Frame FrameSopor 
      Caption         =   " Soportado"
      ForeColor       =   &H00972E0B&
      Height          =   2445
      Left            =   5640
      TabIndex        =   20
      Top             =   2640
      Width           =   4575
      Begin VB.TextBox txtAux 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   7
         Left            =   240
         MaxLength       =   10
         TabIndex        =   11
         Tag             =   "cuenta soportado N/Ded|T|N|||tiposiva|cuentasn|||"
         Text            =   "Text1"
         Top             =   1920
         Width           =   1080
      End
      Begin VB.TextBox txtAux2 
         BackColor       =   &H80000018&
         Height          =   315
         Index           =   7
         Left            =   1440
         MaxLength       =   30
         TabIndex        =   25
         Text            =   "Text2"
         Top             =   1920
         Width           =   2925
      End
      Begin VB.TextBox txtAux 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   5
         Left            =   240
         MaxLength       =   10
         TabIndex        =   9
         Tag             =   "cuenta soportado|T|N|||tiposiva|cuentaso|||"
         Text            =   "Text1"
         Top             =   480
         Width           =   1080
      End
      Begin VB.TextBox txtAux2 
         BackColor       =   &H80000018&
         Height          =   315
         Index           =   5
         Left            =   1440
         MaxLength       =   30
         TabIndex        =   22
         Text            =   "Text2"
         Top             =   480
         Width           =   2925
      End
      Begin VB.TextBox txtAux 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   6
         Left            =   240
         MaxLength       =   10
         TabIndex        =   10
         Tag             =   "cuenta soportado recargo|T|N|||tiposiva|cuentasr|||"
         Text            =   "Text1"
         Top             =   1200
         Width           =   1080
      End
      Begin VB.TextBox txtAux2 
         BackColor       =   &H80000018&
         Height          =   315
         Index           =   6
         Left            =   1440
         MaxLength       =   30
         TabIndex        =   21
         Text            =   "Text2"
         Top             =   1200
         Width           =   2925
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   4
         Left            =   2160
         MousePointer    =   4  'Icon
         Tag             =   "-1"
         ToolTipText     =   "Buscar cuenta contable"
         Top             =   1680
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Cuenta soportado N/Ded"
         Height          =   315
         Index           =   3
         Left            =   240
         TabIndex        =   26
         Top             =   1695
         Width           =   1815
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   2
         Left            =   1590
         MousePointer    =   4  'Icon
         Tag             =   "-1"
         ToolTipText     =   "Buscar cuenta contable"
         Top             =   240
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Cuenta soportado"
         Height          =   315
         Index           =   2
         Left            =   240
         TabIndex        =   24
         Top             =   255
         Width           =   1455
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   3
         Left            =   2190
         MousePointer    =   4  'Icon
         Tag             =   "-1"
         ToolTipText     =   "Buscar cuenta contable"
         Top             =   960
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Cuenta soportado recargo"
         Height          =   315
         Index           =   1
         Left            =   240
         TabIndex        =   23
         Top             =   975
         Width           =   2055
      End
   End
   Begin VB.Frame FrameReper 
      Caption         =   " Repercutido"
      ForeColor       =   &H00972E0B&
      Height          =   1725
      Left            =   5640
      TabIndex        =   15
      Top             =   720
      Width           =   4575
      Begin VB.TextBox txtAux2 
         BackColor       =   &H80000018&
         Height          =   315
         Index           =   4
         Left            =   1440
         MaxLength       =   30
         TabIndex        =   18
         Text            =   "Text2"
         Top             =   1200
         Width           =   2925
      End
      Begin VB.TextBox txtAux 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   4
         Left            =   240
         MaxLength       =   10
         TabIndex        =   8
         Tag             =   "cuenta repercutido recargo|T|N|||tiposiva|cuentarr|||"
         Text            =   "Text1"
         Top             =   1200
         Width           =   1080
      End
      Begin VB.TextBox txtAux2 
         BackColor       =   &H80000018&
         Height          =   315
         Index           =   3
         Left            =   1440
         MaxLength       =   30
         TabIndex        =   16
         Text            =   "Text2"
         Top             =   480
         Width           =   2925
      End
      Begin VB.TextBox txtAux 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   3
         Left            =   240
         MaxLength       =   10
         TabIndex        =   7
         Tag             =   "cuenta repercutido|T|N|||tiposiva|cuentare|||"
         Text            =   "Text1"
         Top             =   480
         Width           =   1080
      End
      Begin VB.Label Label1 
         Caption         =   "Cuenta repercutido recargo"
         Height          =   315
         Index           =   0
         Left            =   240
         TabIndex        =   19
         Top             =   975
         Width           =   2055
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   1
         Left            =   2310
         MousePointer    =   4  'Icon
         Tag             =   "-1"
         ToolTipText     =   "Buscar cuenta contable"
         Top             =   960
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Cuenta repercutido"
         Height          =   315
         Index           =   4
         Left            =   240
         TabIndex        =   17
         Top             =   255
         Width           =   1455
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   0
         Left            =   1710
         MousePointer    =   4  'Icon
         Tag             =   "-1"
         ToolTipText     =   "Buscar cuenta contable"
         Top             =   240
         Width           =   240
      End
   End
   Begin VB.TextBox txtAux 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   290
      Index           =   2
      Left            =   2280
      MaxLength       =   5
      TabIndex        =   5
      Tag             =   "% IVA|N|N|0|99,99|tiposiva|porceiva|#0.00|N|"
      Text            =   "%"
      Top             =   4920
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.TextBox txtAux 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   290
      Index           =   1
      Left            =   720
      MaxLength       =   15
      TabIndex        =   4
      Tag             =   "Desc. Tipo IVA|T|N|||tiposiva|nombriva|||"
      Text            =   "nom caja"
      Top             =   4920
      Visible         =   0   'False
      Width           =   1515
   End
   Begin VB.ComboBox Combo1 
      Appearance      =   0  'Flat
      Height          =   315
      Index           =   0
      ItemData        =   "frmTipIVAConta.frx":000C
      Left            =   3555
      List            =   "frmTipIVAConta.frx":000E
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Tag             =   "Tipo de IVA|N|N|0|3|tiposiva|tipodiva||N|"
      Top             =   4920
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   7260
      TabIndex        =   1
      Top             =   5340
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   8520
      TabIndex        =   2
      Top             =   5340
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox txtAux 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   290
      Index           =   0
      Left            =   240
      MaxLength       =   2
      TabIndex        =   3
      Tag             =   "C�digo IVA|N|N|0|99|tiposiva|codigiva|00|S|"
      Text            =   "cod"
      Top             =   4920
      Visible         =   0   'False
      Width           =   435
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "frmTipIVAConta.frx":0010
      Height          =   4410
      Left            =   120
      TabIndex        =   0
      Top             =   540
      Width           =   5115
      _ExtentX        =   9022
      _ExtentY        =   7779
      _Version        =   393216
      AllowUpdate     =   0   'False
      BorderStyle     =   0
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         AllowFocus      =   0   'False
         AllowRowSizing  =   0   'False
         AllowSizing     =   0   'False
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdRegresar 
      Caption         =   "&Regresar"
      Height          =   375
      Left            =   8520
      TabIndex        =   14
      Top             =   5340
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Height          =   555
      Index           =   1
      Left            =   120
      TabIndex        =   12
      Top             =   5175
      Width           =   2625
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
         Left            =   45
         TabIndex        =   13
         Top             =   240
         Width           =   2415
      End
   End
   Begin MSAdodcLib.Adodc adodc1 
      Height          =   330
      Left            =   3480
      Top             =   120
      Visible         =   0   'False
      Width           =   2055
      _ExtentX        =   3625
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
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   27
      Top             =   0
      Width           =   10470
      _ExtentX        =   18468
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      Style           =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   12
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Style           =   3
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Buscar"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Ver todos"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Nuevo"
            Object.Tag             =   "2"
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Modificar"
            Object.Tag             =   "2"
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Eliminar"
            Object.Tag             =   "2"
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Style           =   3
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Imprimir"
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Salir"
         EndProperty
      EndProperty
      BorderStyle     =   1
      Begin VB.CheckBox chkVistaPrevia 
         Caption         =   "Vista previa"
         Height          =   195
         Left            =   8400
         TabIndex        =   28
         Top             =   120
         Visible         =   0   'False
         Width           =   1215
      End
   End
   Begin VB.Menu mnOpciones 
      Caption         =   "&Opciones"
      Begin VB.Menu mnBuscar 
         Caption         =   "&Buscar"
         Shortcut        =   ^F
      End
      Begin VB.Menu mnVerTodos 
         Caption         =   "&Ver todos"
         Shortcut        =   ^B
      End
      Begin VB.Menu mnBarra1 
         Caption         =   "-"
      End
      Begin VB.Menu mnNuevo 
         Caption         =   "&Nuevo"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnModificar 
         Caption         =   "&Modificar"
         Shortcut        =   ^M
      End
      Begin VB.Menu mnEliminar 
         Caption         =   "&Eliminar"
         Shortcut        =   ^E
      End
      Begin VB.Menu mnBarra2 
         Caption         =   "-"
      End
      Begin VB.Menu mnSalir 
         Caption         =   "&Salir"
         Shortcut        =   ^S
      End
   End
End
Attribute VB_Name = "frmTipIVAConta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' +-+-+-+-+-+-+-+-+-+-+-
' +-+- Autor: LAURA +-+-
' +-+-+-+-+-+-+-+-+-+-+-

' **************** PER A QUE FUNCIONE EN UN ATRE MANTENIMENT ********************
' 0. Posar-li l'atribut Datasource a "adodc1" del Datagrid1. Canviar el Caption
'    del formulari
' 1. Canviar els TAGs i els Maxlength de TextAux(0) i TextAux(1)
' 2. En PonerModo(vModo) repasar els indexs del botons, per si es canvien
' 3. En la funci� BotonAnyadir() canviar la taula i el camp per a SugerirCodigoSiguienteStr
' 4. En la funci� BotonBuscar() canviar el nom de la clau primaria
' 5. En la funci� BotonEliminar() canviar la pregunta, les descripcions de la
'    variable SQL i el contingut del DELETE
' 6. En la funci� posamaxlength() repasar el maxlength de TextAux(0)
' 7. En Form_Load() repasar la barra d'iconos (per si es vol canviar alg�n) i
'    canviar la consulta per a vore tots els registres
' 8. En Toolbar1_ButtonClick repasar els indexs de cada bot� per a que corresponguen
' 9. En la funci� CargaGrid canviar l'ORDER BY (normalment per la clau primaria);
'    canviar adem�s els noms dels camps, el format i si fa falta la cantitat;
'    repasar els index dels botons modificar i eliminar.
'    NOTA: si en Form_load ya li he posat clausula WHERE, canviar el `WHERE` de
'    `SQL = CadenaConsulta & " WHERE " & vSQL` per un `AND`
' 10. En txtAux_LostFocus canviar el mensage i el format del camp
' 11. En la funci� DatosOk() canviar els arguments de DevuelveDesdeBD i el mensage
'    en cas d'error
' 12. En la funci� SepuedeBorrar() canviar les comprovacions per a vore si es pot
'    borrar el registre
' *******************************SI N'HI HA COMBO*******************************
' 0. Comprovar que en el SQL de Form_Load() es fa�a refer�ncia a la taula del Combo
' 1. Pegar el Combo1 al  costat dels TextAux. Canviar-li el TAG
' 2. En BotonModificar() canviar el camp del Combo
' 3. En CargaCombo() canviar la consulta i els noms del camps, o posar els valor
'    a ma si no es llig de cap base de datos els valors del Combo

Option Explicit

Public DatosADevolverBusqueda As String    'Tendra el n� de text que quiere que devuelva, empipados
Public Event DatoSeleccionado(CadenaSeleccion As String)


'codigo que tiene el campo en el momento que se llama desde otro formulario
'nos situamos en ese valor
Public CodigoActual As String
'indica si venimos del punto de facturas varias
Public Facturas As Boolean '0 no viene de facturas varias / 1 viene de facturas varias
Public Conexion As Integer '2 contabilidad de avnics
                           '3 contabilidad de seguros
                           '4 contabilidad de telefonia
                           '5 contabilidad de gasolinera
                           '6 contabilidad de facturas socios
                           
'Public DeConsulta As Boolean



Private WithEvents frmCtas As frmCtasConta 'Cuentas contables de la Contabilidad
Attribute frmCtas.VB_VarHelpID = -1

Private CadenaConsulta As String
Private CadB As String

Dim PrimeraVez As Boolean
Dim Modo As Byte
'----------------------------------------------
'*************** MODOS ********************
'   0.-  Formulari net sense cap camp ple
'   1.-  Preparant per a fer la b�squeda
'   2.-  Ja tenim registres i els anem a recorrer
'        ,podem editar-los Edici� del camp
'   3.-  Inserci� de nou registre
'   4.-  Modificar
'   5.-  Manteniment Llinies
'----------------------------------------------


Private Sub PonerModo(vModo)
Dim b As Boolean
Dim i As Byte

    Modo = vModo
    
    b = (Modo = 2)
    If b Then
        PonerContRegIndicador
    Else
        PonerIndicador lblIndicador, Modo
    End If
    
    For i = 0 To 2
        txtAux(i).visible = Not b
    Next i
    Combo1(0).visible = Not b
    txtAux(8).visible = Not b

    cmdAceptar.visible = Not b
    cmdCancelar.visible = Not b
'    DataGrid1.Enabled = b
    
    'Si es regresar
    If DatosADevolverBusqueda <> "" Then cmdRegresar.visible = b
        
    'bloquear los campos que no estan en el Grid si no estamos insertando/modificando
'    b = (Modo = 2)
    For i = 3 To 7
        BloquearTxt txtAux(i), b
    Next i
        
    'Bloquear los botones para busquedas
    BloquearImgBuscar Me, Modo
        
    'Si estamos modo Modificar bloquear clave primaria
    BloquearTxt txtAux(0), (Modo = 4)
    
    'Poner el tama�o de los campos. Si es modo Busqueda el MaxLength del campo
    'debe ser mayor.
    PonerLongCampos
    PonerModoOpcionesMenu 'Activar/Desact botones de menu segun Modo
    PonerOpcionesMenu  'En funcion del usuario
End Sub


Private Sub PonerLongCampos()
'Modificar el MaxLength del campo en funcion de si es modo de b�squeda o no
'para los campos que permitan introducir criterios m�s largos del tama�o del campo
'Si maxlength=6 para codprove, en modo busqueda Maxlenth=13 'el doble + 1
    
    'para los txtAux
    PonerLongCamposGnral Me, Modo, 3
End Sub


Private Sub PonerModoOpcionesMenu()
'Activa/Desactiva botones del la toobar y del menu, segun el modo en que estemos
Dim b As Boolean

    b = (Modo = 2)
    'busqueda
    Toolbar1.Buttons(2).Enabled = b
    Me.mnBuscar.Enabled = b
    'ver todos
    Toolbar1.Buttons(3).Enabled = b
    Me.mnVerTodos.Enabled = b
    
    'Insertar
    Toolbar1.Buttons(6).Enabled = False
    Me.mnNuevo.Enabled = False
    'Modificar
    Toolbar1.Buttons(7).Enabled = False
    Me.mnModificar.Enabled = False
    'Eliminar
    Toolbar1.Buttons(8).Enabled = False
    Me.mnEliminar.Enabled = False
    
    b = (b And adodc1.Recordset.RecordCount > 0)
    
    'Imprimir
    Toolbar1.Buttons(11).Enabled = b
    
End Sub


Private Sub BotonAnyadir()
'Dim NumF As String
'Dim anc As Single
'Dim i As Byte
End Sub


Private Sub BotonVerTodos()
    CargaGrid ""
    CadB = ""
    PonerModo 2
End Sub

Private Sub BotonBuscar()
Dim i As Byte

    ' ***************** canviar per la clau primaria ********
    CargaGrid "codigiva = -1"
    '*******************************************************************************
    'Buscar
    For i = 0 To txtAux.Count - 1
        txtAux(i).Text = ""
    Next i
    Combo1(0).ListIndex = -1

    LLamaLineas DataGrid1.Top + 206, 1 'Pone el Modo=1, Busqueda
    PonerFoco txtAux(0)
End Sub

Private Sub BotonModificar()
'    Dim cad As String
'    Dim anc As Single
'    Dim i As Integer, j As Integer
End Sub


Private Sub LLamaLineas(alto As Single, xModo As Byte)
    DeseleccionaGrid Me.DataGrid1
    PonerModo xModo
    'Fijamos el ancho
    txtAux(0).Top = alto
    txtAux(1).Top = alto
    txtAux(2).Top = alto
    txtAux(8).Top = alto
    Combo1(0).Top = alto - 15
End Sub

'Private Sub BotonEliminar()
''Dim SQL As String
'End Sub


Private Sub cmdAceptar_Click()
    On Error GoTo EAceptar

    Select Case Modo
        Case 1 'BUSQUEDA
            CadB = ObtenerBusqueda(Me)
            If CadB <> "" Then
                CargaGrid CadB
                PonerModo 2
'                lblIndicador.Caption = "BUSQUEDA: " & PonerContRegistros(Me.adodc1)
                PonerFocoGrid Me.DataGrid1
            End If
    End Select
    
EAceptar:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description
End Sub

Private Sub cmdCancelar_Click()

    Select Case Modo
        Case 1 'b�squeda
            CargaGrid CadB
    End Select
    
    PonerModo 2
    
    If Not adodc1.Recordset.EOF Then DataGrid1_RowColChange 1, 1
    
'    If CadB <> "" Then
'        lblIndicador.Caption = "BUSQUEDA: " & PonerContRegistros(Me.adodc1)
'    Else
'        lblIndicador.Caption = ""
'    End If
    PonerFocoGrid Me.DataGrid1
End Sub


Private Sub cmdRegresar_Click()
Dim cad As String
Dim i As Integer
Dim J As Integer
Dim Aux As String

    If adodc1.Recordset.EOF Then
        MsgBox "Ning�n registro devuelto.", vbExclamation
        Exit Sub
    End If
    cad = ""
    i = 0
    Do
        J = i + 1
        i = InStr(J, DatosADevolverBusqueda, "|")
        If i > 0 Then
            Aux = Mid(DatosADevolverBusqueda, J, i - J)
            J = Val(Aux)
            cad = cad & adodc1.Recordset.Fields(J) & "|"
        End If
    Loop Until i = 0
    RaiseEvent DatoSeleccionado(cad)
    Unload Me
End Sub


Private Sub Combo1_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub DataGrid1_DblClick()
    If cmdRegresar.visible Then cmdRegresar_Click
End Sub

Private Sub DataGrid1_KeyPress(KeyAscii As Integer)
'    If KeyAscii = 13 And Me.cmdRegresar.visible Then
'        DeseleccionaGrid Me.DataGrid1
'        PonerFocoBtn Me.cmdRegresar
'    Else
        KEYpress KeyAscii
'    End If
End Sub

Private Sub DataGrid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)

    If DataGrid1.Columns.Count > 2 And (Modo = 0 Or Modo = 2 Or Modo = 4) Then
        If IsNull(DataGrid1.Bookmark) Then Exit Sub
        '----------cuentas contables
        'cuenta repercutido
        txtAux(3).Text = DataGrid1.Columns(5).Text
        txtAux2(3).Text = PonerNombreCuenta(txtAux(3), Modo, , Conexion, Facturas)
        'cuenta repercutido recargo
        txtAux(4).Text = DataGrid1.Columns(6).Text
        txtAux2(4).Text = PonerNombreCuenta(txtAux(4), Modo, , Conexion, Facturas)
        
        'cuenta soportado
        txtAux(5).Text = DataGrid1.Columns(7).Text
        txtAux2(5).Text = PonerNombreCuenta(txtAux(5), Modo, , Conexion, Facturas)
        'cuenta soportado recargo
        txtAux(6).Text = DataGrid1.Columns(8).Text
        txtAux2(6).Text = PonerNombreCuenta(txtAux(6), Modo, , Conexion, Facturas)
        'cuenta soportado N/Ded
        txtAux(7).Text = DataGrid1.Columns(9).Text
        txtAux2(7).Text = PonerNombreCuenta(txtAux(7), Modo, , Conexion, Facturas)
    End If
    
'
'    If (Modo = 2 Or Modo = 0) Then
'        If CadB = "" Then
'            lblIndicador.Caption = PonerContRegistros(Me.adodc1)
'        Else
'            lblIndicador.Caption = "BUSQUEDA: " & PonerContRegistros(Me.adodc1)
'        End If
'    End If
    PonerContRegIndicador
End Sub

Private Sub LimpiarCampos()
    On Error Resume Next
    
    Limpiar Me   'Metodo general: Limpia los controles TextBox
    lblIndicador.Caption = ""
    If Err.Number <> 0 Then Err.Clear
End Sub



Private Sub Form_Activate()
'    On Error Resume Next
    
    Screen.MousePointer = vbDefault
'    'Ponemos el foco
'    If Modo = 0 Or Modo = 2 Then PonerFocoGrid Me.DataGrid1
'
'    If Err.Number <> 0 Then Err.Clear

    If PrimeraVez Then
        PrimeraVez = False
'        If (DatosADevolverBusqueda <> "") And NuevoCodigo <> "" Then
'            BotonAnyadir
'        Else
            PonerModo 2
            If Me.CodigoActual <> "" Then
                SituarData Me.adodc1, "codigiva=" & CodigoActual, "", True
                PonerFocoGrid Me.DataGrid1
            End If
'        End If
    End If
End Sub


Private Sub Form_Load()
Dim i As Integer
    
    PrimeraVez = True
    
    With Me.Toolbar1
        .HotImageList = frmPpal.imgListComun_OM
        .DisabledImageList = frmPpal.imgListComun_BN
        .ImageList = frmPpal.imgListComun
        'el 1 es separadors
        .Buttons(2).Image = 1   'Buscar
        .Buttons(3).Image = 2   'Todos
        'el 4 i el 5 son separadors
        .Buttons(6).Image = 3   'Insertar
        .Buttons(7).Image = 4   'Modificar
        .Buttons(8).Image = 5   'Borrar
        'el 9 i el 10 son separadors
        .Buttons(11).Image = 10  'Imprimir
        .Buttons(12).Image = 11  'Salir
    End With


    'IMAGES para busqueda
    For i = 0 To Me.imgBuscar.Count - 1
        Me.imgBuscar(i).Picture = frmPpal.imgListImages16.ListImages(1).Picture
    Next i


    '## A mano
'    chkVistaPrevia.Value = CheckValueLeer(Name)
    LimpiarCampos
    CargaCombo (0)
    
    '****************** canviar la consulta *********************************+
    CadenaConsulta = "SELECT codigiva,nombriva,porceiva,porcerec,CASE tipodiva WHEN 0 THEN ""IVA"" WHEN 1 THEN ""IGIC"" WHEN 2 THEN ""BIEN DE INVERSI�N"" WHEN 3 THEN ""R.E.A."" END,cuentare,cuentarr,cuentaso,cuentasr,cuentasn FROM tiposiva "
    '************************************************************************
    CadB = ""
    CargaGrid
    
    If (DatosADevolverBusqueda <> "") Then 'And NuevoCodigo <> "" Then
        BotonAnyadir
    Else
        PonerModo 2
        
    End If

End Sub

Private Sub Form_Unload(Cancel As Integer)
'    CheckValueGuardar Me.Name, Me.chkVistaPrevia.Value
    Screen.MousePointer = vbDefault
End Sub

Private Sub frmCtas_DatoSeleccionado(CadenaSeleccion As String)
'Cuentas contables de la Contabilidad
    txtAux(CStr(imgBuscar(1).Tag)).Text = RecuperaValor(CadenaSeleccion, 1)  'codmacta
    txtAux2(CStr(imgBuscar(1).Tag)).Text = RecuperaValor(CadenaSeleccion, 2) 'des macta
End Sub

Private Sub imgBuscar_Click(Index As Integer)
    TerminaBloquear
    Select Case Index
        Case 1 To 5 'cta contable
            Me.imgBuscar(1).Tag = Index + 3
            Set frmCtas = New frmCtasConta
            frmCtas.NumDigit = 0
            frmCtas.DatosADevolverBusqueda = "0|1|"
            frmCtas.Show vbModal
            Set frmCtas = Nothing
            PonerFoco txtAux(Index)
    End Select
    If Modo = 4 Then BLOQUEADesdeFormulario2 Me, adodc1, 1
End Sub

Private Sub mnBuscar_Click()
    BotonBuscar
End Sub

Private Sub mnEliminar_Click()
'    BotonEliminar
End Sub

Private Sub mnModificar_Click()
'    If BLOQUEADesdeFormulario2(Me, adodc1, 1) Then BotonModificar
End Sub

Private Sub mnNuevo_Click()
'    BotonAnyadir
End Sub

Private Sub mnSalir_Click()
    Unload Me
End Sub

Private Sub mnVerTodos_Click()
    BotonVerTodos
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
        Case 2
                BotonBuscar
        Case 3
                BotonVerTodos
        Case 6
'                BotonAnyadir
        Case 7
'                mnModificar_Click
        Case 8
'                BotonEliminar
                
        Case 11 'Imprimir
                BotonImprimir
                'printNou
        Case 12 'Salir
                mnSalir_Click
    End Select
End Sub


Private Sub CargaGrid(Optional vSQL As String)
    Dim i As Integer
    Dim SQL As String
    Dim tots As String
    
    On Error GoTo ECargaGRid
    
    If Not Facturas Then
        Select Case Conexion
            Case cConta
                adodc1.ConnectionString = ConnConta
            Case cContaSeg
                adodc1.ConnectionString = ConnContaSeg
            Case cContaTel
                adodc1.ConnectionString = ConnContaTel
            Case cContaGas
                adodc1.ConnectionString = ConnContaGas
            Case cContaFacSoc
                adodc1.ConnectionString = ConnContaFacSoc
            Case cContaCV
                adodc1.ConnectionString = ConnContaCV
            Case cContaCVV
                adodc1.ConnectionString = ConnContaCVV
        End Select
    Else
        adodc1.ConnectionString = ConnContaFac 'BD de la Contabilidad facturas varias
    End If
    
    SQL = CadenaConsulta '& " WHERE " & WhereSel
    If vSQL <> "" Then SQL = SQL & " WHERE " & vSQL

    '********************* canviar el ORDER BY *********************++
    SQL = SQL & " ORDER BY codigiva"
    '**************************************************************++
    
    adodc1.RecordSource = SQL
    adodc1.CursorType = adOpenDynamic
    adodc1.LockType = adLockOptimistic
    DataGrid1.ScrollBars = dbgNone
    adodc1.Refresh
    Set DataGrid1.DataSource = adodc1
    
    Set DataGrid1.DataSource = adodc1
    DataGrid1.AllowRowSizing = False
    DataGrid1.RowHeight = 290
    
    'si es visible|control|tipo campo|nombre campo|ancho control|formato campo|
    tots = "S|txtAux(0)|T|Cod.|600|;S|txtAux(1)|T|Descripci�n|1800|;" 'codigiva, nombriva
    tots = tots & "S|txtAux(2)|T|%IVA|600|;S|txtAux(8)|T|%Rec|600|;S|Combo1(0)|C|Tipo|1100|;"
    tots = tots & "N||||0|;N||||0|;N||||0|;N||||0|;N||||0|;" 'cuentas conta
    arregla tots, DataGrid1, Me
    Me.DataGrid1.Columns(2).Alignment = dbgCenter
    DataGrid1.ScrollBars = dbgAutomatic
    
    If Not adodc1.Recordset.EOF Then
        DataGrid1_RowColChange 1, 1
    Else
        For i = 3 To 7 'cuentas
            txtAux(i).Text = ""
            txtAux2(i).Text = ""
        Next i
    End If
    
    DataGrid1.Columns(0).Alignment = dbgRight
    DataGrid1.Columns(2).Alignment = dbgRight
 
ECargaGRid:
    If Err.Number <> 0 Then MuestraError Err.Number, "Cargando Grid", Err.Description  'MsgBox Err.Number & ": " & Err.Description
End Sub


Private Sub txtAux_GotFocus(Index As Integer)
    ConseguirFocoLin txtAux(Index)
End Sub


Private Sub txtAux_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    KEYdown KeyCode
End Sub

Private Sub txtAux_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub


Private Sub txtAux_LostFocus(Index As Integer)
    
    If Not PerderFocoGnral(txtAux(Index), Modo) Then Exit Sub
    
    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub
    
    Select Case Index
            
        Case 3 To 7 'cta contable (va a la BD de la contabilidad)
            If txtAux(Index).Text = "" Then
                txtAux2(Index).Text = ""
                Exit Sub
            End If
            If Modo = 1 And ContieneCaracterBusqueda(txtAux(Index).Text) Then Exit Sub     'Busquedas
            txtAux2(Index).Text = PonerNombreCuenta(txtAux(Index), Modo, , Conexion)
    End Select
End Sub


Private Sub CargaCombo(Index As Integer)
'0-IVA, 1-IGIC, 2-Bien de inversi�n, 3.- REA(Regimen especial agrario)
    Combo1(Index).Clear
   
    Select Case Index
        Case 0 'combo Tipos de IVA
            Combo1(Index).AddItem "IVA"
            Combo1(Index).ItemData(Combo1(Index).NewIndex) = 0
    
            Combo1(Index).AddItem "IGIC"
            Combo1(Index).ItemData(Combo1(Index).NewIndex) = 1
    
            Combo1(Index).AddItem "BIEN DE INVERSI�N"
            Combo1(Index).ItemData(Combo1(Index).NewIndex) = 2
    
            Combo1(Index).AddItem "R.E.A."
            Combo1(Index).ItemData(Combo1(Index).NewIndex) = 3
    End Select
End Sub


Private Sub PonerOpcionesMenu()
    PonerOpcionesMenuGeneral Me
End Sub


Private Function SepuedeBorrar() As Boolean
    SepuedeBorrar = True
End Function


Private Sub KEYpress(KeyAscii As Integer)
    If KeyAscii = 13 Then 'ENTER
        KeyAscii = 0
        SendKeys "{tab}"
    ElseIf KeyAscii = 27 And Modo = 2 Then Unload Me  'ESC
    End If
End Sub


Private Sub PonerContRegIndicador()
'si estamos en modo ver registros muestra el numero de registro en el que estamos
'situados del total de registros mostrados: 1 de 24
Dim cadReg As String

    If (Modo = 2 Or Modo = 0) Then
        cadReg = PonerContRegistros(Me.adodc1)
        If CadB = "" Then
            lblIndicador.Caption = cadReg
        Else
            lblIndicador.Caption = "BUSQUEDA: " & cadReg
        End If
    End If
End Sub



Private Sub BotonImprimir()
Dim cadParam As String
Dim cadFormula As String

    'A�adir el parametro de Empresa
    cadParam = "|pEmpresa=""" & vEmpresa.nomEmpre & """|"
'    cadFormula = "{empresas.codgrupo}=" & vEmpresa.codGrupo
    With frmImprimir
        .FormulaSeleccion = cadFormula
        .OtrosParametros = cadParam
        .NumeroParametros = 1 'Solo parametro de la empresa
        .SoloImprimir = False
        .Opcion = 20 'Opcionlistado
        .Show vbModal
    End With
End Sub

'Private Sub printNou()
'
'    With frmImprimir2
'        .cadTabla2 = "tiposiva"
'        .Informe2 = "rTiposIVA_conta.rpt"
'        If CadB <> "" Then
'            .cadRegSelec = SQL2SF(CadB)
'        Else
'            .cadRegSelec = ""
'        End If
'        .cadRegActua = POS2SF(adodc1, Me)
'        .cadTodosReg = ""
'        '.OtrosParametros2 = "pEmpresa='" & vEmpresa.NomEmpre & "'|pOrden={tarjbanc.nomtarje}|"
'        .OtrosParametros2 = "pEmpresa='" & vEmpresa.nomEmpre & "'|"
'        .NumeroParametros2 = 1
'        .MostrarTree2 = False
'        .InfConta2 = False
'        .ConSubInforme2 = False
'
'        .Show vbModal
'    End With
'End Sub

' ### [DavidV] 26/04/2006: Activar/desactivar la rueda del rat�n.
Private Sub DataGrid1_GotFocus()
  WheelHook DataGrid1
End Sub
Private Sub DataGrid1_Lostfocus()
  WheelUnHook
End Sub

