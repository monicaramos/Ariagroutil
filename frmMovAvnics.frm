VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmMovAvnics 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Movimientos AVNICS"
   ClientHeight    =   8220
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13500
   Icon            =   "frmMovAvnics.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8220
   ScaleWidth      =   13500
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkAux 
      BackColor       =   &H80000005&
      Height          =   255
      Index           =   0
      Left            =   8820
      TabIndex        =   18
      Tag             =   "Contabilizado|N|N|0|1|movim|intconta|0|N|"
      Top             =   2745
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.TextBox txtAux 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   290
      Index           =   6
      Left            =   8100
      MaxLength       =   17
      TabIndex        =   17
      Tag             =   "Retencion|N|N|0|999999.99|movim|timport2|##,###,##0.00||"
      Text            =   "Imp"
      Top             =   2745
      Width           =   555
   End
   Begin VB.TextBox txtAux 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   290
      Index           =   5
      Left            =   7425
      MaxLength       =   17
      TabIndex        =   16
      Tag             =   "Importe|N|N|0|999999.99|movim|timporte|##,###,##0.00||"
      Text            =   "Imp"
      Top             =   2745
      Width           =   555
   End
   Begin VB.CommandButton btnBuscar 
      Appearance      =   0  'Flat
      Caption         =   "+"
      Height          =   290
      Index           =   1
      Left            =   4080
      MaskColor       =   &H00000000&
      TabIndex        =   15
      ToolTipText     =   "Buscar Fecha"
      Top             =   2760
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.TextBox txtAux 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   290
      Index           =   4
      Left            =   6120
      MaxLength       =   17
      TabIndex        =   4
      Tag             =   "Importe|N|N|0|999999.99|movim|timporte|##,###,##0.00||"
      Text            =   "Imp"
      Top             =   2760
      Width           =   555
   End
   Begin VB.TextBox txtAux 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   290
      Index           =   3
      Left            =   4320
      MaxLength       =   35
      TabIndex        =   3
      Tag             =   "Concepto|T|N|||movim|concepto|||"
      Text            =   "Concep"
      Top             =   2760
      Width           =   1755
   End
   Begin VB.TextBox txtAux 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   290
      Index           =   2
      Left            =   3480
      MaxLength       =   10
      TabIndex        =   2
      Tag             =   "Fecha|F|N|||movim|fechamov|dd/mm/yyyy|S|"
      Text            =   "Fecha"
      Top             =   2760
      Width           =   555
   End
   Begin VB.TextBox txtAux 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   290
      Index           =   0
      Left            =   180
      MaxLength       =   6
      TabIndex        =   0
      Tag             =   "Avnic|N|N|1|999999|movim|codavnic|000000|S|"
      Text            =   "Cod"
      Top             =   2760
      Width           =   555
   End
   Begin VB.TextBox txtAux2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   285
      Index           =   1
      Left            =   1560
      TabIndex        =   8
      Top             =   2760
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.CommandButton btnBuscar 
      Appearance      =   0  'Flat
      Caption         =   "+"
      Height          =   290
      Index           =   0
      Left            =   1320
      MaskColor       =   &H00000000&
      TabIndex        =   7
      ToolTipText     =   "Buscar Cliente"
      Top             =   2760
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   10755
      TabIndex        =   5
      Top             =   7545
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   12015
      TabIndex        =   6
      Top             =   7545
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox txtAux 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   290
      Index           =   1
      Left            =   780
      MaxLength       =   4
      TabIndex        =   1
      Tag             =   "Ejercicio|N|N|0|9999|movim|anoejerc|0000|S|"
      Text            =   "Ejer"
      Top             =   2760
      Width           =   555
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "frmMovAvnics.frx":000C
      Height          =   6825
      Left            =   120
      TabIndex        =   11
      Top             =   540
      Width           =   13125
      _ExtentX        =   23151
      _ExtentY        =   12039
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
      Left            =   12015
      TabIndex        =   14
      Top             =   7515
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Height          =   555
      Index           =   1
      Left            =   120
      TabIndex        =   9
      Top             =   7440
      Width           =   2385
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
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   2175
      End
   End
   Begin MSAdodcLib.Adodc adodc1 
      Height          =   330
      Left            =   4440
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
      TabIndex        =   12
      Top             =   0
      Width           =   13500
      _ExtentX        =   23813
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
            Object.ToolTipText     =   "Borrar Turno"
            Object.Tag             =   "2"
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Total selección"
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
         TabIndex        =   13
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
         HelpContextID   =   2
         Shortcut        =   ^N
      End
      Begin VB.Menu mnModificar 
         Caption         =   "&Modificar"
         HelpContextID   =   2
         Shortcut        =   ^M
      End
      Begin VB.Menu mnEliminar 
         Caption         =   "&Eliminar"
         HelpContextID   =   2
         Shortcut        =   ^E
      End
      Begin VB.Menu mnBarra2 
         Caption         =   "-"
      End
      Begin VB.Menu mnBorrarTurno 
         Caption         =   "&Borrar Turno"
         HelpContextID   =   2
      End
      Begin VB.Menu mnTotalSeleccion 
         Caption         =   "&Total Seleccion"
      End
      Begin VB.Menu mnImprimir 
         Caption         =   "&Imprimir"
      End
      Begin VB.Menu mnBarra3 
         Caption         =   "-"
      End
      Begin VB.Menu mnSalir 
         Caption         =   "&Salir"
         Shortcut        =   ^S
      End
   End
End
Attribute VB_Name = "frmMovAvnics"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' +-+-+-+-+-+-+-+-+-+-+-+-
' +-+- Autor: MANOLO  +-+-
' +-+-+-+-+-+-+-+-+-+-+-+-

' **************** PER A QUE FUNCIONE EN UN ATRE MANTENIMENT ********************
' 0. Posar-li l'atribut Datasource a "adodc1" del Datagrid1. Canviar el Caption
'    del formulari
' 1. Canviar els TAGs i els Maxlength de TextAux(0) i TextAux(1)
' 2. En PonerModo(vModo) repasar els indexs del botons, per si es canvien
' 3. En la funció BotonAnyadir() canviar la taula i el camp per a SugerirCodigoSiguienteStr
' 4. En la funció BotonBuscar() canviar el nom de la clau primaria
' 5. En la funció BotonEliminar() canviar la pregunta, les descripcions de la
'    variable SQL i el contingut del DELETE
' 6. En la funció posamaxlength() repasar el maxlength de TextAux(0)
' 7. En Form_Load() repasar la barra d'iconos (per si es vol canviar algún) i
'    canviar la consulta per a vore tots els registres
' 8. En Toolbar1_ButtonClick repasar els indexs de cada botó per a que corresponguen
' 9. En la funció CargaGrid canviar l'ORDER BY (normalment per la clau primaria);
'    canviar ademés els noms dels camps, el format i si fa falta la cantitat;
'    repasar els index dels botons modificar i eliminar.
'    NOTA: si en Form_load ya li he posat clausula WHERE, canviar el `WHERE` de
'    `SQL = CadenaConsulta & " WHERE " & vSQL` per un `AND`
' 10. En txtAux_LostFocus canviar el mensage i el format del camp
' 11. En la funció DatosOk() canviar els arguments de DevuelveDesdeBD i el mensage
'    en cas d'error
' 12. En la funció SepuedeBorrar() canviar les comprovacions per a vore si es pot
'    borrar el registre
' ********************************************************************************

Option Explicit

Public DatosADevolverBusqueda As String    'Tendra el nº de text que quiere que devuelva, empipados
Public Event DatoSeleccionado(CadenaSeleccion As String)

'codi per al registe que s'afegix al cridar des d'atre formulari.
'Obrir en modo Insertar i tornar datos del registre insertat
Public NuevoCodigo As String

'codigo que tiene el campo en el momento que se llama desde otro formulario
'nos situamos en ese valor
Public CodigoActual As String

Private WithEvents frmFPa As frmAvnics
Attribute frmFPa.VB_VarHelpID = -1
Private WithEvents frmB As frmBuscaGrid
Attribute frmB.VB_VarHelpID = -1
Private WithEvents frmC As frmCal
Attribute frmC.VB_VarHelpID = -1
Private WithEvents frmavn As frmAvnics
Attribute frmavn.VB_VarHelpID = -1


Private CadenaConsulta As String
Private CadB As String
Private PrimeraVez As Boolean
Private HaDevueltoDatos As Boolean

Dim ValorAnt As String
' utilizado para buscar por checks
Private BuscaChekc As String

Dim tipoF As String
Dim Modo As Byte

'----------- MODOS --------------------------------
'   0.-  Formulari net sense cap camp ple
'   1.-  Preparant per a fer la búsqueda
'   2.-  Ja tenim registres i els anem a recorrer
'        ,podem editar-los Edició del camp
'   3.-  Inserció de nou registre
'   4.-  Modificar
'--------------------------------------------------

Private Sub PonerModo(vModo)
Dim b As Boolean
Dim i As Byte

    On Error Resume Next
    
    Modo = vModo
    BuscaChekc = ""
'    PonerIndicador lblIndicador, Modo
    b = (Modo = 2)
    If b Then
        PonerContRegIndicador
    Else
        PonerIndicador lblIndicador, Modo
    End If
    
    For i = 0 To txtAux.Count - 1 'els txtAux del grid
        txtAux(i).visible = Not b
    Next i
    btnBuscar(0).visible = Not b
    btnBuscar(1).visible = Not b
    txtAux2(1).visible = Not b
    chkAux(0).visible = Not b

    cmdAceptar.visible = Not b
    cmdCancelar.visible = Not b
    DataGrid1.Enabled = b
    
    'Si es regresar
    If DatosADevolverBusqueda <> "" Then cmdRegresar.visible = b
    
    If Modo = 3 Or Modo = 4 Or Modo = 1 Then i = 4 'Insertar/Modificar o busqueda
    BloquearImgBuscar Me, i
    
    PonerLongCampos
    PonerModoOpcionesMenu 'Activar/Desact botones de menu segun Modo
    PonerOpcionesMenu 'Activar/Desact botones de menu segun permisos del usuario
    
    'Si estamos modo Modificar bloquear clave primaria
    BloquearTxt txtAux(0), (Modo = 4)
    BloquearTxt txtAux(1), (Modo = 4)
    BloquearTxt txtAux(2), (Modo = 4)
    
    PonerFocoGrid Me.DataGrid1
    
    If Err.Number <> 0 Then Err.Clear
End Sub

Private Sub PonerModoOpcionesMenu()
Dim b As Boolean
On Error Resume Next

    b = (Modo = 2)
    'Busqueda
    Toolbar1.Buttons(2).Enabled = b
    Me.mnBuscar.Enabled = b
    'Ver Todos
    Toolbar1.Buttons(3).Enabled = b
    Me.mnVerTodos.Enabled = b
    
    'Insertar
    Toolbar1.Buttons(6).Enabled = b 'And Not DeConsulta
    Me.mnNuevo.Enabled = b 'And Not DeConsulta
    
    b = (b And adodc1.Recordset.RecordCount > 0) 'And Not DeConsulta
    'Modificar
    Toolbar1.Buttons(7).Enabled = b
    Me.mnModificar.Enabled = b
    'Eliminar
    Toolbar1.Buttons(8).Enabled = b
    Me.mnEliminar.Enabled = b
    'Borrar turno
    Toolbar1.Buttons(9).Enabled = b
    Me.mnBorrarTurno.Enabled = b
    'Total seleccion
    Toolbar1.Buttons(10).Enabled = b
    Me.mnTotalSeleccion.Enabled = b
    'Imprimir
    Toolbar1.Buttons(11).Enabled = True
    Me.mnImprimir.Enabled = True
    
    If Err.Number <> 0 Then Err.Clear
End Sub

Private Sub BotonAnyadir()
    Dim NumF As String
    Dim anc As Single
    Dim i As Integer
    
'   ' ### [Monica] 21/09/2006
'   ' cuando añado se carga todo sql grid estaba la instruccion de abajo
     CargaGrid CadB  'primer de tot carregue tot el grid
'    CargaGrid "codavnic = -1" 'primer de tot carregue tot el grid
   
    CadB = ""
    '******************** canviar taula i camp **************************
    If (DatosADevolverBusqueda <> "") And NuevoCodigo <> "" Then
        NumF = NuevoCodigo
    Else
        NumF = "" 'SugerirCodigoSiguienteStr("scaalb", "codclave")
    End If
    '********************************************************************
    
    'Situamos el grid al final
    AnyadirLinea DataGrid1, adodc1
         
    anc = DataGrid1.Top
    If DataGrid1.Row < 0 Then
        anc = anc + 206
    Else
        anc = anc + DataGrid1.RowTop(DataGrid1.Row) + 5
    End If
    
    LLamaLineas anc, 3 '(limpia los campos)
    
    txtAux(0).Text = NumF
 '   FormateaCampo txtAux(0)
    For i = 1 To txtAux.Count - 1
        txtAux(i).Text = ""
    Next i
    txtAux2(1).Text = ""
    txtAux(2).Text = Format(Now, "dd/mm/yyyy") ' Fecha x defecto
    chkAux(0).Value = 0
       
    'Ponemos el foco
    If (DatosADevolverBusqueda <> "") And NuevoCodigo <> "" Then
        PonerFoco txtAux(0)
    Else
        PonerFoco txtAux(0)
    End If
End Sub

Private Sub BotonVerTodos()
    CadB = ""
    CargaGrid ""
    PonerModo 2
End Sub

Private Sub BotonBuscar()
Dim i As Integer
    ' ***************** canviar per la clau primaria ********
    CargaGrid "movim.codavnic = -1"
    '*******************************************************************************
    'Buscar
    For i = 0 To txtAux.Count - 1
        txtAux(i).Text = ""
    Next i
    txtAux2(1).Text = ""
    chkAux(0).Value = 0
    LLamaLineas DataGrid1.Top + 216, 1
    PonerFoco txtAux(0)
End Sub

Private Sub BotonModificar()
    Dim anc As Single
    Dim i As Integer

    Screen.MousePointer = vbHourglass
    
    If DataGrid1.Bookmark < DataGrid1.FirstRow Or DataGrid1.Bookmark > (DataGrid1.FirstRow + DataGrid1.VisibleRows - 1) Then
        i = DataGrid1.Bookmark - DataGrid1.FirstRow
        DataGrid1.Scroll 0, i
        DataGrid1.Refresh
    End If
    
    If DataGrid1.Row < 0 Then
        anc = 320
    Else
        anc = DataGrid1.RowTop(DataGrid1.Row) + 545
    End If

    'Llamamos al form
    For i = 0 To 1
        txtAux(i).Text = DataGrid1.Columns(i).Text
    Next i
    
    For i = 2 To txtAux.Count - 1
        txtAux(i).Text = DataGrid1.Columns(i + 1).Text
    Next i
    Me.chkAux(0).Value = Me.adodc1.Recordset!intconta
    
    txtAux2(1).Text = DataGrid1.Columns(2).Text

    LLamaLineas anc, 4
   
    'Como es modificar
    
    PonerFoco txtAux(3)
    Screen.MousePointer = vbDefault
End Sub

Private Sub LLamaLineas(alto As Single, xModo As Byte)
Dim i As Byte

    DeseleccionaGrid Me.DataGrid1
    PonerModo xModo
    
    'Fijamos el ancho
    For i = 0 To txtAux.Count - 1
        txtAux(i).Top = alto
    Next i
    Me.chkAux(0).Top = alto
    btnBuscar(0).Top = alto
    btnBuscar(1).Top = alto
    txtAux2(1).Top = alto
End Sub

Private Sub BotonEliminar()
Dim SQL As String
Dim temp As Boolean

    On Error GoTo Error2
    
    'Ciertas comprobaciones
    If adodc1.Recordset.EOF Then Exit Sub
'    If Not SepuedeBorrar Then Exit Sub
    
    '*************** canviar els noms i el DELETE **********************************
    SQL = "¿Seguro que desea eliminar el Movimiento?"
    SQL = SQL & vbCrLf & "Fecha: " & adodc1.Recordset.Fields(3)
    
    If MsgBox(SQL, vbQuestion + vbYesNo) = vbYes Then
        'Hay que eliminar
        NumRegElim = adodc1.Recordset.AbsolutePosition
        SQL = "Delete from movim where codavnic=" & adodc1.Recordset!codavnic
        SQL = SQL & " and anoejerc=" & adodc1.Recordset!anoejerc
        SQL = SQL & " and fechamov=" & adodc1.Recordset!Fechamov
        
        conn.Execute SQL
        CargaGrid CadB
        temp = SituarDataTrasEliminar(adodc1, NumRegElim, True)
        PonerModoOpcionesMenu
        adodc1.Recordset.Cancel
    End If
    Exit Sub
    
Error2:
    Screen.MousePointer = vbDefault
    MuestraError Err.Number, "Eliminando registro", Err.Description
End Sub

Private Sub PonerLongCampos()

'Modificar el MaxLength del campo en funcion de si es modo de búsqueda o no
'para los campos que permitan introducir criterios más largos del tamaño del campo
'Si maxlength=6 para codprove, en modo busqueda Maxlenth=13 'el doble + 1
    
    'para los txtAux
    PonerLongCamposGnral Me, Modo, 3
End Sub

Private Sub btnBuscar_Click(Index As Integer)
    TerminaBloquear
    Select Case Index
        Case 0 'Avnic
            Set frmavn = New frmAvnics
            frmavn.DatosADevolverBusqueda = "0|1|4|"
            frmavn.CodigoActual = txtAux(0).Text
            frmavn.Show vbModal
            Set frmavn = Nothing
            PonerFoco txtAux(0)
        Case 1 ' Fecha
            Dim esq As Long
            Dim dalt As Long
            Dim menu As Long
            Dim obj As Object
        
            Set frmC = New frmCal
            
            esq = btnBuscar(Index).Left
            dalt = btnBuscar(Index).Top
                
            Set obj = btnBuscar(Index).Container
              
              While btnBuscar(Index).Parent.Name <> obj.Name
                    esq = esq + obj.Left
                    dalt = dalt + obj.Top
                    Set obj = obj.Container
              Wend
            
            menu = Me.Height - Me.ScaleHeight 'ací tinc el heigth del menú i de la toolbar
        
            frmC.Left = esq + btnBuscar(Index).Parent.Left + 30
            frmC.Top = dalt + btnBuscar(Index).Parent.Top + btnBuscar(Index).Height + menu - 40
        
            btnBuscar(Index).Tag = Index '<===
            ' *** repasar si el camp es txtAux o Text1 ***
            If txtAux(2).Text <> "" Then frmC.NovaData = txtAux(2).Text
            ' ********************************************
        
            frmC.Show vbModal
            Set frmC = Nothing
            ' *** repasar si el camp es txtAux o Text1 ***
            PonerFoco txtAux(2) '<===
            ' ********************************************
        
            
    End Select
    If Modo = 4 Then BLOQUEADesdeFormulario2 Me, adodc1, 1
End Sub

Private Sub chkAux_Click(Index As Integer)
    If Modo = 1 Then
        'Buscqueda
        If InStr(1, BuscaChekc, "chkAux(" & Index & ")") = 0 Then BuscaChekc = BuscaChekc & "chkAux(" & Index & ")|"
    End If
End Sub

Private Sub chkAux_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub cmdAceptar_Click()
Dim i As Integer

    Select Case Modo
        Case 1 'BUSQUEDA
            CadB = ObtenerBusqueda(Me, BuscaChekc)
            If CadB <> "" Then
                CargaGrid CadB
                PonerModo 2
'                lblIndicador.Caption = "BUSQUEDA: " & PonerContRegistros(Me.adodc1)
                PonerFocoGrid Me.DataGrid1
            End If
            
        Case 3 'INSERTAR
            If DatosOk Then
                If InsertarDesdeForm2(Me, 1) Then
                    CargaGrid
                    If (DatosADevolverBusqueda <> "") And NuevoCodigo <> "" Then
                        cmdCancelar_Click
    '                    If Not adodc1.Recordset.EOF Then adodc1.Recordset.MoveLast
                        If Not adodc1.Recordset.EOF Then
                            adodc1.Recordset.Find (adodc1.Recordset.Fields(0).Name & " =" & NuevoCodigo)
                        End If
                        cmdRegresar_Click
                    Else
                        BotonAnyadir
                    End If
                    CadB = ""
                End If
            End If
            
        Case 4 'MODIFICAR
            If DatosOk Then
                If ModificaDesdeFormulario2(Me) Then
                    TerminaBloquear
                    i = adodc1.Recordset.Fields(0)
                    PonerModo 2
                    CargaGrid CadB
                    adodc1.Recordset.Find (adodc1.Recordset.Fields(0).Name & " =" & i)
                    PonerFocoGrid Me.DataGrid1
                    
                End If
            End If
    End Select
End Sub

Private Sub cmdCancelar_Click()
On Error Resume Next

    Select Case Modo
        Case 1 'BUSQUEDA
            CargaGrid CadB
        Case 3 'INSERTAR
            DataGrid1.AllowAddNew = False
            If Not adodc1.Recordset.EOF Then adodc1.Recordset.MoveFirst
        Case 4 'MODIFICAR
            TerminaBloquear
    End Select
    
    If Not adodc1.Recordset.EOF Then
'        CargaForaGrid
    Else
        LimpiarCampos
    End If
    
    PonerModo 2
    
    PonerFocoGrid Me.DataGrid1
    
    If Err.Number <> 0 Then Err.Clear
End Sub

Private Sub cmdRegresar_Click()
Dim cad As String
Dim i As Integer
Dim J As Integer
Dim Aux As String

    If adodc1.Recordset.EOF Then
        MsgBox "Ningún registro devuelto.", vbExclamation
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

Private Sub DataGrid1_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub DataGrid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
Dim i As Byte
    
    For i = 0 To txtAux.Count - 1
       txtAux(i).Text = ""
    Next i
    
    PonerContRegIndicador
End Sub

Private Sub DataGrid1_DblClick()
    If cmdRegresar.visible Then cmdRegresar_Click
End Sub

Private Sub Form_Activate()
    Screen.MousePointer = vbDefault
    
    If PrimeraVez Then
        PrimeraVez = False
        If (DatosADevolverBusqueda <> "") And NuevoCodigo <> "" Then
            BotonAnyadir
        Else
            PonerModo 2
             If Me.CodigoActual <> "" Then
                SituarData Me.adodc1, "codempre=" & CodigoActual, "", True
            End If
        End If
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
'       .Buttons(9).Image = 20   'Eliminar Turno
        .Buttons(10).Image = 21   'Totales
        .Buttons(11).Image = 10  'Imprimir
        .Buttons(12).Image = 11  'Salir
    End With

    'cargar IMAGES de busqueda

    '****************** canviar la consulta *********************************+
    CadenaConsulta = "Select movim.codavnic, movim.anoejerc, avnic.nombrper, "
    CadenaConsulta = CadenaConsulta & "movim.fechamov, movim.concepto, movim.timporte, "
    CadenaConsulta = CadenaConsulta & "movim.timport1, movim.timport2, movim.intconta, IF(intconta=1,'*','') as dintconta  "
    CadenaConsulta = CadenaConsulta & "from movim, avnic WHERE movim.codavnic=avnic.codavnic and movim.anoejerc=avnic.anoejerc "
    '************************************************************************
    
    CadB = ""
    CargaGrid "movim.codavnic = -1"
   
End Sub

Private Sub Form_Unload(Cancel As Integer)
'    CheckValueGuardar Me.Name, Me.chkVistaPrevia.Value
    DesBloqueoManual ("BORTUR")
    Screen.MousePointer = vbDefault
End Sub

'Private Sub frmB_Selecionado(CadenaDevuelta As String)
'   txtAux(6).Text = RecuperaValor(CadenaDevuelta, 1)
'   txtAux(6).Text = Format(txtAux(6).Text, "00000000")
'End Sub

Private Sub frmC_Selec(vFecha As Date)
    ' *** repasar si el camp es txtAux o Text1 ***
    txtAux(2).Text = Format(vFecha, "dd/mm/yyyy") '<===
    ' ********************************************
End Sub

Private Sub frmAvn_DatoSeleccionado(CadenaSeleccion As String)
    txtAux(0).Text = RecuperaValor(CadenaSeleccion, 1)
    FormateaCampo txtAux(0)
    txtAux(1).Text = RecuperaValor(CadenaSeleccion, 2)
    txtAux2(1).Text = RecuperaValor(CadenaSeleccion, 3)
End Sub

Private Sub mnBuscar_Click()
    BotonBuscar
End Sub

Private Sub mnEliminar_Click()
    BotonEliminar
End Sub

Private Sub mnImprimir_Click()
    BotonImprimir
End Sub

Private Sub mnModificar_Click()
    'Comprobaciones
    '--------------
    If adodc1.Recordset.EOF Then Exit Sub
    If adodc1.Recordset.RecordCount < 1 Then Exit Sub
    
    'Preparar para modificar
    '-----------------------
    If BLOQUEADesdeFormulario2(Me, adodc1, 1) Then BotonModificar
End Sub

Private Sub mnNuevo_Click()
    BotonAnyadir
End Sub

Private Sub mnSalir_Click()
    Screen.MousePointer = vbDefault
    Unload Me
End Sub

Private Sub mnTotalSeleccion_Click()
    CalcularSumaPantalla
End Sub

Private Sub mnVerTodos_Click()
    BotonVerTodos
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
        Case 2
                mnBuscar_Click
        Case 3
                mnVerTodos_Click
        Case 6
                mnNuevo_Click
        Case 7
                mnModificar_Click
        Case 8
                mnEliminar_Click
        Case 10
                mnTotalSeleccion_Click
        Case 11 'Imprimir
                mnImprimir_Click
        Case 12 'Salir
              mnSalir_Click
    End Select
End Sub

Private Sub CargaGrid(Optional vSQL As String)
    Dim SQL As String
    Dim tots As String
    
'    adodc1.ConnectionString = Conn
    If vSQL <> "" Then
        SQL = CadenaConsulta & " AND " & vSQL
    Else
        SQL = CadenaConsulta
    End If
    '********************* canviar el ORDER BY *********************++
    SQL = SQL & " ORDER BY codavnic, anoejerc, fechamov"
    '**************************************************************++
    
    CargaGridGnral Me.DataGrid1, Me.adodc1, SQL, PrimeraVez
    
    tots = "S|txtAux(0)|T|Avnic|800|;S|txtAux(1)|T|Ejerc.|650|;S|btnBuscar(0)|B||195|;S|txtAux2(1)|T|Nombre|3200|;S|txtAux(2)|T|Fecha|1200|;S|btnBuscar(1)|B||195|;"
    tots = tots & "S|txtAux(3)|T|Concepto|2800|;S|txtAux(4)|T|Importe|1100|;"
    tots = tots & "S|txtAux(5)|T|Bruto|1100|;S|txtAux(6)|T|Retencion|1100|;N||||0|;S|chkAux(0)|CB|IC|360|;"
    
    arregla tots, DataGrid1, Me
    DataGrid1.ScrollBars = dbgAutomatic
      
    If Not adodc1.Recordset.EOF Then
'        CargaForaGrid
    Else
        LimpiarCampos
    End If
    
    DataGrid1.Columns(0).Alignment = dbgRight
'    DataGrid1.Columns(2).Alignment = dbgRight
      
End Sub

Private Sub txtAux_GotFocus(Index As Integer)
    ConseguirFocoLin txtAux(Index)
End Sub

Private Sub txtAux_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    KEYdown KeyCode
End Sub

Private Sub txtAux_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = teclaBuscar Then
        If Modo = 1 Or Modo = 3 Or Modo = 4 Then
            Select Case Index
                Case 5: 'avnic
                    KeyAscii = 0
                    btnBuscar_Click (0)
                Case 2: 'fecha
                    KeyAscii = 0
                    btnBuscar_Click (1)
            End Select
        End If
    Else
        KEYpress KeyAscii
    End If
End Sub

Private Sub txtAux_LostFocus(Index As Integer)
Dim cadMen As String
Dim Famia As String

    If Modo = 1 Then Exit Sub 'Busquedas
    If Not PerderFocoGnral(txtAux(Index), Modo) Then Exit Sub
    
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub
    
    Select Case Index
        Case 0 'codavnic
            PonerFormatoEntero txtAux(Index)
            If txtAux(0).Text <> "" And txtAux(1).Text <> "" Then
                txtAux2(1).Text = DevuelveDesdeBDNew(cPTours, "avnic", "nombrper", "codavnic", txtAux(0).Text, "N", , "anoejerc", txtAux(1).Text, "N")
                If txtAux2(1).Text = "" Then
                    MsgBox "No existe este avnic. Reintroduzca.", vbExclamation
                    txtAux(0).Text = ""
                    txtAux(1).Text = ""
                    PonerFoco txtAux(0)
                End If
            End If
            
        Case 1
            If txtAux(0).Text <> "" And txtAux(1).Text <> "" Then
                txtAux2(1).Text = DevuelveDesdeBDNew(cPTours, "avnic", "nombrper", "codavnic", txtAux(0).Text, "N", , "anoejerc", txtAux(1).Text, "N")
                If txtAux2(1).Text = "" Then
                    MsgBox "No existe este avnic. Reintroduzca.", vbExclamation
                    txtAux(0).Text = ""
                    txtAux(1).Text = ""
                    PonerFoco txtAux(0)
                End If
            End If
            
        Case 2 'FECHA
            PonerFormatoFecha txtAux(Index)
        
        Case 3 'CONCEPTO
            txtAux(Index).Text = UCase(txtAux(Index).Text)
            
        Case 4, 5, 6 'IMPORTE
            If txtAux(Index).Text <> "" Then PonerFormatoDecimal txtAux(Index), 3
            
    End Select
End Sub

Private Function DatosOk() As Boolean
Dim b As Boolean
Dim SQL As String
Dim Fpag As String

    b = CompForm(Me)
    If Not b Then Exit Function
    
    If Modo = 3 Then   'Estamos insertando
        'comprobar si ya existe el campo de clave primaria
        If ExisteCP(txtAux(0)) Then b = False
        
    End If
    
    DatosOk = b
End Function

Private Sub PonerOpcionesMenu()
    PonerOpcionesMenuGeneral Me
    PonerOpcionesMenuGeneralNew Me
End Sub

Private Sub LimpiarCampos()
Dim i As Integer

    On Error Resume Next

'    For i = 11 To 13
'        txtAux(i).Text = ""
'    Next i
'    txtAux2(11).Text = ""
'    txtAux2(12).Text = ""
    
    If Err.Number <> 0 Then Err.Clear
End Sub

Private Sub BotonImprimir()
        frmInfMovAvnics.Show vbModal
End Sub

Private Sub KEYpress(KeyAscii As Integer)
    If KeyAscii = 13 Then 'ENTER
        KeyAscii = 0
        SendKeys "{tab}"
    ElseIf KeyAscii = 27 Then 'ESC
        If (Modo = 0 Or Modo = 2) Then Unload Me
    End If
End Sub

Private Sub KEYBusqueda(KeyAscii As Integer, indice As Integer)
    KeyAscii = 0
'    imgBuscar_Click (indice)
    btnBuscar_Click (0)
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

' ### [DavidV] 26/04/2006: Activar/desactivar la rueda del ratón.
Private Sub DataGrid1_GotFocus()
  WheelHook DataGrid1
End Sub
Private Sub DataGrid1_Lostfocus()
  WheelUnHook
End Sub
Private Sub CalcularSumaPantalla()
Dim RS As ADODB.Recordset
Dim SQL As String

  If Not adodc1.Recordset.EOF And CadB = "" Then CadB = "codavnic > 0"
  If CadB <> "" Then
     SQL = "select sum(timporte), sum(timport1), sum(timport2) FROM movim "
     SQL = SQL & " WHERE " & CadB
     Set RS = New ADODB.Recordset ' Crear objeto
     RS.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText ' abrir cursor
      If Not RS.EOF Then
        SQL = "Importe : " & Format(RS.Fields(0), "####,##0.00")
        SQL = SQL & vbCrLf & "Bruto : " & Format(RS.Fields(1), "####,##0.00")
        SQL = SQL & vbCrLf & "Retención : " & Format(RS.Fields(2), "####,##0.00")
        MsgBox "Totales Selección: " & vbCrLf & vbCrLf & SQL, vbInformation
      End If
     RS.Close
     Set RS = Nothing
    Else
        MsgBox "Haga primero una selección para ver Totales.", vbInformation
  End If
End Sub
