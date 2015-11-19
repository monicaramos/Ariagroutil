VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmAyuda 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Mensajes"
   ClientHeight    =   7620
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   14160
   Icon            =   "frmAyuda.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7620
   ScaleWidth      =   14160
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FrameVariedades 
      Height          =   5790
      Left            =   -30
      TabIndex        =   0
      Top             =   0
      Width           =   7050
      Begin VB.CommandButton cmdAcepCuentas 
         Caption         =   "Aceptar"
         Height          =   375
         Left            =   4080
         TabIndex        =   3
         Top             =   5160
         Width           =   1215
      End
      Begin VB.CommandButton cmdCanVariedades 
         Caption         =   "Cancelar"
         Height          =   375
         Left            =   5520
         TabIndex        =   2
         Top             =   5160
         Width           =   1215
      End
      Begin MSComctlLib.ListView ListView7 
         Height          =   4155
         Left            =   225
         TabIndex        =   1
         Top             =   810
         Width           =   6525
         _ExtentX        =   11509
         _ExtentY        =   7329
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Código"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Variedad"
            Object.Width           =   3881
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Clase "
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Descripcion"
            Object.Width           =   3706
         EndProperty
      End
      Begin VB.Label Label5 
         Caption         =   "Cuentas Contables"
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
         Height          =   375
         Left            =   270
         TabIndex        =   4
         Top             =   270
         Width           =   5145
      End
      Begin VB.Image imgCheck 
         Height          =   240
         Index           =   5
         Left            =   600
         Picture         =   "frmAyuda.frx":000C
         Top             =   5160
         Width           =   240
      End
      Begin VB.Image imgCheck 
         Height          =   240
         Index           =   4
         Left            =   240
         Picture         =   "frmAyuda.frx":0156
         Top             =   5160
         Width           =   240
      End
   End
End
Attribute VB_Name = "frmAyuda"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'====================== VBLES PUBLICAS ================================

Public Event DatoSeleccionado(CadenaSeleccion As String)

Public OpcionMensaje As Byte
'======================================
'==== FACTURACION =====================
' 1 .- Mensaje de Cobros Pendientes
' 2 .- Mensaje de No hay suficiente Stock para pasar de Pedido a Albaran
' 3 .- Mensaje Acerca de...
' 4 .- Listado de los Nº de Serie de un Articulo
' 5 .- Seleccionar tipo de Componente a Mostrar en Mant. de Nº de Series
' 6 .- Mostrar Prefacturacion de Albaranes
' 7 .- Mostrar Prefacturacion Mantenimientos
' 8 .- Mostrar lista clientes para seleccionar los que queremos imprimir (Etiquetas)
' 9 .- Mostrar lista Proveedores para seleccionar los que queremos imprimir (Etiquetas)
'10 .- Mostrar lista de Errores de las facturas NO contabilizadas
'11 .- Mostrar lista lineas de factura a Rectificar para seleccionar las q queremos traer al Albaran de FAct. Rectificativa
'12 .- Mostrar Albaranes del Rango que no se van a Facturar. (Facturar Albaranes Venta)

'13 .- Mostrar Errores
'14 .- Mostrar Empresas existentes en el sistema



'15 .- Mostrar lista de articulos para imprimir etiquetas estanteria
'16 .- Lista de articulos para corregir importes
'17 .- Etiquetas clientes. LO MISMO QUE EL 8 pero hecho por david
'18 .- Mantenimientos. paso ejercicio siguiente a actual
'19 .- Lista de palets de venta asociados al pedido del que se va a generar el albaran
'20 .- Lista de pedidos sin numero de albaran asignado.
'21 .


'22 .- Facturas a cuenta que se han hecho al cliente

Public cadWHERE As String 'Cadena para pasarle la WHERE de la SELECT de los cobros pendientes o de Pedido(para comp. stock)
                          'o CodArtic para seleccionar los Nº Series
                          'para cargar el ListView
                          
Public cadWHERE2 As String

Public vCampos As String 'Articulo y cantidad Empipados para Nº de Series
                         'Tambien para pasar el nombre de la tabla de lineas (sliped, slirep,...)
                         'Dependiendo desde donde llamemos, de Pedidos o Reparaciones


'====================== VBLES LOCALES ================================

Dim PulsadoSalir As Boolean 'Solo salir con el boton de Salir no con aspa del form
Dim PrimeraVez As Boolean

'Para los Nº de Serie
Dim TotalArray As Integer
Dim codArtic() As String
Dim cantidad() As Integer




Private Sub cmdacepCuentas_Click()
Dim CADENA As String
    'Cargo las variedades marcadas
    CADENA = ""
    For NumRegElim = 1 To ListView7.ListItems.Count
        If ListView7.ListItems(NumRegElim).Checked Then
             CADENA = CADENA & "'" & Trim(ListView7.ListItems(NumRegElim).Text) & "',"
        End If
    Next NumRegElim
    ' quitamos la ultima coma
    If CADENA <> "" Then
        CADENA = Mid(CADENA, 1, Len(CADENA) - 1)
    End If
    
    RaiseEvent DatoSeleccionado(CADENA)
    Unload Me
End Sub

Private Sub cmdCanVariedades_Click()
    RaiseEvent DatoSeleccionado("")
    Unload Me
End Sub

'---monica
'Private Sub cmdCorrecotrPrecios_Click(Index As Integer)
'Dim SQL As String
'
'
'    If Index = 0 Then
'
'
'        'Compruebo si ha seleccionado algun articulo de los de precio ultima compra=0
'        cadWHERE2 = ""
'        SQL = ""
'        For TotalArray = 1 To Me.ListView4.ListItems.Count
'            If ListView4.ListItems(TotalArray).Checked Then
'                If ListView4.ListItems(TotalArray).Tag = "" Then
'                    SQL = SQL & "M"
'                Else
'                    cadWHERE2 = cadWHERE2 & "M"
'                End If
'            End If
'        Next
'
'        If SQL <> "" Then
'            MsgBox "No puede actualizar los articulos cuyo precio ultima compra sea 0", vbExclamation
'            Exit Sub
'        End If
'
'        If cadWHERE2 = "" Then
'            MsgBox "Seleccione algun articulo para actualizar", vbExclamation
'            Exit Sub
'        End If
'
'        'Llegado aqui todo correcto. Hacemos la pregunta de actualizar y a correr
'        SQL = "artículo"
'        If Len(cadWHERE2) > 1 Then SQL = SQL & "s"
'        SQL = "Va a actualizar los precios de " & Len(cadWHERE2) & " " & SQL & vbCrLf & vbCrLf & "¿Desea continuar?"
'        If MsgBox(SQL, vbQuestion + vbYesNo) <> vbYes Then Exit Sub
'
'
'        'Aqui esta el proceso de actualizacion de articulos
'        Me.lblIndicadorCorregir.Caption = "Actualización precios"
'        Me.Refresh
'        espera 0.5
'
'       'Para el LOG
'       SQL = cadWHERE & vbCrLf
'       For TotalArray = 1 To Me.ListView4.ListItems.Count
'            If ListView4.ListItems(TotalArray).Checked Then
'                If ListView4.ListItems(TotalArray).Tag <> "" Then SQL = SQL & ListView4.ListItems(TotalArray).Text & "|"
'            End If
'        Next
'        SQL = Mid(SQL, 1, 237)
'
'        '------------------------------------------------------------------------------
'        '  LOG de acciones
'        Set Log = New cLOG
'        Log.Insertar 4, vUsu, "Correccion precios: " & vbCrLf & SQL
'        Set Log = Nothing
'        '-----------------------------------------------------------------------------
'
'
'
'
'
'
'
'
'
'
'        For TotalArray = 1 To Me.ListView4.ListItems.Count
'            If ListView4.ListItems(TotalArray).Checked Then
'                If ListView4.ListItems(TotalArray).Tag <> "" Then
'
'                    'lo metemos en transaccion. Si queremos vamos
'                    Me.lblIndicadorCorregir.Caption = ListView4.ListItems(TotalArray).Text
'                    Me.lblIndicadorCorregir.Refresh
'
'                    ActualizaPrecios TotalArray
'
'
'                End If
'            End If
'        Next
'
'
'    End If
'    Unload Me
'End Sub
'




Private Sub Form_Activate()
Dim OK As Boolean

    
    Select Case OpcionMensaje
        
        Case 21  'Cuentas Contables
            CargarListaFields False
        
        
    End Select
    
    Screen.MousePointer = vbDefault
End Sub


Private Sub Form_Load()
Dim h As Integer, w As Integer
Dim Cad As String
On Error Resume Next

    FrameVariedades.visible = False
    
    PulsadoSalir = True
    PrimeraVez = True
    
    Select Case OpcionMensaje
        Case 21 'variedades
            h = FrameVariedades.Height
            w = FrameVariedades.Width
            PonerFrameVisible FrameVariedades, True, h, w
                
    End Select
    'Me.cmdCancel(indFrame).Cancel = True
    Me.Height = h + 350
    Me.Width = w + 70
    
    If Err.Number <> 0 Then Err.Clear
End Sub








Private Sub Form_Unload(Cancel As Integer)
    If PulsadoSalir = False Then Cancel = 1
End Sub



Private Function ObtenerTamanyosArray() As Boolean
'Para el frame de los Nº de Serie de los Articulos
'En cada indice pone en CodArtic(i) el codigo del articulo
'y en Cantidad(i) la cantidad solicitada de cada codartic
Dim i As Integer, J As Integer

    ObtenerTamanyosArray = False
    'Primero a los campos de la tabla
    TotalArray = -1
    J = 0
    Do
        i = J + 1
        J = InStr(i, vCampos, "·")
        If J > 0 Then TotalArray = TotalArray + 1
    Loop Until J = 0
    
    If TotalArray < 0 Then Exit Function
    
    'Las redimensionaremos
    ReDim codArtic(TotalArray)
    ReDim cantidad(TotalArray)
    
    ObtenerTamanyosArray = True
End Function


Private Function SeparaCampos() As Boolean
'Para el frame de los Nº de Serie de los Articulos
Dim Grupo As String
Dim i As Integer
Dim J As Integer
Dim C As Integer 'Contador dentro del array

    SeparaCampos = False
    i = 0
    C = 0
    Do
        J = i + 1
        i = InStr(J, vCampos, "·")
        If i > 0 Then
            Grupo = Mid(vCampos, J, i - J)
            'Y en la martriz
            InsertaGrupo Grupo, C
            C = C + 1
        End If
    Loop Until i = 0
    SeparaCampos = True
End Function


Private Sub InsertaGrupo(Grupo As String, Contador As Integer)
Dim J As Integer
Dim Cad As String

    J = 0
    Cad = ""
    
    'Cod Artic
    J = InStr(1, Grupo, "|")
    If J > 0 Then
        Cad = Mid(Grupo, 1, J - 1)
        Grupo = Mid(Grupo, J + 1)
        J = 1
    End If
    codArtic(Contador) = Cad
    
    'Cantidad
    J = InStr(1, Grupo, "|")
    If J > 0 Then
        Cad = Mid(Grupo, 1, J - 1)
        Grupo = Mid(Grupo, J + 1)
    Else
        Cad = Grupo
        Grupo = ""
    End If
    cantidad(Contador) = Cad
End Sub





Private Sub imgCheck_Click(Index As Integer)
Dim b As Boolean
    Select Case Index
       Case 4, 5
            'En el listview7
            b = (Index = 5)
            For TotalArray = 1 To ListView7.ListItems.Count
                ListView7.ListItems(TotalArray).Checked = b
                If (TotalArray Mod 50) = 0 Then DoEvents
            Next TotalArray
    End Select
End Sub



Private Sub OptCompXClien_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub OptCompXDpto_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub OptCompXMant_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub


Private Sub KEYpress(KeyAscii As Integer)
Dim cerrar As Boolean

    KEYpressGnral KeyAscii, 2, cerrar
    If cerrar Then Unload Me
End Sub




Private Sub txtMante_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Public Function ObtenerSQLcomponentes(cadWHERE As String) As String
'Obtiene la consulta SQL que selecciona los articulos con nº de serie
'agrupados por tipo de articulo
Dim sql As String

    sql = "Select distinct sserie.codtipar, nomtipar, count(numserie) as cantidad "
    sql = sql & "FROM sserie INNER JOIN stipar ON sserie.codtipar=stipar.codtipar "
    sql = sql & cadWHERE
    sql = sql & " GROUP by codtipar "
    
    ObtenerSQLcomponentes = sql
End Function




Private Sub CargarListaFields(DadoProducto As Boolean)
Dim sql As String
Dim Rs As ADODB.Recordset
Dim It As ListItem

    Select Case Label5.Caption
        Case "Cuentas Contables"
            sql = "select cuentas.codmacta as codigo, cuentas.nommacta as descripcion from cuentas "
    End Select

'    ' viene de un rango de clases
'    Sql = "select variedades.codvarie, variedades.nomvarie, variedades.codclase, clases.nomclase from variedades, clases "
'    Sql = Sql & " where variedades.codclase = clases.codclase "
'
    If cadWHERE <> "" Then sql = sql & " where (1=1) " & cadWHERE
    
    Set Rs = New ADODB.Recordset
    Rs.Open sql, ConnContaFac, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    ListView7.ColumnHeaders.Clear
    ListView7.ColumnHeaders.Add , , "Cuenta", 2000.0631
    ListView7.ColumnHeaders.Add , , "Nombre Cuenta", 4101.0396
    TotalArray = 0
    While Not Rs.EOF
        Set It = ListView7.ListItems.Add
        Select Case Label5.Caption
            Case "Cuentas Contables"
                It.Text = DBLet(Rs!Codigo, "N")
        End Select
        It.SubItems(1) = DBLet(Rs!Descripcion, "T")
        
        It.Checked = True
        
        Rs.MoveNext
        TotalArray = TotalArray + 1
        If TotalArray > 300 Then
            TotalArray = 0
            DoEvents
        End If
    Wend
    Rs.Close
    
End Sub

