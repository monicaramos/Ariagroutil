VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMensaje 
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   Caption         =   "Mensajes"
   ClientHeight    =   5865
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9435
   LinkTopic       =   "Form1"
   ScaleHeight     =   5865
   ScaleWidth      =   9435
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FrameFrasPteContabilizar 
      Height          =   5790
      Left            =   0
      TabIndex        =   20
      Top             =   0
      Width           =   9345
      Begin VB.ComboBox Combo1 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   0
         ItemData        =   "frmMensaje.frx":0000
         Left            =   270
         List            =   "frmMensaje.frx":000A
         Style           =   2  'Dropdown List
         TabIndex        =   24
         Tag             =   "Tipo de cliente|N|N|0|2|ssocio|tipsocio|||"
         Top             =   240
         Width           =   2055
      End
      Begin VB.CommandButton cmdCerrarFras 
         Caption         =   "Continuar"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   5
         Left            =   7950
         TabIndex        =   21
         Top             =   5280
         Width           =   1215
      End
      Begin MSComctlLib.ListView ListView22 
         Height          =   4545
         Left            =   240
         TabIndex        =   22
         Top             =   630
         Width           =   8940
         _ExtentX        =   15769
         _ExtentY        =   8017
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
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
      Begin VB.Label Label24 
         Alignment       =   1  'Right Justify
         Caption         =   "Facturas Pendientes de Contabilizar"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00972E0B&
         Height          =   375
         Left            =   780
         TabIndex        =   23
         Top             =   300
         Width           =   8355
      End
   End
   Begin VB.Frame frameAcercaDE 
      BorderStyle     =   0  'None
      ForeColor       =   &H00C00000&
      Height          =   4545
      Left            =   90
      TabIndex        =   9
      Top             =   240
      Width           =   5385
      Begin VB.CommandButton cmdAceptar 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   3840
         TabIndex        =   10
         Top             =   3900
         Width           =   1035
      End
      Begin VB.Label Label27 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fax: 96 342 09 38"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   195
         Left            =   3180
         TabIndex        =   17
         Top             =   3540
         Width           =   1560
      End
      Begin VB.Label Label26 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tfno: 902 88 88 78"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   195
         Left            =   780
         TabIndex        =   16
         Top             =   3540
         Width           =   1650
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "AriagroUtil"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   36
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   915
         Left            =   240
         TabIndex        =   15
         Top             =   360
         Width           =   4335
      End
      Begin VB.Label Label25 
         BackStyle       =   0  'Transparent
         Caption         =   "4"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   81.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   1725
         Left            =   3780
         TabIndex        =   14
         Top             =   60
         Width           =   1350
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "46007 - VALENCIA"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   195
         Left            =   3240
         TabIndex        =   13
         Top             =   3120
         Width           =   1620
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "C/ Uruguay, 11 despacho 101"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   195
         Left            =   300
         TabIndex        =   12
         Top             =   3120
         Width           =   2580
      End
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Label13"
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
         Left            =   0
         TabIndex        =   11
         Top             =   1200
         Width           =   3795
      End
      Begin VB.Image Image2 
         Height          =   540
         Left            =   1740
         Top             =   2460
         Width           =   2880
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00B48246&
         BorderWidth     =   5
         FillColor       =   &H80000005&
         FillStyle       =   0  'Solid
         Height          =   4425
         Left            =   90
         Top             =   60
         Width           =   5250
      End
      Begin VB.Image Image1 
         Height          =   4395
         Left            =   -30
         Stretch         =   -1  'True
         Top             =   30
         Width           =   5355
      End
   End
   Begin VB.Frame FrameErrores 
      Height          =   5505
      Left            =   60
      TabIndex        =   4
      Top             =   90
      Width           =   8835
      Begin VB.CommandButton CmdSalir 
         Caption         =   "&Salir"
         Height          =   375
         Left            =   6930
         TabIndex        =   5
         Top             =   4830
         Width           =   1035
      End
      Begin MSComctlLib.ListView ListView2 
         Height          =   4155
         Left            =   210
         TabIndex        =   6
         Top             =   540
         Width           =   7785
         _ExtentX        =   13732
         _ExtentY        =   7329
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin VB.Label Label2 
         Caption         =   "Errores de Comprobación"
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
         Height          =   255
         Left            =   270
         TabIndex        =   8
         Top             =   210
         Width           =   3585
      End
      Begin VB.Label Label1 
         Caption         =   "Label2"
         Height          =   345
         Index           =   2
         Left            =   450
         TabIndex        =   7
         Top             =   1470
         Width           =   3555
      End
   End
   Begin VB.Frame FrameCobrosPtes 
      Height          =   5430
      Left            =   45
      TabIndex        =   0
      Top             =   90
      Width           =   8835
      Begin VB.CommandButton CmdCancelarCobros 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   6975
         TabIndex        =   2
         Top             =   4860
         Width           =   1035
      End
      Begin VB.CommandButton CmdAceptarCobros 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   5805
         TabIndex        =   1
         Top             =   4860
         Width           =   1035
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   4095
         Left            =   360
         TabIndex        =   18
         Top             =   630
         Width           =   7650
         _ExtentX        =   13494
         _ExtentY        =   7223
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Appearance      =   1
         NumItems        =   0
      End
      Begin VB.Label Label3 
         Caption         =   "Errores de Comprobación"
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
         Height          =   255
         Left            =   360
         TabIndex        =   19
         Top             =   225
         Width           =   3585
      End
      Begin VB.Label Label1 
         Caption         =   "Label2"
         Height          =   345
         Index           =   1
         Left            =   450
         TabIndex        =   3
         Top             =   1470
         Width           =   3555
      End
   End
End
Attribute VB_Name = "frmMensaje"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public OpcionMensaje As Byte
Public Cadena As String

'variables que se pasan con valor al llamar al formulario de zoom desde otro formulario


Public pTitulo As String


Dim vAnt As Integer

Private Sub cmdAceptar_Click()
    Unload Me
End Sub

Private Sub cmdCancelar_Click()
    'salimos y no hacemos nada
    Unload Me
End Sub

Private Sub CmdAceptarCobros_Click()
     Unload Me
End Sub

Private Sub CmdCancelarCobros_Click()
    Unload Me
End Sub

Private Sub cmdCerrarFras_Click(Index As Integer)
    Unload Me
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Combo1_Click(Index As Integer)
   Select Case Index
        Case 0
            If vAnt <> Combo1(0).ListIndex Then CargarFacturasPendientesContabilizar
            vAnt = Combo1(0).ListIndex
    End Select
End Sub

Private Sub Combo1_GotFocus(Index As Integer)
    Select Case Index
        Case 0
            vAnt = Combo1(0).ListIndex
    End Select
End Sub

Private Sub Form_Activate()
    PonerFocoBtn Me.cmdAceptar
End Sub


Private Sub Form_Load()
    Me.Shape1.Width = Me.Width - 30
    Me.Shape1.Height = Me.Height - 30

    'obtener el campo correspondiente y mostrarlo en el text
    
    Label1(1).Caption = pTitulo

    If OpcionMensaje <= 3 Then ' Errores al hacer comprobaciones
        PonerFrameCobrosPtesVisible True, 1000, 2000
        CargarListaErrComprobacion
        Me.Caption = "Errores de Comprobacion: "
        PonerFocoBtn Me.cmdSalir
    End If
    

    If OpcionMensaje = 10 Then  'Errores al contabilizar facturas
        PonerFrameCobrosPtesVisible True, 1000, 2000
        CargarListaErrContab
        Me.Caption = "Facturas NO contabilizadas: "
        PonerFocoBtn Me.CmdAceptarCobros
    End If
    
    If OpcionMensaje = 11 Then  'Errores al contabilizar facturas socios
        PonerFrameCobrosPtesVisible True, 1000, 2000
        CargarListaErrContabFacSoc
        Me.Caption = "Facturas NO contabilizadas: "
        Me.Label3.Caption = "Facturas NO contabilizadas: "
        Me.CmdAceptarCobros.visible = False
        PonerFocoBtn Me.CmdCancelarCobros '.CmdAceptarCobros
    End If

    If OpcionMensaje = 6 Then
        PonerFrameCobrosPtesVisible True, 1000, 2000
        CargaImagen
'        Me.Caption = "Acerca de ....."
'        w = Me.frameAcercaDE.Width
'        h = Me.frameAcercaDE.Height
        Me.frameAcercaDE.visible = True
        Label13.Caption = "Versión:  " & App.Major & "." & App.Minor & "." & App.Revision & " "
    End If


    If OpcionMensaje = 12 Then
        PonerFrameCobrosPtesVisible True, 1000, 2000
        CargarFacturasPendientesContabilizar
        PonerFocoBtn Me.CmdCancelarCobros '.CmdAceptarCobros
    
        CargarCombo
        
        Combo1(0).ListIndex = 0
    
    End If

End Sub

Private Sub CargarListaErrContab()
'Muestra la lista Detallada de Facturas que dieron error al contabilizar
'en un ListView
Dim Rs As ADODB.Recordset
Dim ItmX As ListItem
Dim SQL As String

    On Error GoTo ECargarList

    SQL = " SELECT  * "
    SQL = SQL & " FROM tmperrfac "
    
    Set Rs = New ADODB.Recordset
    Rs.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not Rs.EOF Then
        ListView1.Height = 4500
        ListView1.Width = 7400
        ListView1.Left = 500
        ListView1.Top = 500

        'Los encabezados
        ListView1.ColumnHeaders.Clear

        If Rs.Fields(0).Name = "codprove" Then
            'Facturas de Compra
             ListView1.ColumnHeaders.Add , , "Prove.", 700
        Else 'Facturas de Venta
            ListView1.ColumnHeaders.Add , , "Tipo", 600
        End If
        ListView1.ColumnHeaders.Add , , "Factura", 1000, 1
        ListView1.ColumnHeaders.Add , , "Fecha", 1100, 1
        ListView1.ColumnHeaders.Add , , "Error", 4620
    
        While Not Rs.EOF
            Set ItmX = ListView1.ListItems.Add
            'El primer campo será codtipom si llamamos desde Ventas
            ' y será codprove si llamamos desde Compras
            ItmX.Text = Rs.Fields(0).Value
            ItmX.SubItems(1) = Format(Rs!numfactu, "0000000")
            ItmX.SubItems(2) = Rs!fecfactu
            ItmX.SubItems(3) = Rs!Error
            Rs.MoveNext
        Wend
    End If
    Rs.Close
    Set Rs = Nothing

ECargarList:
    If Err.Number <> 0 Then
        Err.Clear
    End If
End Sub

Private Sub PonerFrameCobrosPtesVisible(visible As Boolean, ByRef h As Integer, ByRef w As Integer)
'Pone el Frame de Cobros Pendientes Visible y Ajustado al Formulario, y visualiza los controles
'necesario para el Informe

    h = 4600
        
    Select Case OpcionMensaje
        Case 0, 1, 2, 3
            h = 6000
            w = 9200
'            Me.Label1(0).Top = 4800
'            Me.Label1(0).Left = 3400
            Me.cmdSalir.Caption = "&Salir"
            PonerFrameVisible Me.FrameErrores, visible, h, w
            Me.frameAcercaDE.visible = False
            Me.FrameCobrosPtes.visible = False
            Me.FrameFrasPteContabilizar.visible = False
        
        Case 10  'Errores al contabilizar facturas
            h = 6000
            w = 8400
            Me.CmdAceptarCobros.Top = 5300
            Me.CmdAceptarCobros.Left = 4900
            Me.FrameFrasPteContabilizar.visible = False
    
        Case 11  'Errores al contabilizar facturas socios
            h = 6000
            w = 9200
'            Me.CmdAceptarCobros.Top = 5300
'            Me.CmdAceptarCobros.Left = 4900
            PonerFrameVisible Me.FrameCobrosPtes, visible, h, w
            Me.frameAcercaDE.visible = False
            Me.FrameErrores.visible = False
            Me.FrameFrasPteContabilizar.visible = False
        
        Case 12
            h = 6000
            w = 9500
            PonerFrameVisible Me.FrameFrasPteContabilizar, visible, h, w
            Me.frameAcercaDE.visible = False
            Me.FrameErrores.visible = False
            Me.FrameCobrosPtes.visible = False
    
    
        Case 6 ' Acerca de
            h = 4485
            w = 5415
            Me.Width = w
            Me.Height = h
            Me.Shape1.Width = w
            Me.Shape1.Height = h
            Me.Shape1.Top = 0
            Me.Shape1.Left = 0
            Me.frameAcercaDE.visible = True
            Me.frameAcercaDE.Left = 5
            Me.frameAcercaDE.Top = 5
            Me.frameAcercaDE.Width = w - 5
            Me.frameAcercaDE.Height = h - 5
            Me.FrameCobrosPtes.visible = False
            Me.FrameErrores.visible = False
            Me.FrameFrasPteContabilizar.visible = False
'            PonerFrameVisible Me.frameAcercaDE, visible, h - 20, w - 20

            Exit Sub
    End Select
            
    
End Sub


Private Sub CargarListaErrComprobacion()
'Muestra la lista Detallada de Facturas que dieron error al contabilizar
'en un ListView
Dim Rs As ADODB.Recordset
Dim ItmX As ListItem
Dim SQL As String

    On Error GoTo ECargarListErrComprobacion

    SQL = " SELECT  * "
    SQL = SQL & " FROM tmperrcomprob "
    
    Set Rs = New ADODB.Recordset
    Rs.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not Rs.EOF Then
        'Los encabezados
        ListView2.ColumnHeaders.Clear

        Select Case OpcionMensaje
            Case 0
                ListView2.ColumnHeaders.Add , , "Error en Centros de Coste", 6000
            Case 1
                ListView2.ColumnHeaders.Add , , "Error en letra de serie", 6000
            Case 2
                ListView2.ColumnHeaders.Add , , "Error en cuentas contables", 6000
            Case 3
                ListView2.ColumnHeaders.Add , , "Error en tipos de iva", 6000
        
        End Select
    
        While Not Rs.EOF
            Set ItmX = ListView2.ListItems.Add
            'El primer campo será codtipom si llamamos desde Ventas
            ' y será codprove si llamamos desde Compras
            ItmX.Text = Rs.Fields(0).Value
            Rs.MoveNext
        Wend
    End If
    Rs.Close
    Set Rs = Nothing

ECargarListErrComprobacion:
    If Err.Number <> 0 Then
        Err.Clear
    End If
End Sub

Private Sub CargaImagen()
On Error Resume Next
     Image2.Picture = LoadPicture(App.path & "\logo.jpg")
    Err.Clear
End Sub

Private Sub CargarListaErrContabFacSoc()
'Muestra la lista Detallada de Facturas que dieron error al contabilizar
'en un ListView
Dim Rs As ADODB.Recordset
Dim ItmX As ListItem
Dim SQL As String

    On Error GoTo ECargarList

    SQL = " SELECT  * "
    SQL = SQL & " FROM tmperrfac "
    
    Set Rs = New ADODB.Recordset
    Rs.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not Rs.EOF Then

        'Los encabezados
        ListView1.ColumnHeaders.Clear
        ListView1.ColumnHeaders.Add , , "Factura", 1200
        ListView1.ColumnHeaders.Add , , "Fecha", 1100
        ListView1.ColumnHeaders.Add , , "Error", 9620
    
        While Not Rs.EOF
            Set ItmX = ListView1.ListItems.Add
            'El primer campo será codtipom si llamamos desde Ventas
            ' y será codprove si llamamos desde Compras
            ItmX.Text = Rs!numfactu
            ItmX.SubItems(1) = Rs!fecfactu
            ItmX.SubItems(2) = Rs!Error
            Rs.MoveNext
        Wend
    End If
    Rs.Close
    Set Rs = Nothing

ECargarList:
    If Err.Number <> 0 Then
        Err.Clear
    End If
End Sub



Private Sub CargarFacturasPendientesContabilizar()
Dim SQL As String
Dim sql2 As String
Dim Rs As ADODB.Recordset
Dim IT As ListItem
Dim TotalArray As Long
Dim BdConta As Integer


    SQL = Cadena
    Select Case Combo1(0).ListIndex
        Case 0
        
        Case 1
            SQL = SQL & " and codigo1 = 0 "
        Case 2
            SQL = SQL & " and codigo1 = 1 "
        Case 3
            SQL = SQL & " and codigo1 = 2 "
        Case 4
            SQL = SQL & " and codigo1 = 3 "
    End Select
    SQL = SQL & " order by fecha1 "
     
     
    Set Rs = New ADODB.Recordset
    Rs.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    ListView22.ColumnHeaders.Clear

    ListView22.ColumnHeaders.Add , , "Tipo Factura", 2600
    ListView22.ColumnHeaders.Add , , "Fecha", 1600
    ListView22.ColumnHeaders.Add , , "Factura", 1600, 0
    ListView22.ColumnHeaders.Add , , "Importe", 1800, 1
    
    ListView22.ListItems.Clear
    
    ListView22.SmallIcons = frmPpal.imgListPpal
    
    
    TotalArray = 0
    While Not Rs.EOF
        Set IT = ListView22.ListItems.Add
            
        'It.Tag = DevNombreSQL(RS!codCampo)
        IT.Text = DBLet(Rs!nombre1, "T")
        IT.SubItems(1) = DBLet(Rs!Fecha1, "F")
        IT.SubItems(2) = DBLet(Rs!nombre2, "T")
        IT.SubItems(3) = Format(DBLet(Rs!importe1, "N"), "###,###,##0.00")
        
        Select Case DBLet(Rs!codigo1, "N")
            Case 0 ' facturas varias
                sql2 = "select numconta from seccion "
                BdConta = DevuelveValor(sql2)
                If AbrirConexionContaFac(vParamAplic.UsuarioContaFac, vParamAplic.PasswordContaFac, BdConta) Then
                    Set vEmpresaFac = New CempresaFac
                    If vEmpresaFac.LeerNiveles Then
                        If vEmpresaFac.TieneSII Then
                            If DBLet(Rs!Fecha1, "F") < DateAdd("d", vEmpresaFac.SIIDiasAviso * (-1), Now) Then
                                IT.ForeColor = vbRed
                                IT.ListSubItems.Item(1).ForeColor = vbRed
                                IT.ListSubItems.Item(2).ForeColor = vbRed
                                IT.ListSubItems.Item(3).ForeColor = vbRed
                            End If
                        End If
                    End If
                    Set vEmpresaFac = Nothing
                 End If
                 CerrarConexionContaFac
                 IT.SmallIcon = 21
                 
            Case 1 ' facturas de socio
                If vEmpresaFacSoc.TieneSII Then
                    If DBLet(Rs!Fecha1, "F") < DateAdd("d", vEmpresaFacSoc.SIIDiasAviso * (-1), Now) Then
                        IT.ForeColor = vbRed
                        IT.ListSubItems.Item(1).ForeColor = vbRed
                        IT.ListSubItems.Item(2).ForeColor = vbRed
                        IT.ListSubItems.Item(3).ForeColor = vbRed
                    End If
                End If
                IT.SmallIcon = 18
            
            Case 2 ' facturas de gasolinera
                If vEmpresaGas.TieneSII Then
                    If DBLet(Rs!Fecha1, "F") < DateAdd("d", vEmpresaGas.SIIDiasAviso * (-1), Now) Then
                        IT.ForeColor = vbRed
                        IT.ListSubItems.Item(1).ForeColor = vbRed
                        IT.ListSubItems.Item(2).ForeColor = vbRed
                        IT.ListSubItems.Item(3).ForeColor = vbRed
                    End If
                End If
                IT.SmallIcon = 16

            Case 3 ' facturas de telefonia
                If vEmpresaTel.TieneSII Then
                    If DBLet(Rs!Fecha1, "F") < DateAdd("d", vEmpresaTel.SIIDiasAviso * (-1), Now) Then
                        IT.ForeColor = vbRed
                        IT.ListSubItems.Item(1).ForeColor = vbRed
                        IT.ListSubItems.Item(2).ForeColor = vbRed
                        IT.ListSubItems.Item(3).ForeColor = vbRed
                    End If
                End If
                IT.SmallIcon = 14
        
        End Select
        
        ListView22.Refresh
        
        Rs.MoveNext
        TotalArray = TotalArray + 1
        If TotalArray > 300 Then
            TotalArray = 0
            DoEvents
        End If
    Wend
    Rs.Close
    
End Sub


Private Sub CargarCombo()
    ' *** neteje els combos, els pose valor i seleccione el valor per defecte ***
    Combo1(0).Clear
    
    Combo1(0).AddItem "Todas"
    Combo1(0).ItemData(Combo1(0).NewIndex) = 0
    
    If vParamAplic.FacturasVarias Then
        Combo1(0).AddItem "Fras.Varias"
        Combo1(0).ItemData(Combo1(0).NewIndex) = 1
    End If
    If vParamAplic.FactSocios Then
        Combo1(0).AddItem "Fras.Socio"
        Combo1(0).ItemData(Combo1(0).NewIndex) = 2
    End If
    If vParamAplic.Gasolinera Then
        Combo1(0).AddItem "Fras.Gasolinera"
        Combo1(0).ItemData(Combo1(0).NewIndex) = 3
    End If
    If vParamAplic.Telefonia Then
        Combo1(0).AddItem "Fras.Telefonía"
        Combo1(0).ItemData(Combo1(0).NewIndex) = 4
    End If
End Sub


