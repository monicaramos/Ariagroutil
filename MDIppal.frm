VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm MDIppal 
   BackColor       =   &H8000000C&
   Caption         =   "AriagroUtil"
   ClientHeight    =   7860
   ClientLeft      =   225
   ClientTop       =   1125
   ClientWidth     =   11160
   Icon            =   "MDIppal.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11160
      _ExtentX        =   19685
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      Style           =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   34
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Avnics"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Informe Avnics"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Renovación Avnics"
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Calculo Intereses"
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Cancelación Avnics"
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Movimientos Avnics"
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Grabación 193"
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Ayuda Modelo 123"
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Líneas"
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Pólizas"
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Traspaso Agroweb"
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Paso a Tesoreria Seguros"
         EndProperty
         BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Asiento Contable Seguros"
         EndProperty
         BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button19 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Facturas Telefonía"
         EndProperty
         BeginProperty Button20 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Traspaso Datos Telefonía"
         EndProperty
         BeginProperty Button21 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Contabilización Facturas Telefonía"
         EndProperty
         BeginProperty Button22 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button23 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Secciones"
         EndProperty
         BeginProperty Button24 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Conceptos de Facturas"
         EndProperty
         BeginProperty Button25 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Facturas Varias"
         EndProperty
         BeginProperty Button26 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Integración Contable"
         EndProperty
         BeginProperty Button27 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button28 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Facturas Gasolinera"
         EndProperty
         BeginProperty Button29 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Trapaso Gasolinera"
         EndProperty
         BeginProperty Button30 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Integración Contable"
         EndProperty
         BeginProperty Button31 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button32 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Variedades"
         EndProperty
         BeginProperty Button33 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Facturas Socios"
         EndProperty
         BeginProperty Button34 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Contabilización Facturas Socios"
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin VB.Menu mnParametros 
      Caption         =   "&Datos Básicos"
      Index           =   1
      Begin VB.Menu mnP_Generales 
         Caption         =   "&Datos de Empresa"
         Index           =   1
      End
      Begin VB.Menu mnP_Generales 
         Caption         =   "&Parámetros"
         Index           =   2
      End
      Begin VB.Menu mnP_Generales 
         Caption         =   "&Tipos de Documentos"
         Index           =   3
      End
      Begin VB.Menu mnP_Generales 
         Caption         =   "&Usuarios"
         Index           =   4
      End
      Begin VB.Menu mnP_Generales 
         Caption         =   "-"
         Index           =   5
      End
      Begin VB.Menu mnP_Generales 
         Caption         =   "&Salir"
         Index           =   6
      End
   End
   Begin VB.Menu mnAvnics 
      Caption         =   "&Gestión Avnics"
      Begin VB.Menu mnG_Avnics 
         Caption         =   "Avnics"
         Index           =   1
      End
      Begin VB.Menu mnG_Avnics 
         Caption         =   "&Informe Avnics"
         Index           =   2
      End
      Begin VB.Menu mnG_Avnics 
         Caption         =   "-"
         Index           =   3
      End
      Begin VB.Menu mnG_Avnics 
         Caption         =   "&Renovacion Avnics"
         Index           =   4
      End
      Begin VB.Menu mnG_Avnics 
         Caption         =   "-"
         Index           =   5
      End
      Begin VB.Menu mnG_Avnics 
         Caption         =   "&Cálculo de Intereses"
         Index           =   6
      End
      Begin VB.Menu mnG_Avnics 
         Caption         =   "Conta&bilización Intereses"
         Index           =   7
      End
      Begin VB.Menu mnG_Avnics 
         Caption         =   "-"
         Index           =   8
      End
      Begin VB.Menu mnG_Avnics 
         Caption         =   "Cancelación &Avnics"
         Index           =   9
      End
      Begin VB.Menu mnG_Avnics 
         Caption         =   "&Movimientos Avnics"
         Index           =   10
      End
      Begin VB.Menu mnG_Avnics 
         Caption         =   "-"
         Index           =   11
      End
      Begin VB.Menu mnG_Avnics 
         Caption         =   "Ayuda Modelo 123"
         Index           =   12
      End
      Begin VB.Menu mnG_Avnics 
         Caption         =   "Grabacion Modelo 193"
         Index           =   13
      End
   End
   Begin VB.Menu mnSeguros 
      Caption         =   "&Seguros"
      Begin VB.Menu mnS_Seguros 
         Caption         =   "Importar Datos Agroweb"
         Index           =   1
      End
      Begin VB.Menu mnS_Seguros 
         Caption         =   "Líneas "
         Index           =   2
      End
      Begin VB.Menu mnS_Seguros 
         Caption         =   "Pólizas"
         Index           =   3
      End
      Begin VB.Menu mnS_Seguros 
         Caption         =   "Integración Contable"
         Index           =   4
         Begin VB.Menu mnS_Integ 
            Caption         =   "&Asiento Contable"
            Index           =   1
         End
         Begin VB.Menu mnS_Integ 
            Caption         =   "&Cobros Tesoreria"
            Index           =   2
         End
      End
   End
   Begin VB.Menu mnTelefonia 
      Caption         =   "&Telefonia"
      Begin VB.Menu mnT_Telefonia 
         Caption         =   "Importar Datos Fichero"
         Index           =   1
      End
      Begin VB.Menu mnT_Telefonia 
         Caption         =   "Facturas Telefónicas"
         Index           =   2
      End
      Begin VB.Menu mnT_Telefonia 
         Caption         =   "Contabilización de Facturas"
         Index           =   3
      End
   End
   Begin VB.Menu mnFacturasVarias 
      Caption         =   "Facturas &Varias"
      Begin VB.Menu mnT_FacVarias 
         Caption         =   "&Secciones"
         Index           =   1
      End
      Begin VB.Menu mnT_FacVarias 
         Caption         =   "&Conceptos de Facturación"
         Index           =   2
      End
      Begin VB.Menu mnT_FacVarias 
         Caption         =   "&Entrada de Facturas"
         Index           =   3
      End
      Begin VB.Menu mnT_FacVarias 
         Caption         =   "&Reimpresión de Facturas"
         Index           =   4
      End
      Begin VB.Menu mnT_FacVarias 
         Caption         =   "Envio de &Facturas por email"
         Index           =   5
      End
      Begin VB.Menu mnT_FacVarias 
         Caption         =   "&Integración Contable"
         Index           =   6
      End
      Begin VB.Menu mnT_FacVarias 
         Caption         =   "Integración &Aridoc"
         Index           =   7
      End
   End
   Begin VB.Menu mnAportaciones 
      Caption         =   "&Aportaciones"
      Enabled         =   0   'False
      Visible         =   0   'False
      Begin VB.Menu mnT_Aportaciones 
         Caption         =   "Tipos de Aportación"
         Enabled         =   0   'False
         Index           =   1
      End
   End
   Begin VB.Menu mnGasolinera 
      Caption         =   "&Gasolinera"
      Begin VB.Menu mnT_Gasolinera 
         Caption         =   "&Traspaso Facturas"
         Index           =   1
      End
      Begin VB.Menu mnT_Gasolinera 
         Caption         =   "&Histórico Facturas"
         Index           =   2
      End
      Begin VB.Menu mnT_Gasolinera 
         Caption         =   "&Ventas por Socio"
         Index           =   3
      End
      Begin VB.Menu mnT_Gasolinera 
         Caption         =   "&Integración Contable"
         Index           =   4
      End
   End
   Begin VB.Menu mnFacturasSocios 
      Caption         =   "&Facturas Socios"
      Begin VB.Menu mnT_FacSocios 
         Caption         =   "&Variedades"
         Index           =   1
      End
      Begin VB.Menu mnT_FacSocios 
         Caption         =   "&Entrada Facturas"
         Index           =   2
      End
      Begin VB.Menu mnT_FacSocios 
         Caption         =   "&Reimpresión de Facturas"
         Index           =   3
      End
      Begin VB.Menu mnT_FacSocios 
         Caption         =   "&Contabilización"
         Index           =   4
      End
      Begin VB.Menu mnT_FacSocios 
         Caption         =   "-"
         Index           =   5
      End
      Begin VB.Menu mnT_FacSocios 
         Caption         =   "&Listado de Retenciones"
         Index           =   6
      End
      Begin VB.Menu mnT_FacSocios 
         Caption         =   "&Grabación Modelo 190"
         Index           =   7
      End
      Begin VB.Menu mnT_FacSocios 
         Caption         =   "-"
         Index           =   8
      End
      Begin VB.Menu mnT_FacSocios 
         Caption         =   "Traspaso Liquidación"
         Index           =   9
      End
      Begin VB.Menu mnT_FacSocios 
         Caption         =   "Traspaso Facturas Aceite"
         Index           =   10
      End
      Begin VB.Menu mnT_FacSocios 
         Caption         =   "Traspaso Subv.FEDEPROL"
         Index           =   11
      End
   End
   Begin VB.Menu mnCoarval 
      Caption         =   "&Coarval y Varios"
      Begin VB.Menu mnT_Coarval 
         Caption         =   "&Importar Datos Ficheros"
         Index           =   1
      End
      Begin VB.Menu mnT_Coarval 
         Caption         =   "&Histórico Facturas"
         Index           =   2
      End
      Begin VB.Menu mnT_Coarval 
         Caption         =   "&Contabilización de Facturas"
         Index           =   3
      End
   End
   Begin VB.Menu mnUtil 
      Caption         =   "&Utilidades"
      WindowList      =   -1  'True
      Begin VB.Menu mnE_Util 
         Caption         =   "Revisión de caracteres en Multibase"
         Index           =   1
      End
      Begin VB.Menu mnE_Util 
         Caption         =   "-"
         Index           =   2
      End
      Begin VB.Menu mnE_Util 
         Caption         =   "&Copia de Seguridad local"
         Index           =   3
      End
   End
   Begin VB.Menu mnSoporte 
      Caption         =   "&Soporte"
      Begin VB.Menu mnE_Soporte1 
         Caption         =   "&Web Soporte"
      End
      Begin VB.Menu mnp_Barra2 
         Caption         =   "-"
      End
      Begin VB.Menu mnE_Soporte2 
         Caption         =   "&Acerca de"
      End
   End
End
Attribute VB_Name = "MDIppal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private PrimeraVez As Boolean
Dim TieneEditorDeMenus As Boolean

Public Sub GetIconsFromLibrary(ByVal sLibraryFilePath As String, ByVal op As Integer, ByVal tam As Integer)
    Dim i As Integer
    Dim tRes As ResType, iCount As Integer
        
    opcio = op
    tamany = tam
    ghmodule = LoadLibraryEx(sLibraryFilePath, 0, DONT_RESOLVE_DLL_REFERENCES)

    If ghmodule = 0 Then
        MsgBox "Invalid library file.", vbCritical
        Exit Sub
    End If
        
    For tRes = RT_FIRST To RT_LAST
        DoEvents
        EnumResourceNames ghmodule, tRes, AddressOf EnumResNameProc, 0
    Next
    FreeLibrary ghmodule
             
End Sub

Private Sub MDIForm_Activate()
    If PrimeraVez Then
        PrimeraVez = False
    End If
End Sub

Private Sub MDIForm_Load()
Dim cad As String
Dim i As Integer
    PrimeraVez = True
    CargarImagen
    PonerDatosFormulario

    If vEmpresa Is Nothing Then
        Caption = "ARIAGROUTIL" & " ver. " & App.Major & "." & App.Minor & "." & App.Revision & "   -  " & " FALTA CONFIGURAR"
    Else
        Caption = "ARIAGROUTIL" & " ver. " & App.Major & "." & App.Minor & "." & App.Revision & "   -  Empresa: " & vEmpresa.nomEmpre & cad & _
                  "   -  Usuario: " & vSesion.Nombre
    End If

    ' *** per als iconos XP ***
    GetIconsFromLibrary App.path & "\iconos.dll", 1, 32
    
    GetIconsFromLibrary App.path & "\iconos.dll", 1, 24
    GetIconsFromLibrary App.path & "\iconos_BN.dll", 2, 24
    GetIconsFromLibrary App.path & "\iconos_OM.dll", 3, 24
    
    GetIconsFromLibrary App.path & "\iconosAriagroutil.dll", 4, 24
    
    
  
    'CARGAR LA TOOLBAR DEL FORM PRINCIPAL
    With Me.Toolbar1
'        .HotImageList = frmPpal.imgListComun_OM
'        .DisabledImageList = frmPpal.imgListComun_BN
        .ImageList = frmPpal.imgListPpal

        .Buttons(1).Image = 1   'Avnics
        .Buttons(2).Image = 4  'Informe Avnics
        'el 3 son separadors
        .Buttons(4).Image = 5   'Renovación de Avnics
        .Buttons(5).Image = 8    'Calculo de Intereses
        .Buttons(6).Image = 3    'Cancelacion Avnics
        'el 7 son separadors
        .Buttons(8).Image = 6   'Movimientos Avnics
        .Buttons(9).Image = 7   'Grabacion 193
        .Buttons(10).Image = 15  'Ayuda de modelo 123
        'el 10 es separador
        .Buttons(12).Image = 9  'Lineas de Pólizas
        .Buttons(13).Image = 12  'Pólizas
        'el 14 es separador
        .Buttons(15).Image = 10  'traspaso de Pólizas
        .Buttons(16).Image = 11  'paso a tesoreria
        .Buttons(17).Image = 13  'paso a contabilidad
        'el 18 es separador
        .Buttons(19).Image = 14  'facturas de telefonia
        .Buttons(20).Image = 10  'traspaso desde fichero
        .Buttons(21).Image = 13  'contabilizacion de facturas de telefonia
        'el 22 es separador
        .Buttons(23).Image = 19  'Facturas varias: secciones
        .Buttons(24).Image = 20  'Facturas varias: conceptos de factura
        .Buttons(25).Image = 21  'Facturas varias: entradas de factura
        .Buttons(26).Image = 13  'Facturas varias: contabilizacion
        'el 27 es separador
        .Buttons(28).Image = 16  'Facturas Gasolinera
        .Buttons(29).Image = 10  'Traspaso Gasolinera
        .Buttons(30).Image = 13  'contabilizacion de facturas de gasolinera
        'el 31 es separador
        .Buttons(32).Image = 17  'Variedades
        .Buttons(33).Image = 18  'Facturas Socios
        .Buttons(34).Image = 13  'contabilizacion de facturas de gasolinera
    End With
    
    
    For i = 1 To 10
        Me.Toolbar1.Buttons(i).visible = vParamAplic.Avnics
        Me.Toolbar1.Buttons(i).Enabled = vParamAplic.Avnics
    Next i
    For i = 11 To 17
        Me.Toolbar1.Buttons(i).visible = vParamAplic.Seguros
        Me.Toolbar1.Buttons(i).Enabled = vParamAplic.Seguros
    Next i
    For i = 19 To 21
        Me.Toolbar1.Buttons(i).visible = vParamAplic.Telefonia
        Me.Toolbar1.Buttons(i).Enabled = vParamAplic.Telefonia
    Next i
    For i = 22 To 26
        Me.Toolbar1.Buttons(i).visible = vParamAplic.FacturasVarias
        Me.Toolbar1.Buttons(i).Enabled = vParamAplic.FacturasVarias
    Next i
    For i = 28 To 30
        Me.Toolbar1.Buttons(i).visible = vParamAplic.Gasolinera
        Me.Toolbar1.Buttons(i).Enabled = vParamAplic.Gasolinera
    Next i
    For i = 31 To 34
        Me.Toolbar1.Buttons(i).visible = vParamAplic.FactSocios
        Me.Toolbar1.Buttons(i).Enabled = vParamAplic.FactSocios
    Next i
    
    GetIconsFromLibrary App.path & "\iconos.dll", 1, 16
    GetIconsFromLibrary App.path & "\iconos_BN.dll", 2, 16
    GetIconsFromLibrary App.path & "\iconos_OM.dll", 3, 16

    LeerEditorMenus
    
    PonerDatosFormulario
    
    BloqueoDeMenus
   
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
    AccionesCerrar
    End
End Sub

Private Sub mnE_Soporte1_Click()
    Screen.MousePointer = vbHourglass
    LanzaHome "websoporte"
    espera 2
    Screen.MousePointer = vbDefault
End Sub

Private Sub mnE_Util_Click(Index As Integer)
    SubmnE_Util_Click (Index)
End Sub


Private Sub mnE_Soporte2_Click()
    frmMensaje.OpcionMensaje = 6
    frmMensaje.Show vbModal
End Sub


Private Sub mnG_Avnics_Click(Index As Integer)
    SubmnG_Avnics_Click (Index)
End Sub

Private Sub mnP_Salir1_Click()
'    Unload frmPpal
'    Unload Me
    BotonSalir
End Sub

Private Sub mnP_Salir2_Click()
'    Unload frmPpal
'    Unload Me
    BotonSalir
End Sub

Private Sub BotonSalir()
    Unload frmPpal
    Unload Me
End Sub

Private Sub mnP_Generales_Click(Index As Integer)
    SubmnP_Generales_Click (Index)
End Sub

Private Sub mnS_Integ_Click(Index As Integer)
    SubmnS_IntegSeg_click (Index)
End Sub

Private Sub mnS_Seguros_Click(Index As Integer)
    SubmnG_Seguros_Click (Index)
End Sub

Private Sub mnT_Coarval_Click(Index As Integer)
    SubmnG_FactCoarval_Click (Index)
End Sub

Private Sub mnT_FacSocios_Click(Index As Integer)
    SubmnG_FactSocios_Click (Index)
End Sub

Private Sub mnT_FacVarias_Click(Index As Integer)
    SubmnG_FactVarias_Click (Index)
End Sub

Private Sub mnT_Aportaciones_Click(Index As Integer)
    SubmnG_Aportaciones_Click (Index)
End Sub

Private Sub mnT_Gasolinera_Click(Index As Integer)
    SubmnG_Gasolinera_Click (Index)
End Sub

Private Sub mnT_Telefonia_Click(Index As Integer)
    SubmnG_Telefonia_Click (Index)
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
        ' AVNICs
        Case 1 'Avnics
            SubmnG_Avnics_Click (1)
        Case 2 'Informe Avnics
            SubmnG_Avnics_Click (2)
        Case 4 'Renovacion Avnics
            SubmnG_Avnics_Click (4)
        Case 5 ' Calculo de Intereses
            SubmnG_Avnics_Click (6)
        Case 6 ' Cancelacion Avnics
            SubmnG_Avnics_Click (9)
        Case 8 'Movimientos Avnics
            SubmnG_Avnics_Click (10)
        Case 9 ' grabacion de 193
            SubmnG_Avnics_Click (13)
        Case 10 ' ayuda modelo 123
            SubmnG_Avnics_Click (12)
            
        ' SEGUROS AGRARIOS
        Case 12 'Mantenimiento de Líneas de Seguros
            SubmnG_Seguros_Click (2)
        Case 13 'Mantenimiento de Pólizas
            SubmnG_Seguros_Click (3)
        Case 15 'TRaspaso de polizas de Agroweb
            SubmnG_Seguros_Click (1)
        Case 16 'Paso a tesoreria
            SubmnS_IntegSeg_click (2)
        Case 17 'Asiento contable
            SubmnS_IntegSeg_click (1)
            
        ' TELFONIA
        Case 19 'Facturas Telefonia
            SubmnG_Telefonia_Click (2)
        Case 20 'Traspaso de Telefonia
            SubmnG_Telefonia_Click (1)
        Case 21 'Contabilizacion de Facturas de Telefonia
            SubmnG_Telefonia_Click (3)
            
        ' FACTURAS VARIAS
        Case 23 'Facturas varias: secciones
            SubmnG_FactVarias_Click (1)
        Case 24 'Facturas varias: conceptos facturación
            SubmnG_FactVarias_Click (2)
        Case 25 'Facturas varias: entrada de factura
            SubmnG_FactVarias_Click (3)
        Case 26 'Facturas varias: contabilizacion
            SubmnG_FactVarias_Click (5)
        
        ' GASOLINERA
        Case 28 'Facturas de Gasolinera
            SubmnG_Gasolinera_Click (2)
        Case 29 'Trapaso de Gasolinera
            SubmnG_Gasolinera_Click (1)
        Case 30 'Contabilizacion de facturas de gasolinera
            SubmnG_Gasolinera_Click (4)
            
        ' FACTURAS SOCIOS
        Case 32 'Variedades
            SubmnG_FactSocios_Click (1)
        Case 33 'Facturas Socios
            SubmnG_FactSocios_Click (2)
        Case 34 'Contabilizacion de facturas socios
            SubmnG_FactSocios_Click (4)
    End Select
End Sub

' ### [Monica] 05/09/2006
Private Sub PonerDatosFormulario()
Dim Config As Boolean

    Config = (vEmpresa Is Nothing) Or (vParamAplic Is Nothing)
    
    If Not Config Then HabilitarSoloPrametros_o_Empresas True

    'FijarConerrores
    CadenaDesdeOtroForm = ""

    'Poner datos visible del form
'    PonerDatosVisiblesForm
    
    'Habilitar/Deshabilitar entradas del menu segun el nivel de usuario
'    PonerMenusNivelUsuario

    'Si no hay carpeta interaciones, no habra integraciones
'    Me.mnComprobarPendientes.Enabled = vConfig.Integraciones <> ""


    'Habilitar
    If Config Then HabilitarSoloPrametros_o_Empresas False
    'Panel con el nombre de la empresa
'    If Not vEmpresa Is Nothing Then
'        Me.StatusBar1.Panels(2).Text = "Empresa:   " & vEmpresa.nomempre & "               Código: " & vEmpresa.codempre
'    Else
'        Me.StatusBar1.Panels(2).Text = "Falta configurar"
'    End If


    'Si tiene editor de menus
    If TieneEditorDeMenus Then PoneMenusDelEditor

End Sub

' ### [Monica] 05/09/2006
Private Sub HabilitarSoloPrametros_o_Empresas(Habilitar As Boolean)
Dim T As Control
Dim cad As String

    On Error Resume Next
    For Each T In Me
        cad = T.Name
        If Mid(T.Name, 1, 2) = "mn" Then
            'If LCase(Mid(T.Name, 1, 8)) <> "mn_b" Then
                T.Enabled = Habilitar
            'End If
        End If
    Next
    
    Me.Toolbar1.Enabled = Habilitar
    Me.Toolbar1.visible = Habilitar
    Me.mnParametros(1).Enabled = True
    Me.mnP_Generales(1).Enabled = True
    Me.mnP_Generales(2).Enabled = True
    Me.mnP_Generales(6).Enabled = True
    Me.mnP_Generales(17).Enabled = True
    
'    Me.mnCambioEmpresa.Enabled = True
End Sub


' ### [Monica] 07/11/2006
' añadida esta parte para la personalizacion de menus

Private Sub LeerEditorMenus()
Dim SQL As String
Dim miRsAux As ADODB.Recordset

    On Error GoTo ELeerEditorMenus
    TieneEditorDeMenus = False
    SQL = "Select count(*) from appmenus where aplicacion='Avnics'"
    Set miRsAux = New ADODB.Recordset
    miRsAux.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not miRsAux.EOF Then
        If Not IsNull(miRsAux.Fields(0)) Then
            If miRsAux.Fields(0) > 0 Then TieneEditorDeMenus = True
        End If
    End If
    miRsAux.Close
        

ELeerEditorMenus:
    Set miRsAux = Nothing
    If Err.Number <> 0 Then Err.Clear
End Sub




Private Sub PoneMenusDelEditor()
Dim T As Control
Dim SQL As String
Dim C As String
Dim miRsAux As ADODB.Recordset

    On Error GoTo ELeerEditorMenus
    
    SQL = "Select * from appmenususuario where aplicacion='Avnics' and codusu = " & Val(Right(CStr(vSesion.Codusu), 3))
    Set miRsAux = New ADODB.Recordset
    miRsAux.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    SQL = ""

    While Not miRsAux.EOF
        If Not IsNull(miRsAux.Fields(3)) Then
            SQL = SQL & miRsAux.Fields(3) & "·"
        End If
        miRsAux.MoveNext
    Wend
    miRsAux.Close
        
   
    If SQL <> "" Then
        SQL = "·" & SQL
        For Each T In Me.Controls
            If TypeOf T Is menu Then
                C = DevuelveCadenaMenu(T)
                C = "·" & C & "·"
                If InStr(1, SQL, C) > 0 Then T.visible = False
           
            End If
        Next
    End If
ELeerEditorMenus:
    Set miRsAux = Nothing
    If Err.Number <> 0 Then Err.Clear
End Sub

Private Function DevuelveCadenaMenu(ByRef T As Control) As String

On Error GoTo EDevuelveCadenaMenu
    DevuelveCadenaMenu = T.Name & "|"
    DevuelveCadenaMenu = DevuelveCadenaMenu & T.Index '& "|"   Monica:con esto no funcionaba
    Exit Function
EDevuelveCadenaMenu:
    Err.Clear
    
End Function

Private Sub LanzaHome(Opcion As String)
    Dim i As Integer
    Dim cad As String
    On Error GoTo ELanzaHome
    
    'Obtenemos la pagina web de los parametros
    CadenaDesdeOtroForm = DevuelveDesdeBD("websoporte", "sparam", "codparam", 1, "N")
    If CadenaDesdeOtroForm = "" Then
        MsgBox "Falta configurar los datos en parametros.", vbExclamation
        Exit Sub
    End If
        
    i = FreeFile
    cad = ""
    Open App.path & "\lanzaexp.dat" For Input As #i
    Line Input #i, cad
    Close #i
    
    'Lanzamos
    If cad <> "" Then Shell cad & " " & CadenaDesdeOtroForm, vbMaximizedFocus
    
ELanzaHome:
    If Err.Number <> 0 Then MuestraError Err.Number, cad & vbCrLf & Err.Description
    CadenaDesdeOtroForm = ""
End Sub

Private Sub CargarImagen()

On Error GoTo eCargarImagen
'    Me.Picture = LoadPicture(App.path & "\fondo.dat")
    Exit Sub
eCargarImagen:
    MuestraError Err.Number, "Error cargando imagen. LLame a soporte"
    End
End Sub


Private Sub BloqueoDeMenus()
Dim b As Boolean
    mnAvnics.visible = (vParamAplic.Avnics = 1)
    mnAvnics.Enabled = (vParamAplic.Avnics = 1)

    mnSeguros.visible = (vParamAplic.Seguros = 1)
    mnSeguros.Enabled = (vParamAplic.Seguros = 1)
    
    mnTelefonia.visible = (vParamAplic.Telefonia = 1)
    mnTelefonia.Enabled = (vParamAplic.Telefonia = 1)
    
    mnGasolinera.visible = (vParamAplic.Gasolinera = 1)
    mnGasolinera.Enabled = (vParamAplic.Gasolinera = 1)
    
    mnFacturasSocios.visible = (vParamAplic.FactSocios = 1)
    mnFacturasSocios.Enabled = (vParamAplic.FactSocios = 1)

    mnCoarval.visible = (vParamAplic.Coarval = 1)
    mnCoarval.Enabled = (vParamAplic.Coarval = 1)

    '[Monica]18/11/2013: Integracion a aridoc
    Me.mnT_FacVarias(7).visible = (vParamAplic.HayAridoc = 1)
    Me.mnT_FacVarias(7).Enabled = (vParamAplic.HayAridoc = 1)
    

End Sub

