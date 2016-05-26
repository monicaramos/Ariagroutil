VERSION 5.00
Begin VB.Form frmImpAridoc 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Exportar datos a AriDoc"
   ClientHeight    =   4425
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5940
   Icon            =   "frmImpAridoc.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4425
   ScaleWidth      =   5940
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      Height          =   1770
      Left            =   135
      TabIndex        =   10
      Top             =   1620
      Width           =   5640
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   6
         Left            =   2415
         Locked          =   -1  'True
         TabIndex        =   14
         Top             =   255
         Width           =   2955
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   6
         Left            =   1500
         MaxLength       =   6
         TabIndex        =   1
         Top             =   255
         Width           =   830
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   1
         Left            =   1500
         MaxLength       =   10
         TabIndex        =   3
         Tag             =   "Código Postal|T|S|||clientes|codposta|||"
         Top             =   1320
         Width           =   1050
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   0
         Left            =   1500
         MaxLength       =   10
         TabIndex        =   2
         Tag             =   "Código Postal|T|S|||clientes|codposta|||"
         Top             =   975
         Width           =   1050
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   6
         Left            =   1215
         MouseIcon       =   "frmImpAridoc.frx":000C
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar sección"
         Top             =   255
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Sección"
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
         Index           =   6
         Left            =   330
         TabIndex        =   15
         Top             =   240
         Width           =   540
      End
      Begin VB.Label Label4 
         Caption         =   "Desde"
         Height          =   195
         Index           =   15
         Left            =   570
         TabIndex        =   13
         Top             =   960
         Width           =   465
      End
      Begin VB.Label Label4 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   14
         Left            =   570
         TabIndex        =   12
         Top             =   1320
         Width           =   420
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   0
         Left            =   1230
         Picture         =   "frmImpAridoc.frx":015E
         ToolTipText     =   "Buscar fecha"
         Top             =   960
         Width           =   240
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   1
         Left            =   1230
         Picture         =   "frmImpAridoc.frx":01E9
         ToolTipText     =   "Buscar fecha"
         Top             =   1320
         Width           =   240
      End
      Begin VB.Label Label4 
         Caption         =   "Fecha Factura"
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
         Height          =   255
         Index           =   16
         Left            =   315
         TabIndex        =   11
         Top             =   630
         Width           =   1185
      End
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "Salir"
      Height          =   375
      Left            =   3960
      TabIndex        =   5
      Top             =   3960
      Width           =   1695
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "Aceptar"
      Height          =   375
      Left            =   2070
      TabIndex        =   4
      Top             =   3960
      Width           =   1695
   End
   Begin VB.Frame Frame1 
      Caption         =   "Carpeta de destino: "
      Height          =   1215
      Left            =   135
      TabIndex        =   6
      Top             =   360
      Width           =   5655
      Begin VB.TextBox txtCarp 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   1
         Left            =   1560
         TabIndex        =   8
         Top             =   600
         Width           =   3855
      End
      Begin VB.TextBox txtCarp 
         Enabled         =   0   'False
         Height          =   285
         Index           =   0
         Left            =   495
         TabIndex        =   0
         Top             =   585
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "Código"
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
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   7
         Top             =   330
         Width           =   975
      End
   End
   Begin VB.Label lblInf 
      Alignment       =   2  'Center
      Caption         =   "Información del proceso"
      Height          =   315
      Left            =   240
      TabIndex        =   9
      Top             =   3495
      Width           =   5295
   End
End
Attribute VB_Name = "frmImpAridoc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public tipo As Byte
    'Tipo:  0 Impresion de facturas varias
    
    
Dim DesdeFecha As Date
Dim Hastafecha As Date
Dim frmVis As frmVisReport
Dim impor As ArdImportador

Dim indice As Byte 'Indice del text1 donde se ponen los datos devueltos desde otros Formularios de Mtos
Dim BdConta As Integer



Private WithEvents frmC As frmCal 'calendario fechas
Attribute frmC.VB_VarHelpID = -1
Private WithEvents frmSec As frmManSecciones
Attribute frmSec.VB_VarHelpID = -1


Private Sub cmdAceptar_Click()
    If Not DatosOk() Then Exit Sub
    '-- Cargar facturas de gasolinera entre las fechas seleccionadas
    Select Case tipo
        Case 0 ' facturas de venta
            CargaFacturas DesdeFecha, Hastafecha
            MsgBox "Proceso finalizado", vbInformation
    End Select
    cmdSalir_Click
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Function DatosOk() As Boolean
    DesdeFecha = CDate(txtCodigo(0).Text)
    Hastafecha = CDate(txtCodigo(1).Text)
    If DesdeFecha > Hastafecha Then
        MsgBox "La fecha desde debe ser menor que la fecha hasta", vbInformation
        Exit Function
    End If
    If txtCarp(0) = "" Then
        MsgBox "Debe seleccionar una carpeta de importación.", vbInformation
        Exit Function
    End If
    If tipo = 0 Then
        If txtCodigo(6).Text = "" Then
            MsgBox "Debe introducir la sección de las facturas.", vbInformation
            Exit Function
        End If
    End If
    DatosOk = True
End Function


Private Sub Form_Load()
    txtCodigo(0).Text = Date
    txtCodigo(1).Text = Date
    Set impor = New ArdImportador
    
    Set ardDB = New BaseDatos
    ardDB.tipo = "MYSQL"
    ardDB.abrir "Aridoc", "root", "aritel"
    
    Frame3.Enabled = (tipo <> 2)
    Frame3.visible = (tipo <> 2)
    
   'cargar IMAGES de busqueda
    Me.imgBuscar(6).Picture = frmPpal.imgListImages16.ListImages(1).Picture
   

    Select Case tipo
        Case 0:
            Me.txtCarp(0).Text = vParamAplic.CarpetaFac
    End Select
    txtCarp_LostFocus (0)
End Sub

Private Sub frmSec_DatoSeleccionado(CadenaSeleccion As String)
Dim Cad As String
    txtCodigo(6).Text = RecuperaValor(CadenaSeleccion, 1) 'codsecci
    txtNombre(6).Text = RecuperaValor(CadenaSeleccion, 2) 'nomsecci
    
    Cad = RecuperaValor(CadenaSeleccion, 5)  'numconta
    If Cad <> "" Then BdConta = CInt(Cad)  'numero de conta

End Sub

Private Sub imgBuscar_Click(Index As Integer)
            
    Select Case Index
        Case 6
            indice = 6
            Set frmSec = New frmManSecciones
            frmSec.DatosADevolverBusqueda = "0|1|2|3|4|"
            frmSec.CodigoActual = txtCodigo(6).Text
            frmSec.Show vbModal
            Set frmSec = Nothing
    End Select

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
    PonerFoco txtCodigo(CByte(imgFec(0).Tag))
    ' ***************************
End Sub

Private Sub frmC_Selec(vFecha As Date)
 'Fecha
    txtCodigo(CByte(imgFec(0).Tag)).Text = Format(vFecha, "dd/MM/yyyy")
End Sub

Private Sub txtCarp_GotFocus(Index As Integer)
    ConseguirFoco txtCarp(Index), 3
End Sub

Private Sub txtCarp_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    KEYdown KeyCode
End Sub

Private Sub txtCarp_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub txtCarp_LostFocus(Index As Integer)
Dim Cad As String

    If Index = 0 Then
        If txtCarp(0) <> "" Then 'txtCarp(1) = impor.nombreCarpeta(CLng(txtCarp(0)))
            Cad = CargaPath(txtCarp(Index))
            txtCarp(1).Text = Mid(Cad, 2, Len(Cad))
        End If
    End If
End Sub

Private Sub txtCodigo_GotFocus(Index As Integer)
    ConseguirFoco txtCodigo(Index), 3
End Sub

Private Sub txtCodigo_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    KEYdown KeyCode
End Sub

Private Sub txtCodigo_KeyPress(Index As Integer, KeyAscii As Integer)
'14/02/2007 antes
'    KEYpress KeyAscii
    If KeyAscii = teclaBuscar Then
        Select Case Index
            Case 0: KEYFecha KeyAscii, 0 'fecha desde
            Case 1: KEYFecha KeyAscii, 1 'fecha hasta
        End Select
    Else
        KEYpress KeyAscii
    End If

End Sub

Private Sub KEYFecha(KeyAscii As Integer, indice As Integer)
    KeyAscii = 0
    imgFec_Click (indice)
End Sub

Private Sub KEYpress(KeyAscii As Integer)
    If KeyAscii = 13 Then 'ENTER
        KeyAscii = 0
        SendKeys "{tab}"
    ElseIf KeyAscii = 27 Then Unload Me  'ESC
    End If
End Sub

Private Sub txtCodigo_LostFocus(Index As Integer)
Dim Cad As String, cadTipo As String 'tipo cliente

    'Quitar espacios en blanco por los lados
    txtCodigo(Index).Text = Trim(txtCodigo(Index).Text)
    
    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub

    
    Select Case Index
        Case 0, 1, 2 'FECHAS
            If txtCodigo(Index).Text <> "" Then PonerFormatoFecha txtCodigo(Index)
            
        Case 6 ' seccion
            BdConta = 0
            txtNombre(Index).Text = PonerNombreDeCod(txtCodigo(Index), "seccion", "nomsecci", "codsecci", "N")
            If txtCodigo(Index).Text <> "" Then
                txtCodigo(Index).Text = Format(txtCodigo(Index).Text, "000")
            
                Cad = DevuelveDesdeBDNew(cPTours, "seccion", "numconta", "codsecci", txtCodigo(6).Text, "N") 'numconta
                If Cad <> "" Then BdConta = CByte(Cad)  'numero de conta
            Else
                MsgBox "Debe introducir un código existente en la sección. Revise.", vbExclamation
            End If
        
            
    End Select
End Sub


Private Sub CargaFacturas(DFecha As Date, HFecha As Date)
    Dim db As BaseDatos
    Dim Sql As String
    Dim Rs As ADODB.Recordset
    Dim Rs2 As ADODB.Recordset
    Dim i As Long
    Dim FicheroPDF As String
    Dim c1 As String
    Dim c2 As String
    Dim c3 As String
    Dim c4 As String
    Dim f1 As Date
    Dim f3 As Date
    Dim i1 As Currency
    Dim fr As frmVisReport
On Error GoTo err_CargaFacturas
    
Dim indRPT As Byte 'Indica el tipo de Documento en la tabla "scryst"
Dim nomDocu As String 'Nombre de Informe rpt de crystal
Dim numParam As Byte
Dim cadParam As String

Dim NomClien As String


    Set db = New BaseDatos
    db.tipo = "MYSQL"
    
    db.abrir "vAriagroutil", "root", "aritel"


'    db.abrir "accArigasol", "", ""
    Sql = "select cabfact.*" & _
            " from cabfact where fecfactu >= " & db.Fecha(CDate(txtCodigo(0).Text)) & _
            " and fecfactu <= " & db.Fecha(CDate(txtCodigo(1).Text)) & _
            " and cabfact.codsecci = " & DBSet(txtCodigo(6).Text, "N") & _
            " and cabfact.pasaridoc = 0"
            
    Set Rs = db.cursor(Sql)
    
    
    If Not Rs.EOF Then
        Rs.MoveFirst
        While Not Rs.EOF
            i = i + 1
            lblInf.Caption = "Procesando registro " & CStr(i)
            lblInf.Refresh
            '-- Creamos el pdf
            FicheroPDF = App.path & "\ExpAriDoc.pdf"

'18/02/2010: lo quito para que prueben
'            If Not IntentaMatar(FicheroPDF) Then Err.Raise 53
            
            
            Set fr = New frmVisReport
            
            
            '++monica: seleccionamos que rpt se ha de ejecutar
            cadParam = "|pEmpresa=" & vEmpresa.nomEmpre & "|"
            indRPT = 1 'Impresion de Factura
            
            '[Monica]26/05/2016: otro report para materna
            If EsSeccionMaterna(txtCodigo(6).Text) Then indRPT = 4
            
            
            If Not PonerParamRPT(indRPT, cadParam, numParam, nomDocu) Then Exit Sub
            '++
            fr.NumeroParametros = numParam
            fr.OtrosParametros = cadParam
            fr.ConSubInforme = True
            fr.Informe = App.path & "\Informes\" & nomDocu
            fr.FormulaSeleccion = "{cabfact.codsecci} = " & txtCodigo(6).Text & " and {cabfact.letraser} = '" & Rs!letraser & "' and " & _
                                  "{cabfact.numfactu} =" & CStr(Rs!numfactu) & " and " & _
                                  "{cabfact.fecfactu} = Date(" & Format(Rs!fecfactu, "yyyy") & _
                                                                    "," & Format(Rs!fecfactu, "mm") & _
                                                                    "," & Format(Rs!fecfactu, "dd") & ")"
                                                                    
            '[Monica]18/11/2013: falta indicarle cual es la contabilidad
            fr.Contabilidad = BdConta
            fr.Facturas = True
            fr.FicheroPDF = FicheroPDF
            
            Load fr 'trabaja sin mostrar el formulario
            Screen.MousePointer = vbDefault
            
            NomClien = ""
            If AbrirConexionContaFac(vParamAplic.UsuarioContaFac, vParamAplic.PasswordContaFac, BdConta) Then
                Set vEmpresaFac = New CempresaFac
                If vEmpresaFac.LeerNiveles Then
                    NomClien = DevuelveDesdeBDNewFac("cuentas", "nommacta", "codmacta", CStr(Rs!ctaclien), "T") ' PonerNombreCuenta( CStr(RS!ctaclien), 0, , BdConta, True)
                End If
                Set vEmpresaFac = Nothing
                CerrarConexionContaFac
            End If
            
            
            c1 = CargaParametroFac(vParamAplic.C1Factura, Rs, NomClien)
            c2 = CargaParametroFac(vParamAplic.C2Factura, Rs, NomClien)
            c3 = CargaParametroFac(vParamAplic.C3Factura, Rs, NomClien)
            c4 = CargaParametroFac(vParamAplic.C4Factura, Rs, NomClien)
            
            f1 = Rs!fecfactu
            i1 = Rs!TotalFac
            f3 = Now
            If impor.importaFicheroPDF(FicheroPDF, CLng(txtCarp(0)), c1, c2, c3, c4, f1, f3, i1) Then
                'actualizamos el pasaridoc de facturas
                Sql = "update cabfact set pasaridoc = 1 where codsecci = " & txtCodigo(6).Text & " and letraser = " & DBSet(Rs!letraser, "T")
                Sql = Sql & " and numfactu = " & DBSet(Rs!numfactu, "N") & " and fecfactu = " & DBSet(Rs!fecfactu, "F")
                db.Ejecutar Sql
            End If
            
            Unload fr
            Set fr = Nothing
            
            Rs.MoveNext
        Wend
    End If
    Exit Sub
err_CargaFacturas:
    If Err.Number <> 0 Then
        MsgBox Err.Description & " (" & Err.Number & ")", vbCritical, "CargaFacturas"
    End If
End Sub



Private Function CargaParametroFac(param As Byte, ByRef Rs As ADODB.Recordset, NomClien As String) As String
    Select Case param
        Case 0 'facturas
            CargaParametroFac = Format(Rs!numfactu, "0000000") & "-" & Rs!letraser
        Case 1 'codigo cliente
            CargaParametroFac = Mid(Rs!ctaclien, Len(Rs!ctaclien) - 3, 4)
        Case 2 'nombre cliente
            CargaParametroFac = NomClien
        Case 3 'procedencia
            CargaParametroFac = "ARIAGROUTIL"
        Case Else
            CargaParametroFac = ""
    End Select

End Function


Private Function CargaPath(Codigo As Integer) As String
Dim Nod As Node
Dim J As Integer
Dim i As Integer
Dim C As String
Dim campo1 As String
Dim padre As String
Dim A As String

Dim Sql As String
Dim Rs As ADODB.Recordset

    'distinto del cargapath de parametros de aplicacion

    Sql = "select nombre, padre from carpetas where codcarpeta = " & DBSet(Codigo, "N")
    Set Rs = ardDB.cursor(Sql)

    If Not Rs.EOF Then
        C = "\" & Rs!Nombre
        If Rs!padre > 0 Then
            C = CargaPath(CInt(Rs!padre)) & C
        End If
    End If
    
    CargaPath = C
End Function

Private Function IntentaMatar(FicheroPDF As String) As Boolean
Dim i As Integer

    On Error Resume Next
    i = 1
    IntentaMatar = False
    Do
        If Dir(FicheroPDF, vbArchive) <> "" Then
            Kill FicheroPDF
            If Err.Number <> 0 Then
                Err.Clear
                i = i + 1
            Else
                IntentaMatar = True
                i = 6
            End If
        Else
            IntentaMatar = True
            i = 6
        End If
    Loop Until i < 5 Or IntentaMatar = True
    
    
End Function

