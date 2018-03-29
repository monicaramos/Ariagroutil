VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmRenAvnics 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Renovación AVNICS"
   ClientHeight    =   3810
   ClientLeft      =   45
   ClientTop       =   2430
   ClientWidth     =   5970
   Icon            =   "frmRenAvnics.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3810
   ScaleWidth      =   5970
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtCodigo 
      Alignment       =   1  'Right Justify
      Height          =   285
      Index           =   1
      Left            =   2565
      MaxLength       =   4
      TabIndex        =   3
      Top             =   2520
      Width           =   735
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Cálculo utilizando la fecha del Avnic"
      Height          =   375
      Left            =   720
      TabIndex        =   2
      Top             =   2160
      Width           =   3030
   End
   Begin VB.TextBox txtCodigo 
      Height          =   285
      Index           =   2
      Left            =   2535
      MaxLength       =   10
      TabIndex        =   0
      Top             =   1245
      Width           =   1095
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   3360
      TabIndex        =   4
      Top             =   3135
      Width           =   975
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   4485
      TabIndex        =   5
      Top             =   3135
      Width           =   975
   End
   Begin VB.TextBox txtCodigo 
      Alignment       =   1  'Right Justify
      Height          =   285
      Index           =   0
      Left            =   2535
      MaxLength       =   4
      TabIndex        =   1
      Top             =   1725
      Width           =   735
   End
   Begin MSComDlg.CommonDialog cd1 
      Left            =   1920
      Top             =   3000
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Número de Meses"
      ForeColor       =   &H00000080&
      Height          =   195
      Index           =   1
      Left            =   720
      TabIndex        =   9
      Top             =   2565
      Width           =   1290
   End
   Begin VB.Image imgFec 
      Height          =   240
      Index           =   2
      Left            =   2175
      Picture         =   "frmRenAvnics.frx":000C
      ToolTipText     =   "Buscar fecha"
      Top             =   1245
      Width           =   240
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Seleccione Vencimiento y Ejercicio:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   6
      Left            =   720
      TabIndex        =   8
      Top             =   600
      Width           =   4260
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Fecha Vencimiento"
      ForeColor       =   &H00000080&
      Height          =   195
      Index           =   3
      Left            =   720
      TabIndex        =   7
      Top             =   1245
      Width           =   1365
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Nuevo Ejercicio"
      ForeColor       =   &H00000080&
      Height          =   195
      Index           =   0
      Left            =   720
      TabIndex        =   6
      Top             =   1725
      Width           =   1125
   End
End
Attribute VB_Name = "frmRenAvnics"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' +-+-+-+-+-+-+-+-+-+-+-
' +-+- Autor:MANOLO +-+-
' +-+-+-+-+-+-+-+-+-+-+-

Option Explicit

Private WithEvents frmC As frmCal 'Calendario de Fechas
Attribute frmC.VB_VarHelpID = -1

'GENERALES PARA PASARLE A CRYSTAL REPORT
Private cadFormula As String 'Cadena con la FormulaSelection para Crystal Report
Private cadSelect As String 'Cadena para comprobar si hay datos antes de abrir Informe
Private cadParam As String 'Cadena con los parametros para Crystal Report
Private numParam As Byte 'Numero de parametros que se pasan a Crystal Report

Dim indCodigo As Integer 'indice para txtCodigo
 
'Se inicializan para cada Informe (tabla de BD a la que hace referencia
Dim tabla As String
Dim Codigo As String 'Código para FormulaSelection de Crystal Report
Dim TipCod As String

Dim PrimeraVez As Boolean

Private Sub KEYpress(KeyAscii As Integer)
    If KeyAscii = 13 Then 'ENTER
        KeyAscii = 0
        SendKeys "{tab}"
    ElseIf KeyAscii = 27 Then Unload Me  'ESC
    End If
End Sub

Private Sub Check1_Click()
    If Check1.Value <> 1 Then
        Label1(1).visible = False
        txtCodigo(1).visible = False
    Else
        Label1(1).visible = True
        txtCodigo(1).visible = True
    End If
End Sub

Private Sub cmdAceptar_Click()
'Obtener la cadena SQL para eliminar los registros seleccionados
Dim cDesde As String 'Ejercicio
Dim fDesde As String 'Fecha Vto.
Dim SQL As String

    If Not DatosOk Then Exit Sub

    InicializarVbles
    SQL = ""
    'Valores para Formula seleccion del informe
    cDesde = Trim(txtCodigo(0).Text)
    fDesde = Trim(txtCodigo(2).Text)
    
    SQL = "WHERE avnic.codialta <> 2"
    
    RenovarAvnics SQL
    
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
    If PrimeraVez Then PrimeraVez = False
    
    Screen.MousePointer = vbDefault
    PonerFoco txtCodigo(2)
End Sub

Private Sub Form_Load()

    PrimeraVez = True
    Limpiar Me
    
    '###Descomentar
'    CommitConexion
    
    txtCodigo(0).Text = Format(Now, "yyyy") '
    
    tabla = "avnic"
    
    Label1(1).visible = False
    txtCodigo(1).visible = False
    
    'Esto se consigue poneinedo el cancel en el opcion k corresponda
    Me.cmdCancel.Cancel = True
End Sub

Private Sub frmC_Selec(vFecha As Date)
   txtCodigo(CByte(imgFec(2).Tag)).Text = Format(vFecha, "dd/MM/yyyy")
End Sub

Private Sub imgFec_Click(Index As Integer)
    'Calendario de Fechas
    Dim esq As Long
    Dim dalt As Long
    Dim menu As Long
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
    
    menu = Me.Height - Me.ScaleHeight 'ací tinc el heigth del menú i de la toolbar
   
    ' es desplega baix i cap a la dreta
    frmC.Left = esq + imgFec(Index).Parent.Left + 30
    frmC.Top = dalt + imgFec(Index).Parent.Top + imgFec(Index).Height + menu - 40
    
    ' ***canviar l'index de imgFec pel 1r index de les imagens de buscar data***
    imgFec(2).Tag = Index 'independentment de les dates que tinga, sempre pose l'index en la 27
    If txtCodigo(Index).Text <> "" Then frmC.NovaData = txtCodigo(Index).Text
    
    frmC.Show vbModal
    Set frmC = Nothing
    PonerFoco txtCodigo(CByte(imgFec(2).Tag))
    ' **************************************************************************
End Sub

Private Sub txtCodigo_GotFocus(Index As Integer)
    ConseguirFoco txtCodigo(Index), 3
End Sub

Private Sub txtCodigo_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    KEYdown KeyCode
End Sub

Private Sub txtCodigo_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = teclaBuscar Then
        Select Case Index
            Case 2: KEYFecha KeyAscii, 2 'fecha
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

    'Quitar espacios en blanco por los lados
    txtCodigo(Index).Text = Trim(txtCodigo(Index).Text)
    
    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub
    
    Select Case Index
        Case 0 'EJERCICIO
            If txtCodigo(Index).Text <> "" Then PonerFormatoEntero txtCodigo(Index)

        Case 2 'FECHA
            If txtCodigo(Index).Text <> "" Then PonerFormatoFecha txtCodigo(Index)
    End Select
End Sub

Private Sub InicializarVbles()
    cadFormula = ""
    cadSelect = ""
    cadParam = ""
    numParam = 0
End Sub

Private Function RenovarAvnics(cadW As String) As Boolean
'Eliminar Albaranes de Fecha y Turno: Tabla (scaalb)
'que cumplan los criterios seleccionados en la cadena WHERE cadW

Dim Cad As String, SQL As String
Dim Rs As ADODB.Recordset
Dim todasElim As Boolean

    On Error GoTo EEliminar

    Cad = "Va a renovar los AVNICS." & vbCrLf
    Cad = Cad & vbCrLf & vbCrLf & "¿Desea Comenzar? "
    
    If MsgBox(Cad, vbQuestion + vbYesNoCancel) = vbYes Then     'Empezamos
        'Hay que eliminar
        On Error GoTo EEliminar
        Screen.MousePointer = vbHourglass
        
        If CrearAvnics(txtCodigo(2).Text, txtCodigo(0).Text) Then
            MsgBox "El proceso se realizó correctamente.", vbInformation
            Unload Me
        Else
            MsgBox "ATENCIÓN: Se ha producido error en el proceso.", vbInformation
            Unload Me
        End If
    End If
    Screen.MousePointer = vbDefault
    
EEliminar:
    Screen.MousePointer = vbDefault
    If Err.Number > 0 Then MuestraError Err.Number, "Renovar Avnics", Err.Description
End Function

Private Function CrearAvnics(FecVto As String, AnoEje As String) As Boolean
'Eliminar las lineas y la Cabecera de un Caja. Tablas: cajascab, cajaslin
Dim SQL As String
Dim sql2 As String
Dim b As Boolean
Dim Rs As ADODB.Recordset
Dim Cadena As String
Dim Mes As Integer
Dim Ano As Integer
Dim Fecha1 As String
Dim Fecha As Date

    On Error GoTo ECrearAvnics
    CrearAvnics = False
    b = False
    
    conn.BeginTrans
    
    If Me.Check1.Value = False Then
        SQL = "insert into `avnic` (`codavnic`,`nombrper`,`nifperso`,`nifrepre`,"
        SQL = SQL & "`codposta`,`nomcalle`,`poblacio`,`provinci`,`codialta`,`codbanco`,`codsucur`,"
        SQL = SQL & "`cuentaba`,`digcontr`,`imporper`,`imporret`,`anoejerc`,`nifpers1`,`fechalta`,`nombper1`,"
        SQL = SQL & "`nomcall1`,`poblaci1`,`provinc1`,`codpost1`,`fechavto`,`porcinte`,"
        SQL = SQL & "`importes`,`codmacta`,`observac`,`iban`) "
        SQL = SQL & "SELECT codavnic,nombrper,nifperso,nifrepre,"
        SQL = SQL & "codposta,nomcalle,poblacio,provinci,0,codbanco,codsucur,"
        SQL = SQL & "cuentaba,digcontr,0,0," & AnoEje & ",nifpers1,fechalta,nombper1,"
        SQL = SQL & "nomcall1,poblaci1,provinc1,codpost1,'" & Format(FecVto, FormatoFecha) & "',porcinte,"
        SQL = SQL & "importes,codmacta,observac, iban FROM avnic "
        SQL = SQL & "WHERE codialta <> 2 AND anoejerc =" & AnoEje - 1
        'AVNICS
        conn.Execute SQL
    Else
        SQL = "SELECT codavnic,nombrper,nifperso,nifrepre,"
        SQL = SQL & "codposta,nomcalle,poblacio,provinci,0,codbanco,codsucur,"
        SQL = SQL & "cuentaba,digcontr,0,0," & AnoEje & ",nifpers1,fechalta,nombper1,"
        SQL = SQL & "nomcall1,poblaci1,provinc1,codpost1,fechavto,porcinte,"
        SQL = SQL & "importes,codmacta,observac, iban FROM avnic "
        SQL = SQL & "WHERE codialta <> 2 AND anoejerc =" & AnoEje - 1
    
        
        sql2 = "insert into `avnic` (`codavnic`,`nombrper`,`nifperso`,`nifrepre`,"
        sql2 = sql2 & "`codposta`,`nomcalle`,`poblacio`,`provinci`,`codialta`,`codbanco`,`codsucur`,"
        sql2 = sql2 & "`cuentaba`,`digcontr`,`imporper`,`imporret`,`anoejerc`,`nifpers1`,`fechalta`,`nombper1`,"
        sql2 = sql2 & "`nomcall1`,`poblaci1`,`provinc1`,`codpost1`,`fechavto`,`porcinte`,"
        sql2 = sql2 & "`importes`,`codmacta`,`observac`,`iban`) values "
        
        Set Rs = New ADODB.Recordset
        Rs.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        While Not Rs.EOF
            Fecha1 = CStr(DateAdd("m", CInt(txtCodigo(1).Text), Rs!fechavto)) ' le suma el numero de meses a la fecha de vto
            
            Cadena = "(" & DBSet(Rs!codavnic, "N") & "," & DBSet(Rs!nombrper, "T") & "," & DBSet(Rs!nifperso, "T") & "," & DBSet(Rs!nifrepre, "T") & ","
            Cadena = Cadena & DBSet(Rs!Codposta, "T") & "," & DBSet(Rs!nomcalle, "T") & "," & DBSet(Rs!poblacio, "T") & "," & DBSet(Rs!provinci, "T") & ","
            Cadena = Cadena & "0" & "," & DBSet(Rs!codbanco, "N") & "," & DBSet(Rs!codsucur, "N") & ","
            Cadena = Cadena & DBSet(Rs!cuentaba, "T") & "," & DBSet(Rs!digcontr, "T") & ",0,0," & AnoEje & "," & DBSet(Rs!nifpers1, "T") & ","
            Cadena = Cadena & DBSet(Rs!fechalta, "F") & "," & DBSet(Rs!nombper1, "T") & ","
            Cadena = Cadena & DBSet(Rs!nomcall1, "T") & "," & DBSet(Rs!poblaci1, "T") & "," & DBSet(Rs!provinc1, "T") & "," & DBSet(Rs!codpost1, "T") & ","
            Cadena = Cadena & DBSet(Fecha1, "F") & "," & DBSet(Rs!Porcinte, "N") & ","
            Cadena = Cadena & DBSet(Rs!importes, "N") & "," & DBSet(Rs!Codmacta, "T") & "," & DBSet(Rs!observac, "T") & "," & DBSet(Rs!Iban, "T") & ")"
        
            sql2 = sql2 & Cadena & ","
        
            Rs.MoveNext
        Wend
    
        ' quitamos la ultima coma del ultimo registro
        sql2 = Mid(sql2, 1, Len(sql2) - 1)
        
        conn.Execute sql2
        Set Rs = Nothing
    
    End If
    
    CrearAvnics = True
    b = True
    
ECrearAvnics:
    If Err.Number <> 0 Then
        MuestraError Err.Number, Err.Description, "Insertar registros"
        b = False
    End If
    
    If Not b Then
        conn.RollbackTrans
    Else
        conn.CommitTrans
    End If
    CrearAvnics = b
End Function


Private Function DatosOk() As Boolean
Dim SQL As String
Dim b As Boolean
    b = True
    If Me.Check1.Value = 1 Then
        If txtCodigo(1).Text = "" Then
            MsgBox "Debe introducir un valor en el número de meses que se va a incrementar.", vbExclamation
            b = False
        Else
            If CInt(txtCodigo(1).Text) < 1 Or CInt(txtCodigo(1).Text) > 12 Then
                MsgBox "El rango de meses que podemos incremetar es de uno a doce meses.", vbExclamation
                b = False
            End If
        End If
    End If
    DatosOk = b
End Function
