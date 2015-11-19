VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmDatos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Datos a importar"
   ClientHeight    =   6990
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10380
   Icon            =   "frmDatos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6990
   ScaleWidth      =   10380
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   7920
      TabIndex        =   5
      Top             =   6480
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      Cancel          =   -1  'True
      Caption         =   "Salir"
      Height          =   375
      Left            =   9120
      TabIndex        =   4
      Top             =   6480
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Height          =   375
      Left            =   3360
      Picture         =   "frmDatos.frx":6852
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Comprobar Cta.Contable socio"
      Top             =   6480
      Width           =   375
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   2520
      Top             =   -120
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDatos.frx":6954
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDatos.frx":C576
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDatos.frx":CB10
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDatos.frx":D0AA
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDatos.frx":E83C
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame1 
      Height          =   615
      Left            =   120
      TabIndex        =   1
      Top             =   6360
      Width           =   3015
      Begin VB.Label Label1 
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   2775
      End
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   6015
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   10095
      _ExtentX        =   17806
      _ExtentY        =   10610
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   7
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Socio"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Nombre"
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Fecha"
         Object.Width           =   1940
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Nº Factura"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   4
         Text            =   "Base"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   5
         Text            =   "Iva"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   6
         Text            =   "Total"
         Object.Width           =   1764
      EndProperty
   End
End
Attribute VB_Name = "frmDatos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Cadena As String
Public Digito As String

Dim PrimeraVez As Boolean
Dim errores As String
Dim Vector  'Publico , para no tener que ir pasandolo de un sitio a otro


Private Sub cmdAceptar_Click()
Dim I As Integer
Dim C As String
Dim TodoOk As Boolean
Dim N As Integer
Dim Cont As Integer
Dim Codmacta As String

    C = ""
    For I = 1 To ListView1.ListItems.Count
        If ListView1.ListItems(I).Checked Then
            If ListView1.ListItems(I).SmallIcon = 1 Or ListView1.ListItems(I).SmallIcon = 4 Then
                C = C & "o"
            End If
        End If
    Next I
    
    
    If C = "" Then
        MsgBox "Ninguna dato a integrar", vbExclamation
        Exit Sub
    Else
        C = Len(C)
        C = "Desea generar las facturas(" & C & ") de movil ?"
        If MsgBox(C, vbQuestion + vbYesNo) = vbNo Then Exit Sub
    End If
    
    
    Screen.MousePointer = vbHourglass
    N = ListView1.ListItems.Count
    Cont = 0
    TodoOk = True
    For I = N To 1 Step -1
        If ListView1.ListItems(I).Checked Then
            If ListView1.ListItems(I).SmallIcon = 1 Or ListView1.ListItems(I).SmallIcon = 4 Then
                Label1.Caption = ListView1.ListItems(I).SubItems(1)
                Label1.Refresh
                'Haremos el insert
                C = "INSERT INTO telmovil (numserie,numfactu,fecfactu,codmacta,"
                C = C & "baseimpo,cuotaiva,totalfac,intconta) VALUES (" & DBSet(Trim(vParamAplic.NumSerieTel), "T") & ","
                
'                Codmacta = CtaContableSocio(ListView1.ListItems(I).Text, cContaTel)
                Codmacta = Trim(vParamAplic.RaizCtaSocTel & Format(ListView1.ListItems(I).Text, "00000"))
                
                C = C & DBSet(ListView1.ListItems(I).SubItems(3), "N") & ","
                C = C & DBSet(ListView1.ListItems(I).SubItems(2), "F") & ","
                C = C & DBSet(Codmacta, "T") & ","
                C = C & DBSet(ListView1.ListItems(I).SubItems(4), "N") & ","
                C = C & DBSet(ListView1.ListItems(I).SubItems(5), "N") & ","
                C = C & DBSet(ListView1.ListItems(I).SubItems(6), "N") & ","
                C = C & "0)"
                
                If Not Ejecutar(C) Then
                    TodoOk = False
                    If MsgBox("¿Desea continuar con el resto de facturas?", vbQuestion + vbYesNo) = vbNo Then
                        Exit For
                    End If
                Else
                    ListView1.ListItems.Remove I
                    Cont = Cont + 1
                End If
            End If
        End If
    Next I
    
    Label1.Caption = ""
    Screen.MousePointer = vbDefault
    
    If TodoOk Then
        MsgBox "Se han traspasado " & Cont & " facturas.", vbInformation
        Unload Me
    End If
End Sub


Private Function Ejecutar(ByRef SQL As String) As Boolean
    On Error Resume Next
    conn.Execute SQL
    If Err.Number = 0 Then
        Ejecutar = True
    Else
        MsgBox Err.Description, vbExclamation
        Err.Clear
        Ejecutar = False
    End If
End Function



Private Function TransformaComasPuntos(Cade As String) As String
Dim C As String 'QUITAR
Dim I As Integer

    I = InStr(1, Cade, ",")
    If I > 0 Then
        TransformaComasPuntos = Mid(Cade, 1, I - 1) & "." & Mid(Cade, I + 1)
    Else
        TransformaComasPuntos = Cade
    End If

End Function

Private Sub Command1_Click()
Dim I As Integer
Dim C As String
Dim J As Integer
Dim Inicio As Integer
Dim Desplaza As Integer

    Me.Label1.Caption = "Comprobar Cta. socio"
    Label1.Refresh
    C = ""
    J = 0
    Inicio = 1
    Desplaza = 4
    For I = 1 To ListView1.ListItems.Count
        If ListView1.ListItems(I).Checked Then
'            C = C & " or nifdatos = " & DBSet(ListView1.ListItems(I).Text, "T")
            C = C & " or codmacta = " & DBSet(vParamAplic.RaizCtaSocTel & Format(ListView1.ListItems(I).Text, "00000"), "T")
        End If
        J = J + 1
        If J = Desplaza Then
            'Hacer select
            C = Mid(C, 4)
            HacerSelectYVerificar Inicio, C, Desplaza
            
            'Para hacer mas
            Inicio = I + 1
            C = ""
            J = 0
        End If
    Next I
    
    If J > 0 Then
            C = Mid(C, 4)
            HacerSelectYVerificar Inicio, C, J
    End If
    Label1.Caption = ""
    Me.Command1.Enabled = False
End Sub

Private Sub HacerSelectYVerificar(Incio As Integer, ByRef Ca As String, Desplazamiento As Integer)
Dim RS As ADODB.Recordset
Dim Socios As String
Dim J As Integer

    On Error GoTo EHacerSelectYVerificar
    Cadena = "Select codmacta from cuentas where " & Ca

    Set RS = New ADODB.Recordset
    
    RS.Open Cadena, ConnContaTel, adOpenForwardOnly, adLockPessimistic, adCmdText
    Cadena = "|"
    
    
    While Not RS.EOF
        Cadena = Cadena & RS.Fields(0) & "|"
        RS.MoveNext
    Wend
    RS.Close
    Set RS = Nothing
    Debug.Print Incio & "  -  " & (Incio + Desplazamiento) - 1
    For J = Incio To (Incio + Desplazamiento) - 1
        If ListView1.ListItems(J).Checked Then
            Socios = "|" & Trim(vParamAplic.RaizCtaSocTel & Format(Val(ListView1.ListItems(J).Text), "00000")) & "|"
'            Socios = "|" & CtaContableSocio(ListView1.ListItems(J).Text, cContaTel) & "|"
            If InStr(1, Cadena, Socios) > 0 Then
                ListView1.ListItems(J).SmallIcon = 4   'OK
                
            Else
                ListView1.ListItems(J).SmallIcon = 5
                ListView1.ListItems(J).Checked = False
            End If
        End If
    Next J
    
    Set RS = Nothing
    
    Exit Sub
EHacerSelectYVerificar:
    MsgBox Err.Description, vbCritical
End Sub


Private Sub Command2_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
    If PrimeraVez Then
        PrimeraVez = False
        Label1.Caption = "Procesar fichero"
        ProcesarFichero
        Label1.Caption = ""
        Me.Refresh
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
    
    PrimeraVez = True
    
    If vParamAplic.NumeroContaTel = 0 Then
        MsgBox "Para Realizar este proceso debe tener conexión a contabilidad. Revise", vbExclamation
        Unload Me
    Else
        Set ListView1.SmallIcons = Me.ImageList1
    End If
End Sub



Private Sub ProcesarFichero()
Dim Fin As Boolean
Dim NF As Integer
Dim C As String
Dim Bien As Boolean
Dim I As Integer

    
    NF = FreeFile
    Fin = False
    Bien = False
    Open Cadena For Input As #NF
    If EOF(NF) Then Fin = True
    I = 0
    Me.Refresh
    While Not Fin
        Line Input #NF, C
        If I > 0 Then
            If Not ProcesarLinea(C) Then
                
            Else
                Bien = True
            End If
            
        End If
        I = I + 1
        If (I Mod 20) Then
            Me.Refresh
        End If
        If EOF(NF) Then Fin = True
    Wend
    Close #NF
    If Bien Then
    
    End If
End Sub


Private Function ProcesarLinea(Linea As String) As Boolean
Dim IT As ListItem
Dim Campos As Integer
Dim I As Integer
Dim C As String
Dim S As String
    Vector = Split(Linea, ";")
    Campos = UBound(Vector)
    If Campos > 0 And Campos < 13 Then
            
        'Comprobaremos que los campos a insertar son correcto
        '----------------------------------------------------
        I = ListView1.ListItems.Count + 1
        If Not ComprobarLinea Then
            errores = errores & vbCrLf & Linea & vbCrLf & vbCrLf
            Set IT = ListView1.ListItems.Add(, "C" & CStr(I), Mid(Linea, 1, 15))
            IT.ToolTipText = Linea
            IT.SmallIcon = 2
            IT.Ghosted = True
        Else
        
            C = QuitaPrimeraUltimaComilla(CStr(Vector(5)))
'            Set IT = ListView1.ListItems.Add(, "C" & CStr(I), vParamAplic.NumeroContaTel & Format(C, "00000"))
            Set IT = ListView1.ListItems.Add(, "C" & CStr(I), C)
            IT.Tag = C
            
            'Si el codigo es cero:, NO se integra
            S = QuitaPrimeraUltimaComilla(CStr(Vector(5)))
            If Val(S) > 0 Then
                IT.Checked = True
                IT.SmallIcon = 1
            Else
                IT.Checked = False
                IT.SmallIcon = 3
            End If
            
            IT.SubItems(1) = QuitaPrimeraUltimaComilla(CStr(Vector(4)))
            IT.SubItems(2) = QuitaPrimeraUltimaComilla(CStr(Vector(8)))
            
            C = QuitaPrimeraUltimaComilla(CStr(Vector(6)))
            If Len(C) >= 5 Then C = Right(C, 5)
            IT.SubItems(3) = Digito & C
            
            For I = 4 To 6
                'Para la columna 4 es el 10
                C = PonerImporte(QuitaPrimeraUltimaComilla(CStr(Vector(I + 5)))) ' antes I + 6
                IT.SubItems(I) = C
            Next
             
        End If
    End If
End Function

Private Function PonerImporte(C As String) As String
Dim J As Integer
    
    J = InStr(1, C, "€")
    If J > 0 Then C = Mid(C, 1, J - 1)
    PonerImporte = Trim(C)
End Function

Private Function ComprobarLinea() As Boolean
 On Error GoTo EComprobarLinea
 Dim TienErr As String
 Dim I As Integer
 Dim C As String
    ComprobarLinea = False
    TienErr = ""
    
    
    'Campo 3. codsocio
    C = Vector(5)
    C = QuitaPrimeraUltimaComilla(C)
    
    'añadido
    If C = "0" Then
        ComprobarLinea = True
        Exit Function
    End If
    
    If Not ComprobarCtaConta(C) Then TienErr = TienErr & " Cta.Socio no existe en contabilidad."
        
    
    'Campo 5. Codsocio
    For I = 9 To 11 ' antes de 10 a 12
        If Not IsNumeric(Vector(I)) Then TienErr = TienErr & " Importe incorrecto( " & I - 9 & ")."
    Next I
 
 
    If TienErr = "" Then
        ComprobarLinea = True
    Else
        errores = errores & TienErr
    End If
    Exit Function
EComprobarLinea:
    MsgBox Err.Description, vbExclamation
End Function


Private Function QuitaPrimeraUltimaComilla(C As String) As String
    If Mid(C, 1, 1) = """" Then C = Mid(C, 2)
    If Right(C, 1) = """" Then C = Mid(C, 1, Len(C) - 1)
    QuitaPrimeraUltimaComilla = C
End Function

Private Function ComprobarCtaConta(C As String) As Boolean
    If vParamAplic.NumeroContaTel <> 0 Then
        ComprobarCtaConta = (DevuelveDesdeBDNew(cContaTel, "cuentas", "codmacta", "codmacta", Trim(vParamAplic.RaizCtaSocTel & Format(C, "00000")), "T") <> "")
    End If
End Function
