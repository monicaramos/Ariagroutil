VERSION 5.00
Begin VB.Form frmIdentifica 
   BackColor       =   &H00800000&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   5490
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7965
   LinkTopic       =   "Form1"
   ScaleHeight     =   5490
   ScaleWidth      =   7965
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   210
      Top             =   3270
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      IMEMode         =   3  'DISABLE
      Index           =   1
      Left            =   4320
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   4980
      Width           =   3015
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   0
      Left            =   4320
      TabIndex        =   0
      Top             =   4020
      Width           =   3015
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Label3"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   270
      TabIndex        =   6
      Top             =   90
      Width           =   7305
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   195
      Left            =   930
      TabIndex        =   5
      Top             =   5250
      Width           =   1725
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Cargando ....."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   375
      Index           =   2
      Left            =   5040
      TabIndex        =   4
      Top             =   4920
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Usuario"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   375
      Index           =   0
      Left            =   4320
      TabIndex        =   2
      Top             =   3600
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Password"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   375
      Index           =   1
      Left            =   4320
      TabIndex        =   3
      Top             =   4560
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Image Image1 
      Height          =   5535
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   7935
   End
End
Attribute VB_Name = "frmIdentifica"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' +-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-
' +-+- Autor: DAVID (refet per C�SAR) +-+-
' +-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-

Option Explicit

Dim PrimeraVez As Boolean
Dim T1 As Single
Dim vSegundos As Integer

Private Sub Form_Activate()
    If PrimeraVez Then
        PrimeraVez = False

         Me.Refresh
         PonerVisible True
         If Text1(0).Text <> "" Then
            PonerFoco Text1(1)
         Else
            PonerFoco Text1(0)
         End If
             
         Me.Timer1.Enabled = True
             
         'Leemos el ultimo usuario conectado
         NumeroEmpresaMemorizar True
         
         T1 = T1 + 2.5 - Timer
         If T1 > 0 Then espera T1

         PonerVisible True
         If Text1(0).Text <> "" Then
            Text1(1).SetFocus
        Else
            Text1(0).SetFocus
        End If

    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
'    Screen.MousePointer = vbHourglass
    'PonerVisible False
'    T1 = Timer
    'Text1(0).Text = "root"
 '   Text1(1).Text = "aritel"
    PrimeraVez = True
    CargaImagen
    Label2.Caption = "Ver. " & App.Major & "." & App.Minor & "." & App.Revision
    
    Label3.Caption = ""
    vSegundos = 60
    Label3.Caption = ""
    
End Sub

Private Sub CargaImagen()
    On Error Resume Next
    Me.Image1 = LoadPicture(App.path & "\entrada.dat")
    Me.Height = Me.Image1.Height
    Me.Width = Me.Image1.Width
    If Err.Number <> 0 Then
        MsgBox Err.Description & vbCrLf & vbCrLf & "Error cargando", vbCritical
        Set conn = Nothing
        End
    End If
End Sub


Private Sub Form_Unload(Cancel As Integer)
    NumeroEmpresaMemorizar False
End Sub

Private Sub Text1_GotFocus(Index As Integer)
    With Text1(Index)
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub KEYpress(KeyAscii As Integer)
Dim cerrar As Boolean

    KEYpressGnral KeyAscii, 2, cerrar
    If cerrar Then Unload Me

End Sub

Private Sub Text1_LostFocus(Index As Integer)
    Text1(Index).Text = Trim(Text1(Index).Text)

    'Comprobamos si los dos estan con datos
    If Text1(0).Text <> "" And Text1(1).Text <> "" Then
        Validar
    End If
End Sub


Private Sub Validar()
Dim OK As Byte
Dim Cad As String


    Set vSesion = New CSesion

    If vSesion.leer(Text1(0).Text) = 0 Then
        'Con exito
        If vSesion.PasswdPROPIO = Text1(1).Text Then
            OK = 0
        Else
            OK = 1
        End If
    Else
        If Text1(0).Text = "root" And Text1(1).Text = "aritel" Then
            Cad = "insert into usuarios (codusu, nomusu, login, passwordpropio, nivelusuges) "
            Cad = Cad & " values (0,'root','root','aritel',0)"
            conn.Execute Cad
            OK = 0
        Else
            OK = 2
        End If

    End If

    If OK <> 0 Then
        MsgBox "Usuario o Password Incorrecto", vbExclamation

        Text1(1).Text = ""
        PonerFoco Text1(0)
    Else
        'OK
        If vSesion.Nivel < 0 Then
            MsgBox "Usuario sin Permisos.", vbExclamation
            End
        Else
            PonerVisible False
            Me.Refresh
            espera 0.2
        
        '    If ComprovaVersio Then
        '        MsgBox "Existe una versi�n m�s reciente de la aplicaci�n. Se va a proceder a la actualizaci�n", vbInformation
        '        Shell App.Path & "\PlannerUpdate.exe", vbNormalFocus
        '        End
        '    End If
        
            'Carga Datos de la Empresa y los Niveles de cuentas de Contabilidad de la empresa
            'Crea la Conexion a la BD de la Contabilidad
            LeerDatosEmpresa
        
            InicializarFormatos
            teclaBuscar = 43
        
            Load frmPpal
            
            MDIppal.Show
            
            Unload Me
            
        End If
    
    End If


End Sub





Private Sub PonerVisible(visible As Boolean)
    'Label1(2).visible = Not visible  'Cargando
    Text1(0).visible = visible
    Text1(1).visible = visible
    Label1(0).visible = visible
    Label1(1).visible = visible
End Sub

'Lo que haremos aqui es ver, o guardar, el ultimo numero de empresa
'a la que ha entrado, y el usuario
Private Sub NumeroEmpresaMemorizar(leer As Boolean)
Dim NF As Integer
Dim Cad As String
On Error GoTo ENumeroEmpresaMemorizar

    Cad = App.path & "\ultusu.dat"
    If leer Then
        If Dir(Cad) <> "" Then
            NF = FreeFile
            Open Cad For Input As #NF
            Line Input #NF, Cad
            Close #NF
            Cad = Trim(Cad)
            
                'El primer pipe es el usuario
                Text1(0).Text = Cad
    
        End If
    Else 'Escribir
        NF = FreeFile
        Open Cad For Output As #NF
        Cad = Text1(0).Text
        Print #NF, Cad
        Close #NF
    End If
ENumeroEmpresaMemorizar:
    Err.Clear
End Sub


Private Sub Timer1_Timer()
    'Label3 = "Si no entra en " & vSegundos & " segundos. La aplicaci�n se cerrar�."
    If vSegundos < 50 Then
        Label3 = "Si no hace login, la pantalla se cerrar� autom�ticamente en " & " " & vSegundos & " segundos"
        Me.Refresh
        DoEvents
    End If
    vSegundos = vSegundos - 1
    If vSegundos = -1 Then Unload Me
End Sub

