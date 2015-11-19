VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmTrasTele 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Procesar fichero telefonia"
   ClientHeight    =   1755
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8190
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1755
   ScaleWidth      =   8190
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog cd1 
      Left            =   7440
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox Text1 
      Height          =   300
      Index           =   2
      Left            =   1680
      MaxLength       =   1
      TabIndex        =   1
      Top             =   750
      Width           =   255
   End
   Begin VB.CommandButton Command2 
      Cancel          =   -1  'True
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   6840
      Picture         =   "frmTrasTele.frx":0000
      TabIndex        =   3
      Top             =   1200
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Aceptar"
      Height          =   375
      Left            =   5520
      TabIndex        =   2
      Top             =   1200
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   300
      Index           =   1
      Left            =   1680
      TabIndex        =   0
      Top             =   240
      Width           =   6375
   End
   Begin VB.Label Label1 
      Caption         =   "1er digito factura"
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   5
      Top             =   750
      Width           =   1335
   End
   Begin VB.Image Image1 
      Height          =   360
      Left            =   1260
      Picture         =   "frmTrasTele.frx":5E62
      Top             =   225
      Width           =   360
   End
   Begin VB.Label Label1 
      Caption         =   "PATH  fichero"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   4
      Top             =   270
      Width           =   1215
   End
End
Attribute VB_Name = "frmTrasTele"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Cambiado As Boolean
Dim Texto1 As String


Private Sub Command1_Click()
Dim C As String
Dim i As Integer


    C = "F"
    For i = 1 To 2
        Text1(i).Text = Trim(Text1(i).Text)
        If Text1(i).Text = "" Then
            C = ""
            Exit For
        End If
    Next i
                    
    If C = "" Then
        MsgBox "Todos los campos son requeridos", vbExclamation
        Exit Sub
    End If
    
    
    'Campo ruta del fichero
    If Dir(Text1(1).Text, vbArchive) = "" Then
        MsgBox "El fichero NO existe", vbExclamation
        Exit Sub
    End If
    
    'Campo numerico
    If Not IsNumeric(Text1(2).Text) Then
        MsgBox "Campo numerico", vbExclamation
        Exit Sub
    End If
                    
        
     'Abrir conexion mutibase
     Screen.MousePointer = vbHourglass
' la conexion es la de telefonia
'    If AbrirConexion(Text1(0).Text) Then
        If Cambiado Then Predeterminados False
        'LLEGADO AQUI procesaremos el fichero
        frmDatos.CADENA = Text1(1).Text
        frmDatos.Digito = Text1(2).Text
        frmDatos.Show vbModal
        Unload Me
'    End If
    Screen.MousePointer = vbDefault
End Sub




Private Sub Command2_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
    If vParamAplic.NumeroContaTel = 0 Then
        MsgBox "Para Realizar este proceso debe tener conexión a contabilidad. Revise", vbExclamation
        Unload Me
    End If
End Sub

Private Sub Form_Load()
    Text1(1).Text = ""
    Predeterminados True
    Cambiado = False
End Sub


Private Function LeerLinea(ByRef NFi As Integer) As String
Dim Ca As String
    On Error Resume Next
    Line Input #NFi, Ca
    If Err.Number <> 0 Then
        Err.Clear
        LeerLinea = ""
    Else
        LeerLinea = Ca
    End If
End Function

Private Sub Predeterminados(leer As Boolean)
Dim C As String
Dim NF As Integer
    On Error GoTo ELeerLinea
    
    C = App.path & "\Predet.dat"
    NF = FreeFile
    If leer Then
        Texto1 = ""
        Text1(2).Text = ""
        If Dir(C, vbArchive) <> "" Then
            Open C For Input As #NF
            C = LeerLinea(NF)
            Texto1 = C
            C = LeerLinea(NF)
            Text1(2).Text = C
            Close #NF
        End If

    Else
        'Guardar
        Open C For Output As #NF
        Print #NF, Texto1
        Print #NF, Text1(2).Text
        Close #NF
    End If

    Exit Sub
ELeerLinea:
    MsgBox Err.Description, vbExclamation
End Sub

Private Sub Image1_Click()
    cd1.CancelError = False
    cd1.ShowOpen
    If cd1.FileName <> "" Then Text1(1).Text = cd1.FileName
End Sub

Private Sub Text1_Change(Index As Integer)
    If Index <> 1 Then Cambiado = True
End Sub
