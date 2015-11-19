VERSION 5.00
Begin VB.Form frmCaracteresMB 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Revisión Caracteres Multibase"
   ClientHeight    =   5850
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6090
   Icon            =   "frmCaracteresMB.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5850
   ScaleWidth      =   6090
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame frameMultibase 
      BorderStyle     =   0  'None
      Height          =   5775
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6015
      Begin VB.CheckBox chkMultibase 
         Caption         =   "Lineas"
         Height          =   255
         Index           =   2
         Left            =   720
         TabIndex        =   11
         Top             =   3690
         Value           =   1  'Checked
         Width           =   2145
      End
      Begin VB.CommandButton cmdMultiBase 
         Caption         =   "Aceptar"
         Height          =   375
         Index           =   0
         Left            =   3480
         TabIndex        =   4
         Top             =   5160
         Width           =   1095
      End
      Begin VB.CommandButton cmdMultiBase 
         Caption         =   "Salir"
         Height          =   375
         Index           =   1
         Left            =   4680
         TabIndex        =   3
         Top             =   5160
         Width           =   1095
      End
      Begin VB.CheckBox chkMultibase 
         Caption         =   "Avnics"
         Height          =   255
         Index           =   0
         Left            =   720
         TabIndex        =   2
         Top             =   2640
         Value           =   1  'Checked
         Width           =   2055
      End
      Begin VB.CheckBox chkMultibase 
         Caption         =   "Movimientos"
         Height          =   255
         Index           =   1
         Left            =   720
         TabIndex        =   1
         Top             =   3180
         Value           =   1  'Checked
         Width           =   2145
      End
      Begin VB.Label Label29 
         Caption         =   "Revisión caracteres multibase"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   375
         Left            =   720
         TabIndex        =   10
         Top             =   120
         Width           =   4935
      End
      Begin VB.Label Label30 
         Caption         =   "Utlidad para revisar los caracteres especiales que puedan quedar al realizar integraciones. "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   360
         TabIndex        =   9
         Top             =   720
         Width           =   5775
      End
      Begin VB.Label Label31 
         Caption         =   "No debe trabajar nadie en la aplicación"
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
         Left            =   360
         TabIndex        =   8
         Top             =   1320
         Width           =   4815
      End
      Begin VB.Label Label32 
         Caption         =   "A este proceso le puede costar mucho tiempo."
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
         Left            =   360
         TabIndex        =   7
         Top             =   1680
         Width           =   4815
      End
      Begin VB.Label Label33 
         Caption         =   "Datos a revisar:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   2280
         Width           =   4815
      End
      Begin VB.Label Label34 
         Caption         =   "Label34"
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   4800
         Width           =   5535
      End
      Begin VB.Line Line5 
         X1              =   240
         X2              =   5640
         Y1              =   4140
         Y2              =   4140
      End
   End
End
Attribute VB_Name = "frmCaracteresMB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim SQL As String

Private Sub cmdMultiBase_Click(Index As Integer)
Dim i As Integer
    If Index = 1 Then
        Unload Me
        Exit Sub
    End If
    
    
    'Comprobamos k ha selecionado algun nivel
    NE = 0
    For i = 0 To Me.chkMultibase.Count - 1
        If Me.chkMultibase(i).Value = 1 Then NE = NE + 1
    Next i
    If NE = 0 Then
        MsgBox "Seleccione donde se van a realizar los cambios", vbExclamation
        Exit Sub
    End If
    
    'Comprobacion si hay alguien trabajando
    If UsuariosConectados Then Exit Sub
    
    SQL = "Seguro que desea continuar con el proceso"
    If MsgBox(SQL, vbCritical + vbYesNoCancel) <> vbYes Then Exit Sub
    
    Screen.MousePointer = vbHourglass
    NumRegElim = 0
    For i = 0 To Me.chkMultibase.Count - 1
        If Me.chkMultibase(i).Value = 1 Then
            'Hacemos los cambios para ese valor
            HacerCambios i
        End If
    Next i
'    Bloquear_DesbloquearBD False
    Screen.MousePointer = vbDefault
    Label34.Caption = ""
    SQL = "Proceso finalizado" & vbCrLf & "Se han realizado: " & NumRegElim & " cambio(s)."
    MsgBox SQL, vbInformation
End Sub
Private Sub Form_Activate()
    If PrimeraVez Then
        PrimeraVez = False
    Else
        Screen.MousePointer = vbDefault
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
Dim w, h
    PrimeraVez = True
    Me.frameMultibase.visible = False
    'MULTIBASE
    Me.Caption = "Sustitución caracteres multibase"
    w = Me.frameMultibase.Width
    h = Me.frameMultibase.Height + 300
    Me.frameMultibase.visible = True
    Label34.Caption = ""
    cmdMultiBase(1).Cancel = True
    Me.Width = w + 120
    Me.Height = h + 120
    
    For i = 0 To 2
        chkMultibase(i).Value = 0
    Next i
    
    chkMultibase(0).Enabled = (vParamAplic.Avnics = 1)
    chkMultibase(0).visible = (vParamAplic.Avnics = 1)
    chkMultibase(1).Enabled = (vParamAplic.Avnics = 1)
    chkMultibase(1).visible = (vParamAplic.Avnics = 1)
    chkMultibase(2).Enabled = (vParamAplic.Seguros = 1)
    chkMultibase(2).visible = (vParamAplic.Seguros = 1)
    
End Sub

Private Sub HacerCambios(ByVal Tabla As Integer)
Dim Cambio As String
Dim Inicio As Integer
Dim Fin As Integer
Dim cad As String

    'RevisaCaracterMultibase
    Select Case Tabla
    Case 0
        'Avnics
        SQL = "Select codavnic, anoejerc, nombrper, nomcalle, poblacio, provinci, "
        SQL = SQL & "nombper1, nomcall1, poblaci1, provinc1 "
        SQL = SQL & " FROM avnic"
        Inicio = 2
        Fin = 9
    Case 1
        'Movimientos
        SQL = "Select codavnic, anoejerc, fechamov, concepto FROM movim"
        Inicio = 3
        Fin = 3
    Case 2
        'Lineas de seguros
        SQL = "Select codlinea, nomlinea FROM seglinea"
        Inicio = 1
        Fin = 1
    End Select
    
    Set rs = New adodb.Recordset
    rs.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not rs.EOF Then
        
        While Not rs.EOF
            Label34.Caption = rs.Fields(0) & " - " & rs.Fields(1)
            Label34.Refresh
            Cambio = ""
            
            For i = Inicio To Fin
                'Campo no nulo
                If Not IsNull(rs.Fields(i)) Then
                    SQL = rs.Fields(i)
                    cad = RevisaCaracterMultibase(SQL)
                    If SQL <> cad Then
                        'Han habido cambios
                        If Cambio <> "" Then Cambio = Cambio & ","
'                        Sql = NombreSQL(Cad)
                        SQL = DevNombreSQL(cad)
                        NumRegElim = NumRegElim + 1
                        Cambio = Cambio & rs.Fields(i).Name & " = '" & SQL & "'"
                    End If
                End If
            Next i
            If Cambio <> "" Then
                'OK HAY K CAMBIAR, k updatear
                Select Case Tabla
                Case 0
                    SQL = "UPDATE avnic SET " & Cambio & " WHERE codavnic =" & rs.Fields(0) & " AND anoejerc =" & rs.Fields(1)
            
                Case 1
                    SQL = "UPDATE movim"
                    SQL = SQL & " SET " & Cambio & " WHERE codavnic = " & rs.Fields(0)
                    SQL = SQL & " AND anoejerc =" & rs.Fields(1)
                    SQL = SQL & " AND fechamov =" & DBSet(rs.Fields(2), "F")
                
                Case 2
                    SQL = "UPDATE seglinea"
                    SQL = SQL & " SET " & Cambio & " WHERE codlinea = " & rs.Fields(0)
                End Select
                
                'Ejecutamos
                conn.Execute SQL
            End If
            rs.MoveNext
        Wend
    End If
    rs.Close
    Set rs = Nothing
            
End Sub

