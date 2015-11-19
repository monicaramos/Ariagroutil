VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CSeccion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'variables locales que contienen valores de propiedad

Private mvarSeccion As Integer 'copia local
Private mvarNombreSec As String 'copia local
Private mvarContador As Long  'copia local
Private mvarBDConta As Integer 'copia local
Private mvarLetraSerie As String * 1 'Letra Serie Ariconta

Private RS As ADODB.Recordset
Private SQL As String
Private Mens As String


Public Property Let Seccion(ByVal vData As Integer)
    mvarSeccion = vData
End Property


Public Property Get Seccion() As Integer
    Seccion = mvarSeccion
End Property

Public Property Let NombreSeccion(ByVal vData As String)
    mvarNombreSec = vData
End Property


Public Property Get NombreSeccion() As String
    NombreSeccion = mvarNombreSec
End Property


Public Property Let Contador(ByVal vData As Long)
    mvarContador = vData
End Property


Public Property Get Contador() As Long
    Contador = mvarContador
End Property

Public Property Let LetraSerie(ByVal vData As String)
    mvarLetraSerie = vData
End Property


Public Property Get LetraSerie() As String
    LetraSerie = mvarLetraSerie
End Property

Public Property Let BdConta(ByVal vData As Integer)
    mvarBDConta = vData
End Property


Public Property Get BdConta() As Integer
    BdConta = mvarBDConta
End Property


'//////////////////////////////////////////////////////////////////7
'//
'//     Si es el ultimo numero de la serie devuelve el anterior
Public Function DevolverContador(vCodigo As String, Contador As Long) As Byte
Dim OK As Boolean
Dim C1 As Long
On Error GoTo Err1
    
    'Abrimos bloqueando
    SQL = "Select * from seccion WHERE codsecci = " & vCodigo & " FOR UPDATE " ' "' FOR UPDATE"
    'Esto esba asi antes, sin comentar
    'Conn.Execute "Set autocommit = 0"
    DevolverContador = 1
    OK = False
    Set RS = New ADODB.Recordset
    RS.Open SQL, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    If Not RS.EOF Then
        mvarSeccion = vCodigo
        mvarNombreSec = RS!nomsecci
        C1 = RS!Contador
        OK = True
    End If
    RS.Close
    Set RS = Nothing
    If Not OK Then
        conn.Execute "commit"
        Exit Function
    End If

    OK = False

    OK = (C1 = Contador)
    If OK Then
        'Actualizamos el contador
        SQL = "UPDATE seccion set "

        C1 = C1 - 1
        SQL = SQL & " contador=" & C1
        SQL = SQL & " WHERE codsecci = " & vCodigo
        conn.Execute SQL
    End If

    'Desbloqueamos
    conn.Execute "commit"
    DevolverContador = 1
    
Err1:
    If Err.Number <> 0 Then
        Mens = "Error: " & Err.Number & " : " & Err.Description
        MsgBox Mens, vbExclamation
        DevolverContador = 0
    End If
    conn.Execute "Set autocommit = 1"
End Function


Public Function ConseguirContador(vCodigo As String) As Long
Dim OK As Boolean
Dim C1 As Long 'contador
Dim RS As ADODB.Recordset
On Error GoTo Err1
    
    ConseguirContador = 1
    
    'Abrimos bloqueando
    SQL = "Select * from seccion WHERE codsecci=" & vCodigo & " FOR UPDATE"
    Set RS = New ADODB.Recordset
    RS.Open SQL, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    If Not RS.EOF Then
        mvarSeccion = vCodigo
        mvarNombreSec = RS!nomsecci
        mvarContador = RS!Contador
        C1 = RS!Contador
        OK = True
    End If
    RS.Close
    Set RS = Nothing
    
    If OK Then ConseguirContador = C1 + 1
    Exit Function
Err1:
    Mens = "Error: " & Err.Number & " : " & Err.Description
    MsgBox Mens, vbExclamation
    ConseguirContador = -1
End Function


Public Function IncrementarContador(vCodigo As String) As Boolean
Dim C1 As Long 'contador1: recibo de entrada
On Error GoTo Err1
    
      
    'Actualizamos el contador
    C1 = Contador
    SQL = "UPDATE seccion set "
    C1 = C1 + 1
    mvarContador = C1
    SQL = SQL & " contador=" & C1
    
    SQL = SQL & " WHERE codsecci=" & vCodigo
    conn.Execute SQL
    IncrementarContador = True
    Exit Function
Err1:
    If Err.Number <> 0 Then
        IncrementarContador = False
    Else
        IncrementarContador = True
    End If
End Function



Public Function Leer(vCodigo As String) As Boolean
On Error GoTo Err3
    
    SQL = "SELECT * FROM seccion WHERE codsecci=" & vCodigo
    Set RS = New ADODB.Recordset
    RS.Open SQL, conn, adOpenForwardOnly, adLockOptimistic

    If RS.EOF Then
        Leer = False
        'No existe la sección
        SQL = "No existe la Sección: " & vCodigo
        MsgBox SQL, vbCritical
    Else
        mvarSeccion = RS!codsecci
        mvarNombreSec = RS!nomsecci
        mvarContador = RS!Contador
        mvarLetraSerie = DBLet(Trim(RS!letraser), "T")
        mvarBDConta = RS!NumConta
        Leer = True
    End If

    RS.Close
    Set RS = Nothing
    Exit Function

Err3:
    If Err.Number <> 0 Then
        Mens = "Se ha producido un error." & vbCrLf
        Mens = Mens & "Número: " & Err.Number & vbCrLf
        Mens = Mens & "Descripción: " & Err.Description
        MsgBox Mens, vbExclamation
        RS.Close
        Leer = False
    End If
End Function
