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

Private mvarESMaterna As Integer 'es materna


Private Rs As ADODB.Recordset
Private Sql As String
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


Public Property Let EsMaterna(ByVal vData As Integer)
    mvarESMaterna = vData
End Property


Public Property Get EsMaterna() As Integer
    EsMaterna = mvarESMaterna
End Property



'//////////////////////////////////////////////////////////////////7
'//
'//     Si es el ultimo numero de la serie devuelve el anterior
Public Function DevolverContador(vcodigo As String, Contador As Long) As Byte
Dim OK As Boolean
Dim c1 As Long
On Error GoTo Err1
    
    'Abrimos bloqueando
    Sql = "Select * from seccion WHERE codsecci = " & vcodigo & " FOR UPDATE " ' "' FOR UPDATE"
    'Esto esba asi antes, sin comentar
    'Conn.Execute "Set autocommit = 0"
    DevolverContador = 1
    OK = False
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    If Not Rs.EOF Then
        mvarSeccion = vcodigo
        mvarNombreSec = Rs!nomsecci
        c1 = Rs!Contador
        OK = True
    End If
    Rs.Close
    Set Rs = Nothing
    If Not OK Then
        conn.Execute "commit"
        Exit Function
    End If

    OK = False

    OK = (c1 = Contador)
    If OK Then
        'Actualizamos el contador
        Sql = "UPDATE seccion set "

        c1 = c1 - 1
        Sql = Sql & " contador=" & c1
        Sql = Sql & " WHERE codsecci = " & vcodigo
        conn.Execute Sql
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


Public Function ConseguirContador(vcodigo As String) As Long
Dim OK As Boolean
Dim c1 As Long 'contador
Dim Rs As ADODB.Recordset
On Error GoTo Err1
    
    ConseguirContador = 1
    
    'Abrimos bloqueando
    Sql = "Select * from seccion WHERE codsecci=" & vcodigo & " FOR UPDATE"
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    If Not Rs.EOF Then
        mvarSeccion = vcodigo
        mvarNombreSec = Rs!nomsecci
        mvarContador = Rs!Contador
        c1 = Rs!Contador
        OK = True
    End If
    Rs.Close
    Set Rs = Nothing
    
    If OK Then ConseguirContador = c1 + 1
    Exit Function
Err1:
    Mens = "Error: " & Err.Number & " : " & Err.Description
    MsgBox Mens, vbExclamation
    ConseguirContador = -1
End Function


Public Function IncrementarContador(vcodigo As String) As Boolean
Dim c1 As Long 'contador1: recibo de entrada
On Error GoTo Err1
    
      
    'Actualizamos el contador
    c1 = Contador
    Sql = "UPDATE seccion set "
    c1 = c1 + 1
    mvarContador = c1
    Sql = Sql & " contador=" & c1
    
    Sql = Sql & " WHERE codsecci=" & vcodigo
    conn.Execute Sql
    IncrementarContador = True
    Exit Function
Err1:
    If Err.Number <> 0 Then
        IncrementarContador = False
    Else
        IncrementarContador = True
    End If
End Function



Public Function leer(vcodigo As String) As Boolean
On Error GoTo Err3
    
    Sql = "SELECT * FROM seccion WHERE codsecci=" & vcodigo
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockOptimistic

    If Rs.EOF Then
        leer = False
        'No existe la sección
        Sql = "No existe la Sección: " & vcodigo
        MsgBox Sql, vbCritical
    Else
        mvarSeccion = Rs!codsecci
        mvarNombreSec = Rs!nomsecci
        mvarContador = Rs!Contador
        mvarLetraSerie = DBLet(Trim(Rs!letraser), "T")
        mvarBDConta = Rs!NumConta
        mvarESMaterna = Rs!EsMaterna
        leer = True
    End If

    Rs.Close
    Set Rs = Nothing
    Exit Function

Err3:
    If Err.Number <> 0 Then
        Mens = "Se ha producido un error." & vbCrLf
        Mens = Mens & "Número: " & Err.Number & vbCrLf
        Mens = Mens & "Descripción: " & Err.Description
        MsgBox Mens, vbExclamation
        Rs.Close
        leer = False
    End If
End Function


