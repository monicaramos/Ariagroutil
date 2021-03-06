VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CContador"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'variables locales que contienen valores de propiedad

'Private mTipoContador As String 'copia local
'Private mvarNombreRegistro As String 'copia local

'Contadores a�o actual
'---------------------
'Private mContadA As Integer 'A�o actual
'Private mPVentaA As Single  'contador preventa
'Private mExpIndA As Single  'contador expedientes individuales a�o actual
'Private mExpGrpA As Single  'contador expedientes grupos a�o actual


'Contadores a�o siguiente
'------------------------
'Private mContadB As Integer 'A�o siguiente
'Private mPVentaB As Single  'contador preventa a�o siguiente
'Private mExpIndB As Single  'contador expedientes individuales a�o siguiente
'Private mExpGrpB As Single  'contador expedientes grupos a�o siguiente
'


Private mContador As Long

Private Rs As ADODB.Recordset
Private Sql As String
Private Mens As String


'contador
Public Property Let Contador(ByVal vData As Long)
    mContador = vData
End Property


Public Property Get Contador() As Long
    Contador = mContador
End Property


'///////////////////////////////////////////////////////////
'//
'//  Devuelve el contador que corresponda.
'//  Si es llamado desde dentro de un begin-trans
'//  entonces NO ejecuta las sentencias commit
Public Function ConseguirContador(tipo As String, DentroBeginTrans As Boolean, Optional ByRef db As BaseDatos) As Boolean
Dim OK As Boolean
Dim c1 As Currency
Dim c2 As Currency

    On Error GoTo Err1

    'Abrimos bloqueando
    Sql = "Select contador, letraser from stipom "
    Sql = Sql & " WHERE codtipom = '" & DBLet(tipo, "T") & "' FOR UPDATE"

    ConseguirContador = False
    OK = False

    If IsNull(db) Then
        Set Rs = New ADODB.Recordset
        Rs.Open Sql, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    Else
        Set Rs = db.cursor(Sql)
    End If
    
    If Not Rs.EOF Then
        c1 = Rs!Contador + 1
        mContador = c1
        OK = True
    Else
        MsgBox "No existe contador para ese tipo de movimiento.", vbExclamation
        OK = False
    End If
    
    Rs.Close
    Set Rs = Nothing

    If IsNull(db) Then
        If Not OK Then
            If Not DentroBeginTrans Then conn.Execute "commit"
            Exit Function
        End If

        'Actualizamos el contador
        Sql = "UPDATE stipom set contador = " & DBLet(c1, "N")
        Sql = Sql & " WHERE codtipom = '" & DBLet(tipo, "T") & "'"
        conn.Execute Sql
    
        'Desbloqueamos
        If Not DentroBeginTrans Then conn.Execute "commit"
        ConseguirContador = True
        Exit Function

    Else
    
        If Not OK Then
            If Not DentroBeginTrans Then db.Ejecutar "commit"
            Exit Function
        End If

        'Actualizamos el contador
        Sql = "UPDATE stipom set contador = " & DBLet(c1, "N")
        Sql = Sql & " WHERE codtipom = '" & DBLet(tipo, "T") & "'"
        db.Ejecutar Sql
    
        'Desbloqueamos
        If Not DentroBeginTrans Then db.Ejecutar "commit"
        ConseguirContador = True
        Exit Function
    
    End If
Err1:
    Mens = "Error: " & Err.Number & " : " & Err.Description
    MsgBox Mens, vbExclamation
End Function

'Public Function ConseguirContador(tipo As String, DentroBeginTrans As Boolean) As Byte
'Dim Ok As Boolean
'Dim c1 As Long
'Dim c2 As Long
'    On Error GoTo Err1
'    'Abrimos bloqueando
'    sql = "Select contador, letraser from stipom "
'    sql = sql & " WHERE codtipom = '" & DBLet(tipo, "T") & "' FOR UPDATE"
'
'    ConseguirContador = 1
'    Ok = False
'    Set RS = New ADODB.Recordset
'    RS.Open sql, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
'    If Not RS.EOF Then
'        c1 = RS!Contador
'        Ok = True
'    End If
'    RS.Close
'    If Not Ok Then
'        If Not DentroBeginTrans Then conn.Execute "commit"
'        Exit Function
'    End If
'
'    'Actualizamos el contador
'    sql = "UPDATE stipom set "
'    mContador = c1
'    c1 = c1 + 1
'    sql = sql & " contador=" & c1
'    sql = sql & " WHERE codtipom = '" & tipo & "' "
'
'    conn.Execute sql
'
'    'Desbloqueamos
'    If Not DentroBeginTrans Then conn.Execute "commit"
'    ConseguirContador = 0
'    Exit Function
'Err1:
'    Mens = "Error: " & Err.Number & " : " & Err.Description
'    MsgBox Mens, vbExclamation
'End Function


'Public Function ConseguirContador(Fecha As String, campo1 As String, campo2 As String, DentroBeginTrans As Boolean) As Boolean
'Dim Ok As Boolean
'Dim C1 As Currency
'Dim C2 As Currency
'
'    On Error GoTo Err1
'
'    'Abrimos bloqueando
''    SQL = "Select * from contadores WHERE TipoRegi = '" & LEtra & "' FOR UPDATE"
'    sql = "Select contad_a, contad_b, " & campo1 & " as contado1, " & campo2 & " as contado2 from agencias "
'    sql = sql & " WHERE codempre = " & vSesion.Empresa & " and codagenc=" & vSesion.Agencia & " FOR UPDATE"
'
'    ConseguirContador = False
'    Ok = False
'
'    Set RS = New ADODB.Recordset
'    RS.Open sql, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
'
'    If Not RS.EOF Then
''        mvarTipoContador = LEtra
''        mvarNombreRegistro = RS!nomregis
'        C1 = RS!Contado1
'        C2 = RS!Contado2
'        If Year(Fecha) = RS!contad_a Then 'a�o actual
'            mAnyoAct = True
'            Ok = True
'        ElseIf Year(Fecha) = RS!contad_b Then 'a�o siguiente
'            mAnyoAct = False
'            Ok = True
'        Else
'            Ok = False
'            MsgBox "No existe contador de n� de venta para ese a�o.", vbExclamation
'        End If
'
'    End If
'    RS.Close
'    Set RS = Nothing
'
'    If Not Ok Then
'        If Not DentroBeginTrans Then conn.Execute "commit"
'        Exit Function
'    End If
'
'    'Actualizamos el contador
'    sql = "UPDATE agencias set "
'
'    If AnyoActual Then
'        C1 = C1 + 1
'        mContador1 = C1
'        'c1 = c1 + 1 'ANTES
'        sql = sql & campo1 & "=" & C1
'    Else
'        C2 = C2 + 1
'        mContador2 = C2
'        sql = sql & campo2 & "=" & C2
'    End If
'
'    sql = sql & " WHERE codempre = " & vSesion.Empresa & " and codagenc = " & vSesion.Agencia
'    conn.Execute sql
'
'    'Desbloqueamos
'    If Not DentroBeginTrans Then conn.Execute "commit"
'    ConseguirContador = True
'    Exit Function
'
'Err1:
'    Mens = "Error: " & Err.Number & " : " & Err.Description
'    MsgBox Mens, vbExclamation
'End Function



'//////////////////////////////////////////////////////////////////7
'//
'//     Si es el ultimo numero de la serie devuelve el anterior
Public Function DevolverContador(tipo As String, Contador As Long, Optional db As BaseDatos) As Boolean
Dim OK As Boolean
Dim c1 As Long
Dim c2 As Long
    On Error GoTo Err1
    'Abrimos bloqueando
    Sql = "Select * from stipom WHERE codtipom = '" & tipo & "' FOR UPDATE"
    'Esto esba asi antes, sin comentar
    'Conn.Execute "Set autocommit = 0"
    DevolverContador = False
    OK = False
    If IsNull(db) Then
        Set Rs = New ADODB.Recordset
        Rs.Open Sql, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    Else
        Set Rs = db.cursor(Sql)
    End If
    If Not Rs.EOF Then
        c1 = Rs!Contador
        OK = True
    End If
    
    Rs.Close
    
    If Not OK Then
        If IsNull(db) Then
            conn.Execute "commit"
        Else
            db.Ejecutar "commit"
        End If
        Exit Function
    End If

    OK = False
    'Contador = Contador + 1
    OK = (c1 = Contador)

    If OK Then
        'Actualizamos el contador
        Sql = "UPDATE stipom set "
        c1 = c1 - 1
        Sql = Sql & " contador=" & c1
        Sql = Sql & " WHERE codtipom = '" & DBLet(tipo, "T") & "'"
        If IsNull(db) Then
            conn.Execute Sql
        Else
            db.Ejecutar Sql
        End If
    End If

    'Desbloqueamos
    If IsNull(db) Then
        conn.Execute "commit"
    Else
        db.Ejecutar "commit"
    End If
    DevolverContador = True
Err1:
    If Err.Number <> 0 Then
        Mens = "Error: " & Err.Number & " : " & Err.Description
        MsgBox Mens, vbExclamation
    End If
    If IsNull(db) Then
        conn.Execute "Set autocommit = 1"
    Else
        db.Ejecutar "Set autocommit = 1"
    End If
End Function

