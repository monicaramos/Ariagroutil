VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CProveedor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'Variables locales que contienen valores de propiedad
Private mCodProve As String 'codigo
Private mNomComer As String 'nombre comercial
Private mNIFProve As String 'NIF del proveedor

Private mPContCom As String 'persona de contacto comercial
Private mTfnoCom1 As String 'telefono comercial 1
Private mWebCom As String 'web comercial
Private mMailCom As String 'email comercial




'------------------------------------------------
'Propiedades del modulo CProveedor
'------------------------------------------------
'Codigo proveedor
Public Property Let CodProve(ByVal vData As String)
     mCodProve = vData
End Property

Public Property Get CodProve() As String
     CodProve = mCodProve
End Property


'Nombre comercial
Public Property Let NomComer(ByVal vData As String)
     mNomComer = vData
End Property

Public Property Get NomComer() As String
     NomComer = mNomComer
End Property

'NIF proveedor
Public Property Let NIFProve(ByVal vData As String)
     mNIFProve = vData
End Property

Public Property Get NIFProve() As String
     NIFProve = mNIFProve
End Property




'Persona de contacto comercial
Public Property Let PersConCom(ByVal vData As String)
     mPContCom = vData
End Property

Public Property Get PersConCom() As String
     PersConCom = mPContCom
End Property

'Telefono comercial 1
Public Property Let TelefCom1(ByVal vData As String)
     mTfnoCom1 = vData
End Property

Public Property Get TelefCom1() As String
     TelefCom1 = mTfnoCom1
End Property


'Web de contacto comercial
Public Property Let WebCom(ByVal vData As String)
     mWebCom = vData
End Property

Public Property Get WebCom() As String
     WebCom = mWebCom
End Property


'Mail de contacto comercial
Public Property Let EMailCom(ByVal vData As String)
     mMailCom = vData
End Property

Public Property Get EMailCom() As String
     EMailCom = mMailCom
End Property





'****   LEER  ***
Public Function LeerDatos(codigo As String) As Byte
'Leer los datos del proveedor
'Lee de la BD: PlannerTours, Tabla: Proveedo

    Dim RS As ADODB.Recordset
    Dim SQL As String, mens As String

    On Error GoTo ELeer
    
    SQL = "SELECT * FROM proveedo WHERE codprove=" & codigo 'vUsu.CadenaConexion

    Set RS = New ADODB.Recordset
    RS.Open SQL, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    mens = ""
    If RS.EOF Then
        LeerDatos = 0
    Else
        mCodProve = RS!CodProve
        mNomComer = RS!NomComer
        mNIFProve = DBLet(RS!NIFProve, "T")
        
        mPContCom = DBLet(RS!pcontcom, "T")
        mTfnoCom1 = DBLet(RS!TfnoCom1, "T")
        mWebCom = DBLet(RS!webcomer, "T")
        mMailCom = DBLet(RS!EMailCom, "T")
        
        mens = "Datos del Proveedor"
        LeerDatos = 1
    End If

    RS.Close
    Set RS = Nothing
    Exit Function

ELeer:
    mens = "Se ha producido un error. " & mens & vbCrLf
    mens = mens & "Número: " & Err.Number & vbCrLf
    mens = mens & "Descripción: " & Err.Description
    MsgBox mens, vbExclamation
    Set RS = Nothing
    LeerDatos = 1
End Function





Public Function TieneNoConf() As Boolean
'Comprobar si el proveedor tiene No Conformidades para la empresa conectada

    Dim RS As ADODB.Recordset
    Dim SQL As String

    On Error GoTo ErrNoC
    
    SQL = "SELECT COUNT(*) FROM noconcab WHERE codprove=" & CodProve & " AND codempre=" & vSesion.Empresa
    
    Set RS = New ADODB.Recordset
    RS.Open SQL, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    If Not RS.EOF Then
        If RS.Fields(0).Value > 0 Then TieneNoConf = True
    End If
    Set RS = Nothing

    Exit Function
    
ErrNoC:
    MuestraError Err.Number, "Comprobar no conformidades proveedor.", Err.Description
End Function
