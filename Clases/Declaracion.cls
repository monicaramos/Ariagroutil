VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Declaracion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Incluye una declaración completa a la agencia tributaria de diferentes suministros"
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
'variables locales para almacenar los valores de las propiedades
Private mvarid As String 'copia local
Private mvarcodee As String 'copia local
Private mvartest As Variant 'copia local
Private mvarsituacion As Byte 'copia local
'variables locales para almacenar los valores de las propiedades
Private mvarsuministros As ColSuministro 'copia local
Private sql As String
Private RS As ADODB.Recordset


Public Property Set suministros(ByVal vData As ColSuministro)
Attribute suministros.VB_Description = "Suministros que componen una declaración"
'se usa al asignar un objeto a la propiedad, en la parte izquierda de una instrucción Set.
'Syntax: Set x.suministros = Form1
    Set mvarsuministros = vData
End Property


Public Property Get suministros() As ColSuministro
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.suministros
    Set suministros = mvarsuministros
End Property

Public Property Let situacion(ByVal vData As Byte)
Attribute situacion.VB_Description = "Posibles valores: 0 = Transacción creada pero no enviada, 1 = Transacción enviada a la agencia tributaria."
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.situacion = 5
    mvarsituacion = vData
End Property

Public Property Get situacion() As Byte
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.situacion
    situacion = mvarsituacion
End Property

Public Property Let test(ByVal vData As Variant)
Attribute test.VB_Description = "Si 'S' indica que la transacción es tipo test, es decir, para probar simplemente"
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.test = 5
    mvartest = vData
End Property

Public Property Set test(ByVal vData As Variant)
'se usa al asignar un objeto a la propiedad, en la parte izquierda de una instrucción Set.
'Syntax: Set x.test = Form1
    Set mvartest = vData
End Property

Public Property Get test() As Variant
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.test
    If IsObject(mvartest) Then
        Set test = mvartest
    Else
        test = mvartest
    End If
End Property

Public Property Let codee(ByVal vData As String)
Attribute codee.VB_Description = "Identificador de la entidad emisora de la tarjeta"
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.codee = 5
    mvarcodee = vData
End Property

Public Property Get codee() As String
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.codee
    codee = mvarcodee
End Property

Public Property Let id(ByVal vData As String)
Attribute id.VB_Description = "Identificador de transaccion en la forma AAAAMMDDHHMMSSNNNNNN"
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.id = 5
    mvarid = vData
End Property

Public Property Get id() As String
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.id
    id = mvarid
End Property

Public Function leer(id As String, conn As ADODB.Connection) As Boolean
    Dim vSuministro As Suministro
On Error GoTo err_leer
    sql = "select a.id, a.codee, a.test," & _
            " b.idmovcont, b.cim, b.fechahora, b.codprod, b.lit, b.nif, b.matricula" & _
            " from gp_suministrv2ent as a, gp_suministro as b" & _
            " where a.id = '" & id & "'" & _
            " and b.id = a.id"
    Set RS = New ADODB.Recordset
    RS.Open sql, conn, adOpenForwardOnly
    If Not RS.EOF Then
        RS.MoveFirst
        '-- cargamos los atributos principales
        mvarid = RS!id
        mvarcodee = RS!codee
        mvartest = RS!test
        Set mvarsuministros = New ColSuministro
        While Not RS.EOF
            '-- Cargamos la colección de suministros
            Set vSuministro = New Suministro
            With vSuministro
                .id = RS!id
                .idmovcont = RS!idmovcont
                .cim = RS!cim
                .fecha = CDate(Format(RS!fechahora, "dd/mm/yyyy"))
                .hora = CDate(Format(RS!fechahora, "hh:mm:ss"))
                .codpro = RS!codprod
                .lit = RS!lit
                .nif = RS!nif
                .matricula = RS!matricula
            End With
            mvarsuministros.Add vSuministro, CStr(vSuministro.idmovcont)
            RS.MoveNext
        Wend
        leer = True
    Else
        leer = False
    End If
    Exit Function
err_leer:
    MsgBox Err.Description & " (" & Err.Number & ")", vbCritical, "Declaracion.leer"
End Function

Public Function guardar(conn As ADODB.Connection) As Boolean
    Dim vSuministro As Suministro
On Error GoTo err_guardar
    '-- Generamos el valor de ID
    mvarid = Format(Now, "yyyymmddhhnnss000001")
    '-- Obtenemos el CEE
    If Not obtenerCEE(conn) Then
        MsgBox "No se puede obtener el CEE. Revise parámetros", vbExclamation
        Exit Function
    End If
    sql = "delete from gp_suministrv2ent where id = '" & mvarid & "'"
    conn.Execute sql
    sql = "insert into gp_suministrv2ent (id,codee,test,situacion)" & _
            " values (" & _
            "'" & mvarid & "'," & _
            "'" & mvarcodee & "'," & _
            "'" & mvartest & "'," & _
            "0" & _
            ")"
    conn.Execute sql ' ya tenemos la cabecera
    For Each vSuministro In mvarsuministros
        '-- Confirmamos la clave principal
        vSuministro.id = mvarid
        If Not vSuministro.guardar(conn, True) Then
            Exit Function
        End If
    Next
    guardar = True
    Exit Function
err_guardar:
    MsgBox Err.Description & " (" & Err.Number & ")", vbCritical, "Declaracion.guardar"
End Function

Public Function eliminar(conn As ADODB.Connection) As Boolean
On Error GoTo err_eliminar
    sql = "delete from gp_suministro where id = '" & mvarid & "'"
    conn.Execute sql
    sql = "delete from gp_suministrv2ent where id = '" & mvarid & "'"
    conn.Execute sql
    eliminar = True
    Exit Function
err_eliminar:
    MsgBox Err.Description & " (" & Err.Number & ")", vbCritical, "Declaracion.eliminar"
End Function

Private Function obtenerCEE(conn As ADODB.Connection) As Boolean
On Error GoTo err_obtenerCEE
    sql = "select cee from sparam"
    Set RS = New ADODB.Recordset
    RS.Open sql, conn, adOpenForwardOnly
    If Not RS.EOF Then
        mvarcodee = RS!Cee
        obtenerCEE = True
    Else
        obtenerCEE = False
    End If
    Exit Function
err_obtenerCEE:
    MsgBox Err.Description & " (" & Err.Number & ")", vbCritical, "Declaracion.obtenerCEE"
End Function

Public Function exportarXML() As Boolean
    Dim NF As Integer
    Dim lin As String
    Dim mSuministro As Suministro
On Error GoTo err_exportarXML
    '-- comprobamos que hay algo para exportar
    If mvarsuministros Is Nothing Then Exit Function
    If mvarsuministros.Count = 0 Then Exit Function
    '-- obtenemos el fichero en el que grabaremos el XML
    NF = FreeFile()
    Open App.path & "\D" & mvarid & ".XML" For Output As NF
    '-- comenzamos a grabar los datos de cabecera
        Print #NF, "<?xml version=""1.0"" encoding=""UTF-8""?>"
        Print #NF, "<!--Generated by Avnics (c) Ariadna Software S.L. (http://www.ariadnasoftware.com)-->"
        lin = "<SuministrV2Ent CodEE=""" & mvarcodee & """ Id= """ & mvarid & """"
        If mvartest = "S" Then lin = lin & " Test=""S"""
        lin = lin & " xsi:schemaLocation=""https://www3.aeat.es/ADUA/internet/es/aeat/dit/adu/iegasolprof/SuministrV2Ent.xsd https://www3.aeat.es/ADUA/internet/es/aeat/dit/adu/iegasolprof/SuministrV2Ent.xsd"""
        lin = lin & " xmlns=""https://www3.aeat.es/ADUA/internet/es/aeat/dit/adu/iegasolprof/SuministrV2Ent.xsd"""
        lin = lin & " xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"">"
        Print #NF, lin
        Set mSuministro = New Suministro
        For Each mSuministro In mvarsuministros
            Print #NF, "<Suministro>"
                With mSuministro
                    Print #NF, "<IdMovCont>" & CStr(.idmovcont) & "</IdMovCont>"
                    Print #NF, "<CIM>" & .cim & "</CIM>"
                    Print #NF, "<Fecha>" & Format(.fecha, "yyyymmdd") & "</Fecha>"
                    Print #NF, "<Hora>" & Format(.hora, "hhmm") & "</Hora>"
                    Print #NF, "<CodPro>" & Format(.codpro, "000") & "</CodPro>"
                    Print #NF, "<Lit>" & TransformaComasPuntos(Format(.lit, "0.00")) & "</Lit>"
                    Print #NF, "<NIF>" & .nif & "</NIF>"
                    Print #NF, "<Matricula>" & .matricula & "</Matricula>"
                End With
            Print #NF, "</Suministro>"
        Next
        Print #NF, "</SuministrV2Ent>"
    Close #NF
    exportarXML = True
    Exit Function
err_exportarXML:
    MsgBox Err.Description & " (" & Err.Number & ")", vbCritical, "Declaracion.exportarXML"
End Function
