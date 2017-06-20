Attribute VB_Name = "bus"
'NOTA: en este mòdul, ademés, n'hi han funcions generals que no siguen de formularis (molt bé)
Option Explicit

'Definicion Conexión a BASE DE DATOS
'---------------------------------------------------
'Conexión a la BD Avnics de la empresa
Public conn As ADODB.Connection

'Conexión a la BD de Usuarios
'Public ConnUsuarios As ADODB.Connection

'Conexión a la BD de Contabilidad de la empresa conectada
Public ConnConta As ADODB.Connection ' avnics
Public ConnContaSeg As ADODB.Connection ' seguros
Public ConnContaTel As ADODB.Connection ' telefonia
Public ConnContaFac As ADODB.Connection ' facturas varias
Public ConnContaGas As ADODB.Connection ' gasolinera
Public ConnContaFacSoc As ADODB.Connection ' facturas socios
Public ConnContaCV As ADODB.Connection ' facturas coarval
Public ConnContaCVV As ADODB.Connection ' facturas coarval varias

'Conexión a la BD de Contabilidad de otra empresa distinta a la conectada
Public ConnAuxCon As ADODB.Connection

'Conexión a la BD de Aridoc de la empresa conectada
Public ConnAridoc As ADODB.Connection


'Que conexion a base de datos se va a utilizar
Public Const cPTours As Byte = 1 'trabajaremos con conn (conexion a BD Avnics)
Public Const cConta As Byte = 2 'trabajaremos con connConta (cxion a BD Contabilidad)
Public Const cContaSeg As Byte = 3 'trabajaremos con connContaSeg (cxion a BD Contabilidad para Seguros)
Public Const cContaTel As Byte = 4 'trabajaremos con connContaTel (cxion a BD Contabilidad para Telefonia)
Public Const cContaFac As Byte = 5 'trabajaremos con connContaFac (cxion a BD Contabilidad para Facturas Varias)
Public Const cContaGas As Byte = 6 'trabajaremos con connContaGas (cxion a BD Contabilidad para Gasolinera)
Public Const cContaFacSoc As Byte = 7 'trabajaremos con connContaFacSoc (cxion a BD Contabilidad para Facturas Socios)
Public Const cContaCV As Byte = 8 'trabajaremos con connContaCV (cxion a BD Contabilidad para Facturas Coarval)
Public Const cContaCVV As Byte = 9 'trabajaremos con connContaCV (cxion a BD Contabilidad para Facturas Coarval varias)

Public Const cAridoc As Byte = 10 'trabajaremos con connAridoc (cxion a BD Aridoc)

Public ardDB As BaseDatos ' este es la base de datos que soportará aridoc

'Definicion de clases de la aplicación
'-----------------------------------------------------
Public vEmpresa As Cempresa  'Los datos de la empresa
Public vEmpresaSeg As CempresaSeg  'Los datos de la empresa
Public vEmpresaTel As CempresaTel  'Los datos de la empresa
Public vEmpresaFac As CempresaFac  'Los datos de la empresa
Public vEmpresaGas As CempresaGas  'Los datos de la empresa
Public vEmpresaFacSoc As CempresaFacSoc  'Los datos de la empresa
Public vEmpresaCV As CempresaCV  'Los datos de la empresa
Public vEmpresaCVV As CempresaCVV  'Los datos de la empresa


Public vParamAplic As CParamAplic   'parametros de la aplicacion
Public vSesion As CSesion   'Los datos del usuario que hizo login



'Definicion de FORMATOS
'---------------------------------------------------
Public FormatoFecha As String
Public FormatoHora As String
Public FormatoImporte As String 'Decimal(12,2)
Public FormatoPrecio As String 'Decimal(8,3)
'Public FormatoCantidad As String 'Decimal(10,2)
Public FormatoPorcen As String 'Decimal(5,2) 'Porcentajes
Public FormatoExp As String  'Expedientes

Public FormatoDec10d2 As String 'Decimal(10,2)
Public FormatoDec10d3 As String 'Decimal(10,3)
Public FormatoDec5d4 As String 'Decimal(5,4)
Public FormatoDec10d4 As String 'Decimal(10,4)

Public FIni As String
Public FFin As String

Public FIniSeg As String 'fecha de inicio de ejercicio de la contabilidad de Seguros
Public FFinSeg As String 'fecha de fin de ejercicio de la contabilidad de Seguros

Public FIniTel As String 'fecha de inicio de ejercicio de la contabilidad de Telefonia
Public FFinTel As String 'fecha de fin de ejercicio de la contabilidad de Telefonia

Public FIniGas As String 'fecha de inicio de ejercicio de la contabilidad de Gasolinera
Public FFinGas As String 'fecha de fin de ejercicio de la contabilidad de Gasolinera

Public FIniFacSoc As String 'fecha de inicio de ejercicio de la contabilidad de facturas de socio
Public FFinFacSoc As String 'fecha de fin de ejercicio de la contabilidad de facturas de socio

Public FIniCV As String 'fecha de inicio de ejercicio de la contabilidad de facturas de coarval
Public FFinCV As String 'fecha de fin de ejercicio de la contabilidad de facturas de coarval

Public FIniCVV As String 'fecha de inicio de ejercicio de la contabilidad de facturas de coarval
Public FFinCVV As String 'fecha de fin de ejercicio de la contabilidad de facturas de coarval



'Public FormatoKms As String 'Decimal(8,4)


Public teclaBuscar As Integer 'llamada desde prismaticos

Public CadenaDesdeOtroForm As String

'Global para nº de registro eliminado
Public NumRegElim  As Long

'publica para almacenar control cambios en registros de formularios
'se utiliza en InsertarCambios
Public CadenaCambio As String
Public ValorAnterior As String

Public MensError As String

'Para algunos campos de texto suletos controlarlos
'Public miTag As CTag

'Variable para saber si se ha actualizado algun asiento
'Public AlgunAsientoActualizado As Boolean
'Public TieneIntegracionesPendientes As Boolean

'Public miRsAux As ADODB.Recordset

Public AnchoLogin As String  'Para fijar los anchos de columna

Public Const SerieFraPro = "1"

'Inicio Aplicación
Public Sub Main()

    If App.PrevInstance Then
        MsgBox "Avnics ya se esta ejecutando", vbExclamation
        End
     End If
     
     'obric la conexio
    If AbrirConexionAvnics("root", "aritel") = False Then
        MsgBox "La aplicación no puede continuar sin acceso a los datos. ", vbCritical
        End
    End If

    Load frmIdentifica
    'CadenaDesdeOtroForm = ""
    
    'Necesitaremos el archivo login.dat
    frmIdentifica.Show
    
End Sub

Public Function ComprovaVersio() As Boolean
  
'    Dim RS2 As Recordset
'    Dim cad2 As String
'    Dim major_ul As Integer
'    Dim minor_ul As Integer
'    Dim revis_ul As Integer
'
'    ComprovaVersio = False
'
'    cad2 = "SELECT * FROM ulversio"
'
'    Set RS2 = New ADODB.Recordset
'    RS2.Open cad2, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
'
'    If Not RS2.EOF Then
'        major_ul = RS2.Fields!major_ul
'        minor_ul = RS2.Fields!minor_ul
'        revis_ul = RS2.Fields!revis_ul
'    Else
'        MsgBox "Error al consultar la última versión disponible", vbCritical
'        'ulVersio = False
'        Exit Function
'    End If
'
'    RS2.Close
'    Set RS2 = Nothing
'
'    If (App.Major <> major_ul) Or (App.Minor <> minor_ul) Or (App.Revision <> revis_ul) Then
'        ComprovaVersio = True
'    End If
'
'    Exit Function
    
End Function

'espera els segon que li digam
Public Function espera(Segundos As Single)
    Dim T1
    T1 = Timer
    Do
    Loop Until Timer - T1 > Segundos
End Function


Public Function AbrirConexionAvnics(Usuario As String, Pass As String) As Boolean
Dim Cad As String
On Error GoTo EAbrirConexion
    
    AbrirConexionAvnics = False
    Set conn = Nothing
    Set conn = New Connection
    'Conn.CursorLocation = adUseClient
    conn.CursorLocation = adUseServer
'    cad = "DSN=plannertours;DESC=MySQL ODBC 3.51 Driver DSN;DATABASE=plannertours;SERVER=srvcentral;UID=" & Usuario & ";PASSWORD=" & Pass & ";PORT=3306;OPTION=3;STMT=;"

'    cad = "DSN=arigasol;DESC=MySQL ODBC 3.51 Driver DSN;DATABASE=arigasol;UID=" & Usuario & ";PASSWORD=" & Pass & ";PORT=3306;OPTION=3;STMT=;"
'--monica
'    cad = "DSN=vAriagroutil;DESC=MySQL ODBC 3.51 Driver DSN;UID=" & Usuario & ";PASSWORD=" & Pass & ";PORT=3306;OPTION=3;STMT=;"
    
'++ de david
    Cad = "DRIVER={MySQL ODBC 3.51 Driver};DESC=;DATA SOURCE= vAriagroutil"
    Cad = Cad & ";UID=" & Usuario
    Cad = Cad & ";PWD=" & Pass
    Cad = Cad & ";Persist Security Info=true"
    
    conn.ConnectionString = Cad
    conn.Open
    AbrirConexionAvnics = True
    Exit Function
    
EAbrirConexion:
    MuestraError Err.Number, "Abrir conexión Avnics.", Err.Description
End Function


Public Function AbrirConexionConta(Usuario As String, Pass As String) As Boolean
Dim Cad As String
Dim nomConta As String 'nombre de la BD de la contabilidad
Dim serConta As String 'servidor donde esta la BD de la contabilidad
On Error GoTo EAbrirConexion
    
    AbrirConexionConta = False
    
    Set ConnConta = Nothing
    Set ConnConta = New Connection
'    Conn.CursorLocation = adUseClient   'Si ponemos este hay opciones k no van ej select con rs!campo
    ConnConta.CursorLocation = adUseServer   'Si ponemos esta alguns errores de Conn no se muestran correctamente

' ### [Monica] 06/09/2006
    If vParamAplic.ContabilidadNueva Then
        nomConta = "ariconta" & vParamAplic.NumeroConta
    Else
        nomConta = "conta" & vParamAplic.NumeroConta
    End If
'    vEmpresa.BDConta = nomConta
    If vParamAplic.NumeroConta <> 0 Then
    '    cad = "DRIVER={MySQL ODBC 3.51 Driver};DESC=;DATABASE=conta" & vParamConta.NumeroConta & ";SERVER=" & vParamConta.ServidorConta & ";"
    '    cad = cad & ";UID=" & vParamConta.UsuarioConta
    '    cad = cad & ";PWD=" & vParamConta.PasswordConta
    '    cad = "DSN=vConta;DESC=MySQL ODBC 3.51 Driver DSN;DATABASE=conta2;SERVER=david;UID=" & Usuario & ";PASSWORD=" & Pass & ";PORT=3306;OPTION=3;STMT=;"
'--monica: quitado por lo de david
'        If vParamAplic.ServidorConta <> "" Then  'especificamos servidor
'            cad = "DSN=vConta;DESC=MySQL ODBC 3.51 Driver DSN;DATABASE=" & nomConta & ";SERVER=" & vParamAplic.ServidorConta & ";UID=" & Usuario & ";PASSWORD=" & Pass & ";PORT=3306;OPTION=3;STMT=;"
'        Else 'por defecto cogera la BD del servidor que haya en el ODBC
'            cad = "DSN=vConta;DESC=MySQL ODBC 3.51 Driver DSN;DATABASE=" & nomConta & ";UID=" & Usuario & ";PASSWORD=" & Pass & ";PORT=3306;OPTION=3;STMT=;"
'        End If
'de david
        Cad = "DRIVER={MySQL ODBC 3.51 Driver};DESC=;DATA SOURCE= vConta;DATABASE=" & nomConta
        If vParamAplic.ServidorContaFac <> "" Then  'especificamos servidor
            Cad = Cad & ";SERVER=" & vParamAplic.ServidorContaFac
        End If
        
        Cad = Cad & ";UID=" & Usuario
        Cad = Cad & ";PWD=" & Pass
        Cad = Cad & ";Persist Security Info=true"
        
        ConnConta.ConnectionString = Cad
        ConnConta.Open
        ConnConta.Execute "Set AUTOCOMMIT = 1"
        AbrirConexionConta = True
    Else
        AbrirConexionConta = False
    End If
    Exit Function
EAbrirConexion:
    MuestraError Err.Number, "Abrir conexión Contabilidad Avnics.", Err.Description
End Function

Public Function AbrirConexionContaSeg(Usuario As String, Pass As String) As Boolean
Dim Cad As String
Dim nomConta As String 'nombre de la BD de la contabilidad
Dim serConta As String 'servidor donde esta la BD de la contabilidad
On Error GoTo EAbrirConexionSeg
    
    AbrirConexionContaSeg = False
    
    Set ConnContaSeg = Nothing
    Set ConnContaSeg = New Connection
'    Conn.CursorLocation = adUseClient   'Si ponemos este hay opciones k no van ej select con rs!campo
    ConnContaSeg.CursorLocation = adUseServer   'Si ponemos esta alguns errores de Conn no se muestran correctamente

' ### [Monica] 06/09/2006
    If vParamAplic.ContabilidadNueva Then
        nomConta = "ariconta" & vParamAplic.NumeroContaSeg
    Else
        nomConta = "conta" & vParamAplic.NumeroContaSeg
    End If
'    vEmpresa.BDConta = nomConta
    If vParamAplic.NumeroContaSeg <> 0 Then
    '    cad = "DRIVER={MySQL ODBC 3.51 Driver};DESC=;DATABASE=conta" & vParamConta.NumeroConta & ";SERVER=" & vParamConta.ServidorConta & ";"
    '    cad = cad & ";UID=" & vParamConta.UsuarioConta
    '    cad = cad & ";PWD=" & vParamConta.PasswordConta
    '    cad = "DSN=vConta;DESC=MySQL ODBC 3.51 Driver DSN;DATABASE=conta2;SERVER=david;UID=" & Usuario & ";PASSWORD=" & Pass & ";PORT=3306;OPTION=3;STMT=;"
'--monica: cambiado por lo de david
'        If vParamAplic.ServidorContaSeg <> "" Then  'especificamos servidor
'            cad = "DSN=vConta;DESC=MySQL ODBC 3.51 Driver DSN;DATABASE=" & nomConta & ";SERVER=" & vParamAplic.ServidorContaSeg & ";UID=" & Usuario & ";PASSWORD=" & Pass & ";PORT=3306;OPTION=3;STMT=;"
'        Else 'por defecto cogera la BD del servidor que haya en el ODBC
'            cad = "DSN=vConta;DESC=MySQL ODBC 3.51 Driver DSN;DATABASE=" & nomConta & ";UID=" & Usuario & ";PASSWORD=" & Pass & ";PORT=3306;OPTION=3;STMT=;"
'        End If
'de david
        Cad = "DRIVER={MySQL ODBC 3.51 Driver};DESC=;DATA SOURCE= vConta;DATABASE=" & nomConta
        If vParamAplic.ServidorContaSeg <> "" Then  'especificamos servidor
            Cad = Cad & ";SERVER=" & vParamAplic.ServidorContaSeg
        End If
        Cad = Cad & ";UID=" & Usuario
        Cad = Cad & ";PWD=" & Pass
        Cad = Cad & ";Persist Security Info=true"

        ConnContaSeg.ConnectionString = Cad
        ConnContaSeg.Open
        ConnContaSeg.Execute "Set AUTOCOMMIT = 1"
        AbrirConexionContaSeg = True
    Else
        AbrirConexionContaSeg = False
    End If
    Exit Function
EAbrirConexionSeg:
    MuestraError Err.Number, "Abrir conexión Contabilidad Seguros.", Err.Description
End Function

Public Function AbrirConexionContaTel(Usuario As String, Pass As String) As Boolean
Dim Cad As String
Dim nomConta As String 'nombre de la BD de la contabilidad
Dim serConta As String 'servidor donde esta la BD de la contabilidad
On Error GoTo EAbrirConexionTel
    
    AbrirConexionContaTel = False
    
    Set ConnContaTel = Nothing
    Set ConnContaTel = New Connection
'    Conn.CursorLocation = adUseClient   'Si ponemos este hay opciones k no van ej select con rs!campo
    ConnContaTel.CursorLocation = adUseServer   'Si ponemos esta alguns errores de Conn no se muestran correctamente

' ### [Monica] 06/09/2006
    If vParamAplic.ContabilidadNueva Then
        nomConta = "ariconta" & vParamAplic.NumeroContaTel
    Else
        nomConta = "conta" & vParamAplic.NumeroContaTel
    End If
'    vEmpresa.BDConta = nomConta
    If vParamAplic.NumeroContaTel <> 0 Then
    '    cad = "DRIVER={MySQL ODBC 3.51 Driver};DESC=;DATABASE=conta" & vParamConta.NumeroConta & ";SERVER=" & vParamConta.ServidorConta & ";"
    '    cad = cad & ";UID=" & vParamConta.UsuarioConta
    '    cad = cad & ";PWD=" & vParamConta.PasswordConta
    '    cad = "DSN=vConta;DESC=MySQL ODBC 3.51 Driver DSN;DATABASE=conta2;SERVER=david;UID=" & Usuario & ";PASSWORD=" & Pass & ";PORT=3306;OPTION=3;STMT=;"
'--monica: cambiado por lo de david
'        If vParamAplic.ServidorContaTel <> "" Then  'especificamos servidor
'            cad = "DSN=vConta;DESC=MySQL ODBC 3.51 Driver DSN;DATABASE=" & nomConta & ";SERVER=" & vParamAplic.ServidorContaTel & ";UID=" & Usuario & ";PASSWORD=" & Pass & ";PORT=3306;OPTION=3;STMT=;"
'        Else 'por defecto cogera la BD del servidor que haya en el ODBC
'            cad = "DSN=vConta;DESC=MySQL ODBC 3.51 Driver DSN;DATABASE=" & nomConta & ";UID=" & Usuario & ";PASSWORD=" & Pass & ";PORT=3306;OPTION=3;STMT=;"
'        End If
        
        Cad = "DRIVER={MySQL ODBC 3.51 Driver};DESC=;DATA SOURCE= vConta;DATABASE=" & nomConta
        If vParamAplic.ServidorContaTel <> "" Then  'especificamos servidor
            Cad = Cad & ";SERVER=" & vParamAplic.ServidorContaTel
        End If
        Cad = Cad & ";UID=" & Usuario
        Cad = Cad & ";PWD=" & Pass
        Cad = Cad & ";Persist Security Info=true"


        ConnContaTel.ConnectionString = Cad
        ConnContaTel.Open
        ConnContaTel.Execute "Set AUTOCOMMIT = 1"
        AbrirConexionContaTel = True
    Else
        AbrirConexionContaTel = False
    End If
    Exit Function
EAbrirConexionTel:
    MuestraError Err.Number, "Abrir conexión Contabilidad Telefonía.", Err.Description
End Function

Public Function AbrirConexionContaCV(Usuario As String, Pass As String) As Boolean
Dim Cad As String
Dim nomConta As String 'nombre de la BD de la contabilidad
Dim serConta As String 'servidor donde esta la BD de la contabilidad
On Error GoTo EAbrirConexionCV
    
    AbrirConexionContaCV = False
    
    Set ConnContaCV = Nothing
    Set ConnContaCV = New Connection
'    Conn.CursorLocation = adUseClient   'Si ponemos este hay opciones k no van ej select con rs!campo
    ConnContaCV.CursorLocation = adUseServer   'Si ponemos esta alguns errores de Conn no se muestran correctamente

' ### [Monica] 06/09/2006
    If vParamAplic.ContabilidadNueva Then
        nomConta = "ariconta" & vParamAplic.NumeroContaCV
    Else
        nomConta = "conta" & vParamAplic.NumeroContaCV
    End If
'    vEmpresa.BDConta = nomConta
    If vParamAplic.NumeroContaCV <> 0 Then
    '    cad = "DRIVER={MySQL ODBC 3.51 Driver};DESC=;DATABASE=conta" & vParamConta.NumeroConta & ";SERVER=" & vParamConta.ServidorConta & ";"
    '    cad = cad & ";UID=" & vParamConta.UsuarioConta
    '    cad = cad & ";PWD=" & vParamConta.PasswordConta
    '    cad = "DSN=vConta;DESC=MySQL ODBC 3.51 Driver DSN;DATABASE=conta2;SERVER=david;UID=" & Usuario & ";PASSWORD=" & Pass & ";PORT=3306;OPTION=3;STMT=;"
'--monica: cambiado por lo de david
'        If vParamAplic.ServidorContaTel <> "" Then  'especificamos servidor
'            cad = "DSN=vConta;DESC=MySQL ODBC 3.51 Driver DSN;DATABASE=" & nomConta & ";SERVER=" & vParamAplic.ServidorContaTel & ";UID=" & Usuario & ";PASSWORD=" & Pass & ";PORT=3306;OPTION=3;STMT=;"
'        Else 'por defecto cogera la BD del servidor que haya en el ODBC
'            cad = "DSN=vConta;DESC=MySQL ODBC 3.51 Driver DSN;DATABASE=" & nomConta & ";UID=" & Usuario & ";PASSWORD=" & Pass & ";PORT=3306;OPTION=3;STMT=;"
'        End If
        
        Cad = "DRIVER={MySQL ODBC 3.51 Driver};DESC=;DATA SOURCE= vConta;DATABASE=" & nomConta
        If vParamAplic.ServidorContaCV <> "" Then 'especificamos servidor
            Cad = Cad & ";SERVER=" & vParamAplic.ServidorContaCV
        End If
        Cad = Cad & ";UID=" & Usuario
        Cad = Cad & ";PWD=" & Pass
        Cad = Cad & ";Persist Security Info=true"


        ConnContaCV.ConnectionString = Cad
        ConnContaCV.Open
        ConnContaCV.Execute "Set AUTOCOMMIT = 1"
        AbrirConexionContaCV = True
    Else
        AbrirConexionContaCV = False
    End If
    Exit Function
EAbrirConexionCV:
    MuestraError Err.Number, "Abrir conexión Contabilidad Facturas CV.", Err.Description
End Function


Public Function AbrirConexionContaCVV(Usuario As String, Pass As String) As Boolean
Dim Cad As String
Dim nomConta As String 'nombre de la BD de la contabilidad
Dim serConta As String 'servidor donde esta la BD de la contabilidad
On Error GoTo EAbrirConexionCVV
    
    AbrirConexionContaCVV = False
    
    Set ConnContaCVV = Nothing
    Set ConnContaCVV = New Connection
'    Conn.CursorLocation = adUseClient   'Si ponemos este hay opciones k no van ej select con rs!campo
    ConnContaCVV.CursorLocation = adUseServer   'Si ponemos esta alguns errores de Conn no se muestran correctamente

' ### [Monica] 06/09/2006
    If vParamAplic.ContabilidadNueva Then
        nomConta = "ariconta" & vParamAplic.NumeroContaCVV
    Else
        nomConta = "conta" & vParamAplic.NumeroContaCVV
    End If
'    vEmpresa.BDConta = nomConta
    If vParamAplic.NumeroContaCVV <> 0 Then
    '    cad = "DRIVER={MySQL ODBC 3.51 Driver};DESC=;DATABASE=conta" & vParamConta.NumeroConta & ";SERVER=" & vParamConta.ServidorConta & ";"
    '    cad = cad & ";UID=" & vParamConta.UsuarioConta
    '    cad = cad & ";PWD=" & vParamConta.PasswordConta
    '    cad = "DSN=vConta;DESC=MySQL ODBC 3.51 Driver DSN;DATABASE=conta2;SERVER=david;UID=" & Usuario & ";PASSWORD=" & Pass & ";PORT=3306;OPTION=3;STMT=;"
'--monica: cambiado por lo de david
'        If vParamAplic.ServidorContaTel <> "" Then  'especificamos servidor
'            cad = "DSN=vConta;DESC=MySQL ODBC 3.51 Driver DSN;DATABASE=" & nomConta & ";SERVER=" & vParamAplic.ServidorContaTel & ";UID=" & Usuario & ";PASSWORD=" & Pass & ";PORT=3306;OPTION=3;STMT=;"
'        Else 'por defecto cogera la BD del servidor que haya en el ODBC
'            cad = "DSN=vConta;DESC=MySQL ODBC 3.51 Driver DSN;DATABASE=" & nomConta & ";UID=" & Usuario & ";PASSWORD=" & Pass & ";PORT=3306;OPTION=3;STMT=;"
'        End If
        
        Cad = "DRIVER={MySQL ODBC 3.51 Driver};DESC=;DATA SOURCE= vConta;DATABASE=" & nomConta
        If vParamAplic.ServidorContaCV <> "" Then 'especificamos servidor
            Cad = Cad & ";SERVER=" & vParamAplic.ServidorContaCV
        End If
        Cad = Cad & ";UID=" & Usuario
        Cad = Cad & ";PWD=" & Pass
        Cad = Cad & ";Persist Security Info=true"


        ConnContaCVV.ConnectionString = Cad
        ConnContaCVV.Open
        ConnContaCVV.Execute "Set AUTOCOMMIT = 1"
        AbrirConexionContaCVV = True
    Else
        AbrirConexionContaCVV = False
    End If
    Exit Function
EAbrirConexionCVV:
    MuestraError Err.Number, "Abrir conexión Contabilidad Facturas CV Varias.", Err.Description
End Function





Public Function AbrirConexionContaFac(Usuario As String, Pass As String, NumConta As Integer) As Boolean
Dim Cad As String
Dim nomConta As String 'nombre de la BD de la contabilidad
Dim serConta As String 'servidor donde esta la BD de la contabilidad
On Error GoTo EAbrirConexion
    
    AbrirConexionContaFac = False
    
    Set ConnContaFac = Nothing
    Set ConnContaFac = New Connection
'    Conn.CursorLocation = adUseClient   'Si ponemos este hay opciones k no van ej select con rs!campo
    ConnContaFac.CursorLocation = adUseServer   'Si ponemos esta alguns errores de Conn no se muestran correctamente

' ### [Monica] 06/09/2006
    If vParamAplic.ContabilidadNueva Then
        nomConta = "ariconta" & NumConta
    Else
        nomConta = "conta" & NumConta
    End If
    If NumConta <> 0 Then
    '    cad = "DRIVER={MySQL ODBC 3.51 Driver};DESC=;DATABASE=conta" & vParamConta.NumeroConta & ";SERVER=" & vParamConta.ServidorConta & ";"
    '    cad = cad & ";UID=" & vParamConta.UsuarioConta
    '    cad = cad & ";PWD=" & vParamConta.PasswordConta
    '    cad = "DSN=vConta;DESC=MySQL ODBC 3.51 Driver DSN;DATABASE=conta2;SERVER=david;UID=" & Usuario & ";PASSWORD=" & Pass & ";PORT=3306;OPTION=3;STMT=;"
    
'--monica. cambiado por lo de David
'        If vParamAplic.ServidorContaFac <> "" Then  'especificamos servidor
'            cad = "DSN=vConta;DESC=MySQL ODBC 3.51 Driver DSN;DATABASE=" & nomConta & ";SERVER=" & vParamAplic.ServidorContaFac & ";UID=" & Usuario & ";PASSWORD=" & Pass & ";PORT=3306;OPTION=3;STMT=;"
'        Else 'por defecto cogera la BD del servidor que haya en el ODBC
'            cad = "DSN=vConta;DESC=MySQL ODBC 3.51 Driver DSN;DATABASE=" & nomConta & ";UID=" & Usuario & ";PASSWORD=" & Pass & ";PORT=3306;OPTION=3;STMT=;"
'        End If
'de david
        Cad = "DRIVER={MySQL ODBC 3.51 Driver};DESC=;DATA SOURCE= vConta;DATABASE=" & nomConta
        If vParamAplic.ServidorContaFac <> "" Then  'especificamos servidor
            Cad = Cad & ";SERVER=" & vParamAplic.ServidorContaFac
        End If
        
        Cad = Cad & ";UID=" & Usuario
        Cad = Cad & ";PWD=" & Pass
        Cad = Cad & ";Persist Security Info=true"
        
        ConnContaFac.ConnectionString = Cad
        ConnContaFac.Open
        ConnContaFac.Execute "Set AUTOCOMMIT = 1"
        AbrirConexionContaFac = True
    Else
        AbrirConexionContaFac = False
    End If
    Exit Function
EAbrirConexion:
    MuestraError Err.Number, "Abrir conexión Contabilidad Factura.", Err.Description
End Function


Public Function AbrirConexionContaGas(Usuario As String, Pass As String) As Boolean
Dim Cad As String
Dim nomConta As String 'nombre de la BD de la contabilidad
Dim serConta As String 'servidor donde esta la BD de la contabilidad

On Error GoTo EAbrirConexionGas
    
    AbrirConexionContaGas = False
    
    Set ConnContaGas = Nothing
    Set ConnContaGas = New Connection
'    Conn.CursorLocation = adUseClient   'Si ponemos este hay opciones k no van ej select con rs!campo
    ConnContaGas.CursorLocation = adUseServer   'Si ponemos esta alguns errores de Conn no se muestran correctamente

' ### [Monica] 06/09/2006
    If vParamAplic.ContabilidadNueva Then
        nomConta = "ariconta" & vParamAplic.NumeroContaGas
    Else
        nomConta = "conta" & vParamAplic.NumeroContaGas
    End If
'    vEmpresa.BDConta = nomConta
    If vParamAplic.NumeroContaGas <> 0 Then
    '    cad = "DRIVER={MySQL ODBC 3.51 Driver};DESC=;DATABASE=conta" & vParamConta.NumeroConta & ";SERVER=" & vParamConta.ServidorConta & ";"
    '    cad = cad & ";UID=" & vParamConta.UsuarioConta
    '    cad = cad & ";PWD=" & vParamConta.PasswordConta
    '    cad = "DSN=vConta;DESC=MySQL ODBC 3.51 Driver DSN;DATABASE=conta2;SERVER=david;UID=" & Usuario & ";PASSWORD=" & Pass & ";PORT=3306;OPTION=3;STMT=;"
'--monica: cambiado por lo de david
'        If vParamAplic.ServidorContaSeg <> "" Then  'especificamos servidor
'            cad = "DSN=vConta;DESC=MySQL ODBC 3.51 Driver DSN;DATABASE=" & nomConta & ";SERVER=" & vParamAplic.ServidorContaSeg & ";UID=" & Usuario & ";PASSWORD=" & Pass & ";PORT=3306;OPTION=3;STMT=;"
'        Else 'por defecto cogera la BD del servidor que haya en el ODBC
'            cad = "DSN=vConta;DESC=MySQL ODBC 3.51 Driver DSN;DATABASE=" & nomConta & ";UID=" & Usuario & ";PASSWORD=" & Pass & ";PORT=3306;OPTION=3;STMT=;"
'        End If
'de david
        Cad = "DRIVER={MySQL ODBC 3.51 Driver};DESC=;DATA SOURCE= vConta;DATABASE=" & nomConta
        If vParamAplic.ServidorContaGas <> "" Then  'especificamos servidor
            Cad = Cad & ";SERVER=" & vParamAplic.ServidorContaGas
        End If
        Cad = Cad & ";UID=" & Usuario
        Cad = Cad & ";PWD=" & Pass
        Cad = Cad & ";Persist Security Info=true"

        ConnContaGas.ConnectionString = Cad
        ConnContaGas.Open
        ConnContaGas.Execute "Set AUTOCOMMIT = 1"
        AbrirConexionContaGas = True
    Else
        AbrirConexionContaGas = False
    End If
    Exit Function
EAbrirConexionGas:
    MuestraError Err.Number, "Abrir conexión Contabilidad Gasolinera.", Err.Description
End Function


Public Function AbrirConexionContaFacSoc(Usuario As String, Pass As String) As Boolean
Dim Cad As String
Dim nomConta As String 'nombre de la BD de la contabilidad
Dim serConta As String 'servidor donde esta la BD de la contabilidad

On Error GoTo EAbrirConexionfacsoc
    
    AbrirConexionContaFacSoc = False
    
    Set ConnContaFacSoc = Nothing
    Set ConnContaFacSoc = New Connection
'    Conn.CursorLocation = adUseClient   'Si ponemos este hay opciones k no van ej select con rs!campo
    ConnContaFacSoc.CursorLocation = adUseServer   'Si ponemos esta alguns errores de Conn no se muestran correctamente

' ### [Monica] 06/09/2006
    If vParamAplic.ContabilidadNueva Then
        nomConta = "ariconta" & vParamAplic.NumeroContaFacSoc
    Else
        nomConta = "conta" & vParamAplic.NumeroContaFacSoc
    End If
'    vEmpresa.BDConta = nomConta
    If vParamAplic.NumeroContaFacSoc <> 0 Then
    '    cad = "DRIVER={MySQL ODBC 3.51 Driver};DESC=;DATABASE=conta" & vParamConta.NumeroConta & ";SERVER=" & vParamConta.ServidorConta & ";"
    '    cad = cad & ";UID=" & vParamConta.UsuarioConta
    '    cad = cad & ";PWD=" & vParamConta.PasswordConta
    '    cad = "DSN=vConta;DESC=MySQL ODBC 3.51 Driver DSN;DATABASE=conta2;SERVER=david;UID=" & Usuario & ";PASSWORD=" & Pass & ";PORT=3306;OPTION=3;STMT=;"
'--monica: cambiado por lo de david
'        If vParamAplic.ServidorContaSeg <> "" Then  'especificamos servidor
'            cad = "DSN=vConta;DESC=MySQL ODBC 3.51 Driver DSN;DATABASE=" & nomConta & ";SERVER=" & vParamAplic.ServidorContaSeg & ";UID=" & Usuario & ";PASSWORD=" & Pass & ";PORT=3306;OPTION=3;STMT=;"
'        Else 'por defecto cogera la BD del servidor que haya en el ODBC
'            cad = "DSN=vConta;DESC=MySQL ODBC 3.51 Driver DSN;DATABASE=" & nomConta & ";UID=" & Usuario & ";PASSWORD=" & Pass & ";PORT=3306;OPTION=3;STMT=;"
'        End If
'de david
        Cad = "DRIVER={MySQL ODBC 3.51 Driver};DESC=;DATA SOURCE= vConta;DATABASE=" & nomConta
        If vParamAplic.ServidorContaFacSoc <> "" Then  'especificamos servidor
            Cad = Cad & ";SERVER=" & vParamAplic.ServidorContaFacSoc
        End If
        Cad = Cad & ";UID=" & Usuario
        Cad = Cad & ";PWD=" & Pass
        Cad = Cad & ";Persist Security Info=true"

        ConnContaFacSoc.ConnectionString = Cad
        ConnContaFacSoc.Open
        ConnContaFacSoc.Execute "Set AUTOCOMMIT = 1"
        AbrirConexionContaFacSoc = True
    Else
        AbrirConexionContaFacSoc = False
    End If
    Exit Function
EAbrirConexionfacsoc:
    MuestraError Err.Number, "Abrir conexión Contabilidad Facturas Socios.", Err.Description
End Function





Public Function AbrirConexionAuxCon(Empresa As String, Usuario As String, Pass As String) As Boolean
Dim Cad As String
Dim nomConta As String 'nombre de la BD de la contabilidad
Dim serConta As String 'servidor donde esta la BD de la contabilidad
On Error GoTo EAbrirConexion

    AbrirConexionAuxCon = False

    Set ConnAuxCon = Nothing
    Set ConnAuxCon = New Connection
'    Conn.CursorLocation = adUseClient   'Si ponemos este hay opciones k no van ej select con rs!campo
    ConnAuxCon.CursorLocation = adUseServer   'Si ponemos esta alguns errores de Conn no se muestran correctamente

    'Obtener la BD de contabilidad
'    SQL = "select bdaconta FROM paramcon WHERE codempre=" & codEmpre
    serConta = "serconta"
    nomConta = DevuelveDesdeBDNew(2, "sparam", "bdaconta", "codempre", Empresa, "N", serConta)
'    vEmpresa.BDConta = nomConta
    If nomConta <> "" Then
    '    cad = "DRIVER={MySQL ODBC 3.51 Driver};DESC=;DATABASE=conta" & vParamConta.NumeroConta & ";SERVER=" & vParamConta.ServidorConta & ";"
    '    cad = cad & ";UID=" & vParamConta.UsuarioConta
    '    cad = cad & ";PWD=" & vParamConta.PasswordConta
    '    cad = "DSN=vConta;DESC=MySQL ODBC 3.51 Driver DSN;DATABASE=conta2;SERVER=david;UID=" & Usuario & ";PASSWORD=" & Pass & ";PORT=3306;OPTION=3;STMT=;"
        If serConta <> "" Then 'especificamos servidor
            Cad = "DSN=vConta;DESC=MySQL ODBC 3.51 Driver DSN;DATABASE=" & nomConta & ";SERVER=" & serConta & ";UID=" & Usuario & ";PASSWORD=" & Pass & ";PORT=3306;OPTION=3;STMT=;"
        Else 'por defecto cogera la BD del servidor que haya en el ODBC
            Cad = "DSN=vConta;DESC=MySQL ODBC 3.51 Driver DSN;DATABASE=" & nomConta & ";UID=" & Usuario & ";PASSWORD=" & Pass & ";PORT=3306;OPTION=3;STMT=;"
        End If
        ConnAuxCon.ConnectionString = Cad
        ConnAuxCon.Open
        ConnAuxCon.Execute "Set AUTOCOMMIT = 1"
        AbrirConexionAuxCon = True
    Else
        AbrirConexionAuxCon = False
    End If
    Exit Function
EAbrirConexion:
    MuestraError Err.Number, "Abrir conexión Contabilidad.", Err.Description
End Function



Public Function AbrirConexionAridoc(Usuario As String, Pass As String) As Boolean
'Abre
Dim Cad As String
On Error GoTo EAbrirConexion

    
    AbrirConexionAridoc = False
    Set ConnAridoc = Nothing
    Set ConnAridoc = New Connection
'    Conn.CursorLocation = adUseClient   'Si ponemos este hay opciones k no van ej select con rs!campo
    ConnAridoc.CursorLocation = adUseServer   'Si ponemos esta alguns errores de Conn no se muestran correctamente
                        
'    cad = "DSN=Aridoc;DESC=MySQL ODBC 3.51 Driver DSN;UID=" & Usuario & ";PASSWORD=" & Pass & ";PORT=3306;OPTION=3;STMT=;"
    Cad = "DRIVER={MySQL ODBC 3.51 Driver};DESC=;DATA SOURCE=Aridoc;DATABASE=Aridoc;"
    Cad = Cad & ";UID=" & Usuario
    Cad = Cad & ";PWD=" & Pass
                     
    '++monica:tema del vista
    Cad = Cad & ";Persist Security Info=true"
    '++
                     
                     
    ConnAridoc.ConnectionString = Cad
    ConnAridoc.Open
    ConnAridoc.Execute "Set AUTOCOMMIT = 1"
    AbrirConexionAridoc = True
    Exit Function
EAbrirConexion:
    MuestraError Err.Number, "Abrir conexión Aridoc.", Err.Description
End Function






Public Function CerrarConexionConta()
  'Cerramos la conexion con BD: Contabilidad
  On Error Resume Next
   ConnConta.Close
   If Err.Number <> 0 Then Err.Clear
End Function

Public Function CerrarConexionContaSeg()
  'Cerramos la conexion con BD: Contabilidad
  On Error Resume Next
   ConnContaSeg.Close
   If Err.Number <> 0 Then Err.Clear
End Function

Public Function CerrarConexionContaTel()
  'Cerramos la conexion con BD: Contabilidad
  On Error Resume Next
   ConnContaTel.Close
   If Err.Number <> 0 Then Err.Clear
End Function

Public Function CerrarConexionContaCV()
  'Cerramos la conexion con BD: Contabilidad
  On Error Resume Next
   ConnContaCV.Close
   If Err.Number <> 0 Then Err.Clear
End Function

Public Function CerrarConexionContaCVV()
  'Cerramos la conexion con BD: Contabilidad
  On Error Resume Next
   ConnContaCVV.Close
   If Err.Number <> 0 Then Err.Clear
End Function

Public Function CerrarConexionContaFac()
  'Cerramos la conexion con BD: Contabilidad
  On Error Resume Next
   ConnContaFac.Close
   If Err.Number <> 0 Then Err.Clear
End Function

Public Function CerrarConexionContaGas()
  'Cerramos la conexion con BD: Contabilidad
  On Error Resume Next
   ConnContaGas.Close
   If Err.Number <> 0 Then Err.Clear
End Function


Public Function CerrarConexionContaFacSoc()
  'Cerramos la conexion con BD: Contabilidad
  On Error Resume Next
   ConnContaFacSoc.Close
   If Err.Number <> 0 Then Err.Clear
End Function


Public Function CerrarConexionAridoc()
  'Cerramos la conexion con BD: Aridoc
  On Error Resume Next
   ConnAridoc.Close
   If Err.Number <> 0 Then Err.Clear
End Function



Public Sub LeerDatosEmpresa()
'Crea instancia de la clase Cempresa con los valores en
'Tabla: Empresas
'BDatos: PTours y Conta
 
    Set vEmpresa = New Cempresa
    Set vEmpresaSeg = New CempresaSeg
    Set vEmpresaTel = New CempresaTel
    Set vEmpresaGas = New CempresaGas
    Set vEmpresaFacSoc = New CempresaFacSoc
    Set vEmpresaCV = New CempresaCV
    Set vEmpresaCVV = New CempresaCVV
    
    If vEmpresa.LeerDatos(1) = False Then  'De Avnics
        MsgBox "No se han podido cargar los datos de la empresa. Debe configurar la aplicación.", vbExclamation
        Set vEmpresa = Nothing
       ' Set vSesion = Nothing
       ' Set conn = Nothing
        Exit Sub
    End If
    
    ' ### [Monica] 06/09/2006
    ' añadido
    Set vParamAplic = New CParamAplic
    If vParamAplic.leer = 1 Then
        MsgBox "No se han podido cargar los parámetros de contabilidad. Debe configurar la aplicación.", vbExclamation
        
        Set vParamAplic = Nothing
        Exit Sub
    Else
        If vParamAplic.Avnics = 1 Then
            If vParamAplic.NumeroConta <> 0 Then
            
                'Abrir conexión a la BDatos de Contabilidad para acceder a
                'Tablas: Cuentas, Tipos IVA,...
                If AbrirConexionConta(vParamAplic.UsuarioConta, vParamAplic.PasswordConta) = False Then
                    MsgBox "La aplicación no puede continuar sin acceso a los datos de Contabilidad. ", vbCritical
                    AccionesCerrar
                    End
                End If
            ' ### [Monica] 06/09/2006
            ' comento los niveles de contabilidad pq solo tengo las cuentas
                If vEmpresa.LeerNiveles() = False Then  'De Contabilidad
                    MsgBox "No se han podido cargar los niveles de la contabilidad de la empresa. Debe configurar la aplicación.", vbExclamation
                    AccionesCerrar
                    End
                End If
                
                FechasEjercicioConta FIni, FFin
            End If
        End If
        If vParamAplic.Seguros = 1 Then
            If vParamAplic.NumeroContaSeg <> 0 Then
            
                'Abrir conexión a la BDatos de Contabilidad para acceder a
                'Tablas: Cuentas, Tipos IVA,...
                If AbrirConexionContaSeg(vParamAplic.UsuarioContaSeg, vParamAplic.PasswordContaSeg) = False Then
                    MsgBox "La aplicación no puede continuar sin acceso a los datos de Contabilidad. ", vbCritical
                    AccionesCerrar
                    End
                End If
            ' ### [Monica] 06/09/2006
            ' comento los niveles de contabilidad pq solo tengo las cuentas
                If vEmpresaSeg.LeerNiveles() = False Then  'De Contabilidad
                    MsgBox "No se han podido cargar los niveles de la contabilidad de la empresa. Debe configurar la aplicación.", vbExclamation
                    AccionesCerrar
                    End
                End If
                
                FechasEjercicioContaSeg FIniSeg, FFinSeg
            End If
        End If
        If vParamAplic.Telefonia = 1 Then
            If vParamAplic.NumeroContaTel <> 0 Then
            
                'Abrir conexión a la BDatos de Contabilidad para acceder a
                'Tablas: Cuentas, Tipos IVA,...
                If AbrirConexionContaTel(vParamAplic.UsuarioContaTel, vParamAplic.PasswordContaTel) = False Then
                    MsgBox "La aplicación no puede continuar sin acceso a los datos de Contabilidad. ", vbCritical
                    AccionesCerrar
                    End
                End If
            ' ### [Monica] 06/09/2006
            ' comento los niveles de contabilidad pq solo tengo las cuentas
                If vEmpresaTel.LeerNiveles() = False Then  'De Contabilidad
                    MsgBox "No se han podido cargar los niveles de la contabilidad de la empresa. Debe configurar la aplicación.", vbExclamation
                    AccionesCerrar
                    End
                End If
                
                FechasEjercicioContaTel FIniTel, FFinTel
            End If
        End If
        If vParamAplic.Gasolinera = 1 Then
            If vParamAplic.NumeroContaGas <> 0 Then
            
                'Abrir conexión a la BDatos de Contabilidad para acceder a
                'Tablas: Cuentas, Tipos IVA,...
                If AbrirConexionContaGas(vParamAplic.UsuarioContaGas, vParamAplic.PasswordContaGas) = False Then
                    MsgBox "La aplicación no puede continuar sin acceso a los datos de Contabilidad. ", vbCritical
                    AccionesCerrar
                    End
                End If
            ' ### [Monica] 06/09/2006
            ' comento los niveles de contabilidad pq solo tengo las cuentas
                If vEmpresaGas.LeerNiveles() = False Then  'De Contabilidad
                    MsgBox "No se han podido cargar los niveles de la contabilidad de la empresa. Debe configurar la aplicación.", vbExclamation
                    AccionesCerrar
                    End
                End If
                
                FechasEjercicioContaGas FIniGas, FFinGas
            End If
        End If
        If vParamAplic.FactSocios = 1 Then
            If vParamAplic.NumeroContaFacSoc <> 0 Then
            
                'Abrir conexión a la BDatos de Contabilidad para acceder a
                'Tablas: Cuentas, Tipos IVA,...
                If AbrirConexionContaFacSoc(vParamAplic.UsuarioContaFacSoc, vParamAplic.PasswordContaFacSoc) = False Then
                    MsgBox "La aplicación no puede continuar sin acceso a los datos de Contabilidad. ", vbCritical
                    AccionesCerrar
                    End
                End If
            ' ### [Monica] 06/09/2006
            ' comento los niveles de contabilidad pq solo tengo las cuentas
                If vEmpresaFacSoc.LeerNiveles() = False Then  'De Contabilidad
                    MsgBox "No se han podido cargar los niveles de la contabilidad de la empresa. Debe configurar la aplicación.", vbExclamation
                    AccionesCerrar
                    End
                End If
                
                FechasEjercicioContaFacSoc FIniFacSoc, FFinFacSoc
            End If
        End If
        If vParamAplic.Coarval = 1 Then
            If vParamAplic.NumeroContaCV <> 0 Then
            
                'Abrir conexión a la BDatos de Contabilidad para acceder a
                'Tablas: Cuentas, Tipos IVA,...
                If AbrirConexionContaCV(vParamAplic.UsuarioContaCV, vParamAplic.PasswordContaCV) = False Then
                    MsgBox "La aplicación no puede continuar sin acceso a los datos de Contabilidad. ", vbCritical
                    AccionesCerrar
                    End
                End If
            ' ### [Monica] 06/09/2006
            ' comento los niveles de contabilidad pq solo tengo las cuentas
                If vEmpresaCV.LeerNiveles() = False Then  'De Contabilidad
                    MsgBox "No se han podido cargar los niveles de la contabilidad de la empresa. Debe configurar la aplicación.", vbExclamation
                    AccionesCerrar
                    End
                End If
                
                FechasEjercicioContaCV FIniTel, FFinTel
            End If
        
            If vParamAplic.NumeroContaCVV <> 0 Then
            
                'Abrir conexión a la BDatos de Contabilidad para acceder a
                'Tablas: Cuentas, Tipos IVA,...
                If AbrirConexionContaCVV(vParamAplic.UsuarioContaCV, vParamAplic.PasswordContaCV) = False Then
                    MsgBox "La aplicación no puede continuar sin acceso a los datos de Contabilidad. ", vbCritical
                    AccionesCerrar
                    End
                End If
            ' ### [Monica] 06/09/2006
            ' comento los niveles de contabilidad pq solo tengo las cuentas
                If vEmpresaCVV.LeerNiveles() = False Then  'De Contabilidad
                    MsgBox "No se han podido cargar los niveles de la contabilidad de la empresa. Debe configurar la aplicación.", vbExclamation
                    AccionesCerrar
                    End
                End If
                
                FechasEjercicioContaCVV FIniTel, FFinTel
            End If
        
        
        
        End If
        
    End If
'    Set vParam = New Cparametros
'    If vParam.Leer = False Then   'De AriGasol
'        MsgBox "No se han podido cargar los parámetros de la empresa. Debe configurar la aplicación.", vbExclamation
'        Set vEmpresa = Nothing
'        Set vSesion = Nothing
'        Set Conn = Nothing
'        End
'    End If
End Sub


Public Function PonerDatosPpal()
    If Not vEmpresa Is Nothing Then
        MDIppal.Caption = "ARIAGROUTIL" & " ver. " & App.Major & "." & App.Minor & "." & App.Revision & "   -  Empresa: " & vEmpresa.nomEmpre
    End If
    If Err.Number <> 0 Then MuestraError Err.Description, "Poniendo datos de la pantalla principal", Err.Description
End Function

    

Public Sub MuestraError(numero As Long, Optional Cadena As String, Optional Desc As String)
    Dim Cad As String
    Dim Aux As String
    
    'Con este sub pretendemos unificar el msgbox para todos los errores
    'que se produzcan
    On Error Resume Next
    Cad = "Se ha producido un error: " & vbCrLf
    If Cadena <> "" Then
        Cad = Cad & vbCrLf & Cadena & vbCrLf & vbCrLf
    End If
    'Numeros de errores que contolamos
    If conn.Errors.Count > 0 Then
        ControlamosError Aux
        conn.Errors.Clear
    Else
        Aux = ""
    End If
    If Aux <> "" Then Desc = Aux
    If Desc <> "" Then Cad = Cad & vbCrLf & Desc & vbCrLf & vbCrLf
    If Aux = "" Then Cad = Cad & "Número: " & numero & vbCrLf & "Descripción: " & Error(numero)
    MsgBox Cad, vbExclamation
End Sub

Public Function DBSet(vData As Variant, Tipo As String, Optional EsNulo As String) As Variant
'Establece el valor del dato correcto antes de Insertar en la BD
Dim Cad As String

        If IsNull(vData) Then
            DBSet = ValorNulo
            Exit Function
        End If

        If Tipo <> "" Then
            Select Case Tipo
                Case "T"    'Texto
                    If vData = "" Then
                        If EsNulo = "N" Then
                            DBSet = "''"
                        Else
                            DBSet = ValorNulo
                        End If
                    Else
                        Cad = (CStr(vData))
                        NombreSQL Cad
                        DBSet = "'" & Cad & "'"
                    End If
                    
                Case "N"    'Numero
                    If vData = "" Or vData = 0 Then
                        If EsNulo <> "" Then
                            If EsNulo = "S" Then
                                DBSet = ValorNulo
                            Else
                                DBSet = 0
                            End If
                        Else
                            DBSet = 0
                        End If
                    Else
                        Cad = CStr(ImporteFormateado(CStr(vData)))
                        DBSet = TransformaComasPuntos(Cad)
                    End If
                    
                Case "F"    'Fecha
'                     '==David
''                    DBLet = "0:00:00"
'                     '==Laura
                    If vData = "" Then
                        If EsNulo = "S" Then
                            DBSet = ValorNulo
                        Else
                            DBSet = "'1900-01-01'"
                        End If
                    Else
                        DBSet = "'" & Format(vData, FormatoFecha) & "'"
                    End If
                    
                Case "FH" 'Fecha/Hora
                    If vData = "" Then
                        If EsNulo = "S" Then DBSet = ValorNulo
                    Else
                        DBSet = "'" & Format(vData, "yyyy-mm-dd hh:mm:ss") & "'"
                    End If
                    
                Case "H" 'Hora
                    If vData = "" Then
                    Else
                        DBSet = "'" & Format(vData, "hh:mm:ss") & "'"
                    End If
                    
                Case "B"  'Boolean
                    If vData Then
                        DBSet = 1
                    Else
                        DBSet = 0
                    End If
            End Select
        End If
End Function

Public Function DBLetMemo(vData As Variant) As Variant
    On Error Resume Next
    
    DBLetMemo = vData
    
    
    
    If Err.Number <> 0 Then
        Err.Clear
        DBLetMemo = ""
    End If
End Function



Public Function DBLet(vData As Variant, Optional Tipo As String) As Variant
'Para cuando recupera Datos de la BD
    If IsNull(vData) Then
        DBLet = ""
        If Tipo <> "" Then
            Select Case Tipo
                Case "T"    'Texto
                    DBLet = ""
                Case "N"    'Numero
                    DBLet = 0
                Case "F"    'Fecha
                     '==David
'                    DBLet = "0:00:00"
                     '==Laura
'                     DBLet = "0000-00-00"
                      DBLet = ""
                Case "D"
                    DBLet = 0
                Case "B"  'Boolean
                    DBLet = False
                Case Else
                    DBLet = ""
            End Select
        End If
    Else
        DBLet = vData
    End If
End Function

'/////////////////////////////////////////////////
'   Esto lo ejecutaremos justo antes de bloquear
'   Prepara la conexion para bloquear
Public Sub PreparaBloquear()
    conn.Execute "commit"
    conn.Execute "set autocommit=0"
End Sub

'/////////////////////////////////////////////////
'   Esto lo ejecutaremos justo despues de un bloque
'   Prepara la conexion para bloquear
Public Sub TerminaBloquear()
    conn.Execute "commit"
    conn.Execute "set autocommit=1"
End Sub

'///////////////////////////////////////////////////////////////
'
'   Cogemos un numero formateado: 1.256.256,98  y deevolvemos 1256256,98
'   Tiene que venir numérico
Public Function ImporteFormateado(Importe As String) As Currency
Dim i As Integer

    If Importe = "" Then
        ImporteFormateado = 0
    Else
        'Primero quitamos los puntos
        Do
            i = InStr(1, Importe, ".")
            If i > 0 Then Importe = Mid(Importe, 1, i - 1) & Mid(Importe, i + 1)
        Loop Until i = 0
        ImporteFormateado = Importe
    End If
End Function

' ### [Monica] 11/09/2006
Public Function ImporteSinFormato(Cadena As String) As String
Dim i As Integer
'Quitamos puntos
Do
    i = InStr(1, Cadena, ".")
    If i > 0 Then Cadena = Mid(Cadena, 1, i - 1) & Mid(Cadena, i + 1)
Loop Until i = 0
ImporteSinFormato = TransformaPuntosComas(Cadena)
End Function



'Cambia los puntos de los numeros decimales
'por comas
Public Function TransformaComasPuntos(Cadena As String) As String
Dim i As Integer
    Do
        i = InStr(1, Cadena, ",")
        If i > 0 Then
            Cadena = Mid(Cadena, 1, i - 1) & "." & Mid(Cadena, i + 1)
        End If
    Loop Until i = 0
    TransformaComasPuntos = Cadena
End Function

'Para los nombre que pueden tener ' . Para las comillas habra que hacer dentro otro INSTR
Public Sub NombreSQL(ByRef Cadena As String)
Dim J As Integer
Dim i As Integer
Dim Aux As String
    J = 1
    Do
        i = InStr(J, Cadena, "'")
        If i > 0 Then
            Aux = Mid(Cadena, 1, i - 1) & "\"
            Cadena = Aux & Mid(Cadena, i)
            J = i + 2
        End If
    Loop Until i = 0
End Sub

Public Function EsFechaOKString(ByRef T As String) As Boolean
Dim Cad As String
    
    Cad = T
    If InStr(1, Cad, "/") = 0 Then
        If Len(T) = 8 Then
            Cad = Mid(Cad, 1, 2) & "/" & Mid(Cad, 3, 2) & "/" & Mid(Cad, 5)
        Else
            If Len(T) = 6 Then Cad = Mid(Cad, 1, 2) & "/" & Mid(Cad, 3, 2) & "/" & Mid(Cad, 5)
        End If
    End If
    If IsDate(Cad) Then
        EsFechaOKString = True
        T = Format(Cad, "dd/mm/yyyy")
    Else
        EsFechaOKString = False
    End If
End Function

Public Function DevNombreSQL(Cadena As String) As String
Dim J As Integer
Dim i As Integer
Dim Aux As String
    J = 1
    Do
        i = InStr(J, Cadena, "'")
        If i > 0 Then
            Aux = Mid(Cadena, 1, i - 1) & "\"
            Cadena = Aux & Mid(Cadena, i)
            J = i + 2
        End If
    Loop Until i = 0
    DevNombreSQL = Cadena
End Function


Public Function DevuelveDesdeBD(kCampo As String, Ktabla As String, Kcodigo As String, ValorCodigo As String, Optional Tipo As String, Optional ByRef OtroCampo As String) As String
    Dim Rs As Recordset
    Dim Cad As String
    Dim Aux As String
    
    On Error GoTo EDevuelveDesdeBD
    DevuelveDesdeBD = ""
    Cad = "Select " & kCampo
    If OtroCampo <> "" Then Cad = Cad & ", " & OtroCampo
    Cad = Cad & " FROM " & Ktabla
    Cad = Cad & " WHERE " & Kcodigo & " = "
    If Tipo = "" Then Tipo = "N"
    Select Case Tipo
    Case "N"
        'No hacemos nada
        Cad = Cad & ValorCodigo
    Case "T", "F"
        Cad = Cad & "'" & ValorCodigo & "'"
    Case Else
        MsgBox "Tipo : " & Tipo & " no definido", vbExclamation
        Exit Function
    End Select
    
    
    
    'Creamos el sql
    Set Rs = New ADODB.Recordset
    Rs.Open Cad, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    If Not Rs.EOF Then
        DevuelveDesdeBD = DBLet(Rs.Fields(0))
        If OtroCampo <> "" Then OtroCampo = DBLet(Rs.Fields(1))
    End If
    Rs.Close
    Set Rs = Nothing
    Exit Function
EDevuelveDesdeBD:
        MuestraError Err.Number, "Devuelve DesdeBD.", Err.Description
End Function



''Este metodo sustituye a DevuelveDesdeBD
''Funciona para claves primarias formadas por 2 campos
'Public Function DevuelveDesdeBDnew(vBD As Byte, Ktabla As String, kCampo As String, Kcodigo1 As String, valorCodigo1 As String, Optional tipo1 As String, Optional ByRef otroCampo As String, Optional KCodigo2 As String, Optional ValorCodigo2 As String, Optional tipo2 As String) As String
''IN: vBD --> Base de Datos a la que se accede
'Dim RS As Recordset
'Dim cad As String
'Dim Aux As String
'
'On Error GoTo EDevuelveDesdeBDnew
'    DevuelveDesdeBDnew = ""
'    If valorCodigo1 = "" And ValorCodigo2 = "" Then Exit Function
'    cad = "Select " & kCampo
'    If otroCampo <> "" Then cad = cad & ", " & otroCampo
'    cad = cad & " FROM " & Ktabla
'    cad = cad & " WHERE " & Kcodigo1 & " = "
'    If tipo1 = "" Then tipo1 = "N"
'    Select Case tipo1
'        Case "N"
'            'No hacemos nada
'            If IsNumeric(valorCodigo1) Then
'                cad = cad & Val(valorCodigo1)
'            Else
'                MsgBox "El campo debe ser numérico.", vbExclamation
'                DevuelveDesdeBDnew = "Error"
'                Exit Function
'            End If
'        Case "T", "F"
'            cad = cad & "'" & valorCodigo1 & "'"
'        Case Else
'            MsgBox "Tipo : " & tipo1 & " no definido", vbExclamation
'            Exit Function
'    End Select
'
'    If KCodigo2 <> "" Then
'        cad = cad & " AND " & KCodigo2 & " = "
'        If tipo2 = "" Then tipo2 = "N"
'        Select Case tipo2
'        Case "N"
'            'No hacemos nada
'            If ValorCodigo2 = "" Then
'                cad = cad & "-1"
'            Else
'                cad = cad & Val(ValorCodigo2)
'            End If
'        Case "T"
'            cad = cad & "'" & ValorCodigo2 & "'"
'        Case "F"
'            cad = cad & "'" & Format(ValorCodigo2, FormatoFecha) & "'"
'        Case Else
'            MsgBox "Tipo : " & tipo2 & " no definido", vbExclamation
'            Exit Function
'        End Select
'    End If
'
'
'    'Creamos el sql
'    Set RS = New ADODB.Recordset
'
'    Select Case vBD
'        Case cPTours 'vBD=1: PlannerTours
'            RS.Open cad, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
'        Case cConta 'BD 2: Contabilidad
'            RS.Open cad, ConnConta, adOpenForwardOnly, adLockOptimistic, adCmdText
'        Case 3 'vBD=3: contabilidad distinta a la de la empresa conectada
'            RS.Open cad, ConnAuxCon, adOpenForwardOnly, adLockOptimistic, adCmdText
'    End Select
''    If vBD = cPTours Then 'vBD=1: PlannerTours
''        RS.Open cad, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
''    ElseIf vBD = cConta Then  'BD 2: Contabilidad
''        RS.Open cad, ConnConta, adOpenForwardOnly, adLockOptimistic, adCmdText
''    End If
'
'    If Not RS.EOF Then
'        DevuelveDesdeBDnew = DBLet(RS.Fields(0))
'        If otroCampo <> "" Then otroCampo = DBLet(RS.Fields(1))
'    End If
'    RS.Close
'    Set RS = Nothing
'    Exit Function
'
'EDevuelveDesdeBDnew:
'        MuestraError Err.Number, "Devuelve DesdeBD.", Err.Description
'End Function


'LAURA
'Este metodo sustituye a DevuelveDesdeBD
'Funciona para claves primarias formadas por 3 campos
Public Function DevuelveDesdeBDNew(vBD As Byte, Ktabla As String, kCampo As String, Kcodigo1 As String, valorCodigo1 As String, Optional tipo1 As String, Optional ByRef OtroCampo As String, Optional KCodigo2 As String, Optional ValorCodigo2 As String, Optional tipo2 As String, Optional KCodigo3 As String, Optional ValorCodigo3 As String, Optional tipo3 As String) As String
'IN: vBD --> Base de Datos a la que se accede
Dim Rs As Recordset
Dim Cad As String
Dim Aux As String
    
On Error GoTo EDevuelveDesdeBDnew
    DevuelveDesdeBDNew = ""
'    If valorCodigo1 = "" And ValorCodigo2 = "" Then Exit Function
    Cad = "Select " & kCampo
    If OtroCampo <> "" Then Cad = Cad & ", " & OtroCampo
    Cad = Cad & " FROM " & Ktabla
    If Kcodigo1 <> "" Then
        Cad = Cad & " WHERE " & Kcodigo1 & " = "
        If tipo1 = "" Then tipo1 = "N"
        Select Case tipo1
            Case "N"
                'No hacemos nada
                Cad = Cad & Val(valorCodigo1)
            Case "T"
                Cad = Cad & DBSet(valorCodigo1, "T")
            Case "F"
                Cad = Cad & DBSet(valorCodigo1, "F")
            Case Else
                MsgBox "Tipo : " & tipo1 & " no definido", vbExclamation
                Exit Function
        End Select
    End If
    
    If KCodigo2 <> "" Then
        Cad = Cad & " AND " & KCodigo2 & " = "
        If tipo2 = "" Then tipo2 = "N"
        Select Case tipo2
        Case "N"
            'No hacemos nada
            If ValorCodigo2 = "" Then
                Cad = Cad & "-1"
            Else
                Cad = Cad & Val(ValorCodigo2)
            End If
        Case "T"
'            cad = cad & "'" & ValorCodigo2 & "'"
            Cad = Cad & DBSet(ValorCodigo2, "T")
        Case "F"
            Cad = Cad & "'" & Format(ValorCodigo2, FormatoFecha) & "'"
        Case Else
            MsgBox "Tipo : " & tipo2 & " no definido", vbExclamation
            Exit Function
        End Select
    End If
    
    If KCodigo3 <> "" Then
        Cad = Cad & " AND " & KCodigo3 & " = "
        If tipo3 = "" Then tipo3 = "N"
        Select Case tipo3
        Case "N"
            'No hacemos nada
            If ValorCodigo3 = "" Then
                Cad = Cad & "-1"
            Else
                Cad = Cad & Val(ValorCodigo3)
            End If
        Case "T"
            Cad = Cad & "'" & ValorCodigo3 & "'"
        Case "F"
            Cad = Cad & "'" & Format(ValorCodigo3, FormatoFecha) & "'"
        Case Else
            MsgBox "Tipo : " & tipo3 & " no definido", vbExclamation
            Exit Function
        End Select
    End If
    
    'Creamos el sql
    Set Rs = New ADODB.Recordset
    Select Case vBD
        Case cPTours   'BD 1: Ariges
            Rs.Open Cad, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
        Case cConta    'BD 2: Conta Avnics
            Rs.Open Cad, ConnConta, adOpenForwardOnly, adLockOptimistic, adCmdText
        Case cContaSeg 'BD 3: Conta Seguros
            Rs.Open Cad, ConnContaSeg, adOpenForwardOnly, adLockOptimistic, adCmdText
        Case cContaTel 'BD 4: Conta Telefonia
            Rs.Open Cad, ConnContaTel, adOpenForwardOnly, adLockOptimistic, adCmdText
        Case cContaGas 'BD 6: Conta Gasolinera
            Rs.Open Cad, ConnContaGas, adOpenForwardOnly, adLockOptimistic, adCmdText
        Case cContaFacSoc 'BD 7: Facturas Socio
            Rs.Open Cad, ConnContaFacSoc, adOpenForwardOnly, adLockOptimistic, adCmdText
        Case cContaCV 'BD 8: Facturas Coarval
            Rs.Open Cad, ConnContaCV, adOpenForwardOnly, adLockOptimistic, adCmdText
        Case cContaCVV 'BD 9: Facturas Coarval Ventas
            Rs.Open Cad, ConnContaCVV, adOpenForwardOnly, adLockOptimistic, adCmdText
        '[Monica]12/11/2013: base datos de aridoc
        Case cAridoc
            Rs.Open Cad, ConnAridoc, adOpenForwardOnly, adLockOptimistic, adCmdText
    End Select
    
    If Not Rs.EOF Then
        DevuelveDesdeBDNew = DBLet(Rs.Fields(0))
        If OtroCampo <> "" Then OtroCampo = DBLet(Rs.Fields(1))
    End If
    Rs.Close
    Set Rs = Nothing
    Exit Function
    
EDevuelveDesdeBDnew:
        MuestraError Err.Number, "Devuelve DesdeBD.", Err.Description
End Function


'LAURA
'Este metodo sustituye a DevuelveDesdeBD
'Funciona para claves primarias formadas por 3 campos
Public Function DevuelveDesdeBDNewFac(Ktabla As String, kCampo As String, Kcodigo1 As String, valorCodigo1 As String, Optional tipo1 As String, Optional ByRef OtroCampo As String, Optional KCodigo2 As String, Optional ValorCodigo2 As String, Optional tipo2 As String, Optional KCodigo3 As String, Optional ValorCodigo3 As String, Optional tipo3 As String) As String
'IN: vBD --> Base de Datos a la que se accede
Dim Rs As Recordset
Dim Cad As String
Dim Aux As String
    
On Error GoTo EDevuelveDesdeBDnewFac
    DevuelveDesdeBDNewFac = ""
'    If valorCodigo1 = "" And ValorCodigo2 = "" Then Exit Function
    Cad = "Select " & kCampo
    If OtroCampo <> "" Then Cad = Cad & ", " & OtroCampo
    Cad = Cad & " FROM " & Ktabla
    If Kcodigo1 <> "" Then
        Cad = Cad & " WHERE " & Kcodigo1 & " = "
        If tipo1 = "" Then tipo1 = "N"
    Select Case tipo1
        Case "N"
            'No hacemos nada
            Cad = Cad & Val(valorCodigo1)
        Case "T"
            Cad = Cad & DBSet(valorCodigo1, "T")
        Case "F"
            Cad = Cad & DBSet(valorCodigo1, "F")
        Case Else
            MsgBox "Tipo : " & tipo1 & " no definido", vbExclamation
            Exit Function
    End Select
    End If
    
    If KCodigo2 <> "" Then
        Cad = Cad & " AND " & KCodigo2 & " = "
        If tipo2 = "" Then tipo2 = "N"
        Select Case tipo2
        Case "N"
            'No hacemos nada
            If ValorCodigo2 = "" Then
                Cad = Cad & "-1"
            Else
                Cad = Cad & Val(ValorCodigo2)
            End If
        Case "T"
'            cad = cad & "'" & ValorCodigo2 & "'"
            Cad = Cad & DBSet(ValorCodigo2, "T")
        Case "F"
            Cad = Cad & "'" & Format(ValorCodigo2, FormatoFecha) & "'"
        Case Else
            MsgBox "Tipo : " & tipo2 & " no definido", vbExclamation
            Exit Function
        End Select
    End If
    
    If KCodigo3 <> "" Then
        Cad = Cad & " AND " & KCodigo3 & " = "
        If tipo3 = "" Then tipo3 = "N"
        Select Case tipo3
        Case "N"
            'No hacemos nada
            If ValorCodigo3 = "" Then
                Cad = Cad & "-1"
            Else
                Cad = Cad & Val(ValorCodigo3)
            End If
        Case "T"
            Cad = Cad & "'" & ValorCodigo3 & "'"
        Case "F"
            Cad = Cad & "'" & Format(ValorCodigo3, FormatoFecha) & "'"
        Case Else
            MsgBox "Tipo : " & tipo3 & " no definido", vbExclamation
            Exit Function
        End Select
    End If
    
    
    'Creamos el sql
    Set Rs = New ADODB.Recordset
    Rs.Open Cad, ConnContaFac, adOpenForwardOnly, adLockOptimistic, adCmdText
    
    If Not Rs.EOF Then
        DevuelveDesdeBDNewFac = DBLet(Rs.Fields(0))
        If OtroCampo <> "" Then OtroCampo = DBLet(Rs.Fields(1))
    End If
    Rs.Close
    Set Rs = Nothing
    Exit Function
    
EDevuelveDesdeBDnewFac:
        MuestraError Err.Number, "Devuelve DesdeBD.", Err.Description
End Function




'CESAR
Public Function DevuelveDesdeBDnew2(kBD As Integer, kCampo As String, Ktabla As String, Kcodigo As String, ValorCodigo As String, Optional Tipo As String, Optional num As Byte, Optional ByRef OtroCampo As String) As String
Dim Rs As Recordset
Dim Cad As String
Dim Aux As String
Dim v_aux As Integer
Dim campo As String
Dim Valor As String
Dim tip As String

On Error GoTo EDevuelveDesdeBDnew2
DevuelveDesdeBDnew2 = ""

Cad = "Select " & kCampo
If OtroCampo <> "" Then Cad = Cad & ", " & OtroCampo
Cad = Cad & " FROM " & Ktabla

If Kcodigo <> "" Then Cad = Cad & " where "

For v_aux = 1 To num
    campo = RecuperaValor(Kcodigo, v_aux)
    Valor = RecuperaValor(ValorCodigo, v_aux)
    tip = RecuperaValor(Tipo, v_aux)
        
    Cad = Cad & campo & "="
    If tip = "" Then Tipo = "N"
    
    Select Case tip
            Case "N"
                'No hacemos nada
                Cad = Cad & Valor
            Case "T", "F"
                Cad = Cad & "'" & Valor & "'"
            Case Else
                MsgBox "Tipo : " & tip & " no definido", vbExclamation
            Exit Function
    End Select
    
    If v_aux < num Then Cad = Cad & " AND "
  Next v_aux

'Creamos el sql
Set Rs = New ADODB.Recordset
Select Case kBD
    Case 1
        Rs.Open Cad, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
End Select

If Not Rs.EOF Then
    DevuelveDesdeBDnew2 = DBLet(Rs.Fields(0))
    If OtroCampo <> "" Then OtroCampo = DBLet(Rs.Fields(1))
Else
     If OtroCampo <> "" Then OtroCampo = ""
End If
Rs.Close
Set Rs = Nothing
Exit Function
EDevuelveDesdeBDnew2:
    MuestraError Err.Number, "Devuelve DesdeBDnew2.", Err.Description
End Function


Public Function EsEntero(Texto As String) As Boolean
Dim i As Integer
Dim C As Integer
Dim L As Integer
Dim res As Boolean

    res = True
    EsEntero = False

    If Not IsNumeric(Texto) Then
        res = False
    Else
        'Vemos si ha puesto mas de un punto
        C = 0
        L = 1
        Do
            i = InStr(L, Texto, ".")
            If i > 0 Then
                L = i + 1
                C = C + 1
            End If
        Loop Until i = 0
        If C > 1 Then res = False
        
        'Si ha puesto mas de una coma y no tiene puntos
        If C = 0 Then
            L = 1
            Do
                i = InStr(L, Texto, ",")
                If i > 0 Then
                    L = i + 1
                    C = C + 1
                End If
            Loop Until i = 0
            If C > 1 Then res = False
        End If
        
    End If
        EsEntero = res
End Function

Public Function TransformaPuntosComas(Cadena As String) As String
    Dim i As Integer
    Do
        i = InStr(1, Cadena, ".")
        If i > 0 Then
            Cadena = Mid(Cadena, 1, i - 1) & "," & Mid(Cadena, i + 1)
        End If
        Loop Until i = 0
    TransformaPuntosComas = Cadena
End Function

Public Sub InicializarFormatos()
    FormatoFecha = "yyyy-mm-dd"
    FormatoHora = "hh:mm:ss"
'    FormatoFechaHora = "yyyy-mm-dd hh:mm:ss"
    FormatoImporte = "#,###,###,##0.00"  'Decimal(12,2)
    FormatoPrecio = "##,##0.000"  'Decimal(8,3) antes decimal(10,4)
'    FormatoCantidad = "##,###,##0.00"   'Decimal(10,2)
    FormatoPorcen = "##0.00" 'Decima(5,2) para porcentajes
    
    FormatoDec10d2 = "##,###,##0.00"   'Decimal(10,2)
    FormatoDec10d3 = "##,###,##0.000"   'Decimal(10,3)
    FormatoDec5d4 = "0.0000"   'Decimal(5,4)
    FormatoDec10d4 = "###,##0.0000" ' Decimal(10,4)
    FormatoExp = "0000000000"
'    FormatoKms = "#,##0.00##" 'Decimal(8,4)
End Sub


Public Sub AccionesCerrar()
'cosas que se deben hacen cuando finaliza la aplicacion
    On Error Resume Next
    
    'cerrar clases q estan abiertas durante la ejecucion
    Set vEmpresa = Nothing
    Set vEmpresaSeg = Nothing
    Set vEmpresaTel = Nothing
    Set vEmpresaCV = Nothing
    Set vEmpresaCVV = Nothing
    Set vSesion = Nothing
    
'    Set vParam = Nothing
'    Set vParamAplic = Nothing
'    Set vParamConta = Nothing
    
    
    'Cerrar Conexiones a bases de datos
    conn.Close
    ConnConta.Close
    ConnContaSeg.Close
    ConnContaTel.Close
    ConnContaCV.Close
    ConnContaCVV.Close
    Set conn = Nothing
    Set ConnConta = Nothing
    Set ConnContaSeg = Nothing
    Set ConnContaTel = Nothing
    Set ConnContaCV = Nothing
    Set ConnContaCVV = Nothing
    
    If Err.Number <> 0 Then Err.Clear
End Sub

Public Sub AccionesCerrarContabilidades()
'cosas que se deben hacen cuando finaliza la aplicacion
    On Error Resume Next
    
    'cerrar clases q estan abiertas durante la ejecucion
    Set vEmpresa = Nothing
    Set vEmpresaSeg = Nothing
    Set vEmpresaTel = Nothing
    Set vEmpresaCV = Nothing
    Set vEmpresaCVV = Nothing
    Set vEmpresaGas = Nothing
    Set vEmpresaFacSoc = Nothing
    
'    Set vParam = Nothing
'    Set vParamAplic = Nothing
'    Set vParamConta = Nothing
    
    'Cerrar Conexiones a bases de datos
    ConnConta.Close
    ConnContaSeg.Close
    ConnContaTel.Close
    ConnContaCV.Close
    ConnContaCVV.Close
    ConnContaGas.Close
    ConnContaFacSoc.Close
    
    Set ConnConta = Nothing
    Set ConnContaSeg = Nothing
    Set ConnContaTel = Nothing
    Set ConnContaCV = Nothing
    Set ConnContaCVV = Nothing
    Set ConnContaGas = Nothing
    Set ConnContaFacSoc = Nothing
    
    If Err.Number <> 0 Then Err.Clear
End Sub



Public Function OtrosPCsContraAplicacion() As String
Dim MiRS As Recordset
Dim Cad As String
Dim Equipo As String

    Set MiRS = New ADODB.Recordset
    Cad = "show processlist"
    MiRS.Open Cad, conn, adOpenKeyset, adLockOptimistic, adCmdText
    Cad = ""
    While Not MiRS.EOF
        If MiRS.Fields(3) = vSesion.CadenaConexion Then
            Equipo = MiRS.Fields(2)
            'Primero quitamos los dos puntos del puerot
            NumRegElim = InStr(1, Equipo, ":")
            If NumRegElim > 0 Then Equipo = Mid(Equipo, 1, NumRegElim - 1)
            
            'El punto del dominio
            NumRegElim = InStr(1, Equipo, ".")
            If NumRegElim > 0 Then Equipo = Mid(Equipo, 1, NumRegElim - 1)
            
            Equipo = UCase(Equipo)
            
            If Equipo <> vSesion.Codusu Then
                    If Equipo <> "LOCALHOST" Then
                        If InStr(1, Cad, Equipo & "|") = 0 Then Cad = Cad & Equipo & "|"
                    End If
            End If
        End If
        'Siguiente
        MiRS.MoveNext
    Wend
    NumRegElim = 0
    MiRS.Close
    Set MiRS = Nothing
    OtrosPCsContraAplicacion = Cad
End Function


Public Function UsuariosConectados() As Boolean
Dim i As Integer
Dim Cad As String
Dim metag As String
Dim Sql As String
Cad = OtrosPCsContraAplicacion
UsuariosConectados = False
If Cad <> "" Then
    UsuariosConectados = True
    i = 1
    metag = "Los siguientes PC's están conectados a: " & vEmpresa.nomEmpre & " (" & vSesion.CadenaConexion & ")" & vbCrLf & vbCrLf
    Do
        Sql = RecuperaValor(Cad, i)
        If Sql <> "" Then
            metag = metag & "    - " & Sql & vbCrLf
            i = i + 1
        End If
    Loop Until Sql = ""
    MsgBox metag, vbExclamation
End If
End Function

'--------------------------------------------------------------------
'-------------------------------------------------------------------
'Para el envio de los mails
Public Function PrepararCarpetasEnvioMail(Optional NoBorrar As Boolean) As Boolean
    On Error GoTo EPrepararCarpetasEnvioMail
    PrepararCarpetasEnvioMail = False

    If Dir(App.path & "\temp", vbDirectory) = "" Then
        MkDir App.path & "\temp"
    Else
        If Not NoBorrar Then
            If Dir(App.path & "\temp\*.*", vbArchive) <> "" Then Kill App.path & "\temp\*.*"
        End If
    End If


    PrepararCarpetasEnvioMail = True
    Exit Function
EPrepararCarpetasEnvioMail:
    MuestraError Err.Number, "", "Preparar Carpetas temporal para envio eMail. Borrando tmp "
End Function



Public Function ejecutar(ByRef Sql As String, OcultarMsg As Boolean) As Boolean
    On Error Resume Next
    conn.Execute Sql
    If Err.Number <> 0 Then
        If Not OcultarMsg Then MuestraError Err.Number, Err.Description, Sql
        ejecutar = False
    Else
        ejecutar = True
    End If
End Function


