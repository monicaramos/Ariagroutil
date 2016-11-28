Attribute VB_Name = "ModContabilizar"
' copia del ariges

Option Explicit



'===================================================================================
'CONTABILIZAR FACTURAS:
'Modulo para el traspaso de registros de cabecera y lineas de tablas de FACTURACION
'A las tablas de FACTURACION de Contabilidad
'====================================================================================

Private BaseImp As Currency
Private IvaImp As Currency

Private CCoste As String


Private vTipoIva(2) As Currency
Private vPorcIva(2) As Currency
Private vPorcRec(2) As Currency
Private vBaseIva(2) As Currency
Private vImpIva(2) As Currency
Private vImpRec(2) As Currency





Public Function CrearTMPFacturas(cadTABLA As String, cadwhere As String, Optional Facturas As Boolean, Optional Telefono As Boolean) As Boolean
'Crea una temporal donde inserta la clave primaria de las
'facturas seleccionadas para facturar y trabaja siempre con ellas
' facturas indica si viene de facturas varias o de telefonia
Dim SQL As String
    
    On Error GoTo ECrear
    
    CrearTMPFacturas = False
    
    SQL = "CREATE TEMPORARY TABLE tmpfactu ( "
    If Facturas Then
        SQL = SQL & "codsecci smallint(2) NOT NULL default 0,"
    End If
    SQL = SQL & "numserie char(3) NOT NULL default '',"
    SQL = SQL & "numfactu mediumint(7) unsigned NOT NULL default '0',"
    SQL = SQL & "fecfactu date NOT NULL default '0000-00-00') "
    conn.Execute SQL
     
    If Facturas Then
        SQL = "SELECT codsecci, letraser, numfactu, fecfactu"
    Else
        If Telefono Then
            SQL = "SELECT numserie, numfactu, fecfactu"
        Else
            SQL = "SELECT letraser, numfactu, fecfactu"
        End If
    End If
    SQL = SQL & " FROM " & cadTABLA
    SQL = SQL & " WHERE " & cadwhere
    SQL = " INSERT INTO tmpfactu " & SQL
    conn.Execute SQL

    CrearTMPFacturas = True
    
ECrear:
     If Err.Number <> 0 Then
        CrearTMPFacturas = False
        'Borrar la tabla temporal
        SQL = " DROP TABLE IF EXISTS tmpfactu;"
        conn.Execute SQL
    End If
End Function

Public Function CrearTMPFacturasCV(cadTABLA As String, cadwhere As String, Optional Facturas As Boolean, Optional Telefono As Boolean) As Boolean
'Crea una temporal donde inserta la clave primaria de las
'facturas seleccionadas para facturar y trabaja siempre con ellas
' facturas indica si viene de facturas varias o de telefonia
Dim SQL As String
    
    On Error GoTo ECrear
    
    CrearTMPFacturasCV = False
    
    SQL = "CREATE TEMPORARY TABLE tmpfactu ( "
    If Facturas Then
        SQL = SQL & "codsecci smallint(2) NOT NULL default 0,"
    End If
    SQL = SQL & "numserie char(3) NOT NULL default '',"
    SQL = SQL & "numfactu varchar(10) NOT NULL default '',"
    SQL = SQL & "fecfactu date NOT NULL default '0000-00-00') "
    conn.Execute SQL
     
    If Facturas Then
        SQL = "SELECT codsecci, letraser, numfactu, fecfactu"
    Else
        If Telefono Then
            SQL = "SELECT numserie, numfactu, fecfactu"
        Else
            SQL = "SELECT letraser, numfactu, fecfactu"
        End If
    End If
    SQL = SQL & " FROM " & cadTABLA
    SQL = SQL & " WHERE " & cadwhere
    SQL = " INSERT INTO tmpfactu " & SQL
    conn.Execute SQL

    CrearTMPFacturasCV = True
    
ECrear:
     If Err.Number <> 0 Then
        CrearTMPFacturasCV = False
        'Borrar la tabla temporal
        SQL = " DROP TABLE IF EXISTS tmpfactu;"
        conn.Execute SQL
    End If
End Function


Public Function CrearTMPFacturasProveedor(cadTABLA As String, cadwhere As String) As Boolean
'Crea una temporal donde inserta la clave primaria de las
'facturas seleccionadas para facturar y trabaja siempre con ellas
' facturas indica si viene de facturas varias o de telefonia
Dim SQL As String
    
    On Error GoTo ECrear
    
    CrearTMPFacturasProveedor = False
    
    SQL = "CREATE TEMPORARY TABLE tmpfactu ( "
    SQL = SQL & "numfactu mediumint(7) unsigned NOT NULL default '0',"
    SQL = SQL & "fecfactu date NOT NULL default '0000-00-00', "
    SQL = SQL & "codmacta varchar(10) NOT NULL ) "
    
    conn.Execute SQL
     
    SQL = "SELECT numfactu, fecfactu, codmacta"
    SQL = SQL & " FROM " & cadTABLA
    SQL = SQL & " WHERE " & cadwhere
    SQL = " INSERT INTO tmpfactu " & SQL
    conn.Execute SQL

    CrearTMPFacturasProveedor = True
    
ECrear:
     If Err.Number <> 0 Then
        CrearTMPFacturasProveedor = False
        'Borrar la tabla temporal
        SQL = " DROP TABLE IF EXISTS tmpfactu;"
        conn.Execute SQL
    End If
End Function



Public Sub BorrarTMPFacturas()
On Error Resume Next

    conn.Execute " DROP TABLE IF EXISTS tmpfactu;"
    If Err.Number <> 0 Then Err.Clear
End Sub


Public Function CrearTMPErrFact(cadTABLA As String) As Boolean
'Crea una temporal donde insertara la clave primaria de las
'facturas erroneas al facturar
Dim SQL As String
    
    On Error GoTo ECrear
    
    CrearTMPErrFact = False
    
    SQL = "CREATE TEMPORARY TABLE tmperrfac ( "
    If cadTABLA = "schfac" Or cadTABLA = "telmovil" Then
        SQL = SQL & "codtipom char(3) NOT NULL default '',"
        SQL = SQL & "numfactu mediumint(7) unsigned NOT NULL default '0',"
    Else
        If cadTABLA = "cvfacturas" Then
            SQL = SQL & "codtipom char(3) NOT NULL default '',"
            SQL = SQL & "numfactu varchar(10), "
        End If
    End If
    SQL = SQL & "fecfactu date NOT NULL default '0000-00-00', "
    SQL = SQL & "error varchar(400) NULL )"
    conn.Execute SQL
     
    CrearTMPErrFact = True
    
ECrear:
     If Err.Number <> 0 Then
        CrearTMPErrFact = False
        'Borrar la tabla temporal
        SQL = " DROP TABLE IF EXISTS tmperrfac;"
        conn.Execute SQL
    End If
End Function


Public Function CrearTMPErrComprob() As Boolean
'Crea una temporal donde insertara la clave primaria de las
'facturas erroneas al facturar
Dim SQL As String
    
    On Error GoTo ECrear
    
    CrearTMPErrComprob = False
    
    SQL = "CREATE TEMPORARY TABLE tmperrcomprob ( "
    SQL = SQL & "error varchar(100) NULL )"
    conn.Execute SQL
     
    CrearTMPErrComprob = True
    
ECrear:
     If Err.Number <> 0 Then
        CrearTMPErrComprob = False
        'Borrar la tabla temporal
        SQL = " DROP TABLE IF EXISTS tmperrcomprob;"
        conn.Execute SQL
    End If
End Function



Public Sub BorrarTMPErrFact()
On Error Resume Next
    conn.Execute " DROP TABLE IF EXISTS tmperrfac;"
    If Err.Number <> 0 Then Err.Clear
End Sub

Public Sub BorrarTMPErrComprob()
On Error Resume Next
    conn.Execute " DROP TABLE IF EXISTS tmperrcomprob;"
    If Err.Number <> 0 Then Err.Clear
End Sub

Public Sub BorrarTMPAsiento()
On Error Resume Next
    conn.Execute " DROP TABLE IF EXISTS tmpasien;"
    If Err.Number <> 0 Then Err.Clear
End Sub


Public Function ComprobarLetraSerie(bd As Byte) As Boolean
'Para Facturas VENTA a clientes
'Comprueba que la letra del serie del tipo de movimiento es  correcta
Dim SQL As String
Dim Rs As ADODB.Recordset
Dim RSconta As ADODB.Recordset
Dim b As Boolean
Dim Cad As String, devuelve As String

On Error GoTo EComprobarLetra

    ComprobarLetraSerie = False
    
    'Comprobar que existe la letra de serie en contabilidad
        'cargamos el RSConta con la tabla contadores de BD: Contabilidad
        'donde estan todas las letra de serie que existen en la contabilidad
    SQL = "Select distinct tiporegi from contadores"
    Set RSconta = New ADODB.Recordset
    Select Case bd
        Case cConta
            RSconta.Open SQL, ConnConta, adOpenDynamic, adLockPessimistic, adCmdText
        Case cContaSeg
            RSconta.Open SQL, ConnContaSeg, adOpenDynamic, adLockPessimistic, adCmdText
        Case cContaTel
            RSconta.Open SQL, ConnContaTel, adOpenDynamic, adLockPessimistic, adCmdText
    End Select
    
    If RSconta.EOF Then
        RSconta.Close
        Set RSconta = Nothing
        Exit Function
    End If
        

    'obtenemos los distintos tipos de movimiento que vamos a contabilizar
    'de las facturas seleccionadas
    SQL = "select distinct numserie from tmpfactu "

    Set Rs = New ADODB.Recordset
    Rs.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    Cad = ""
    b = True
    While Not Rs.EOF 'And b
        'comprobar que todas las letras serie existen en Arigasol
'        Sql = "letraser"
'        devuelve = DevuelveDesdeBD("letraser", "stipom", "letraser", DBLet(RS!numserie), "T", Sql)
'        If devuelve = "" Then
'            b = False
'            cad = RS!numserie & " en BD de Ariagroutil."
'            InsertarError "No existe la letra de serie " & cad
'        Else
            'comprobar que todas las letras serie existen en la contabilidad
            devuelve = "tiporegi= '" & Trim(Rs!numserie) & "'" '& devuelve & "'"
            RSconta.MoveFirst
            RSconta.Find (devuelve), , adSearchForward
            If RSconta.EOF Then
                'no encontrado
                b = False
                Cad = SQL & " en BD de Contabilidad."
                InsertarError "No existe la letra de serie " & Cad
            End If
'        End If
        If b Then Cad = Cad & DBSet(Trim(Rs!numserie), "T") & ","
        Rs.MoveNext
    Wend
    Rs.Close
    Set Rs = Nothing
    RSconta.Close
    Set RSconta = Nothing
    
    If Not b Then 'Hay algun movimiento que no existe
        devuelve = "No existe el tipo de movimiento: " & Cad & vbCrLf
        devuelve = devuelve & "Consulte con el administrador."
'            MsgBox devuelve, vbExclamation
        Exit Function
    End If
    
'    'Todos los Tipo de movimiento existen
'    If cad <> "" Then
'        cad = Mid(cad, 1, Len(cad) - 1) 'quitamos ult. coma
'
'        'miramos si hay algun movimiento de factura que la letra serie sea nulo
'        Sql = "select count(*) from stipom "
'        Sql = Sql & "where letraser IN (" & cad & ") and (isnull(letraser) or letraser='')"
'        If RegistrosAListar(Sql) > 0 Then
'            Sql = "Hay algun tipo de movimiento de Facturación que no tiene letra serie." & vbCrLf
'            Sql = Sql & "Comprobar en la tabla de tipos de movimiento: " & cad
'            InsertarError Sql
''                MsgBox sql, vbExclamation
'            Exit Function
'        End If
'    End If
    ComprobarLetraSerie = True

EComprobarLetra:
    If Err.Number <> 0 Then
        MuestraError Err.Number, "Comprobar Letra Serie", Err.Description
    End If
End Function


Public Function ComprobarLetraSerieFac() As Boolean
'Para Facturas VENTA a clientes
'Comprueba que la letra del serie del tipo de movimiento es  correcta
Dim SQL As String
Dim Rs As ADODB.Recordset
Dim RSconta As ADODB.Recordset
Dim b As Boolean
Dim Cad As String, devuelve As String

On Error GoTo EComprobarLetraFac

    ComprobarLetraSerieFac = False
    
    'Comprobar que existe la letra de serie en contabilidad
        'cargamos el RSConta con la tabla contadores de BD: Contabilidad
        'donde estan todas las letra de serie que existen en la contabilidad
    SQL = "Select distinct tiporegi from contadores"
    Set RSconta = New ADODB.Recordset
    RSconta.Open SQL, ConnContaFac, adOpenDynamic, adLockPessimistic, adCmdText
    
    If RSconta.EOF Then
        RSconta.Close
        Set RSconta = Nothing
        Exit Function
    End If
        

    'obtenemos los distintos tipos de movimiento que vamos a contabilizar
    'de las facturas seleccionadas
    SQL = "select distinct numserie from tmpfactu "

    Set Rs = New ADODB.Recordset
    Rs.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    Cad = ""
    b = True
    While Not Rs.EOF 'And b
        'comprobar que todas las letras serie existen en la contabilidad
        devuelve = "tiporegi= '" & Trim(Rs!numserie) & "'" '& devuelve & "'"
        RSconta.MoveFirst
        RSconta.Find (devuelve), , adSearchForward
        If RSconta.EOF Then
            'no encontrado
            b = False
            Cad = DBLet(Rs!numserie, "T") & " en BD de Contabilidad."
            InsertarError "No existe la letra de serie " & Cad
        End If
        Rs.MoveNext
    Wend
    Rs.Close
    Set Rs = Nothing
    RSconta.Close
    Set RSconta = Nothing
    
    ComprobarLetraSerieFac = b '26/11/2009  antes true

EComprobarLetraFac:
    If Err.Number <> 0 Then
        MuestraError Err.Number, "Comprobar Letra Serie", Err.Description
    End If
End Function


Public Function ComprobarLetraSerieGas() As Boolean
'Para Facturas Gasolinera a Socios
'Comprueba que la letra del serie del tipo de movimiento es  correcta
Dim SQL As String
Dim Rs As ADODB.Recordset
Dim RSconta As ADODB.Recordset
Dim b As Boolean
Dim Cad As String, devuelve As String

On Error GoTo EComprobarLetraGas

    ComprobarLetraSerieGas = False
    
    'Comprobar que existe la letra de serie en contabilidad
        'cargamos el RSConta con la tabla contadores de BD: Contabilidad
        'donde estan todas las letra de serie que existen en la contabilidad
    SQL = "Select distinct tiporegi from contadores"
    Set RSconta = New ADODB.Recordset
    RSconta.Open SQL, ConnContaGas, adOpenDynamic, adLockPessimistic, adCmdText
    
    If RSconta.EOF Then
        RSconta.Close
        Set RSconta = Nothing
        Exit Function
    End If
        

    'obtenemos los distintos tipos de movimiento que vamos a contabilizar
    'de las facturas seleccionadas
    SQL = "select distinct numserie from tmpfactu "

    Set Rs = New ADODB.Recordset
    Rs.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    Cad = ""
    b = True
    While Not Rs.EOF 'And b
        'comprobar que todas las letras serie existen en Arigasol
'        Sql = "letraser"
'        devuelve = DevuelveDesdeBD("letraser", "stipom", "letraser", DBLet(RS!numserie), "T", Sql)
'        If devuelve = "" Then
'            b = False
'            cad = RS!numserie & " en BD de Ariagroutil."
'            InsertarError "No existe la letra de serie " & cad
'        Else
            'comprobar que todas las letras serie existen en la contabilidad
            devuelve = "tiporegi= '" & Trim(Rs!numserie) & "'" '& devuelve & "'"
            RSconta.MoveFirst
            RSconta.Find (devuelve), , adSearchForward
            If RSconta.EOF Then
                'no encontrado
                b = False
                Cad = SQL & " en BD de Contabilidad."
                InsertarError "No existe la letra de serie " & Cad
            End If
'        End If
        If b Then Cad = Cad & DBSet(Trim(Rs!numserie), "T") & ","
        Rs.MoveNext
    Wend
    Rs.Close
    Set Rs = Nothing
    RSconta.Close
    Set RSconta = Nothing
    
    If Not b Then 'Hay algun movimiento que no existe
        devuelve = "No existe el tipo de movimiento: " & Cad & vbCrLf
        devuelve = devuelve & "Consulte con el administrador."
'            MsgBox devuelve, vbExclamation
        Exit Function
    End If
    
    ComprobarLetraSerieGas = True

EComprobarLetraGas:
    If Err.Number <> 0 Then
        MuestraError Err.Number, "Comprobar Letra Serie", Err.Description
    End If
End Function



Public Function ComprobarLetraSerieCV(tipo As Byte) As Boolean
'Para Facturas Gasolinera a Socios
'Comprueba que la letra del serie del tipo de movimiento es  correcta
Dim SQL As String
Dim Rs As ADODB.Recordset
Dim RSconta As ADODB.Recordset
Dim b As Boolean
Dim Cad As String, devuelve As String

On Error GoTo EComprobarLetraCV

    ComprobarLetraSerieCV = False
    
    'Comprobar que existe la letra de serie en contabilidad
        'cargamos el RSConta con la tabla contadores de BD: Contabilidad
        'donde estan todas las letra de serie que existen en la contabilidad
    SQL = "Select distinct tiporegi from contadores"
    Set RSconta = New ADODB.Recordset
    If tipo = 0 Then
        RSconta.Open SQL, ConnContaCVV, adOpenDynamic, adLockPessimistic, adCmdText
    Else
        RSconta.Open SQL, ConnContaCV, adOpenDynamic, adLockPessimistic, adCmdText
    End If
    
    If RSconta.EOF Then
        RSconta.Close
        Set RSconta = Nothing
        Exit Function
    End If
        

    'obtenemos los distintos tipos de movimiento que vamos a contabilizar
    'de las facturas seleccionadas
    SQL = "select distinct numserie from tmpfactu "

    Set Rs = New ADODB.Recordset
    Rs.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    Cad = ""
    b = True
    While Not Rs.EOF 'And b
        'comprobar que todas las letras serie existen en Arigasol
'        Sql = "letraser"
'        devuelve = DevuelveDesdeBD("letraser", "stipom", "letraser", DBLet(RS!numserie), "T", Sql)
'        If devuelve = "" Then
'            b = False
'            cad = RS!numserie & " en BD de Ariagroutil."
'            InsertarError "No existe la letra de serie " & cad
'        Else
            'comprobar que todas las letras serie existen en la contabilidad
            devuelve = "tiporegi= '" & Trim(Rs!numserie) & "'" '& devuelve & "'"
            RSconta.MoveFirst
            RSconta.Find (devuelve), , adSearchForward
            SQL = Rs!numserie
            If RSconta.EOF Then
                'no encontrado
                b = False
                Cad = SQL & " en BD de Contabilidad."
                InsertarError "No existe la letra de serie " & Cad
            End If
'        End If
        If b Then Cad = Cad & DBSet(Trim(Rs!numserie), "T") & ","
        Rs.MoveNext
    Wend
    Rs.Close
    Set Rs = Nothing
    RSconta.Close
    Set RSconta = Nothing
    
    If Not b Then 'Hay algun movimiento que no existe
        devuelve = "No existe el tipo de movimiento: " & Cad & vbCrLf
        devuelve = devuelve & "Consulte con el administrador."
'            MsgBox devuelve, vbExclamation
        Exit Function
    End If
    
    ComprobarLetraSerieCV = True

EComprobarLetraCV:
    If Err.Number <> 0 Then
        MuestraError Err.Number, "Comprobar Letra Serie", Err.Description
    End If
End Function







Public Function ComprobarNumFacturas(bd As Byte, cadWConta) As Boolean
'Comprobar que no exista ya en la contabilidad un nº de factura para la fecha que
'vamos a contabilizar
Dim SQL As String
Dim Rs As ADODB.Recordset
Dim RSconta As ADODB.Recordset
Dim b As Boolean

    On Error GoTo ECompFactu

    ComprobarNumFacturas = False
    
    SQL = "SELECT numserie,codfaccl,anofaccl FROM cabfact "
    SQL = SQL & " WHERE " & cadWConta
    
    Set RSconta = New ADODB.Recordset
    Select Case bd
        Case cConta
            RSconta.Open SQL, ConnConta, adOpenForwardOnly, adLockPessimistic, adCmdText
        Case cContaSeg
            RSconta.Open SQL, ConnContaSeg, adOpenForwardOnly, adLockPessimistic, adCmdText
        Case cContaTel
            RSconta.Open SQL, ConnContaTel, adOpenForwardOnly, adLockPessimistic, adCmdText
        Case cContaGas
            RSconta.Open SQL, ConnContaGas, adOpenForwardOnly, adLockPessimistic, adCmdText
        Case cContaCV
            RSconta.Open SQL, ConnContaCV, adOpenForwardOnly, adLockPessimistic, adCmdText
        Case cContaCVV
            RSconta.Open SQL, ConnContaCVV, adOpenForwardOnly, adLockPessimistic, adCmdText
    End Select

    If Not RSconta.EOF Then
        'Seleccionamos las distintas facturas que vamos a facturar
        SQL = "SELECT DISTINCT tmpfactu.numserie,tmpfactu.numfactu,tmpfactu.fecfactu "
        SQL = SQL & " FROM tmpfactu "
        
        Set Rs = New ADODB.Recordset
        Rs.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        b = True
        While Not Rs.EOF 'And b
' quitado el 12022007
'            SQL = "(numserie= " & DBSet(RS!letraser, "T") & " AND codfaccl=" & DBSet(RS!numfactu, "N") & " AND anofaccl=" & Year(RS!fecfactu) & ")"
'            If SituarRSetMULTI(RSconta, SQL) Then
            SQL = ""
            Select Case bd
                Case cConta
                    SQL = DevuelveDesdeBDNew(cConta, "cabfact", "codfaccl", "codfaccl", Rs!numfactu, "N", , "numserie", Trim(Rs!numserie), "T", "anofaccl", Year(Rs!fecfactu), "N")
                Case cContaSeg
                    SQL = DevuelveDesdeBDNew(cContaSeg, "cabfact", "codfaccl", "codfaccl", Rs!numfactu, "N", , "numserie", Trim(Rs!numserie), "T", "anofaccl", Year(Rs!fecfactu), "N")
                Case cContaTel
                    SQL = DevuelveDesdeBDNew(cContaTel, "cabfact", "codfaccl", "codfaccl", Rs!numfactu, "N", , "numserie", Trim(Rs!numserie), "T", "anofaccl", Year(Rs!fecfactu), "N")
                Case cContaGas
                    SQL = DevuelveDesdeBDNew(cContaGas, "cabfact", "codfaccl", "codfaccl", Rs!numfactu, "N", , "numserie", Trim(Rs!numserie), "T", "anofaccl", Year(Rs!fecfactu), "N")
                Case cContaCV
                    SQL = DevuelveDesdeBDNew(cContaCV, "cabfact", "codfaccl", "codfaccl", Rs!numfactu, "N", , "numserie", Trim(Rs!numserie), "T", "anofaccl", Year(Rs!fecfactu), "N")
                Case cContaCVV
                    SQL = DevuelveDesdeBDNew(cContaCVV, "cabfact", "codfaccl", "codfaccl", Rs!numfactu, "N", , "numserie", Trim(Rs!numserie), "T", "anofaccl", Year(Rs!fecfactu), "N")
            End Select
            If SQL <> "" Then
                b = False
                SQL = "          Nº Fac.: " & Format(Rs!numfactu, "0000000") & vbCrLf
                SQL = SQL & "          Fecha: " & Rs!fecfactu
                
                SQL = "Ya existe la factura: " & vbCrLf & SQL
                InsertarError SQL
            
            End If
            Rs.MoveNext
        Wend
        Rs.Close
        Set Rs = Nothing
        
        If Not b Then
            SQL = "Ya existe la factura: " & vbCrLf & SQL
            SQL = "Comprobando Nº Facturas en Contabilidad...       " & vbCrLf & vbCrLf & SQL
            
            'MsgBox sql, vbExclamation
            ComprobarNumFacturas = False
        Else
            ComprobarNumFacturas = True
        End If
    Else
        ComprobarNumFacturas = True
    End If
    RSconta.Close
    Set RSconta = Nothing
    
ECompFactu:
     If Err.Number <> 0 Then
        MuestraError Err.Number, "Comprobar Nº Facturas", Err.Description
    End If
End Function



Public Function ComprobarNumFacturasContaNueva(bd As Byte, cadWConta) As Boolean
'Comprobar que no exista ya en la contabilidad un nº de factura para la fecha que
'vamos a contabilizar
Dim SQL As String
Dim Rs As ADODB.Recordset
Dim RSconta As ADODB.Recordset
Dim b As Boolean

    On Error GoTo ECompFactu

    ComprobarNumFacturasContaNueva = False
    
    SQL = "SELECT numserie,numfactu,anofactu FROM factcli "
    SQL = SQL & " WHERE " & cadWConta
    
    Set RSconta = New ADODB.Recordset
    Select Case bd
        Case cConta
            RSconta.Open SQL, ConnConta, adOpenForwardOnly, adLockPessimistic, adCmdText
        Case cContaSeg
            RSconta.Open SQL, ConnContaSeg, adOpenForwardOnly, adLockPessimistic, adCmdText
        Case cContaTel
            RSconta.Open SQL, ConnContaTel, adOpenForwardOnly, adLockPessimistic, adCmdText
        Case cContaGas
            RSconta.Open SQL, ConnContaGas, adOpenForwardOnly, adLockPessimistic, adCmdText
        Case cContaCV
            RSconta.Open SQL, ConnContaCV, adOpenForwardOnly, adLockPessimistic, adCmdText
        Case cContaCVV
            RSconta.Open SQL, ConnContaCVV, adOpenForwardOnly, adLockPessimistic, adCmdText
    End Select

    If Not RSconta.EOF Then
        'Seleccionamos las distintas facturas que vamos a facturar
        SQL = "SELECT DISTINCT tmpfactu.numserie,tmpfactu.numfactu,tmpfactu.fecfactu "
        SQL = SQL & " FROM tmpfactu "
        
        Set Rs = New ADODB.Recordset
        Rs.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        b = True
        While Not Rs.EOF 'And b
' quitado el 12022007
'            SQL = "(numserie= " & DBSet(RS!letraser, "T") & " AND codfaccl=" & DBSet(RS!numfactu, "N") & " AND anofaccl=" & Year(RS!fecfactu) & ")"
'            If SituarRSetMULTI(RSconta, SQL) Then
            SQL = ""
            Select Case bd
                Case cConta
                    SQL = DevuelveDesdeBDNew(cConta, "factcli", "numfactu", "numfactu", Rs!numfactu, "N", , "numserie", Trim(Rs!numserie), "T", "anofactu", Year(Rs!fecfactu), "N")
                Case cContaSeg
                    SQL = DevuelveDesdeBDNew(cContaSeg, "factcli", "numfactu", "numfactu", Rs!numfactu, "N", , "numserie", Trim(Rs!numserie), "T", "anofactu", Year(Rs!fecfactu), "N")
                Case cContaTel
                    SQL = DevuelveDesdeBDNew(cContaTel, "factcli", "numfactu", "numfactu", Rs!numfactu, "N", , "numserie", Trim(Rs!numserie), "T", "anofactu", Year(Rs!fecfactu), "N")
                Case cContaGas
                    SQL = DevuelveDesdeBDNew(cContaGas, "factcli", "numfactu", "numfactu", Rs!numfactu, "N", , "numserie", Trim(Rs!numserie), "T", "anofactu", Year(Rs!fecfactu), "N")
                Case cContaCV
                    SQL = DevuelveDesdeBDNew(cContaCV, "factcli", "numfactu", "numfactu", Rs!numfactu, "N", , "numserie", Trim(Rs!numserie), "T", "anofactu", Year(Rs!fecfactu), "N")
                Case cContaCVV
                    SQL = DevuelveDesdeBDNew(cContaCVV, "factcli", "numfactu", "numfactu", Rs!numfactu, "N", , "numserie", Trim(Rs!numserie), "T", "anofactu", Year(Rs!fecfactu), "N")
            End Select
            If SQL <> "" Then
                b = False
                SQL = "          Nº Fac.: " & Format(Rs!numfactu, "0000000") & vbCrLf
                SQL = SQL & "          Fecha: " & Rs!fecfactu
                
                SQL = "Ya existe la factura: " & vbCrLf & SQL
                InsertarError SQL
            
            End If
            Rs.MoveNext
        Wend
        Rs.Close
        Set Rs = Nothing
        
        If Not b Then
            SQL = "Ya existe la factura: " & vbCrLf & SQL
            SQL = "Comprobando Nº Facturas en Contabilidad...       " & vbCrLf & vbCrLf & SQL
            
            'MsgBox sql, vbExclamation
            ComprobarNumFacturasContaNueva = False
        Else
            ComprobarNumFacturasContaNueva = True
        End If
    Else
        ComprobarNumFacturasContaNueva = True
    End If
    RSconta.Close
    Set RSconta = Nothing
    
ECompFactu:
     If Err.Number <> 0 Then
        MuestraError Err.Number, "Comprobar Nº Facturas", Err.Description
    End If
End Function





Public Function ComprobarNumFacturasFac(cadWConta) As Boolean
'Comprobar que no exista ya en la contabilidad un nº de factura para la fecha que
'vamos a contabilizar
Dim SQL As String
Dim Rs As ADODB.Recordset
Dim RSconta As ADODB.Recordset
Dim b As Boolean

    On Error GoTo ECompFactuFac

    ComprobarNumFacturasFac = False
    
    SQL = "SELECT numserie,codfaccl,anofaccl FROM cabfact "
    SQL = SQL & " WHERE " & cadWConta
    
    Set RSconta = New ADODB.Recordset
    RSconta.Open SQL, ConnContaFac, adOpenForwardOnly, adLockPessimistic, adCmdText

    If Not RSconta.EOF Then
        'Seleccionamos las distintas facturas que vamos a facturar
        SQL = "SELECT DISTINCT tmpfactu.numserie,tmpfactu.numfactu,tmpfactu.fecfactu "
        SQL = SQL & " FROM tmpfactu "
        
        Set Rs = New ADODB.Recordset
        Rs.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        b = True
        While Not Rs.EOF 'And b
' quitado el 12022007
'            SQL = "(numserie= " & DBSet(RS!letraser, "T") & " AND codfaccl=" & DBSet(RS!numfactu, "N") & " AND anofaccl=" & Year(RS!fecfactu) & ")"
'            If SituarRSetMULTI(RSconta, SQL) Then
            SQL = ""
            SQL = DevuelveDesdeBDNewFac("cabfact", "codfaccl", "codfaccl", Rs!numfactu, "N", , "numserie", Trim(Rs!numserie), "T", "anofaccl", Year(Rs!fecfactu), "N")
            If SQL <> "" Then
                b = False
                SQL = "          Nº Fac.: " & Format(Rs!numfactu, "0000000") & vbCrLf
                SQL = SQL & "          Fecha: " & Rs!fecfactu
                
                SQL = "Ya existe la factura: " & vbCrLf & SQL
                InsertarError SQL
            
            End If
            Rs.MoveNext
        Wend
        Rs.Close
        Set Rs = Nothing
        
        If Not b Then
            SQL = "Ya existe la factura: " & vbCrLf & SQL
            SQL = "Comprobando Nº Facturas en Contabilidad...       " & vbCrLf & vbCrLf & SQL
            
            'MsgBox sql, vbExclamation
            ComprobarNumFacturasFac = False
        Else
            ComprobarNumFacturasFac = True
        End If
    Else
        ComprobarNumFacturasFac = True
    End If
    RSconta.Close
    Set RSconta = Nothing
    
ECompFactuFac:
     If Err.Number <> 0 Then
        MuestraError Err.Number, "Comprobar Nº Facturas Varias", Err.Description
    End If
End Function



Public Function ComprobarNumFacturasFacContaNueva(cadWConta) As Boolean
'Comprobar que no exista ya en la contabilidad un nº de factura para la fecha que
'vamos a contabilizar
Dim SQL As String
Dim Rs As ADODB.Recordset
Dim RSconta As ADODB.Recordset
Dim b As Boolean

    On Error GoTo ECompFactuFac

    ComprobarNumFacturasFacContaNueva = False
    
    SQL = "SELECT numserie,numfactu,anofactu FROM factcli "
    SQL = SQL & " WHERE " & cadWConta
    
    Set RSconta = New ADODB.Recordset
    RSconta.Open SQL, ConnContaFac, adOpenForwardOnly, adLockPessimistic, adCmdText

    If Not RSconta.EOF Then
        'Seleccionamos las distintas facturas que vamos a facturar
        SQL = "SELECT DISTINCT tmpfactu.numserie,tmpfactu.numfactu,tmpfactu.fecfactu "
        SQL = SQL & " FROM tmpfactu "
        
        Set Rs = New ADODB.Recordset
        Rs.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        b = True
        While Not Rs.EOF 'And b
' quitado el 12022007
'            SQL = "(numserie= " & DBSet(RS!letraser, "T") & " AND codfaccl=" & DBSet(RS!numfactu, "N") & " AND anofaccl=" & Year(RS!fecfactu) & ")"
'            If SituarRSetMULTI(RSconta, SQL) Then
            SQL = ""
            SQL = DevuelveDesdeBDNewFac("factcli", "numfactu", "numfactu", Rs!numfactu, "N", , "numserie", Trim(Rs!numserie), "T", "anofactu", Year(Rs!fecfactu), "N")
            If SQL <> "" Then
                b = False
                SQL = "          Nº Fac.: " & Format(Rs!numfactu, "0000000") & vbCrLf
                SQL = SQL & "          Fecha: " & Rs!fecfactu
                
                SQL = "Ya existe la factura: " & vbCrLf & SQL
                InsertarError SQL
            
            End If
            Rs.MoveNext
        Wend
        Rs.Close
        Set Rs = Nothing
        
        If Not b Then
            SQL = "Ya existe la factura: " & vbCrLf & SQL
            SQL = "Comprobando Nº Facturas en Contabilidad...       " & vbCrLf & vbCrLf & SQL
            
            'MsgBox sql, vbExclamation
            ComprobarNumFacturasFacContaNueva = False
        Else
            ComprobarNumFacturasFacContaNueva = True
        End If
    Else
        ComprobarNumFacturasFacContaNueva = True
    End If
    RSconta.Close
    Set RSconta = Nothing
    
ECompFactuFac:
     If Err.Number <> 0 Then
        MuestraError Err.Number, "Comprobar Nº Facturas Varias", Err.Description
    End If
End Function






Public Function ComprobarCtaContable(cadTABLA As String, Opcion As Byte, Optional cadwhere As String, Optional bd As Byte, Optional tipo As Byte) As Boolean
'Comprobar que todas las ctas contables de los distintos clientes de las facturas
'que vamos a contabilizar existan en la contabilidad
Dim SQL As String
Dim Rs As ADODB.Recordset
Dim RSconta As ADODB.Recordset
Dim b As Boolean
Dim cadG As String
Dim enc As String
    
    On Error GoTo ECompCta

    ComprobarCtaContable = False
    
    SQL = "SELECT codmacta FROM cuentas "
    SQL = SQL & " WHERE apudirec='S'"
    If cadG <> "" Then SQL = SQL & cadG
    
    Set RSconta = New ADODB.Recordset
    Select Case bd
        Case cConta
            RSconta.Open SQL, ConnConta, adOpenStatic, adLockPessimistic, adCmdText
        Case cContaSeg
            RSconta.Open SQL, ConnContaSeg, adOpenStatic, adLockPessimistic, adCmdText
        Case cContaTel
            RSconta.Open SQL, ConnContaTel, adOpenStatic, adLockPessimistic, adCmdText
        Case cContaFacSoc
            RSconta.Open SQL, ConnContaFacSoc, adOpenStatic, adLockPessimistic, adCmdText
        Case cContaCV
            RSconta.Open SQL, ConnContaCV, adOpenStatic, adLockPessimistic, adCmdText
        Case cContaCVV
            RSconta.Open SQL, ConnContaCVV, adOpenStatic, adLockPessimistic, adCmdText
    End Select

    If Not RSconta.EOF Then
        If Opcion = 1 Then
                SQL = "SELECT DISTINCT avnic.codmacta, avnic.codavnic  "
                SQL = SQL & " FROM avnic, movim  "
                SQL = SQL & " where " & cadwhere & " and avnic.codavnic = movim.codavnic and avnic.anoejerc = movim.anoejerc "
        ElseIf Opcion = 2 Then
                SQL = "SELECT distinct segpoliza.codmacta, segpoliza.codrefer  "
                SQL = SQL & " from segpoliza "
                SQL = SQL & " where " & cadwhere
        ElseIf Opcion = 3 Then
                'si hay analitica comprobar que todas las cuentas
                'empiezan por el digito que hay en conta.parametros.grupovta
                cadG = DevuelveDesdeBDNew(cConta, "parametros", "grupovta", "", "", "")
        
                SQL = "SELECT distinct sartic.codartic "
                SQL = SQL & ", sartic.codmacta, sartic.codmaccl"
                SQL = SQL & " from ((slhfac "
                SQL = SQL & " INNER JOIN tmpfactu ON slhfac.letraser=tmpfactu.letraser AND slhfac.numfactu=tmpfactu.numfactu AND slhfac.fecfactu=tmpfactu.fecfactu) "
                SQL = SQL & "INNER JOIN sartic ON slhfac.codartic=sartic.codartic) "
                SQL = SQL & " where sartic.codmacta "
                If cadG <> "" Then
                     SQL = SQL & " AND not ((sartic.codmacta like '" & cadG & "%') and (sartic.codmaccl like '" & cadG & "%'))"
                End If
        ElseIf Opcion = 4 Then
            SQL = "select codmacta from telmovil "
        ElseIf Opcion = 5 Then
            SQL = "select ctabancoseg from sparam "
        ElseIf Opcion = 6 Then
            SQL = "select ctagasto from sparam"
        ElseIf Opcion = 7 Then
            SQL = "select ctareten from sparam"
        ElseIf Opcion = 8 Then
            SQL = "SELECT distinct factsocio.codmacta "
            SQL = SQL & " from (factsocio "
            SQL = SQL & " INNER JOIN tmpfactu ON factsocio.numfactu=tmpfactu.numfactu AND factsocio.fecfactu=tmpfactu.fecfactu AND factsocio.codmacta=tmpfactu.codmacta) "
        ElseIf Opcion = 9 Then
            SQL = " select variedad.codmacta "
            SQL = SQL & " from ((factsocio "
            SQL = SQL & " INNER JOIN tmpfactu ON factsocio.numfactu=tmpfactu.numfactu AND factsocio.fecfactu=tmpfactu.fecfactu AND factsocio.codmacta=tmpfactu.codmacta) "
            SQL = SQL & " INNER JOIN variedad ON factsocio.codvarie=variedad.codvarie) "
        ElseIf Opcion = 10 Then
            SQL = "select ctaretenfacsoc from sparam"
        ElseIf Opcion = 11 Then
            SQL = "SELECT distinct variedad.codmacta "
            SQL = SQL & " from (factsocio "
            SQL = SQL & " INNER JOIN tmpfactu ON factsocio.numfactu=tmpfactu.numfactu AND factsocio.fecfactu=tmpfactu.fecfactu AND factsocio.codmacta=tmpfactu.codmacta) "
            SQL = SQL & " INNER JOIN variedad on factsocio.codvarie = variedad.codvarie "
            SQL = SQL & " where not variedad.codmacta like '" & vEmpresaFacSoc.DigGrupoGto & "%'"
        ElseIf Opcion = 12 Then
            SQL = "SELECT distinct cvfacturas.codmactasoc codmacta "
            SQL = SQL & " from (cvfacturas "
            SQL = SQL & " INNER JOIN tmpfactu ON cvfacturas.numfactu=tmpfactu.numfactu AND cvfacturas.fecfactu=tmpfactu.fecfactu and cvfacturas.letraser=tmpfactu.numserie) "
        ElseIf Opcion = 13 Then
            SQL = "SELECT distinct cvfacturas.codmactavta codmacta "
            SQL = SQL & " from (cvfacturas "
            SQL = SQL & " INNER JOIN tmpfactu ON cvfacturas.numfactu=tmpfactu.numfactu AND cvfacturas.fecfactu=tmpfactu.fecfactu and cvfacturas.letraser=tmpfactu.numserie) "
            
        End If
        
        Set Rs = New ADODB.Recordset
        Rs.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        b = True
        While Not Rs.EOF 'And b
            If Opcion = 3 Then
                SQL = Rs!Codmacta & " o " & Rs!CodmacCl
                SQL = "La cuenta " & SQL & " del articulo " & Rs!CodArtic & " no es del grupo correcto."
                InsertarError SQL
            Else
                If Opcion = 11 Then
                    SQL = Rs!Codmacta
                    SQL = "La cuenta " & SQL & " de la variedad no es del grupo correcto."
                    InsertarError SQL
                Else
                    SQL = "codmacta= " & DBLet(Rs.Fields(0).Value, "T") 'DBSet(RS.Fields(0).Value, "T") '& " and apudirec='S' "
                End If
            End If
            enc = ""
            Select Case bd
                Case cConta
                    enc = DevuelveDesdeBDNew(cConta, "cuentas", "codmacta", "codmacta", Rs.Fields(0).Value, "T")
                Case cContaSeg
                    enc = DevuelveDesdeBDNew(cContaSeg, "cuentas", "codmacta", "codmacta", Rs.Fields(0).Value, "T")
                Case cContaTel
                    enc = DevuelveDesdeBDNew(cContaTel, "cuentas", "codmacta", "codmacta", Rs.Fields(0).Value, "T")
                Case cContaFacSoc
                    enc = DevuelveDesdeBDNew(cContaFacSoc, "cuentas", "codmacta", "codmacta", Rs.Fields(0).Value, "T")
                Case cContaCV
                    enc = DevuelveDesdeBDNew(cContaCV, "cuentas", "codmacta", "codmacta", Rs.Fields(0).Value, "T")
                Case cContaCVV
                    enc = DevuelveDesdeBDNew(cContaCVV, "cuentas", "codmacta", "codmacta", Rs.Fields(0).Value, "T")
            End Select
                 
            If enc = "" Then
                b = False 'no encontrado
                If Opcion = 1 Then
                        SQL = Rs!Codmacta & " del Código Avnic " & Format(Rs!codavnic, "0000000")
                        SQL = "No existe la cta contable " & SQL
                        InsertarError SQL
                End If
                If Opcion = 2 Then
                    SQL = Rs!Codmacta & " de la Póliza " & Rs!Codrefer
                    SQL = "No existe la cta contable " & SQL
                    InsertarError SQL
                End If
                If Opcion = 4 Then
                    SQL = "No existe la cta contable " & SQL
                    InsertarError SQL
                End If
                If Opcion = 5 Or Opcion = 6 Or Opcion = 7 Or Opcion = 8 Or Opcion = 9 Then
                    SQL = "No existe la cta contable " & SQL
                    InsertarError SQL
                End If
                If Opcion = 10 Then
                    SQL = "No existe la cta contable de retencion " & SQL
                    InsertarError SQL
                End If
                If Opcion = 12 Or Opcion = 13 Then
                    SQL = "No existe la cta contable  " & DBLet(Rs.Fields(0).Value)
                    InsertarError SQL
                End If
            Else
                If Opcion = 11 Then
                    b = False
                End If
            End If
                
            Rs.MoveNext
        Wend
        Rs.Close
        Set Rs = Nothing
        
        If Not b Then
            ComprobarCtaContable = False
        Else
            ComprobarCtaContable = True
        End If
    Else
        ComprobarCtaContable = True
    End If
    RSconta.Close
    Set RSconta = Nothing
    
ECompCta:
     If Err.Number <> 0 Then
        MuestraError Err.Number, "Comprobar Ctas Contables", Err.Description
    End If
End Function


Public Function ComprobarCtaContableFac(Opcion As Byte, Optional cadwhere As String) As Boolean
'Comprobar que todas las ctas contables de los distintos clientes de las facturas
'que vamos a contabilizar existan en la contabilidad
Dim SQL As String
Dim Rs As ADODB.Recordset
Dim RSconta As ADODB.Recordset
Dim b As Boolean
Dim cadG As String
Dim enc As String
    
    On Error GoTo ECompCta

    ComprobarCtaContableFac = False
    
    SQL = "SELECT codmacta FROM cuentas "
    SQL = SQL & " WHERE apudirec='S'"
    If cadG <> "" Then SQL = SQL & cadG
    
    Set RSconta = New ADODB.Recordset
    RSconta.Open SQL, ConnContaFac, adOpenStatic, adLockPessimistic, adCmdText

    If Not RSconta.EOF Then
        If Opcion = 1 Then
                SQL = "SELECT DISTINCT cabfact.ctaclien, cabfact.numfactu   "
                SQL = SQL & " FROM cabfact  "
                SQL = SQL & " where " & cadwhere
        ElseIf Opcion = 2 Then
                SQL = "SELECT distinct concefact.codmacta, concefact.codconce "
                SQL = SQL & " from concefact, linfact, cabfact "
                SQL = SQL & " where " & cadwhere & " and concefact.codconce = linfact.codconce"
                SQL = SQL & " and cabfact.codsecci = linfact.codsecci "
                SQL = SQL & " and cabfact.letraser = linfact.letraser "
                SQL = SQL & " and cabfact.numfactu = linfact.numfactu "
                SQL = SQL & " and cabfact.fecfactu = linfact.fecfactu "
        ElseIf Opcion = 3 Then
                'si hay analitica comprobar que todas las cuentas
                'empiezan por el digito que hay en conta.parametros.grupovta
                cadG = DevuelveDesdeBDNewFac("parametros", "grupovta", "", "", "")
        
                SQL = "SELECT distinct concefact.codconce "
                SQL = SQL & ", concefact.codmacta"
                SQL = SQL & " from ((linfact "
                SQL = SQL & " INNER JOIN tmpfactu ON linfact.codsecci=tmpfactu.codsecci and linfact.letraser=tmpfactu.numserie AND linfact.numfactu=tmpfactu.numfactu AND linfact.fecfactu=tmpfactu.fecfactu) "
                SQL = SQL & " INNER JOIN concefact on linfact.codconce = concefact.codconce) "
                SQL = SQL & " where concefact.codmacta "
                If cadG <> "" Then
                     SQL = SQL & " AND not (concefact.codmacta like '" & cadG & "%') "
                End If
        ElseIf Opcion = 4 Then
            b = True
            enc = ""
            enc = DevuelveDesdeBDNewFac("cuentas", "codmacta", "codmacta", cadwhere, "T")
            If enc = "" Then
                b = False
                SQL = "No existe la cta contable de banco" & cadwhere
                InsertarError SQL
            End If
        ElseIf Opcion = 5 Then
            SQL = "select ctabancoseg from sparam "
        ElseIf Opcion = 6 Then
            SQL = "select ctagasto from sparam"
        ElseIf Opcion = 7 Then
            SQL = "select ctareten from sparam"
        ElseIf Opcion = 8 Then
            SQL = "SELECT DISTINCT cabfact.cuereten   "
            SQL = SQL & " FROM cabfact  "
            SQL = SQL & " where " & cadwhere
        End If
        If Opcion <> 4 Then
            Set Rs = New ADODB.Recordset
            Rs.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            b = True
            While Not Rs.EOF 'And b
                If Opcion = 3 Then
                    SQL = Rs!Codmacta
                    SQL = "La cuenta " & SQL & " del concepto " & Rs!codConce & " no es del grupo correcto."
                    InsertarError SQL
                Else
                    SQL = "codmacta= " & DBLet(Rs.Fields(0).Value, "T") '& " and apudirec='S' "
                End If
                
                enc = ""
                If Opcion <> 8 Then
                    enc = DevuelveDesdeBDNewFac("cuentas", "codmacta", "codmacta", Rs.Fields(0).Value, "T")
                Else
                    If DBLet(Rs.Fields(0).Value, "T") <> "" Then
                        enc = DevuelveDesdeBDNewFac("cuentas", "codmacta", "codmacta", Rs.Fields(0).Value, "T")
                    End If
                End If
                     
                If enc = "" Then
                    If Opcion <> 8 Then b = False 'no encontrado
                    If Opcion = 1 Then
                            SQL = Rs!Codmacta & " de la Factura " & Format(Rs!numfactu, "0000000")
                            SQL = "No existe la cta contable " & SQL
                            InsertarError SQL
                    End If
                    If Opcion = 2 Then
                        SQL = Rs!Codmacta & " del Concepto " & Rs!codConce
                        SQL = "No existe la cta contable " & SQL
                        InsertarError SQL
                    End If
                    If Opcion = 4 Then
                        SQL = "No existe la cta contable " & SQL
                        InsertarError SQL
                    End If
                    If Opcion = 5 Or Opcion = 6 Or Opcion = 7 Then
                        SQL = "No existe la cta contable " & SQL
                        InsertarError SQL
                    End If
                    If Opcion = 8 Then
                        SQL = DBLet(Rs!cuereten, "T")
                        If SQL <> "" Then
                            b = False
                            SQL = "No existe la cta contable de retención: " & SQL
                            InsertarError SQL
                        End If
                    End If
                End If
                    
                Rs.MoveNext
            Wend
            Rs.Close
            Set Rs = Nothing
        End If
        If Not b Then
            ComprobarCtaContableFac = False
        Else
            ComprobarCtaContableFac = True
        End If
    Else
        ComprobarCtaContableFac = True
    End If
    RSconta.Close
    Set RSconta = Nothing
    
ECompCta:
     If Err.Number <> 0 Then
        MuestraError Err.Number, "Comprobar Ctas Contables", Err.Description
    End If
End Function


Public Function ComprobarCtaContableGas(Opcion As Byte, Optional cadwhere As String) As Boolean
'Comprobar que todas las ctas contables de los distintos socios de las facturas
'que vamos a contabilizar existan en la contabilidad
Dim SQL As String
Dim Rs As ADODB.Recordset
Dim RSconta As ADODB.Recordset
Dim b As Boolean
Dim cadG As String
Dim enc As String
Dim longitud As Integer
    
    On Error GoTo ECompCta

    ComprobarCtaContableGas = False
    
    SQL = "SELECT codmacta FROM cuentas "
    SQL = SQL & " WHERE apudirec='S'"
    If cadG <> "" Then SQL = SQL & cadG
    
    Set RSconta = New ADODB.Recordset
    RSconta.Open SQL, ConnContaGas, adOpenStatic, adLockPessimistic, adCmdText

    If Not RSconta.EOF Then
        If Opcion = 1 Then
                longitud = vEmpresaGas.DigitosUltimoNivel - vEmpresaGas.DigitosNivelAnterior
                SQL = "SELECT DISTINCT concat(" & vParamAplic.RaizCtaSocGas & ", right(concat('0000000000', codsocio)," & longitud & ")) as codmacta, numfactu  "
                SQL = SQL & " FROM gascabfac  "
                SQL = SQL & " where " & cadwhere
        ElseIf Opcion = 2 Then
                SQL = "SELECT " & vParamAplic.CtaVentasGas & " as codmacta "
                SQL = SQL & " from sparam "
        ElseIf Opcion = 3 Then
                'si hay analitica comprobar que la cuenta de ventas
                'empieza por el digito que hay en conta.parametros.grupovta
                cadG = DevuelveDesdeBDNew(cContaGas, "parametros", "grupovta", "", "", "")
        
                SQL = "SELECT  " & vParamAplic.CtaVentasGas & " as codmacta "
                SQL = SQL & " from sparam "
                If cadG <> "" Then
                     SQL = SQL & " where not (" & vParamAplic.CtaVentasGas & " like '" & cadG & "%') "
                End If
        ElseIf Opcion = 4 Then
                SQL = "select " & vParamAplic.CtaContraGas & " as codmacta "
                SQL = SQL & " from sparam "
        ElseIf Opcion = 5 Then
            SQL = "select ctabancoseg from sparam "
        ElseIf Opcion = 6 Then
            SQL = "select ctagasto from sparam"
        ElseIf Opcion = 7 Then
            SQL = "select ctareten from sparam"
        End If
        Set Rs = New ADODB.Recordset
        Rs.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        b = True
        While Not Rs.EOF 'And b
            If Opcion = 3 Then
                SQL = vParamAplic.CtaVentasGas
                SQL = "La cuenta " & SQL & " no es del grupo correcto."
                b = False
                InsertarError SQL
            Else
                SQL = "codmacta= " & DBSet(Rs.Fields(0).Value, "T") '& " and apudirec='S' "
            End If
        
            enc = ""
            enc = DevuelveDesdeBDNew(cContaGas, "cuentas", "codmacta", "codmacta", Rs.Fields(0).Value, "T")
                 
            If enc = "" Then
                b = False 'no encontrado
                If Opcion = 1 Then
                    SQL = Rs!Codmacta & " de la Factura " & Format(Rs!numfactu, "0000000")
                    SQL = "No existe la cta contable " & SQL
                    InsertarError SQL
                End If
                If Opcion = 2 Then
                    SQL = Rs!Codmacta & " de ventas de gasolinera "
                    SQL = "No existe la cta contable " & SQL
                    InsertarError SQL
                End If
                If Opcion = 3 Then
                    SQL = Rs!Codmacta
                    SQL = "La cuenta de ventas " & SQL & " no es del grupo correcto."
                    InsertarError SQL
                End If
                If Opcion = 4 Then
                    SQL = "No existe la cta contable " & SQL
                    InsertarError SQL
                End If
                If Opcion = 5 Or Opcion = 6 Or Opcion = 7 Then
                    SQL = "No existe la cta contable " & SQL
                    InsertarError SQL
                End If
            End If
                
            Rs.MoveNext
        Wend
        Rs.Close
        Set Rs = Nothing
        If Not b Then
            ComprobarCtaContableGas = False
        Else
            ComprobarCtaContableGas = True
        End If
    Else
        ComprobarCtaContableGas = True
    End If
    RSconta.Close
    Set RSconta = Nothing
    
ECompCta:
     If Err.Number <> 0 Then
        MuestraError Err.Number, "Comprobar Ctas Contables", Err.Description
    End If
End Function





Public Function ComprobarTiposIVA(Seccion As String) As Boolean
'Comprobar que todos los Tipos de IVA de las distintas facturas (scafac.codigiva1, codigiv2,codigiv3)
'que vamos a contabilizar existan en la contabilidad
Dim SQL As String
Dim Rs As ADODB.Recordset
Dim RSconta As ADODB.Recordset
Dim b As Boolean
Dim i As Byte
'Dim CodigIVA As String

    On Error GoTo ECompIVA

    ComprobarTiposIVA = False
    
    SQL = "SELECT distinct codigiva FROM tiposiva "
    
    Set RSconta = New ADODB.Recordset
    RSconta.Open SQL, ConnContaFac, adOpenStatic, adLockPessimistic, adCmdText

    If Not RSconta.EOF Then
        'Seleccionamos los distintos tipos de IVA de las facturas a Contabilizar
        For i = 1 To 3
            SQL = "SELECT DISTINCT cabfact.tipoiva" & i
            SQL = SQL & " FROM cabfact "
            SQL = SQL & " INNER JOIN tmpfactu ON cabfact.letraser=tmpfactu.numserie AND cabfact.numfactu=tmpfactu.numfactu AND cabfact.fecfactu=tmpfactu.fecfactu "
            SQL = SQL & " WHERE not isnull(tipoiva" & i & ")"
            '[Monica]19/05/2016: pasamos la seccion
            If Seccion <> "" Then SQL = SQL & " and cabfact.codsecci = " & DBSet(Seccion, "N")
            
            Set Rs = New ADODB.Recordset
            Rs.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            b = True
            While Not Rs.EOF 'And b
                If Rs.Fields(0) <> 0 Then ' añadido pq en arigasol sino tiene tipo de iva pone ceros
                    SQL = "codigiva= " & DBSet(Rs.Fields(0), "N")
                    RSconta.MoveFirst
                    RSconta.Find (SQL), , adSearchForward
                    If RSconta.EOF Then
                        b = False 'no encontrado
                        SQL = "No existe el " & SQL
                        SQL = "Tipo de IVA: " & Rs.Fields(0)
                        InsertarError SQL
                    End If
                End If
                Rs.MoveNext
            Wend
            Rs.Close
            Set Rs = Nothing
        
            If Not b Then
                ComprobarTiposIVA = False
                Exit For
            Else
                ComprobarTiposIVA = True
            End If
        Next i
    End If
    RSconta.Close
    Set RSconta = Nothing
    
ECompIVA:
    If Err.Number <> 0 Then
        MuestraError Err.Number, "Comprobar Tipo de IVA.", Err.Description
    End If
End Function


Public Function ComprobarTiposIVAGas() As Boolean
'Comprobar que todos los Tipos de IVA de las distintas facturas (scafac.codigiva1, codigiv2,codigiv3)
'que vamos a contabilizar existan en la contabilidad
Dim SQL As String
Dim Rs As ADODB.Recordset
Dim RSconta As ADODB.Recordset
Dim b As Boolean
Dim i As Byte
'Dim CodigIVA As String

    On Error GoTo ECompIVA

    ComprobarTiposIVAGas = False
    
    SQL = "SELECT distinct codigiva FROM tiposiva "
    
    Set RSconta = New ADODB.Recordset
    RSconta.Open SQL, ConnContaGas, adOpenStatic, adLockPessimistic, adCmdText

    If Not RSconta.EOF Then
        'Seleccionamos los distintos tipos de IVA de las facturas a Contabilizar
        SQL = "SELECT DISTINCT gascabfac.codiva "
        SQL = SQL & " FROM gascabfac "
        SQL = SQL & " INNER JOIN tmpfactu ON gascabfac.letraser=tmpfactu.numserie AND gascabfac.numfactu=tmpfactu.numfactu AND gascabfac.fecfactu=tmpfactu.fecfactu "
        SQL = SQL & " WHERE not isnull(codiva)"

        Set Rs = New ADODB.Recordset
        Rs.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        b = True
        While Not Rs.EOF 'And b
            If Rs.Fields(0) <> 0 Then ' añadido pq en arigasol sino tiene tipo de iva pone ceros
                SQL = "codigiva= " & DBSet(Rs.Fields(0), "N")
                RSconta.MoveFirst
                RSconta.Find (SQL), , adSearchForward
                If RSconta.EOF Then
                    b = False 'no encontrado
                    SQL = "No existe el " & SQL
                    SQL = "Tipo de IVA: " & Rs.Fields(0)
                    InsertarError SQL
                End If
            End If
            Rs.MoveNext
        Wend
        Rs.Close
        Set Rs = Nothing
    
        If Not b Then
            ComprobarTiposIVAGas = False
        Else
            ComprobarTiposIVAGas = True
        End If
    End If
    RSconta.Close
    Set RSconta = Nothing
    
ECompIVA:
    If Err.Number <> 0 Then
        MuestraError Err.Number, "Comprobar Tipo de IVA.", Err.Description
    End If
End Function


Public Function ComprobarTiposIVAFacSoc() As Boolean
'Comprobar que todos los Tipos de IVA de las distintas facturas (factsocio.tipoiva)
'que vamos a contabilizar existan en la contabilidad
Dim SQL As String
Dim Rs As ADODB.Recordset
Dim RSconta As ADODB.Recordset
Dim b As Boolean
Dim i As Byte
'Dim CodigIVA As String

    On Error GoTo ECompIVA

    ComprobarTiposIVAFacSoc = False
    
    SQL = "SELECT distinct codigiva FROM tiposiva "
    
    Set RSconta = New ADODB.Recordset
    RSconta.Open SQL, ConnContaFacSoc, adOpenStatic, adLockPessimistic, adCmdText

    If Not RSconta.EOF Then
        'Seleccionamos los distintos tipos de IVA de las facturas a Contabilizar
            SQL = "SELECT DISTINCT factsocio.tipoiva"
            SQL = SQL & " FROM factsocio "
            SQL = SQL & " INNER JOIN tmpfactu ON factsocio.numfactu=tmpfactu.numfactu AND factsocio.fecfactu=tmpfactu.fecfactu AND factsocio.codmacta=tmpfactu.codmacta "
            SQL = SQL & " WHERE not isnull(tipoiva)"

            Set Rs = New ADODB.Recordset
            Rs.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            b = True
            While Not Rs.EOF 'And b
                If Rs.Fields(0) <> 0 Then ' añadido pq en arigasol sino tiene tipo de iva pone ceros
                    SQL = "codigiva= " & DBSet(Rs.Fields(0), "N")
                    RSconta.MoveFirst
                    RSconta.Find (SQL), , adSearchForward
                    If RSconta.EOF Then
                        b = False 'no encontrado
                        SQL = "No existe el " & SQL
                        SQL = "Tipo de IVA: " & Rs.Fields(0)
                        InsertarError SQL
                    End If
                End If
                Rs.MoveNext
            Wend
            Rs.Close
            Set Rs = Nothing
        
            If Not b Then
                ComprobarTiposIVAFacSoc = False
            Else
                ComprobarTiposIVAFacSoc = True
            End If
    End If
    RSconta.Close
    Set RSconta = Nothing
    
ECompIVA:
    If Err.Number <> 0 Then
        MuestraError Err.Number, "Comprobar Tipo de IVA.", Err.Description
    End If
End Function


Public Function ComprobarTiposIVACV(tipo As Byte) As Boolean
'Comprobar que todos los Tipos de IVA de las distintas facturas (scafac.codigiva1, codigiv2,codigiv3)
'que vamos a contabilizar existan en la contabilidad
Dim SQL As String
Dim Rs As ADODB.Recordset
Dim RSconta As ADODB.Recordset
Dim b As Boolean
Dim i As Byte
'Dim CodigIVA As String

    On Error GoTo ECompIVA

    ComprobarTiposIVACV = False
    
    SQL = "SELECT distinct codigiva FROM tiposiva "
    
    Set RSconta = New ADODB.Recordset
    If tipo = 0 Then
        RSconta.Open SQL, ConnContaCVV, adOpenStatic, adLockPessimistic, adCmdText
    Else
        RSconta.Open SQL, ConnContaCV, adOpenStatic, adLockPessimistic, adCmdText
    End If

    If Not RSconta.EOF Then
        'Seleccionamos los distintos tipos de IVA de las facturas a Contabilizar
        SQL = "SELECT DISTINCT cvfacturas.codiva, cvfacturas.codiva2, cvfacturas.codiva3 "
        SQL = SQL & " FROM cvfacturas "
        SQL = SQL & " INNER JOIN tmpfactu ON cvfacturas.letraser=tmpfactu.numserie AND cvfacturas.numfactu=tmpfactu.numfactu AND cvfacturas.fecfactu=tmpfactu.fecfactu "
        SQL = SQL & " WHERE not isnull(codiva)"

        Set Rs = New ADODB.Recordset
        Rs.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        b = True
        While Not Rs.EOF 'And b
            If Rs.Fields(0) <> 0 Then ' añadido pq en arigasol sino tiene tipo de iva pone ceros
                SQL = "codigiva= " & DBSet(Rs.Fields(0), "N")
                RSconta.MoveFirst
                RSconta.Find (SQL), , adSearchForward
                If RSconta.EOF Then
                    b = False 'no encontrado
                    SQL = "No existe el " & SQL
                    SQL = "Tipo de IVA: " & Rs.Fields(0)
                    InsertarError SQL
                End If
            End If
            If Rs.Fields(1) <> 0 Then ' añadido pq en arigasol sino tiene tipo de iva pone ceros
                SQL = "codigiva= " & DBSet(Rs.Fields(1), "N")
                RSconta.MoveFirst
                RSconta.Find (SQL), , adSearchForward
                If RSconta.EOF Then
                    b = False 'no encontrado
                    SQL = "No existe el " & SQL
                    SQL = "Tipo de IVA: " & Rs.Fields(1)
                    InsertarError SQL
                End If
            End If
            If Rs.Fields(2) <> 0 Then ' añadido pq en arigasol sino tiene tipo de iva pone ceros
                SQL = "codigiva= " & DBSet(Rs.Fields(2), "N")
                RSconta.MoveFirst
                RSconta.Find (SQL), , adSearchForward
                If RSconta.EOF Then
                    b = False 'no encontrado
                    SQL = "No existe el " & SQL
                    SQL = "Tipo de IVA: " & Rs.Fields(2)
                    InsertarError SQL
                End If
            End If
            
            
            
            Rs.MoveNext
        Wend
        Rs.Close
        Set Rs = Nothing
    
        If Not b Then
            ComprobarTiposIVACV = False
        Else
            ComprobarTiposIVACV = True
        End If
    End If
    RSconta.Close
    Set RSconta = Nothing
    
ECompIVA:
    If Err.Number <> 0 Then
        MuestraError Err.Number, "Comprobar Tipo de IVA.", Err.Description
    End If
End Function



Public Function PasarFactura(cadwhere As String, FecVenci As String, ctaVta As String, CtaBanco As String, FPago As String, Tipiva As String, CodCCost As String) As Boolean
'Insertar en tablas cabecera/lineas de la Contabilidad la factura
' arigasol.schfac --> conta.cabfact
' arigasol.slhfac --> conta.linfact
'Actualizar la tabla ariges.scafac.inconta=1 para indicar que ya esta contabilizada
Dim b As Boolean
Dim cadMen As String
Dim SQL As String
Dim vsocio As CSocio
Dim codsoc As Long
Dim Rs As ADODB.Recordset

Dim Rsx As ADODB.Recordset
Dim sql2 As String
Dim codfor As Integer
Dim TipForpa As String
Dim PorIva As String

    On Error GoTo EContab

    ConnContaTel.BeginTrans
    conn.BeginTrans
     
    PorIva = ""
    PorIva = DevuelveDesdeBDNew(cContaTel, "tiposiva", "porceiva", "codigiva", Tipiva, "N")
    'Insertar en la conta Cabecera Factura
    b = InsertarCabFact(cadwhere, Tipiva, cadMen, PorIva, FPago)
    cadMen = "Insertando Cab. Factura: " & cadMen
    
    ' insertar en tesoreria
    If b Then
        SQL = "select * from telmovil where " & cadwhere
        Set Rsx = New ADODB.Recordset
        Rsx.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        
        cadMen = ""
        b = InsertarEnTesoreriaNew3(Rsx, FecVenci, FPago, CtaBanco, cadMen)
        cadMen = "Insertando en Tesoreria: " & cadMen
    End If

    If b Then
        'Insertar lineas de Factura en la Conta
        cadMen = ""
        b = InsertarLinFact("telmovil", cadwhere, cadMen, ctaVta, CodCCost, Tipiva, PorIva)
        cadMen = "Insertando Lin. Factura: " & cadMen

        If b Then
            'Poner intconta=1 en arigasol.scafac
            cadMen = ""
            b = ActualizarCabFact("telmovil", cadwhere, cadMen)
            cadMen = "Actualizando Factura: " & cadMen
        End If
    End If
    
'    If Not b Then
'        SQL = "Insert into tmperrfac(codtipom,numfactu,fecfactu,error) "
'        SQL = SQL & " Select *," & DBSet(cadMen, "T") & " as error From tmpfactu "
'        SQL = SQL & " WHERE " & Replace(cadWhere, "telmovil", "tmpfactu")
'        conn.Execute SQL
'    End If
    
EContab:
    If Err.Number <> 0 Then
        b = False
        MuestraError Err.Number, "Contabilizando Factura", Err.Description
    End If
    If b Then
        ConnContaTel.CommitTrans
        conn.CommitTrans
        PasarFactura = True
    Else
       
        ConnContaTel.RollbackTrans
        conn.RollbackTrans
        
        SQL = "Insert into tmperrfac(codtipom,numfactu,fecfactu,error) "
        SQL = SQL & " Select *," & DBSet(cadMen, "T") & " as error From tmpfactu "
        SQL = SQL & " WHERE " & Replace(cadwhere, "telmovil", "tmpfactu")
        conn.Execute SQL
        
        
        PasarFactura = False
    End If
End Function


Public Function PasarFacturaCV(cadwhere As String, FecVenci As String, tipo As String, CtaBanco As String, CodCCost As String) As Boolean
'Insertar en tablas cabecera/lineas de la Contabilidad la factura
' arigasol.schfac --> conta.cabfact
' arigasol.slhfac --> conta.linfact
'Actualizar la tabla ariges.scafac.inconta=1 para indicar que ya esta contabilizada
Dim b As Boolean
Dim cadMen As String
Dim SQL As String
Dim vsocio As CSocio
Dim codsoc As Long
Dim Rs As ADODB.Recordset

Dim Rsx As ADODB.Recordset
Dim sql2 As String
Dim codfor As Integer
Dim TipForpa As String
Dim PorIva As String

    On Error GoTo EContab

    If CInt(tipo) = 0 Then
        ConnContaCVV.BeginTrans
    Else
        ConnContaCV.BeginTrans
    End If
    
    conn.BeginTrans
     
'    PorIva = ""
'    PorIva = DevuelveDesdeBDNew(cContaTel, "tiposiva", "porceiva", "codigiva", Tipiva, "N")
    'Insertar en la conta Cabecera Factura
    b = InsertarFacturaCV(cadwhere, tipo, cadMen)
    cadMen = "Insertando Factura: " & cadMen
    
    ' insertar en tesoreria
    If b Then
        SQL = "select * from cvfacturas where " & cadwhere
        Set Rsx = New ADODB.Recordset
        Rsx.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        
        cadMen = ""
        b = InsertarEnTesoreriaNew4(Rsx, FecVenci, tipo, CtaBanco, cadMen)
        
        cadMen = "Insertando en Tesoreria: " & cadMen
    End If

    If b Then
        If b Then
            'Poner intconta=1 en arigasol.scafac
            cadMen = ""
            b = ActualizarCabFact("cvfacturas", cadwhere, cadMen)
            cadMen = "Actualizando Factura: " & cadMen
        End If
    End If
    

EContab:
    If Err.Number <> 0 Then
        b = False
        MuestraError Err.Number, "Contabilizando Factura", Err.Description
    End If
    If b Then
        If CInt(tipo) = 0 Then
            ConnContaCVV.CommitTrans
        Else
            ConnContaCV.CommitTrans
        End If
        conn.CommitTrans
        PasarFacturaCV = True
    Else
        If CInt(tipo) = 0 Then
            ConnContaCVV.RollbackTrans
        Else
            ConnContaCV.RollbackTrans
        End If
        conn.RollbackTrans
        
        SQL = "Insert into tmperrfac "
        SQL = SQL & " Select *, " & DBSet(cadMen, "T") & "  as error From tmpfactu "
        SQL = SQL & " WHERE (numserie,numfactu,fecfactu) in (select letraser,numfactu,fecfactu from cvfacturas where " & cadwhere & ")"
        conn.Execute SQL
        
        
        PasarFacturaCV = False
    End If
End Function





Public Function PasarFacturaFac(cadwhere As String, FecVenci As String, CtaBanco As String, CodCCost As String) As Boolean
'Insertar en tablas cabecera/lineas de la Contabilidad la factura
' ariagroutil.cabfact --> conta.cabfact
' ariagroutil.linfact --> conta.linfact
'Actualizar la tabla ariagroutil.cabfact.inconta=1 para indicar que ya esta contabilizada

Dim b As Boolean
Dim cadMen As String
Dim SQL As String
Dim vsocio As CSocio
Dim codsoc As Long
Dim Rs As ADODB.Recordset

Dim Rsx As ADODB.Recordset
Dim sql2 As String
Dim codfor As Integer
Dim TipForpa As String
Dim PorIva As String

    On Error GoTo EContab

    ConnContaFac.BeginTrans
    conn.BeginTrans
     
    'Insertar en la conta Cabecera Factura
    b = InsertarCabFactFac(cadwhere, cadMen)
    cadMen = "Insertando Cab. Factura: " & cadMen
    
    ' insertar en tesoreria
    If b Then
        SQL = "select * from cabfact where " & cadwhere
        Set Rsx = New ADODB.Recordset
        Rsx.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        
        b = InsertarEnTesoreriaNewFac(Rsx, FecVenci, CtaBanco, "")
        cadMen = "Insertando en Tesoreria: " & cadMen
        Set Rsx = Nothing
    End If

    If b Then
        'Insertar lineas de Factura en la Conta
        If vParamAplic.ContabilidadNueva Then
            b = InsertarLinFactFacContaNueva("linfact", Replace(cadwhere, "cabfact", "linfact"), cadMen, CodCCost)
        Else
            b = InsertarLinFactFac("linfact", Replace(cadwhere, "cabfact", "linfact"), cadMen, CodCCost)
        End If
        cadMen = "Insertando Lin. Factura: " & cadMen

        If b Then
            'Poner intconta=1 en ariagroutil.cabfact
            b = ActualizarCabFact("cabfact", cadwhere, cadMen)
            cadMen = "Actualizando Factura: " & cadMen
        End If
    End If
    
    If Not b Then
        SQL = "Insert into tmperrfac(codtipom,numfactu,fecfactu,error) "
        SQL = SQL & " Select *," & DBSet(cadMen, "T") & " as error From tmpfactu "
        SQL = SQL & " WHERE " & Replace(cadwhere, "cabfact", "tmpfactu")
        conn.Execute SQL
    End If
    
EContab:
    If Err.Number <> 0 Then
        b = False
        MuestraError Err.Number, "Contabilizando Factura", Err.Description
    End If
    If b Then
        ConnContaFac.CommitTrans
        conn.CommitTrans
        PasarFacturaFac = True
    Else
        ConnContaFac.RollbackTrans
        conn.RollbackTrans
        PasarFacturaFac = False
    End If
End Function


Public Function PasarFacturaGas(cadwhere As String, fecfactu As String, NumAsien As String, NumLinea As String, CodCCost As String, ByRef Diferencia As Currency) As Boolean
'Insertar en tablas cabecera/lineas de la Contabilidad la factura
' arigasol.schfac --> conta.cabfact
' arigasol.slhfac --> conta.linfact
'Actualizar la tabla ariges.scafac.inconta=1 para indicar que ya esta contabilizada
Dim b As Boolean
Dim cadMen As String
Dim SQL As String
Dim vsocio As CSocio
Dim codsoc As Long
Dim Rs As ADODB.Recordset

Dim Rsx As ADODB.Recordset
Dim sql2 As String
Dim codfor As Integer
Dim TipForpa As String
'Dim PorIva As String
Dim CuentaSocio As String

Dim numdocum As String
Dim ampliacion As String
Dim ampliaciond As String
Dim ampliacionh As String
Dim ImporteD As Currency
Dim ImporteH As Currency
Dim Cad As String

    On Error GoTo EContab

'    PorIva = ""
'    PorIva = DevuelveDesdeBDNew(cContaGas, "tiposiva", "porceiva", "codigiva", vParamAplic.CodIvaGas, "N")
    
    'Insertar en la conta Cabecera Factura
    b = InsertarCabFactGas(cadwhere, vParamAplic.CodIvaGas, cadMen)
    cadMen = "Insertando Cab. Factura: " & cadMen
    
    ' insertar en linea de asiento
    If b Then
        
        SQL = ""
        'i = 0
        
        ImporteD = 0
        ImporteH = 0
        
        ampliacion = "Fact.Gasolinera"
        ampliaciond = Trim(DevuelveDesdeBDNew(cContaGas, "conceptos", "nomconce", "codconce", vParamAplic.ConceDebeGas, "N")) & " " & ampliacion
        ampliacionh = Trim(DevuelveDesdeBDNew(cContaGas, "conceptos", "nomconce", "codconce", vParamAplic.ConceHaberGas, "N")) & " " & ampliacion

        ' ******************IMPORTE de la poliza
        
        SQL = "select * from gascabfac where " & cadwhere
        Set Rs = New ADODB.Recordset
        Rs.Open SQL, conn, adOpenStatic, adLockPessimistic, adCmdText
        
        numdocum = DBLet(Rs!numfactu, "N")
        
        CuentaSocio = Trim(vParamAplic.RaizCtaSocGas & Format(Rs!Codsocio, Repeat("0", vEmpresaGas.DigitosUltimoNivel - vEmpresaGas.DigitosNivelAnterior)))
        
        Cad = DBSet(vParamAplic.NumDiarioGas, "N") & "," & DBSet(fecfactu, "F") & "," & DBSet(NumAsien, "N") & ","
        Cad = Cad & DBSet(NumLinea, "N") & "," & DBSet(CuentaSocio, "T") & "," & DBSet(numdocum, "T") & ","
        
        ' COMPROBAMOS EL SIGNO DEL IMPORTE PQ NO PERMITIMOS INTRODUCIR APUNTES CON IMPORTES NEGATIVOS
        If DBLet(Rs!total, "N") < 0 Then
            ' importe al debe en negativo, cambiamos el signo
            Cad = Cad & DBSet(vParamAplic.ConceDebeGas, "N") & "," & DBSet(ampliaciond, "T") & "," & DBSet(Rs!total * (-1), "N") & "," & ValorNulo & ","
            Cad = Cad & DBSet(CodCCost, "T") & "," & DBSet(vParamAplic.CtaContraGas, "T") & "," & ValorNulo & ",0"
        
            ImporteD = ImporteD + (CCur(Rs!total) * (-1))
        Else
            ' importe al haber en positivo
            Cad = Cad & DBSet(vParamAplic.ConceHaberGas, "N") & "," & DBSet(ampliacionh, "T") & "," & ValorNulo & "," & DBSet(Rs!total, "N") & ","
            Cad = Cad & DBSet(CodCCost, "T") & "," & DBSet(vParamAplic.CtaContraGas, "T") & "," & ValorNulo & ",0"
        
            ImporteH = ImporteH + CCur(Rs!total)
        End If
        
        Diferencia = Diferencia + (ImporteH - ImporteD)
        
        Cad = "(" & Cad & ")"
                     
        b = InsertarLinAsientoDia(Cad, cadMen, cContaGas)
        cadMen = "Insertando Linea Asiento: " & cadMen
    End If

    If b Then
        'Insertar lineas de Factura en la Conta
        b = InsertarLinFactGas("gascabfac", cadwhere, cadMen, CodCCost)
        cadMen = "Insertando Lin. Factura: " & cadMen

        If b Then
            'Poner intconta=1 en arigasol.gascabfac
            b = ActualizarCabFact("gascabfac", cadwhere, cadMen)
            cadMen = "Actualizando Factura: " & cadMen
        End If
    End If
    
    If Not b Then
        SQL = "Insert into tmperrfac(codtipom,numfactu,fecfactu,error) "
        SQL = SQL & " Select *," & DBSet(cadMen, "T") & " as error From tmpfactu "
        SQL = SQL & " WHERE " & Replace(cadwhere, "gascabfac", "tmpfactu")
        conn.Execute SQL
    End If
    
EContab:
    If Err.Number <> 0 Then
        b = False
        MuestraError Err.Number, "Contabilizando Factura", Err.Description
    End If
    If b Then
'        ConnContaGas.CommitTrans
'        conn.CommitTrans
        PasarFacturaGas = True
    Else
'        ConnContaGas.RollbackTrans
'        conn.RollbackTrans
        PasarFacturaGas = False
    End If
End Function



Private Function InsertarCabFact(cadwhere As String, TipoIva As String, caderr As String, PorcIva As String, FPago As String) As Boolean
'Insertando en tabla conta.cabfact
Dim SQL As String
Dim Rs As ADODB.Recordset
Dim Cad As String
Dim SqlDatos As String
Dim RsDatos As ADODB.Recordset
Dim sql2 As String
Dim CadenaInsertFaclin2 As String

    On Error GoTo EInsertar
    
    SQL = " SELECT numserie,numfactu,fecfactu,codmacta, year(fecfactu) as anofaccl,"
    SQL = SQL & "baseimpo,cuotaiva,totalfac "
    SQL = SQL & " FROM telmovil "
    SQL = SQL & " WHERE " & cadwhere
    
    Set Rs = New ADODB.Recordset
    Rs.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    Cad = ""
    If Not Rs.EOF Then
        'guardamos estos valores para utilizarlos cuando insertemos las lineas de la factura
        SQL = ""
        SQL = DBSet(Trim(Rs!numserie), "T") & "," & DBSet(Rs!numfactu, "N") & "," & DBSet(Rs!fecfactu, "F") & "," & DBSet(Rs!Codmacta, "T") & "," & Year(Rs!fecfactu) & ",'FACTURACION',"
        
        If Not vParamAplic.ContabilidadNueva Then
            SQL = SQL & DBSet(Rs!BaseImpo, "N") & "," & ValorNulo & "," & ValorNulo & "," & DBSet(PorcIva, "N") & "," & ValorNulo & "," & ValorNulo & ","
            SQL = SQL & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & DBSet(Rs!CuotaIva, "N", "N") & "," & ValorNulo & "," & ValorNulo & ","
            SQL = SQL & ValorNulo & "," & ValorNulo & "," & ValorNulo & ","
            SQL = SQL & DBSet(Rs!TotalFac, "N") & "," & DBSet(TipoIva, "N") & "," & ValorNulo & "," & ValorNulo & ",0,"
            SQL = SQL & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & ","
            SQL = SQL & DBSet(Rs!fecfactu, "F")
            Cad = Cad & "(" & SQL & ")"
    '        RS.MoveNext
        
            'Insertar en la contabilidad
            SQL = "INSERT INTO cabfact (numserie,codfaccl,fecfaccl,codmacta,anofaccl,confaccl,ba1faccl,ba2faccl,ba3faccl,"
            SQL = SQL & "pi1faccl,pi2faccl,pi3faccl,pr1faccl,pr2faccl,pr3faccl,ti1faccl,ti2faccl,ti3faccl,tr1faccl,tr2faccl,tr3faccl,"
            SQL = SQL & "totfaccl,tp1faccl,tp2faccl,tp3faccl,intracom,retfaccl,trefaccl,cuereten,numdiari,fechaent,numasien,fecliqcl) "
            SQL = SQL & " VALUES " & Cad
            ConnContaTel.Execute SQL
        Else
            SqlDatos = "select * from cuentas where codmacta = " & DBSet(Rs!ctaclien, "T")
            Set RsDatos = New ADODB.Recordset
            RsDatos.Open SqlDatos, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            If Not RsDatos.EOF Then
                SQL = SQL & "'0'"
                SQL = SQL & "0," & DBSet(FPago, "N") & "," & DBSet(Rs!BaseImpo, "N") & "," & ValorNulo & "," & DBSet(Rs!CuotaIva, "N") & ","
                SQL = SQL & ValorNulo & "," & DBSet(Rs!TotalFac, "N") & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & ",0," & DBSet(Rs!fecfactu, "F") & ","
                SQL = SQL & DBSet(RsDatos!Nommacta, "T", "S") & "," & DBSet(RsDatos!dirdatos, "T", "S") & "," & DBSet(RsDatos!desPobla, "T", "S") & ","
                SQL = SQL & DBSet(RsDatos!Codposta, "T", "S") & "," & DBSet(RsDatos!desProvi, "T", "S") & "," & DBSet(RsDatos!nifdatos, "T", "S") & ",'ES')"
            Else
                SQL = SQL & "'0'"
                SQL = SQL & "0," & DBSet(FPago, "N") & "," & DBSet(Rs!BaseImpo, "N") & "," & ValorNulo & "," & DBSet(Rs!CuotaIva, "N") & ","
                SQL = SQL & ValorNulo & "," & DBSet(Rs!TotalFac, "N") & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & ",0," & DBSet(Rs!fecfactu, "F") & ","
                SQL = SQL & ValorNulo & "," & ValorNulo & "," & ValorNulo & ","
                SQL = SQL & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & ")"
            End If
            Set RsDatos = Nothing
            
            SQL = "(" & SQL & ")"
            
            sql2 = "INSERT INTO factcli (numserie,numfactu,fecfactu,codmacta,anofactu,observa,codconce340,codopera,codforpa,totbases,totbasesret,totivas,"
            sql2 = sql2 & "totrecargo,totfaccl, retfaccl,trefaccl,cuereten,tiporeten,fecliqcl,nommacta,dirdatos,codpobla,despobla, desprovi,nifdatos,"
            sql2 = sql2 & "codpais,codagente)"
            sql2 = sql2 & " VALUES " & SQL
            
            ConnContaTel.Execute sql2
            
            CadenaInsertFaclin2 = ""
            
            'numlinea,baseimpo,codigiva,porciva,porcrec,impoiva,imporec)
            'IVA 1, siempre existe
            sql2 = "'" & Rs!numserie & "'," & Rs!numfactu & "," & DBSet(Rs!fecfactu, "F") & "," & Year(Rs!fecfactu) & ","
            sql2 = sql2 & "1," & DBSet(Rs!BaseImpo, "N") & "," & DBSet(TipoIva, "N") & "," & DBSet(PorcIva, "N") & ","
            sql2 = sql2 & ValorNulo & "," & DBSet(Rs!CuotaIva, "N") & "," & ValorNulo
            CadenaInsertFaclin2 = CadenaInsertFaclin2 & "(" & sql2 & ")"
            
            'para las lineas
            vTipoIva(0) = TipoIva
            vPorcIva(0) = PorcIva
            vPorcRec(0) = 0
            vImpIva(0) = Rs!CuotaIva
            vImpRec(0) = 0
            vBaseIva(0) = Rs!BaseImpo
    
            SQL = "INSERT INTO factcli_totales(numserie,numfactu,fecfactu,anofactu,numlinea,baseimpo,codigiva,"
            SQL = SQL & "porciva,porcrec,impoiva,imporec) VALUES " & CadenaInsertFaclin2
            ConnContaTel.Execute SQL
        
        End If
    End If
    Rs.Close
    Set Rs = Nothing
    
    
EInsertar:
    If Err.Number <> 0 Then
        InsertarCabFact = False
        caderr = Err.Description
    Else
        InsertarCabFact = True
    End If
End Function

Private Function InsertarFacturaCV(cadwhere As String, tipo As String, caderr As String) As Boolean
'Insertando en tabla conta.cabfact
Dim SQL As String
Dim Rs As ADODB.Recordset
Dim Cad As String
Dim CodIva2 As String
Dim CodIva3 As String
Dim PorcIva2 As String
Dim PorcIva3 As String
Dim BaseImpo As Currency
Dim numfactu As Long

Dim Mc As CContadorContab

    On Error GoTo EInsertar
    
    SQL = " SELECT letraser,numfactu,fecfactu,codmactasoc,codmactavta,  year(fecfactu) as anofac,"
    SQL = SQL & "baseimpo,porciva,codiva,cuotaiva,baseimpo2,porciva2,codiva2,cuotaiva2,baseimpo3,porciva3,codiva3,cuotaiva3,totalfac "
    SQL = SQL & " FROM cvfacturas "
    SQL = SQL & " WHERE " & cadwhere
    
    Set Rs = New ADODB.Recordset
    Rs.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    Cad = ""
    If Not Rs.EOF Then
        'guardamos estos valores para utilizarlos cuando insertemos las lineas de la factura
        CodIva2 = ValorNulo
        CodIva3 = ValorNulo
        PorcIva2 = 0
        PorcIva3 = 0
        If DBLet(Rs!BaseImpo2, "N") <> 0 Then
            CodIva2 = DBSet(Rs!CodIva2, "N")
            PorcIva2 = DBSet(Rs!PorcIva2, "N")
        End If
        If DBLet(Rs!BaseImpo3, "N") <> 0 Then
            CodIva3 = DBSet(Rs!CodIva3, "N")
            PorcIva3 = DBSet(Rs!PorcIva3, "N")
        End If
'        If Rs!NumFactu = "FC00003781" Then
'            MsgBox "a"
'        End If
        
        If tipo <= 1 Then
            ' factura de ventas
            ' cabecera
            If tipo = 1 Then
                numfactu = Trim(Right(Rs!numfactu, 7))
            Else
                numfactu = DBLet(Rs!numfactu)
            End If
            
            SQL = ""
            SQL = DBSet(Trim(Rs!letraser), "T") & "," & DBSet(numfactu, "N") & "," & DBSet(Rs!fecfactu, "F") & "," & DBSet(Rs!CodmactaSoc, "T") & "," & Year(Rs!fecfactu) & ",'FACTURACION',"
            SQL = SQL & DBSet(Rs!BaseImpo, "N") & "," & DBSet(Rs!BaseImpo2, "N", "S") & "," & DBSet(Rs!BaseImpo3, "N", "S") & "," & DBSet(Rs!PorcIva, "N") & "," & DBSet(PorcIva2, "N", "S") & "," & DBSet(PorcIva3, "N", "S") & ","
            SQL = SQL & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & DBSet(Rs!CuotaIva, "N", "N") & "," & DBSet(Rs!cuotaiva2, "N", "S") & "," & DBSet(Rs!cuotaiva3, "N", "S") & ","
            SQL = SQL & ValorNulo & "," & ValorNulo & "," & ValorNulo & ","
            SQL = SQL & DBSet(Rs!TotalFac, "N") & "," & DBSet(Rs!CodIVA, "N") & "," & CodIva2 & "," & CodIva3 & ",0,"
            SQL = SQL & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & ","
            SQL = SQL & DBSet(Rs!fecfactu, "F")
            Cad = Cad & "(" & SQL & ")"
        
            'Insertar en la contabilidad
            SQL = "INSERT INTO cabfact (numserie,codfaccl,fecfaccl,codmacta,anofaccl,confaccl,ba1faccl,ba2faccl,ba3faccl,"
            SQL = SQL & "pi1faccl,pi2faccl,pi3faccl,pr1faccl,pr2faccl,pr3faccl,ti1faccl,ti2faccl,ti3faccl,tr1faccl,tr2faccl,tr3faccl,"
            SQL = SQL & "totfaccl,tp1faccl,tp2faccl,tp3faccl,intracom,retfaccl,trefaccl,cuereten,numdiari,fechaent,numasien,fecliqcl) "
            SQL = SQL & " VALUES " & Cad
            If tipo = 0 Then
                ConnContaCVV.Execute SQL
            Else
                ConnContaCV.Execute SQL
            End If
            
            ' linea de factura de ventas
            BaseImpo = DBLet(Rs!BaseImpo, "N") + DBLet(Rs!BaseImpo2, "N") + DBLet(Rs!BaseImpo3, "N")
            
            SQL = ""
            SQL = "'" & Trim(Rs!letraser) & "'," & DBSet(numfactu, "N") & "," & Year(Rs!fecfactu) & ",1,"
            SQL = SQL & DBSet(Rs!CodmactaVta, "T")
            SQL = SQL & "," & DBSet(BaseImpo, "N") & ","
            If CCoste = "" Then
                SQL = SQL & ValorNulo
            Else
                SQL = SQL & DBSet(CCoste, "T")
            End If
        
            Cad = "(" & SQL & ")"
            'Insertar en la contabilidad
            If Cad <> "" Then
                SQL = "INSERT INTO linfact (numserie,codfaccl,anofaccl,numlinea,codtbase,impbascl,codccost) "
                SQL = SQL & " VALUES " & Cad
                If tipo = 0 Then
                    ConnContaCVV.Execute SQL
                Else
                    ConnContaCV.Execute SQL
                End If
            End If
            
        Else
            ' factura de compras
            ' cabecera
            Set Mc = New CContadorContab
            
            If Mc.ConseguirContador("1", (Rs!fecfactu <= CDate(FFinCV)), True, cContaCV) = 0 Then
                SQL = ""
                SQL = Mc.Contador & "," & DBSet(Rs!fecfactu, "F") & "," & DBLet(Rs!anofac, "N") & "," & DBSet(Rs!fecfactu, "F") & "," & DBSet(Rs!numfactu, "T") & "," & DBSet(Rs!CodmactaSoc, "T") & "," & ValorNulo & ","
                SQL = SQL & DBSet(Rs!BaseImpo, "N") & "," & DBSet(Rs!BaseImpo2, "N", "S") & "," & DBSet(Rs!BaseImpo3, "N", "S") & ","
                SQL = SQL & DBSet(Rs!PorcIva, "N") & "," & DBSet(PorcIva2, "N", "S") & "," & DBSet(PorcIva3, "N", "S") & ","
                SQL = SQL & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & DBSet(Rs!CuotaIva, "N") & "," & DBSet(Rs!cuotaiva2, "N", "S") & "," & DBSet(Rs!cuotaiva3, "N", "S") & ","
                SQL = SQL & ValorNulo & "," & ValorNulo & "," & ValorNulo & ","
                SQL = SQL & DBSet(Rs!TotalFac, "N") & "," & DBSet(Rs!CodIVA, "N") & "," & CodIva2 & "," & CodIva3 & ",0,"
                SQL = SQL & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & DBSet(Rs!fecfactu, "F") & ",0"
                Cad = Cad & "(" & SQL & ")"
                
                'Insertar en la contabilidad
                SQL = "INSERT INTO cabfactprov (numregis,fecfacpr,anofacpr,fecrecpr,numfacpr,codmacta,confacpr,ba1facpr,ba2facpr,ba3facpr,"
                SQL = SQL & "pi1facpr,pi2facpr,pi3facpr,pr1facpr,pr2facpr,pr3facpr,ti1facpr,ti2facpr,ti3facpr,tr1facpr,tr2facpr,tr3facpr,"
                SQL = SQL & "totfacpr,tp1facpr,tp2facpr,tp3facpr,extranje,retfacpr,trefacpr,cuereten,numdiari,fechaent,numasien,fecliqpr,nodeducible) "
                SQL = SQL & " VALUES " & Cad
                ConnContaCV.Execute SQL
            End If
            'linea
            
            ' linea de factura de ventas
            BaseImpo = DBLet(Rs!BaseImpo, "N") + DBLet(Rs!BaseImpo2, "N") + DBLet(Rs!BaseImpo3, "N")
            
            SQL = ""
            SQL = Mc.Contador & "," & Year(Rs!fecfactu) & ",1,"
            SQL = SQL & DBSet(Rs!CodmactaVta, "T")
            SQL = SQL & "," & DBSet(BaseImpo, "N") & ","
            
            If CCoste = "" Then
                SQL = SQL & ValorNulo
            Else
                SQL = SQL & DBSet(CCoste, "T")
            End If
            Cad = "(" & SQL & ")"
            
            'Insertar en la contabilidad
            If Cad <> "" Then
                SQL = "INSERT INTO linfactprov (numregis,anofacpr,numlinea,codtbase,impbaspr,codccost) "
                SQL = SQL & " VALUES " & Cad
                ConnContaCV.Execute SQL
            End If
                    
        End If
    End If
    Rs.Close
    Set Rs = Nothing
    
    
EInsertar:
    If Err.Number <> 0 Then
        InsertarFacturaCV = False
        caderr = Err.Description
    Else
        InsertarFacturaCV = True
    End If
End Function




Private Function InsertarCabFactGas(cadwhere As String, TipoIva As String, caderr As String) As Boolean
'Insertando en tabla conta.cabfact
Dim SQL As String
Dim Rs As ADODB.Recordset
Dim Cad As String
Dim Codmacta As String
Dim FPago As String
Dim RsDatos As ADODB.Recordset
Dim SqlDatos As String
Dim sql2 As String
Dim CadenaInsertFaclin2 As String



    On Error GoTo EInsertar
    
    SQL = " SELECT letraser,numfactu,fecfactu,codsocio, year(fecfactu) as anofaccl,"
    SQL = SQL & "base as baseimpo,iva as cuotaiva,total as totalfac, codiva, porciva "
    SQL = SQL & " FROM gascabfac "
    SQL = SQL & " WHERE " & cadwhere
    
    Set Rs = New ADODB.Recordset
    Rs.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    Cad = ""
    If Not Rs.EOF Then
        'guardamos estos valores para utilizarlos cuando insertemos las lineas de la factura
        
        Codmacta = vParamAplic.RaizCtaSocGas & Right("0000000000" & DBLet(Rs!Codsocio, "N"), vEmpresaGas.DigitosUltimoNivel - vEmpresaGas.DigitosNivelAnterior)
        
        SQL = ""
        SQL = DBSet(Trim(Rs!letraser), "T") & "," & DBSet(Rs!numfactu, "N") & "," & DBSet(Rs!fecfactu, "F") & "," & DBSet(Codmacta, "T") & "," & Year(Rs!fecfactu) & ",'FACTURACION',"
        
        If Not vParamAplic.ContabilidadNueva Then
            SQL = SQL & DBSet(Rs!BaseImpo, "N") & "," & ValorNulo & "," & ValorNulo & "," & DBSet(Rs!PorcIva, "N") & "," & ValorNulo & "," & ValorNulo & ","
            SQL = SQL & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & DBSet(Rs!CuotaIva, "N", "N") & "," & ValorNulo & "," & ValorNulo & ","
            SQL = SQL & ValorNulo & "," & ValorNulo & "," & ValorNulo & ","
            SQL = SQL & DBSet(Rs!TotalFac, "N") & "," & DBSet(Rs!CodIVA, "N") & "," & ValorNulo & "," & ValorNulo & ",0,"
            SQL = SQL & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & ","
            SQL = SQL & DBSet(Rs!fecfactu, "F")
            Cad = Cad & "(" & SQL & ")"
    
            'Insertar en la contabilidad
            SQL = "INSERT INTO cabfact (numserie,codfaccl,fecfaccl,codmacta,anofaccl,confaccl,ba1faccl,ba2faccl,ba3faccl,"
            SQL = SQL & "pi1faccl,pi2faccl,pi3faccl,pr1faccl,pr2faccl,pr3faccl,ti1faccl,ti2faccl,ti3faccl,tr1faccl,tr2faccl,tr3faccl,"
            SQL = SQL & "totfaccl,tp1faccl,tp2faccl,tp3faccl,intracom,retfaccl,trefaccl,cuereten,numdiari,fechaent,numasien,fecliqcl) "
            SQL = SQL & " VALUES " & Cad
            ConnContaGas.Execute SQL
        Else
            FPago = DevuelveDesdeBDNew(cContaGas, "formapago", "min(codforpa)", "", "")
        
            SqlDatos = "select * from cuentas where codmacta = " & DBSet(Codmacta, "T")
            Set RsDatos = New ADODB.Recordset
            RsDatos.Open SqlDatos, ConnContaGas, adOpenForwardOnly, adLockPessimistic, adCmdText
            If Not RsDatos.EOF Then
                SQL = SQL & "'0',"
                SQL = SQL & "0," & DBSet(FPago, "N") & "," & DBSet(Rs!BaseImpo, "N") & "," & ValorNulo & "," & DBSet(Rs!CuotaIva, "N") & ","
                SQL = SQL & ValorNulo & "," & DBSet(Rs!TotalFac, "N") & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & ",0," & DBSet(Rs!fecfactu, "F") & ","
                SQL = SQL & DBSet(RsDatos!Nommacta, "T", "S") & "," & DBSet(RsDatos!dirdatos, "T", "S") & "," & DBSet(RsDatos!desPobla, "T", "S") & ","
                SQL = SQL & DBSet(RsDatos!Codposta, "T", "S") & "," & DBSet(RsDatos!desProvi, "T", "S") & "," & DBSet(RsDatos!nifdatos, "T", "S") & ",'ES'"
            Else
                SQL = SQL & "'0',"
                SQL = SQL & "0," & DBSet(FPago, "N") & "," & DBSet(Rs!BaseImpo, "N") & "," & ValorNulo & "," & DBSet(Rs!CuotaIva, "N") & ","
                SQL = SQL & ValorNulo & "," & DBSet(Rs!TotalFac, "N") & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & ",0," & DBSet(Rs!fecfactu, "F") & ","
                SQL = SQL & ValorNulo & "," & ValorNulo & "," & ValorNulo & ","
                SQL = SQL & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo
            End If
            Set RsDatos = Nothing
            SQL = "(" & SQL & ")"
            
            sql2 = "INSERT INTO factcli (numserie,numfactu,fecfactu,codmacta,anofactu,observa,codconce340,codopera,codforpa,totbases,totbasesret,totivas,"
            sql2 = sql2 & "totrecargo,totfaccl, retfaccl,trefaccl,cuereten,tiporeten,fecliqcl,nommacta,dirdatos,codpobla,despobla, desprovi,nifdatos,"
            sql2 = sql2 & "codpais)"
            sql2 = sql2 & " VALUES " & SQL
            
            ConnContaGas.Execute sql2
            
            CadenaInsertFaclin2 = ""
            
            'numlinea,baseimpo,codigiva,porciva,porcrec,impoiva,imporec)
            'IVA 1, siempre existe
            sql2 = "'" & Rs!letraser & "'," & Rs!numfactu & "," & DBSet(Rs!fecfactu, "F") & "," & Year(Rs!fecfactu) & ","
            sql2 = sql2 & "1," & DBSet(Rs!BaseImpo, "N") & "," & DBSet(Rs!CodIVA, "N") & "," & DBSet(Rs!PorcIva, "N") & ","
            sql2 = sql2 & ValorNulo & "," & DBSet(Rs!CuotaIva, "N") & "," & ValorNulo
            CadenaInsertFaclin2 = CadenaInsertFaclin2 & "(" & sql2 & ")"
            
            'para las lineas
            vTipoIva(0) = Rs!CodIVA
            vPorcIva(0) = Rs!PorcIva
            vPorcRec(0) = 0
            vImpIva(0) = Rs!CuotaIva
            vImpRec(0) = 0
            vBaseIva(0) = Rs!BaseImpo
    
            SQL = "INSERT INTO factcli_totales(numserie,numfactu,fecfactu,anofactu,numlinea,baseimpo,codigiva,"
            SQL = SQL & "porciva,porcrec,impoiva,imporec) VALUES " & CadenaInsertFaclin2
            ConnContaGas.Execute SQL
        
        End If
    
    
    End If
    Rs.Close
    Set Rs = Nothing
    
    
    
EInsertar:
    If Err.Number <> 0 Then
        InsertarCabFactGas = False
        caderr = Err.Description
    Else
        InsertarCabFactGas = True
    End If
End Function



Private Function InsertarCabFactFac(cadwhere As String, caderr As String) As Boolean
'Insertando en tabla conta.cabfact
Dim SQL As String
Dim Rs As ADODB.Recordset
Dim Cad As String
Dim SqlDatos As String
Dim RsDatos As ADODB.Recordset
Dim sql2 As String
Dim CadenaInsertFaclin2 As String

    On Error GoTo EInsertar
    
    SQL = " SELECT letraser,numfactu,fecfactu,ctaclien, year(fecfactu) as anofaccl,"
    SQL = SQL & "baseiva1, baseiva2, baseiva3, impoiva1, impoiva2, impoiva3, imporec1,"
    SQL = SQL & "imporec2, imporec3, totalfac, tipoiva1, tipoiva2, tipoiva3, porciva1,"
    SQL = SQL & "porciva2, porciva3, porcrec1, porcrec2, porcrec3, totalfac, retfaccl, "
    SQL = SQL & "trefaccl, cuereten, codforpa "
    SQL = SQL & " FROM cabfact "
    SQL = SQL & " WHERE " & cadwhere
    
    Set Rs = New ADODB.Recordset
    Rs.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    Cad = ""
    If Not Rs.EOF Then
        'guardamos estos valores para utilizarlos cuando insertemos las lineas de la factura
        
        SQL = ""
        SQL = DBSet(Trim(Rs!letraser), "T") & "," & DBSet(Rs!numfactu, "N") & "," & DBSet(Rs!fecfactu, "F") & "," & DBSet(Rs!ctaclien, "T") & "," & Year(Rs!fecfactu) & ",'FACTURACION',"
        
        BaseImp = Rs!baseiva1 + CCur(DBLet(Rs!baseiva2, "N")) + CCur(DBLet(Rs!baseiva3, "N"))
        IvaImp = DBLet(Rs!impoiva1, "N") + DBLet(Rs!impoiva2, "N") + DBLet(Rs!impoiva3, "N")
        
        
        If Not vParamAplic.ContabilidadNueva Then
        
            SQL = SQL & DBSet(Rs!baseiva1, "N") & "," & DBSet(Rs!baseiva2, "N") & "," & DBSet(Rs!baseiva3, "N") & "," & DBSet(Rs!porciva1, "N") & "," & DBSet(Rs!PorcIva2, "N") & "," & DBSet(Rs!PorcIva3, "N") & ","
            SQL = SQL & DBSet(Rs!porcrec1, "N") & "," & DBSet(Rs!porcrec2, "N") & "," & DBSet(Rs!porcrec3, "N") & "," & DBSet(Rs!impoiva1, "N") & "," & DBSet(Rs!impoiva2, "N") & "," & DBSet(Rs!impoiva3, "N") & ","
            SQL = SQL & DBSet(Rs!imporec1, "N") & "," & DBSet(Rs!imporec2, "N") & "," & DBSet(Rs!imporec3, "N") & ","
            SQL = SQL & DBSet(Rs!TotalFac, "N") & "," & DBSet(Rs!tipoiva1, "N") & "," & DBSet(Rs!TipoIVA2, "N") & "," & DBSet(Rs!TipoIVA3, "N") & ",0,"
            SQL = SQL & DBSet(Rs!retfaccl, "N") & "," & DBSet(Rs!trefaccl, "N") & "," & DBSet(Rs!cuereten, "T") & ","
            SQL = SQL & ValorNulo & "," & ValorNulo & "," & ValorNulo & ","
            SQL = SQL & DBSet(Rs!fecfactu, "F")
        
            Cad = Cad & "(" & SQL & ")"
            
            'Insertar en la contabilidad
            SQL = "INSERT INTO cabfact (numserie,codfaccl,fecfaccl,codmacta,anofaccl,confaccl,ba1faccl,ba2faccl,ba3faccl,"
            SQL = SQL & "pi1faccl,pi2faccl,pi3faccl,pr1faccl,pr2faccl,pr3faccl,ti1faccl,ti2faccl,ti3faccl,tr1faccl,tr2faccl,tr3faccl,"
            SQL = SQL & "totfaccl,tp1faccl,tp2faccl,tp3faccl,intracom,retfaccl,trefaccl,cuereten,numdiari,fechaent,numasien,fecliqcl) "
            SQL = SQL & " VALUES " & Cad
            ConnContaFac.Execute SQL
            
        
        Else
            SqlDatos = "select * from cuentas where codmacta = " & DBSet(Rs!ctaclien, "T")
            Set RsDatos = New ADODB.Recordset
            RsDatos.Open SqlDatos, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            If Not RsDatos.EOF Then
                SQL = SQL & "'0'"
                SQL = SQL & "0," & DBSet(Rs!CodForpa, "N") & "," & DBSet(BaseImp, "N") & "," & ValorNulo & "," & DBSet(IvaImp, "N") & ","
                SQL = SQL & ValorNulo & "," & DBSet(Rs!TotalFac, "N") & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & ",0," & DBSet(Rs!fecfactu, "F") & ","
                SQL = SQL & DBSet(RsDatos!Nommacta, "T", "S") & "," & DBSet(RsDatos!dirdatos, "T", "S") & "," & DBSet(RsDatos!desPobla, "T", "S") & ","
                SQL = SQL & DBSet(RsDatos!Codposta, "T", "S") & "," & DBSet(RsDatos!desProvi, "T", "S") & "," & DBSet(RsDatos!nifdatos, "T", "S") & ",'ES')"
            Else
                SQL = SQL & "'0'"
                SQL = SQL & "0," & DBSet(Rs!CodForpa, "N") & "," & DBSet(BaseImp, "N") & "," & ValorNulo & "," & DBSet(IvaImp, "N") & ","
                SQL = SQL & ValorNulo & "," & DBSet(Rs!TotalFac, "N") & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & ",0," & DBSet(Rs!fecfactu, "F") & ","
                SQL = SQL & ValorNulo & "," & ValorNulo & "," & ValorNulo & ","
                SQL = SQL & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & ")"
            End If
            SQL = "(" & SQL & ")"
            
            sql2 = "INSERT INTO factcli (numserie,numfactu,fecfactu,codmacta,anofactu,observa,codconce340,codopera,codforpa,totbases,totbasesret,totivas,"
            sql2 = sql2 & "totrecargo,totfaccl, retfaccl,trefaccl,cuereten,tiporeten,fecliqcl,nommacta,dirdatos,codpobla,despobla, desprovi,nifdatos,"
            sql2 = sql2 & "codpais,codagente)"
            sql2 = sql2 & " VALUES " & SQL
            
            ConnContaFac.Execute sql2
            
            CadenaInsertFaclin2 = ""
            
            'numlinea,baseimpo,codigiva,porciva,porcrec,impoiva,imporec)
            'IVA 1, siempre existe
            sql2 = "'" & Rs!letraser & "'," & Rs!numfactu & "," & DBSet(Rs!fecfactu, "F") & "," & Year(Rs!fecfactu) & ","
            sql2 = sql2 & "1," & DBSet(Rs!baseimp1, "N") & "," & Rs!tipoiva1 & "," & DBSet(Rs!porciva1, "N") & ","
            sql2 = sql2 & ValorNulo & "," & DBSet(Rs!impoiva1, "N") & "," & ValorNulo
            CadenaInsertFaclin2 = CadenaInsertFaclin2 & "(" & sql2 & ")"
            
            'para las lineas
            vTipoIva(0) = Rs!tipoiva1
            vPorcIva(0) = Rs!porciva1
            vPorcRec(0) = DBLet(Rs!porcrec1, "N")
            vImpIva(0) = Rs!impoiva1
            vImpRec(0) = DBLet(Rs!imporec1, "N")
            vBaseIva(0) = Rs!baseimp1
            
            vTipoIva(1) = 0: vTipoIva(2) = 0
            
            If Not IsNull(Rs!PorcIva2) Then
                sql2 = "'" & Rs!letraser & "'," & Rs!numfactu & "," & DBSet(Rs!fecfactu, "F") & "," & Year(Rs!fecfactu) & ","
                sql2 = sql2 & "2," & DBSet(Rs!baseimp2, "N") & "," & Rs!codiiva2 & "," & DBSet(Rs!PorcIva2, "N") & ","
                sql2 = sql2 & ValorNulo & "," & DBSet(Rs!impoiva2, "N") & "," & ValorNulo
                CadenaInsertFaclin2 = CadenaInsertFaclin2 & " , (" & sql2 & ")"
                vTipoIva(1) = Rs!TipoIVA2
                vPorcIva(1) = Rs!PorcIva2
                vPorcRec(1) = DBLet(Rs!porcrec2, "N")
                vImpIva(1) = Rs!impoiva2
                vImpRec(1) = DBLet(Rs!imporec2, "N")
                vBaseIva(1) = Rs!baseimp2
            End If
            If Not IsNull(Rs!PorcIva3) Then
                sql2 = "'" & Rs!letraser & "'," & Rs!numfactu & "," & DBSet(Rs!fecfactu, "F") & "," & Year(Rs!fecfactu) & ","
                sql2 = sql2 & "3," & DBSet(Rs!baseimp3, "N") & "," & Rs!codiiva3 & "," & DBSet(Rs!PorcIva3, "N") & ","
                sql2 = sql2 & ValorNulo & "," & DBSet(Rs!impoiva3, "N") & "," & ValorNulo
                CadenaInsertFaclin2 = CadenaInsertFaclin2 & " , (" & sql2 & ")"
                vTipoIva(2) = Rs!TipoIVA3
                vPorcIva(2) = Rs!PorcIva3
                vPorcRec(2) = DBLet(Rs!porcrec3, "N")
                vImpIva(2) = Rs!impoiva3
                vImpRec(2) = DBLet(Rs!imporec3, "N")
                vBaseIva(2) = Rs!baseimp3
            End If
    
            SQL = "INSERT INTO factcli_totales(numserie,numfactu,fecfactu,anofactu,numlinea,baseimpo,codigiva,"
            SQL = SQL & "porciva,porcrec,impoiva,imporec) VALUES " & CadenaInsertFaclin2
            ConnContaFac.Execute SQL
            
            
        End If
            
'        RS.MoveNext
    End If
    Rs.Close
    Set Rs = Nothing
    
    
    
EInsertar:
    If Err.Number <> 0 Then
        InsertarCabFactFac = False
        caderr = Err.Description
    Else
        InsertarCabFactFac = True
    End If
End Function



Private Function InsertarLinFact(cadTABLA As String, cadwhere As String, caderr As String, CtaVenta As String, CCoste As String, Tipiva As String, PorcIva As String) As Boolean
'cadWHere: selecciona un registro de scafac
'codtipom=x and numfactu=y and fecfactu=z
Dim SQL As String
Dim Rs As ADODB.Recordset
Dim Cad As String, Aux As String
Dim i As Long
Dim totimp As Currency, ImpLinea As Currency
Dim CodIVA As String
Dim Iva As String
Dim vIva As Currency


    On Error GoTo EInLinea
    SQL = " SELECT numserie,numfactu,fecfactu,codmacta, "
    SQL = SQL & "baseimpo,cuotaiva,totalfac "
    SQL = SQL & " FROM telmovil "

    SQL = " SELECT numserie, numfactu, fecfactu, baseimpo,year(fecfactu) as anofaccl,cuotaiva,totalfac from " & cadTABLA & " where " & cadwhere
    Set Rs = New ADODB.Recordset
    Rs.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText

    Cad = ""
    i = 1
    totimp = 0
    If Not Rs.EOF Then
        'concatenamos linea para insertar en la tabla de conta.linfact
        SQL = ""
        SQL = "'" & Trim(Rs!numserie) & "'," & Rs!numfactu & "," & Year(Rs!fecfactu) & "," & i & ","
        
        'dependiendo del colectivo del socio cogemos la cta contable cliente o socio del articulo
        SQL = SQL & DBSet(CtaVenta, "T")
        
        SQL = SQL & "," & DBSet(Rs!BaseImpo, "N") & ","
        
        If CCoste = "" Then
            SQL = SQL & ValorNulo
        Else
            SQL = SQL & DBSet(CCoste, "T")
        End If
        
        If vParamAplic.ContabilidadNueva Then
            SQL = SQL & "," & DBSet(Rs!fecfactu, "F") & "," & DBSet(Tipiva, "N") & "," & DBSet(PorcIva, "N") & "," & ValorNulo & "," & DBSet(Rs!CuotaIva, "N") & "," & ValorNulo
        End If
        
        Cad = Cad & "(" & SQL & ")" & ","
        
        i = i + 1
        Rs.MoveNext
    End If
    
    Rs.Close
    Set Rs = Nothing

    'Insertar en la contabilidad
    If Cad <> "" Then
        Cad = Mid(Cad, 1, Len(Cad) - 1) 'quitar la ult. coma
        '$$$
        If vParamAplic.ContabilidadNueva Then
            SQL = "INSERT INTO factcli_lineas (numserie,numfactu,anofactu,numlinea,codmacta,baseimpo,codccost,fecfactu,codigiva,porciva,porcrec,impoiva,imporec) "
        Else
            SQL = "INSERT INTO linfact (numserie,codfaccl,anofaccl,numlinea,codtbase,impbascl,codccost) "
        End If
        SQL = SQL & " VALUES " & Cad
        ConnContaTel.Execute SQL
    End If

EInLinea:
    If Err.Number <> 0 Then
        InsertarLinFact = False
        caderr = Err.Description
    Else
        InsertarLinFact = True
    End If
End Function




Private Function InsertarLinFactGas(cadTABLA As String, cadwhere As String, caderr As String, CCoste As String) As Boolean
'cadWHere: selecciona un registro de gascabfac
'codtipom=x and numfactu=y and fecfactu=z
Dim SQL As String
Dim Rs As ADODB.Recordset
Dim Cad As String, Aux As String
Dim i As Long
Dim totimp As Currency, ImpLinea As Currency
Dim CodIVA As String
Dim Iva As String
Dim vIva As Currency


    On Error GoTo EInLinea

    SQL = " SELECT letraser, numfactu, fecfactu, base, codiva, iva, total, porciva from " & cadTABLA & " where " & cadwhere
    Set Rs = New ADODB.Recordset
    Rs.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText

    Cad = ""
    i = 1
    totimp = 0
    If Not Rs.EOF Then
        'concatenamos linea para insertar en la tabla de conta.linfact
        SQL = ""
        SQL = "'" & Trim(Rs!letraser) & "'," & Rs!numfactu & "," & Year(Rs!fecfactu) & "," & i & ","
        
        'dependiendo del colectivo del socio cogemos la cta contable cliente o socio del articulo
        SQL = SQL & DBSet(vParamAplic.CtaVentasGas, "T")
        
        SQL = SQL & "," & DBSet(Rs!Base, "N") & ","
        
        If CCoste = "" Then
            SQL = SQL & ValorNulo
        Else
            SQL = SQL & DBSet(CCoste, "T")
        End If
        
        If vParamAplic.ContabilidadNueva Then
            SQL = SQL & "," & DBSet(Rs!fecfactu, "F") & "," & DBSet(Rs!CodIVA, "N") & "," & DBSet(Rs!PorcIva, "N") & ","
            SQL = SQL & ValorNulo & "," & DBSet(Rs!Iva, "N") & "," & ValorNulo
        End If
        
        Cad = Cad & "(" & SQL & ")" & ","
        
        i = i + 1
        Rs.MoveNext
    End If
    
    Rs.Close
    Set Rs = Nothing

    'Insertar en la contabilidad
    If Cad <> "" Then
        Cad = Mid(Cad, 1, Len(Cad) - 1) 'quitar la ult. coma
        If vParamAplic.ContabilidadNueva Then
            SQL = "INSERT INTO factcli_lineas (numserie,numfactu,anofactu,numlinea,codmacta,baseimpo,codccost,fecfactu,codigiva,porciva,porcrec,impoiva,imporec) "
        Else
            SQL = "INSERT INTO linfact (numserie,codfaccl,anofaccl,numlinea,codtbase,impbascl,codccost) "
        End If
        SQL = SQL & " VALUES " & Cad
        ConnContaGas.Execute SQL
    End If

EInLinea:
    If Err.Number <> 0 Then
        InsertarLinFactGas = False
        caderr = Err.Description
    Else
        InsertarLinFactGas = True
    End If
End Function



Private Function InsertarLinFactFac(cadTABLA As String, cadwhere As String, caderr As String, CCoste As String) As Boolean
'cadWHere: selecciona un registro de scafac
'codtipom=x and numfactu=y and fecfactu=z
Dim SQL As String
Dim Rs As ADODB.Recordset
Dim Cad As String, Aux As String
Dim i As Long
Dim totimp As Currency, ImpLinea As Currency
Dim CodIVA As String
Dim Iva As String
Dim vIva As Currency


    On Error GoTo EInLinea

    SQL = " SELECT letraser, numfactu, fecfactu, concefact.codmacta, concefact.codccost, sum(importe) from " & cadTABLA
    SQL = SQL & ", concefact where " & cadwhere
    SQL = SQL & " and concefact.codconce = " & cadTABLA & ".codconce"
    SQL = SQL & " GROUP BY 1,2,3,4,5 "
    
    Set Rs = New ADODB.Recordset
    Rs.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText

    Cad = ""
    i = 1
    totimp = 0
    While Not Rs.EOF
        'concatenamos linea para insertar en la tabla de conta.linfact
        SQL = ""
        SQL = "'" & Trim(Rs!letraser) & "'," & Rs!numfactu & "," & Year(Rs!fecfactu) & "," & i & ","
        
        'dependiendo del colectivo del socio cogemos la cta contable cliente o socio del articulo
        SQL = SQL & DBSet(Rs!Codmacta, "T")
        
        SQL = SQL & "," & DBSet(Rs.Fields(5).Value, "N") & ","
        
        If DBLet(Rs!CodCCost, "T") = "" Then
            SQL = SQL & ValorNulo
        Else
            '[Monica]05/07/2012: comprobamos aqui que si hay analitica y tiene centro de coste, la cuenta debe comenzar por
            '                    los digitos que indican la contabilidad
            Dim GrupoGto As String
            Dim GrupoVta As String
            Dim GrupoOrd As String
            
            GrupoGto = DevuelveDesdeBDNewFac("parametros", "grupogto", "", "", "T")
            GrupoVta = DevuelveDesdeBDNewFac("parametros", "grupovta", "", "", "T")
            GrupoOrd = DevuelveDesdeBDNewFac("parametros", "grupoord", "", "", "T")
            
            If vEmpresaFac.TieneAnalitica And (Mid(Trim(Rs!Codmacta), 1, 1) = GrupoGto Or Mid(Trim(Rs!Codmacta), 1, 1) = GrupoVta Or Mid(Trim(Rs!Codmacta), 1, 1) = GrupoOrd) Then
                SQL = SQL & DBSet(Rs!CodCCost, "T")
            Else
                SQL = SQL & ValorNulo
            End If
        End If
        
        Cad = Cad & "(" & SQL & ")" & ","
        
        i = i + 1
        Rs.MoveNext
        
    Wend
    
    Rs.Close
    Set Rs = Nothing

    'Insertar en la contabilidad
    If Cad <> "" Then
        Cad = Mid(Cad, 1, Len(Cad) - 1) 'quitar la ult. coma
        SQL = "INSERT INTO linfact (numserie,codfaccl,anofaccl,numlinea,codtbase,impbascl,codccost) "
        SQL = SQL & " VALUES " & Cad
        ConnContaFac.Execute SQL
    End If

EInLinea:
    If Err.Number <> 0 Then
        InsertarLinFactFac = False
        caderr = Err.Description
    Else
        InsertarLinFactFac = True
    End If
End Function


Private Function InsertarLinFactFacContaNueva(cadTABLA As String, cadwhere As String, caderr As String, CCoste As String) As Boolean
'cadWHere: selecciona un registro de scafac
'codtipom=x and numfactu=y and fecfactu=z
Dim SQL As String
Dim Rs As ADODB.Recordset
Dim Cad As String, Aux As String
Dim i As Long
Dim totimp As Currency, ImpLinea As Currency
Dim CodIVA As String
Dim Iva As String
Dim vIva As Currency

Dim NumeroIVA As Byte
Dim k As Integer
Dim HayQueAjustar As Boolean
Dim ImpImva As Currency
Dim ImpRec As Currency

    
    On Error GoTo EInLinea

    SQL = " SELECT letraser, numfactu, fecfactu, concefact.codmacta, concefact.codccost, tipoiva, sum(importe) from " & cadTABLA
    SQL = SQL & ", concefact where " & cadwhere
    SQL = SQL & " and concefact.codconce = " & cadTABLA & ".codconce"
    SQL = SQL & " GROUP BY 1,2,3,4,5,6 "
    
    Set Rs = New ADODB.Recordset
    Rs.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText

    Cad = ""
    i = 1
    totimp = 0
    While Not Rs.EOF
        'concatenamos linea para insertar en la tabla de conta.linfact
        SQL = ""
        SQL = "'" & Trim(Rs!letraser) & "'," & Rs!numfactu & "," & Year(Rs!fecfactu) & "," & i & ","
        
        'dependiendo del colectivo del socio cogemos la cta contable cliente o socio del articulo
        SQL = SQL & DBSet(Rs!Codmacta, "T")
        
        SQL = SQL & "," & DBSet(Rs.Fields(5).Value, "N") & ","
        
        
        'Vemos que tipo de IVA es en el vector de importes
        NumeroIVA = 127
        For k = 0 To 2
            If Rs!CodigIva = vTipoIva(k) Then
                NumeroIVA = k
                Exit For
            End If
        Next
        If NumeroIVA > 100 Then Err.Raise 513, "Error obteniendo IVA: " & Rs!CodigIva
        
        
        
        If DBLet(Rs!CodCCost, "T") = "" Then
            SQL = SQL & ValorNulo
        Else
            '[Monica]05/07/2012: comprobamos aqui que si hay analitica y tiene centro de coste, la cuenta debe comenzar por
            '                    los digitos que indican la contabilidad
            Dim GrupoGto As String
            Dim GrupoVta As String
            Dim GrupoOrd As String
            
            GrupoGto = DevuelveDesdeBDNewFac("parametros", "grupogto", "", "", "T")
            GrupoVta = DevuelveDesdeBDNewFac("parametros", "grupovta", "", "", "T")
            GrupoOrd = DevuelveDesdeBDNewFac("parametros", "grupoord", "", "", "T")
            
            If vEmpresaFac.TieneAnalitica And (Mid(Trim(Rs!Codmacta), 1, 1) = GrupoGto Or Mid(Trim(Rs!Codmacta), 1, 1) = GrupoVta Or Mid(Trim(Rs!Codmacta), 1, 1) = GrupoOrd) Then
                SQL = SQL & DBSet(Rs!CodCCost, "T")
            Else
                SQL = SQL & ValorNulo
            End If
        End If
        
        vBaseIva(NumeroIVA) = vBaseIva(NumeroIVA) - ImpLinea   'Para ajustar el importe y que no haya descuadre
        HayQueAjustar = False
        If vBaseIva(NumeroIVA) <> 0 Then
            'falta importe.
            'Puede ser que hayan mas lineas, o haya descuadre. Como esta ordenado por tipo de iva
            Rs.MoveNext
            If Rs.EOF Then
                'No hay mas lineas
                'Hay que ajustar SI o SI
                HayQueAjustar = True
            Else
                'Si que hay mas lineas.
                'Son del mismo tipo de IVA
                If Rs!TipoIva <> vTipoIva(0) Then
                    'NO es el mismo tipo de IVA
                    'Hay que ajustar
                    HayQueAjustar = True
                End If
            End If
            Rs.MovePrevious
        End If
        
        SQL = SQL & "," & vTipoIva(NumeroIVA) & "," & DBSet(vPorcIva(NumeroIVA), "N") & "," & DBSet(vPorcRec(NumeroIVA), "N", "S") & ","
        
        If HayQueAjustar Then
            Stop
        Else
        
        End If

        
        'Caluclo el importe de IVA y el de recargo de equivalencia
        ImpImva = vPorcIva(NumeroIVA) / 100
        ImpImva = Round2(ImpLinea * ImpImva, 2)
        If vPorcRec(NumeroIVA) = 0 Then
            ImpRec = 0
        Else
            ImpRec = vPorcRec(NumeroIVA) / 100
            ImpRec = Round2(ImpLinea * ImpRec, 2)
        End If
        vImpIva(NumeroIVA) = vImpIva(NumeroIVA) - ImpImva
        vImpRec(NumeroIVA) = vImpRec(NumeroIVA) - ImpRec
        
        
        ' baseimpo , impoiva, imporec
        SQL = SQL & DBSet(ImpLinea, "N") & "," & DBSet(ImpImva, "N") & "," & DBSet(ImpRec, "N", "S")
        
        Cad = Cad & "(" & SQL & ")" & ","
        
        i = i + 1
        Rs.MoveNext
        
    Wend
    
    Rs.Close
    Set Rs = Nothing

    'Insertar en la contabilidad
    If Cad <> "" Then
        Cad = Mid(Cad, 1, Len(Cad) - 1) 'quitar la ult. coma
        SQL = "INSERT INTO factcli_lineas (numserie,numfactu,anofactu,numlinea,codmacta,codccost,fecfactu,codigiva,porciva,porcrec,baseimpo,impoiva,imporec) "
        SQL = SQL & " VALUES " & Cad
        ConnContaFac.Execute SQL
    End If

EInLinea:
    If Err.Number <> 0 Then
        InsertarLinFactFacContaNueva = False
        caderr = Err.Description
    Else
        InsertarLinFactFacContaNueva = True
    End If
End Function







Private Function ActualizarCabFact(cadTABLA As String, cadwhere As String, caderr As String) As Boolean
'Poner la factura como contabilizada
Dim SQL As String

    On Error GoTo EActualizar
    
    SQL = "UPDATE " & cadTABLA & " SET intconta=1 "
    SQL = SQL & " WHERE " & cadwhere

    conn.Execute SQL
    
EActualizar:
    If Err.Number <> 0 Then
        ActualizarCabFact = False
        caderr = Err.Description
    Else
        ActualizarCabFact = True
    End If
End Function

Private Sub InsertarError(Cadena As String)
Dim SQL As String

    SQL = "insert into tmperrcomprob values ('" & Cadena & "')"
    conn.Execute SQL

End Sub


Public Function InsertarCabAsientoDia(Diario As String, Asiento As String, Fecha As String, Obs As String, caderr As String, bd As Byte) As Boolean
'Insertando en tabla conta.cabfact
Dim SQL As String
Dim Rs As ADODB.Recordset
Dim Cad As String
Dim Nulo2 As String
Dim Nulo3 As String

    On Error GoTo EInsertar
       
    If vParamAplic.ContabilidadNueva Then
        Cad = Format(Diario, "00") & ", " & DBSet(Fecha, "F") & "," & Format(Asiento, "000000") & ","
        Cad = Cad & DBSet(Obs, "T")
        Cad = Cad & "," & DBSet(Now, "FH") & "," & DBSet(vSesion.Login, "T") & ",'ARIAGROUTIL'"
        Cad = "(" & Cad & ")"

        'Insertar en la contabilidad
        SQL = "INSERT INTO hcabapu (numdiari, fechaent, numasien, obsdiari, feccreacion, usucreacion, desdeaplicacion) "
        SQL = SQL & " VALUES " & Cad
    Else
        Cad = Format(Diario, "00") & ", " & DBSet(Fecha, "F") & "," & Format(Asiento, "000000") & ","
        Cad = Cad & "''," & ValorNulo & "," & DBSet(Obs, "T")
        Cad = "(" & Cad & ")"
    
        'Insertar en la contabilidad
        SQL = "INSERT INTO cabapu (numdiari, fechaent, numasien, bloqactu, numaspre, obsdiari) "
        SQL = SQL & " VALUES " & Cad
    End If
    
    Select Case bd
        Case cConta
            ConnConta.Execute SQL
        Case cContaSeg
            ConnContaSeg.Execute SQL
        Case cContaGas
            ConnContaGas.Execute SQL
    End Select

EInsertar:
    If Err.Number <> 0 Then
        InsertarCabAsientoDia = False
        caderr = Err.Description
    Else
        InsertarCabAsientoDia = True
    End If
End Function


Public Function InsertarLinAsientoDia(Cad As String, caderr As String, bd As Byte) As Boolean
' el Tipo me indica desde donde viene la llamada
' tipo = 0 srecau.codmacta
' tipo = 1 scaalb.codmacta

Dim Rs As ADODB.Recordset
Dim Aux As String
Dim SQL As String
Dim i As Byte
Dim totimp As Currency, ImpLinea As Currency

    On Error GoTo EInLinea

    If vParamAplic.ContabilidadNueva Then
        SQL = "INSERT INTO hlinapu (numdiari, fechaent, numasien, linliapu, codmacta, numdocum, codconce, "
        SQL = SQL & " ampconce, timporteD, timporteH, codccost, ctacontr, idcontab, punteada) "
        SQL = SQL & " VALUES " & Cad
    Else
        SQL = "INSERT INTO linapu (numdiari, fechaent, numasien, linliapu, codmacta, numdocum, codconce, "
        SQL = SQL & " ampconce, timporteD, timporteH, codccost, ctacontr, idcontab, punteada) "
        SQL = SQL & " VALUES " & Cad
    End If
    
    Select Case bd
        Case cConta
            ConnConta.Execute SQL
        Case cContaSeg
            ConnContaSeg.Execute SQL
        Case cContaGas
            ConnContaGas.Execute SQL
    End Select

EInLinea:
    If Err.Number <> 0 Then
        InsertarLinAsientoDia = False
        caderr = Err.Description
    Else
        InsertarLinAsientoDia = True
    End If
End Function

Public Function ActualizarMovimientos(cadwhere As String, caderr As String) As Boolean
'Poner el movimiento como contabilizada
Dim SQL As String

    On Error GoTo EActualizar
    
    SQL = "UPDATE movim SET intconta=1 "
    SQL = SQL & " WHERE " & cadwhere

    conn.Execute SQL
    
EActualizar:
    If Err.Number <> 0 Then
        ActualizarMovimientos = False
        caderr = Err.Description
    Else
        ActualizarMovimientos = True
    End If
End Function

Public Sub FechasEjercicioConta(FIni As String, FFin As String)
Dim Rs As ADODB.Recordset

    On Error GoTo EFechas

    FIni = "Select fechaini,fechafin From parametros"
    Set Rs = New ADODB.Recordset
    Rs.Open FIni, ConnConta, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not Rs.EOF Then
        FIni = DBLet(Rs!FechaIni, "F")
        FFin = DBLet(Rs!FechaFin, "F")
    End If
    Rs.Close
    Set Rs = Nothing

EFechas:
    If Err.Number <> 0 Then Err.Clear
End Sub

Public Sub FechasEjercicioContaSeg(FIniSeg As String, FFinSeg As String)
Dim Rs As ADODB.Recordset

    On Error GoTo EFechas

    FIniSeg = "Select fechaini,fechafin From parametros"
    Set Rs = New ADODB.Recordset
    Rs.Open FIniSeg, ConnContaSeg, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not Rs.EOF Then
        FIniSeg = DBLet(Rs!FechaIni, "F")
        FFinSeg = DBLet(Rs!FechaFin, "F")
    End If
    Rs.Close
    Set Rs = Nothing

EFechas:
    If Err.Number <> 0 Then Err.Clear
End Sub

Public Sub FechasEjercicioContaTel(FIniTel As String, FFinTel As String)
Dim Rs As ADODB.Recordset

    On Error GoTo EFechas

    FIniTel = "Select fechaini,fechafin From parametros"
    Set Rs = New ADODB.Recordset
    Rs.Open FIniTel, ConnContaTel, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not Rs.EOF Then
        FIniTel = DBLet(Rs!FechaIni, "F")
        FFinTel = DBLet(Rs!FechaFin, "F")
    End If
    Rs.Close
    Set Rs = Nothing

EFechas:
    If Err.Number <> 0 Then Err.Clear
End Sub

Public Sub FechasEjercicioContaGas(FIniGas As String, FFinGas As String)
Dim Rs As ADODB.Recordset

    On Error GoTo EFechas

    FIniGas = "Select fechaini,fechafin From parametros"
    Set Rs = New ADODB.Recordset
    Rs.Open FIniGas, ConnContaGas, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not Rs.EOF Then
        FIniGas = DBLet(Rs!FechaIni, "F")
        FFinGas = DBLet(Rs!FechaFin, "F")
    End If
    Rs.Close
    Set Rs = Nothing

EFechas:
    If Err.Number <> 0 Then Err.Clear
End Sub

Public Sub FechasEjercicioContaFacSoc(FIniFacSoc As String, FFinFacSoc As String)
Dim Rs As ADODB.Recordset

    On Error GoTo EFechas

    FIniFacSoc = "Select fechaini,fechafin From parametros"
    Set Rs = New ADODB.Recordset
    Rs.Open FIniFacSoc, ConnContaFacSoc, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not Rs.EOF Then
        FIniFacSoc = DBLet(Rs!FechaIni, "F")
        FFinFacSoc = DBLet(Rs!FechaFin, "F")
    End If
    Rs.Close
    Set Rs = Nothing

EFechas:
    If Err.Number <> 0 Then Err.Clear
End Sub


Public Sub FechasEjercicioContaCV(FIniCV As String, FFinCV As String)
Dim Rs As ADODB.Recordset

    On Error GoTo EFechas

    FIniCV = "Select fechaini,fechafin From parametros"
    Set Rs = New ADODB.Recordset
    Rs.Open FIniCV, ConnContaCV, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not Rs.EOF Then
        FIniCV = DBLet(Rs!FechaIni, "F")
        FFinCV = DBLet(Rs!FechaFin, "F")
    End If
    Rs.Close
    Set Rs = Nothing

EFechas:
    If Err.Number <> 0 Then Err.Clear
End Sub


Public Sub FechasEjercicioContaCVV(FIniCVV As String, FFinCVV As String)
Dim Rs As ADODB.Recordset

    On Error GoTo EFechas

    FIniCVV = "Select fechaini,fechafin From parametros"
    Set Rs = New ADODB.Recordset
    Rs.Open FIniCVV, ConnContaCVV, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not Rs.EOF Then
        FIniCVV = DBLet(Rs!FechaIni, "F")
        FFinCVV = DBLet(Rs!FechaFin, "F")
    End If
    Rs.Close
    Set Rs = Nothing

EFechas:
    If Err.Number <> 0 Then Err.Clear
End Sub




Public Function CrearTMPAsiento() As Boolean
'Crea una temporal donde inserta la clave primaria de las
'facturas seleccionadas para facturar y trabaja siempre con ellas
Dim SQL As String
    
    On Error GoTo ECrear
    
    CrearTMPAsiento = False
    
    SQL = "CREATE TEMPORARY TABLE tmpasien ( "
    SQL = SQL & "fecalbar date NOT NULL default '0000-00-00',"
    SQL = SQL & "codturno tinyint(1) NOT NULL default '0',"
    SQL = SQL & "codmacta varchar(10) NOT NULL default ' ',"
    SQL = SQL & "importel decimal(10,2)  NOT NULL default '0.00')"
    conn.Execute SQL

    CrearTMPAsiento = True
    
ECrear:
     If Err.Number <> 0 Then
        CrearTMPAsiento = False
        'Borrar la tabla temporal
        SQL = " DROP TABLE IF EXISTS tmpasien;"
        conn.Execute SQL
    End If
End Function

' ### [Monica] 07/05/2007
Public Function InsertarEnTesoreriaNew(Fechamov As String, FecVenci As String, codavnic As String, anoejerc As Integer, Codmacta2 As String, Concepto As String, forpa As String, MenError As String) As Boolean
'Guarda datos de Tesoreria en tablas: ariges.svenci y en conta.scobros
Dim b As Boolean
Dim Rs As ADODB.Recordset
Dim Rsx As ADODB.Recordset
Dim SQL As String, text1csb As String, text2csb As String
Dim sql2 As String
Dim Sql3 As String
Dim Sql4 As String
Dim Rs3 As ADODB.Recordset
Dim Rs4 As ADODB.Recordset

Dim CadValues As String, cadValuesAux As String 'para insertar en svenci
Dim CadValues2 As String, CadValuesAux2 As String 'para insertar en conta.scobro
Dim FecVenci1 As Date
Dim ImpVenci As Single
Dim i As Byte
Dim CodmacBPr As String
Dim cadWHERE2 As String
Dim DigConta As String

Dim vvIban As String


    On Error GoTo EInsertarTesoreriaNew

    b = False
    InsertarEnTesoreriaNew = False
    CadValues = ""
    CadValues2 = ""

    SQL = "select * from movim where fechamov = " & DBSet(Fechamov, "F") & " and codavnic = " & DBSet(codavnic, "N")
    
    Set Rs = New ADODB.Recordset
    Rs.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    If Not Rs.EOF Then
    
        text1csb = "'Nro:" & Format(codavnic, "0000000") & " " & Format(Fechamov, "dd/mm/yy")
        text1csb = text1csb & " de " & DBSet(Rs!timporte, "N") & "'"
        text2csb = Concepto
        
              
        Sql4 = "select codmacta, codbanco, codsucur, digcontr, cuentaba, iban, nombrper nommacta,nomcalle dirdatos, "
        Sql4 = Sql4 & " poblacio despobla,codposta,provinci desprovi,nifperso nifdatos from avnic "
        Sql4 = Sql4 & " where codavnic = " & codavnic & " and anoejerc = " & DBSet(anoejerc, "N")
        
        Set Rs4 = New ADODB.Recordset
        Rs4.Open Sql4, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        If Not Rs4.EOF Then
            DigConta = DBLet(Rs4!digcontr, "T")
            If DBLet(Rs4!digcontr, "T") = "**" Then DigConta = "00"
        
            CadValuesAux2 = "(" & DBSet(Rs4!Codmacta, "T") & "," & DBSet(codavnic, "N") & ", " & DBSet(Fechamov, "F") & ", 1,"
            CadValues2 = CadValuesAux2 & DBSet(forpa, "N") & "," & DBSet(FecVenci, "F") & "," & DBSet(Rs!timporte, "N") & "," & ValorNulo & "," & ValorNulo
            CadValues2 = CadValues2 & "," & DBSet(Codmacta2, "T") & ","
            
            If vParamAplic.ContabilidadNueva Then
                vvIban = MiFormat(Rs4!Iban, "") & MiFormat(Rs4!codbanco, "0000") & MiFormat(Rs4!codsucur, "0000") & MiFormat(DigConta, "00") & MiFormat(Rs4!cuentaba, "0000000000")
            
                CadValues2 = CadValues2 & "," & text1csb & "," & DBSet(text2csb, "T") & ","
                CadValues2 = CadValues2 & DBSet(vvIban, "T", "S") & ", "
                CadValues2 = CadValues2 & DBSet(Rs4!Nommacta, "T", "S") & "," & DBSet(Rs4!dirdatos, "T", "S") & "," & DBSet(Rs4!desPobla, "T", "S") & ","
                CadValues2 = CadValues2 & DBSet(Rs4!Codposta, "T", "S") & "," & DBSet(Rs4!desProvi, "T", "S") & "," & DBSet(Rs4!nifdatos, "T", "S") & ",'ES')"
                
                SQL = "INSERT INTO pagos (numserie, codmacta, numfactu, fecfactu, numorden, codforpa, fecefect, impefect, ctabanc1,text1csb,text2csb, iban,"
                SQL = SQL & "nomprove, domprove, pobprove, cpprove, proprove, nifprove, codpais)"
            
            Else
                CadValues2 = CadValues2 & ValorNulo & "," & "0,0," & text1csb & "," & DBSet(text2csb, "T") & "," & DBSet(Rs4!codbanco, "N") & ", "
                CadValues2 = CadValues2 & DBSet(Rs4!codsucur, "N") & ", " & DBSet(DigConta, "T") & ", " & DBSet(Rs4!cuentaba, "T") & ", " & ValorNulo ' & ") "
                '[Monica]22/11/2013: tema iban
                If vEmpresa.HayNorma19_34Nueva = 1 Then
                   CadValues2 = CadValues2 & "," & DBSet(Rs4!Iban, "T", "S") & ")"
                Else
                   CadValues2 = CadValues2 & ")"
                End If
                
                'Insertamos en la tabla scobro de la CONTA
                SQL = "INSERT INTO spagop (ctaprove, numfactu, fecfactu, numorden, codforpa, fecefect, impefect,  "
                SQL = SQL & "fecultpa, imppagad, ctabanc1, ctabanc2, emitdocum, contdocu, text1csb, text2csb, entidad, "
                SQL = SQL & "oficina, cc, cuentaba, transfer" ' ) "
                
                '[Monica]22/11/2013: tema iban
                If vEmpresa.HayNorma19_34Nueva = 1 Then
                   SQL = SQL & ",iban)"
                Else
                   SQL = SQL & ")"
                End If
            End If
            
            SQL = SQL & " VALUES " & CadValues2
            ConnConta.Execute SQL
        End If

    End If

    b = True

EInsertarTesoreriaNew:
    If Err.Number <> 0 Then b = False
    InsertarEnTesoreriaNew = b
End Function



' ### [Monica] 07/05/2007
Public Function InsertarEnTesoreriaNew2(ByRef Rsx As ADODB.Recordset, FecVenci As String, forpa As String, CtaBan As String, MenError As String) As Boolean
'Guarda datos de Tesoreria en tablas: conta.scobros
Dim b As Boolean
Dim SQL As String, text33csb As String, text41csb As String
Dim Sql4 As String
Dim Rs4 As ADODB.Recordset

Dim CadValues As String, cadValuesAux As String 'para insertar en svenci
Dim CadValues2 As String, CadValuesAux2 As String 'para insertar en conta.scobro
Dim i As Byte
Dim DigConta As String
Dim CC As String
Dim vrefer As String

    On Error GoTo EInsertarTesoreriaNew2

    b = False
    InsertarEnTesoreriaNew2 = False
    CadValues = ""
    CadValues2 = ""

    If vParamAplic.ContabilidadNueva Then
        Sql4 = "select * "
    Else
        Sql4 = "select entidad, oficina, CC, cuentaba "
        '[Monica]22/11/2013: tema iban
        If vEmpresaSeg.HayNorma19_34Nueva = 1 Then
           Sql4 = Sql4 & ",iban"
        Else
           Sql4 = Sql4 & ""
        End If
    End If
    Sql4 = Sql4 & " from cuentas where codmacta = " & DBSet(Rsx!Codmacta, "T")
    Set Rs4 = New ADODB.Recordset
    
    Rs4.Open Sql4, ConnContaSeg, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    If Not Rs4.EOF Then
        text33csb = "'Nro:" & DBLet(Rsx!Codrefer, "T") & " " & Format(DBLet(Rsx!FechaEnv, "F"), "dd/mm/yy")
        text33csb = text33csb & " de " & DBSet(Rsx!imppoliz, "N") & "'"
        text41csb = "Linea:" & Format(Rsx!codlinea, "0000") & " Plan:" & Format(Rsx!CodiPlan, "0000")
        text41csb = text41csb & " Colectivo:" & Format(Rsx!Colectiv, "0000000")
              
        If Not vParamAplic.ContabilidadNueva Then
            CC = DBLet(Rs4!CC, "T")
            If DBLet(Rs4!CC, "T") = "**" Then CC = "00"
        End If
        
        vrefer = Mid(Rsx!Codrefer, 2, 8) 'Mid(Rsx!Codrefer, 1, 6) & Mid(Rsx!Codrefer, 8, 1)
    
        CadValuesAux2 = "('S', " & DBSet(vrefer, "N") & "," & DBSet(Rsx!FechaEnv, "F") & ", 1," & DBSet(Rsx!Codmacta, "T") & ","
        CadValues2 = CadValuesAux2 & DBSet(forpa, "N") & "," & DBSet(FecVenci, "F") & "," & DBSet(Rsx!imppoliz, "N") & "," & DBSet(Rsx!impinter, "N") & ","
        
        If Not vParamAplic.ContabilidadNueva Then
            CadValues2 = CadValues2 & DBSet(CtaBan, "T") & "," & DBSet(Rs4!entidad, "N") & "," & DBSet(Rs4!oficina, "N") & ","
            CadValues2 = CadValues2 & DBSet(CC, "T") & "," & DBSet(Rs4!cuentaba, "T") & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & ","
            CadValues2 = CadValues2 & text33csb & "," & DBSet(text41csb, "T") & ",1" ')"
            '[Monica]22/11/2013: tema iban
            If vEmpresaSeg.HayNorma19_34Nueva = 1 Then
               CadValues2 = CadValues2 & "," & DBSet(Rs4!Iban, "T", "S") & ")"
            Else
               CadValues2 = CadValues2 & ")"
            End If
        
            'Insertamos en la tabla scobro de la CONTA
            SQL = "INSERT INTO scobro (numserie, codfaccl, fecfaccl, numorden, codmacta, codforpa, fecvenci, impvenci, gastos,"
            SQL = SQL & "ctabanc1, codbanco, codsucur, digcontr, cuentaba, ctabanc2, fecultco, impcobro,  "
            SQL = SQL & " text33csb, text41csb, agente" ') "
            
            '[Monica]22/11/2013: tema iban
            If vEmpresaSeg.HayNorma19_34Nueva = 1 Then
               SQL = SQL & ",iban)"
            Else
               SQL = SQL & ")"
            End If
        Else
            CadValues2 = CadValues2 & DBSet(CtaBan, "T") & "," & ValorNulo & "," & ValorNulo & ","
            CadValues2 = CadValues2 & text33csb & "," & DBSet(text41csb, "T") & ",1," & DBSet(Rs4!Iban, "T", "S") & ","
            CadValues2 = CadValues2 & DBSet(Rs4!Nommacta, "T", "S") & "," & DBSet(Rs4!dirdatos, "T", "S") & "," & DBSet(Rs4!desPobla, "T", "S") & ","
            CadValues2 = CadValues2 & DBSet(Rs4!Codposta, "T", "S") & "," & DBSet(Rs4!desProvi, "T", "S") & "," & DBSet(Rs4!nifdatos, "T", "S") & ",'ES')"
        
            SQL = "INSERT INTO cobros (numserie, numfactu, fecfactu, numorden, codmacta, codforpa, fecvenci, impvenci, gastos, "
            SQL = SQL & "ctabanc1, fecultco, impcobro, "
            SQL = SQL & " text33csb, text41csb, agente, iban, "
            SQL = SQL & "nomclien, domclien, pobclien, cpclien, proclien, nifclien, codpais"
            SQL = SQL & ") "
        
        End If
        
        SQL = SQL & " VALUES " & CadValues2
        ConnContaSeg.Execute SQL

    End If

    b = True

EInsertarTesoreriaNew2:
    If Err.Number <> 0 Then b = False
    InsertarEnTesoreriaNew2 = b
End Function



' ### [Monica] 07/05/2007
Public Function InsertarEnTesoreriaNew3(ByRef Rsx As ADODB.Recordset, FecVenci As String, forpa As String, CtaBan As String, ByRef MenError As String) As Boolean
'Guarda datos de Tesoreria en tablas: conta.scobros
Dim b As Boolean
Dim SQL As String, text33csb As String, text41csb As String
Dim Sql4 As String
Dim Rs4 As ADODB.Recordset

Dim CadValues As String, cadValuesAux As String 'para insertar en svenci
Dim CadValues2 As String, CadValuesAux2 As String 'para insertar en conta.scobro
Dim i As Byte
Dim DigConta As String
Dim CC As String
Dim vrefer As String

    On Error GoTo EInsertarTesoreriaNew3

    b = False
    InsertarEnTesoreriaNew3 = False
    CadValues = ""
    CadValues2 = ""

    
    If vParamAplic.ContabilidadNueva Then
        Sql4 = "select * "
    Else
        Sql4 = "select entidad, oficina, CC, cuentaba "
        '[Monica]22/11/2013: tema iban
        If vEmpresaTel.HayNorma19_34Nueva = 1 Then
           Sql4 = Sql4 & ",iban "
        End If
    End If
    Sql4 = Sql4 & "from cuentas where codmacta = " & Rsx!Codmacta
    
    Set Rs4 = New ADODB.Recordset
    
    Rs4.Open Sql4, ConnContaTel, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    If Not Rs4.EOF Then
        text33csb = "'Factura:" & DBLet(Trim(Rsx!numserie), "T") & "-" & DBLet(Rsx!numfactu, "T") & " " & Format(DBLet(Rsx!fecfactu, "F"), "dd/mm/yy") & "'"
        text41csb = "de " & DBSet(Rsx!TotalFac, "N")
              
              
        If Not vParamAplic.ContabilidadNueva Then
            CC = DBLet(Rs4!CC, "T")
            If DBLet(Rs4!CC, "T") = "**" Then CC = "00"
        End If
    
        CadValuesAux2 = "(" & DBSet(Trim(Rsx!numserie), "T") & "," & DBSet(Rsx!numfactu, "N") & "," & DBSet(Rsx!fecfactu, "F") & ", 1," & DBSet(Rsx!Codmacta, "T") & ","
        CadValues2 = CadValuesAux2 & DBSet(forpa, "N") & "," & DBSet(FecVenci, "F") & "," & DBSet(Rsx!TotalFac, "N") & ","
        CadValues2 = CadValues2 & DBSet(CtaBan, "T") & ","
        
        If Not vParamAplic.ContabilidadNueva Then
            CadValues2 = CadValues2 & DBSet(Rs4!entidad, "N") & "," & DBSet(Rs4!oficina, "N") & ","
            CadValues2 = CadValues2 & DBSet(CC, "T") & "," & DBSet(Rs4!cuentaba, "T") & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & ","
            CadValues2 = CadValues2 & text33csb & "," & DBSet(text41csb, "T") & ",1" ')"
            '[Monica]22/11/2013: tema iban
            If vEmpresaTel.HayNorma19_34Nueva = 1 Then
                CadValues2 = CadValues2 & ", " & DBSet(Rs4!Iban, "T", "S") & ")"
            Else
                CadValues2 = CadValues2 & ")"
            End If
            
            'Insertamos en la tabla scobro de la CONTA
            SQL = "INSERT INTO scobro (numserie, codfaccl, fecfaccl, numorden, codmacta, codforpa, fecvenci, impvenci, "
            SQL = SQL & "ctabanc1, codbanco, codsucur, digcontr, cuentaba, ctabanc2, fecultco, impcobro, "
            SQL = SQL & " text33csb, text41csb, agente" ') "
            '[Monica]22/11/2013: tema iban
            If vEmpresaTel.HayNorma19_34Nueva = 1 Then
               SQL = SQL & ",iban)"
            Else
               SQL = SQL & ")"
            End If
        Else
        
            CadValues2 = CadValues2 & ValorNulo & "," & ValorNulo & ","
            CadValues2 = CadValues2 & text33csb & "," & DBSet(text41csb, "T") & ",1," & DBSet(Rs4!Iban, "T", "S") & ","
            
            CadValues2 = CadValues2 & DBSet(Rs4!Nommacta, "T", "S") & "," & DBSet(Rs4!dirdatos, "T", "S") & "," & DBSet(Rs4!desPobla, "T", "S") & ","
            CadValues2 = CadValues2 & DBSet(Rs4!Codposta, "T", "S") & "," & DBSet(Rs4!desProvi, "T", "S") & "," & DBSet(Rs4!nifdatos, "T", "S") & ",'ES')"
        
            SQL = "INSERT INTO cobros (numserie, numfactu, fecfactu, numorden, codmacta, codforpa, fecvenci, impvenci, "
            SQL = SQL & "ctabanc1, fecultco, impcobro, "
            SQL = SQL & " text33csb, text41csb, agente, iban, "
            SQL = SQL & "nomclien, domclien, pobclien, cpclien, proclien, nifclien, codpais"
            SQL = SQL & ") "
        
        End If
        SQL = SQL & " VALUES " & CadValues2
        ConnContaTel.Execute SQL

    End If

    b = True

EInsertarTesoreriaNew3:
    If Err.Number <> 0 Then
        b = False
        MenError = Err.Description
    End If
    InsertarEnTesoreriaNew3 = b
End Function


Public Function InsertarEnTesoreriaNew4(ByRef Rsx As ADODB.Recordset, FecVenci As String, tipo As String, CtaBan As String, ByRef MenError As String) As Boolean
'Guarda datos de Tesoreria en tablas: conta.scobros
Dim b As Boolean
Dim SQL As String, text33csb As String, text41csb As String
Dim Sql4 As String, text1csb As String, text2csb As String
Dim Rs4 As ADODB.Recordset

Dim CadValues As String, cadValuesAux As String 'para insertar en svenci
Dim CadValues2 As String, CadValuesAux2 As String 'para insertar en conta.scobro
Dim i As Byte
Dim DigConta As String
Dim CC As String
Dim vrefer As String
Dim numfactu As Long

    On Error GoTo EInsertarTesoreriaNew4

    b = False
    InsertarEnTesoreriaNew4 = False
    CadValues = ""
    CadValues2 = ""

    If tipo <= 1 Then ' facturas de venta
        If tipo = 1 Then
            numfactu = Trim(Right(Rsx!numfactu, 7))
        Else
            numfactu = DBLet(Rsx!numfactu)
        End If
        
        If Not vParamAplic.ContabilidadNueva Then
            Sql4 = "select entidad, oficina, CC, cuentaba"
            '[Monica]22/11/2013: tema iban
            If vEmpresaCVV.HayNorma19_34Nueva = 1 Then
               Sql4 = Sql4 & ",iban"
            End If
        Else
            Sql4 = "select * "
        End If
        Sql4 = Sql4 & " from cuentas where codmacta = " & Rsx!CodmactaSoc
        
        Set Rs4 = New ADODB.Recordset
        If tipo = 0 Then
            Rs4.Open Sql4, ConnContaCVV, adOpenForwardOnly, adLockPessimistic, adCmdText
        Else
            Rs4.Open Sql4, ConnContaCV, adOpenForwardOnly, adLockPessimistic, adCmdText
        End If
        If Not Rs4.EOF Then
            text33csb = "'Factura:" & DBLet(Trim(Rsx!letraser), "T") & "-" & DBLet(numfactu, "N") & " " & Format(DBLet(Rsx!fecfactu, "F"), "dd/mm/yy") & "'"
            text41csb = "de " & DBSet(Rsx!TotalFac, "N")
                  
            If Not vParamAplic.ContabilidadNueva Then
                CC = DBLet(Rs4!CC, "T")
                If DBLet(Rs4!CC, "T") = "**" Then CC = "00"
            End If
        
            CadValuesAux2 = "(" & DBSet(Trim(Rsx!letraser), "T") & "," & DBSet(numfactu, "N") & "," & DBSet(Rsx!fecfactu, "F") & ", 1," & DBSet(Rsx!CodmactaSoc, "T") & ","
            CadValues2 = CadValuesAux2 & DBSet(Rsx!CodForpa, "N") & "," & DBSet(FecVenci, "F") & "," & DBSet(Rsx!TotalFac, "N") & ","
            CadValues2 = CadValues2 & DBSet(CtaBan, "T")
            
            If Not vParamAplic.ContabilidadNueva Then
                CadValues2 = CadValues2 & "," & DBSet(Rs4!entidad, "N") & "," & DBSet(Rs4!oficina, "N") & ","
                CadValues2 = CadValues2 & DBSet(CC, "T") & "," & DBSet(Rs4!cuentaba, "T") & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & ","
                CadValues2 = CadValues2 & text33csb & "," & DBSet(text41csb, "T") & ",1" ')"
                
                '[Monica]22/11/2013: tema iban
                If vEmpresaCVV.HayNorma19_34Nueva = 1 Then
                   CadValues2 = CadValues2 & "," & DBSet(Rs4!Iban, "T", "S") & ")"
                Else
                   CadValues2 = CadValues2 & ")"
                End If
                
                
                'Insertamos en la tabla scobro de la CONTA
                SQL = "INSERT INTO scobro (numserie, codfaccl, fecfaccl, numorden, codmacta, codforpa, fecvenci, impvenci, "
                SQL = SQL & "ctabanc1, codbanco, codsucur, digcontr, cuentaba, ctabanc2, fecultco, impcobro, "
                SQL = SQL & " text33csb, text41csb, agente" ') "
                '[Monica]22/11/2013: tema iban
                If vEmpresaCVV.HayNorma19_34Nueva = 1 Then
                   SQL = SQL & ",iban) "
                Else
                   SQL = SQL & ")"
                End If
            Else
                CadValues2 = CadValues2 & ValorNulo & "," & ValorNulo & ","
                CadValues2 = CadValues2 & text33csb & "," & DBSet(text41csb, "T") & ",1," & DBSet(Rs4!Iban, "T", "S") & ","
                
                CadValues2 = CadValues2 & DBSet(Rs4!Nommacta, "T", "S") & "," & DBSet(Rs4!dirdatos, "T", "S") & "," & DBSet(Rs4!desPobla, "T", "S") & ","
                CadValues2 = CadValues2 & DBSet(Rs4!Codposta, "T", "S") & "," & DBSet(Rs4!desProvi, "T", "S") & "," & DBSet(Rs4!nifdatos, "T", "S") & ",'ES')"
            
                SQL = "INSERT INTO cobros (numserie, numfactu, fecfactu, numorden, codmacta, codforpa, fecvenci, impvenci, "
                SQL = SQL & "ctabanc1, fecultco, impcobro, "
                SQL = SQL & " text33csb, text41csb, agente, iban, "
                SQL = SQL & "nomclien, domclien, pobclien, cpclien, proclien, nifclien, codpais"
                SQL = SQL & ") "
            
            End If
            
            SQL = SQL & " VALUES " & CadValues2
            If tipo = 0 Then
                ConnContaCVV.Execute SQL
            Else
                ConnContaCV.Execute SQL
            End If
        End If

    Else ' factura de compras
        text1csb = "'Nro Factura:" & DBLet(Rsx!numfactu, "T") & " Fecha: " & Format(DBLet(Rsx!fecfactu, "F"), "dd/mm/yy") & "'"
        text2csb = ""
        
        Sql4 = "select entidad, oficina, CC, cuentaba from cuentas where codmacta = " & DBLet(Rsx!CodmactaSoc, "T")
        Set Rs4 = New ADODB.Recordset
        
        Rs4.Open Sql4, ConnContaCV, adOpenForwardOnly, adLockPessimistic, adCmdText
        
        If Not Rs4.EOF Then
            DigConta = DBLet(Rs4!CC, "T")
            If DBLet(Rs4!CC, "T") = "**" Then DigConta = "00"
        
            CadValuesAux2 = "(" & DBSet(Rsx!CodmactaSoc, "T") & "," & DBSet(Rsx!numfactu, "T") & ", " & DBSet(Rsx!fecfactu, "F") & ", 1,"
            CadValues2 = CadValuesAux2 & DBSet(Rsx!CodForpa, "N") & "," & DBSet(FecVenci, "F") & "," & DBSet(Rsx!TotalFac, "N") & "," & ValorNulo & "," & ValorNulo
            CadValues2 = CadValues2 & "," & DBSet(CtaBan, "T") & "," & ValorNulo & "," & "0,0," & text1csb & "," & DBSet(text2csb, "T") & "," & DBSet(Rs4!entidad, "N") & ", "
            CadValues2 = CadValues2 & DBSet(Rs4!oficina, "N") & ", " & DBSet(DigConta, "T") & ", " & DBSet(Rs4!cuentaba, "T") & ", " & ValorNulo '& ") "
            
            '[Monica]22/11/2013: tema iban
            If vEmpresaCVV.HayNorma19_34Nueva = 1 Then
               CadValues2 = CadValues2 & "," & DBSet(Rs4!Iban, "T", "S") & ")"
            Else
               CadValues2 = CadValues2 & ")"
            End If
        
            'Insertamos en la tabla scobro de la CONTA
            SQL = "INSERT INTO spagop (ctaprove, numfactu, fecfactu, numorden, codforpa, fecefect, impefect,  "
            SQL = SQL & "fecultpa, imppagad, ctabanc1, ctabanc2, emitdocum, contdocu, text1csb, text2csb, entidad, "
            SQL = SQL & "oficina, cc, cuentaba, transfer" ') "
            '[Monica]22/11/2013: tema iban
            If vEmpresaCVV.HayNorma19_34Nueva = 1 Then
               SQL = SQL & ",iban)"
            Else
               SQL = SQL & ")"
            End If
            SQL = SQL & " VALUES " & CadValues2
            ConnContaCV.Execute SQL
        End If

    End If

    b = True

EInsertarTesoreriaNew4:
    If Err.Number <> 0 Then
        b = False
        MenError = Err.Description
    End If
    InsertarEnTesoreriaNew4 = b
End Function






' ### [Monica] 07/05/2007
Public Function InsertarEnTesoreriaNewFac(ByRef Rsx As ADODB.Recordset, FecVenci As String, CtaBan As String, MenError As String) As Boolean
'Guarda datos de Tesoreria en tablas: conta.scobros
Dim b As Boolean
Dim SQL As String, text33csb As String, text41csb As String
Dim Sql4 As String
Dim Rs4 As ADODB.Recordset

Dim CadValues As String, cadValuesAux As String 'para insertar en svenci
Dim CadValues2 As String, CadValuesAux2 As String 'para insertar en conta.scobro
Dim i As Byte
Dim DigConta As String
Dim CC As String
Dim vrefer As String

    On Error GoTo EInsertarTesoreriaNewFac

    b = False
    InsertarEnTesoreriaNewFac = False
    CadValues = ""
    CadValues2 = ""

    Sql4 = "select entidad, oficina, CC, cuentaba "
    
    '[Monica]22/11/2013: tema iban
    If vEmpresaFac.HayNorma19_34Nueva = 1 Then
        Sql4 = Sql4 & ", iban "
    End If
    Sql4 = Sql4 & " from cuentas where codmacta = " & Rsx!ctaclien
    
    Set Rs4 = New ADODB.Recordset
    
    Rs4.Open Sql4, ConnContaFac, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    If Not Rs4.EOF Then
        text33csb = "'Factura:" & DBLet(Trim(Rsx!letraser), "T") & "-" & DBLet(Rsx!numfactu, "T") & " " & Format(DBLet(Rsx!fecfactu, "F"), "dd/mm/yy") & "'"
        text41csb = "de " & DBSet(Rsx!TotalFac, "N")
              
        CC = DBLet(Rs4!CC, "T")
        If DBLet(Rs4!CC, "T") = "**" Then CC = "00"
    
    
        CadValuesAux2 = "(" & DBSet(Trim(Rsx!letraser), "T") & "," & DBSet(Rsx!numfactu, "N") & "," & DBSet(Rsx!fecfactu, "F") & ", 1," & DBSet(Rsx!ctaclien, "T") & ","
        CadValues2 = CadValuesAux2 & DBSet(Rsx!CodForpa, "N") & "," & DBSet(FecVenci, "F") & "," & DBSet(Rsx!TotalFac, "N") & ","
        CadValues2 = CadValues2 & DBSet(CtaBan, "T") & "," & DBSet(Rs4!entidad, "N") & "," & DBSet(Rs4!oficina, "N") & ","
        CadValues2 = CadValues2 & DBSet(CC, "T") & "," & DBSet(Rs4!cuentaba, "T") & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & ","
        CadValues2 = CadValues2 & text33csb & "," & DBSet(text41csb, "T") & ",1" ')"
        '[Monica]22/11/2013: tema iban
        If vEmpresaFac.HayNorma19_34Nueva = 1 Then
           CadValues2 = CadValues2 & "," & DBSet(Rs4!Iban, "T", "S") & ")"
        Else
           CadValues2 = CadValues2 & ")"
        End If
        
        'Insertamos en la tabla scobro de la CONTA
        SQL = "INSERT INTO scobro (numserie, codfaccl, fecfaccl, numorden, codmacta, codforpa, fecvenci, impvenci, "
        SQL = SQL & "ctabanc1, codbanco, codsucur, digcontr, cuentaba, ctabanc2, fecultco, impcobro, "
        SQL = SQL & " text33csb, text41csb, agente" ') "
        '[Monica]22/11/2013: tema iban
        If vEmpresaFac.HayNorma19_34Nueva = 1 Then
           SQL = SQL & ",iban)"
        Else
           SQL = SQL & ")"
        End If
        
        SQL = SQL & " VALUES " & CadValues2
        ConnContaFac.Execute SQL

    End If

    b = True

EInsertarTesoreriaNewFac:
    If Err.Number <> 0 Then b = False
    InsertarEnTesoreriaNewFac = b
End Function




Public Function ComprobarFormadePago(cadCC As String) As Boolean
Dim SQL As String
Dim sql2 As String
Dim Rs As ADODB.Recordset
Dim b As Boolean

    ComprobarFormadePago = False
    If vParamAplic.ContabilidadNueva Then
        SQL = DevuelveDesdeBDNew(cContaFacSoc, "formapago", "codforpa", "codforpa", cadCC, "N")
    Else
        SQL = DevuelveDesdeBDNew(cContaFacSoc, "sforpa", "codforpa", "codforpa", cadCC, "N")
    End If
    If SQL = "" Then
        b = False
        sql2 = "No existe la forma de Pago: " & cadCC
        InsertarError sql2
    Else
        b = True
    End If
    ComprobarFormadePago = b
    
End Function


'----------------------------------------------------------------------
' FACTURAS SOCIOS
'----------------------------------------------------------------------

Public Function PasarFacturaSoc(cadwhere As String, FecVenci As String, CtaBan As String, forpa As String) As Boolean
'Insertar en tablas cabecera/lineas de la Contabilidad la factura socios
' ariagroutil.factsocio  --> conta.cabfactprov
'                        --> conta.linfactprov
'Actualizar la tabla ariagroutil.factsocio.inconta=1 para indicar que ya esta contabilizada
Dim b As Boolean
Dim cadMen As String
Dim SQL As String
Dim Mc As CContadorContab


    On Error GoTo EContab

    ConnContaFacSoc.BeginTrans
    conn.BeginTrans
        
    
    Set Mc = New CContadorContab
    
    '---- Insertar en la conta Cabecera Factura
    b = InsertarCabFactSoc(cadwhere, cadMen, Mc, forpa)
    cadMen = "Insertando Cab. Factura: " & cadMen
    
    ' insertar en tesoreria
    If b Then
        b = InsertarEnTesoreriaFacSoc(vEmpresaFacSoc.FechaFin, FecVenci, cadwhere, CtaBan, forpa, cadMen)
        cadMen = "Insertando en Tesoreria: " & cadMen
    End If
    
    
    If b Then
        CCoste = ""
        '---- Insertar lineas de Factura en la Conta
        If vParamAplic.ContabilidadNueva Then
            b = InsertarLinFactSocContaNueva("factsocio", cadwhere, cadMen, Mc.Contador)
        Else
            b = InsertarLinFactSoc("factsocio", cadwhere, cadMen, Mc.Contador)
        End If
        cadMen = "Insertando Lin. Factura: " & cadMen

        If b Then
            '---- Poner intconta=1 en ariges.scafac
            b = ActualizarCabFact("factsocio", cadwhere, cadMen)
            cadMen = "Actualizando Factura: " & cadMen
        End If
    End If
    
'    If Not b Then
'        SQL = "Insert into tmpErrFac(codprove,numfactu,fecfactu,error) "
'        SQL = SQL & " Select *," & DBSet(Mid(cadMen, 1, 200), "T") & " as error From tmpFactu "
'        SQL = SQL & " WHERE " & Replace(cadWhere, "scafpc", "tmpFactu")
'        Conn.Execute SQL
'    End If
    
EContab:
    If Err.Number <> 0 Then
        b = False
        MuestraError Err.Number, "Contabilizando Factura", Err.Description
    End If
    If b Then
        ConnContaFacSoc.CommitTrans
        conn.CommitTrans
        PasarFacturaSoc = True
    Else
        ConnContaFacSoc.RollbackTrans
        conn.RollbackTrans
        PasarFacturaSoc = False
        If Not b Then
            SQL = "Insert into tmperrfac(codtipom, numfactu,fecfactu,error) "
            SQL = SQL & " select 1,numfactu, fecfactu, " & DBSet(cadMen, "T") & " from factsocio where " & cadwhere
            conn.Execute SQL
        End If
    End If
End Function


Private Function InsertarCabFactSoc(cadwhere As String, caderr As String, ByRef Mc As CContadorContab, forpa As String) As Boolean
'Insertando en tabla conta.cabfact
'(OUT) AnyoFacPr: aqui devolvemos el año de fecha recepcion para insertarlo en las lineas de factura de proveedor de la conta
Dim SQL As String
Dim Rs As ADODB.Recordset
Dim Cad As String
Dim Nulo2 As String
Dim Nulo3 As String
Dim Intracom As Integer
Dim BaseImp As Currency
Dim TotalFac As Currency

Dim SqlDatos As String
Dim RsDatos As ADODB.Recordset
Dim TipoOpera As String
Dim Aux As String
Dim sql2 As String

Dim CadenaInsertFaclin2 As String


    On Error GoTo EInsertar
       
   
    SQL = " SELECT fecfactu,year(fecfactu) as anofacpr,codmacta, numfactu, "
    SQL = SQL & "baseimpo,porciva,cuotaiva,basereten,porcreten,impreten,"
    SQL = SQL & "totalfac,tipoiva "
    SQL = SQL & " FROM " & "factsocio "
    SQL = SQL & " WHERE " & cadwhere
    
    Set Rs = New ADODB.Recordset
    Rs.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    Cad = ""
    If Not Rs.EOF Then
                                                                    '16/01/2012: antes era: CDate(vEmpresaFacSoc.FechaFin) - 365
        If Mc.ConseguirContador("1", (Rs!fecfactu <= CDate(vEmpresaFacSoc.FechaFin)), True, cContaFacSoc) = 0 Then
            'guardamos estos valores para utilizarlos cuando insertemos las lineas de la factura
'            DtoPPago = RS!DtoPPago
'            DtoGnral = RS!DtoGnral
            BaseImp = DBLet(Rs!BaseImpo, "N")
            TotalFac = BaseImp + DBLet(Rs!CuotaIva, "N")
'            AnyoFacPr = RS!anofacpr
            
            SQL = ""
            If vParamAplic.ContabilidadNueva Then SQL = SQL & DBSet(SerieFraPro, "T") & ","
            
            SQL = SQL & Mc.Contador & "," & DBSet(Rs!fecfactu, "F") & "," & DBLet(Rs!anofacpr, "N") & "," & DBSet(Rs!fecfactu, "F") & "," & DBSet(Rs!numfactu, "T") & "," & DBSet(Rs!Codmacta, "T") & "," & ValorNulo & ","
            
            
            If Not vParamAplic.ContabilidadNueva Then
            
                SQL = SQL & DBSet(Rs!BaseImpo, "N") & "," & ValorNulo & "," & ValorNulo & ","
                SQL = SQL & DBSet(Rs!PorcIva, "N") & "," & ValorNulo & "," & ValorNulo & ","
                SQL = SQL & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & DBSet(Rs!CuotaIva, "N") & "," & ValorNulo & "," & ValorNulo & ","
                SQL = SQL & ValorNulo & "," & ValorNulo & "," & ValorNulo & ","
                SQL = SQL & DBSet(TotalFac, "N") & "," & DBSet(Rs!TipoIva, "N") & "," & ValorNulo & "," & ValorNulo & ",0,"
                If DBLet(Rs!ImpReten, "N") <> 0 Then
                    SQL = SQL & DBSet(Rs!PorcReten, "N") & "," & DBSet(Rs!ImpReten, "N") & "," & DBSet(vParamAplic.CtaRetenFacSoc, "T") & ","
                Else
                    SQL = SQL & ValorNulo & "," & ValorNulo & "," & ValorNulo & ","
                End If
                SQL = SQL & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & DBSet(Rs!fecfactu, "F") & ",0"
                
                Cad = Cad & "(" & SQL & ")"
                
                'Insertar en la contabilidad
                SQL = "INSERT INTO cabfactprov (numregis,fecfacpr,anofacpr,fecrecpr,numfacpr,codmacta,confacpr,ba1facpr,ba2facpr,ba3facpr,"
                SQL = SQL & "pi1facpr,pi2facpr,pi3facpr,pr1facpr,pr2facpr,pr3facpr,ti1facpr,ti2facpr,ti3facpr,tr1facpr,tr2facpr,tr3facpr,"
                SQL = SQL & "totfacpr,tp1facpr,tp2facpr,tp3facpr,extranje,retfacpr,trefacpr,cuereten,numdiari,fechaent,numasien,fecliqpr,nodeducible) "
                SQL = SQL & " VALUES " & Cad
                ConnContaFacSoc.Execute SQL
                
            Else
                
                SqlDatos = "select * from cuentas where codmacta = " & DBSet(Rs!Codmacta, "T")
                Set RsDatos = New ADODB.Recordset
                RsDatos.Open SqlDatos, ConnContaGas, adOpenForwardOnly, adLockPessimistic, adCmdText
                If Not RsDatos.EOF Then
                
                    SQL = SQL & DBSet(RsDatos!Nommacta, "T", "S") & "," & DBSet(RsDatos!dirdatos, "T", "S") & "," & DBSet(RsDatos!desPobla, "T", "S") & ","
                    SQL = SQL & DBSet(RsDatos!Codposta, "T", "S") & "," & DBSet(RsDatos!desProvi, "T", "S") & "," & DBSet(RsDatos!nifdatos, "T", "S") & ",'ES',"
            
                Else
                
                    SQL = SQL & ValorNulo & "," & ValorNulo & "," & ValorNulo & ","
                    SQL = SQL & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & ","
                 
                End If
                Set RsDatos = Nothing
                
                SQL = SQL & DBSet(forpa, "N") & ","
                
                '$$$
                TipoOpera = 0
                
                Aux = "0"
                'codopera,codconce340,codintra
                SQL = SQL & TipoOpera & "," & DBSet(Aux, "T") & "," & ValorNulo & ","
                
                '[Monica]10/11/2016: en totalfac llevabamos base + impiva pq antes retencion estaba en lineas
                '                    en la nueva conta está en la cabecera
                TotalFac = DBLet(Rs!TotalFac, "N")
                
                'para las lineas
                'factpro_totales(numserie,numregis,fecharec,anofactu,numlinea,baseimpo,codigiva,porciva,porcrec,impoiva,imporec)
                'IVA 1, siempre existe
                Aux = "'" & SerieFraPro & "'," & Mc.Contador & "," & DBSet(Rs!fecfactu, "F") & "," & Rs!anofacpr & ","
                
                sql2 = Aux & "1," & DBSet(Rs!BaseImpo, "N") & "," & DBSet(Rs!TipoIva, "N") & "," & DBSet(Rs!PorcIva, "N") & ","
                sql2 = sql2 & ValorNulo & "," & DBSet(Rs!CuotaIva, "N") & "," & ValorNulo
                CadenaInsertFaclin2 = CadenaInsertFaclin2 & "(" & sql2 & ")"
                    
                'Los totales
                'totbases,totbasesret,totivas,totrecargo,totfacpr,
                SQL = SQL & DBSet(Rs!BaseImpo, "N") & "," & DBSet(Rs!BaseReten, "N") & ","
                'totivas
                SQL = SQL & DBSet(Rs!CuotaIva, "N") & "," & DBSet(TotalFac, "N") & ","
                If DBLet(Rs!PorcReten, "N") <> 0 Then
                    'porcreten,impreten,
                    SQL = SQL & DBSet(Rs!PorcReten, "N") & "," & DBSet(Rs!ImpReten, "N") & "," & DBSet(vParamAplic.CtaRetenFacSoc, "T") & ",2"
                Else
                    SQL = SQL & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo
                End If
                
                Cad = "(" & SQL & ")"
            
                'Insertar en la contabilidad
                SQL = "INSERT INTO factpro(numserie,numregis,fecfactu,anofactu,fecharec,fecliqpr,numfactu,codmacta,observa,nommacta,"
                SQL = SQL & "dirdatos,codpobla,despobla,desprovi,nifdatos,codpais,codforpa,codopera,codconce340,codintra,"
                SQL = SQL & "totbases,totbasesret,totivas,totfacpr,retfacpr,trefacpr,cuereten,tiporeten)"
                SQL = SQL & " VALUES " & Cad
                ConnConta.Execute SQL
            
                'Las  lineas de IVA
                SQL = "INSERT INTO factpro_totales(numserie,numregis,fecharec,anofactu,numlinea,baseimpo,codigiva,porciva,porcrec,impoiva,imporec)"
                SQL = SQL & " VALUES " & CadenaInsertFaclin2
                ConnConta.Execute SQL
            
            End If
'            'añadido como david para saber que numero de registro corresponde a cada factura
'            'Para saber el numreo de registro que le asigna a la factrua
'            SQL = "INSERT INTO tmpinformes (codusu,codigo1,nombre1,nombre2,importe1) VALUES (" & vUsu.Codigo & "," & Mc.Contador
'            SQL = SQL & ",'" & DevNombreSQL(RS!numfactu) & " @ " & Format(RS!FecFactu, "dd/mm/yyyy") & "','" & DevNombreSQL(RS!nomprove) & "'," & RS!codProve & ")"
'            conn.Execute SQL
        End If
    End If
    Rs.Close
    Set Rs = Nothing
    
EInsertar:
    If Err.Number <> 0 Then
        InsertarCabFactSoc = False
        caderr = Err.Description
    Else
        InsertarCabFactSoc = True
    End If
End Function


' ### [Monica] 07/05/2007
Public Function InsertarEnTesoreriaFacSoc(FechaFin As String, FecVenci As String, vWhere As String, CtaBanco As String, forpa As String, MenError As String) As Boolean
'Guarda datos de Tesoreria en tablas: ariges.svenci y en conta.scobros
Dim b As Boolean
Dim Rs As ADODB.Recordset
Dim Rsx As ADODB.Recordset
Dim SQL As String, text1csb As String, text2csb As String
Dim sql2 As String
Dim Sql3 As String
Dim Sql4 As String
Dim Rs3 As ADODB.Recordset
Dim Rs4 As ADODB.Recordset

Dim CadValues As String, cadValuesAux As String 'para insertar en svenci
Dim CadValues2 As String, CadValuesAux2 As String 'para insertar en conta.scobro
Dim FecVenci1 As Date
Dim ImpVenci As Single
Dim i As Byte
Dim CodmacBPr As String
Dim cadWHERE2 As String
Dim DigConta As String
Dim Variedad As String

    On Error GoTo EInsertarTesoreriaNew

    b = False
    InsertarEnTesoreriaFacSoc = False
    CadValues = ""
    CadValues2 = ""

    SQL = "select * from factsocio where " & vWhere
    
    Set Rs = New ADODB.Recordset
    Rs.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    If Not Rs.EOF Then
    
        Variedad = ""
        Variedad = DevuelveDesdeBDNew(cPTours, "variedad", "nomvarie", "codvarie", Rs!codvarie, "N")
    
        text1csb = "'Nro Factura:" & Format(DBLet(Rs!numfactu, "N"), "0000000") & " Fecha: " & Format(DBLet(Rs!fecfactu, "F"), "dd/mm/yy") & "'"
        text2csb = "Variedad: " & Variedad
        
        If Not vParamAplic.ContabilidadNueva Then
            '[Monica]22/11/2013: tema iban
            Sql4 = "select entidad, oficina, CC, cuentaba "
            
            '[Monica]22/11/2013: tema iban
            If vEmpresaFacSoc.HayNorma19_34Nueva = 1 Then
                Sql4 = Sql4 & ",iban"
            End If
        Else
            Sql4 = "select * "
        End If
        Sql4 = Sql4 & " from cuentas where codmacta = " & DBLet(Rs!Codmacta, "T")
        
        Set Rs4 = New ADODB.Recordset
        
        Rs4.Open Sql4, ConnContaFacSoc, adOpenForwardOnly, adLockPessimistic, adCmdText
        
        If Not Rs4.EOF Then
            If Not vParamAplic.ContabilidadNueva Then
                DigConta = DBLet(Rs4!CC, "T")
                If DBLet(Rs4!CC, "T") = "**" Then DigConta = "00"
            End If
        
            CadValuesAux2 = "(" & DBSet(Rs!Codmacta, "T") & "," & DBSet(Rs!numfactu, "N") & ", " & DBSet(Rs!fecfactu, "F") & ", 1,"
            CadValues2 = CadValuesAux2 & DBSet(forpa, "N") & "," & DBSet(FecVenci, "F") & "," & DBSet(Rs!TotalFac, "N") & "," & ValorNulo & "," & ValorNulo
            CadValues2 = CadValues2 & "," & DBSet(CtaBanco, "T")
            
            If vParamAplic.ContabilidadNueva Then
                CadValues2 = CadValues2 & "," & text1csb & "," & DBSet(text2csb, "T") & ","
                CadValues2 = CadValues2 & DBSet(Rs4!Iban, "T", "S") & ", "
                CadValues2 = CadValues2 & DBSet(Rs4!Nommacta, "T", "S") & "," & DBSet(Rs4!dirdatos, "T", "S") & "," & DBSet(Rs4!desPobla, "T", "S") & ","
                CadValues2 = CadValues2 & DBSet(Rs4!Codposta, "T", "S") & "," & DBSet(Rs4!desProvi, "T", "S") & "," & DBSet(Rs4!nifdatos, "T", "S") & ",'ES')"
                
                SQL = "INSERT INTO pagos (numserie, codmacta, numfactu, fecfactu, numorden, codforpa, fecefect, impefect, ctabanc1,text1csb,text2csb, iban,"
                SQL = SQL & "nomprove, domprove, pobprove, cpprove, proprove, nifprove, codpais)"
                
            Else
                CadValues2 = CadValues2 & "," & ValorNulo & "," & "0,0," & text1csb & "," & DBSet(text2csb, "T") & ","
                CadValues2 = CadValues2 & DBSet(Rs4!entidad, "N") & ", " & DBSet(Rs4!oficina, "N") & ", " & DBSet(DigConta, "T") & ", " & DBSet(Rs4!cuentaba, "T") & ", " & ValorNulo '& ") "
            
                '[Monica]22/11/2013: tema iban
                If vEmpresaFacSoc.HayNorma19_34Nueva = 1 Then
                   CadValues2 = CadValues2 & "," & DBSet(Rs4!Iban, "T", "S") & ")"
                Else
                   CadValues2 = CadValues2 & ")"
                End If
            
            
                'Insertamos en la tabla scobro de la CONTA
                SQL = "INSERT INTO spagop (ctaprove, numfactu, fecfactu, numorden, codforpa, fecefect, impefect,  "
                SQL = SQL & "fecultpa, imppagad, ctabanc1, ctabanc2, emitdocum, contdocu, text1csb, text2csb, entidad, "
                SQL = SQL & "oficina, cc, cuentaba, transfer" ') "
                
                '[Monica]22/11/2013: tema iban
                If vEmpresaFacSoc.HayNorma19_34Nueva = 1 Then
                   SQL = SQL & ", iban)"
                Else
                   SQL = SQL & ")"
                End If
            
            End If
        
        
        End If
        
        
        SQL = SQL & " VALUES " & CadValues2
        ConnContaFacSoc.Execute SQL

    End If

    b = True

EInsertarTesoreriaNew:
    If Err.Number <> 0 Then b = False
    InsertarEnTesoreriaFacSoc = b
End Function

Private Function InsertarLinFactSoc(cadTABLA As String, cadwhere As String, caderr As String, Optional numRegis As Long) As Boolean
'cadWHere: selecciona un registro de scafac
'codtipom=x and numfactu=y and fecfactu=z
Dim SQL As String
Dim SQLaux As String
Dim sql2 As String
Dim Rs As ADODB.Recordset
Dim Cad As String, Aux As String
Dim i As Byte
Dim totimp As Currency, ImpLinea As Currency

    On Error GoTo EInLinea

    SQL = " SELECT factsocio.codvarie, variedad.codmacta, factsocio.baseimpo as importe, "
    SQL = SQL & " factsocio.porcreten, factsocio.impreten, factsocio.basereten, factsocio.fecfactu, "
    SQL = SQL & " factsocio.codmacta as ctasocio "
    SQL = SQL & " FROM (factsocio  "
    SQL = SQL & " inner join variedad on factsocio.codvarie=variedad.codvarie) "
    SQL = SQL & " WHERE " & cadwhere
    
    Set Rs = New ADODB.Recordset
    Rs.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText

    Cad = ""
    i = 1
    totimp = 0
    
    SQLaux = Cad
    'calculamos la Base Imp del total del importe para cada cta cble ventas
    '---- Laura: 10/10/2006
    ImpLinea = DBLet(Rs!Importe, "N")
    '----
    totimp = totimp + ImpLinea
    
    'concatenamos linea para insertar en la tabla de conta.linfact
    ' linea de base sobre la cta contable de la variedad
    SQL = ""
    SQL = numRegis & "," & Year(Rs!fecfactu) & "," & i & ","
    SQL = SQL & DBSet(Rs!Codmacta, "T")
    SQL = SQL & "," & DBSet(ImpLinea, "N") & ","
    
    If CCoste = "" Then
        SQL = SQL & ValorNulo
    Else
        SQL = SQL & DBSet(CCoste, "T")
    End If
    
    Cad = Cad & "(" & SQL & ")"
    
    
    If DBLet(Rs!ImpReten, "N") <> 0 Then
        'linea de base del importe de retencion en positivo sobre la cuenta del socio
        i = i + 1
        SQL = ""
        SQL = numRegis & "," & Year(Rs!fecfactu) & "," & i & ","
        SQL = SQL & DBSet(Rs!CtaSocio, "T")
        SQL = SQL & "," & DBSet(Rs!ImpReten, "N") & ","
        
        If CCoste = "" Then
            SQL = SQL & ValorNulo
        Else
            SQL = SQL & DBSet(CCoste, "T")
        End If
        
        Cad = Cad & ",(" & SQL & ")"
        
        'linea de base del importe de retencion en negativo sobre la cuenta de retencion de parametros
        i = i + 1
        ImpLinea = DBLet(Rs!ImpReten, "N") * (-1)
        SQL = ""
        SQL = numRegis & "," & Year(Rs!fecfactu) & "," & i & ","
        SQL = SQL & DBSet(vParamAplic.CtaRetenFacSoc, "T")
        SQL = SQL & "," & DBSet(ImpLinea, "N") & ","
        
        If CCoste = "" Then
            SQL = SQL & ValorNulo
        Else
            SQL = SQL & DBSet(CCoste, "T")
        End If
        
        Cad = Cad & ",(" & SQL & ")"
    End If
    
    'Insertar en la contabilidad
    If Cad <> "" Then
        SQL = "INSERT INTO linfactprov (numregis,anofacpr,numlinea,codtbase,impbaspr,codccost) "
        SQL = SQL & " VALUES " & Cad
        ConnContaFacSoc.Execute SQL
    End If

EInLinea:
    If Err.Number <> 0 Then
        InsertarLinFactSoc = False
        caderr = Err.Description
    Else
        InsertarLinFactSoc = True
    End If
End Function


Private Function InsertarLinFactSocContaNueva(cadTABLA As String, cadwhere As String, caderr As String, Optional numRegis As Long) As Boolean
'cadWHere: selecciona un registro de scafac
'codtipom=x and numfactu=y and fecfactu=z
Dim SQL As String
Dim SQLaux As String
Dim sql2 As String
Dim Rs As ADODB.Recordset
Dim Cad As String, Aux As String
Dim i As Byte
Dim totimp As Currency, ImpLinea As Currency

    On Error GoTo EInLinea

    SQL = " SELECT factsocio.codvarie, variedad.codmacta, factsocio.baseimpo as importe, "
    SQL = SQL & " factsocio.porcreten, factsocio.impreten, factsocio.basereten, factsocio.fecfactu, "
    SQL = SQL & " factsocio.codmacta as ctasocio, factsocio.tipoiva, factsocio.porciva, factsocio.cuotaiva "
    SQL = SQL & " FROM (factsocio  "
    SQL = SQL & " inner join variedad on factsocio.codvarie=variedad.codvarie) "
    SQL = SQL & " WHERE " & cadwhere
    
    Set Rs = New ADODB.Recordset
    Rs.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText

    Cad = ""
    i = 1
    totimp = 0
    
    SQLaux = Cad
    'calculamos la Base Imp del total del importe para cada cta cble ventas
    '---- Laura: 10/10/2006
    ImpLinea = DBLet(Rs!Importe, "N")
    '----
    totimp = totimp + ImpLinea
    
    'concatenamos linea para insertar en la tabla de conta.linfact
    ' linea de base sobre la cta contable de la variedad
    SQL = DBSet(SerieFraPro, "T") & "," & numRegis & "," & DBSet(Rs!fecfactu, "F") & "," & Year(Rs!fecfactu) & "," & i & ","
    SQL = SQL & DBSet(Rs!Codmacta, "T") & ","
    
    If CCoste = "" Then
        SQL = SQL & ValorNulo
    Else
        SQL = SQL & DBSet(CCoste, "T")
    End If
    SQL = SQL & "," & DBSet(Rs!TipoIva, "N") & "," & DBSet(Rs!PorcIva, "N") & "," & DBSet(Rs!PorcReten, "N") & "," & DBSet(ImpLinea, "N") & ","
    SQL = SQL & "," & DBSet(Rs!CuotaIva, "N") & "," & DBSet(Rs!ImpReten, "N") & ",1"
    
    
    Cad = Cad & "(" & SQL & ")"
    
    
    'Insertar en la contabilidad
    If Cad <> "" Then
        SQL = "INSERT INTO factpro_lineas (numserie,numregis,fecharec,anofactu,numlinea,codmacta,codccost,codigiva,porciva,porcrec,"
        SQL = SQL & "baseimpo,impoiva,imporec,aplicret) "
        SQL = SQL & " VALUES " & Cad
        ConnContaFacSoc.Execute SQL
    End If

EInLinea:
    If Err.Number <> 0 Then
        InsertarLinFactSocContaNueva = False
        caderr = Err.Description
    Else
        InsertarLinFactSocContaNueva = True
    End If
End Function




Private Sub InsertarTMPErrFac(MenError As String, cadwhere As String)
Dim SQL As String

    On Error Resume Next
    SQL = "Insert into tmpErrFac(codprove,numfactu,fecfactu,error) "
    SQL = SQL & " Select *," & DBSet(Mid(MenError, 1, 200), "T") & " as error From tmpFactu "
    SQL = SQL & " WHERE " & Replace(cadwhere, "factsocio", "tmpFactu")
    conn.Execute SQL
    
    If Err.Number <> 0 Then Err.Clear
End Sub


Public Function ComprobarCCoste() As Boolean
Dim SQL As String
Dim Rs As ADODB.Recordset
Dim b As Boolean


    On Error GoTo ECCoste

    ComprobarCCoste = False
            
    SQL = "SELECT distinct concefact.codconce "
    SQL = SQL & ", concefact.codccost"
    SQL = SQL & " from ((linfact "
    SQL = SQL & " INNER JOIN tmpfactu ON linfact.codsecci=tmpfactu.codsecci and linfact.letraser=tmpfactu.numserie AND linfact.numfactu=tmpfactu.numfactu AND linfact.fecfactu=tmpfactu.fecfactu) "
    SQL = SQL & " INNER JOIN concefact on linfact.codconce = concefact.codconce) "

    Set Rs = New ADODB.Recordset
    Rs.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText

    b = True
    
    While Not Rs.EOF
       'comprobar que el Centro de Coste existe en la Contabilidad
       If DBLet(Rs.Fields(1).Value, "T") <> "" Then
            SQL = DevuelveDesdeBDNewFac("cabccost", "codccost", "codccost", Rs.Fields(1).Value, "T")
            If SQL = "" Then
                b = False
                SQL = "No existe el centro de coste: " & DBLet(Rs.Fields(1).Value, "T")
                SQL = SQL & " del concepto: " & DBLet(Rs.Fields(0).Value, "N")
                InsertarError SQL
            End If
       Else
            b = False
            SQL = "El concepto: " & DBLet(Rs.Fields(0).Value, "N")
            SQL = SQL & " no tiene centro de coste asociado. "
            InsertarError SQL
       End If
       Rs.MoveNext
    Wend
    
    ComprobarCCoste = b
    Set Rs = Nothing
    
ECCoste:
     If Err.Number <> 0 Then
        MuestraError Err.Number, "Comprobar Centros de Coste", Err.Description
    End If
End Function

