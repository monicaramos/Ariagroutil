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
Dim Sql As String
    
    On Error GoTo ECrear
    
    CrearTMPFacturas = False
    
    Sql = "CREATE TEMPORARY TABLE tmpfactu ( "
    If Facturas Then
        Sql = Sql & "codsecci smallint(2) NOT NULL default 0,"
    End If
    Sql = Sql & "numserie char(3) NOT NULL default '',"
    Sql = Sql & "numfactu mediumint(7) unsigned NOT NULL default '0',"
    Sql = Sql & "fecfactu date NOT NULL default '0000-00-00') "
    conn.Execute Sql
     
    If Facturas Then
        Sql = "SELECT codsecci, letraser, numfactu, fecfactu"
    Else
        If Telefono Then
            Sql = "SELECT numserie, numfactu, fecfactu"
        Else
            Sql = "SELECT letraser, numfactu, fecfactu"
        End If
    End If
    Sql = Sql & " FROM " & cadTABLA
    Sql = Sql & " WHERE " & cadwhere
    Sql = " INSERT INTO tmpfactu " & Sql
    conn.Execute Sql

    CrearTMPFacturas = True
    
ECrear:
     If Err.Number <> 0 Then
        CrearTMPFacturas = False
        'Borrar la tabla temporal
        Sql = " DROP TABLE IF EXISTS tmpfactu;"
        conn.Execute Sql
    End If
End Function

Public Function CrearTMPFacturasCV(cadTABLA As String, cadwhere As String, Optional Facturas As Boolean, Optional Telefono As Boolean) As Boolean
'Crea una temporal donde inserta la clave primaria de las
'facturas seleccionadas para facturar y trabaja siempre con ellas
' facturas indica si viene de facturas varias o de telefonia
Dim Sql As String
    
    On Error GoTo ECrear
    
    CrearTMPFacturasCV = False
    
    Sql = "CREATE TEMPORARY TABLE tmpfactu ( "
    If Facturas Then
        Sql = Sql & "codsecci smallint(2) NOT NULL default 0,"
    End If
    Sql = Sql & "numserie char(3) NOT NULL default '',"
    Sql = Sql & "numfactu varchar(10) NOT NULL default '',"
    Sql = Sql & "fecfactu date NOT NULL default '0000-00-00') "
    conn.Execute Sql
     
    If Facturas Then
        Sql = "SELECT codsecci, letraser, numfactu, fecfactu"
    Else
        If Telefono Then
            Sql = "SELECT numserie, numfactu, fecfactu"
        Else
            Sql = "SELECT letraser, numfactu, fecfactu"
        End If
    End If
    Sql = Sql & " FROM " & cadTABLA
    Sql = Sql & " WHERE " & cadwhere
    Sql = " INSERT INTO tmpfactu " & Sql
    conn.Execute Sql

    CrearTMPFacturasCV = True
    
ECrear:
     If Err.Number <> 0 Then
        CrearTMPFacturasCV = False
        'Borrar la tabla temporal
        Sql = " DROP TABLE IF EXISTS tmpfactu;"
        conn.Execute Sql
    End If
End Function


Public Function CrearTMPFacturasProveedor(cadTABLA As String, cadwhere As String) As Boolean
'Crea una temporal donde inserta la clave primaria de las
'facturas seleccionadas para facturar y trabaja siempre con ellas
' facturas indica si viene de facturas varias o de telefonia
Dim Sql As String
    
    On Error GoTo ECrear
    
    CrearTMPFacturasProveedor = False
    
    Sql = "CREATE TEMPORARY TABLE tmpfactu ( "
    Sql = Sql & "numfactu mediumint(7) unsigned NOT NULL default '0',"
    Sql = Sql & "fecfactu date NOT NULL default '0000-00-00', "
    Sql = Sql & "codmacta varchar(10) NOT NULL ) "
    
    conn.Execute Sql
     
    Sql = "SELECT numfactu, fecfactu, codmacta"
    Sql = Sql & " FROM " & cadTABLA
    Sql = Sql & " WHERE " & cadwhere
    Sql = " INSERT INTO tmpfactu " & Sql
    conn.Execute Sql

    CrearTMPFacturasProveedor = True
    
ECrear:
     If Err.Number <> 0 Then
        CrearTMPFacturasProveedor = False
        'Borrar la tabla temporal
        Sql = " DROP TABLE IF EXISTS tmpfactu;"
        conn.Execute Sql
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
Dim Sql As String
    
    On Error GoTo ECrear
    
    CrearTMPErrFact = False
    
    Sql = "CREATE TEMPORARY TABLE tmperrfac ( "
    If cadTABLA = "schfac" Or cadTABLA = "telmovil" Then
        Sql = Sql & "codtipom char(3) NOT NULL default '',"
        Sql = Sql & "numfactu mediumint(7) unsigned NOT NULL default '0',"
    Else
        If cadTABLA = "cvfacturas" Then
            Sql = Sql & "codtipom char(3) NOT NULL default '',"
            Sql = Sql & "numfactu varchar(10), "
        End If
    End If
    Sql = Sql & "fecfactu date NOT NULL default '0000-00-00', "
    Sql = Sql & "error varchar(400) NULL )"
    conn.Execute Sql
     
    CrearTMPErrFact = True
    
ECrear:
     If Err.Number <> 0 Then
        CrearTMPErrFact = False
        'Borrar la tabla temporal
        Sql = " DROP TABLE IF EXISTS tmperrfac;"
        conn.Execute Sql
    End If
End Function


Public Function CrearTMPErrComprob() As Boolean
'Crea una temporal donde insertara la clave primaria de las
'facturas erroneas al facturar
Dim Sql As String
    
    On Error GoTo ECrear
    
    CrearTMPErrComprob = False
    
    Sql = "CREATE TEMPORARY TABLE tmperrcomprob ( "
    Sql = Sql & "error varchar(100) NULL )"
    conn.Execute Sql
     
    CrearTMPErrComprob = True
    
ECrear:
     If Err.Number <> 0 Then
        CrearTMPErrComprob = False
        'Borrar la tabla temporal
        Sql = " DROP TABLE IF EXISTS tmperrcomprob;"
        conn.Execute Sql
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
Dim Sql As String
Dim Rs As ADODB.Recordset
Dim RSconta As ADODB.Recordset
Dim b As Boolean
Dim Cad As String, devuelve As String

On Error GoTo EComprobarLetra

    ComprobarLetraSerie = False
    
    'Comprobar que existe la letra de serie en contabilidad
        'cargamos el RSConta con la tabla contadores de BD: Contabilidad
        'donde estan todas las letra de serie que existen en la contabilidad
    Sql = "Select distinct tiporegi from contadores"
    Set RSconta = New ADODB.Recordset
    Select Case bd
        Case cConta
            RSconta.Open Sql, ConnConta, adOpenDynamic, adLockPessimistic, adCmdText
        Case cContaSeg
            RSconta.Open Sql, ConnContaSeg, adOpenDynamic, adLockPessimistic, adCmdText
        Case cContaTel
            RSconta.Open Sql, ConnContaTel, adOpenDynamic, adLockPessimistic, adCmdText
    End Select
    
    If RSconta.EOF Then
        RSconta.Close
        Set RSconta = Nothing
        Exit Function
    End If
        

    'obtenemos los distintos tipos de movimiento que vamos a contabilizar
    'de las facturas seleccionadas
    Sql = "select distinct numserie from tmpfactu "

    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
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
                Cad = Sql & " en BD de Contabilidad."
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
Dim Sql As String
Dim Rs As ADODB.Recordset
Dim RSconta As ADODB.Recordset
Dim b As Boolean
Dim Cad As String, devuelve As String

On Error GoTo EComprobarLetraFac

    ComprobarLetraSerieFac = False
    
    'Comprobar que existe la letra de serie en contabilidad
        'cargamos el RSConta con la tabla contadores de BD: Contabilidad
        'donde estan todas las letra de serie que existen en la contabilidad
    Sql = "Select distinct tiporegi from contadores"
    Set RSconta = New ADODB.Recordset
    RSconta.Open Sql, ConnContaFac, adOpenDynamic, adLockPessimistic, adCmdText
    
    If RSconta.EOF Then
        RSconta.Close
        Set RSconta = Nothing
        Exit Function
    End If
        

    'obtenemos los distintos tipos de movimiento que vamos a contabilizar
    'de las facturas seleccionadas
    Sql = "select distinct numserie from tmpfactu "

    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
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
Dim Sql As String
Dim Rs As ADODB.Recordset
Dim RSconta As ADODB.Recordset
Dim b As Boolean
Dim Cad As String, devuelve As String

On Error GoTo EComprobarLetraGas

    ComprobarLetraSerieGas = False
    
    'Comprobar que existe la letra de serie en contabilidad
        'cargamos el RSConta con la tabla contadores de BD: Contabilidad
        'donde estan todas las letra de serie que existen en la contabilidad
    Sql = "Select distinct tiporegi from contadores"
    Set RSconta = New ADODB.Recordset
    RSconta.Open Sql, ConnContaGas, adOpenDynamic, adLockPessimistic, adCmdText
    
    If RSconta.EOF Then
        RSconta.Close
        Set RSconta = Nothing
        Exit Function
    End If
        

    'obtenemos los distintos tipos de movimiento que vamos a contabilizar
    'de las facturas seleccionadas
    Sql = "select distinct numserie from tmpfactu "

    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
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
                Cad = Sql & " en BD de Contabilidad."
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
Dim Sql As String
Dim Rs As ADODB.Recordset
Dim RSconta As ADODB.Recordset
Dim b As Boolean
Dim Cad As String, devuelve As String

On Error GoTo EComprobarLetraCV

    ComprobarLetraSerieCV = False
    
    'Comprobar que existe la letra de serie en contabilidad
        'cargamos el RSConta con la tabla contadores de BD: Contabilidad
        'donde estan todas las letra de serie que existen en la contabilidad
    Sql = "Select distinct tiporegi from contadores"
    Set RSconta = New ADODB.Recordset
    If tipo = 0 Then
        RSconta.Open Sql, ConnContaCVV, adOpenDynamic, adLockPessimistic, adCmdText
    Else
        RSconta.Open Sql, ConnContaCV, adOpenDynamic, adLockPessimistic, adCmdText
    End If
    
    If RSconta.EOF Then
        RSconta.Close
        Set RSconta = Nothing
        Exit Function
    End If
        

    'obtenemos los distintos tipos de movimiento que vamos a contabilizar
    'de las facturas seleccionadas
    Sql = "select distinct numserie from tmpfactu "

    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
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
            Sql = Rs!numserie
            If RSconta.EOF Then
                'no encontrado
                b = False
                Cad = Sql & " en BD de Contabilidad."
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
Dim Sql As String
Dim Rs As ADODB.Recordset
Dim RSconta As ADODB.Recordset
Dim b As Boolean

    On Error GoTo ECompFactu

    ComprobarNumFacturas = False
    
    Sql = "SELECT numserie,codfaccl,anofaccl FROM cabfact "
    Sql = Sql & " WHERE " & cadWConta
    
    Set RSconta = New ADODB.Recordset
    Select Case bd
        Case cConta
            RSconta.Open Sql, ConnConta, adOpenForwardOnly, adLockPessimistic, adCmdText
        Case cContaSeg
            RSconta.Open Sql, ConnContaSeg, adOpenForwardOnly, adLockPessimistic, adCmdText
        Case cContaTel
            RSconta.Open Sql, ConnContaTel, adOpenForwardOnly, adLockPessimistic, adCmdText
        Case cContaGas
            RSconta.Open Sql, ConnContaGas, adOpenForwardOnly, adLockPessimistic, adCmdText
        Case cContaCV
            RSconta.Open Sql, ConnContaCV, adOpenForwardOnly, adLockPessimistic, adCmdText
        Case cContaCVV
            RSconta.Open Sql, ConnContaCVV, adOpenForwardOnly, adLockPessimistic, adCmdText
    End Select

    If Not RSconta.EOF Then
        'Seleccionamos las distintas facturas que vamos a facturar
        Sql = "SELECT DISTINCT tmpfactu.numserie,tmpfactu.numfactu,tmpfactu.fecfactu "
        Sql = Sql & " FROM tmpfactu "
        
        Set Rs = New ADODB.Recordset
        Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        b = True
        While Not Rs.EOF 'And b
' quitado el 12022007
'            SQL = "(numserie= " & DBSet(RS!letraser, "T") & " AND codfaccl=" & DBSet(RS!numfactu, "N") & " AND anofaccl=" & Year(RS!fecfactu) & ")"
'            If SituarRSetMULTI(RSconta, SQL) Then
            Sql = ""
            Select Case bd
                Case cConta
                    Sql = DevuelveDesdeBDNew(cConta, "cabfact", "codfaccl", "codfaccl", Rs!numfactu, "N", , "numserie", Trim(Rs!numserie), "T", "anofaccl", Year(Rs!fecfactu), "N")
                Case cContaSeg
                    Sql = DevuelveDesdeBDNew(cContaSeg, "cabfact", "codfaccl", "codfaccl", Rs!numfactu, "N", , "numserie", Trim(Rs!numserie), "T", "anofaccl", Year(Rs!fecfactu), "N")
                Case cContaTel
                    Sql = DevuelveDesdeBDNew(cContaTel, "cabfact", "codfaccl", "codfaccl", Rs!numfactu, "N", , "numserie", Trim(Rs!numserie), "T", "anofaccl", Year(Rs!fecfactu), "N")
                Case cContaGas
                    Sql = DevuelveDesdeBDNew(cContaGas, "cabfact", "codfaccl", "codfaccl", Rs!numfactu, "N", , "numserie", Trim(Rs!numserie), "T", "anofaccl", Year(Rs!fecfactu), "N")
                Case cContaCV
                    Sql = DevuelveDesdeBDNew(cContaCV, "cabfact", "codfaccl", "codfaccl", Rs!numfactu, "N", , "numserie", Trim(Rs!numserie), "T", "anofaccl", Year(Rs!fecfactu), "N")
                Case cContaCVV
                    Sql = DevuelveDesdeBDNew(cContaCVV, "cabfact", "codfaccl", "codfaccl", Rs!numfactu, "N", , "numserie", Trim(Rs!numserie), "T", "anofaccl", Year(Rs!fecfactu), "N")
            End Select
            If Sql <> "" Then
                b = False
                Sql = "          Nº Fac.: " & Format(Rs!numfactu, "0000000") & vbCrLf
                Sql = Sql & "          Fecha: " & Rs!fecfactu
                
                Sql = "Ya existe la factura: " & vbCrLf & Sql
                InsertarError Sql
            
            End If
            Rs.MoveNext
        Wend
        Rs.Close
        Set Rs = Nothing
        
        If Not b Then
            Sql = "Ya existe la factura: " & vbCrLf & Sql
            Sql = "Comprobando Nº Facturas en Contabilidad...       " & vbCrLf & vbCrLf & Sql
            
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
Dim Sql As String
Dim Rs As ADODB.Recordset
Dim RSconta As ADODB.Recordset
Dim b As Boolean

    On Error GoTo ECompFactu

    ComprobarNumFacturasContaNueva = False
    
    Sql = "SELECT numserie,numfactu,anofactu FROM factcli "
    Sql = Sql & " WHERE " & cadWConta
    
    Set RSconta = New ADODB.Recordset
    Select Case bd
        Case cConta
            RSconta.Open Sql, ConnConta, adOpenForwardOnly, adLockPessimistic, adCmdText
        Case cContaSeg
            RSconta.Open Sql, ConnContaSeg, adOpenForwardOnly, adLockPessimistic, adCmdText
        Case cContaTel
            RSconta.Open Sql, ConnContaTel, adOpenForwardOnly, adLockPessimistic, adCmdText
        Case cContaGas
            RSconta.Open Sql, ConnContaGas, adOpenForwardOnly, adLockPessimistic, adCmdText
        Case cContaCV
            RSconta.Open Sql, ConnContaCV, adOpenForwardOnly, adLockPessimistic, adCmdText
        Case cContaCVV
            RSconta.Open Sql, ConnContaCVV, adOpenForwardOnly, adLockPessimistic, adCmdText
    End Select

    If Not RSconta.EOF Then
        'Seleccionamos las distintas facturas que vamos a facturar
        Sql = "SELECT DISTINCT tmpfactu.numserie,tmpfactu.numfactu,tmpfactu.fecfactu "
        Sql = Sql & " FROM tmpfactu "
        
        Set Rs = New ADODB.Recordset
        Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        b = True
        While Not Rs.EOF 'And b
' quitado el 12022007
'            SQL = "(numserie= " & DBSet(RS!letraser, "T") & " AND codfaccl=" & DBSet(RS!numfactu, "N") & " AND anofaccl=" & Year(RS!fecfactu) & ")"
'            If SituarRSetMULTI(RSconta, SQL) Then
            Sql = ""
            Select Case bd
                Case cConta
                    Sql = DevuelveDesdeBDNew(cConta, "factcli", "numfactu", "numfactu", Rs!numfactu, "N", , "numserie", Trim(Rs!numserie), "T", "anofactu", Year(Rs!fecfactu), "N")
                Case cContaSeg
                    Sql = DevuelveDesdeBDNew(cContaSeg, "factcli", "numfactu", "numfactu", Rs!numfactu, "N", , "numserie", Trim(Rs!numserie), "T", "anofactu", Year(Rs!fecfactu), "N")
                Case cContaTel
                    Sql = DevuelveDesdeBDNew(cContaTel, "factcli", "numfactu", "numfactu", Rs!numfactu, "N", , "numserie", Trim(Rs!numserie), "T", "anofactu", Year(Rs!fecfactu), "N")
                Case cContaGas
                    Sql = DevuelveDesdeBDNew(cContaGas, "factcli", "numfactu", "numfactu", Rs!numfactu, "N", , "numserie", Trim(Rs!numserie), "T", "anofactu", Year(Rs!fecfactu), "N")
                Case cContaCV
                    Sql = DevuelveDesdeBDNew(cContaCV, "factcli", "numfactu", "numfactu", Rs!numfactu, "N", , "numserie", Trim(Rs!numserie), "T", "anofactu", Year(Rs!fecfactu), "N")
                Case cContaCVV
                    Sql = DevuelveDesdeBDNew(cContaCVV, "factcli", "numfactu", "numfactu", Rs!numfactu, "N", , "numserie", Trim(Rs!numserie), "T", "anofactu", Year(Rs!fecfactu), "N")
            End Select
            If Sql <> "" Then
                b = False
                Sql = "          Nº Fac.: " & Format(Rs!numfactu, "0000000") & vbCrLf
                Sql = Sql & "          Fecha: " & Rs!fecfactu
                
                Sql = "Ya existe la factura: " & vbCrLf & Sql
                InsertarError Sql
            
            End If
            Rs.MoveNext
        Wend
        Rs.Close
        Set Rs = Nothing
        
        If Not b Then
            Sql = "Ya existe la factura: " & vbCrLf & Sql
            Sql = "Comprobando Nº Facturas en Contabilidad...       " & vbCrLf & vbCrLf & Sql
            
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
Dim Sql As String
Dim Rs As ADODB.Recordset
Dim RSconta As ADODB.Recordset
Dim b As Boolean

    On Error GoTo ECompFactuFac

    ComprobarNumFacturasFac = False
    
    Sql = "SELECT numserie,codfaccl,anofaccl FROM cabfact "
    Sql = Sql & " WHERE " & cadWConta
    
    Set RSconta = New ADODB.Recordset
    RSconta.Open Sql, ConnContaFac, adOpenForwardOnly, adLockPessimistic, adCmdText

    If Not RSconta.EOF Then
        'Seleccionamos las distintas facturas que vamos a facturar
        Sql = "SELECT DISTINCT tmpfactu.numserie,tmpfactu.numfactu,tmpfactu.fecfactu "
        Sql = Sql & " FROM tmpfactu "
        
        Set Rs = New ADODB.Recordset
        Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        b = True
        While Not Rs.EOF 'And b
' quitado el 12022007
'            SQL = "(numserie= " & DBSet(RS!letraser, "T") & " AND codfaccl=" & DBSet(RS!numfactu, "N") & " AND anofaccl=" & Year(RS!fecfactu) & ")"
'            If SituarRSetMULTI(RSconta, SQL) Then
            Sql = ""
            Sql = DevuelveDesdeBDNewFac("cabfact", "codfaccl", "codfaccl", Rs!numfactu, "N", , "numserie", Trim(Rs!numserie), "T", "anofaccl", Year(Rs!fecfactu), "N")
            If Sql <> "" Then
                b = False
                Sql = "          Nº Fac.: " & Format(Rs!numfactu, "0000000") & vbCrLf
                Sql = Sql & "          Fecha: " & Rs!fecfactu
                
                Sql = "Ya existe la factura: " & vbCrLf & Sql
                InsertarError Sql
            
            End If
            Rs.MoveNext
        Wend
        Rs.Close
        Set Rs = Nothing
        
        If Not b Then
            Sql = "Ya existe la factura: " & vbCrLf & Sql
            Sql = "Comprobando Nº Facturas en Contabilidad...       " & vbCrLf & vbCrLf & Sql
            
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
Dim Sql As String
Dim Rs As ADODB.Recordset
Dim RSconta As ADODB.Recordset
Dim b As Boolean

    On Error GoTo ECompFactuFac

    ComprobarNumFacturasFacContaNueva = False
    
    Sql = "SELECT numserie,numfactu,anofactu FROM factcli "
    Sql = Sql & " WHERE " & cadWConta
    
    Set RSconta = New ADODB.Recordset
    RSconta.Open Sql, ConnContaFac, adOpenForwardOnly, adLockPessimistic, adCmdText

    If Not RSconta.EOF Then
        'Seleccionamos las distintas facturas que vamos a facturar
        Sql = "SELECT DISTINCT tmpfactu.numserie,tmpfactu.numfactu,tmpfactu.fecfactu "
        Sql = Sql & " FROM tmpfactu "
        
        Set Rs = New ADODB.Recordset
        Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        b = True
        While Not Rs.EOF 'And b
' quitado el 12022007
'            SQL = "(numserie= " & DBSet(RS!letraser, "T") & " AND codfaccl=" & DBSet(RS!numfactu, "N") & " AND anofaccl=" & Year(RS!fecfactu) & ")"
'            If SituarRSetMULTI(RSconta, SQL) Then
            Sql = ""
            Sql = DevuelveDesdeBDNewFac("factcli", "numfactu", "numfactu", Rs!numfactu, "N", , "numserie", Trim(Rs!numserie), "T", "anofactu", Year(Rs!fecfactu), "N")
            If Sql <> "" Then
                b = False
                Sql = "          Nº Fac.: " & Format(Rs!numfactu, "0000000") & vbCrLf
                Sql = Sql & "          Fecha: " & Rs!fecfactu
                
                Sql = "Ya existe la factura: " & vbCrLf & Sql
                InsertarError Sql
            
            End If
            Rs.MoveNext
        Wend
        Rs.Close
        Set Rs = Nothing
        
        If Not b Then
            Sql = "Ya existe la factura: " & vbCrLf & Sql
            Sql = "Comprobando Nº Facturas en Contabilidad...       " & vbCrLf & vbCrLf & Sql
            
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
Dim Sql As String
Dim Rs As ADODB.Recordset
Dim RSconta As ADODB.Recordset
Dim b As Boolean
Dim cadG As String
Dim enc As String
    
    On Error GoTo ECompCta

    ComprobarCtaContable = False
    
    Sql = "SELECT codmacta FROM cuentas "
    Sql = Sql & " WHERE apudirec='S'"
    If cadG <> "" Then Sql = Sql & cadG
    
    Set RSconta = New ADODB.Recordset
    Select Case bd
        Case cConta
            RSconta.Open Sql, ConnConta, adOpenStatic, adLockPessimistic, adCmdText
        Case cContaSeg
            RSconta.Open Sql, ConnContaSeg, adOpenStatic, adLockPessimistic, adCmdText
        Case cContaTel
            RSconta.Open Sql, ConnContaTel, adOpenStatic, adLockPessimistic, adCmdText
        Case cContaFacSoc
            RSconta.Open Sql, ConnContaFacSoc, adOpenStatic, adLockPessimistic, adCmdText
        Case cContaCV
            RSconta.Open Sql, ConnContaCV, adOpenStatic, adLockPessimistic, adCmdText
        Case cContaCVV
            RSconta.Open Sql, ConnContaCVV, adOpenStatic, adLockPessimistic, adCmdText
    End Select

    If Not RSconta.EOF Then
        If Opcion = 1 Then
                Sql = "SELECT DISTINCT avnic.codmacta, avnic.codavnic  "
                Sql = Sql & " FROM avnic, movim  "
                Sql = Sql & " where " & cadwhere & " and avnic.codavnic = movim.codavnic and avnic.anoejerc = movim.anoejerc "
        ElseIf Opcion = 2 Then
                Sql = "SELECT distinct segpoliza.codmacta, segpoliza.codrefer  "
                Sql = Sql & " from segpoliza "
                Sql = Sql & " where " & cadwhere
        ElseIf Opcion = 3 Then
                'si hay analitica comprobar que todas las cuentas
                'empiezan por el digito que hay en conta.parametros.grupovta
                cadG = DevuelveDesdeBDNew(cConta, "parametros", "grupovta", "", "", "")
        
                Sql = "SELECT distinct sartic.codartic "
                Sql = Sql & ", sartic.codmacta, sartic.codmaccl"
                Sql = Sql & " from ((slhfac "
                Sql = Sql & " INNER JOIN tmpfactu ON slhfac.letraser=tmpfactu.letraser AND slhfac.numfactu=tmpfactu.numfactu AND slhfac.fecfactu=tmpfactu.fecfactu) "
                Sql = Sql & "INNER JOIN sartic ON slhfac.codartic=sartic.codartic) "
                Sql = Sql & " where sartic.codmacta "
                If cadG <> "" Then
                     Sql = Sql & " AND not ((sartic.codmacta like '" & cadG & "%') and (sartic.codmaccl like '" & cadG & "%'))"
                End If
        ElseIf Opcion = 4 Then
            Sql = "select codmacta from telmovil "
        ElseIf Opcion = 5 Then
            Sql = "select ctabancoseg from sparam "
        ElseIf Opcion = 6 Then
            Sql = "select ctagasto from sparam"
        ElseIf Opcion = 7 Then
            Sql = "select ctareten from sparam"
        ElseIf Opcion = 8 Then
            Sql = "SELECT distinct factsocio.codmacta "
            Sql = Sql & " from (factsocio "
            Sql = Sql & " INNER JOIN tmpfactu ON factsocio.numfactu=tmpfactu.numfactu AND factsocio.fecfactu=tmpfactu.fecfactu AND factsocio.codmacta=tmpfactu.codmacta) "
        ElseIf Opcion = 9 Then
            Sql = " select variedad.codmacta "
            Sql = Sql & " from ((factsocio "
            Sql = Sql & " INNER JOIN tmpfactu ON factsocio.numfactu=tmpfactu.numfactu AND factsocio.fecfactu=tmpfactu.fecfactu AND factsocio.codmacta=tmpfactu.codmacta) "
            Sql = Sql & " INNER JOIN variedad ON factsocio.codvarie=variedad.codvarie) "
        ElseIf Opcion = 10 Then
            Sql = "select ctaretenfacsoc from sparam"
        ElseIf Opcion = 11 Then
            Sql = "SELECT distinct variedad.codmacta "
            Sql = Sql & " from (factsocio "
            Sql = Sql & " INNER JOIN tmpfactu ON factsocio.numfactu=tmpfactu.numfactu AND factsocio.fecfactu=tmpfactu.fecfactu AND factsocio.codmacta=tmpfactu.codmacta) "
            Sql = Sql & " INNER JOIN variedad on factsocio.codvarie = variedad.codvarie "
            Sql = Sql & " where not variedad.codmacta like '" & vEmpresaFacSoc.DigGrupoGto & "%'"
        ElseIf Opcion = 12 Then
            Sql = "SELECT distinct cvfacturas.codmactasoc codmacta "
            Sql = Sql & " from (cvfacturas "
            Sql = Sql & " INNER JOIN tmpfactu ON cvfacturas.numfactu=tmpfactu.numfactu AND cvfacturas.fecfactu=tmpfactu.fecfactu and cvfacturas.letraser=tmpfactu.numserie) "
        ElseIf Opcion = 13 Then
            Sql = "SELECT distinct cvfacturas.codmactavta codmacta "
            Sql = Sql & " from (cvfacturas "
            Sql = Sql & " INNER JOIN tmpfactu ON cvfacturas.numfactu=tmpfactu.numfactu AND cvfacturas.fecfactu=tmpfactu.fecfactu and cvfacturas.letraser=tmpfactu.numserie) "
            
        End If
        
        Set Rs = New ADODB.Recordset
        Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        b = True
        While Not Rs.EOF 'And b
            If Opcion = 3 Then
                Sql = Rs!Codmacta & " o " & Rs!CodmacCl
                Sql = "La cuenta " & Sql & " del articulo " & Rs!CodArtic & " no es del grupo correcto."
                InsertarError Sql
            Else
                If Opcion = 11 Then
                    Sql = Rs!Codmacta
                    Sql = "La cuenta " & Sql & " de la variedad no es del grupo correcto."
                    InsertarError Sql
                Else
                    Sql = "codmacta= " & DBLet(Rs.Fields(0).Value, "T") 'DBSet(RS.Fields(0).Value, "T") '& " and apudirec='S' "
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
                        Sql = Rs!Codmacta & " del Código Avnic " & Format(Rs!codavnic, "0000000")
                        Sql = "No existe la cta contable " & Sql
                        InsertarError Sql
                End If
                If Opcion = 2 Then
                    Sql = Rs!Codmacta & " de la Póliza " & Rs!Codrefer
                    Sql = "No existe la cta contable " & Sql
                    InsertarError Sql
                End If
                If Opcion = 4 Then
                    Sql = "No existe la cta contable " & Sql
                    InsertarError Sql
                End If
                If Opcion = 5 Or Opcion = 6 Or Opcion = 7 Or Opcion = 8 Or Opcion = 9 Then
                    Sql = "No existe la cta contable " & Sql
                    InsertarError Sql
                End If
                If Opcion = 10 Then
                    Sql = "No existe la cta contable de retencion " & Sql
                    InsertarError Sql
                End If
                If Opcion = 12 Or Opcion = 13 Then
                    Sql = "No existe la cta contable  " & DBLet(Rs.Fields(0).Value)
                    InsertarError Sql
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
Dim Sql As String
Dim Rs As ADODB.Recordset
Dim RSconta As ADODB.Recordset
Dim b As Boolean
Dim cadG As String
Dim enc As String
    
    On Error GoTo ECompCta

    ComprobarCtaContableFac = False
    
    Sql = "SELECT codmacta FROM cuentas "
    Sql = Sql & " WHERE apudirec='S'"
    If cadG <> "" Then Sql = Sql & cadG
    
    Set RSconta = New ADODB.Recordset
    RSconta.Open Sql, ConnContaFac, adOpenStatic, adLockPessimistic, adCmdText

    If Not RSconta.EOF Then
        If Opcion = 1 Then
                Sql = "SELECT DISTINCT cabfact.ctaclien, cabfact.numfactu   "
                Sql = Sql & " FROM cabfact  "
                Sql = Sql & " where " & cadwhere
        ElseIf Opcion = 2 Then
                Sql = "SELECT distinct concefact.codmacta, concefact.codconce "
                Sql = Sql & " from concefact, linfact, cabfact "
                Sql = Sql & " where " & cadwhere & " and concefact.codconce = linfact.codconce"
                Sql = Sql & " and cabfact.codsecci = linfact.codsecci "
                Sql = Sql & " and cabfact.letraser = linfact.letraser "
                Sql = Sql & " and cabfact.numfactu = linfact.numfactu "
                Sql = Sql & " and cabfact.fecfactu = linfact.fecfactu "
        ElseIf Opcion = 3 Then
                'si hay analitica comprobar que todas las cuentas
                'empiezan por el digito que hay en conta.parametros.grupovta
                cadG = DevuelveDesdeBDNewFac("parametros", "grupovta", "", "", "")
        
                Sql = "SELECT distinct concefact.codconce "
                Sql = Sql & ", concefact.codmacta"
                Sql = Sql & " from ((linfact "
                Sql = Sql & " INNER JOIN tmpfactu ON linfact.codsecci=tmpfactu.codsecci and linfact.letraser=tmpfactu.numserie AND linfact.numfactu=tmpfactu.numfactu AND linfact.fecfactu=tmpfactu.fecfactu) "
                Sql = Sql & " INNER JOIN concefact on linfact.codconce = concefact.codconce) "
                Sql = Sql & " where concefact.codmacta "
                If cadG <> "" Then
                     Sql = Sql & " AND not (concefact.codmacta like '" & cadG & "%') "
                End If
        ElseIf Opcion = 4 Then
            b = True
            enc = ""
            enc = DevuelveDesdeBDNewFac("cuentas", "codmacta", "codmacta", cadwhere, "T")
            If enc = "" Then
                b = False
                Sql = "No existe la cta contable de banco" & cadwhere
                InsertarError Sql
            End If
        ElseIf Opcion = 5 Then
            Sql = "select ctabancoseg from sparam "
        ElseIf Opcion = 6 Then
            Sql = "select ctagasto from sparam"
        ElseIf Opcion = 7 Then
            Sql = "select ctareten from sparam"
        ElseIf Opcion = 8 Then
            Sql = "SELECT DISTINCT cabfact.cuereten   "
            Sql = Sql & " FROM cabfact  "
            Sql = Sql & " where " & cadwhere
        End If
        If Opcion <> 4 Then
            Set Rs = New ADODB.Recordset
            Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            b = True
            While Not Rs.EOF 'And b
                If Opcion = 3 Then
                    Sql = Rs!Codmacta
                    Sql = "La cuenta " & Sql & " del concepto " & Rs!codConce & " no es del grupo correcto."
                    InsertarError Sql
                Else
                    Sql = "codmacta= " & DBLet(Rs.Fields(0).Value, "T") '& " and apudirec='S' "
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
                            Sql = Rs!Codmacta & " de la Factura " & Format(Rs!numfactu, "0000000")
                            Sql = "No existe la cta contable " & Sql
                            InsertarError Sql
                    End If
                    If Opcion = 2 Then
                        Sql = Rs!Codmacta & " del Concepto " & Rs!codConce
                        Sql = "No existe la cta contable " & Sql
                        InsertarError Sql
                    End If
                    If Opcion = 4 Then
                        Sql = "No existe la cta contable " & Sql
                        InsertarError Sql
                    End If
                    If Opcion = 5 Or Opcion = 6 Or Opcion = 7 Then
                        Sql = "No existe la cta contable " & Sql
                        InsertarError Sql
                    End If
                    If Opcion = 8 Then
                        Sql = DBLet(Rs!cuereten, "T")
                        If Sql <> "" Then
                            b = False
                            Sql = "No existe la cta contable de retención: " & Sql
                            InsertarError Sql
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
Dim Sql As String
Dim Rs As ADODB.Recordset
Dim RSconta As ADODB.Recordset
Dim b As Boolean
Dim cadG As String
Dim enc As String
Dim longitud As Integer
    
    On Error GoTo ECompCta

    ComprobarCtaContableGas = False
    
    Sql = "SELECT codmacta FROM cuentas "
    Sql = Sql & " WHERE apudirec='S'"
    If cadG <> "" Then Sql = Sql & cadG
    
    Set RSconta = New ADODB.Recordset
    RSconta.Open Sql, ConnContaGas, adOpenStatic, adLockPessimistic, adCmdText

    If Not RSconta.EOF Then
        If Opcion = 1 Then
                longitud = vEmpresaGas.DigitosUltimoNivel - vEmpresaGas.DigitosNivelAnterior
                Sql = "SELECT DISTINCT concat(" & vParamAplic.RaizCtaSocGas & ", right(concat('0000000000', codsocio)," & longitud & ")) as codmacta, numfactu  "
                Sql = Sql & " FROM gascabfac  "
                Sql = Sql & " where " & cadwhere
        ElseIf Opcion = 2 Then
                Sql = "SELECT " & vParamAplic.CtaVentasGas & " as codmacta "
                Sql = Sql & " from sparam "
        ElseIf Opcion = 3 Then
                'si hay analitica comprobar que la cuenta de ventas
                'empieza por el digito que hay en conta.parametros.grupovta
                cadG = DevuelveDesdeBDNew(cContaGas, "parametros", "grupovta", "", "", "")
        
                Sql = "SELECT  " & vParamAplic.CtaVentasGas & " as codmacta "
                Sql = Sql & " from sparam "
                If cadG <> "" Then
                     Sql = Sql & " where not (" & vParamAplic.CtaVentasGas & " like '" & cadG & "%') "
                End If
        ElseIf Opcion = 4 Then
                Sql = "select " & vParamAplic.CtaContraGas & " as codmacta "
                Sql = Sql & " from sparam "
        ElseIf Opcion = 5 Then
            Sql = "select ctabancoseg from sparam "
        ElseIf Opcion = 6 Then
            Sql = "select ctagasto from sparam"
        ElseIf Opcion = 7 Then
            Sql = "select ctareten from sparam"
        End If
        Set Rs = New ADODB.Recordset
        Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        b = True
        While Not Rs.EOF 'And b
            If Opcion = 3 Then
                Sql = vParamAplic.CtaVentasGas
                Sql = "La cuenta " & Sql & " no es del grupo correcto."
                b = False
                InsertarError Sql
            Else
                Sql = "codmacta= " & DBSet(Rs.Fields(0).Value, "T") '& " and apudirec='S' "
            End If
        
            enc = ""
            enc = DevuelveDesdeBDNew(cContaGas, "cuentas", "codmacta", "codmacta", Rs.Fields(0).Value, "T")
                 
            If enc = "" Then
                b = False 'no encontrado
                If Opcion = 1 Then
                    Sql = Rs!Codmacta & " de la Factura " & Format(Rs!numfactu, "0000000")
                    Sql = "No existe la cta contable " & Sql
                    InsertarError Sql
                End If
                If Opcion = 2 Then
                    Sql = Rs!Codmacta & " de ventas de gasolinera "
                    Sql = "No existe la cta contable " & Sql
                    InsertarError Sql
                End If
                If Opcion = 3 Then
                    Sql = Rs!Codmacta
                    Sql = "La cuenta de ventas " & Sql & " no es del grupo correcto."
                    InsertarError Sql
                End If
                If Opcion = 4 Then
                    Sql = "No existe la cta contable " & Sql
                    InsertarError Sql
                End If
                If Opcion = 5 Or Opcion = 6 Or Opcion = 7 Then
                    Sql = "No existe la cta contable " & Sql
                    InsertarError Sql
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
Dim Sql As String
Dim Rs As ADODB.Recordset
Dim RSconta As ADODB.Recordset
Dim b As Boolean
Dim i As Byte
'Dim CodigIVA As String

    On Error GoTo ECompIVA

    ComprobarTiposIVA = False
    
    Sql = "SELECT distinct codigiva FROM tiposiva "
    
    Set RSconta = New ADODB.Recordset
    RSconta.Open Sql, ConnContaFac, adOpenStatic, adLockPessimistic, adCmdText

    If Not RSconta.EOF Then
        'Seleccionamos los distintos tipos de IVA de las facturas a Contabilizar
        For i = 1 To 3
            Sql = "SELECT DISTINCT cabfact.tipoiva" & i
            Sql = Sql & " FROM cabfact "
            Sql = Sql & " INNER JOIN tmpfactu ON cabfact.letraser=tmpfactu.numserie AND cabfact.numfactu=tmpfactu.numfactu AND cabfact.fecfactu=tmpfactu.fecfactu "
            Sql = Sql & " WHERE not isnull(tipoiva" & i & ")"
            '[Monica]19/05/2016: pasamos la seccion
            If Seccion <> "" Then Sql = Sql & " and cabfact.codsecci = " & DBSet(Seccion, "N")
            
            Set Rs = New ADODB.Recordset
            Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            b = True
            While Not Rs.EOF 'And b
                If Rs.Fields(0) <> 0 Then ' añadido pq en arigasol sino tiene tipo de iva pone ceros
                    Sql = "codigiva= " & DBSet(Rs.Fields(0), "N")
                    RSconta.MoveFirst
                    RSconta.Find (Sql), , adSearchForward
                    If RSconta.EOF Then
                        b = False 'no encontrado
                        Sql = "No existe el " & Sql
                        Sql = "Tipo de IVA: " & Rs.Fields(0)
                        InsertarError Sql
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
Dim Sql As String
Dim Rs As ADODB.Recordset
Dim RSconta As ADODB.Recordset
Dim b As Boolean
Dim i As Byte
'Dim CodigIVA As String

    On Error GoTo ECompIVA

    ComprobarTiposIVAGas = False
    
    Sql = "SELECT distinct codigiva FROM tiposiva "
    
    Set RSconta = New ADODB.Recordset
    RSconta.Open Sql, ConnContaGas, adOpenStatic, adLockPessimistic, adCmdText

    If Not RSconta.EOF Then
        'Seleccionamos los distintos tipos de IVA de las facturas a Contabilizar
        Sql = "SELECT DISTINCT gascabfac.codiva "
        Sql = Sql & " FROM gascabfac "
        Sql = Sql & " INNER JOIN tmpfactu ON gascabfac.letraser=tmpfactu.numserie AND gascabfac.numfactu=tmpfactu.numfactu AND gascabfac.fecfactu=tmpfactu.fecfactu "
        Sql = Sql & " WHERE not isnull(codiva)"

        Set Rs = New ADODB.Recordset
        Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        b = True
        While Not Rs.EOF 'And b
            If Rs.Fields(0) <> 0 Then ' añadido pq en arigasol sino tiene tipo de iva pone ceros
                Sql = "codigiva= " & DBSet(Rs.Fields(0), "N")
                RSconta.MoveFirst
                RSconta.Find (Sql), , adSearchForward
                If RSconta.EOF Then
                    b = False 'no encontrado
                    Sql = "No existe el " & Sql
                    Sql = "Tipo de IVA: " & Rs.Fields(0)
                    InsertarError Sql
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
Dim Sql As String
Dim Rs As ADODB.Recordset
Dim RSconta As ADODB.Recordset
Dim b As Boolean
Dim i As Byte
'Dim CodigIVA As String

    On Error GoTo ECompIVA

    ComprobarTiposIVAFacSoc = False
    
    Sql = "SELECT distinct codigiva FROM tiposiva "
    
    Set RSconta = New ADODB.Recordset
    RSconta.Open Sql, ConnContaFacSoc, adOpenStatic, adLockPessimistic, adCmdText

    If Not RSconta.EOF Then
        'Seleccionamos los distintos tipos de IVA de las facturas a Contabilizar
            Sql = "SELECT DISTINCT factsocio.tipoiva"
            Sql = Sql & " FROM factsocio "
            Sql = Sql & " INNER JOIN tmpfactu ON factsocio.numfactu=tmpfactu.numfactu AND factsocio.fecfactu=tmpfactu.fecfactu AND factsocio.codmacta=tmpfactu.codmacta "
            Sql = Sql & " WHERE not isnull(tipoiva)"

            Set Rs = New ADODB.Recordset
            Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            b = True
            While Not Rs.EOF 'And b
                If Rs.Fields(0) <> 0 Then ' añadido pq en arigasol sino tiene tipo de iva pone ceros
                    Sql = "codigiva= " & DBSet(Rs.Fields(0), "N")
                    RSconta.MoveFirst
                    RSconta.Find (Sql), , adSearchForward
                    If RSconta.EOF Then
                        b = False 'no encontrado
                        Sql = "No existe el " & Sql
                        Sql = "Tipo de IVA: " & Rs.Fields(0)
                        InsertarError Sql
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
Dim Sql As String
Dim Rs As ADODB.Recordset
Dim RSconta As ADODB.Recordset
Dim b As Boolean
Dim i As Byte
'Dim CodigIVA As String

    On Error GoTo ECompIVA

    ComprobarTiposIVACV = False
    
    Sql = "SELECT distinct codigiva FROM tiposiva "
    
    Set RSconta = New ADODB.Recordset
    If tipo = 0 Then
        RSconta.Open Sql, ConnContaCVV, adOpenStatic, adLockPessimistic, adCmdText
    Else
        RSconta.Open Sql, ConnContaCV, adOpenStatic, adLockPessimistic, adCmdText
    End If

    If Not RSconta.EOF Then
        'Seleccionamos los distintos tipos de IVA de las facturas a Contabilizar
        Sql = "SELECT DISTINCT cvfacturas.codiva, cvfacturas.codiva2, cvfacturas.codiva3 "
        Sql = Sql & " FROM cvfacturas "
        Sql = Sql & " INNER JOIN tmpfactu ON cvfacturas.letraser=tmpfactu.numserie AND cvfacturas.numfactu=tmpfactu.numfactu AND cvfacturas.fecfactu=tmpfactu.fecfactu "
        Sql = Sql & " WHERE not isnull(codiva)"

        Set Rs = New ADODB.Recordset
        Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        b = True
        While Not Rs.EOF 'And b
            If Rs.Fields(0) <> 0 Then ' añadido pq en arigasol sino tiene tipo de iva pone ceros
                Sql = "codigiva= " & DBSet(Rs.Fields(0), "N")
                RSconta.MoveFirst
                RSconta.Find (Sql), , adSearchForward
                If RSconta.EOF Then
                    b = False 'no encontrado
                    Sql = "No existe el " & Sql
                    Sql = "Tipo de IVA: " & Rs.Fields(0)
                    InsertarError Sql
                End If
            End If
            If Rs.Fields(1) <> 0 Then ' añadido pq en arigasol sino tiene tipo de iva pone ceros
                Sql = "codigiva= " & DBSet(Rs.Fields(1), "N")
                RSconta.MoveFirst
                RSconta.Find (Sql), , adSearchForward
                If RSconta.EOF Then
                    b = False 'no encontrado
                    Sql = "No existe el " & Sql
                    Sql = "Tipo de IVA: " & Rs.Fields(1)
                    InsertarError Sql
                End If
            End If
            If Rs.Fields(2) <> 0 Then ' añadido pq en arigasol sino tiene tipo de iva pone ceros
                Sql = "codigiva= " & DBSet(Rs.Fields(2), "N")
                RSconta.MoveFirst
                RSconta.Find (Sql), , adSearchForward
                If RSconta.EOF Then
                    b = False 'no encontrado
                    Sql = "No existe el " & Sql
                    Sql = "Tipo de IVA: " & Rs.Fields(2)
                    InsertarError Sql
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



Public Function PasarFactura(cadwhere As String, FecVenci As String, ctaVta As String, CtaBanco As String, FPago As String, Tipiva As String, CodCCost As String, ByRef vContaFra As cContabilizarFacturas) As Boolean
'Insertar en tablas cabecera/lineas de la Contabilidad la factura
' arigasol.schfac --> conta.cabfact
' arigasol.slhfac --> conta.linfact
'Actualizar la tabla ariges.scafac.inconta=1 para indicar que ya esta contabilizada
Dim b As Boolean
Dim cadMen As String
Dim Sql As String
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
    b = InsertarCabFact(cadwhere, Tipiva, cadMen, PorIva, FPago, vContaFra)
    cadMen = "Insertando Cab. Factura: " & cadMen
    
    ' insertar en tesoreria
    If b Then
        Sql = "select * from telmovil where " & cadwhere
        Set Rsx = New ADODB.Recordset
        Rsx.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        
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
            If vParamAplic.ContabilidadNueva Then vContaFra.AnyadeElError vContaFra.IntegraLaFacturaCliente(vContaFra.NumeroFactura, vContaFra.Anofac, vContaFra.Serie)
            
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
        
        Sql = "Insert into tmperrfac(codtipom,numfactu,fecfactu,error) "
        Sql = Sql & " Select *," & DBSet(cadMen, "T") & " as error From tmpfactu "
        Sql = Sql & " WHERE " & Replace(cadwhere, "telmovil", "tmpfactu")
        conn.Execute Sql
        
        
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
Dim Sql As String
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
        Sql = "select * from cvfacturas where " & cadwhere
        Set Rsx = New ADODB.Recordset
        Rsx.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        
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
        
        Sql = "Insert into tmperrfac "
        Sql = Sql & " Select *, " & DBSet(cadMen, "T") & "  as error From tmpfactu "
        Sql = Sql & " WHERE (numserie,numfactu,fecfactu) in (select letraser,numfactu,fecfactu from cvfacturas where " & cadwhere & ")"
        conn.Execute Sql
        
        
        PasarFacturaCV = False
    End If
End Function





Public Function PasarFacturaFac(cadwhere As String, FecVenci As String, CtaBanco As String, CodCCost As String, ByRef vContaFra As cContabilizarFacturas) As Boolean
'Insertar en tablas cabecera/lineas de la Contabilidad la factura
' ariagroutil.cabfact --> conta.cabfact
' ariagroutil.linfact --> conta.linfact
'Actualizar la tabla ariagroutil.cabfact.inconta=1 para indicar que ya esta contabilizada

Dim b As Boolean
Dim cadMen As String
Dim Sql As String
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
    b = InsertarCabFactFac(cadwhere, cadMen, vContaFra)
    cadMen = "Insertando Cab. Factura: " & cadMen
    
    ' insertar en tesoreria
    If b Then
        Sql = "select * from cabfact where " & cadwhere
        Set Rsx = New ADODB.Recordset
        Rsx.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        
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
            If vParamAplic.ContabilidadNueva Then vContaFra.AnyadeElError vContaFra.IntegraLaFacturaCliente(vContaFra.NumeroFactura, vContaFra.Anofac, vContaFra.Serie)
        
        
            'Poner intconta=1 en ariagroutil.cabfact
            b = ActualizarCabFact("cabfact", cadwhere, cadMen)
            cadMen = "Actualizando Factura: " & cadMen
        End If
    End If
    
    If Not b Then
        Sql = "Insert into tmperrfac(codtipom,numfactu,fecfactu,error) "
        Sql = Sql & " Select *," & DBSet(cadMen, "T") & " as error From tmpfactu "
        Sql = Sql & " WHERE " & Replace(cadwhere, "cabfact", "tmpfactu")
        conn.Execute Sql
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


Public Function PasarFacturaGas(cadwhere As String, fecfactu As String, NumAsien As String, NumLinea As String, CodCCost As String, ByRef Diferencia As Currency, ByRef vContaFra As cContabilizarFacturas) As Boolean
'Insertar en tablas cabecera/lineas de la Contabilidad la factura
' arigasol.schfac --> conta.cabfact
' arigasol.slhfac --> conta.linfact
'Actualizar la tabla ariges.scafac.inconta=1 para indicar que ya esta contabilizada
Dim b As Boolean
Dim cadMen As String
Dim Sql As String
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
    b = InsertarCabFactGas(cadwhere, vParamAplic.CodIvaGas, cadMen, vContaFra)
    cadMen = "Insertando Cab. Factura: " & cadMen
    
    ' insertar en linea de asiento
    If b Then
        
        Sql = ""
        'i = 0
        
        ImporteD = 0
        ImporteH = 0
        
        ampliacion = "Fact.Gasolinera"
        ampliaciond = Trim(DevuelveDesdeBDNew(cContaGas, "conceptos", "nomconce", "codconce", vParamAplic.ConceDebeGas, "N")) & " " & ampliacion
        ampliacionh = Trim(DevuelveDesdeBDNew(cContaGas, "conceptos", "nomconce", "codconce", vParamAplic.ConceHaberGas, "N")) & " " & ampliacion

        ' ******************IMPORTE de la poliza
        
        Sql = "select * from gascabfac where " & cadwhere
        Set Rs = New ADODB.Recordset
        Rs.Open Sql, conn, adOpenStatic, adLockPessimistic, adCmdText
        
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
            If vParamAplic.ContabilidadNueva Then vContaFra.AnyadeElError vContaFra.IntegraLaFacturaCliente(vContaFra.NumeroFactura, vContaFra.Anofac, vContaFra.Serie)
            
            'Poner intconta=1 en arigasol.gascabfac
            b = ActualizarCabFact("gascabfac", cadwhere, cadMen)
            cadMen = "Actualizando Factura: " & cadMen
        End If
    End If
    
    If Not b Then
        Sql = "Insert into tmperrfac(codtipom,numfactu,fecfactu,error) "
        Sql = Sql & " Select *," & DBSet(cadMen, "T") & " as error From tmpfactu "
        Sql = Sql & " WHERE " & Replace(cadwhere, "gascabfac", "tmpfactu")
        conn.Execute Sql
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



Private Function InsertarCabFact(cadwhere As String, TipoIva As String, caderr As String, PorcIva As String, FPago As String, ByRef vContaFra As cContabilizarFacturas) As Boolean
'Insertando en tabla conta.cabfact
Dim Sql As String
Dim Rs As ADODB.Recordset
Dim Cad As String
Dim SqlDatos As String
Dim RsDatos As ADODB.Recordset
Dim sql2 As String
Dim CadenaInsertFaclin2 As String

    On Error GoTo EInsertar
    
    Sql = " SELECT numserie,numfactu,fecfactu,codmacta, year(fecfactu) as anofaccl,"
    Sql = Sql & "baseimpo,cuotaiva,totalfac "
    Sql = Sql & " FROM telmovil "
    Sql = Sql & " WHERE " & cadwhere
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    Cad = ""
    If Not Rs.EOF Then
        'guardamos estos valores para utilizarlos cuando insertemos las lineas de la factura
        vContaFra.NumeroFactura = DBLet(Rs!numfactu)
        vContaFra.Serie = DBLet(Rs!numserie)
        vContaFra.Anofac = DBLet(Rs!anofaccl)
        
        Sql = ""
        Sql = DBSet(Trim(Rs!numserie), "T") & "," & DBSet(Rs!numfactu, "N") & "," & DBSet(Rs!fecfactu, "F") & "," & DBSet(Rs!Codmacta, "T") & "," & Year(Rs!fecfactu) & ",'FACTURACION',"
        
        If Not vParamAplic.ContabilidadNueva Then
            Sql = Sql & DBSet(Rs!BaseImpo, "N") & "," & ValorNulo & "," & ValorNulo & "," & DBSet(PorcIva, "N") & "," & ValorNulo & "," & ValorNulo & ","
            Sql = Sql & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & DBSet(Rs!CuotaIva, "N", "N") & "," & ValorNulo & "," & ValorNulo & ","
            Sql = Sql & ValorNulo & "," & ValorNulo & "," & ValorNulo & ","
            Sql = Sql & DBSet(Rs!TotalFac, "N") & "," & DBSet(TipoIva, "N") & "," & ValorNulo & "," & ValorNulo & ",0,"
            Sql = Sql & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & ","
            Sql = Sql & DBSet(Rs!fecfactu, "F")
            Cad = Cad & "(" & Sql & ")"
    '        RS.MoveNext
        
            'Insertar en la contabilidad
            Sql = "INSERT INTO cabfact (numserie,codfaccl,fecfaccl,codmacta,anofaccl,confaccl,ba1faccl,ba2faccl,ba3faccl,"
            Sql = Sql & "pi1faccl,pi2faccl,pi3faccl,pr1faccl,pr2faccl,pr3faccl,ti1faccl,ti2faccl,ti3faccl,tr1faccl,tr2faccl,tr3faccl,"
            Sql = Sql & "totfaccl,tp1faccl,tp2faccl,tp3faccl,intracom,retfaccl,trefaccl,cuereten,numdiari,fechaent,numasien,fecliqcl) "
            Sql = Sql & " VALUES " & Cad
            ConnContaTel.Execute Sql
        Else
            SqlDatos = "select * from cuentas where codmacta = " & DBSet(Rs!ctaclien, "T")
            Set RsDatos = New ADODB.Recordset
            RsDatos.Open SqlDatos, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            If Not RsDatos.EOF Then
                Sql = Sql & "'0'"
                Sql = Sql & "0," & DBSet(FPago, "N") & "," & DBSet(Rs!BaseImpo, "N") & "," & ValorNulo & "," & DBSet(Rs!CuotaIva, "N") & ","
                Sql = Sql & ValorNulo & "," & DBSet(Rs!TotalFac, "N") & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & ",0," & DBSet(Rs!fecfactu, "F") & ","
                Sql = Sql & DBSet(RsDatos!Nommacta, "T", "S") & "," & DBSet(RsDatos!dirdatos, "T", "S") & "," & DBSet(RsDatos!desPobla, "T", "S") & ","
                Sql = Sql & DBSet(RsDatos!Codposta, "T", "S") & "," & DBSet(RsDatos!desProvi, "T", "S") & "," & DBSet(RsDatos!nifdatos, "T", "S") & ",'ES')"
            Else
                Sql = Sql & "'0'"
                Sql = Sql & "0," & DBSet(FPago, "N") & "," & DBSet(Rs!BaseImpo, "N") & "," & ValorNulo & "," & DBSet(Rs!CuotaIva, "N") & ","
                Sql = Sql & ValorNulo & "," & DBSet(Rs!TotalFac, "N") & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & ",0," & DBSet(Rs!fecfactu, "F") & ","
                Sql = Sql & ValorNulo & "," & ValorNulo & "," & ValorNulo & ","
                Sql = Sql & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & ")"
            End If
            Set RsDatos = Nothing
            
            Sql = "(" & Sql & ")"
            
            sql2 = "INSERT INTO factcli (numserie,numfactu,fecfactu,codmacta,anofactu,observa,codconce340,codopera,codforpa,totbases,totbasesret,totivas,"
            sql2 = sql2 & "totrecargo,totfaccl, retfaccl,trefaccl,cuereten,tiporeten,fecliqcl,nommacta,dirdatos,codpobla,despobla, desprovi,nifdatos,"
            sql2 = sql2 & "codpais,codagente)"
            sql2 = sql2 & " VALUES " & Sql
            
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
    
            Sql = "INSERT INTO factcli_totales(numserie,numfactu,fecfactu,anofactu,numlinea,baseimpo,codigiva,"
            Sql = Sql & "porciva,porcrec,impoiva,imporec) VALUES " & CadenaInsertFaclin2
            ConnContaTel.Execute Sql
        
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
Dim Sql As String
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
    
    Sql = " SELECT letraser,numfactu,fecfactu,codmactasoc,codmactavta,  year(fecfactu) as anofac,"
    Sql = Sql & "baseimpo,porciva,codiva,cuotaiva,baseimpo2,porciva2,codiva2,cuotaiva2,baseimpo3,porciva3,codiva3,cuotaiva3,totalfac "
    Sql = Sql & " FROM cvfacturas "
    Sql = Sql & " WHERE " & cadwhere
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
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
            
            Sql = ""
            Sql = DBSet(Trim(Rs!letraser), "T") & "," & DBSet(numfactu, "N") & "," & DBSet(Rs!fecfactu, "F") & "," & DBSet(Rs!CodmactaSoc, "T") & "," & Year(Rs!fecfactu) & ",'FACTURACION',"
            Sql = Sql & DBSet(Rs!BaseImpo, "N") & "," & DBSet(Rs!BaseImpo2, "N", "S") & "," & DBSet(Rs!BaseImpo3, "N", "S") & "," & DBSet(Rs!PorcIva, "N") & "," & DBSet(PorcIva2, "N", "S") & "," & DBSet(PorcIva3, "N", "S") & ","
            Sql = Sql & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & DBSet(Rs!CuotaIva, "N", "N") & "," & DBSet(Rs!cuotaiva2, "N", "S") & "," & DBSet(Rs!cuotaiva3, "N", "S") & ","
            Sql = Sql & ValorNulo & "," & ValorNulo & "," & ValorNulo & ","
            Sql = Sql & DBSet(Rs!TotalFac, "N") & "," & DBSet(Rs!CodIVA, "N") & "," & CodIva2 & "," & CodIva3 & ",0,"
            Sql = Sql & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & ","
            Sql = Sql & DBSet(Rs!fecfactu, "F")
            Cad = Cad & "(" & Sql & ")"
        
            'Insertar en la contabilidad
            Sql = "INSERT INTO cabfact (numserie,codfaccl,fecfaccl,codmacta,anofaccl,confaccl,ba1faccl,ba2faccl,ba3faccl,"
            Sql = Sql & "pi1faccl,pi2faccl,pi3faccl,pr1faccl,pr2faccl,pr3faccl,ti1faccl,ti2faccl,ti3faccl,tr1faccl,tr2faccl,tr3faccl,"
            Sql = Sql & "totfaccl,tp1faccl,tp2faccl,tp3faccl,intracom,retfaccl,trefaccl,cuereten,numdiari,fechaent,numasien,fecliqcl) "
            Sql = Sql & " VALUES " & Cad
            If tipo = 0 Then
                ConnContaCVV.Execute Sql
            Else
                ConnContaCV.Execute Sql
            End If
            
            ' linea de factura de ventas
            BaseImpo = DBLet(Rs!BaseImpo, "N") + DBLet(Rs!BaseImpo2, "N") + DBLet(Rs!BaseImpo3, "N")
            
            Sql = ""
            Sql = "'" & Trim(Rs!letraser) & "'," & DBSet(numfactu, "N") & "," & Year(Rs!fecfactu) & ",1,"
            Sql = Sql & DBSet(Rs!CodmactaVta, "T")
            Sql = Sql & "," & DBSet(BaseImpo, "N") & ","
            If CCoste = "" Then
                Sql = Sql & ValorNulo
            Else
                Sql = Sql & DBSet(CCoste, "T")
            End If
        
            Cad = "(" & Sql & ")"
            'Insertar en la contabilidad
            If Cad <> "" Then
                Sql = "INSERT INTO linfact (numserie,codfaccl,anofaccl,numlinea,codtbase,impbascl,codccost) "
                Sql = Sql & " VALUES " & Cad
                If tipo = 0 Then
                    ConnContaCVV.Execute Sql
                Else
                    ConnContaCV.Execute Sql
                End If
            End If
            
        Else
            ' factura de compras
            ' cabecera
            Set Mc = New CContadorContab
            
            If Mc.ConseguirContador("1", (Rs!fecfactu <= CDate(FFinCV)), True, cContaCV) = 0 Then
                Sql = ""
                Sql = Mc.Contador & "," & DBSet(Rs!fecfactu, "F") & "," & DBLet(Rs!Anofac, "N") & "," & DBSet(Rs!fecfactu, "F") & "," & DBSet(Rs!numfactu, "T") & "," & DBSet(Rs!CodmactaSoc, "T") & "," & ValorNulo & ","
                Sql = Sql & DBSet(Rs!BaseImpo, "N") & "," & DBSet(Rs!BaseImpo2, "N", "S") & "," & DBSet(Rs!BaseImpo3, "N", "S") & ","
                Sql = Sql & DBSet(Rs!PorcIva, "N") & "," & DBSet(PorcIva2, "N", "S") & "," & DBSet(PorcIva3, "N", "S") & ","
                Sql = Sql & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & DBSet(Rs!CuotaIva, "N") & "," & DBSet(Rs!cuotaiva2, "N", "S") & "," & DBSet(Rs!cuotaiva3, "N", "S") & ","
                Sql = Sql & ValorNulo & "," & ValorNulo & "," & ValorNulo & ","
                Sql = Sql & DBSet(Rs!TotalFac, "N") & "," & DBSet(Rs!CodIVA, "N") & "," & CodIva2 & "," & CodIva3 & ",0,"
                Sql = Sql & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & DBSet(Rs!fecfactu, "F") & ",0"
                Cad = Cad & "(" & Sql & ")"
                
                'Insertar en la contabilidad
                Sql = "INSERT INTO cabfactprov (numregis,fecfacpr,anofacpr,fecrecpr,numfacpr,codmacta,confacpr,ba1facpr,ba2facpr,ba3facpr,"
                Sql = Sql & "pi1facpr,pi2facpr,pi3facpr,pr1facpr,pr2facpr,pr3facpr,ti1facpr,ti2facpr,ti3facpr,tr1facpr,tr2facpr,tr3facpr,"
                Sql = Sql & "totfacpr,tp1facpr,tp2facpr,tp3facpr,extranje,retfacpr,trefacpr,cuereten,numdiari,fechaent,numasien,fecliqpr,nodeducible) "
                Sql = Sql & " VALUES " & Cad
                ConnContaCV.Execute Sql
            End If
            'linea
            
            ' linea de factura de ventas
            BaseImpo = DBLet(Rs!BaseImpo, "N") + DBLet(Rs!BaseImpo2, "N") + DBLet(Rs!BaseImpo3, "N")
            
            Sql = ""
            Sql = Mc.Contador & "," & Year(Rs!fecfactu) & ",1,"
            Sql = Sql & DBSet(Rs!CodmactaVta, "T")
            Sql = Sql & "," & DBSet(BaseImpo, "N") & ","
            
            If CCoste = "" Then
                Sql = Sql & ValorNulo
            Else
                Sql = Sql & DBSet(CCoste, "T")
            End If
            Cad = "(" & Sql & ")"
            
            'Insertar en la contabilidad
            If Cad <> "" Then
                Sql = "INSERT INTO linfactprov (numregis,anofacpr,numlinea,codtbase,impbaspr,codccost) "
                Sql = Sql & " VALUES " & Cad
                ConnContaCV.Execute Sql
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




Private Function InsertarCabFactGas(cadwhere As String, TipoIva As String, caderr As String, ByRef vContaFra As cContabilizarFacturas) As Boolean
'Insertando en tabla conta.cabfact
Dim Sql As String
Dim Rs As ADODB.Recordset
Dim Cad As String
Dim Codmacta As String
Dim FPago As String
Dim RsDatos As ADODB.Recordset
Dim SqlDatos As String
Dim sql2 As String
Dim CadenaInsertFaclin2 As String



    On Error GoTo EInsertar
    
    Sql = " SELECT letraser,numfactu,fecfactu,codsocio, year(fecfactu) as anofaccl,"
    Sql = Sql & "base as baseimpo,iva as cuotaiva,total as totalfac, codiva, porciva "
    Sql = Sql & " FROM gascabfac "
    Sql = Sql & " WHERE " & cadwhere
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    Cad = ""
    If Not Rs.EOF Then
        'guardamos estos valores para utilizarlos cuando insertemos las lineas de la factura
        
        vContaFra.NumeroFactura = DBLet(Rs!numfactu)
        vContaFra.Serie = DBLet(Rs!letraser)
        vContaFra.Anofac = DBLet(Rs!anofaccl)
        
        Codmacta = vParamAplic.RaizCtaSocGas & Right("0000000000" & DBLet(Rs!Codsocio, "N"), vEmpresaGas.DigitosUltimoNivel - vEmpresaGas.DigitosNivelAnterior)
        
        Sql = ""
        Sql = DBSet(Trim(Rs!letraser), "T") & "," & DBSet(Rs!numfactu, "N") & "," & DBSet(Rs!fecfactu, "F") & "," & DBSet(Codmacta, "T") & "," & Year(Rs!fecfactu) & ",'FACTURACION',"
        
        If Not vParamAplic.ContabilidadNueva Then
            Sql = Sql & DBSet(Rs!BaseImpo, "N") & "," & ValorNulo & "," & ValorNulo & "," & DBSet(Rs!PorcIva, "N") & "," & ValorNulo & "," & ValorNulo & ","
            Sql = Sql & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & DBSet(Rs!CuotaIva, "N", "N") & "," & ValorNulo & "," & ValorNulo & ","
            Sql = Sql & ValorNulo & "," & ValorNulo & "," & ValorNulo & ","
            Sql = Sql & DBSet(Rs!TotalFac, "N") & "," & DBSet(Rs!CodIVA, "N") & "," & ValorNulo & "," & ValorNulo & ",0,"
            Sql = Sql & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & ","
            Sql = Sql & DBSet(Rs!fecfactu, "F")
            Cad = Cad & "(" & Sql & ")"
    
            'Insertar en la contabilidad
            Sql = "INSERT INTO cabfact (numserie,codfaccl,fecfaccl,codmacta,anofaccl,confaccl,ba1faccl,ba2faccl,ba3faccl,"
            Sql = Sql & "pi1faccl,pi2faccl,pi3faccl,pr1faccl,pr2faccl,pr3faccl,ti1faccl,ti2faccl,ti3faccl,tr1faccl,tr2faccl,tr3faccl,"
            Sql = Sql & "totfaccl,tp1faccl,tp2faccl,tp3faccl,intracom,retfaccl,trefaccl,cuereten,numdiari,fechaent,numasien,fecliqcl) "
            Sql = Sql & " VALUES " & Cad
            ConnContaGas.Execute Sql
        Else
            FPago = DevuelveDesdeBDNew(cContaGas, "formapago", "min(codforpa)", "", "")
        
            SqlDatos = "select * from cuentas where codmacta = " & DBSet(Codmacta, "T")
            Set RsDatos = New ADODB.Recordset
            RsDatos.Open SqlDatos, ConnContaGas, adOpenForwardOnly, adLockPessimistic, adCmdText
            If Not RsDatos.EOF Then
                Sql = Sql & "'0',"
                Sql = Sql & "0," & DBSet(FPago, "N") & "," & DBSet(Rs!BaseImpo, "N") & "," & ValorNulo & "," & DBSet(Rs!CuotaIva, "N") & ","
                Sql = Sql & ValorNulo & "," & DBSet(Rs!TotalFac, "N") & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & ",0," & DBSet(Rs!fecfactu, "F") & ","
                Sql = Sql & DBSet(RsDatos!Nommacta, "T", "S") & "," & DBSet(RsDatos!dirdatos, "T", "S") & "," & DBSet(RsDatos!desPobla, "T", "S") & ","
                Sql = Sql & DBSet(RsDatos!Codposta, "T", "S") & "," & DBSet(RsDatos!desProvi, "T", "S") & "," & DBSet(RsDatos!nifdatos, "T", "S") & ",'ES'"
            Else
                Sql = Sql & "'0',"
                Sql = Sql & "0," & DBSet(FPago, "N") & "," & DBSet(Rs!BaseImpo, "N") & "," & ValorNulo & "," & DBSet(Rs!CuotaIva, "N") & ","
                Sql = Sql & ValorNulo & "," & DBSet(Rs!TotalFac, "N") & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & ",0," & DBSet(Rs!fecfactu, "F") & ","
                Sql = Sql & ValorNulo & "," & ValorNulo & "," & ValorNulo & ","
                Sql = Sql & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo
            End If
            Set RsDatos = Nothing
            Sql = "(" & Sql & ")"
            
            sql2 = "INSERT INTO factcli (numserie,numfactu,fecfactu,codmacta,anofactu,observa,codconce340,codopera,codforpa,totbases,totbasesret,totivas,"
            sql2 = sql2 & "totrecargo,totfaccl, retfaccl,trefaccl,cuereten,tiporeten,fecliqcl,nommacta,dirdatos,codpobla,despobla, desprovi,nifdatos,"
            sql2 = sql2 & "codpais)"
            sql2 = sql2 & " VALUES " & Sql
            
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
    
            Sql = "INSERT INTO factcli_totales(numserie,numfactu,fecfactu,anofactu,numlinea,baseimpo,codigiva,"
            Sql = Sql & "porciva,porcrec,impoiva,imporec) VALUES " & CadenaInsertFaclin2
            ConnContaGas.Execute Sql
        
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



Private Function InsertarCabFactFac(cadwhere As String, caderr As String, ByRef vContaFra As cContabilizarFacturas) As Boolean
'Insertando en tabla conta.cabfact
Dim Sql As String
Dim Rs As ADODB.Recordset
Dim Cad As String
Dim SqlDatos As String
Dim RsDatos As ADODB.Recordset
Dim sql2 As String
Dim CadenaInsertFaclin2 As String

    On Error GoTo EInsertar
    
    Sql = " SELECT letraser,numfactu,fecfactu,ctaclien, year(fecfactu) as anofaccl,"
    Sql = Sql & "baseiva1, baseiva2, baseiva3, impoiva1, impoiva2, impoiva3, imporec1,"
    Sql = Sql & "imporec2, imporec3, totalfac, tipoiva1, tipoiva2, tipoiva3, porciva1,"
    Sql = Sql & "porciva2, porciva3, porcrec1, porcrec2, porcrec3, totalfac, retfaccl, "
    Sql = Sql & "trefaccl, cuereten, codforpa "
    '[Monica]27/11/2017: hemos insertado todos los datos fiscales en la tabla
    Sql = Sql & ", nommacta, dirdatos, codposta, despobla, desprovi, nifdatos, codpais "
    Sql = Sql & " FROM cabfact "
    Sql = Sql & " WHERE " & cadwhere
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    Cad = ""
    If Not Rs.EOF Then
        'guardamos estos valores para utilizarlos cuando insertemos las lineas de la factura
        vContaFra.NumeroFactura = DBLet(Rs!numfactu)
        vContaFra.Serie = DBLet(Rs!letraser)
        vContaFra.Anofac = DBLet(Rs!anofaccl)
        
        Sql = ""
        Sql = DBSet(Trim(Rs!letraser), "T") & "," & DBSet(Rs!numfactu, "N") & "," & DBSet(Rs!fecfactu, "F") & "," & DBSet(Rs!ctaclien, "T") & "," & Year(Rs!fecfactu) & ",'FACTURACION',"
        
        BaseImp = Rs!baseiva1 + CCur(DBLet(Rs!baseiva2, "N")) + CCur(DBLet(Rs!baseiva3, "N"))
        IvaImp = DBLet(Rs!impoiva1, "N") + DBLet(Rs!impoiva2, "N") + DBLet(Rs!impoiva3, "N")
        
        
        If Not vParamAplic.ContabilidadNueva Then
        
            Sql = Sql & DBSet(Rs!baseiva1, "N") & "," & DBSet(Rs!baseiva2, "N") & "," & DBSet(Rs!baseiva3, "N") & "," & DBSet(Rs!porciva1, "N") & "," & DBSet(Rs!PorcIva2, "N") & "," & DBSet(Rs!PorcIva3, "N") & ","
            Sql = Sql & DBSet(Rs!porcrec1, "N") & "," & DBSet(Rs!porcrec2, "N") & "," & DBSet(Rs!porcrec3, "N") & "," & DBSet(Rs!impoiva1, "N") & "," & DBSet(Rs!impoiva2, "N") & "," & DBSet(Rs!impoiva3, "N") & ","
            Sql = Sql & DBSet(Rs!imporec1, "N") & "," & DBSet(Rs!imporec2, "N") & "," & DBSet(Rs!imporec3, "N") & ","
            Sql = Sql & DBSet(Rs!TotalFac, "N") & "," & DBSet(Rs!tipoiva1, "N") & "," & DBSet(Rs!tipoiva2, "N") & "," & DBSet(Rs!tipoiva3, "N") & ",0,"
            Sql = Sql & DBSet(Rs!retfaccl, "N") & "," & DBSet(Rs!trefaccl, "N") & "," & DBSet(Rs!cuereten, "T") & ","
            Sql = Sql & ValorNulo & "," & ValorNulo & "," & ValorNulo & ","
            Sql = Sql & DBSet(Rs!fecfactu, "F")
        
            Cad = Cad & "(" & Sql & ")"
            
            'Insertar en la contabilidad
            Sql = "INSERT INTO cabfact (numserie,codfaccl,fecfaccl,codmacta,anofaccl,confaccl,ba1faccl,ba2faccl,ba3faccl,"
            Sql = Sql & "pi1faccl,pi2faccl,pi3faccl,pr1faccl,pr2faccl,pr3faccl,ti1faccl,ti2faccl,ti3faccl,tr1faccl,tr2faccl,tr3faccl,"
            Sql = Sql & "totfaccl,tp1faccl,tp2faccl,tp3faccl,intracom,retfaccl,trefaccl,cuereten,numdiari,fechaent,numasien,fecliqcl) "
            Sql = Sql & " VALUES " & Cad
            ConnContaFac.Execute Sql
            
        
        Else
            '[Monica]28/11/2017: ahora los datos los cogemos de la factura
'            SqlDatos = "select * from cuentas where codmacta = " & DBSet(Rs!ctaclien, "T")
'            Set RsDatos = New ADODB.Recordset
'            RsDatos.Open SqlDatos, ConnContaFac, adOpenForwardOnly, adLockPessimistic, adCmdText
'            If Not RsDatos.EOF Then
'                Sql = Sql & "'0',"
'                Sql = Sql & "0," & DBSet(Rs!CodForpa, "N") & "," & DBSet(BaseImp, "N") & "," & ValorNulo & "," & DBSet(IvaImp, "N") & ","
'                Sql = Sql & ValorNulo & "," & DBSet(Rs!TotalFac, "N") & "," & DBSet(Rs!retfaccl, "N") & "," & DBSet(Rs!trefaccl, "N") & "," & DBSet(Rs!cuereten, "T")
'                If DBLet(Rs!trefaccl, "N") <> 0 Then
'                    Sql = Sql & ",1," & DBSet(Rs!fecfactu, "F") & ","
'                Else
'                    Sql = Sql & ",0," & DBSet(Rs!fecfactu, "F") & ","
'                End If
'                Sql = Sql & DBSet(RsDatos!Nommacta, "T", "S") & "," & DBSet(RsDatos!dirdatos, "T", "S") & "," & DBSet(RsDatos!desPobla, "T", "S") & ","
'                Sql = Sql & DBSet(RsDatos!Codposta, "T", "S") & "," & DBSet(RsDatos!desProvi, "T", "S") & "," & DBSet(RsDatos!nifdatos, "T", "S") & ",'ES',1"
'            Else
'                Sql = Sql & "'0',"
'                Sql = Sql & "0," & DBSet(Rs!CodForpa, "N") & "," & DBSet(BaseImp, "N") & "," & ValorNulo & "," & DBSet(IvaImp, "N") & ","
'                Sql = Sql & ValorNulo & "," & DBSet(Rs!TotalFac, "N") & "," & DBSet(Rs!retfaccl, "N") & "," & DBSet(Rs!trefaccl, "N") & "," & DBSet(Rs!cuereten, "T")
'                If DBLet(Rs!trefaccl, "N") <> 0 Then
'                    Sql = Sql & ",1," & DBSet(Rs!fecfactu, "F") & ","
'                Else
'                    Sql = Sql & ",0," & DBSet(Rs!fecfactu, "F") & ","
'                End If
'
'                Sql = Sql & ValorNulo & "," & ValorNulo & "," & ValorNulo & ","
'                Sql = Sql & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & ",1"
'            End If
            
            Sql = Sql & "'0',"
            Sql = Sql & "0," & DBSet(Rs!CodForpa, "N") & "," & DBSet(BaseImp, "N") & "," & ValorNulo & "," & DBSet(IvaImp, "N") & ","
            Sql = Sql & ValorNulo & "," & DBSet(Rs!TotalFac, "N") & "," & DBSet(Rs!retfaccl, "N") & "," & DBSet(Rs!trefaccl, "N") & "," & DBSet(Rs!cuereten, "T")
            If DBLet(Rs!trefaccl, "N") <> 0 Then
                Sql = Sql & ",1," & DBSet(Rs!fecfactu, "F") & ","
            Else
                Sql = Sql & ",0," & DBSet(Rs!fecfactu, "F") & ","
            End If
            Sql = Sql & DBSet(Rs!Nommacta, "T", "S") & "," & DBSet(Rs!dirdatos, "T", "S") & "," & DBSet(Rs!desPobla, "T", "S") & ","
            Sql = Sql & DBSet(Rs!Codposta, "T", "S") & "," & DBSet(Rs!desProvi, "T", "S") & "," & DBSet(Rs!nifdatos, "T", "S") & "," & DBSet(Rs!codpais, "T") & ",1"
            
            
            Sql = "(" & Sql & ")"
            
            sql2 = "INSERT INTO factcli (numserie,numfactu,fecfactu,codmacta,anofactu,observa,codconce340,codopera,codforpa,totbases,totbasesret,totivas,"
            sql2 = sql2 & "totrecargo,totfaccl, retfaccl,trefaccl,cuereten,tiporeten,fecliqcl,nommacta,dirdatos,codpobla,despobla, desprovi,nifdatos,"
            sql2 = sql2 & "codpais,codagente)"
            sql2 = sql2 & " VALUES " & Sql
            
            ConnContaFac.Execute sql2
            
            CadenaInsertFaclin2 = ""
            
            'numlinea,baseimpo,codigiva,porciva,porcrec,impoiva,imporec)
            'IVA 1, siempre existe
            sql2 = "'" & Rs!letraser & "'," & Rs!numfactu & "," & DBSet(Rs!fecfactu, "F") & "," & Year(Rs!fecfactu) & ","
            sql2 = sql2 & "1," & DBSet(Rs!baseiva1, "N") & "," & Rs!tipoiva1 & "," & DBSet(Rs!porciva1, "N") & ","
            sql2 = sql2 & ValorNulo & "," & DBSet(Rs!impoiva1, "N") & "," & ValorNulo
            CadenaInsertFaclin2 = CadenaInsertFaclin2 & "(" & sql2 & ")"
            
            'para las lineas
            vTipoIva(0) = Rs!tipoiva1
            vPorcIva(0) = Rs!porciva1
            vPorcRec(0) = DBLet(Rs!porcrec1, "N")
            vImpIva(0) = Rs!impoiva1
            vImpRec(0) = DBLet(Rs!imporec1, "N")
            vBaseIva(0) = Rs!baseiva1
            
            vTipoIva(1) = 0: vTipoIva(2) = 0
            
            If Not IsNull(Rs!PorcIva2) Then
                sql2 = "'" & Rs!letraser & "'," & Rs!numfactu & "," & DBSet(Rs!fecfactu, "F") & "," & Year(Rs!fecfactu) & ","
                sql2 = sql2 & "2," & DBSet(Rs!baseiva2, "N") & "," & Rs!tipoiva2 & "," & DBSet(Rs!PorcIva2, "N") & ","
                sql2 = sql2 & ValorNulo & "," & DBSet(Rs!impoiva2, "N") & "," & ValorNulo
                CadenaInsertFaclin2 = CadenaInsertFaclin2 & " , (" & sql2 & ")"
                vTipoIva(1) = Rs!tipoiva2
                vPorcIva(1) = Rs!PorcIva2
                vPorcRec(1) = DBLet(Rs!porcrec2, "N")
                vImpIva(1) = Rs!impoiva2
                vImpRec(1) = DBLet(Rs!imporec2, "N")
                vBaseIva(1) = Rs!baseiva2
            End If
            If Not IsNull(Rs!PorcIva3) Then
                sql2 = "'" & Rs!letraser & "'," & Rs!numfactu & "," & DBSet(Rs!fecfactu, "F") & "," & Year(Rs!fecfactu) & ","
                sql2 = sql2 & "3," & DBSet(Rs!baseiva3, "N") & "," & Rs!tipoiva3 & "," & DBSet(Rs!PorcIva3, "N") & ","
                sql2 = sql2 & ValorNulo & "," & DBSet(Rs!impoiva3, "N") & "," & ValorNulo
                CadenaInsertFaclin2 = CadenaInsertFaclin2 & " , (" & sql2 & ")"
                vTipoIva(2) = Rs!tipoiva3
                vPorcIva(2) = Rs!PorcIva3
                vPorcRec(2) = DBLet(Rs!porcrec3, "N")
                vImpIva(2) = Rs!impoiva3
                vImpRec(2) = DBLet(Rs!imporec3, "N")
                vBaseIva(2) = Rs!baseiva3
            End If
    
            Sql = "INSERT INTO factcli_totales(numserie,numfactu,fecfactu,anofactu,numlinea,baseimpo,codigiva,"
            Sql = Sql & "porciva,porcrec,impoiva,imporec) VALUES " & CadenaInsertFaclin2
            ConnContaFac.Execute Sql
            
            
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
Dim Sql As String
Dim Rs As ADODB.Recordset
Dim Cad As String, Aux As String
Dim i As Long
Dim totimp As Currency, ImpLinea As Currency
Dim CodIVA As String
Dim Iva As String
Dim vIva As Currency


    On Error GoTo EInLinea
    Sql = " SELECT numserie,numfactu,fecfactu,codmacta, "
    Sql = Sql & "baseimpo,cuotaiva,totalfac "
    Sql = Sql & " FROM telmovil "

    Sql = " SELECT numserie, numfactu, fecfactu, baseimpo,year(fecfactu) as anofaccl,cuotaiva,totalfac from " & cadTABLA & " where " & cadwhere
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText

    Cad = ""
    i = 1
    totimp = 0
    If Not Rs.EOF Then
        'concatenamos linea para insertar en la tabla de conta.linfact
        Sql = ""
        Sql = "'" & Trim(Rs!numserie) & "'," & Rs!numfactu & "," & Year(Rs!fecfactu) & "," & i & ","
        
        'dependiendo del colectivo del socio cogemos la cta contable cliente o socio del articulo
        Sql = Sql & DBSet(CtaVenta, "T")
        
        Sql = Sql & "," & DBSet(Rs!BaseImpo, "N") & ","
        
        If CCoste = "" Then
            Sql = Sql & ValorNulo
        Else
            Sql = Sql & DBSet(CCoste, "T")
        End If
        
        If vParamAplic.ContabilidadNueva Then
            Sql = Sql & "," & DBSet(Rs!fecfactu, "F") & "," & DBSet(Tipiva, "N") & "," & DBSet(PorcIva, "N") & "," & ValorNulo & "," & DBSet(Rs!CuotaIva, "N") & "," & ValorNulo
        End If
        
        Cad = Cad & "(" & Sql & ")" & ","
        
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
            Sql = "INSERT INTO factcli_lineas (numserie,numfactu,anofactu,numlinea,codmacta,baseimpo,codccost,fecfactu,codigiva,porciva,porcrec,impoiva,imporec) "
        Else
            Sql = "INSERT INTO linfact (numserie,codfaccl,anofaccl,numlinea,codtbase,impbascl,codccost) "
        End If
        Sql = Sql & " VALUES " & Cad
        ConnContaTel.Execute Sql
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
Dim Sql As String
Dim Rs As ADODB.Recordset
Dim Cad As String, Aux As String
Dim i As Long
Dim totimp As Currency, ImpLinea As Currency
Dim CodIVA As String
Dim Iva As String
Dim vIva As Currency


    On Error GoTo EInLinea

    Sql = " SELECT letraser, numfactu, fecfactu, base, codiva, iva, total, porciva from " & cadTABLA & " where " & cadwhere
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText

    Cad = ""
    i = 1
    totimp = 0
    If Not Rs.EOF Then
        'concatenamos linea para insertar en la tabla de conta.linfact
        Sql = ""
        Sql = "'" & Trim(Rs!letraser) & "'," & Rs!numfactu & "," & Year(Rs!fecfactu) & "," & i & ","
        
        'dependiendo del colectivo del socio cogemos la cta contable cliente o socio del articulo
        Sql = Sql & DBSet(vParamAplic.CtaVentasGas, "T")
        
        Sql = Sql & "," & DBSet(Rs!Base, "N") & ","
        
        If CCoste = "" Then
            Sql = Sql & ValorNulo
        Else
            Sql = Sql & DBSet(CCoste, "T")
        End If
        
        If vParamAplic.ContabilidadNueva Then
            Sql = Sql & "," & DBSet(Rs!fecfactu, "F") & "," & DBSet(Rs!CodIVA, "N") & "," & DBSet(Rs!PorcIva, "N") & ","
            Sql = Sql & ValorNulo & "," & DBSet(Rs!Iva, "N") & "," & ValorNulo
        End If
        
        Cad = Cad & "(" & Sql & ")" & ","
        
        i = i + 1
        Rs.MoveNext
    End If
    
    Rs.Close
    Set Rs = Nothing

    'Insertar en la contabilidad
    If Cad <> "" Then
        Cad = Mid(Cad, 1, Len(Cad) - 1) 'quitar la ult. coma
        If vParamAplic.ContabilidadNueva Then
            Sql = "INSERT INTO factcli_lineas (numserie,numfactu,anofactu,numlinea,codmacta,baseimpo,codccost,fecfactu,codigiva,porciva,porcrec,impoiva,imporec) "
        Else
            Sql = "INSERT INTO linfact (numserie,codfaccl,anofaccl,numlinea,codtbase,impbascl,codccost) "
        End If
        Sql = Sql & " VALUES " & Cad
        ConnContaGas.Execute Sql
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
Dim Sql As String
Dim Rs As ADODB.Recordset
Dim Cad As String, Aux As String
Dim i As Long
Dim totimp As Currency, ImpLinea As Currency
Dim CodIVA As String
Dim Iva As String
Dim vIva As Currency


    On Error GoTo EInLinea

    Sql = " SELECT letraser, numfactu, fecfactu, concefact.codmacta, concefact.codccost, sum(importe) from " & cadTABLA
    Sql = Sql & ", concefact where " & cadwhere
    Sql = Sql & " and concefact.codconce = " & cadTABLA & ".codconce"
    Sql = Sql & " GROUP BY 1,2,3,4,5 "
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText

    Cad = ""
    i = 1
    totimp = 0
    While Not Rs.EOF
        'concatenamos linea para insertar en la tabla de conta.linfact
        Sql = ""
        Sql = "'" & Trim(Rs!letraser) & "'," & Rs!numfactu & "," & Year(Rs!fecfactu) & "," & i & ","
        
        'dependiendo del colectivo del socio cogemos la cta contable cliente o socio del articulo
        Sql = Sql & DBSet(Rs!Codmacta, "T")
        
        Sql = Sql & "," & DBSet(Rs.Fields(5).Value, "N") & ","
        
        If DBLet(Rs!CodCCost, "T") = "" Then
            Sql = Sql & ValorNulo
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
                Sql = Sql & DBSet(Rs!CodCCost, "T")
            Else
                Sql = Sql & ValorNulo
            End If
        End If
        
        Cad = Cad & "(" & Sql & ")" & ","
        
        i = i + 1
        Rs.MoveNext
        
    Wend
    
    Rs.Close
    Set Rs = Nothing

    'Insertar en la contabilidad
    If Cad <> "" Then
        Cad = Mid(Cad, 1, Len(Cad) - 1) 'quitar la ult. coma
        Sql = "INSERT INTO linfact (numserie,codfaccl,anofaccl,numlinea,codtbase,impbascl,codccost) "
        Sql = Sql & " VALUES " & Cad
        ConnContaFac.Execute Sql
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
Dim Sql As String
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


    Sql = " SELECT letraser, numfactu, fecfactu, concefact.codmacta, concefact.codccost, " & cadTABLA & ".tipoiva , sum(importe) from " & cadTABLA
    Sql = Sql & ", concefact where " & cadwhere
    Sql = Sql & " and concefact.codconce = " & cadTABLA & ".codconce"
    Sql = Sql & " GROUP BY 1,2,3,4,5,6 "
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText

    Cad = ""
    i = 1
    totimp = 0
    While Not Rs.EOF
        'concatenamos linea para insertar en la tabla de conta.linfact
        Sql = ""
        Sql = "'" & Trim(Rs!letraser) & "'," & Rs!numfactu & "," & Year(Rs!fecfactu) & "," & i & ","
        
        'dependiendo del colectivo del socio cogemos la cta contable cliente o socio del articulo
        Sql = Sql & DBSet(Rs!Codmacta, "T") & ","
        
        
        ImpLinea = DBLet(Rs.Fields(6).Value, "N")
        'Vemos que tipo de IVA es en el vector de importes
        NumeroIVA = 127
        For k = 0 To 2
            If Rs!TipoIva = vTipoIva(k) Then
                NumeroIVA = k
                Exit For
            End If
        Next
        If NumeroIVA > 100 Then Err.Raise 513, "Error obteniendo IVA: " & Rs!CodigIva
        
        
        
        If DBLet(Rs!CodCCost, "T") = "" Then
            Sql = Sql & ValorNulo
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
                Sql = Sql & DBSet(Rs!CodCCost, "T")
            Else
                Sql = Sql & ValorNulo
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

        Sql = Sql & "," & DBSet(Rs!fecfactu, "F") & "," & vTipoIva(NumeroIVA) & "," & DBSet(vPorcIva(NumeroIVA), "N") & "," & DBSet(vPorcRec(NumeroIVA), "N", "S") & ","

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
        Sql = Sql & DBSet(ImpLinea, "N") & "," & DBSet(ImpImva, "N") & "," & DBSet(ImpRec, "N", "S")
        
        Cad = Cad & "(" & Sql & ")" & ","
        
        i = i + 1
        Rs.MoveNext
        
    Wend
    
    Rs.Close
    Set Rs = Nothing

    'Insertar en la contabilidad
    If Cad <> "" Then
        Cad = Mid(Cad, 1, Len(Cad) - 1) 'quitar la ult. coma
        Sql = "INSERT INTO factcli_lineas (numserie,numfactu,anofactu,numlinea,codmacta,codccost,fecfactu,codigiva,porciva,porcrec,baseimpo,impoiva,imporec) "
        Sql = Sql & " VALUES " & Cad
        ConnContaFac.Execute Sql
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
Dim Sql As String

    On Error GoTo EActualizar
    
    Sql = "UPDATE " & cadTABLA & " SET intconta=1 "
    Sql = Sql & " WHERE " & cadwhere

    conn.Execute Sql
    
EActualizar:
    If Err.Number <> 0 Then
        ActualizarCabFact = False
        caderr = Err.Description
    Else
        ActualizarCabFact = True
    End If
End Function

Private Sub InsertarError(Cadena As String)
Dim Sql As String

    Sql = "insert into tmperrcomprob values ('" & Cadena & "')"
    conn.Execute Sql

End Sub


Public Function InsertarCabAsientoDia(Diario As String, Asiento As String, Fecha As String, Obs As String, caderr As String, bd As Byte) As Boolean
'Insertando en tabla conta.cabfact
Dim Sql As String
Dim Rs As ADODB.Recordset
Dim Cad As String
Dim Nulo2 As String
Dim Nulo3 As String

    On Error GoTo EInsertar
       
    If vParamAplic.ContabilidadNueva Then
        Cad = Format(Diario, "00") & ", " & DBSet(Fecha, "F") & "," & Format(Asiento, "000000") & ","
        Cad = Cad & DBSet(Obs, "T")
        Cad = Cad & "," & DBSet(Now, "FH") & "," & DBSet(vSesion.Login, "T") & ",'ARIAGRO UTILIDADES'"
        Cad = "(" & Cad & ")"

        'Insertar en la contabilidad
        Sql = "INSERT INTO hcabapu (numdiari, fechaent, numasien, obsdiari, feccreacion, usucreacion, desdeaplicacion) "
        Sql = Sql & " VALUES " & Cad
    Else
        Cad = Format(Diario, "00") & ", " & DBSet(Fecha, "F") & "," & Format(Asiento, "000000") & ","
        Cad = Cad & "''," & ValorNulo & "," & DBSet(Obs, "T")
        Cad = "(" & Cad & ")"
    
        'Insertar en la contabilidad
        Sql = "INSERT INTO cabapu (numdiari, fechaent, numasien, bloqactu, numaspre, obsdiari) "
        Sql = Sql & " VALUES " & Cad
    End If
    
    Select Case bd
        Case cConta
            ConnConta.Execute Sql
        Case cContaSeg
            ConnContaSeg.Execute Sql
        Case cContaGas
            ConnContaGas.Execute Sql
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
Dim Sql As String
Dim i As Byte
Dim totimp As Currency, ImpLinea As Currency

    On Error GoTo EInLinea

    If vParamAplic.ContabilidadNueva Then
        Sql = "INSERT INTO hlinapu (numdiari, fechaent, numasien, linliapu, codmacta, numdocum, codconce, "
        Sql = Sql & " ampconce, timporteD, timporteH, codccost, ctacontr, idcontab, punteada) "
        Sql = Sql & " VALUES " & Cad
    Else
        Sql = "INSERT INTO linapu (numdiari, fechaent, numasien, linliapu, codmacta, numdocum, codconce, "
        Sql = Sql & " ampconce, timporteD, timporteH, codccost, ctacontr, idcontab, punteada) "
        Sql = Sql & " VALUES " & Cad
    End If
    
    Select Case bd
        Case cConta
            ConnConta.Execute Sql
        Case cContaSeg
            ConnContaSeg.Execute Sql
        Case cContaGas
            ConnContaGas.Execute Sql
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
Dim Sql As String

    On Error GoTo EActualizar
    
    Sql = "UPDATE movim SET intconta=1 "
    Sql = Sql & " WHERE " & cadwhere

    conn.Execute Sql
    
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
Dim Sql As String
    
    On Error GoTo ECrear
    
    CrearTMPAsiento = False
    
    Sql = "CREATE TEMPORARY TABLE tmpasien ( "
    Sql = Sql & "fecalbar date NOT NULL default '0000-00-00',"
    Sql = Sql & "codturno tinyint(1) NOT NULL default '0',"
    Sql = Sql & "codmacta varchar(10) NOT NULL default ' ',"
    Sql = Sql & "importel decimal(10,2)  NOT NULL default '0.00')"
    conn.Execute Sql

    CrearTMPAsiento = True
    
ECrear:
     If Err.Number <> 0 Then
        CrearTMPAsiento = False
        'Borrar la tabla temporal
        Sql = " DROP TABLE IF EXISTS tmpasien;"
        conn.Execute Sql
    End If
End Function

' ### [Monica] 07/05/2007
Public Function InsertarEnTesoreriaNew(Fechamov As String, FecVenci As String, codavnic As String, anoejerc As Integer, Codmacta2 As String, Concepto As String, forpa As String, MenError As String) As Boolean
'Guarda datos de Tesoreria en tablas: ariges.svenci y en conta.scobros
Dim b As Boolean
Dim Rs As ADODB.Recordset
Dim Rsx As ADODB.Recordset
Dim Sql As String, text1csb As String, text2csb As String
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

    Sql = "select * from movim where fechamov = " & DBSet(Fechamov, "F") & " and codavnic = " & DBSet(codavnic, "N")
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
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
        
            CadValuesAux2 = "("
            If vParamAplic.ContabilidadNueva Then CadValuesAux2 = CadValuesAux2 & DBSet(SerieFraPro, "T") & ","
            CadValuesAux2 = CadValuesAux2 & DBSet(Rs4!Codmacta, "T") & "," & DBSet(codavnic, "N") & ", " & DBSet(Fechamov, "F") & ", 1,"
            
            CadValues2 = CadValuesAux2 & DBSet(forpa, "N") & "," & DBSet(FecVenci, "F") & "," & DBSet(Rs!timporte, "N") & ","
            
            If vParamAplic.ContabilidadNueva Then
                vvIban = MiFormat(Rs4!Iban, "") & MiFormat(Rs4!codbanco, "0000") & MiFormat(Rs4!codsucur, "0000") & MiFormat(DigConta, "00") & MiFormat(Rs4!cuentaba, "0000000000")
            
                CadValues2 = CadValues2 & DBSet(Codmacta2, "T") & "," & text1csb & "," & DBSet(text2csb, "T") & ","
                CadValues2 = CadValues2 & DBSet(vvIban, "T", "S") & ", "
                CadValues2 = CadValues2 & DBSet(Rs4!Nommacta, "T", "S") & "," & DBSet(Rs4!dirdatos, "T", "S") & "," & DBSet(Rs4!desPobla, "T", "S") & ","
                CadValues2 = CadValues2 & DBSet(Rs4!Codposta, "T", "S") & "," & DBSet(Rs4!desProvi, "T", "S") & "," & DBSet(Rs4!nifdatos, "T", "S") & ",'ES')"
                
                Sql = "INSERT INTO pagos (numserie, codmacta, numfactu, fecfactu, numorden, codforpa, fecefect, impefect, ctabanc1,text1csb,text2csb, iban,"
                Sql = Sql & "nomprove, domprove, pobprove, cpprove, proprove, nifprove, codpais)"
            
            Else
                CadValues2 = CadValues2 & ValorNulo & "," & ValorNulo & "," & DBSet(Codmacta2, "T") & ","
                CadValues2 = CadValues2 & ValorNulo & "," & "0,0," & text1csb & "," & DBSet(text2csb, "T") & "," & DBSet(Rs4!codbanco, "N") & ", "
                CadValues2 = CadValues2 & DBSet(Rs4!codsucur, "N") & ", " & DBSet(DigConta, "T") & ", " & DBSet(Rs4!cuentaba, "T") & ", " & ValorNulo ' & ") "
                '[Monica]22/11/2013: tema iban
                If vEmpresa.HayNorma19_34Nueva = 1 Then
                   CadValues2 = CadValues2 & "," & DBSet(Rs4!Iban, "T", "S") & ")"
                Else
                   CadValues2 = CadValues2 & ")"
                End If
                
                'Insertamos en la tabla scobro de la CONTA
                Sql = "INSERT INTO spagop (ctaprove, numfactu, fecfactu, numorden, codforpa, fecefect, impefect,  "
                Sql = Sql & "fecultpa, imppagad, ctabanc1, ctabanc2, emitdocum, contdocu, text1csb, text2csb, entidad, "
                Sql = Sql & "oficina, cc, cuentaba, transfer" ' ) "
                
                '[Monica]22/11/2013: tema iban
                If vEmpresa.HayNorma19_34Nueva = 1 Then
                   Sql = Sql & ",iban)"
                Else
                   Sql = Sql & ")"
                End If
            End If
            
            Sql = Sql & " VALUES " & CadValues2
            ConnConta.Execute Sql
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
Dim Sql As String, text33csb As String, text41csb As String
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
            Sql = "INSERT INTO scobro (numserie, codfaccl, fecfaccl, numorden, codmacta, codforpa, fecvenci, impvenci, gastos,"
            Sql = Sql & "ctabanc1, codbanco, codsucur, digcontr, cuentaba, ctabanc2, fecultco, impcobro,  "
            Sql = Sql & " text33csb, text41csb, agente" ') "
            
            '[Monica]22/11/2013: tema iban
            If vEmpresaSeg.HayNorma19_34Nueva = 1 Then
               Sql = Sql & ",iban)"
            Else
               Sql = Sql & ")"
            End If
        Else
            CadValues2 = CadValues2 & DBSet(CtaBan, "T") & "," & ValorNulo & "," & ValorNulo & ","
            CadValues2 = CadValues2 & text33csb & "," & DBSet(text41csb, "T") & ",1," & DBSet(Rs4!Iban, "T", "S") & ","
            CadValues2 = CadValues2 & DBSet(Rs4!Nommacta, "T", "S") & "," & DBSet(Rs4!dirdatos, "T", "S") & "," & DBSet(Rs4!desPobla, "T", "S") & ","
            CadValues2 = CadValues2 & DBSet(Rs4!Codposta, "T", "S") & "," & DBSet(Rs4!desProvi, "T", "S") & "," & DBSet(Rs4!nifdatos, "T", "S") & ",'ES')"
        
            Sql = "INSERT INTO cobros (numserie, numfactu, fecfactu, numorden, codmacta, codforpa, fecvenci, impvenci, gastos, "
            Sql = Sql & "ctabanc1, fecultco, impcobro, "
            Sql = Sql & " text33csb, text41csb, agente, iban, "
            Sql = Sql & "nomclien, domclien, pobclien, cpclien, proclien, nifclien, codpais"
            Sql = Sql & ") "
        
        End If
        
        Sql = Sql & " VALUES " & CadValues2
        ConnContaSeg.Execute Sql

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
Dim Sql As String, text33csb As String, text41csb As String
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
            Sql = "INSERT INTO scobro (numserie, codfaccl, fecfaccl, numorden, codmacta, codforpa, fecvenci, impvenci, "
            Sql = Sql & "ctabanc1, codbanco, codsucur, digcontr, cuentaba, ctabanc2, fecultco, impcobro, "
            Sql = Sql & " text33csb, text41csb, agente" ') "
            '[Monica]22/11/2013: tema iban
            If vEmpresaTel.HayNorma19_34Nueva = 1 Then
               Sql = Sql & ",iban)"
            Else
               Sql = Sql & ")"
            End If
        Else
        
            CadValues2 = CadValues2 & ValorNulo & "," & ValorNulo & ","
            CadValues2 = CadValues2 & text33csb & "," & DBSet(text41csb, "T") & ",1," & DBSet(Rs4!Iban, "T", "S") & ","
            
            CadValues2 = CadValues2 & DBSet(Rs4!Nommacta, "T", "S") & "," & DBSet(Rs4!dirdatos, "T", "S") & "," & DBSet(Rs4!desPobla, "T", "S") & ","
            CadValues2 = CadValues2 & DBSet(Rs4!Codposta, "T", "S") & "," & DBSet(Rs4!desProvi, "T", "S") & "," & DBSet(Rs4!nifdatos, "T", "S") & ",'ES')"
        
            Sql = "INSERT INTO cobros (numserie, numfactu, fecfactu, numorden, codmacta, codforpa, fecvenci, impvenci, "
            Sql = Sql & "ctabanc1, fecultco, impcobro, "
            Sql = Sql & " text33csb, text41csb, agente, iban, "
            Sql = Sql & "nomclien, domclien, pobclien, cpclien, proclien, nifclien, codpais"
            Sql = Sql & ") "
        
        End If
        Sql = Sql & " VALUES " & CadValues2
        ConnContaTel.Execute Sql

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
Dim Sql As String, text33csb As String, text41csb As String
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
                Sql = "INSERT INTO scobro (numserie, codfaccl, fecfaccl, numorden, codmacta, codforpa, fecvenci, impvenci, "
                Sql = Sql & "ctabanc1, codbanco, codsucur, digcontr, cuentaba, ctabanc2, fecultco, impcobro, "
                Sql = Sql & " text33csb, text41csb, agente" ') "
                '[Monica]22/11/2013: tema iban
                If vEmpresaCVV.HayNorma19_34Nueva = 1 Then
                   Sql = Sql & ",iban) "
                Else
                   Sql = Sql & ")"
                End If
            Else
                CadValues2 = CadValues2 & ValorNulo & "," & ValorNulo & ","
                CadValues2 = CadValues2 & text33csb & "," & DBSet(text41csb, "T") & ",1," & DBSet(Rs4!Iban, "T", "S") & ","
                
                CadValues2 = CadValues2 & DBSet(Rs4!Nommacta, "T", "S") & "," & DBSet(Rs4!dirdatos, "T", "S") & "," & DBSet(Rs4!desPobla, "T", "S") & ","
                CadValues2 = CadValues2 & DBSet(Rs4!Codposta, "T", "S") & "," & DBSet(Rs4!desProvi, "T", "S") & "," & DBSet(Rs4!nifdatos, "T", "S") & ",'ES')"
            
                Sql = "INSERT INTO cobros (numserie, numfactu, fecfactu, numorden, codmacta, codforpa, fecvenci, impvenci, "
                Sql = Sql & "ctabanc1, fecultco, impcobro, "
                Sql = Sql & " text33csb, text41csb, agente, iban, "
                Sql = Sql & "nomclien, domclien, pobclien, cpclien, proclien, nifclien, codpais"
                Sql = Sql & ") "
            
            End If
            
            Sql = Sql & " VALUES " & CadValues2
            If tipo = 0 Then
                ConnContaCVV.Execute Sql
            Else
                ConnContaCV.Execute Sql
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
            Sql = "INSERT INTO spagop (ctaprove, numfactu, fecfactu, numorden, codforpa, fecefect, impefect,  "
            Sql = Sql & "fecultpa, imppagad, ctabanc1, ctabanc2, emitdocum, contdocu, text1csb, text2csb, entidad, "
            Sql = Sql & "oficina, cc, cuentaba, transfer" ') "
            '[Monica]22/11/2013: tema iban
            If vEmpresaCVV.HayNorma19_34Nueva = 1 Then
               Sql = Sql & ",iban)"
            Else
               Sql = Sql & ")"
            End If
            Sql = Sql & " VALUES " & CadValues2
            ConnContaCV.Execute Sql
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
Dim Sql As String, text33csb As String, text41csb As String
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

    If Not vParamAplic.ContabilidadNueva Then

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
            Sql = "INSERT INTO scobro (numserie, codfaccl, fecfaccl, numorden, codmacta, codforpa, fecvenci, impvenci, "
            Sql = Sql & "ctabanc1, codbanco, codsucur, digcontr, cuentaba, ctabanc2, fecultco, impcobro, "
            Sql = Sql & " text33csb, text41csb, agente" ') "
            '[Monica]22/11/2013: tema iban
            If vEmpresaFac.HayNorma19_34Nueva = 1 Then
               Sql = Sql & ",iban)"
            Else
               Sql = Sql & ")"
            End If
            
            Sql = Sql & " VALUES " & CadValues2
            ConnContaFac.Execute Sql
    
        End If
    Else

'[Monica]28/11/2017: para el caso de contabilidad nueva, metemos los datos que hemos dado en la factura
        Sql4 = "select * "
        Sql4 = Sql4 & " from cuentas where codmacta = " & Rsx!ctaclien

        Set Rs4 = New ADODB.Recordset

        Rs4.Open Sql4, ConnContaFac, adOpenForwardOnly, adLockPessimistic, adCmdText

        If Not Rs4.EOF Then

            text33csb = "'Factura:" & DBLet(Trim(Rsx!letraser), "T") & "-" & DBLet(Rsx!numfactu, "T") & " " & Format(DBLet(Rsx!fecfactu, "F"), "dd/mm/yy") & "'"
            text41csb = "de " & DBSet(Rsx!TotalFac, "N")
                  
        
            CadValuesAux2 = "(" & DBSet(Trim(Rsx!letraser), "T") & "," & DBSet(Rsx!numfactu, "N") & "," & DBSet(Rsx!fecfactu, "F") & ", 1," & DBSet(Rsx!ctaclien, "T") & ","
            CadValues2 = CadValuesAux2 & DBSet(Rsx!CodForpa, "N") & "," & DBSet(FecVenci, "F") & "," & DBSet(Rsx!TotalFac, "N") & "," & ValorNulo & ","
            CadValues2 = CadValues2 & DBSet(CtaBan, "T") & "," & ValorNulo & "," & ValorNulo & ","
            CadValues2 = CadValues2 & text33csb & "," & DBSet(text41csb, "T") & ",1," & DBSet(Rs4!Iban, "T", "S") & ","
            '[Monica]28/11/2017: antes era del rs4, ahora lo tenemos en la factura
            CadValues2 = CadValues2 & DBSet(Rsx!Nommacta, "T", "S") & "," & DBSet(Rsx!dirdatos, "T", "S") & "," & DBSet(Rsx!desPobla, "T", "S") & ","
            CadValues2 = CadValues2 & DBSet(Rsx!Codposta, "T", "S") & "," & DBSet(Rsx!desProvi, "T", "S") & "," & DBSet(Rsx!nifdatos, "T", "S") & "," & DBSet(Rsx!codpais, "T") & ")"
           
            'Insertamos en la tabla cobros de la CONTA
            Sql = "INSERT INTO cobros (numserie, numfactu, fecfactu, numorden, codmacta, codforpa, fecvenci, impvenci, gastos, "
            Sql = Sql & "ctabanc1, fecultco, impcobro, "
            Sql = Sql & " text33csb, text41csb, agente, iban, "
            Sql = Sql & "nomclien, domclien, pobclien, cpclien, proclien, nifclien, codpais"
            Sql = Sql & ") "
            
            Sql = Sql & " VALUES " & CadValues2
            ConnContaFac.Execute Sql
        End If
    End If

    b = True

EInsertarTesoreriaNewFac:
    If Err.Number <> 0 Then b = False
    InsertarEnTesoreriaNewFac = b
End Function




Public Function ComprobarFormadePago(cadCC As String) As Boolean
Dim Sql As String
Dim sql2 As String
Dim Rs As ADODB.Recordset
Dim b As Boolean

    ComprobarFormadePago = False
    If vParamAplic.ContabilidadNueva Then
        Sql = DevuelveDesdeBDNew(cContaFacSoc, "formapago", "codforpa", "codforpa", cadCC, "N")
    Else
        Sql = DevuelveDesdeBDNew(cContaFacSoc, "sforpa", "codforpa", "codforpa", cadCC, "N")
    End If
    If Sql = "" Then
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

Public Function PasarFacturaSoc(cadwhere As String, FecVenci As String, CtaBan As String, forpa As String, ByRef vContaFra As cContabilizarFacturas) As Boolean
'Insertar en tablas cabecera/lineas de la Contabilidad la factura socios
' ariagroutil.factsocio  --> conta.cabfactprov
'                        --> conta.linfactprov
'Actualizar la tabla ariagroutil.factsocio.inconta=1 para indicar que ya esta contabilizada
Dim b As Boolean
Dim cadMen As String
Dim Sql As String
Dim Mc As CContadorContab


    On Error GoTo EContab

    ConnContaFacSoc.BeginTrans
    conn.BeginTrans
        
    
    Set Mc = New CContadorContab
    
    '---- Insertar en la conta Cabecera Factura
    b = InsertarCabFactSoc(cadwhere, cadMen, Mc, forpa, vContaFra)
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
            If vParamAplic.ContabilidadNueva Then vContaFra.AnyadeElError vContaFra.IntegraLaFacturaProv(vContaFra.NumeroFactura, vContaFra.Anofac)
            
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
            Sql = "Insert into tmperrfac(codtipom, numfactu,fecfactu,error) "
            Sql = Sql & " select 1,numfactu, fecfactu, " & DBSet(cadMen, "T") & " from factsocio where " & cadwhere
            conn.Execute Sql
        End If
    End If
End Function


Private Function InsertarCabFactSoc(cadwhere As String, caderr As String, ByRef Mc As CContadorContab, forpa As String, ByRef vContaFra As cContabilizarFacturas) As Boolean
'Insertando en tabla conta.cabfact
'(OUT) AnyoFacPr: aqui devolvemos el año de fecha recepcion para insertarlo en las lineas de factura de proveedor de la conta
Dim Sql As String
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
       
   
    Sql = " SELECT fecfactu,year(fecfactu) as anofacpr,codmacta, numfactu, "
    Sql = Sql & "baseimpo,porciva,cuotaiva,basereten,porcreten,impreten,"
    Sql = Sql & "totalfac,tipoiva "
    Sql = Sql & " FROM " & "factsocio "
    Sql = Sql & " WHERE " & cadwhere
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    Cad = ""
    If Not Rs.EOF Then
                                                                    '16/01/2012: antes era: CDate(vEmpresaFacSoc.FechaFin) - 365
        If Mc.ConseguirContador("1", (Rs!fecfactu <= CDate(vEmpresaFacSoc.FechaFin)), True, cContaFacSoc) = 0 Then
            'guardamos estos valores para utilizarlos cuando insertemos las lineas de la factura
            
            vContaFra.NumeroFactura = Mc.Contador
            vContaFra.Anofac = DBLet(Rs!anofacpr, "N")
            
'            DtoPPago = RS!DtoPPago
'            DtoGnral = RS!DtoGnral
            BaseImp = DBLet(Rs!BaseImpo, "N")
            TotalFac = BaseImp + DBLet(Rs!CuotaIva, "N")
'            AnyoFacPr = RS!anofacpr
            
            Sql = ""
            If vParamAplic.ContabilidadNueva Then Sql = Sql & DBSet(SerieFraPro, "T") & ","
            
            Sql = Sql & Mc.Contador & "," & DBSet(Rs!fecfactu, "F") & "," & DBLet(Rs!anofacpr, "N") & "," & DBSet(Rs!fecfactu, "F") & "," & DBSet(Rs!numfactu, "T") & "," & DBSet(Rs!Codmacta, "T") & "," & ValorNulo & ","
            
            
            If Not vParamAplic.ContabilidadNueva Then
            
                Sql = Sql & DBSet(Rs!BaseImpo, "N") & "," & ValorNulo & "," & ValorNulo & ","
                Sql = Sql & DBSet(Rs!PorcIva, "N") & "," & ValorNulo & "," & ValorNulo & ","
                Sql = Sql & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & DBSet(Rs!CuotaIva, "N") & "," & ValorNulo & "," & ValorNulo & ","
                Sql = Sql & ValorNulo & "," & ValorNulo & "," & ValorNulo & ","
                Sql = Sql & DBSet(TotalFac, "N") & "," & DBSet(Rs!TipoIva, "N") & "," & ValorNulo & "," & ValorNulo & ",0,"
                If DBLet(Rs!ImpReten, "N") <> 0 Then
                    Sql = Sql & DBSet(Rs!PorcReten, "N") & "," & DBSet(Rs!ImpReten, "N") & "," & DBSet(vParamAplic.CtaRetenFacSoc, "T") & ","
                Else
                    Sql = Sql & ValorNulo & "," & ValorNulo & "," & ValorNulo & ","
                End If
                Sql = Sql & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & DBSet(Rs!fecfactu, "F") & ",0"
                
                Cad = Cad & "(" & Sql & ")"
                
                'Insertar en la contabilidad
                Sql = "INSERT INTO cabfactprov (numregis,fecfacpr,anofacpr,fecrecpr,numfacpr,codmacta,confacpr,ba1facpr,ba2facpr,ba3facpr,"
                Sql = Sql & "pi1facpr,pi2facpr,pi3facpr,pr1facpr,pr2facpr,pr3facpr,ti1facpr,ti2facpr,ti3facpr,tr1facpr,tr2facpr,tr3facpr,"
                Sql = Sql & "totfacpr,tp1facpr,tp2facpr,tp3facpr,extranje,retfacpr,trefacpr,cuereten,numdiari,fechaent,numasien,fecliqpr,nodeducible) "
                Sql = Sql & " VALUES " & Cad
                ConnContaFacSoc.Execute Sql
                
            Else
                
                SqlDatos = "select * from cuentas where codmacta = " & DBSet(Rs!Codmacta, "T")
                Set RsDatos = New ADODB.Recordset
                RsDatos.Open SqlDatos, ConnContaGas, adOpenForwardOnly, adLockPessimistic, adCmdText
                If Not RsDatos.EOF Then
                
                    Sql = Sql & DBSet(RsDatos!Nommacta, "T", "S") & "," & DBSet(RsDatos!dirdatos, "T", "S") & "," & DBSet(RsDatos!desPobla, "T", "S") & ","
                    Sql = Sql & DBSet(RsDatos!Codposta, "T", "S") & "," & DBSet(RsDatos!desProvi, "T", "S") & "," & DBSet(RsDatos!nifdatos, "T", "S") & ",'ES',"
            
                Else
                
                    Sql = Sql & ValorNulo & "," & ValorNulo & "," & ValorNulo & ","
                    Sql = Sql & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & ","
                 
                End If
                Set RsDatos = Nothing
                
                Sql = Sql & DBSet(forpa, "N") & ","
                
                '$$$
                TipoOpera = 0
                
                Aux = "0"
                'codopera,codconce340,codintra
                Sql = Sql & TipoOpera & "," & DBSet(Aux, "T") & "," & ValorNulo & ","
                
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
                Sql = Sql & DBSet(Rs!BaseImpo, "N") & "," & DBSet(Rs!BaseReten, "N") & ","
                'totivas
                Sql = Sql & DBSet(Rs!CuotaIva, "N") & "," & DBSet(TotalFac, "N") & ","
                If DBLet(Rs!PorcReten, "N") <> 0 Then
                    'porcreten,impreten,
                    Sql = Sql & DBSet(Rs!PorcReten, "N") & "," & DBSet(Rs!ImpReten, "N") & "," & DBSet(vParamAplic.CtaRetenFacSoc, "T") & ",2"
                Else
                    Sql = Sql & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo
                End If
                
                Sql = Sql & "," & DBSet(Rs!fecfactu, "F")
                
                Cad = "(" & Sql & ")"
            
                'Insertar en la contabilidad
                Sql = "INSERT INTO factpro(numserie,numregis,fecfactu,anofactu,fecharec,numfactu,codmacta,observa,nommacta,"
                Sql = Sql & "dirdatos,despobla,codpobla,desprovi,nifdatos,codpais,codforpa,codopera,codconce340,codintra,"
                Sql = Sql & "totbases,totbasesret,totivas,totfacpr,retfacpr,trefacpr,cuereten,tiporeten,fecliqpr)"
                Sql = Sql & " VALUES " & Cad
                ConnContaFacSoc.Execute Sql
            
                'Las  lineas de IVA
                Sql = "INSERT INTO factpro_totales(numserie,numregis,fecharec,anofactu,numlinea,baseimpo,codigiva,porciva,porcrec,impoiva,imporec)"
                Sql = Sql & " VALUES " & CadenaInsertFaclin2
                ConnContaFacSoc.Execute Sql
            
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
Dim Sql As String, text1csb As String, text2csb As String
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

    Sql = "select * from factsocio where " & vWhere
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
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
                CadValuesAux2 = "("
            Else
                CadValuesAux2 = "(" & DBSet(SerieFraPro, "T") & ","
            End If
        
            CadValuesAux2 = CadValuesAux2 & DBSet(Rs!Codmacta, "T") & "," & DBSet(Rs!numfactu, "N") & ", " & DBSet(Rs!fecfactu, "F") & ", 1,"
            
            CadValues2 = CadValuesAux2 & DBSet(forpa, "N") & "," & DBSet(FecVenci, "F") & "," & DBSet(Rs!TotalFac, "N") & "," & ValorNulo & "," & ValorNulo
            CadValues2 = CadValues2 & "," & DBSet(CtaBanco, "T")
            
            If vParamAplic.ContabilidadNueva Then
                CadValues2 = CadValues2 & "," & text1csb & "," & DBSet(text2csb, "T") & ","
                CadValues2 = CadValues2 & DBSet(Rs4!Iban, "T", "S") & ", "
                CadValues2 = CadValues2 & DBSet(Rs4!Nommacta, "T", "S") & "," & DBSet(Rs4!dirdatos, "T", "S") & "," & DBSet(Rs4!desPobla, "T", "S") & ","
                CadValues2 = CadValues2 & DBSet(Rs4!Codposta, "T", "S") & "," & DBSet(Rs4!desProvi, "T", "S") & "," & DBSet(Rs4!nifdatos, "T", "S") & ",'ES')"
                
                Sql = "INSERT INTO pagos (numserie, codmacta, numfactu, fecfactu, numorden, codforpa, fecefect, impefect,fecultpa, imppagad, ctabanc1,text1csb,text2csb, iban,"
                Sql = Sql & "nomprove, domprove, pobprove, cpprove, proprove, nifprove, codpais)"
                
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
                Sql = "INSERT INTO spagop (ctaprove, numfactu, fecfactu, numorden, codforpa, fecefect, impefect,  "
                Sql = Sql & "fecultpa, imppagad, ctabanc1, ctabanc2, emitdocum, contdocu, text1csb, text2csb, entidad, "
                Sql = Sql & "oficina, cc, cuentaba, transfer" ') "
                
                '[Monica]22/11/2013: tema iban
                If vEmpresaFacSoc.HayNorma19_34Nueva = 1 Then
                   Sql = Sql & ", iban)"
                Else
                   Sql = Sql & ")"
                End If
            
            End If
        
        
        End If
        
        
        Sql = Sql & " VALUES " & CadValues2
        ConnContaFacSoc.Execute Sql

    End If

    b = True

EInsertarTesoreriaNew:
    If Err.Number <> 0 Then b = False
    InsertarEnTesoreriaFacSoc = b
End Function

Private Function InsertarLinFactSoc(cadTABLA As String, cadwhere As String, caderr As String, Optional NumRegis As Long) As Boolean
'cadWHere: selecciona un registro de scafac
'codtipom=x and numfactu=y and fecfactu=z
Dim Sql As String
Dim SQLaux As String
Dim sql2 As String
Dim Rs As ADODB.Recordset
Dim Cad As String, Aux As String
Dim i As Byte
Dim totimp As Currency, ImpLinea As Currency

    On Error GoTo EInLinea

    Sql = " SELECT factsocio.codvarie, variedad.codmacta, factsocio.baseimpo as importe, "
    Sql = Sql & " factsocio.porcreten, factsocio.impreten, factsocio.basereten, factsocio.fecfactu, "
    Sql = Sql & " factsocio.codmacta as ctasocio "
    Sql = Sql & " FROM (factsocio  "
    Sql = Sql & " inner join variedad on factsocio.codvarie=variedad.codvarie) "
    Sql = Sql & " WHERE " & cadwhere
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText

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
    Sql = ""
    Sql = NumRegis & "," & Year(Rs!fecfactu) & "," & i & ","
    Sql = Sql & DBSet(Rs!Codmacta, "T")
    Sql = Sql & "," & DBSet(ImpLinea, "N") & ","
    
    If CCoste = "" Then
        Sql = Sql & ValorNulo
    Else
        Sql = Sql & DBSet(CCoste, "T")
    End If
    
    Cad = Cad & "(" & Sql & ")"
    
    
    If DBLet(Rs!ImpReten, "N") <> 0 Then
        'linea de base del importe de retencion en positivo sobre la cuenta del socio
        i = i + 1
        Sql = ""
        Sql = NumRegis & "," & Year(Rs!fecfactu) & "," & i & ","
        Sql = Sql & DBSet(Rs!CtaSocio, "T")
        Sql = Sql & "," & DBSet(Rs!ImpReten, "N") & ","
        
        If CCoste = "" Then
            Sql = Sql & ValorNulo
        Else
            Sql = Sql & DBSet(CCoste, "T")
        End If
        
        Cad = Cad & ",(" & Sql & ")"
        
        'linea de base del importe de retencion en negativo sobre la cuenta de retencion de parametros
        i = i + 1
        ImpLinea = DBLet(Rs!ImpReten, "N") * (-1)
        Sql = ""
        Sql = NumRegis & "," & Year(Rs!fecfactu) & "," & i & ","
        Sql = Sql & DBSet(vParamAplic.CtaRetenFacSoc, "T")
        Sql = Sql & "," & DBSet(ImpLinea, "N") & ","
        
        If CCoste = "" Then
            Sql = Sql & ValorNulo
        Else
            Sql = Sql & DBSet(CCoste, "T")
        End If
        
        Cad = Cad & ",(" & Sql & ")"
    End If
    
    'Insertar en la contabilidad
    If Cad <> "" Then
        Sql = "INSERT INTO linfactprov (numregis,anofacpr,numlinea,codtbase,impbaspr,codccost) "
        Sql = Sql & " VALUES " & Cad
        ConnContaFacSoc.Execute Sql
    End If

EInLinea:
    If Err.Number <> 0 Then
        InsertarLinFactSoc = False
        caderr = Err.Description
    Else
        InsertarLinFactSoc = True
    End If
End Function


Private Function InsertarLinFactSocContaNueva(cadTABLA As String, cadwhere As String, caderr As String, Optional NumRegis As Long) As Boolean
'cadWHere: selecciona un registro de scafac
'codtipom=x and numfactu=y and fecfactu=z
Dim Sql As String
Dim SQLaux As String
Dim sql2 As String
Dim Rs As ADODB.Recordset
Dim Cad As String, Aux As String
Dim i As Byte
Dim totimp As Currency, ImpLinea As Currency

    On Error GoTo EInLinea

    Sql = " SELECT factsocio.codvarie, variedad.codmacta, factsocio.baseimpo as importe, "
    Sql = Sql & " factsocio.porcreten, factsocio.impreten, factsocio.basereten, factsocio.fecfactu, "
    Sql = Sql & " factsocio.codmacta as ctasocio, factsocio.tipoiva, factsocio.porciva, factsocio.cuotaiva "
    Sql = Sql & " FROM (factsocio  "
    Sql = Sql & " inner join variedad on factsocio.codvarie=variedad.codvarie) "
    Sql = Sql & " WHERE " & cadwhere
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText

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
    Sql = DBSet(SerieFraPro, "T") & "," & NumRegis & "," & DBSet(Rs!fecfactu, "F") & "," & Year(Rs!fecfactu) & "," & i & ","
    Sql = Sql & DBSet(Rs!Codmacta, "T") & ","
    
    If CCoste = "" Then
        Sql = Sql & ValorNulo
    Else
        Sql = Sql & DBSet(CCoste, "T")
    End If
    Sql = Sql & "," & DBSet(Rs!TipoIva, "N") & "," & DBSet(Rs!PorcIva, "N") & "," & DBSet(Rs!PorcReten, "N") & "," & DBSet(ImpLinea, "N")
    Sql = Sql & "," & DBSet(Rs!CuotaIva, "N") & "," & DBSet(Rs!ImpReten, "N") & ",1"
    
    
    Cad = Cad & "(" & Sql & ")"
    
    
    'Insertar en la contabilidad
    If Cad <> "" Then
        Sql = "INSERT INTO factpro_lineas (numserie,numregis,fecharec,anofactu,numlinea,codmacta,codccost,codigiva,porciva,porcrec,"
        Sql = Sql & "baseimpo,impoiva,imporec,aplicret) "
        Sql = Sql & " VALUES " & Cad
        ConnContaFacSoc.Execute Sql
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
Dim Sql As String

    On Error Resume Next
    Sql = "Insert into tmpErrFac(codprove,numfactu,fecfactu,error) "
    Sql = Sql & " Select *," & DBSet(Mid(MenError, 1, 200), "T") & " as error From tmpFactu "
    Sql = Sql & " WHERE " & Replace(cadwhere, "factsocio", "tmpFactu")
    conn.Execute Sql
    
    If Err.Number <> 0 Then Err.Clear
End Sub


Public Function ComprobarCCoste() As Boolean
Dim Sql As String
Dim Rs As ADODB.Recordset
Dim b As Boolean


    On Error GoTo ECCoste

    ComprobarCCoste = False
            
    Sql = "SELECT distinct concefact.codconce "
    Sql = Sql & ", concefact.codccost"
    Sql = Sql & " from ((linfact "
    Sql = Sql & " INNER JOIN tmpfactu ON linfact.codsecci=tmpfactu.codsecci and linfact.letraser=tmpfactu.numserie AND linfact.numfactu=tmpfactu.numfactu AND linfact.fecfactu=tmpfactu.fecfactu) "
    Sql = Sql & " INNER JOIN concefact on linfact.codconce = concefact.codconce) "

    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText

    b = True
    
    While Not Rs.EOF
       'comprobar que el Centro de Coste existe en la Contabilidad
       If DBLet(Rs.Fields(1).Value, "T") <> "" Then
            Sql = DevuelveDesdeBDNewFac("cabccost", "codccost", "codccost", Rs.Fields(1).Value, "T")
            If Sql = "" Then
                b = False
                Sql = "No existe el centro de coste: " & DBLet(Rs.Fields(1).Value, "T")
                Sql = Sql & " del concepto: " & DBLet(Rs.Fields(0).Value, "N")
                InsertarError Sql
            End If
       Else
            b = False
            Sql = "El concepto: " & DBLet(Rs.Fields(0).Value, "N")
            Sql = Sql & " no tiene centro de coste asociado. "
            InsertarError Sql
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

