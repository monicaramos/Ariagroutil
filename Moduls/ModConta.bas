Attribute VB_Name = "ModConta"
Option Explicit

'=============================================================================
'   MODULO PARA ACCEDER A LOS DATOS DE LA CONTABILIDAD
'=============================================================================


'=============================================================================
'==========     CUENTAS
'=============================================================================
'LAURA
Public Function PonerNombreCuenta(ByRef Txt As TextBox, Modo As Byte, Optional clien As String, Optional bd As Integer, Optional Facturas As Boolean) As String
'Obtener el nombre de una cuenta
' Facturas --> true indica que viene de una cuenta de facturas varias
'          --> false, viene de avnics, seguros o telefonia
Dim DevfrmCCtas As String
Dim Cad As String

' ### [Monica] 07/09/2006 añadida la linea siguiente condicion vParamAplic.NumeroConta = 0
' para que no saque nada si no hay contabilidad
    If Not vParamAplic Is Nothing Then
        If Not Facturas Then
            Select Case bd
                Case cConta
                    If vParamAplic.NumeroConta = 0 Then
                        PonerNombreCuenta = ""
                        Exit Function
                    End If
                Case cContaSeg
                    If vParamAplic.NumeroContaSeg = 0 Then
                        PonerNombreCuenta = ""
                        Exit Function
                    End If
                Case cContaTel
                    If vParamAplic.NumeroContaTel = 0 Then
                        PonerNombreCuenta = ""
                        Exit Function
                    End If
                Case cContaGas
                    If vParamAplic.NumeroContaGas = 0 Then
                        PonerNombreCuenta = ""
                        Exit Function
                    End If
                Case cContaFacSoc
                    If vParamAplic.NumeroContaFacSoc = 0 Then
                        PonerNombreCuenta = ""
                        Exit Function
                    End If
                Case cContaCV
                    If vParamAplic.NumeroContaCV = 0 Then
                        PonerNombreCuenta = ""
                        Exit Function
                    End If
                Case cContaCVV
                    If vParamAplic.NumeroContaCVV = 0 Then
                        PonerNombreCuenta = ""
                        Exit Function
                    End If
                    
            End Select
       Else
            If bd = 0 Then
                PonerNombreCuenta = ""
                Exit Function
            End If
       End If
    End If
    If Txt.Text = "" Then
         PonerNombreCuenta = ""
         Exit Function
    End If
    DevfrmCCtas = Txt.Text
    If CuentaCorrectaUltimoNivel(DevfrmCCtas, Cad, bd, Facturas) Then
        ' ### [Monica] 07/09/2006
        If InStr(Cad, "No existe la cuenta") > 0 Then
            Txt.Text = DevfrmCCtas
            If (Modo = 3 Or Modo = 4) And clien <> "" Then 'si insertar o modificar
                Cad = Cad & "  ¿Desea crearla?"
                If MsgBox(Cad, vbYesNo) = vbYes Then
                    If InStr(1, Txt.Tag, "ssocio") > 0 Then
                        InsertarCuentaCble DevfrmCCtas, clien, , bd
                    ElseIf InStr(1, Txt.Tag, "sprove") > 0 Then
                        InsertarCuentaCble DevfrmCCtas, "", clien, bd
                    End If
                    PonerNombreCuenta = clien
                End If
            Else
                MsgBox Cad, vbExclamation
            End If
        Else
            Txt.Text = DevfrmCCtas
            PonerNombreCuenta = Cad
        End If
    Else
        MsgBox Cad, vbExclamation
'        Txt.Text = ""
        PonerNombreCuenta = ""
'        PonerFoco Txt
    End If
    DevfrmCCtas = ""

End Function




'DAVID: Cuentas del la Contabilidad
Public Function CuentaCorrectaUltimoNivel(ByRef Cuenta As String, ByRef devuelve As String, Optional bd As Integer, Optional Facturas As Boolean) As Boolean
'Comprueba si es numerica
' Facturas = true viene de facturas varias, false viene de avnics, seguros o telefonia
    Dim SQL As String
    Dim otroCampo As String
    
    CuentaCorrectaUltimoNivel = False
    If Cuenta = "" Then
        devuelve = "Cuenta vacia"
        Exit Function
    End If

    If Not IsNumeric(Cuenta) Then
        devuelve = "La cuenta debe de ser numérica: " & Cuenta
        Exit Function
    End If

    'Rellenamos si procede
    Cuenta = RellenaCodigoCuenta(Cuenta, bd, Facturas)

    '==========
    If Not EsCuentaUltimoNivel(Cuenta, bd, Facturas) Then
        devuelve = "No es cuenta de último nivel: " & Cuenta
        Exit Function
    End If
    '==================

    otroCampo = "apudirec"
    'BD 2: conexion a BD Conta
    If Not Facturas Then
        Select Case bd
            Case cConta
                SQL = DevuelveDesdeBDNew(cConta, "cuentas", "nommacta", "codmacta", Cuenta, "T", otroCampo)
            Case cContaSeg
                SQL = DevuelveDesdeBDNew(cContaSeg, "cuentas", "nommacta", "codmacta", Cuenta, "T", otroCampo)
            Case cContaTel
                SQL = DevuelveDesdeBDNew(cContaTel, "cuentas", "nommacta", "codmacta", Cuenta, "T", otroCampo)
            Case cContaGas
                SQL = DevuelveDesdeBDNew(cContaGas, "cuentas", "nommacta", "codmacta", Cuenta, "T", otroCampo)
            Case cContaGas
                SQL = DevuelveDesdeBDNew(cContaGas, "cuentas", "nommacta", "codmacta", Cuenta, "T", otroCampo)
            Case cContaFacSoc
                SQL = DevuelveDesdeBDNew(cContaFacSoc, "cuentas", "nommacta", "codmacta", Cuenta, "T", otroCampo)
            Case cContaCV
                SQL = DevuelveDesdeBDNew(cContaCV, "cuentas", "nommacta", "codmacta", Cuenta, "T", otroCampo)
            Case cContaCVV
                SQL = DevuelveDesdeBDNew(cContaCVV, "cuentas", "nommacta", "codmacta", Cuenta, "T", otroCampo)
        End Select
    Else
        SQL = DevuelveDesdeBDNewFac("cuentas", "nommacta", "codmacta", Cuenta, "T", otroCampo)
    End If
    If SQL = "" Then
        devuelve = "No existe la cuenta : " & Cuenta
        CuentaCorrectaUltimoNivel = True
        Exit Function
    End If

    'Llegados aqui, si que existe la cuenta
    If otroCampo = "S" Then 'Si es apunte directo
        CuentaCorrectaUltimoNivel = True
        devuelve = SQL
    Else
        devuelve = "No es apunte directo: " & Cuenta
    End If
End Function


'DAVID
Public Function RellenaCodigoCuenta(vCodigo As String, bd As Integer, Optional Facturas As Boolean) As String
'Rellena con ceros hasta poner una cuenta.
'Facturas = true viene de facturas varias y false viene de avnics, seguros o telefonia
'Ejemplo: 43.1 --> 430000001
Dim i As Integer
Dim J As Integer
Dim Cont As Integer
Dim Cad As String

    RellenaCodigoCuenta = vCodigo
    If Not Facturas Then
        Select Case bd
            Case cConta
                If Len(vCodigo) > vEmpresa.DigitosUltimoNivel Then Exit Function
            Case cContaSeg
                If Len(vCodigo) > vEmpresaSeg.DigitosUltimoNivel Then Exit Function
            Case cContaTel
                If Len(vCodigo) > vEmpresaTel.DigitosUltimoNivel Then Exit Function
            Case cContaGas
                If Len(vCodigo) > vEmpresaGas.DigitosUltimoNivel Then Exit Function
            Case cContaFacSoc
                If Len(vCodigo) > vEmpresaFacSoc.DigitosUltimoNivel Then Exit Function
            Case cContaCV
                If Len(vCodigo) > vEmpresaCV.DigitosUltimoNivel Then Exit Function
            Case cContaCVV
                If Len(vCodigo) > vEmpresaCVV.DigitosUltimoNivel Then Exit Function
        End Select
    Else
        If Len(vCodigo) > vEmpresaFac.DigitosUltimoNivel Then Exit Function
    End If
    i = 0: Cont = 0
    Do
        i = i + 1
        i = InStr(i, vCodigo, ".")
        If i > 0 Then
            If Cont > 0 Then Cont = 1000
            Cont = Cont + i
        End If
    Loop Until i = 0

    'Habia mas de un punto
    If Cont > 1000 Or Cont = 0 Then Exit Function

    'Cambiamos el punto por 0's  .-Utilizo la variable maximocaracteres, para no tener k definir mas
    i = Len(vCodigo) - 1 'el punto lo quito
    If Not Facturas Then
        Select Case bd
            Case cConta
                J = vEmpresa.DigitosUltimoNivel - i
            Case cContaSeg
                J = vEmpresaSeg.DigitosUltimoNivel - i
            Case cContaTel
                J = vEmpresaTel.DigitosUltimoNivel - i
            Case cContaGas
                J = vEmpresaGas.DigitosUltimoNivel - i
            Case cContaFacSoc
                J = vEmpresaFacSoc.DigitosUltimoNivel - i
            Case cContaFacSoc
                J = vEmpresaFacSoc.DigitosUltimoNivel - i
            Case cContaCV
                J = vEmpresaCV.DigitosUltimoNivel - i
            Case cContaCVV
                J = vEmpresaCVV.DigitosUltimoNivel - i
        End Select
    Else
        J = vEmpresaFac.DigitosUltimoNivel - i
    End If
    Cad = ""
    For i = 1 To J
        Cad = Cad & "0"
    Next i

    Cad = Mid(vCodigo, 1, Cont - 1) & Cad
    Cad = Cad & Mid(vCodigo, Cont + 1)
    RellenaCodigoCuenta = Cad
End Function

'DAVID
Public Function EsCuentaUltimoNivel(Cuenta As String, Optional bd As Integer, Optional Facturas As Boolean) As Boolean
    If Not Facturas Then
        Select Case bd
            Case cConta
                EsCuentaUltimoNivel = (Len(Cuenta) = vEmpresa.DigitosUltimoNivel)
            Case cContaSeg
                EsCuentaUltimoNivel = (Len(Cuenta) = vEmpresaSeg.DigitosUltimoNivel)
            Case cContaTel
                EsCuentaUltimoNivel = (Len(Cuenta) = vEmpresaTel.DigitosUltimoNivel)
            Case cContaGas
                EsCuentaUltimoNivel = (Len(Cuenta) = vEmpresaGas.DigitosUltimoNivel)
            Case cContaFacSoc
                EsCuentaUltimoNivel = (Len(Cuenta) = vEmpresaFacSoc.DigitosUltimoNivel)
            Case cContaCV
                EsCuentaUltimoNivel = (Len(Cuenta) = vEmpresaCV.DigitosUltimoNivel)
            Case cContaCVV
                EsCuentaUltimoNivel = (Len(Cuenta) = vEmpresaCVV.DigitosUltimoNivel)
        End Select
    Else
        EsCuentaUltimoNivel = (Len(Cuenta) = vEmpresaFac.DigitosUltimoNivel)
    End If
End Function

' ### [Monica] 07/09/2006
' copia de la gestion
Private Function InsertarCuentaCble(Cuenta As String, cadSocio As String, Optional cadProve As String, Optional bd As Integer, Optional Facturas As Boolean) As Boolean
Dim SQL As String
Dim vsocio As CSocio
Dim b As Boolean

    On Error GoTo EInsCta
    
    SQL = "INSERT INTO cuentas (codmacta,nommacta,apudirec,model347,razosoci,dirdatos,codposta,despobla,desprovi,nifdatos,maidatos,webdatos,obsdatos,pais, entidad, oficina, cc, cuentaba) "
    SQL = SQL & " VALUES (" & DBSet(Cuenta, "T") & ","
    
    If cadSocio <> "" Then
        Set vsocio = New CSocio
        If vsocio.LeerDatos(cadSocio) Then
            SQL = SQL & DBSet(vsocio.Nombre, "T") & ",'S',1," & DBSet(Cuenta, "T") & "," & DBSet(vsocio.Domicilio, "T") & ","
            SQL = SQL & DBSet(vsocio.CPostal, "T") & "," & DBSet(vsocio.Poblacion, "T") & "," & DBSet(vsocio.Provincia, "T") & "," & DBSet(vsocio.nif, "T") & "," & DBSet(vsocio.EMailAdm, "T") & "," & DBSet(vsocio.Websocio, "T") & "," & ValorNulo & "," & ValorNulo & "," & DBSet(vsocio.Banco, "N") & "," & DBSet(vsocio.Sucursal, "N") & "," & DBSet(vsocio.Digcontrol, "T") & "," & DBSet(vsocio.CuentaBan, "T") & ")"
            If Not Facturas Then
                Select Case bd
                    Case cConta
                        ConnConta.Execute SQL
                    Case cContaSeg
                        ConnContaSeg.Execute SQL
                    Case cContaTel
                        ConnContaTel.Execute SQL
                    Case cContaGas
                        ConnContaGas.Execute SQL
                    Case cContaFacSoc
                        ConnContaFacSoc.Execute SQL
                    Case cContaCV
                        ConnContaCV.Execute SQL
                    Case cContaCVV
                        ConnContaCVV.Execute SQL
                        
                End Select
            Else
                ConnContaFac.Execute SQL
            End If
            cadSocio = vsocio.Nombre
            b = True
        Else
            b = False
        End If
        Set vsocio = Nothing
    End If
    
    
EInsCta:
    If Err.Number <> 0 Then
        b = False
        MuestraError Err.Description, "Insertando cuenta contable", Err.Description
    End If
    InsertarCuentaCble = b
End Function


'=============================================================================
'==========     CENTROS DE COSTE
'=============================================================================
'LAURA
Public Function PonerNombreCCoste(Empresa As String, ByRef Txt As TextBox) As String
'Obtener el nombre de un centro de coste
Dim codCCoste As String
Dim Cad As String

     If Txt.Text = "" Then
         PonerNombreCCoste = ""
         Exit Function
    End If
    codCCoste = Txt.Text
    If CCosteCorrecto(Empresa, codCCoste, Cad) Then
        Txt.Text = codCCoste
        PonerNombreCCoste = Cad
    Else
        MsgBox Cad, vbExclamation
'        Txt.Text = ""
        PonerNombreCCoste = ""
        PonerFoco Txt
    End If
'    codCCoste = ""
End Function

'LAURA
Public Function CCosteCorrecto(Empresa As String, ByRef Centro As String, ByRef devuelve As String) As Boolean
    Dim SQL As String
    
    CCosteCorrecto = False
 
    'BD 2: conexion a BD Conta
    If Val(Empresa) <> Val(vEmpresa.codEmpre) Then
        SQL = DevuelveDesdeBDNew(3, "cabccost", "nomccost", "codccost", Centro, "T")
    Else
        SQL = DevuelveDesdeBDNew(cConta, "cabccost", "nomccost", "codccost", Centro, "T")
    End If
    If SQL = "" Then
        devuelve = "No existe el Centro de coste : " & Centro
        Exit Function
    Else
        devuelve = SQL
        CCosteCorrecto = True
    End If
End Function




'=============================================================================
'==========     CONCEPTOS
'=============================================================================
'LAURA
Public Function PonerNombreConcepto(ByRef Txt As TextBox, Conec As Byte, Optional Facturas As Boolean) As String
'Obtener el nombre de un concepto
Dim codConce As String
Dim Cad As String

     If Txt.Text = "" Then
         PonerNombreConcepto = ""
         Exit Function
    End If
    codConce = Txt.Text
    If ConceptoCorrecto(codConce, Cad, Conec, Facturas) Then
        Txt.Text = Format(codConce, "000")
        PonerNombreConcepto = Cad
    Else
        MsgBox Cad, vbExclamation
        Txt.Text = ""
        PonerNombreConcepto = ""
        PonerFoco Txt
    End If
End Function


'LAURA
Public Function ConceptoCorrecto(ByRef Concep As String, ByRef devuelve As String, Conec As Byte, Optional Facturas As Boolean) As Boolean
' Facturas = true --> viene de facturas varias
'          = false --> viene de avnics, seguros o telefonia
    Dim SQL As String
    
    ConceptoCorrecto = False
 
    'BD 2: conexion a BD Conta
    If Not Facturas Then
        SQL = DevuelveDesdeBDNew(Conec, "conceptos", "nomconce", "codconce", Concep, "N")
    Else
        SQL = DevuelveDesdeBDNewFac("conceptos", "nomconce", "codconce", Concep, "N")
    End If
    If SQL = "" Then
        devuelve = "No existe el concepto : " & Concep
        Exit Function
    Else
        devuelve = SQL
        ConceptoCorrecto = True
    End If
End Function

' ### [Monica] 27/09/2006
Public Function FacturaContabilizada(numserie As String, numfactu As String, Anofactu As String) As Boolean
Dim SQL As String
Dim NumAsi As Currency

    FacturaContabilizada = False
    SQL = ""
    SQL = DevuelveDesdeBDNew(cConta, "cabfact", "numasien", "numserie", Trim(numserie), "T", , "codfaccl", numfactu, "N", "anofaccl", Anofactu, "N")
    
    If SQL = "" Then Exit Function
    
    NumAsi = DBLet(SQL, "N")
    
    If NumAsi <> 0 Then FacturaContabilizada = True

End Function

' ### [Monica] 27/09/2006
Public Function FacturaRemesada(numserie As String, numfactu As String, fecfactu As String) As Boolean
Dim SQL As String
Dim NumRem As Currency

    FacturaRemesada = False
    
    SQL = ""
    If vParamAplic.ContabilidadNueva Then
        SQL = DevuelveDesdeBDNew(cConta, "cobros", "codrem", "numserie", Trim(numserie), "T", , "numfactu", numfactu, "N", "fecfactu", fecfactu, "F")
    Else
        SQL = DevuelveDesdeBDNew(cConta, "scobro", "codrem", "numserie", Trim(numserie), "T", , "codfaccl", numfactu, "N", "fecfaccl", fecfactu, "F")
    End If
    
    If SQL = "" Then Exit Function
    
    NumRem = DBLet(SQL, "N")
    
    If NumRem <> 0 Then FacturaRemesada = True
    
End Function

' ### [Monica] 27/09/2006
Public Function FacturaCobrada(numserie As String, numfactu As String, fecfactu As String) As Boolean
Dim SQL As String
Dim ImpCob As Currency

    FacturaCobrada = False
    SQL = ""
    If vParamAplic.ContabilidadNueva Then
        SQL = DevuelveDesdeBDNew(cConta, "cobros", "impcobro", "numserie", Trim(numserie), "T", , "numfactu", numfactu, "N", "fecfactu", fecfactu, "F")
    Else
        SQL = DevuelveDesdeBDNew(cConta, "scobro", "impcobro", "numserie", Trim(numserie), "T", , "codfaccl", numfactu, "N", "fecfaccl", fecfactu, "F")
    End If
    If SQL = "" Then Exit Function
    ImpCob = DBLet(SQL, "N")
    
    If ImpCob <> 0 Then FacturaCobrada = True
    
End Function

' ### [Monica] 27/09/2006
Public Function ModificaCtaClienteFacturaContabilidad(letraser As String, numfactu As String, fecfactu As String, CtaConta As String) As Boolean
Dim SQL As String
Dim Anyo As Currency

    On Error GoTo eModificaCtaClienteFacturaContabilidad

    ModificaCtaClienteFacturaContabilidad = False

    Anyo = Year(CDate(fecfactu))
    
    If vParamAplic.ContabilidadNueva Then
        SQL = "update factcli set codmacta = " & DBSet(CtaConta, "T") & " where numserie = " & DBSet(Trim(letraser), "T") & " and " & _
                  "numfactu = " & DBSet(numfactu, "N") & " and anofactu = " & DBSet(Anyo, "N")
        ConnContaFac.Execute SQL
        
        SQL = "update cobros set codmacta = " & DBSet(CtaConta, "T") & " where numserie = " & DBSet(Trim(letraser), "T") & " and " & _
                  "numfactu = " & DBSet(numfactu, "N") & " and fecfactu = " & DBSet(fecfactu, "F")
                  
        ConnContaFac.Execute SQL
    Else
        SQL = "update cabfact set codmacta = " & DBSet(CtaConta, "T") & " where numserie = " & DBSet(Trim(letraser), "T") & " and " & _
                  "codfaccl = " & DBSet(numfactu, "N") & " and anofaccl = " & DBSet(Anyo, "N")
        ConnContaFac.Execute SQL
        
        SQL = "update scobro set codmacta = " & DBSet(CtaConta, "T") & " where numserie = " & DBSet(Trim(letraser), "T") & " and " & _
                  "codfaccl = " & DBSet(numfactu, "N") & " and fecfaccl = " & DBSet(fecfactu, "F")
                  
        ConnContaFac.Execute SQL
    End If
    ModificaCtaClienteFacturaContabilidad = True
    
eModificaCtaClienteFacturaContabilidad:
    If Err.Number <> 0 Then
        MsgBox "Error en ModificaCtaClienteFacturaContabilidad: " & Err.Description, vbExclamation
    End If

End Function

' ### [Monica] 27/09/2006
Public Sub ModificaFormaPagoTesoreria(letraser As String, numfactu As String, fecfactu As String, forpa As String, forpaant As String)
Dim SQL As String
Dim SQL1 As String
Dim TipForpa As String
Dim TipForpaAnt As String
Dim cadwhere As String

    If vParamAplic.ContabilidadNueva Then
        cadwhere = " numserie = " & DBSet(Trim(letraser), "T") & " and " & _
                  "numfactu = " & numfactu & " and fecfactu = " & DBSet(fecfactu, "F")
        
        SQL = "update scobro set codforpa = " & forpa & " where " & cadwhere
    Else
        cadwhere = " numserie = " & DBSet(Trim(letraser), "T") & " and " & _
                  "codfaccl = " & numfactu & " and fecfaccl = " & DBSet(fecfactu, "F")
        
        SQL = "update scobro set codforpa = " & forpa & " where " & cadwhere
    End If
    ConnConta.Execute SQL

End Sub

'' ### [Monica] 29/09/2006
Public Function ModificaImportesFacturaContabilidad(letraser As String, numfactu As String, fecfactu As String, Importe As String, forpa As String, vTabla As String) As Boolean
Dim SQL As String
Dim vWhere As String
Dim b As Boolean
Dim CadValues As String
Dim vsocio As CSocio
Dim Rs As ADODB.Recordset
Dim TipForpa As String

'    On Error GoTo eModificaImportesFacturaContabilidad
'
'    b = False
'
'    vWhere = "numserie = " & DBSet(letraser, "T") & " and codfaccl = " & _
'              numfactu & " and anofaccl = " & Format(Year(fecfactu), "0000")
'
'
'    sql = "select codsocio from " & vTabla & " where letraser = " & DBSet(letraser, "T") & " and numfactu = " & _
'           numfactu & " and fecfactu = " & DBSet(fecfactu, "F")
'
'    Set RS = New adodb.Recordset
'    RS.Open sql, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
'    If Not RS.EOF Then RS.MoveFirst
'
'    Set vsocio = New CSocio
'    If vsocio.LeerDatos(RS.Fields(0).Value) Then
'    '********************************+estoy aqui
'
'        If vTabla = "schfac" Then
'            sql = "delete from linfact where " & vWhere
'            ConnConta.Execute sql
'
'            sql = "delete from cabfact where " & vWhere
'            ConnConta.Execute sql
'
'            sql = "schfac.letraser = " & DBSet(letraser, "T") & " and numfactu = " & numfactu
'            sql = sql & " and fecfactu = " & DBSet(fecfactu, "F")
'
'
'            b = CrearTMPErrFact("schfac")
'            If b Then b = PasarFactura2(sql, vsocio)
'        Else
'            b = CrearTMPErrFact("schfacr")
'        End If
'
'        ' 09/02/2007
'        TipForpa = DevuelveDesdeBDNew(cPTours, "sforpa", "tipforpa", "codforpa", forpa, "N")
'        If TipForpa <> "0" And b Then
'            b = ModificaCobroTesoreria(letraser, numfactu, fecfactu, vsocio, vTabla)
'        End If
'    End If
'
'    ModificaImportesFacturaContabilidad = b
'
eModificaImportesFacturaContabilidad:
    If Err.Number <> 0 Then
        MsgBox "Error en ModificaImportesFacturaContabilidad: " & Err.Description, vbExclamation
    End If
End Function


Public Function CalcularIva(Importe As String, articulo As String) As Currency
'devuelve el iva del Importe
'Ej el 16% de 120 = 19.2
Dim vImp As Currency
Dim vIva As Currency
Dim vArt As Currency
Dim CodIVA As String

Dim IvaArt As Integer
Dim Iva As String
Dim impiva As Currency
On Error Resume Next

    Importe = ComprobarCero(Importe)
    articulo = ComprobarCero(articulo)
    
    CodIVA = DevuelveDesdeBD("codigiva", "sartic", "codartic", articulo, "N")
    Iva = DevuelveDesdeBDNew(cConta, "tiposiva", "porceiva", "codigiva", CodIVA, "N")
    
    vImp = CCur(Importe)
    vIva = CCur(Iva)
    
    impiva = ((vImp * vIva) / 100)
    impiva = Round(impiva, 2)
    
    CalcularIva = CStr(impiva)
    If Err.Number <> 0 Then Err.Clear

End Function


Public Function CalcularBase(Importe As String, articulo As String) As Currency
'devuelve la base del Importe
'Ej el 16% de 120 = 120-19.2 = 100.8
Dim vImp As Currency
Dim vIva As Currency
Dim vArt As Currency
Dim CodIVA As String

Dim IvaArt As Integer
Dim Iva As String
Dim impiva As Currency
On Error Resume Next

    Importe = ComprobarCero(Importe)
    articulo = ComprobarCero(articulo)
    
    CodIVA = DevuelveDesdeBD("codigiva", "sartic", "codartic", articulo, "N")
    Iva = DevuelveDesdeBDNew(cConta, "tiposiva", "porceiva", "codigiva", CodIVA, "N")
    
    vImp = CCur(Importe)
    vIva = CCur(Iva)
    
    impiva = Round2(Importe / (1 + (vIva / 100)), 2)
    
    CalcularBase = CStr(impiva)
    If Err.Number <> 0 Then Err.Clear

End Function

Public Function CtaContableSocio(nif As String, bd As Byte) As String
Dim SQL As String
Dim Rs As ADODB.Recordset
Dim NumDigi As Byte
Dim Encontrado As Boolean

    CtaContableSocio = ""
    
    SQL = "select codmacta from cuentas where nifdatos = " & DBSet(nif, "T")
    Set Rs = New ADODB.Recordset
    
    Select Case bd
        Case cContaSeg
            Rs.Open SQL, ConnContaSeg, adOpenForwardOnly, adLockOptimistic, adCmdText
            NumDigi = Len(vParamAplic.RaizCtaSocSeg)
            
        Case cContaTel
            Rs.Open SQL, ConnContaTel, adOpenForwardOnly, adLockOptimistic, adCmdText
            NumDigi = Len(vParamAplic.RaizCtaSocTel)
    End Select
    
    Encontrado = False
    While Not Rs.EOF And Not Encontrado
        Select Case bd
            Case cContaSeg
                If Mid(Rs.Fields(0).Value, 1, NumDigi) = vParamAplic.RaizCtaSocSeg Then
                    Encontrado = True
                    CtaContableSocio = Rs.Fields(0).Value
                End If
            Case cContaTel
                If Mid(Rs.Fields(0).Value, 1, NumDigi) = vParamAplic.RaizCtaSocTel Then
                    Encontrado = True
                    CtaContableSocio = Rs.Fields(0).Value
                End If
        End Select
        
        Rs.MoveNext
    Wend
    
    Set Rs = Nothing
    
End Function


'=============================================================================
'==========     FORMA DE PAGO
'=============================================================================
Public Function PonerNombreFPago(ByRef Txt As TextBox, Conec As Byte) As String
'Obtener el nombre de un concepto
Dim codFPago As String
Dim Cad As String

     If Txt.Text = "" Then
         PonerNombreFPago = ""
         Exit Function
    End If
    codFPago = Txt.Text
    If FPagoCorrecto(codFPago, Cad, Conec) Then
        Txt.Text = Format(codFPago, "000")
        PonerNombreFPago = Cad
    Else
        MsgBox Cad, vbExclamation
        Txt.Text = ""
        PonerNombreFPago = ""
        PonerFoco Txt
    End If
End Function


Public Function FPagoCorrecto(ByRef FPago As String, ByRef devuelve As String, Conec As Byte) As Boolean
    Dim SQL As String
    
    FPagoCorrecto = False
 
    'BD 2: conexion a BD Conta
    If vParamAplic.ContabilidadNueva Then
        SQL = DevuelveDesdeBDNew(Conec, "formapago", "nomforpa", "codforpa", FPago, "N")
    Else
        SQL = DevuelveDesdeBDNew(Conec, "sforpa", "nomforpa", "codforpa", FPago, "N")
    End If
    If SQL = "" Then
        devuelve = "No existe la Forma de Pago : " & FPago
        Exit Function
    Else
        devuelve = SQL
        FPagoCorrecto = True
    End If
End Function


'=============================================================================
'==========     TIPO DE IVA
'=============================================================================
Public Function PonerNombreTIva(ByRef Txt As TextBox, Conec As Byte) As String
'Obtener el nombre de un tipo de iva
Dim codTIva As String
Dim Cad As String

     If Txt.Text = "" Then
         PonerNombreTIva = ""
         Exit Function
    End If
    codTIva = Txt.Text
    If TIvaCorrecto(codTIva, Cad, Conec) Then
        Txt.Text = Format(codTIva, "0")
        PonerNombreTIva = Cad
    Else
        MsgBox Cad, vbExclamation
        Txt.Text = ""
        PonerNombreTIva = ""
        PonerFoco Txt
    End If
End Function


Public Function TIvaCorrecto(ByRef TIva As String, ByRef devuelve As String, Conec As Byte) As Boolean
    Dim SQL As String
    
    TIvaCorrecto = False
 
    'BD 2: conexion a BD Conta
    SQL = DevuelveDesdeBDNew(Conec, "tiposiva", "nombriva", "codigiva", TIva, "N")
    If SQL = "" Then
        devuelve = "No existe el Tipo de Iva : " & TIva
        Exit Function
    Else
        devuelve = SQL
        TIvaCorrecto = True
    End If
End Function


