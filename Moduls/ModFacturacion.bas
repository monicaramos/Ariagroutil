Attribute VB_Name = "ModFacturacion"
' Modulo en donde se encuentran los procedimintos para la facturacion
'
'Dim db As BaseDatos
Dim RS As ADODB.Recordset
Dim ImpFactu As Currency
Dim TotalImp As Currency
Dim numser As String
Dim dc As Dictionary
Dim baseimpo As Dictionary


Public Function TraspasoHistoricoFacturas(db As BaseDatos, Sql As String, desde As String, hasta As String, ByRef Pb1 As ProgressBar) As Boolean
    
    Dim importel As Currency
    Dim impbase As Currency
    Dim actFactura As Long
    Dim antfactura As Long
    Dim AntFecha As Date
    Dim AntSocio As Long
    Dim AntForpa As Integer
    Dim HayReg As Boolean
    
    Dim SQL1 As String
    
    Dim NumError As Long

    On Error GoTo eTraspasoHistoricoFacturas
    
'    Set db = New BaseDatos
'
'    db.abrir "arigasol", "root", "aritel"
'    db.Tipo = "MYSQL"
        
    Set baseimpo = New Dictionary
      
    ' abrimos la transaccion
    db.AbrirTrans
      
      ' traemos el numero de serie de la factura de tipo FAC de la tabla stipom
      numser = ""
      numser = DevuelveDesdeBD("letraser", "stipom", "codtipom", "FAT", "T")
      
      TotalImp = 0
      Set RS = db.cursor(Sql)
      HayReg = False
      NumError = 0
      If Not RS.EOF Then
          RS.MoveFirst
          antfactura = RS!numfactu
          AntFecha = RS!FecAlbar
          AntSocio = RS!Codsocio
          AntForpa = RS!CodForpa
          
          While Not RS.EOF And NumError = 0
              actFactura = RS!numfactu
              HayReg = True
              If actFactura <> antfactura Then ' after group of numfactu
                 If NumError = 0 Then NumError = InsertCabe(db, baseimpo, antfactura, AntFecha, AntSocio, AntForpa, 0)
                 Set baseimpo = Nothing
                 Set baseimpo = New Dictionary
                 TotalImp = 0
                 antfactura = actFactura
                 AntFecha = RS!FecAlbar
                 AntSocio = RS!Codsocio
                 AntForpa = RS!CodForpa
              End If
              '-------
              ' tenemos que calcular el impuesto multiplicando cantidad de linea por impuesto por articulo
              importel = DBLet(RS!impuesto, "N") ' Comprueba si es nulo y lo pone a 0 o ""
              If EsArticuloCombustible(RS!codartic) Then
                TotalImp = TotalImp + Round2((RS!cantidad * importel), 2)
              End If
              baseimpo(Val(RS!CodigIva)) = DBLet(baseimpo(Val(RS!CodigIva)), "N") + DBLet(RS!importel, "N")
              
              If NumError = 0 Then NumError = InsertLinea(db, RS)
              
              If NumError = 0 Then
                    Pb1.Value = Pb1.Value + 1
                    Pb1.Refresh
                    
                    RS.MoveNext
              End If
          Wend
          If HayReg And NumError = 0 Then NumError = InsertCabe(db, baseimpo, actFactura, AntFecha, AntSocio, AntForpa, 0)


          ' hacemos el borrado masivo de albaranes de las los albaranes
'          If NumError = 0 Then NumError = BorradoAlbaranes(db, desde, hasta)

          ' aprovechamos para borrar todas las pruebas de manguera
'          If NumError = 0 Then NumError = BorradoAlbaranesPrueba(db, desde, hasta)

        End If
    Set RS = Nothing
    If NumError <> 0 Then Err.Raise NumError
        
eTraspasoHistoricoFacturas:
    If Err.Number <> 0 Then
        MuestraError Err.Number, "Error en el traspaso al histórico. Llame a soporte." & vbCrLf & vbCrLf & MensError
        db.RollbackTrans
        TraspasoHistoricoFacturas = False
        Pb1.visible = False
    Else
        db.CommitTrans
        TraspasoHistoricoFacturas = True
    End If
End Function

'Insertar Cabecera de factura
Public Function InsertCabe(ByRef db As BaseDatos, ByRef dc As Dictionary, numfactu As Long, fecha As Date, socio As Long, forpa As Integer, tipo As Byte) As Long   ', db As Database)
' tipo 0 en la schfac
' tipo 1 en la schfacr

    Dim i As Integer
    Dim Imptot(2)
    Dim Tipiva(2)
    Dim Impbas(2)
    Dim impiva(2)
    Dim PorIva(2)
    Dim TotFac
    Dim Sql As String
    Dim NumCoop As String
    
    '05012007
    On Error GoTo eInsertCabe
    MensError = ""
    ' inicializamos los importes de los totales de la cabecera
    TotFac = 0
    For i = 0 To 2
         Tipiva(i) = Null
         Imptot(i) = Null
         Impbas(i) = Null
         impiva(i) = Null
         PorIva(i) = Null
    Next i
    
    For i = 0 To dc.Count - 1
        If i <= 2 Then
            Tipiva(i) = dc.Keys(i)
            Imptot(i) = dc.Items(i)
            PorIva(i) = DevuelveDesdeBDNew(cConta, "tiposiva", "porceiva", "codigiva", CStr(Tipiva(i)), "N")
            Impbas(i) = Round2(Imptot(i) / (1 + (PorIva(i) / 100)), 2)
            impiva(i) = Imptot(i) - Impbas(i)
            TotFac = TotFac + Imptot(i)
        End If
    Next i
'    TotFac = TotFac - totalimp
    
    NumCoop = DevuelveDesdeBD("codcoope", "ssocio", "codsocio", CStr(socio), "N")
    
    If tipo = 0 Then
        Sql = "INSERT into schfac "
    Else
        Sql = "INSERT into schfacr "
    End If

    Sql = Sql & "(letraser, numfactu, fecfactu, codsocio, codcoope, " & _
           "codforpa, baseimp1, baseimp2, baseimp3, impoiva1, " & _
           "impoiva2, impoiva3, tipoiva1, tipoiva2, tipoiva3, " & _
           "porciva1, porciva2, porciva3, totalfac, impuesto, " & _
           "intconta)" & _
           "values " & _
           "(" & db.Texto(numser) & "," & db.numero(numfactu) & "," & db.fecha(fecha) & "," & db.numero(socio) & "," & db.numero(NumCoop) & "," & _
           db.numero(forpa) & "," & db.numero(Impbas(0)) & "," & db.numero(Impbas(1)) & "," & db.numero(Impbas(2)) & "," & db.numero(impiva(0)) & "," & _
           db.numero(impiva(1)) & "," & db.numero(impiva(2)) & "," & db.numero(Tipiva(0)) & "," & db.numero(Tipiva(1)) & "," & db.numero(Tipiva(2)) & "," & _
           db.numero(PorIva(0)) & "," & db.numero(PorIva(1)) & "," & db.numero(PorIva(2)) & "," & db.numero(TotFac) & "," & db.numero(TotalImp) & "," & _
           "0" & ")"
    InsertCabe = db.Ejecutar(Sql)

eInsertCabe:
    If Err.Number <> 0 Or InsertCabe <> 0 Then
        MensError = "Error en la inserción en schfac de la factura " & numfactu & " del socio " & socio
        If InsertCabe = 0 Then InsertCabe = 1
    End If
End Function

'Insertar Linea de factura
Public Function InsertLinea(db As BaseDatos, RS As ADODB.Recordset) As Long  ' , db As Database) As Boolean

    Dim Sql As String
    Dim Implinea As Currency
    
'05012007
On Error GoTo eInsertLinea
    MensError = ""
    
        Sql = "INSERT into slhfac " & _
           "(letraser, numfactu, fecfactu, numlinea, numalbar, " & _
           "fecalbar, horalbar, codturno, numtarje, codartic, " & _
           "cantidad, preciove, implinea) " & _
           "values " & _
           "(" & db.Texto(numser) & "," & db.numero(RS!numfactu) & "," & db.fecha(RS!FecAlbar) & "," & db.numero(RS!NumLinea) & "," & db.Texto(RS!numalbar) & "," & _
           db.fecha(RS!FecAlbar) & "," & db.fechahora(RS!FecAlbar & " " & Format(RS!horalbar, "hh:mm:ss")) & "," & db.numero(RS!codTurno) & "," & db.numero(RS!Numtarje) & "," & db.numero(RS!codartic) & "," & _
           db.numero(RS!cantidad) & "," & db.numero(RS!PrecioVe) & "," & db.numero(RS!importel) & ")"
    InsertLinea = db.Ejecutar(Sql)
    
eInsertLinea:
    If Err.Number <> 0 Or InsertLinea <> 0 Then
        MensError = "Se ha producido un error en la inserción de la linea de factura correspondiente al albaran " & RS!numalbar
        If InsertLinea = 0 Then InsertLinea = 1
    End If
    
End Function


Public Function ExisteEnHistorico(cDesde As String, cHasta As String, ctipo As String) As Boolean
Dim Sql As String
Dim tipo As String

    ExisteEnHistorico = False
    
    Sql = "select count(*) from slhfac, scaalb where letraser = '" & DBSet(Trim(tipo), "T") & "' and " & _
          " slhfac.numfactu= scaalb.numfactu and slhfac.numlinea = scaalb.numlinea "
    
    If cDesde <> "" Then Sql = Sql & " and scaalb.fecalbar >= '" & Format(cDesde, FormatoFecha) & "' "
    If cHasta <> "" Then Sql = Sql & " and scaalb.fecalbar <= '" & Format(cHasta, FormatoFecha) & "' "

    ExisteEnHistorico = (TotalRegistros(Sql) <> 0)
    
End Function


Public Sub RecalculoBasesIvaFactura(ByRef RS As ADODB.Recordset, ByRef Imptot As Variant, ByRef Tipiva As Variant, ByRef Impbas As Variant, ByRef impiva As Variant, ByRef PorIva As Variant, ByRef TotFac As Currency, ByRef ImpRec As Variant, ByRef PorRec As Variant, ByRef PorRet As Variant, ByRef ImpRet As Variant)

    Dim i As Integer
    Dim Sql As String
    Dim baseimpo As Dictionary
    Dim CodIVA As Integer

    Set baseimpo = New Dictionary

    ' inicializamos los importes de los totales de la cabecera
    TotFac = 0
    totimp = 0
    Base = 0
    ImpRet = 0
    For i = 0 To 2
         Tipiva(i) = 0
         Imptot(i) = 0
         Impbas(i) = 0
         impiva(i) = 0
         PorIva(i) = 0
         PorRec(i) = 0
         ImpRec(i) = 0
    Next i

    ' recorremos todas las lineas de la factura
    If Not RS.EOF Then RS.MoveFirst
    While Not RS.EOF
        CodIVA = DBLet(RS!TipoIva, "N") ' DevuelveDesdeBDNewFac("tiposiva", "codigiva", "sartic", "codartic", DBLet(RS!codartic), "N")
        baseimpo(Val(CodIVA)) = DBLet(baseimpo(Val(CodIVA)), "N") + DBLet(RS!Importe, "N")

        RS.MoveNext
    Wend

    For i = 0 To baseimpo.Count - 1
        If i <= 2 Then
            Tipiva(i) = baseimpo.Keys(i)
            Impbas(i) = baseimpo.Items(i)
 
            PorIva(i) = DevuelveDesdeBDNewFac("tiposiva", "porceiva", "codigiva", CStr(Tipiva(i)), "N")
            PorRec(i) = DevuelveDesdeBDNewFac("tiposiva", "porcerec", "codigiva", CStr(Tipiva(i)), "N")
            impiva(i) = DBLet(Round2(Impbas(i) * PorIva(i) / 100, 2), "N")
            ImpRec(i) = DBLet(Round2(Impbas(i) * PorRec(i) / 100, 2), "N")
            Imptot(i) = Impbas(i) + impiva(i) + ImpRec(i)
            TotFac = TotFac + Imptot(i)
 
'antes el iva estaba incluido
'            PorIva(i) = DevuelveDesdeBDNewFac(cConta, "tiposiva", "porceiva", "codigiva", CStr(Tipiva(i)), "N")
'            Impbas(i) = Round2(Imptot(i) / (1 + (PorIva(i) / 100)), 2)
'            impiva(i) = Imptot(i) - Impbas(i)
'            TotFac = TotFac + Imptot(i)
        
        
        End If
    Next i
    'si hay retencion la calculamos
    If PorRet <> 0 Then
        Base = 0
        For i = 0 To baseimpo.Count - 1
            Base = Base + Impbas(i)
        Next i
        ImpRet = Round2(Base * PorRet / 100, 2)
        TotFac = TotFac - ImpRet
    Else
        ImpRet = 0
    End If
End Sub

Public Function InsertaLineaFactura(ByRef db As BaseDatos, RS As ADODB.Recordset, numser As String, numfac As Long, fecfac As Date, Linea As Integer, tipo As Byte) As Long
' tipo = 0 facturacion
' tipo = 1 facturacion ajena

    Dim Sql As String
    Dim Implinea As Currency
    
    On Error GoTo eInsertaLineaFactura
    MensError = ""
    
    If tipo = 0 Then
        Sql = "INSERT into slhfac "
    Else
        Sql = "INSERT into slhfacr "
    End If
     
     Sql = Sql & "(letraser, numfactu, fecfactu, numlinea, numalbar, " & _
           "fecalbar, horalbar, codturno, numtarje, codartic, " & _
           "cantidad, preciove, implinea) " & _
           "values " & _
           "(" & db.Texto(numser) & "," & db.numero(numfac) & "," & db.fecha(fecfac) & "," & db.numero(Linea) & "," & db.Texto(RS!numalbar) & "," & _
           db.fecha(RS!FecAlbar) & "," & db.fechahora(RS!FecAlbar & " " & Format(RS!horalbar, "hh:mm:ss")) & "," & db.numero(RS!codTurno) & "," & db.numero(RS!Numtarje) & "," & db.numero(RS!codartic) & "," & _
           db.numero(RS!cantidad) & "," & db.numero(RS!PrecioVe) & "," & db.numero(RS!importel) & ")"
           
    InsertaLineaFactura = db.Ejecutar(Sql)

eInsertaLineaFactura:
    If Err.Number <> 0 Or InsertaLineaFactura <> 0 Then
        MensError = "Error en la inserción de la línea de factura en el albaran " & RS!numalbar
        If InsertaLineaFactura = 0 Then InsertaLineaFactura = 1
    End If
    
End Function

' en la facturacion ajena hemos de insertar en la temporal para luego hacer la factura global
Public Function InsertaLineaFacturaTemporal(ByRef db As BaseDatos, codartic As String, cantidad As String, importel As String) As Long
' importe1 = codartic
' importe2 = cantidad
' importe3 = importel

    Dim Sql As String
    Dim Implinea As Currency
    
    On Error GoTo eInsertaLineaFacturaTemporal
    MensError = ""
    
    Sql = "select count(*) from tmpinformes where importe1 = " & db.numero(codartic) & " and codusu = " & vSesion.Codigo
    
    If TotalRegistros(Sql) <> 0 Then
        Sql = "update tmpinformes set importe2 = importe2 + " & db.numero(cantidad) & ","
        Sql = Sql & "importe3 = importe3 + " & db.numero(importel)
        Sql = Sql & " where codusu = " & vSesion.Codigo & " and importe1 = " & db.numero(codartic)
    Else
        Sql = "insert into tmpinformes (codusu, importe1, importe2, importe3) values ("
        Sql = Sql & vSesion.Codigo & "," & db.numero(codartic) & "," & db.numero(cantidad) & ","
        Sql = Sql & db.numero(importel) & ")"
        
    End If
           
    InsertaLineaFacturaTemporal = db.Ejecutar(Sql)

eInsertaLineaFacturaTemporal:
    If Err.Number <> 0 Or InsertaLineaFacturaTemporal <> 0 Then
        MensError = "Error en la inserción en temporal de la línea de albaran " & RS!numalbar
        If InsertaLineaFacturaTemporal = 0 Then InsertaLineaFacturaTemporal = 1
    End If
    
End Function

Public Function FechaSuperiorUltimaLiquidacion(fec As Date) As Boolean
Dim Sql As String
Dim RS As ADODB.Recordset
Dim Mensual As Boolean
Dim Anofactu As Integer
Dim PeriodoFactu As Integer
Dim FechaDesde As Date

    On Error GoTo eFechaSuperiorUltimaLiquidacion

    FechaSuperiorUltimaLiquidacion = False

    ' en caso de que haya contabilidad comprobamos que la fecha de factura introducida
    ' no sea inferior a la ultima liquidacion de iva.
    If vParamAplic.NumeroConta <> 0 Then
        Sql = "select periodos, anofactu, perfactu from parametros"
        Set RS = New ADODB.Recordset
        RS.Open Sql, ConnConta, adOpenDynamic, adLockOptimistic
        
        If Not RS.EOF Then
            Mensual = (RS.Fields(0).Value = 1)
            Anofactu = RS.Fields(1).Value
            PeriodoFactu = RS.Fields(2).Value
            
            If Mensual Then ' facturacion mensual
                If PeriodoFactu = 12 Then
                    FechaDesde = CDate("01/01/" & Format(Anofactu + 1, "0000"))
                Else
                    FechaDesde = CDate("01/" & Format(PeriodoFactu + 1, "00") & "/" & Format(Anofactu, "0000"))
                End If
            Else ' facturacion trimestral
                If PeriodoFactu = 4 Then
                    FechaDesde = CDate("01/01/" & Format(Anofactu + 1, "0000"))
                Else
                    FechaDesde = CDate("01/" & Format((PeriodoFactu * 3) + 1, "00") & "/" & Format(Anofactu, "0000"))
                End If
            End If
            
            FechaSuperiorUltimaLiquidacion = (fec >= FechaDesde)
        End If
    End If

eFechaSuperiorUltimaLiquidacion:
    If Err.Number <> 0 Then
         MuestraError Err.Number, Err.Description
    End If
End Function


Public Function FechaDentroPeriodoContable(fec As Date) As Boolean
Dim Sql As String
Dim RS As ADODB.Recordset
Dim Mensual As Boolean
Dim Anofactu As Integer
Dim PeriodoFactu As Integer
Dim FechaDesde As Date

    On Error GoTo eFechaDentroPeriodoContable

    FechaDentroPeriodoContable = (CDate(FIni) <= fec) And (fec <= (CDate(FFin) + 365))

eFechaDentroPeriodoContable:
    If Err.Number <> 0 Then
         MuestraError Err.Number, Err.Description
    End If
End Function

Public Function BorramosTemporal(ByRef db As BaseDatos) As Long
Dim Sql As String

    Sql = "delete from tmpinformes where codusu = " & vSesion.Codigo
    BorramosTemporal = db.Ejecutar(Sql)
    
End Function
