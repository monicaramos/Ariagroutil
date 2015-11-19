Attribute VB_Name = "ModArtic"
' ### [Monica] 29/09/2006
' function que indica si un articulo pertenece a la familia de combustibles

Public Function EsArticuloCombustible(articulo As String) As Boolean
Dim Famia As String
Dim tipoF As String

    EsArticuloCombustible = False
    Famia = ""
    Famia = DevuelveDesdeBD("codfamia", "sartic", "codartic", articulo, "N")
    If Famia = "" Then Exit Function
    tipoF = ""
    tipoF = DevuelveDesdeBD("tipfamia", "sfamia", "codfamia", Famia, "N")
    If tipoF = "" Then Exit Function
    If tipoF = "1" Then EsArticuloCombustible = True

End Function

Public Function EsArticuloDescuento(articulo As String) As Boolean
Dim Famia As String
Dim tipoF As String

    EsArticuloDescuento = False
    Famia = ""
    Famia = DevuelveDesdeBD("codfamia", "sartic", "codartic", articulo, "N")
    If Famia = "" Then Exit Function
    tipoF = ""
    tipoF = DevuelveDesdeBD("tipfamia", "sfamia", "codfamia", Famia, "N")
    If tipoF = "" Then Exit Function
    If tipoF = "2" Then EsArticuloDescuento = True

End Function



Public Function ImpuestoArticulo(articulo As String) As Currency
Dim Sql As String

    ImpuestoArticulo = 0
    Sql = DevuelveDesdeBD("impuesto", "sartic", "codartic", articulo, "N")
    If Sql <> "" Then ImpuestoArticulo = DBLet(CCur(Sql), "N")

End Function

Public Function InsertarFamiliaSiNoExiste(Fam As String) As Boolean
Dim Sql As String

On Error GoTo eInsertarFamiliaSiNoExiste

    InsertarFamiliaSiNoExiste = True
    Sql = ""
    Sql = DevuelveDesdeBD("codfamia", "sfamia", "codfamia", Fam, "N")
    If Sql = "" Then
        Sql = "insert into sfamia (codfamia, nomfamia, tipfamia) values ("
        Sql = Sql & DBSet(Fam, "N") & ",'AUTOMATICA',0)"
        
        conn.Execute Sql
    End If
    
eInsertarFamiliaSiNoExiste:
    If Err.Number <> 0 Then
        InsertarFamiliaSiNoExiste = False
    End If
End Function

Public Function NombreFichero(path As String) As String
Dim cad As String
Dim cad1 As String
Dim b As Boolean
Dim longitud As Integer
Dim i As Integer
Dim J As Integer

    cad = path
    i = 1
    J = Len(cad)
    b = False
    While Not b
        If InStr(cad, "\") = 0 Then
            b = True
        Else
            cad = Mid(cad, InStr(cad, "\") + 1, J)
            J = Len(cad)
        End If
    Wend
    NombreFichero = cad
End Function
