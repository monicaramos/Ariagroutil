Attribute VB_Name = "ModBasico"
Option Explicit

Public Sub AyudaPais(frmBas As frmBasico2, Optional CodActual As String, Optional cWhere As String, Optional Conexion As Byte)

    frmBas.CadenaTots = "S|txtAux(0)|T|Código|1405|;S|txtAux(1)|T|Descripción|4695|;S|txtAux(2)|T|Intracom|900|;"
    frmBas.CadenaConsulta = "SELECT paises.codpais, paises.nompais, if(paises.intracom=0,'No','Si') intracom "
    frmBas.CadenaConsulta = frmBas.CadenaConsulta & " FROM paises "
    frmBas.CadenaConsulta = frmBas.CadenaConsulta & " WHERE (1=1) "
    If cWhere <> "" Then frmBas.CadenaConsulta = frmBas.CadenaConsulta & " and " & cWhere
    frmBas.Tag1 = "Codigo|T|N|||paises|codpais||S|"
    frmBas.Tag2 = "Descripcion|T|S|||paises|nompais|||"
    frmBas.Tag3 = "Intracom.|T|N|||paises|intracom|||"
    
    frmBas.Maxlen1 = 10
    frmBas.Maxlen2 = 40
    frmBas.Maxlen3 = 4
    
    frmBas.pConn = Conexion
    
    frmBas.tabla = "paises"
    frmBas.CampoCP = "codpais"
    frmBas.Caption = "Países"
    frmBas.DeConsulta = True
    frmBas.DatosADevolverBusqueda = "0|1|"
    frmBas.CodigoActual = 0
    If CodActual <> "" Then frmBas.CodigoActual = CodActual
    frmBas.Show vbModal
    
End Sub

