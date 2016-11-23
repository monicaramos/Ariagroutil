VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CParamAplic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"

Option Explicit

'Variables locales que contienen valores de propiedad
'Variables asociadas a cada campo de la tabla de parametros SPARAM

'Datos de la Contabilidad
Private mvarSerConta As String 'Servidor BD de Contabilidad
Private mvarUsuConta As String 'Usuario BD de Contabilidad
Private mvarPasConta As String 'Password BD de Contabilidad
Private mvarNumConta As Byte 'Numero BD de Contabilidad
Private mvarCtaGasto As String 'Cuenta Gastos
Private mvarCtaReten As String 'Cuenta Retencion
Private mvarPorcinte As String
Private mvarPorcrete As String
Private mvarConceDebe As String 'Concepto al debe
Private mvarConceHaber As String 'Concepto al haber
Private mvarNumDiario As String 'Numero de Diario
Private mvarSeguros As String 'hay seguros
Private mvarAvnics As String 'hay avnics
Private mvarTelefonia As String 'hay telefonia
Private mvarGasolinera As String 'hay gasolinera
Private mvarFactSocios As String 'hay facturas socios
Private mvarCoarval As String 'hay facturas coarval

' para seguros
Private mvarSerContaSeg As String 'Servidor BD de Contabilidad
Private mvarUsuContaSeg As String 'Usuario BD de Contabilidad
Private mvarPasContaSeg As String 'Password BD de Contabilidad
Private mvarNumContaSeg As Byte 'Numero BD de Contabilidad
Private mvarCtaBancoSeg As String 'Cuenta contable de banco
Private mvarConceDebeSeg As String 'Concepto al debe
Private mvarConceHaberSeg As String 'Concepto al haber
Private mvarNumDiarioSeg As String 'Numero de Diario
Private mvarRaizCtaSocSeg As String 'raiz cta contable socio

' para telefonia
Private mvarSerContaTel As String 'Servidor BD de Contabilidad
Private mvarUsuContaTel As String 'Usuario BD de Contabilidad
Private mvarPasContaTel As String 'Password BD de Contabilidad
Private mvarNumContaTel As Byte 'Numero BD de Contabilidad
Private mvarConceDebeTel As String 'Concepto al debe
Private mvarConceHaberTel As String 'Concepto al haber
Private mvarNumDiarioTel As String 'Numero de Diario
Private mvarRaizCtaSocTel As String 'raiz cta contable socio
Private mvarCtaBancoTel As String 'Cuenta contable de banco
Private mvarCtaVentaTel As String 'Cuenta contable de venta
Private mvarNumSerieTel As String 'Numero de Serie de Telefonia
Private mvarTipoFicheroTel As Byte 'Tipo de fichero de telefonia

' para factura varias
Private mvarFacturasVarias As Byte 'indicador de si hay facturas varias
Private mvarSerContaFac As String 'Servidor BD de Contabilidad
Private mvarUsuContaFac As String 'Usuario BD de Contabilidad
Private mvarPasContaFac As String 'Password BD de Contabilidad

' para gasolinera
Private mvarSerContaGas As String 'Servidor BD de Contabilidad
Private mvarUsuContaGas As String 'Usuario BD de Contabilidad
Private mvarPasContaGas As String 'Password BD de Contabilidad
Private mvarNumContaGas As Byte 'Numero BD de Contabilidad
Private mvarConceDebeGas As String 'Concepto al debe
Private mvarConceHaberGas As String 'Concepto al haber
Private mvarNumDiarioGas As String 'Numero de Diario
Private mvarRaizCtaSocGas As String 'raiz cta contable socio
Private mvarCtaVentaGas As String 'Cuenta contable de venta
Private mvarCtaContraGas As String 'Cuenta contable de contrapartida
Private mvarNumSerieGas As String 'Numero de Serie de Gasolinera
Private mvarNumSerieGasB As String 'Numero de Serie de Gasoleo B
Private mvarIncreFacGas As String 'Incremento de numero de factura de gaslinera
Private mvarCodIvaGas As String 'Codigo de iva de gasolinera

' para facturas socios
Private mvarSerContaFacSoc As String 'Servidor BD de Contabilidad
Private mvarUsuContaFacSoc As String 'Usuario BD de Contabilidad
Private mvarPasContaFacSoc As String 'Password BD de Contabilidad
Private mvarNumContaFacSoc As Byte 'Numero BD de Contabilidad
Private mvarCtaRetenFacSoc As String 'Cuenta contable de retención
Private mvarRaizCtaFacSoc As String 'raiz cta contable socio
Private mvarPorcReteFacSoc As String ' porcentaje de retencion

Private mvarWebSoporte As String  'Web Soporte
Private mvarDireMail As String 'direccion mail
Private mvarSmtphost As String 'Host
Private mvarSmtpUser As String 'User
Private mvarSmtpPass As String 'Password

Private mvarEnvioDesdeOutlook As Boolean  'Si envia por outlook
Private mvarExeEnvioMail As String


' para facturas coarval
Private mvarSerContaCV As String 'Servidor BD de Contabilidad
Private mvarUsuContaCV As String 'Usuario BD de Contabilidad
Private mvarPasContaCV As String 'Password BD de Contabilidad
Private mvarNumContaCV As Byte 'Numero BD de Contabilidad
Private mvarRaizCtaSocCV As String 'raiz cta contable socio
Private mvarCtaBancoCV As String 'Cuenta contable de banco
Private mvarCtaVentaCV As String 'Cuenta contable de venta
Private mvarLetraSerCV As String 'Numero de Serie de tickets de venta coarval
Private mvarLetraSerFCV As String 'Numero de Serie de facturas coarval
Private mvarLetraSerFinCV As String 'Numero de Serie de facturas internas coarval

Private mvarCtaVentaTickCV As String 'Cuenta contable de venta tickets
Private mvarCtaClienTickCV As String 'Cuenta contable de cliente venta tickets
Private mvarCtaVentaFacCV As String 'Cuenta contable de venta facturas
Private mvarCtaVentaFacInCV As String 'Cuenta contable de venta facturas Interna

Private mvarCodforpaBanCV As String 'forma de pago banco
Private mvarCodforpaConCV As String 'forma de pago contado

Private mvarNumContaCVV As Byte 'Numero BD de Contabilidad
Private mvarCtaBancoCVV As String 'Cuenta contable de banco coarval varias
Private mvarLetraSerCVV As String 'Numero de Serie de tickets de venta coarval varias
Private mvarCodforpaBanCVV As String 'forma de pago banco varias
Private mvarCodforpaConCVV As String 'forma de pago contado varias


Private SQL As String  'Para la sentencia
Private Mens As String    'Para los mensajes


Private mvarHayAridoc As Byte  'Hay Aridoc
Private mvarCarpetaFac As Integer 'Carpeta Facturas
Private mvarExtension As Byte  'Extension del documento
Private mvarC1Factura As Byte  'Campo 1 Factura
Private mvarC2Factura As Byte  'Campo 2 Factura
Private mvarC3Factura As Byte  'Campo 3 Factura
Private mvarC4Factura As Byte  'Campo 4 Factura

Private mvarContabilidadNueva As Boolean




'------------------------------------------------
'Propiedades del modulo CParamAplic
'------------------------------------------------

'servidor conta
Public Property Let ServidorConta(ByVal vData As String)
     mvarSerConta = vData
End Property

Public Property Get ServidorConta() As String
     ServidorConta = mvarSerConta
End Property

Public Property Let UsuarioConta(ByVal vData As String)
     mvarUsuConta = vData
End Property

Public Property Get UsuarioConta() As String
     UsuarioConta = mvarUsuConta
End Property

Public Property Let PasswordConta(ByVal vData As String)
     mvarPasConta = vData
End Property

Public Property Get PasswordConta() As String
     PasswordConta = mvarPasConta
End Property

Public Property Let NumeroConta(ByVal vData As Byte)
     mvarNumConta = vData
End Property

Public Property Get NumeroConta() As Byte
     NumeroConta = mvarNumConta
End Property

Public Property Let CtaGasto(ByVal vData As String)
     mvarCtaGasto = vData
End Property

Public Property Get CtaGasto() As String
     CtaGasto = mvarCtaGasto
End Property

Public Property Let CtaReten(ByVal vData As String)
     mvarCtaReten = vData
End Property

Public Property Get CtaReten() As String
     CtaReten = mvarCtaReten
End Property

Public Property Let Porcinte(ByVal vData As String)
     mvarPorcinte = vData
End Property

Public Property Get Porcinte() As String
     Porcinte = mvarPorcinte
End Property

Public Property Let Porcrete(ByVal vData As String)
     mvarPorcrete = vData
End Property

Public Property Get Porcrete() As String
     Porcrete = mvarPorcrete
End Property

Public Property Let ConceDebe(ByVal vData As Long)
     mvarConceDebe = vData
End Property

Public Property Get ConceDebe() As Long
     ConceDebe = mvarConceDebe
End Property

Public Property Let ConceHaber(ByVal vData As Long)
     mvarConceHaber = vData
End Property

Public Property Get ConceHaber() As Long
     ConceHaber = mvarConceHaber
End Property

Public Property Let NumDiario(ByVal vData As Long)
     mvarNumDiario = vData
End Property

Public Property Get NumDiario() As Long
     NumDiario = mvarNumDiario
End Property


Public Property Let WebSoporte(ByVal vData As String)
     mvarWebSoporte = vData
End Property

Public Property Get WebSoporte() As String
     WebSoporte = mvarWebSoporte
End Property

Public Property Let DireMail(ByVal vData As String)
     mvarDireMail = vData
End Property

Public Property Get DireMail() As String
     DireMail = mvarDireMail
End Property

Public Property Let Smtphost(ByVal vData As String)
     mvarSmtphost = vData
End Property

Public Property Get Smtphost() As String
     Smtphost = mvarSmtphost
End Property

Public Property Let SmtpUser(ByVal vData As String)
     mvarSmtpUser = vData
End Property

Public Property Get SmtpUser() As String
     SmtpUser = mvarSmtpUser
End Property

Public Property Let Smtppass(ByVal vData As String)
     mvarSmtpPass = vData
End Property

Public Property Get Smtppass() As String
     Smtppass = mvarSmtpPass
End Property

Public Property Let Seguros(ByVal vData As Byte)
     mvarSeguros = vData
End Property

Public Property Get Seguros() As Byte
     Seguros = mvarSeguros
End Property

Public Property Let Avnics(ByVal vData As Byte)
     mvarAvnics = vData
End Property

Public Property Get Avnics() As Byte
     Avnics = mvarAvnics
End Property


Public Property Let Telefonia(ByVal vData As Byte)
     mvarTelefonia = vData
End Property

Public Property Get Telefonia() As Byte
     Telefonia = mvarTelefonia
End Property

Public Property Let FacturasVarias(ByVal vData As Byte)
     mvarFacturasVarias = vData
End Property

Public Property Get FacturasVarias() As Byte
     FacturasVarias = mvarFacturasVarias
End Property

Public Property Let Gasolinera(ByVal vData As Byte)
     mvarGasolinera = vData
End Property

Public Property Get Gasolinera() As Byte
     Gasolinera = mvarGasolinera
End Property

Public Property Let FactSocios(ByVal vData As Byte)
     mvarFactSocios = vData
End Property

Public Property Get FactSocios() As Byte
     FactSocios = mvarFactSocios
End Property


Public Property Let Coarval(ByVal vData As Byte)
     mvarCoarval = vData
End Property

Public Property Get Coarval() As Byte
     Coarval = mvarCoarval
End Property


'servidor conta
Public Property Let ServidorContaSeg(ByVal vData As String)
     mvarSerContaSeg = vData
End Property

Public Property Get ServidorContaSeg() As String
     ServidorContaSeg = mvarSerContaSeg
End Property

Public Property Let UsuarioContaSeg(ByVal vData As String)
     mvarUsuContaSeg = vData
End Property

Public Property Get UsuarioContaSeg() As String
     UsuarioContaSeg = mvarUsuContaSeg
End Property

Public Property Let PasswordContaSeg(ByVal vData As String)
     mvarPasContaSeg = vData
End Property

Public Property Get PasswordContaSeg() As String
     PasswordContaSeg = mvarPasContaSeg
End Property

Public Property Let NumeroContaSeg(ByVal vData As Byte)
     mvarNumContaSeg = vData
End Property

Public Property Get NumeroContaSeg() As Byte
     NumeroContaSeg = mvarNumContaSeg
End Property

Public Property Let CtaBancoSeg(ByVal vData As String)
     mvarCtaBancoSeg = vData
End Property

Public Property Get CtaBancoSeg() As String
     CtaBancoSeg = mvarCtaBancoSeg
End Property

Public Property Let ConceDebeSeg(ByVal vData As Long)
     mvarConceDebeSeg = vData
End Property

Public Property Get ConceDebeSeg() As Long
     ConceDebeSeg = mvarConceDebeSeg
End Property

Public Property Let ConceHaberSeg(ByVal vData As Long)
     mvarConceHaberSeg = vData
End Property

Public Property Get ConceHaberSeg() As Long
     ConceHaberSeg = mvarConceHaberSeg
End Property

Public Property Let NumDiarioSeg(ByVal vData As Long)
     mvarNumDiarioSeg = vData
End Property

Public Property Get NumDiarioSeg() As Long
     NumDiarioSeg = mvarNumDiarioSeg
End Property

Public Property Let RaizCtaSocSeg(ByVal vData As String)
     mvarRaizCtaSocSeg = vData
End Property

Public Property Get RaizCtaSocSeg() As String
     RaizCtaSocSeg = mvarRaizCtaSocSeg
End Property


' telefonia
'servidor conta
Public Property Let ServidorContaTel(ByVal vData As String)
     mvarSerContaTel = vData
End Property

Public Property Get ServidorContaTel() As String
     ServidorContaTel = mvarSerContaTel
End Property

Public Property Let UsuarioContaTel(ByVal vData As String)
     mvarUsuContaTel = vData
End Property

Public Property Get UsuarioContaTel() As String
     UsuarioContaTel = mvarUsuContaTel
End Property

Public Property Let PasswordContaTel(ByVal vData As String)
     mvarPasContaTel = vData
End Property

Public Property Get PasswordContaTel() As String
     PasswordContaTel = mvarPasContaTel
End Property

Public Property Let NumeroContaTel(ByVal vData As Byte)
     mvarNumContaTel = vData
End Property

Public Property Get NumeroContaTel() As Byte
     NumeroContaTel = mvarNumContaTel
End Property

Public Property Let ConceDebeTel(ByVal vData As Long)
     mvarConceDebeTel = vData
End Property

Public Property Get ConceDebeTel() As Long
     ConceDebeTel = mvarConceDebeTel
End Property

Public Property Let ConceHaberTel(ByVal vData As Long)
     mvarConceHaberTel = vData
End Property

Public Property Get ConceHaberTel() As Long
     ConceHaberTel = mvarConceHaberTel
End Property

Public Property Let NumDiarioTel(ByVal vData As Long)
     mvarNumDiarioTel = vData
End Property

Public Property Get NumDiarioTel() As Long
     NumDiarioTel = mvarNumDiarioTel
End Property

Public Property Let RaizCtaSocTel(ByVal vData As String)
     mvarRaizCtaSocTel = vData
End Property

Public Property Get RaizCtaSocTel() As String
     RaizCtaSocTel = mvarRaizCtaSocTel
End Property

Public Property Let CtaBancoTel(ByVal vData As String)
     mvarCtaBancoTel = vData
End Property

Public Property Get CtaBancoTel() As String
     CtaBancoTel = mvarCtaBancoTel
End Property

Public Property Let CtaVentaTel(ByVal vData As String)
     mvarCtaVentaTel = vData
End Property

Public Property Get CtaVentaTel() As String
     CtaVentaTel = mvarCtaVentaTel
End Property

Public Property Let NumSerieTel(ByVal vData As String)
     mvarNumSerieTel = vData
End Property

Public Property Get NumSerieTel() As String
     NumSerieTel = mvarNumSerieTel
End Property

Public Property Let TipoFicheroTel(ByVal vData As Byte)
     mvarTipoFicheroTel = vData
End Property

Public Property Get TipoFicheroTel() As Byte
     TipoFicheroTel = mvarTipoFicheroTel
End Property




'facturas varias
Public Property Let ServidorContaFac(ByVal vData As String)
     mvarSerContaFac = vData
End Property

Public Property Get ServidorContaFac() As String
     ServidorContaFac = mvarSerContaFac
End Property

Public Property Let UsuarioContaFac(ByVal vData As String)
     mvarUsuContaFac = vData
End Property

Public Property Get UsuarioContaFac() As String
     UsuarioContaFac = mvarUsuContaFac
End Property

Public Property Let PasswordContaFac(ByVal vData As String)
     mvarPasContaFac = vData
End Property

Public Property Get PasswordContaFac() As String
     PasswordContaFac = mvarPasContaFac
End Property

' gasolinera
Public Property Let ServidorContaGas(ByVal vData As String)
     mvarSerContaGas = vData
End Property

Public Property Get ServidorContaGas() As String
     ServidorContaGas = mvarSerContaGas
End Property

Public Property Let UsuarioContaGas(ByVal vData As String)
     mvarUsuContaGas = vData
End Property

Public Property Get UsuarioContaGas() As String
     UsuarioContaGas = mvarUsuContaGas
End Property

Public Property Let PasswordContaGas(ByVal vData As String)
     mvarPasContaGas = vData
End Property

Public Property Get PasswordContaGas() As String
     PasswordContaGas = mvarPasContaGas
End Property

Public Property Let NumeroContaGas(ByVal vData As Byte)
     mvarNumContaGas = vData
End Property

Public Property Get NumeroContaGas() As Byte
     NumeroContaGas = mvarNumContaGas
End Property

Public Property Let ConceDebeGas(ByVal vData As Long)
     mvarConceDebeGas = vData
End Property

Public Property Get ConceDebeGas() As Long
     ConceDebeGas = mvarConceDebeGas
End Property

Public Property Let ConceHaberGas(ByVal vData As Long)
     mvarConceHaberGas = vData
End Property

Public Property Get ConceHaberGas() As Long
     ConceHaberGas = mvarConceHaberGas
End Property

Public Property Let NumDiarioGas(ByVal vData As Long)
     mvarNumDiarioGas = vData
End Property

Public Property Get NumDiarioGas() As Long
     NumDiarioGas = mvarNumDiarioGas
End Property

Public Property Let RaizCtaSocGas(ByVal vData As String)
     mvarRaizCtaSocGas = vData
End Property

Public Property Get RaizCtaSocGas() As String
     RaizCtaSocGas = mvarRaizCtaSocGas
End Property

Public Property Let CtaVentasGas(ByVal vData As String)
     mvarCtaVentaGas = vData
End Property

Public Property Get CtaVentasGas() As String
     CtaVentasGas = mvarCtaVentaGas
End Property

Public Property Let CtaContraGas(ByVal vData As String)
     mvarCtaContraGas = vData
End Property

Public Property Get CtaContraGas() As String
     CtaContraGas = mvarCtaContraGas
End Property

Public Property Let NumSerieGas(ByVal vData As String)
     mvarNumSerieGas = vData
End Property

Public Property Get NumSerieGas() As String
     NumSerieGas = mvarNumSerieGas
End Property


'[Monica]04/07/2013: introducimos la letra de serie de gasoleo B
Public Property Let NumSerieGasB(ByVal vData As String)
     mvarNumSerieGasB = vData
End Property

Public Property Get NumSerieGasB() As String
     NumSerieGasB = mvarNumSerieGasB
End Property




Public Property Let IncreFactGas(ByVal vData As Long)
     mvarIncreFacGas = vData
End Property

Public Property Get IncreFactGas() As Long
     IncreFactGas = mvarIncreFacGas
End Property

Public Property Let CodIvaGas(ByVal vData As Long)
     mvarCodIvaGas = vData
End Property

Public Property Get CodIvaGas() As Long
     CodIvaGas = mvarCodIvaGas
End Property



' facturas de socio
Public Property Let ServidorContaFacSoc(ByVal vData As String)
     mvarSerContaFacSoc = vData
End Property

Public Property Get ServidorContaFacSoc() As String
     ServidorContaFacSoc = mvarSerContaFacSoc
End Property

Public Property Let UsuarioContaFacSoc(ByVal vData As String)
     mvarUsuContaFacSoc = vData
End Property

Public Property Get UsuarioContaFacSoc() As String
     UsuarioContaFacSoc = mvarUsuContaFacSoc
End Property

Public Property Let PasswordContaFacSoc(ByVal vData As String)
     mvarPasContaFacSoc = vData
End Property

Public Property Get PasswordContaFacSoc() As String
     PasswordContaFacSoc = mvarPasContaFacSoc
End Property

Public Property Let NumeroContaFacSoc(ByVal vData As Byte)
     mvarNumContaFacSoc = vData
End Property

Public Property Get NumeroContaFacSoc() As Byte
     NumeroContaFacSoc = mvarNumContaFacSoc
End Property

Public Property Let RaizCtaFacSoc(ByVal vData As String)
     mvarRaizCtaFacSoc = vData
End Property

Public Property Get RaizCtaFacSoc() As String
     RaizCtaFacSoc = mvarRaizCtaFacSoc
End Property

Public Property Let PorcreteFacSoc(ByVal vData As String)
     mvarPorcReteFacSoc = vData
End Property

Public Property Get PorcreteFacSoc() As String
     PorcreteFacSoc = mvarPorcReteFacSoc
End Property

Public Property Let CtaRetenFacSoc(ByVal vData As String)
     mvarCtaRetenFacSoc = vData
End Property

Public Property Get CtaRetenFacSoc() As String
     CtaRetenFacSoc = mvarCtaRetenFacSoc
End Property

Public Property Get EnvioDesdeOutlook() As Boolean
     EnvioDesdeOutlook = mvarEnvioDesdeOutlook
End Property

Public Property Let EnvioDesdeOutlook(ByVal vData As Boolean)
     mvarEnvioDesdeOutlook = vData
End Property

Public Property Get ExeEnvioMail() As String
     ExeEnvioMail = mvarExeEnvioMail
End Property

Public Property Let ExeEnvioMail(ByVal vData As String)
     mvarExeEnvioMail = vData
End Property


' facturas coarval
'servidor conta
Public Property Let ServidorContaCV(ByVal vData As String)
     mvarSerContaCV = vData
End Property

Public Property Get ServidorContaCV() As String
     ServidorContaCV = mvarSerContaCV
End Property

Public Property Let UsuarioContaCV(ByVal vData As String)
     mvarUsuContaCV = vData
End Property

Public Property Get UsuarioContaCV() As String
     UsuarioContaCV = mvarUsuContaCV
End Property

Public Property Let PasswordContaCV(ByVal vData As String)
     mvarPasContaCV = vData
End Property

Public Property Get PasswordContaCV() As String
     PasswordContaCV = mvarPasContaCV
End Property

Public Property Let NumeroContaCV(ByVal vData As Byte)
     mvarNumContaCV = vData
End Property

Public Property Get NumeroContaCV() As Byte
     NumeroContaCV = mvarNumContaCV
End Property

Public Property Let RaizCtaSocCV(ByVal vData As String)
     mvarRaizCtaSocCV = vData
End Property

Public Property Get RaizCtaSocCV() As String
     RaizCtaSocCV = mvarRaizCtaSocCV
End Property

Public Property Let CtaBancoCV(ByVal vData As String)
     mvarCtaBancoCV = vData
End Property

Public Property Get CtaBancoCV() As String
     CtaBancoCV = mvarCtaBancoCV
End Property

Public Property Let CtaVentaCV(ByVal vData As String)
     mvarCtaVentaCV = vData
End Property

Public Property Get CtaVentaCV() As String
     CtaVentaCV = mvarCtaVentaCV
End Property

Public Property Let LetraSerCV(ByVal vData As String)
     mvarLetraSerCV = vData
End Property

Public Property Get LetraSerCV() As String
     LetraSerCV = mvarLetraSerCV
End Property


Public Property Let LetraSerFCV(ByVal vData As String)
     mvarLetraSerFCV = vData
End Property

Public Property Get LetraSerFCV() As String
     LetraSerFCV = mvarLetraSerFCV
End Property

Public Property Let LetraSerFinCV(ByVal vData As String)
     mvarLetraSerFinCV = vData
End Property

Public Property Get LetraSerFinCV() As String
     LetraSerFinCV = mvarLetraSerFinCV
End Property


Public Property Let CtaVentaTickCV(ByVal vData As String)
     mvarCtaVentaTickCV = vData
End Property

Public Property Get CtaVentaTickCV() As String
     CtaVentaTickCV = mvarCtaVentaTickCV
End Property


Public Property Let CtaClienTickCV(ByVal vData As String)
     mvarCtaClienTickCV = vData
End Property

Public Property Get CtaClienTickCV() As String
     CtaClienTickCV = mvarCtaClienTickCV
End Property


Public Property Let CtaVentaFacCV(ByVal vData As String)
     mvarCtaVentaFacCV = vData
End Property

Public Property Get CtaVentaFacCV() As String
     CtaVentaFacCV = mvarCtaVentaFacCV
End Property

Public Property Let CtaVentaFacInCV(ByVal vData As String)
     mvarCtaVentaFacInCV = vData
End Property

Public Property Get CtaVentaFacInCV() As String
     CtaVentaFacInCV = mvarCtaVentaFacInCV
End Property

Public Property Let CodforpaBanCV(ByVal vData As String)
     mvarCodforpaBanCV = vData
End Property

Public Property Get CodforpaBanCV() As String
     CodforpaBanCV = mvarCodforpaBanCV
End Property

Public Property Let CodforpaConCV(ByVal vData As String)
     mvarCodforpaConCV = vData
End Property

Public Property Get CodforpaConCV() As String
     CodforpaConCV = mvarCodforpaConCV
End Property


Public Property Let NumeroContaCVV(ByVal vData As Byte)
     mvarNumContaCVV = vData
End Property

Public Property Get NumeroContaCVV() As Byte
     NumeroContaCVV = mvarNumContaCVV
End Property

Public Property Let CtaBancoCVV(ByVal vData As String)
     mvarCtaBancoCVV = vData
End Property

Public Property Get CtaBancoCVV() As String
     CtaBancoCVV = mvarCtaBancoCVV
End Property

Public Property Let LetraSerCVV(ByVal vData As String)
     mvarLetraSerCVV = vData
End Property

Public Property Get LetraSerCVV() As String
     LetraSerCVV = mvarLetraSerCVV
End Property

Public Property Let CodforpaBanCVV(ByVal vData As String)
     mvarCodforpaBanCVV = vData
End Property

Public Property Get CodforpaBanCVV() As String
     CodforpaBanCVV = mvarCodforpaBanCVV
End Property

Public Property Let CodforpaConCVV(ByVal vData As String)
     mvarCodforpaConCVV = vData
End Property

Public Property Get CodforpaConCVV() As String
     CodforpaConCVV = mvarCodforpaConCVV
End Property

'[Monica]12/11/2013: Aridoc
Public Property Let HayAridoc(ByVal vData As Byte)
     mvarHayAridoc = vData
End Property

Public Property Get HayAridoc() As Byte
     HayAridoc = mvarHayAridoc
End Property

Public Property Let CarpetaFac(ByVal vData As Integer)
     mvarCarpetaFac = vData
End Property

Public Property Get CarpetaFac() As Integer
     CarpetaFac = mvarCarpetaFac
End Property

Public Property Let Extension(ByVal vData As Byte)
     mvarExtension = vData
End Property

Public Property Get Extension() As Byte
     Extension = mvarExtension
End Property

Public Property Let C1Factura(ByVal vData As Byte)
     mvarC1Factura = vData
End Property

Public Property Get C1Factura() As Byte
     C1Factura = mvarC1Factura
End Property

Public Property Let C2Factura(ByVal vData As Byte)
     mvarC2Factura = vData
End Property

Public Property Get C2Factura() As Byte
     C2Factura = mvarC2Factura
End Property

Public Property Let C3Factura(ByVal vData As Byte)
     mvarC3Factura = vData
End Property

Public Property Get C3Factura() As Byte
     C3Factura = mvarC3Factura
End Property

Public Property Let C4Factura(ByVal vData As Byte)
     mvarC4Factura = vData
End Property

Public Property Get C4Factura() As Byte
     C4Factura = mvarC4Factura
End Property

'[Monica]23/11/2016: si hay contabilidad nueva
Public Property Get ContabilidadNueva() As Boolean
     ContabilidadNueva = mvarContabilidadNueva
End Property



'******
'------------------------------------------------
'FUNCIONES del modulo Cparametros
'------------------------------------------------

'****   LEER  ***
Public Function leer() As Byte
Dim Rs As ADODB.Recordset
On Error GoTo Err3
    
    SQL = "SELECT * FROM sparam "
    Set Rs = New ADODB.Recordset
    Rs.Open SQL, conn, adOpenForwardOnly, adLockOptimistic

    If Rs.EOF Then
        leer = 1
    Else
        'Datos Contabilidad avnics
        mvarSerConta = DBLet(Rs!serConta, "T") 'servidor conta
        mvarUsuConta = DBLet(Rs!UsuConta, "T") 'usuario conta
        mvarPasConta = DBLet(Rs!pasconta, "T") 'password
        mvarNumConta = DBLet(Rs!NumConta, "N") 'numero conta
        mvarCtaGasto = DBLet(Rs!CtaGasto, "T")
        mvarCtaReten = DBLet(Rs!CtaReten, "T")
        mvarPorcinte = DBLet(Rs!Porcinte, "N")
        mvarPorcrete = DBLet(Rs!Porcrete, "N")
        mvarConceDebe = DBLet(Rs!ConceDebe, "N") 'concepto al debe
        mvarConceHaber = DBLet(Rs!ConceHaber, "N") 'concepto al haber
        mvarNumDiario = DBLet(Rs!NumDiario, "N") 'numero de diario
        
        mvarWebSoporte = DBLet(Rs!WebSoporte, "T") 'web soporte
        mvarDireMail = DBLet(Rs!DireMail, "T") 'mail
        mvarSmtphost = DBLet(Rs!Smtphost, "T") 'host
        mvarSmtpUser = DBLet(Rs!SmtpUser, "T") 'user
        mvarSmtpPass = DBLet(Rs!Smtppass, "T") 'pass
        
        mvarSeguros = DBLet(Rs!Seguros, "N") 'hay seguros agrarios
        mvarAvnics = DBLet(Rs!Avnics, "N") 'hay gestion de avnics
        mvarTelefonia = DBLet(Rs!Telefonia, "N") 'hay telefonia
        mvarFacturasVarias = DBLet(Rs!FacturasVarias, "N") 'hay facturasvarias
        mvarGasolinera = DBLet(Rs!Gasolinera, "N") 'hay gasolinera
        mvarFactSocios = DBLet(Rs!FactSocios, "N") ' hay facturas socios
        mvarCoarval = DBLet(Rs!Coarval, "N") ' hay facturas coarval
        
        'datos contabilidad para seguros
        mvarSerContaSeg = DBLet(Rs!serContaseg, "T") 'servidor conta
        mvarUsuContaSeg = DBLet(Rs!UsuContaseg, "T") 'usuario conta
        mvarPasContaSeg = DBLet(Rs!pascontaseg, "T") 'password
        mvarNumContaSeg = DBLet(Rs!NumContaseg, "N") 'numero conta
        mvarConceDebeSeg = DBLet(Rs!ConceDebeSeg, "N") 'concepto al debe
        mvarConceHaberSeg = DBLet(Rs!ConceHaberSeg, "N") 'concepto al haber
        mvarNumDiarioSeg = DBLet(Rs!NumDiarioSeg, "N") 'numero de diario
        mvarCtaBancoSeg = DBLet(Rs!CtaBancoSeg, "T") 'cta contable banco seguros
        mvarRaizCtaSocSeg = DBLet(Rs!RaizCtaSocSeg, "T") 'raiz cta contable socio seguros
        
        'datos contabilidad para telefonia
        mvarSerContaTel = DBLet(Rs!serContatel, "T") 'servidor conta
        mvarUsuContaTel = DBLet(Rs!UsuContatel, "T") 'usuario conta
        mvarPasContaTel = DBLet(Rs!pascontatel, "T") 'password
        mvarNumContaTel = DBLet(Rs!NumContatel, "N") 'numero conta
        mvarConceDebeTel = DBLet(Rs!ConceDebeTel, "N") 'concepto al debe
        mvarConceHaberTel = DBLet(Rs!ConceHaberTel, "N") 'concepto al haber
        mvarNumDiarioTel = DBLet(Rs!NumDiarioTel, "N") 'numero de diario
        mvarCtaBancoTel = DBLet(Rs!CtaBancoTel, "T") 'cta contable banco telefonia
        mvarRaizCtaSocTel = DBLet(Rs!RaizCtaSocTel, "T") 'raiz cta contable socio telefonia
        mvarCtaVentaTel = DBLet(Rs!CtaVentaTel, "T") 'cta contable venta telefonia
        mvarNumSerieTel = DBLet(Rs!NumSerieTel, "T") 'Numero de Serie de telefonia
        mvarTipoFicheroTel = DBLet(Rs!TipoFicheroTel, "N") 'Tipo de Fichero de Telefonia
        
        
        'datos contabilidad para facturas varias
        mvarSerContaFac = DBLet(Rs!serContafac, "T") 'servidor conta
        mvarUsuContaFac = DBLet(Rs!UsuContafac, "T") 'usuario conta
        mvarPasContaFac = DBLet(Rs!pascontafac, "T") 'password
        
        'datos contabilidad para gasolinera
        mvarSerContaGas = DBLet(Rs!serContagas, "T") 'servidor conta
        mvarUsuContaGas = DBLet(Rs!UsuContagas, "T") 'usuario conta
        mvarPasContaGas = DBLet(Rs!pascontagas, "T") 'password
        mvarNumContaGas = DBLet(Rs!NumcontaGas, "N") 'numero conta
        mvarConceDebeGas = DBLet(Rs!ConceDebeGas, "N") 'concepto al debe
        mvarConceHaberGas = DBLet(Rs!ConceHaberGas, "N") 'concepto al haber
        mvarNumDiarioGas = DBLet(Rs!NumDiariGas, "N") 'numero de diario
        mvarCtaVentaGas = DBLet(Rs!CtaVentasGas, "T") 'cta contable ventas gasolinera
        mvarRaizCtaSocGas = DBLet(Rs!RaizCtaSocGas, "T") 'raiz cta contable socio gasolinera
        mvarCtaContraGas = DBLet(Rs!CtaContraGas, "T") 'cta contable contrapartida gasolinera
        mvarNumSerieGas = DBLet(Rs!LetraSerGas, "T") 'Numero de Serie de gasolinera
        mvarNumSerieGasB = DBLet(Rs!LetraSerGasB, "T") 'Numero de Serie de gasoleo B
        mvarIncreFacGas = DBLet(Rs!increfacgas, "N") 'Incremento de factura de gasolinera
        mvarCodIvaGas = DBLet(Rs!CodIvaGas, "N") 'codigo de iva de gasolinera
        
        'datos contabilidad para facturas socios
        mvarSerContaFacSoc = DBLet(Rs!serContafacsoc, "T") 'servidor conta
        mvarUsuContaFacSoc = DBLet(Rs!UsuContafacsoc, "T") 'usuario conta
        mvarPasContaFacSoc = DBLet(Rs!pascontafacsoc, "T") 'password
        mvarNumContaFacSoc = DBLet(Rs!NumContafacsoc, "N") 'numero conta
        mvarCtaRetenFacSoc = DBLet(Rs!CtaRetenFacSoc, "T") 'cta contable retencion factura socios
        mvarRaizCtaFacSoc = DBLet(Rs!RaizCtaFacSoc, "T") 'raiz cta contable socio facturas socios
        mvarPorcReteFacSoc = DBLet(Rs!PorcretenFacSoc, "N") 'porcentaje de retencion de la factura
        
        'datos contabilidad para facturas coarval
        mvarSerContaCV = DBLet(Rs!serContacv, "T") 'servidor conta
        mvarUsuContaCV = DBLet(Rs!UsuContacv, "T") 'usuario conta
        mvarPasContaCV = DBLet(Rs!pascontacv, "T") 'password
        mvarNumContaCV = DBLet(Rs!NumContacv, "N") 'numero conta
        mvarCtaBancoCV = DBLet(Rs!CtaBancoCV, "T") 'cta contable banco coarval
        mvarRaizCtaSocCV = DBLet(Rs!RaizCtaSocCV, "T") 'raiz cta contable socio coarval
        mvarCtaVentaCV = DBLet(Rs!CtaVentaCV, "T") 'cta contable venta coarval
        mvarLetraSerCV = DBLet(Rs!LetraSerCV, "T") 'Numero de Serie tickets de coarval
        mvarLetraSerFCV = DBLet(Rs!LetraSerFCV, "T") 'Numero de Serie facturas venta de coarval
        mvarLetraSerFinCV = DBLet(Rs!LetraSerFinCV, "T") 'Numero de Serie facturas venta internas de coarval
        mvarCtaVentaTickCV = DBLet(Rs!CtaVentaTickCV, "T") 'cta contable venta coarval ticket
        mvarCtaClienTickCV = DBLet(Rs!CtaClienTickCV, "T") 'cta contable cliente venta coarval ticket
        mvarCtaVentaFacCV = DBLet(Rs!CtaVentaFacCV, "T") 'cta contable venta coarval factura
        mvarCtaVentaFacInCV = DBLet(Rs!CtaVentaFacInCV, "T") 'cta contable venta coarval factura interna
        mvarCodforpaBanCV = DBLet(Rs!CodforpaBanCV, "T") 'forma de pago de banco
        mvarCodforpaConCV = DBLet(Rs!CodforpaConCV, "T") 'forma de pago de contado
        mvarNumContaCVV = DBLet(Rs!NumContacvv, "N") 'numero conta
        mvarCtaBancoCVV = DBLet(Rs!CtaBancoCVV, "T") 'cta contable banco coarval varias
        mvarLetraSerCVV = DBLet(Rs!LetraSerCVV, "T") 'Letra de Serie tickets de coarval
        mvarCodforpaBanCVV = DBLet(Rs!CodforpaBanCVV, "T") 'forma de pago de banco
        mvarCodforpaConCVV = DBLet(Rs!CodforpaConCVV, "T") 'forma de pago de contado
        
        Me.EnvioDesdeOutlook = DBLet(Rs!EnvioDesdeOutlook, "N") = 1
        mvarExeEnvioMail = DBLet(Rs!arigesmail, "T")
       
        'aridoc
        mvarHayAridoc = DBLet(Rs!HayAridoc, "N") ' hay aridoc
        mvarCarpetaFac = DBLet(Rs!codcarpetafvar, "N") 'carpeta factura
        mvarExtension = DBLet(Rs!codextension, "N") 'extension documento
        mvarC1Factura = DBLet(Rs!c1fvararidoc, "N") 'campo 1 factura
        mvarC2Factura = DBLet(Rs!c2fvararidoc, "N") 'campo 2 factura
        mvarC3Factura = DBLet(Rs!c3fvararidoc, "N") 'campo 3 factura
        mvarC4Factura = DBLet(Rs!c4fvararidoc, "N") 'campo 4 factura
       
        '[Monica]23/11/2016: nueva contabilidad
        mvarContabilidadNueva = (DBLet(Rs!nuevacontabilidad, "N") = 1)
       
       
        leer = 0
    End If

    Rs.Close
    Set Rs = Nothing
    Exit Function

Err3:
    Mens = "Se ha producido un error." & vbCrLf
    Mens = Mens & "Número: " & Err.Number & vbCrLf
    Mens = Mens & "Descripción: " & Err.Description
    MsgBox Mens, vbExclamation
    Rs.Close
    Set Rs = Nothing
    leer = 1
End Function


Public Function Modificar(Codigo As Byte) As Byte
'Modifica los parametros de la Aplicación
'Modifica la Tabla: sparam, BD: Arigasol
On Error GoTo EModificar

    SQL = "UPDATE sparam SET "
    SQL = SQL & " serconta= '" & mvarSerConta & "'"
    SQL = SQL & ", usuconta= " & DBSet(mvarUsuConta, "T")
    SQL = SQL & ", pasconta= " & DBSet(mvarPasConta, "T")
    SQL = SQL & ", numconta= " & DBSet(mvarNumConta, "N", "S")
    SQL = SQL & ", ctagasto= " & DBSet(mvarCtaGasto, "T")
    SQL = SQL & ", ctareten = " & DBSet(mvarCtaReten, "T")
    SQL = SQL & ", porcinte= " & DBSet(mvarPorcinte, "N", "S")
    SQL = SQL & ", porcrete= " & DBSet(mvarPorcrete, "N", "S")
    SQL = SQL & ", websoporte= " & DBSet(mvarWebSoporte, "T")
    SQL = SQL & ", diremail= " & DBSet(mvarDireMail, "T")
    SQL = SQL & ", smtphost= " & DBSet(mvarSmtphost, "T")
    SQL = SQL & ", smtpuser= " & DBSet(mvarSmtpUser, "T")
    SQL = SQL & ", smtppass= " & DBSet(mvarSmtpPass, "T")
    SQL = SQL & ", concedebe= " & DBSet(mvarConceDebe, "N")
    SQL = SQL & ", concehaber= " & DBSet(mvarConceHaber, "N")
    SQL = SQL & ", numdiario= " & DBSet(mvarNumDiario, "N")
    SQL = SQL & ", seguros= " & DBSet(mvarSeguros, "N")
    SQL = SQL & ", avnics= " & DBSet(mvarAvnics, "N")
    SQL = SQL & ", telefonia= " & DBSet(mvarTelefonia, "N")
    SQL = SQL & ", facturasvarias= " & DBSet(mvarFacturasVarias, "N")
    SQL = SQL & ", gasolinera= " & DBSet(mvarGasolinera, "N")
    SQL = SQL & ", factsocios= " & DBSet(mvarFactSocios, "N")
    SQL = SQL & ", coarval= " & DBSet(mvarCoarval, "N")
    'seguros
    SQL = SQL & ", sercontaseg= '" & mvarSerContaSeg & "'"
    SQL = SQL & ", usucontaseg= " & DBSet(mvarUsuContaSeg, "T")
    SQL = SQL & ", pascontaseg= " & DBSet(mvarPasContaSeg, "T")
    SQL = SQL & ", numcontaseg= " & DBSet(mvarNumContaSeg, "N", "S")
    SQL = SQL & ", concedebeseg= " & DBSet(mvarConceDebeSeg, "N")
    SQL = SQL & ", concehaberseg= " & DBSet(mvarConceHaberSeg, "N")
    SQL = SQL & ", numdiarioseg= " & DBSet(mvarNumDiarioSeg, "N")
    SQL = SQL & ", ctabancoseg= " & DBSet(mvarCtaBancoSeg, "T")
    SQL = SQL & ", raizctasocseg= " & DBSet(mvarRaizCtaSocSeg, "T")
    
    'telefonia
    SQL = SQL & ", sercontatel= '" & mvarSerContaTel & "'"
    SQL = SQL & ", usucontatel= " & DBSet(mvarUsuContaTel, "T")
    SQL = SQL & ", pascontatel= " & DBSet(mvarPasContaTel, "T")
    SQL = SQL & ", numcontatel= " & DBSet(mvarNumContaTel, "N", "S")
    SQL = SQL & ", concedebetel= " & DBSet(mvarConceDebeTel, "N")
    SQL = SQL & ", concehabertel= " & DBSet(mvarConceHaberTel, "N")
    SQL = SQL & ", numdiariotel= " & DBSet(mvarNumDiarioTel, "N")
    SQL = SQL & ", ctabancotel= " & DBSet(mvarCtaBancoTel, "T")
    SQL = SQL & ", raizctasoctel= " & DBSet(mvarRaizCtaSocTel, "T")
    SQL = SQL & ", ctaventatel= " & DBSet(mvarCtaVentaTel, "T")
    SQL = SQL & ", numserietel= " & DBSet(mvarNumSerieTel, "T")
    SQL = SQL & ", tipoficherotel= " & DBSet(mvarTipoFicheroTel, "T")
    'facturas varias
    SQL = SQL & ", sercontafac= '" & mvarSerContaFac & "'"
    SQL = SQL & ", usucontafac= " & DBSet(mvarUsuContaFac, "T")
    SQL = SQL & ", pascontafac= " & DBSet(mvarPasContaFac, "T")
    'gasolinera
    SQL = SQL & ", sercontagas= '" & mvarSerContaGas & "'"
    SQL = SQL & ", usucontagas= " & DBSet(mvarUsuContaGas, "T")
    SQL = SQL & ", pascontagas= " & DBSet(mvarPasContaGas, "T")
    SQL = SQL & ", numcontagas= " & DBSet(mvarNumContaGas, "N", "S")
    SQL = SQL & ", concedebegas= " & DBSet(mvarConceDebeGas, "N")
    SQL = SQL & ", concehabergas= " & DBSet(mvarConceHaberGas, "N")
    SQL = SQL & ", numdiarigas= " & DBSet(mvarNumDiarioGas, "N")
    SQL = SQL & ", ctacontragas= " & DBSet(mvarCtaContraGas, "T")
    SQL = SQL & ", raizctasocgas= " & DBSet(mvarRaizCtaSocGas, "T")
    SQL = SQL & ", ctaventasgas= " & DBSet(mvarCtaVentaGas, "T")
    SQL = SQL & ", letrasergas= " & DBSet(mvarNumSerieGas, "T")
    '[Monica]04/07/2013: gasoleo bonificado
    SQL = SQL & ", letrasergasB= " & DBSet(mvarNumSerieGasB, "T")
    SQL = SQL & ", increfacgas= " & DBSet(mvarIncreFacGas, "N")
    SQL = SQL & ", codivagas= " & DBSet(mvarCodIvaGas, "N")
    'facturas socios
    SQL = SQL & ", sercontafacsoc= '" & mvarSerContaFacSoc & "'"
    SQL = SQL & ", usucontafacsoc= " & DBSet(mvarUsuContaFacSoc, "T")
    SQL = SQL & ", pascontafacsoc= " & DBSet(mvarPasContaFacSoc, "T")
    SQL = SQL & ", numcontafacsoc= " & DBSet(mvarNumContaFacSoc, "N", "S")
    SQL = SQL & ", ctaretenfacsoc= " & DBSet(mvarCtaRetenFacSoc, "T")
    SQL = SQL & ", raizctafacsoc= " & DBSet(mvarRaizCtaFacSoc, "T")
    SQL = SQL & ", porcretenfacsoc= " & DBSet(mvarPorcReteFacSoc, "N", "S")
    
    'coarval
    SQL = SQL & ", sercontacv= '" & mvarSerContaCV & "'"
    SQL = SQL & ", usucontacv= " & DBSet(mvarUsuContaCV, "T")
    SQL = SQL & ", pascontacv= " & DBSet(mvarPasContaCV, "T")
    SQL = SQL & ", numcontacv= " & DBSet(mvarNumContaCV, "N", "S")
    SQL = SQL & ", ctabancocv= " & DBSet(mvarCtaBancoCV, "T")
    SQL = SQL & ", raizctasoccv= " & DBSet(mvarRaizCtaSocCV, "T")
    SQL = SQL & ", ctaventacv= " & DBSet(mvarCtaVentaCV, "T")
    SQL = SQL & ", ctaventatickcv= " & DBSet(mvarCtaVentaTickCV, "T")
    SQL = SQL & ", ctaclientickcv= " & DBSet(mvarCtaClienTickCV, "T")
    SQL = SQL & ", ctaventafaccv= " & DBSet(mvarCtaVentaFacCV, "T")
    SQL = SQL & ", ctaventafacincv= " & DBSet(mvarCtaVentaFacInCV, "T")
    SQL = SQL & ", letrasercv= " & DBSet(mvarLetraSerCV, "T")
    SQL = SQL & ", letraserfcv= " & DBSet(mvarLetraSerFCV, "T")
    SQL = SQL & ", letraserfincv= " & DBSet(mvarLetraSerFinCV, "T")
    SQL = SQL & ", codforpabancv= " & DBSet(mvarCodforpaBanCV, "T")
    SQL = SQL & ", codforpaconcv= " & DBSet(mvarCodforpaConCV, "T")
    SQL = SQL & ", numcontacvv= " & DBSet(mvarNumContaCVV, "N", "S")
    SQL = SQL & ", ctabancocvv= " & DBSet(mvarCtaBancoCVV, "T")
    SQL = SQL & ", letrasercvv= " & DBSet(mvarLetraSerCVV, "T")
    SQL = SQL & ", codforpabancvv= " & DBSet(mvarCodforpaBanCVV, "T")
    SQL = SQL & ", codforpaconcvv= " & DBSet(mvarCodforpaConCVV, "T")
    
    '[Monica]08/02/2012: arigesmail
    SQL = SQL & ", EnvioDesdeOutlook = " & DBSet(Me.EnvioDesdeOutlook, "B")
    SQL = SQL & ", arigesmail = " & DBSet(ExeEnvioMail, "T", "S")
    
    'aridoc
    SQL = SQL & ", hayaridoc= " & DBSet(mvarHayAridoc, "N") ' hay aridoc
    SQL = SQL & ", codcarpetafvar = " & DBSet(mvarCarpetaFac, "N") 'carpeta facturas varias
    SQL = SQL & ", codextension = " & DBSet(mvarExtension, "N") 'extension documento
    SQL = SQL & ", c1fvararidoc = " & DBSet(mvarC1Factura, "N") 'campo 1 factura
    SQL = SQL & ", c2fvararidoc = " & DBSet(mvarC2Factura, "N") 'campo 2 factura
    SQL = SQL & ", c3fvararidoc = " & DBSet(mvarC3Factura, "N") 'campo 3 factura
    SQL = SQL & ", c4fvararidoc = " & DBSet(mvarC4Factura, "N") 'campo 4 factura
    
    SQL = SQL & " WHERE codparam =" & Codigo & ";"

    conn.Execute SQL

    Modificar = 1
    Exit Function

EModificar:
    Mens = "Se ha producido un error." & vbCrLf
    Mens = Mens & "Número: " & Err.Number & vbCrLf
    Mens = Mens & "Descripción: " & Err.Description
    MsgBox Mens, vbExclamation
    Modificar = 0
End Function

