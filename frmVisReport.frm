VERSION 5.00
Object = "{8767A745-088E-4CA6-8594-073D6D2DE57A}#9.2#0"; "crviewer9.dll"
Begin VB.Form frmVisReport 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Visor de informes"
   ClientHeight    =   4260
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5925
   Icon            =   "frmVisReport.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4260
   ScaleWidth      =   5925
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin CRVIEWER9LibCtl.CRViewer9 CRViewer1 
      Height          =   3015
      Left            =   60
      TabIndex        =   0
      Top             =   240
      Width           =   4695
      lastProp        =   600
      _cx             =   8281
      _cy             =   5318
      DisplayGroupTree=   -1  'True
      DisplayToolbar  =   -1  'True
      EnableGroupTree =   -1  'True
      EnableNavigationControls=   -1  'True
      EnableStopButton=   -1  'True
      EnablePrintButton=   -1  'True
      EnableZoomControl=   -1  'True
      EnableCloseButton=   -1  'True
      EnableProgressControl=   -1  'True
      EnableSearchControl=   -1  'True
      EnableRefreshButton=   -1  'True
      EnableDrillDown =   -1  'True
      EnableAnimationControl=   -1  'True
      EnableSelectExpertButton=   0   'False
      EnableToolbar   =   -1  'True
      DisplayBorder   =   -1  'True
      DisplayTabs     =   -1  'True
      DisplayBackgroundEdge=   -1  'True
      SelectionFormula=   ""
      EnablePopupMenu =   0   'False
      EnableExportButton=   -1  'True
      EnableSearchExpertButton=   0   'False
      EnableHelpButton=   0   'False
      LaunchHTTPHyperlinksInNewBrowser=   -1  'True
      EnableLogonPrompts=   -1  'True
   End
End
Attribute VB_Name = "frmVisReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' +-+-+-+-+-+-+-+-+-+-+-
' +-+- Autor: DAVID +-+-
' +-+-+-+-+-+-+-+-+-+-+-

Option Explicit
Public Informe As String
Public InfConta As Boolean 'Enlazar a la Contabilidas

'SubInforme con conexion a la contabilidad. Conectar a las
'tablas de la BDatos correspondiente a la empresa: conta1, conta2, etc.

Public Facturas As Boolean
Public ConSubInforme As Boolean 'Si tiene subinforme ejecta la funcion AbrirSubInforme para enlazar esta a la BD correspondiente
Public Contabilidad As Byte ' indica la contabilidad que es
                            'cconta = 2 --> contabilidad de avnics
                            'cconta = 3 --> contabilidad de seguros
                            'cconta = 4 --> contabilidad de telefonia
                            'cconta = 5 --> contabilidad de gasolinera
                            'cconta = 6 --> contabilidad de facturas socios
                            

'estas varriables las trae del formulario de impresion
Public FormulaSeleccion As String
Public SoloImprimir As Boolean
Public OtrosParametros As String   ' El grupo acaba en |                            ' param1=valor1|param2=valor2|
Public NumeroParametros As Integer   'Cuantos parametros hay.  EMPRESA(EMP) no es parametro. Es fijo en todos los informes
Public MostrarTree As Boolean
Public Opcion As Integer
Public ExportarPDF As Boolean
Public EstaImpreso As Boolean
Public SubInformeConta As String

Public FicheroPDF As String


Dim mapp As CRAXDRT.Application
Dim mrpt As CRAXDRT.Report
Dim smrpt As CRAXDRT.Report

Dim Argumentos() As String
Dim PrimeraVez As Boolean



Private Sub CRViewer1_PrintButtonClicked(UseDefault As Boolean)
'    CRViewer1.PrintReport
    EstaImpreso = True
End Sub

Private Sub Form_Activate()
    If PrimeraVez Then
        PrimeraVez = False
        If SoloImprimir Or Me.ExportarPDF Then
            Screen.MousePointer = vbHourglass
            Unload Me
        End If
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
Dim i As Integer
On Error GoTo Err_Carga
    
    Screen.MousePointer = vbHourglass
    Set mapp = CreateObject("CrystalRuntime.Application")
'    Informe = "C:\Programas\Ariges 4\Informes\rptMarcas.rpt"
    Set mrpt = mapp.OpenReport(Informe)
       
       
       
    'Conectar a la BD de la Empresa
    '####Descomentar
    For i = 1 To mrpt.Database.Tables.Count
       mrpt.Database.Tables(i).SetLogOnInfo "vAriagroutil", , "root", "aritel"
       If InStr(1, mrpt.Database.Tables(i).Name, "_") = 0 Then
'               mrpt.Database.Tables(i).Location = vEmpresa.BDAriges & "." & mrpt.Database.Tables(i).Name
               mrpt.Database.Tables(i).Location = mrpt.Database.Tables(i).Name
       End If
    Next i

    If InfConta Then
        For i = 1 To mrpt.Database.Tables.Count
           If vParamAplic.ContabilidadNueva Then
                mrpt.Database.Tables(i).SetLogOnInfo "vconta", "ariconta" & vParamAplic.NumeroConta, "root", "aritel"
            
                If InStr(1, mrpt.Database.Tables(i).Name, "_") = 0 Then
                       mrpt.Database.Tables(i).Location = "ariconta" & vParamAplic.NumeroConta & "." & mrpt.Database.Tables(i).Name
                End If
           Else
                mrpt.Database.Tables(i).SetLogOnInfo "vconta", "conta" & vParamAplic.NumeroConta, "root", "aritel"
            
                If InStr(1, mrpt.Database.Tables(i).Name, "_") = 0 Then
                       mrpt.Database.Tables(i).Location = "conta" & vParamAplic.NumeroConta & "." & mrpt.Database.Tables(i).Name
                End If
           End If
        Next i
    End If
    
'    If SubInformeConta <> "" Then
'        Set smrpt = mrpt.OpenSubreport(SubInformeConta)
'        For i = 1 To smrpt.Database.Tables.Count
'            smrpt.Database.Tables(i).SetLogOnInfo "vConta", "conta" & vParamAplic.NumeroConta, vParamAplic.UsuarioConta, vParamAplic.PasswordConta
'            smrpt.Database.Tables(i).Location = "conta" & vParamAplic.NumeroConta & "." & smrpt.Database.Tables(i).Name
'        Next i
'    End If
    If ConSubInforme Then AbrirSubreport
    
    
    PrimeraVez = True
    
    CargaArgumentos
    CRViewer1.EnableGroupTree = MostrarTree
    CRViewer1.DisplayGroupTree = MostrarTree
    mrpt.RecordSelectionFormula = FormulaSeleccion
    
    'poner en la select del subinforme los mismos criterios que los del informe
   'If ConSubInforme Then SelectSubreport
   If ConSubInforme Then SelectSubreport
    
 
    
'    If Opcion = 6 Then
'        Dim crD As CRAXDRT.DatabaseFieldDefinition
'        Dim crF As CRAXDRT.FieldObject
'        Dim crS As CRAXDRT.Section
'        Set crD = mrpt.Database.Tables(1).Fields(6)
'        mrpt.AddGroup 0, crD, crGCAnyValue, crAscendingOrder
'        Set crS = mrpt.Sections.Item("GH")
'        Set crF = crS.AddFieldObject("{sfamia.nomfamia}", 100, 0)
''        mrpt.RecordSortFields.Item(3).Parent = mrpt.RecordSortFields.Item(1)
''        mrpt.RecordSortFields.Item(3).Parent = mrpt.RecordSortFields.Item(2)
'    End If
   
    
    'Si es a mail
    If Me.ExportarPDF Then
        Exportar
        Exit Sub
    End If
    
'[Monica]15/11/2013: para el aridoc de Rafa
    If FicheroPDF <> "" Then
        mrpt.ExportOptions.DestinationType = crEDTDiskFile
        mrpt.ExportOptions.DiskFileName = FicheroPDF
        mrpt.ExportOptions.FormatType = crEFTPortableDocFormat
        mrpt.ExportOptions.PDFExportAllPages = True
        mrpt.Export False
        Exit Sub
    End If
    
    
    
    'lOS MARGENES
    PonerMargen
    
    EstaImpreso = False
    CRViewer1.ReportSource = mrpt
   
    
    If SoloImprimir Then
        mrpt.PrintOut False
        EstaImpreso = True
    Else
        CRViewer1.ViewReport
    End If
    Exit Sub
Err_Carga:
    MsgBox "Error: " & Err.Number & vbCrLf & Err.Description & vbCrLf & Informe, vbCritical
    Set mapp = Nothing
    Set mrpt = Nothing
    Set smrpt = Nothing
End Sub

Private Sub Form_Resize()
    CRViewer1.Top = 0
    CRViewer1.Left = 0
    CRViewer1.Height = ScaleHeight
    CRViewer1.Width = ScaleWidth
End Sub

Private Sub CargaArgumentos()
Dim Parametro As String
Dim i As Integer
    'El primer parametro es el nombre de la empresa para todas las empresas
    ' Por lo tanto concaatenaremos con otros parametros
    ' Y sumaremos uno
    'Luego iremos recogiendo para cada formula su valor y viendo si esta en
    ' La cadena de parametros
    'Si esta asignaremos su valor
    
'    OtrosParametros = "|Emp= """ & vEmpresa.nomempre & """|" & OtrosParametros
Select Case NumeroParametros
Case 0
    '====Comenta: LAura
    'Solo se vacian los campos de formula que empiezan con "p" ya que estas
    'formulas se corresponden con paso de parametros al Report
    For i = 1 To mrpt.FormulaFields.Count
        If Left(Mid(mrpt.FormulaFields(i).Name, 3), 1) = "p" Then
            mrpt.FormulaFields(i).Text = """"""
        End If
    Next i
    '====
Case 1
    
    For i = 1 To mrpt.FormulaFields.Count
        Parametro = mrpt.FormulaFields(i).Name
        Parametro = Mid(Parametro, 3)  'Quitamos el {@
        Parametro = Mid(Parametro, 1, Len(Parametro) - 1) ' el } del final
        'Debug.Print Parametro
        If DevuelveValor(Parametro) Then
            mrpt.FormulaFields(i).Text = Parametro
        Else
'            mrpt.FormulaFields(I).Text = """"""
        End If
    Next i
    
Case Else
    NumeroParametros = NumeroParametros + 1
    
    For i = 1 To mrpt.FormulaFields.Count
        Parametro = mrpt.FormulaFields(i).Name
        Parametro = Mid(Parametro, 3)  'Quitamos el {@
        Parametro = Mid(Parametro, 1, Len(Parametro) - 1) ' el } del final
        If DevuelveValor(Parametro) Then
            mrpt.FormulaFields(i).Text = Parametro
        End If
    Next i
'    mrpt.RecordSelectionFormula
End Select
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set mrpt = Nothing
    Set mapp = Nothing
    Set smrpt = Nothing
End Sub


Private Function DevuelveValor(ByRef Valor As String) As Boolean
Dim i As Integer
Dim J As Integer

    Valor = "|" & Valor & "="
    DevuelveValor = False
    i = InStr(1, OtrosParametros, Valor, vbTextCompare)
    If i > 0 Then
        i = i + Len(Valor)
        J = InStr(i, OtrosParametros, "|")
        If J > 0 Then
            Valor = Mid(OtrosParametros, i, J - i)
            If Valor = "" Then
                Valor = " "
            Else
                CompruebaComillas Valor
            End If
            DevuelveValor = True
        End If
    End If
End Function


Private Sub CompruebaComillas(ByRef Valor1 As String)
Dim Aux As String
Dim J As Integer
Dim i As Integer

    If Mid(Valor1, 1, 1) = Chr(34) Then
        'Tiene comillas. Con lo cual tengo k poner las dobles
        Aux = Mid(Valor1, 2, Len(Valor1) - 2)
        i = -1
        Do
            J = i + 2
            i = InStr(J, Aux, """")
            If i > 0 Then
              Aux = Mid(Aux, 1, i - 1) & """" & Mid(Aux, i)
            End If
        Loop Until i = 0
        Aux = """" & Aux & """"
        Valor1 = Aux
    End If
End Sub

Private Sub Exportar()
    mrpt.ExportOptions.DiskFileName = App.path & "\docum.pdf"
    mrpt.ExportOptions.DestinationType = crEDTDiskFile
    mrpt.ExportOptions.PDFExportAllPages = True
    mrpt.ExportOptions.FormatType = crEFTPortableDocFormat
    mrpt.Export False
    'Si ha generado bien entonces
    CadenaDesdeOtroForm = "OK"
End Sub

Private Sub PonerMargen()
Dim Cad As String
Dim i As Integer
    On Error GoTo EPon
    Cad = Dir(App.path & "\*.mrg")
    If Cad <> "" Then
        i = InStr(1, Cad, ".")
        If i > 0 Then
            Cad = Mid(Cad, 1, i - 1)
            If IsNumeric(Cad) Then
                If Val(Cad) > 4000 Then Cad = "4000"
                If Val(Cad) > 0 Then
                    mrpt.BottomMargin = mrpt.BottomMargin + Val(Cad)
                End If
            End If
        End If
    End If
    Exit Sub
EPon:
    Err.Clear
End Sub



Private Sub SelectSubreport()
''Para cada subReport que encuentre en el Informe pone a la del subReport
''la select del report
'Dim crxSection As CRAXDRT.Section
'Dim crxObject As Object
'Dim crxSubreportObject As CRAXDRT.SubreportObject
'Dim i As Integer
'
'    For Each crxSection In mrpt.Sections
'        For Each crxObject In crxSection.ReportObjects
'             If TypeOf crxObject Is SubreportObject Then
'                Set crxSubreportObject = crxObject
''                If crxSubreportObject.SubreportName <> SubInformeConta Then
'                    Set smrpt = mrpt.OpenSubreport(crxSubreportObject.SubreportName)
'                    smrpt.RecordSelectionFormula = mrpt.RecordSelectionFormula
'
''                    For i = 1 To smrpt.Database.Tables.Count
''                         smrpt.Database.Tables(i).SetLogOnInfo "vAriagroutil", , "root", "aritel"
''                         If (InStr(1, smrpt.Database.Tables(i).Name, "_") = 0) Then
''                            smrpt.Database.Tables(i).Location = smrpt.Database.Tables(i).Name
''                         End If
''                    Next i
'                'End If
'             End If
'        Next crxObject
'    Next crxSection
''
'    Set crxSubreportObject = Nothing

End Sub




Private Sub AbrirSubreport()
'Para cada subReport que encuentre en el Informe pone las tablas del subReport
'apuntando a la BD correspondiente
Dim crxSection As CRAXDRT.Section
Dim crxObject As Object
Dim crxSubreportObject As CRAXDRT.SubreportObject
Dim i As Byte

    For Each crxSection In mrpt.Sections
        For Each crxObject In crxSection.ReportObjects
             If TypeOf crxObject Is SubreportObject Then
                Set crxSubreportObject = crxObject
                Set smrpt = mrpt.OpenSubreport(crxSubreportObject.SubreportName)
                For i = 1 To smrpt.Database.Tables.Count 'para cada tabla
                    '------ A�ade Laura: 09/06/2005
                    If smrpt.Database.Tables(i).ConnectionProperties.Item("DSN") = "vAriagroutil" Then
                        smrpt.Database.Tables(i).SetLogOnInfo "vAriagroutil", "ariagroutil", "root", "aritel"
                        If (InStr(1, smrpt.Database.Tables(i).Name, "_") = 0) Then
                           smrpt.Database.Tables(i).Location = "ariagroutil" & "." & smrpt.Database.Tables(i).Name
                        End If
                    ElseIf smrpt.Database.Tables(i).ConnectionProperties.Item("DSN") = "vConta" Then
                        If vParamAplic.ContabilidadNueva Then
                        
                            If Not Facturas Then
                                'monica 05/06/2007
                                Select Case Contabilidad
                                    Case cConta
                                        smrpt.Database.Tables(i).SetLogOnInfo "vConta", "ariconta" & vParamAplic.NumeroConta, vParamAplic.UsuarioConta, vParamAplic.PasswordConta
                                    Case cContaSeg
                                        smrpt.Database.Tables(i).SetLogOnInfo "vConta", "ariconta" & vParamAplic.NumeroContaSeg, vParamAplic.UsuarioContaSeg, vParamAplic.PasswordContaSeg
                                    Case cContaTel
                                        smrpt.Database.Tables(i).SetLogOnInfo "vConta", "ariconta" & vParamAplic.NumeroContaTel, vParamAplic.UsuarioContaTel, vParamAplic.PasswordContaTel
                                    Case cContaGas
                                        smrpt.Database.Tables(i).SetLogOnInfo "vConta", "ariconta" & vParamAplic.NumeroContaGas, vParamAplic.UsuarioContaGas, vParamAplic.PasswordContaGas
                                    Case cContaFacSoc
                                        smrpt.Database.Tables(i).SetLogOnInfo "vConta", "ariconta" & vParamAplic.NumeroContaFacSoc, vParamAplic.UsuarioContaGas, vParamAplic.PasswordContaFacSoc
                                    Case cContaCV
                                        smrpt.Database.Tables(i).SetLogOnInfo "vConta", "ariconta" & vParamAplic.NumeroContaCV, vParamAplic.UsuarioContaCV, vParamAplic.PasswordContaCV
                                    Case cContaCVV
                                        smrpt.Database.Tables(i).SetLogOnInfo "vConta", "ariconta" & vParamAplic.NumeroContaCVV, vParamAplic.UsuarioContaCV, vParamAplic.PasswordContaCV
                                End Select
                                If (InStr(1, smrpt.Database.Tables(i).Name, "_") = 0) Then
                                    Select Case Contabilidad
                                        Case cConta
                                           smrpt.Database.Tables(i).Location = "ariconta" & vParamAplic.NumeroConta & "." & smrpt.Database.Tables(i).Name
                                        Case cContaSeg
                                           smrpt.Database.Tables(i).Location = "ariconta" & vParamAplic.NumeroContaSeg & "." & smrpt.Database.Tables(i).Name
                                        Case cContaTel
                                           smrpt.Database.Tables(i).Location = "ariconta" & vParamAplic.NumeroContaTel & "." & smrpt.Database.Tables(i).Name
                                        Case cContaGas
                                           smrpt.Database.Tables(i).Location = "ariconta" & vParamAplic.NumeroContaGas & "." & smrpt.Database.Tables(i).Name
                                        Case cContaFacSoc
                                           smrpt.Database.Tables(i).Location = "ariconta" & vParamAplic.NumeroContaFacSoc & "." & smrpt.Database.Tables(i).Name
                                        Case cContaCV
                                           smrpt.Database.Tables(i).Location = "ariconta" & vParamAplic.NumeroContaCV & "." & smrpt.Database.Tables(i).Name
                                        Case cContaCVV
                                           smrpt.Database.Tables(i).Location = "ariconta" & vParamAplic.NumeroContaCVV & "." & smrpt.Database.Tables(i).Name
                                    End Select
                                        
                                End If
                            Else ' venimos de facturas varias
    '                            smrpt.Database.Tables(i).SetLogOnInfo vParamAplic.ServidorContaFac, "conta" & Contabilidad, vParamAplic.UsuarioContaFac, vParamAplic.PasswordContaFac
                                smrpt.Database.Tables(i).SetLogOnInfo "vConta", "ariconta" & Contabilidad, vParamAplic.UsuarioContaFac, vParamAplic.PasswordContaFac
                                If (InStr(1, smrpt.Database.Tables(i).Name, "_") = 0) Then
                                    If smrpt.Database.Tables(i).Name = "sforpa" Then
                                        smrpt.Database.Tables(i).Location = "ariconta" & Contabilidad & ".formapago"
                                    Else
                                        smrpt.Database.Tables(i).Location = "ariconta" & Contabilidad & "." & smrpt.Database.Tables(i).Name
                                    End If
                                End If
                            End If
                        
                        Else
                    
                            If Not Facturas Then
                                'monica 05/06/2007
                                Select Case Contabilidad
                                    Case cConta
                                        smrpt.Database.Tables(i).SetLogOnInfo "vConta", "conta" & vParamAplic.NumeroConta, vParamAplic.UsuarioConta, vParamAplic.PasswordConta
                                    Case cContaSeg
                                        smrpt.Database.Tables(i).SetLogOnInfo "vConta", "conta" & vParamAplic.NumeroContaSeg, vParamAplic.UsuarioContaSeg, vParamAplic.PasswordContaSeg
                                    Case cContaTel
                                        smrpt.Database.Tables(i).SetLogOnInfo "vConta", "conta" & vParamAplic.NumeroContaTel, vParamAplic.UsuarioContaTel, vParamAplic.PasswordContaTel
                                    Case cContaGas
                                        smrpt.Database.Tables(i).SetLogOnInfo "vConta", "conta" & vParamAplic.NumeroContaGas, vParamAplic.UsuarioContaGas, vParamAplic.PasswordContaGas
                                    Case cContaFacSoc
                                        smrpt.Database.Tables(i).SetLogOnInfo "vConta", "conta" & vParamAplic.NumeroContaFacSoc, vParamAplic.UsuarioContaGas, vParamAplic.PasswordContaFacSoc
                                    Case cContaCV
                                        smrpt.Database.Tables(i).SetLogOnInfo "vConta", "conta" & vParamAplic.NumeroContaCV, vParamAplic.UsuarioContaCV, vParamAplic.PasswordContaCV
                                    Case cContaCVV
                                        smrpt.Database.Tables(i).SetLogOnInfo "vConta", "conta" & vParamAplic.NumeroContaCVV, vParamAplic.UsuarioContaCV, vParamAplic.PasswordContaCV
                                End Select
                                If (InStr(1, smrpt.Database.Tables(i).Name, "_") = 0) Then
                                    Select Case Contabilidad
                                        Case cConta
                                           smrpt.Database.Tables(i).Location = "conta" & vParamAplic.NumeroConta & "." & smrpt.Database.Tables(i).Name
                                        Case cContaSeg
                                           smrpt.Database.Tables(i).Location = "conta" & vParamAplic.NumeroContaSeg & "." & smrpt.Database.Tables(i).Name
                                        Case cContaTel
                                           smrpt.Database.Tables(i).Location = "conta" & vParamAplic.NumeroContaTel & "." & smrpt.Database.Tables(i).Name
                                        Case cContaGas
                                           smrpt.Database.Tables(i).Location = "conta" & vParamAplic.NumeroContaGas & "." & smrpt.Database.Tables(i).Name
                                        Case cContaFacSoc
                                           smrpt.Database.Tables(i).Location = "conta" & vParamAplic.NumeroContaFacSoc & "." & smrpt.Database.Tables(i).Name
                                        Case cContaCV
                                           smrpt.Database.Tables(i).Location = "conta" & vParamAplic.NumeroContaCV & "." & smrpt.Database.Tables(i).Name
                                        Case cContaCVV
                                           smrpt.Database.Tables(i).Location = "conta" & vParamAplic.NumeroContaCVV & "." & smrpt.Database.Tables(i).Name
                                    End Select
                                        
                                End If
                            Else ' venimos de facturas varias
    '                            smrpt.Database.Tables(i).SetLogOnInfo vParamAplic.ServidorContaFac, "conta" & Contabilidad, vParamAplic.UsuarioContaFac, vParamAplic.PasswordContaFac
                                smrpt.Database.Tables(i).SetLogOnInfo "vConta", "conta" & Contabilidad, vParamAplic.UsuarioContaFac, vParamAplic.PasswordContaFac
                                If (InStr(1, smrpt.Database.Tables(i).Name, "_") = 0) Then
                                    smrpt.Database.Tables(i).Location = "conta" & Contabilidad & "." & smrpt.Database.Tables(i).Name
                                End If
                            End If
                        End If
                    End If
                    '------
                Next i
                
                Set crxSubreportObject = Nothing

             
             End If
        Next crxObject
    Next crxSection

    Set crxSubreportObject = Nothing
End Sub

