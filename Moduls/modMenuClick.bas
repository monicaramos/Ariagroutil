Attribute VB_Name = "modMenuClick"
Option Explicit

Private Sub Construc(nom As String)
    MsgBox nom & ": en construcció..."
End Sub

' ******* DATOS BASICOS *********

Public Sub SubmnP_Generales_Click(Index As Integer)

    Select Case Index
        Case 1: frmConfParamGral.Show vbModal
                PonerDatosPpal
        Case 2: frmConfParamAplic.Show vbModal
        Case 3: frmConfParamRpt.Show vbModal
        Case 4: frmMantenusu.Show vbModal
        Case 6: End
    End Select
End Sub

' *******  AVNICS *********

Public Sub SubmnG_Avnics_Click(Index As Integer)
    Select Case Index
        Case 1: frmAvnics.Show vbModal
        Case 2: frmInfAvnics.Show vbModal
        Case 4: frmRenAvnics.Show vbModal
        Case 6: frmCalculInte.Show vbModal ' calculo de intereses
        Case 7: frmContaInte.Show vbModal  ' contabilizacion de intereses
        Case 9: frmCancelAvnic.Show vbModal ' cancelacion
        Case 10: frmMovAvnics.Show vbModal ' mantenimiento de movimientos avnics
        Case 12: frmAyudaMod123.Show vbModal    ' ayuda para el modelo 123
        Case 13: frmMod193.Show vbModal    ' grabacion de modelo 193
    End Select
End Sub

' *******  SEGUROS *********

Public Sub SubmnG_Seguros_Click(Index As Integer)
    Select Case Index
        Case 1: DesBloqueoManual ("TRASSEG")
                If Not BloqueoManual("TRASSEG", "1") Then
                    MsgBox "No se puede realizar la Importación de Datos. Hay otro usuario realizándolo.", vbExclamation
                    Screen.MousePointer = vbDefault
                Else
                    frmTrasDatosSeg.Show vbModal ' Importar Datos del AgroWeb
                End If
        Case 2: frmManLineas.Show vbModal ' Lineas de seguros
        Case 3: frmManPolizas.Show vbModal ' polizas
    End Select
End Sub

Public Sub SubmnS_IntegSeg_click(Index As Integer)
    Select Case Index
        Case 1: frmIntContaSeg.Show vbModal ' Importar Datos del AgroWeb
        Case 2: frmIntTesorSeg.Show vbModal ' Lineas de seguros
    End Select
End Sub

' *******  TELEFONIA *********

Public Sub SubmnG_Telefonia_Click(Index As Integer)
    Select Case Index
        Case 1:
                If vParamAplic.TipoFicheroTel = 0 Then
                    ' traspaso del fichero de telefonia de Catadau
                    frmTrasTele.Show vbModal
                Else
                    ' traspaso del fichero de telefonia de Bolbaite
                    frmTrasTeleBolb.Show vbModal
                End If
        Case 2: frmManTelef.Show vbModal
        Case 3: frmContabFactTel.Show
    End Select
End Sub

' *******  FACTURAS VARIAS *********

Public Sub SubmnG_FactVarias_Click(Index As Integer)
    Select Case Index
        Case 1: frmManSecciones.Show vbModal 'Mantenimiento de secciones
        Case 2: frmManConceptos.Show vbModal 'Mantenimiento de Conceptos de Facturacion
        Case 3: frmManFactVarias.Show vbModal 'Entrada de Facturas
        Case 4: frmFactVar.Show vbModal 'Impresión de Facturas
        Case 5: AbrirListadoOfer (315)
        Case 6: frmContabFact.Show vbModal 'Integracion Contable facturas varias
        Case 7: frmImpAridoc.Show vbModal 'Integracion aridoc
    End Select
End Sub

' *******  APORTACIONES *********

Public Sub SubmnG_Aportaciones_Click(Index As Integer)
    Select Case Index
        Case 1: frmManTipoApor.Show vbModal 'Mantenimiento de secciones
    End Select
End Sub

' *******  GASOLINERA *********

Public Sub SubmnG_Gasolinera_Click(Index As Integer)
    Select Case Index
        Case 1: frmTrasGasol.Show vbModal
        Case 2: frmHcoFactGas.Show vbModal
        Case 3: frmEstCliGas.Show vbModal
        Case 4: frmContabFactGas.Show vbModal
    End Select
End Sub

' *******  FACTURAS SOCIOS *********

Public Sub SubmnG_FactSocios_Click(Index As Integer)
    Select Case Index
        Case 1: frmManVariedad.Show vbModal 'Mantenimiento de variedades
        Case 2: frmManFactSocios.Show vbModal 'Mantenimiento de Facturas Socios
        Case 3: frmFactSoc.Show vbModal 'Reimpresion de Facturas Socios
        Case 4: frmContaFacSoc.Show vbModal 'Contabilizacion de Facturas Socios
        
        Case 6: frmListRetSoc.OpcionListado = 0 ' Listado de Retenciones a Socios
                frmListRetSoc.Show vbModal
        
        Case 7: frmListRetSoc.OpcionListado = 1 ' Grabacion del Modelo 190
                frmListRetSoc.Show vbModal
        
        Case 9: frmTrasFacLiq.Show vbModal 'Traspaso de Liquidaciones
        Case 10:  frmTrasFacAce.Show vbModal 'Traspaso de Facturas de Aceite
        Case 11:  frmTrasFacSub.Show vbModal 'Traspaso de Subvenciones de FEDEPROL
    
    End Select
End Sub

' *******  FACTURAS COARVAL/VARIOS *********

Public Sub SubmnG_FactCoarval_Click(Index As Integer)
    Select Case Index
        Case 1: frmTrasFacCV.Show vbModal 'Mantenimiento de variedades
        Case 2: frmManFactCV.Show vbModal 'Mantenimiento de facturas coarval
        Case 3: frmContabFactCV.Show vbModal ' Integracion contable
    
    End Select
End Sub

' *******  UTILIDADES *********

Public Sub SubmnE_Util_Click(Index As Integer)
    Select Case Index
        Case 1: frmCaracteresMB.Show vbModal ' comprobacion de caracteres de multibase
        Case 3: frmBackUP.Show vbModal
    End Select
End Sub

