Attribute VB_Name = "ModContadores"
Option Explicit







'Private Function ObtenerContador() As Boolean
''Obtiene el contador para la venta
'Dim cadFec As String
'Dim vCont As CContador
'Dim b As Boolean
'
'        cadFec = Trim(Text1(1).Text)
'        If cadFec = "" Then
'            MsgBox "El campo fecha de la venta debe tener valor para obtener un contador.", vbExclamation
'            b = False
'        Else
'            Set vCont = New CContador
'            If vCont.ConseguirContador(cadFec, "pventa_a", "pventa_b", True) Then
'                If vCont.AnyoActual Then
'                    Text1(0).Text = vCont.Contador1
'                Else
'                    Text1(0).Text = vCont.Contador2
'                End If
'                FormateaCampo Text1(0)
'                b = True
'            End If
'            Set vCont = Nothing
'        End If
'
'        ObtenerContador = b
'End Function
