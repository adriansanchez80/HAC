Attribute VB_Name = "Expiration_date_copyright"
Option Explicit

Private Sub Workbook_Open()

Dim FechaCaducidad As Date

FechaCaducidad = #12/31/2016#

If FechaCaducidad > Date Then
MsgBox "El contenido de este libro Excel se encuentra protegido por el art. 97 texto refundido de la Ley de Propiedad Intelectual aprobado por Real Decreto Legislativo 1/1996, de 12 de abril." & vbCrLf & vbCrLf & "Por favor lee al menos una vez la declaraci�n de privacidad." & vbCrLf & "Dispone de: " & (FechaCaducidad - Date) & " dias de prueba" & vbCrLf & vbCrLf & Chr(13) & "M�s informaci�n: adrian.sanchez@meyss.es", vbInformation, "PER�ODO DE PRUEBA"
Else
'Mensaje cuando el per�odo de prueba ha vencido
MsgBox "HA VENCIDO EL PER�ODO DE PRUEBA :(" & Chr(13) & "El archivo se cerrar� inmediatamente" & Chr(13) & "M�s informaci�n: adrian.sanchez@meyss.es", vbCritical, "FIN DE LA PRUEBA"

Application.DisplayAlerts = False
ActiveWorkbook.ChangeFileAccess xlReadOnly
Kill ActiveWorkbook.FullName
ThisWorkbook.Close
End If

End Sub
