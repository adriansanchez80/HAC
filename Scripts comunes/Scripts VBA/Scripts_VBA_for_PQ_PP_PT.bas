Attribute VB_Name = "Scripts_VBA_for_PQ_PP_PT"

Option Explicit

'el modulo contiene código vba para seleccionar las carpetas y refrescar

Sub ObtenerDirectorio_W()
' en un modulo de codigo normal
    Dim getfolder As String
    Dim Folder As FileDialog
    On Error Resume Next ' por si el usuario pulsa {esc} y no selecciona nada :)) '
    ' Open the file dialog
    Set Folder = Application.FileDialog(msoFileDialogFolderPicker)
    With Folder
    .Title = "NÓMINAS: Seleccionar carpeta contenedora de archivos"
    .InitialFileName = "u:\1.- OS 2016\"
    .AllowMultiSelect = False
    .Show

    If .SelectedItems.Count <> 0 Then
    getfolder = .SelectedItems(1)
    Range("B2").Value = getfolder
    Else
    MsgBox "No se ha seleccionado ninguna carpeta.", 48, "Operación cancelada!!!"
    End If
End With
End Sub


Sub ObtenerDirectorio_HAC()
    Dim getfolder As String
    Dim Folder As FileDialog
    On Error Resume Next ' por si el usuario pulsa {esc} y no selecciona nada :)) '
    ' Open the file dialog
    Set Folder = Application.FileDialog(msoFileDialogFolderPicker)
    With Folder
    .Title = "HAC: Seleccionar carpeta contenedora de archivos"
    .InitialFileName = "u:\1.- OS 2016\"
    .AllowMultiSelect = False
    .Show

    If .SelectedItems.Count <> 0 Then
    getfolder = .SelectedItems(1)
    Range("B3").Value = getfolder
    Else
    MsgBox "No se ha seleccionado ninguna carpeta.", 48, "Operación cancelada!!!"
    End If
  End With
End Sub


Sub RefreshPQ()
'refresh all PQ content
Dim lTest As Long, cn As WorkbookConnection
On Error Resume Next
For Each cn In ThisWorkbook.Connections
lTest = InStr(1, cn.OLEDBConnection.Connection, "Provider=Microsoft.Mashup.OleDb.1")
If Err.Number <> 0 Then
Err.Clear
Exit For
End If
If lTest > 0 Then cn.Refresh
Next cn
End Sub

Sub RefreshPP_PT()
'refresh all PP and PTs' content
ActiveWorkbook.RefreshAll
End Sub
