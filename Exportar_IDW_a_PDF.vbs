Dim InventorApp, Doc, PDFAddIn, TransContext, Options, DataMedium
On Error Resume Next

' Obtener el archivo seleccionado desde el Explorador de Windows
Set objArgs = WScript.Arguments
If objArgs.Count = 0 Then
    MsgBox "Error: No se ha seleccionado un archivo IDW.", 48, "Exportación a PDF"
    WScript.Quit
End If

FilePath = objArgs(0)
OutputFile = Replace(FilePath, ".idw", ".pdf")

' Iniciar Autodesk Inventor si no está abierto
Set InventorApp = GetObject(, "Inventor.Application")
If Err.Number <> 0 Then
    Set InventorApp = CreateObject("Inventor.Application")
    InventorApp.Visible = False
    InventorStartedByScript = True
    Err.Clear
Else
    InventorStartedByScript = False
End If
On Error GoTo 0 ' Restablecer control de errores

' Abrir el archivo en Inventor
Set Doc = InventorApp.Documents.Open(FilePath, False)
If Doc Is Nothing Then
    MsgBox "Error: No se pudo abrir el archivo " & FilePath, 48, "Error"
    WScript.Quit
End If

' Obtener el Add-in de PDF en Inventor
Set PDFAddIn = Nothing
For Each AddIn In InventorApp.ApplicationAddIns
    If InStr(1, AddIn.DisplayName, "PDF", vbTextCompare) > 0 Then
        Set PDFAddIn = AddIn
        Exit For
    End If
Next

If PDFAddIn Is Nothing Then
    MsgBox "Error: No se encontró el Add-in de exportación a PDF.", 48, "Error"
    WScript.Quit
End If

' Configurar exportación a PDF
Set TransContext = InventorApp.TransientObjects.CreateTranslationContext
Set Options = InventorApp.TransientObjects.CreateNameValueMap
Set DataMedium = InventorApp.TransientObjects.CreateDataMedium
DataMedium.FileName = OutputFile

' Intentar exportar a PDF con manejo de errores
On Error Resume Next
Doc.SaveAs OutputFile, True
If Err.Number <> 0 Then
    MsgBox "Error en SaveCopyAs: " & Err.Description & " (Código " & Err.Number & ")", 48, "Error"
    Err.Clear
    WScript.Quit
End If
On Error GoTo 0

' Verificar si el archivo PDF se ha creado
Set fso = CreateObject("Scripting.FileSystemObject")
Dim TiempoInicio, TiempoMax
TiempoInicio = Timer
TiempoMax = 10 ' Segundos máximos de espera

Do While Not fso.FileExists(OutputFile) And (Timer - TiempoInicio) < TiempoMax
    WScript.Sleep 1000
Loop

If fso.FileExists(OutputFile) Then
    MsgBox "Archivo exportado correctamente: " & OutputFile, 64, "Exportación Completa"
Else
    MsgBox "Error: El archivo no se ha creado correctamente. Verifica permisos y la ruta de salida.", 48, "Error"
End If

Set fso = Nothing

' Cerrar documento sin guardar cambios
Doc.Close True

' Cerrar Inventor solo si fue iniciado por el script
If InventorStartedByScript Then
    InventorApp.Quit
End If

Set InventorApp = Nothing
