Dim InventorApp, Doc, StepAddIn, TransContext, Options, DataMedium
On Error Resume Next

' Obtener el archivo seleccionado
Set objArgs = WScript.Arguments
If objArgs.Count = 0 Then
    MsgBox "Error: No se ha seleccionado un archivo IPT o IAM.", 48, "Exportación a STEP"
    WScript.Quit
End If

FilePath = objArgs(0)
OutputFile = Replace(FilePath, ".ipt", ".step")
OutputFile = Replace(OutputFile, ".iam", ".step")

' Verificar si Inventor está en ejecución o iniciarlo
Set InventorApp = GetObject(, "Inventor.Application")
If Err.Number <> 0 Then
    Set InventorApp = CreateObject("Inventor.Application")
    InventorApp.Visible = False
    InventorStartedByScript = True
    Err.Clear
Else
    InventorStartedByScript = False
End If
On Error GoTo 0 ' Restablecer el control de errores

' Abrir archivo en segundo plano
Set Doc = InventorApp.Documents.Open(FilePath, False)

If Doc Is Nothing Then
    MsgBox "Error: No se pudo abrir el archivo " & FilePath, 48, "Error"
    WScript.Quit
End If

' Buscar el Add-in de exportación STEP dinámicamente
Set StepAddIn = Nothing
For Each AddIn In InventorApp.ApplicationAddIns
    If InStr(1, AddIn.DisplayName, "STEP", vbTextCompare) > 0 Then
        Set StepAddIn = AddIn
        Exit For
    End If
Next

If StepAddIn Is Nothing Then
    MsgBox "Error: No se encontró el Add-in de exportación STEP.", 48, "Error"
    WScript.Quit
End If

' Configurar exportación
Set TransContext = InventorApp.TransientObjects.CreateTranslationContext
Set Options = InventorApp.TransientObjects.CreateNameValueMap
Set DataMedium = InventorApp.TransientObjects.CreateDataMedium
DataMedium.FileName = OutputFile

' Intentar exportar a STEP con manejo de errores
On Error Resume Next
Doc.SaveAs OutputFile, True
If Err.Number <> 0 Then
    MsgBox "Error en SaveCopyAs: " & Err.Description & " (Código " & Err.Number & ")", 48, "Error"
    Err.Clear
    WScript.Quit
End If
On Error GoTo 0 ' Restablecer el control de errores

' Esperar hasta 10 segundos para verificar si el archivo se ha creado
Set fso = CreateObject("Scripting.FileSystemObject")
Dim TiempoInicio, TiempoMax
TiempoInicio = Timer
TiempoMax = 10

Do While Not fso.FileExists(OutputFile) And (Timer - TiempoInicio) < TiempoMax
    WScript.Sleep 1000
Loop

If fso.FileExists(OutputFile) Then
    MsgBox "Archivo exportado correctamente: " & OutputFile, 64, "Exportación Completa"
Else
    MsgBox "Error: El archivo no se ha creado correctamente. Verifica los permisos de la carpeta o la ruta de salida.", 48, "Error"
End If

Set fso = Nothing

' Cerrar documento sin guardar cambios
Doc.Close True

' Cerrar Inventor solo si fue iniciado por el script
If InventorStartedByScript Then
    InventorApp.Quit
End If

Set InventorApp = Nothing
