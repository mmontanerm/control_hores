
# Això és informació a la consola
Write-Host "Inicilitzant script de enviament de PDF's..."

# Paràmetres del Servidor SMTP
$SMTPServer = "smtp.gmail.com"  # has de canviar i posar el de office365 segurament
$SMTPPort = 587
$Usuario = "montaner.marc@gmail.com" # el teu mail
$Contraseña = "maho nwqh nxye vvzf" # el teu password
 
# Llegeix la carpeta de treball
$RutaCarpeta = $args[0]

if (-Not (Test-Path $RutaCarpeta)) {
    Write-Host "Error: La carpeta seleccionada no existeix."
    exit
}

Write-Host "Carpeta seleccionada: $RutaCarpeta"

# Buscar últim PDF modificat a la carpeta
$UltimoPDF = Get-ChildItem -Path $RutaCarpeta -Filter "*.pdf" -Recurse | 
             Sort-Object LastWriteTime -Descending | 
             Select-Object -First 1

if (-Not $UltimoPDF) {
    Write-Host "A la carpeta que has seleccionat no hi ha cap fitxer PDF!"
    exit
}

Write-Host "Últim PDF trobat: $($UltimoPDF.Name) - Modificat el $($UltimoPDF.LastWriteTime)"

# Solicitar el destinatari
$Destinatario = Read-Host "Introdueix el correu electrònic del destinatari"

# Validar el format del mail
if ($Destinatario -notmatch "^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$") {
    Write-Host "Error: Adreça de mail no válida."
    exit
}

Write-Host "Correu electrònic del destinatari: $Destinatario"

# Confirmar abans d'enviar
$Confirmacion = Read-Host "¿Enviar el PDF '$($UltimoPDF.Name)' a $Destinatario? (s/n)"
if ($Confirmacion -ne "s") {
    Write-Host "Enviament cancelat."
    exit
}

Write-Host "Preparant correu..."

# Dades del mail
$Asunto = "Últim document PDF a $($UltimoPDF.Directory.Name)"
$Mensaje = @"
Iep!

Aquí tens l'últim document PDF de la carpeta seleccionada.

📂 Carpeta: $RutaCarpeta
📄 Document: $($UltimoPDF.Name)
📅 Última modificació: $($UltimoPDF.LastWriteTime)

Salutacions,
El sistema automàtic del teu cor!
"@

# Aquí fem el correu
$Correo = New-Object System.Net.Mail.MailMessage
$Correo.From = $Usuario
$Correo.To.Add($Destinatario)
$Correo.Subject = $Asunto
$Correo.Body = $Mensaje
$Correo.IsBodyHtml = $false

# Adjuntar el PDF
$Correo.Attachments.Add($UltimoPDF.FullName)

# Configurar el client SMTP
$SMTP = New-Object Net.Mail.SmtpClient($SMTPServer, $SMTPPort)
$SMTP.EnableSSL = $true
$SMTP.Credentials = New-Object System.Net.NetworkCredential($Usuario, $Contraseña)

# Enviar el correu
Try {
    Write-Host "Enviant correu..."
    $SMTP.Send($Correo)
    Write-Host "Correu enviat correctament a $Destinatario amb el document $($UltimoPDF.Name)"
} Catch {
    Write-Host "Error en l'enviament del correu: $_"
}

Write-Host "Apreta una tecla ..."
while ($true) { Start-Sleep 1 }