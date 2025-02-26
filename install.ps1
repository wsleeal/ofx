# Verifica se está rodando como administrador
$admin = [System.Security.Principal.WindowsPrincipal] [System.Security.Principal.WindowsIdentity]::GetCurrent()
$adminRole = [System.Security.Principal.WindowsBuiltInRole]::Administrator

if (-not $admin.IsInRole($adminRole)) {
    # Reexecuta o script com privilégios administrativos
    $scriptPath = $PSCommandPath
    Start-Process powershell -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File `"$scriptPath`"" -Verb RunAs
    exit
}

# Define os links de download
$gitInstallerUrl = "https://github.com/git-for-windows/git/releases/download/v2.48.1.windows.1/Git-2.48.1-64-bit.exe"
$pythonInstallerUrl = "https://www.python.org/ftp/python/3.13.2/python-3.13.2-amd64.exe"

# Define os caminhos temporários para os instaladores
$gitInstallerPath = "$env:TEMP\Git-Installer.exe"
$pythonInstallerPath = "$env:TEMP\Python-Installer.exe"

# Baixa os instaladores usando curl.exe
Write-Host "Baixando Git..."
Start-Process -FilePath "curl.exe" -ArgumentList "-L -o `"$gitInstallerPath`" `"$gitInstallerUrl`"" -Wait -NoNewWindow

Write-Host "Baixando Python..."
Start-Process -FilePath "curl.exe" -ArgumentList "-L -o `"$pythonInstallerPath`" `"$pythonInstallerUrl`"" -Wait -NoNewWindow

# Instala o Git silenciosamente
Write-Host "Instalando Git..."
Start-Process -FilePath $gitInstallerPath -ArgumentList "/VERYSILENT /NORESTART" -Wait

# Instala o Python silenciosamente
Write-Host "Instalando Python..."
Start-Process -FilePath $pythonInstallerPath -ArgumentList "/quiet InstallAllUsers=1 PrependPath=1" -Wait

# Aguarda a Instalacao e remove os instaladores
Remove-Item -Path $gitInstallerPath -Force
Remove-Item -Path $pythonInstallerPath -Force

# Aguarda a Instalacao do Python e obtém o caminho do executável
$pythonPath = (Get-Command python).Source
$pipPath = (Get-Command pip).Source

# Aguarda alguns segundos para garantir que o PATH seja atualizado
Start-Sleep -Seconds 5

# Verifica se o pip está instalado corretamente
Write-Host "Verificando Instalacao do pip..."
Start-Process -FilePath $pythonPath -ArgumentList "-m ensurepip --default-pip" -Wait -NoNewWindow

# Instala a biblioteca desejada
$lib = "git+https://github.com/wsleeal/ofx.git"
Write-Host "Instalando a biblioteca $lib..."
Start-Process -FilePath $pipPath -ArgumentList "install $lib --no-cache" -Wait -NoNewWindow

# Cria um atalho na área de trabalho
Write-Host "Criando atalho na area de trabalho..."
$WshShell = New-Object -ComObject WScript.Shell
$desktopPath = [System.IO.Path]::Combine([System.Environment]::GetFolderPath("Desktop"), "OFX.lnk")

$shortcut = $WshShell.CreateShortcut($desktopPath)
$shortcut.TargetPath = "pythonw.exe"
$shortcut.Arguments = "-m ofx"
$shortcut.WorkingDirectory = [System.IO.Path]::GetDirectoryName($pythonPath)
$shortcut.IconLocation = "$pythonPath,0"
$shortcut.Save()

Write-Host "Instalacao concluida!"

Read-Host "Pressione Enter para continuar..."