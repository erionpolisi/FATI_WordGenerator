# publish.ps1

Write-Host "Publishing .NET Application..." -ForegroundColor Cyan

# Projekt publishen

dotnet publish -c Release -r win-x64 --self-contained true -p:PublishSingleFile=true

# Desktop Pfad

$desktop = [Environment]::GetFolderPath("Desktop")

# Zielordner

$output = "$desktop\FATI_WordGenerator"

# Ordner neu erstellen

if (Test-Path $output) {
Remove-Item $output -Recurse -Force
}

New-Item -ItemType Directory -Path $output | Out-Null

# Publish Ordner

$publishFolder = ".\bin\Release\net8.0\win-x64\publish"

# Dateien kopieren

Copy-Item "$publishFolder*" $output -Recurse

Write-Host ""
Write-Host "Publish completed!" -ForegroundColor Green
Write-Host "Output copied to Desktop:"
Write-Host $output
