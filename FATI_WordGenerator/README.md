1. resourcesPath == path zu den Resources im Projekt (für jeden User unterschiedlich)
2. für die exe datei muss man das Program auf Release stellen und im Terminal folgendes eingeben: 

dotnet publish -c Release -r win-x64 --self-contained true -p:PublishSingleFile=true

Somit wird im Ordner in dem die .exe Datei abgespielt wird eine Datei mit dem Ordnernamen "Rechnungen" erstellt und mithilfe der Excel-Iput-Parametern ein neues Word-Dokument erstellt.