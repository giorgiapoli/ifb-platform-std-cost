# Avvia l'app "Variazioni Standard Cost" in locale e la apre nel browser.
# Se l'app e' gia' in esecuzione la riusa (non ne avvia una seconda).
# Controlla che sulla porta risponda davvero l'app (e non un altro programma).
$dir  = Split-Path -Parent $MyInvocation.MyCommand.Path
$port = 8777
$url  = "http://localhost:$port"

function Test-App {
    try {
        $r = Invoke-WebRequest -Uri "$url/_stcore/health" -UseBasicParsing -TimeoutSec 2
        return ($r.Content -match "ok")
    } catch { return $false }
}

if (-not (Test-App)) {
    $py = (Get-Command python -ErrorAction SilentlyContinue).Source
    if (-not $py) { throw "Python non trovato nel PATH" }
    Start-Process -FilePath $py -WorkingDirectory $dir -WindowStyle Hidden -ArgumentList @(
        "-m", "streamlit", "run", "app.py",
        "--server.port", $port, "--server.address", "127.0.0.1", "--server.headless", "true",
        "--browser.gatherUsageStats", "false")
    for ($i = 0; $i -lt 60; $i++) {          # attende fino a 30 secondi che sia pronta
        Start-Sleep -Milliseconds 500
        if (Test-App) { break }
    }
}
Start-Process $url
