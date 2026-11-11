# Avvia l'app "Variazioni Standard Cost" in locale e la apre nel browser.
# - Se l'app e' gia' in esecuzione la riusa (non ne avvia una seconda) e apre subito la pagina.
# - Se non c'e', apre SUBITO una pagina di attesa (indirizzo http://localhost:8778, quindi sempre nel browser) che passa da
#   sola all'app appena e' pronta; l'app parte nel frattempo.
$dir  = Split-Path -Parent $MyInvocation.MyCommand.Path
$port = 8777
$url  = "http://localhost:$port"
$salute = "http://127.0.0.1:$port/_stcore/health"

function Test-App {
    try {
        $r = Invoke-WebRequest -Uri $salute -UseBasicParsing -TimeoutSec 2
        return ($r.Content -match "ok")
    } catch { return $false }
}

if (Test-App) {
    Start-Process $url
    exit
}

$py = (Get-Command python -ErrorAction SilentlyContinue).Source
if (-not $py) { $py = "C:\Users\giorgia.poli\AppData\Local\Programs\Python\Python313\python.exe" }

# piccolo server per la pagina di attesa
$pagina = @"
<!DOCTYPE html><html lang="it"><head><meta charset="utf-8"><title>Standard Cost - avvio</title>
<style>
body{font-family:Segoe UI,Arial,sans-serif;background:#F6F7F9;color:#14213D;display:flex;align-items:center;justify-content:center;height:100vh;margin:0}
.c{background:#fff;border:1px solid #E4E7EC;border-radius:16px;padding:36px 48px;text-align:center;box-shadow:0 8px 28px rgba(16,24,40,.08)}
.s{width:42px;height:42px;border:5px solid #CCFBF1;border-top-color:#0F766E;border-radius:50%;margin:0 auto 18px;animation:r 1s linear infinite}
@keyframes r{to{transform:rotate(360deg)}} h1{font-size:1.3rem;margin:0 0 6px} p{color:#667085;margin:0}
</style></head><body><div class="c"><div class="s"></div><h1>Sto avviando Standard Cost...</h1>
<p>Ci vuole qualche secondo: la pagina si apre da sola.</p></div>
<script>
var u="$url", n=0;
function prova(){ fetch(u+"/_stcore/health",{mode:"no-cors",cache:"no-store"}).then(function(){location.replace(u);}).catch(function(){ if(++n<300) setTimeout(prova,500); else document.querySelector("p").textContent="Non riesco ad avviare l'app: riprova dal collegamento."; }); }
prova();
</script></body></html>
"@
$ascolto = $null
try {
    $ascolto = New-Object System.Net.HttpListener
    $ascolto.Prefixes.Add("http://localhost:8778/")
    $ascolto.Start()
} catch { $ascolto = $null }

Start-Process -FilePath $py -WorkingDirectory $dir -WindowStyle Hidden -ArgumentList @(
    "-m", "streamlit", "run", "app.py",
    "--server.port", $port, "--server.address", "127.0.0.1", "--server.headless", "true",
    "--browser.gatherUsageStats", "false")

if ($ascolto) {
    Start-Process "http://localhost:8778/"
    $fine = (Get-Date).AddSeconds(150)
    $servita = $false
    while ((Get-Date) -lt $fine) {
        $task = $ascolto.GetContextAsync()
        while (-not $task.AsyncWaitHandle.WaitOne(300)) {
            if ($servita -and (Test-App)) { $fine = Get-Date; break }          # pagina gia' servita e app pronta: finito
            if ((Get-Date) -ge $fine) { break }
        }
        if (-not $task.IsCompleted) { continue }
        $ctx = $task.Result
        $b = [Text.Encoding]::UTF8.GetBytes($pagina)
        $ctx.Response.ContentType = "text/html; charset=utf-8"
        $ctx.Response.ContentLength64 = $b.Length
        $ctx.Response.OutputStream.Write($b, 0, $b.Length)
        $ctx.Response.OutputStream.Close()
        $servita = $true
    }
    $ascolto.Stop()
} else {
    # senza pagina di attesa: aspetto che l'app sia pronta e poi apro il browser
    for ($i = 0; $i -lt 60; $i++) { Start-Sleep -Milliseconds 500; if (Test-App) { break } }
    Start-Process $url
}
