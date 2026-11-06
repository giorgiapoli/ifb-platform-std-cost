# Ferma l'app Variazioni Standard Cost (chiude solo il processo Streamlit sulla porta 8777).
$port = 8777
Get-NetTCPConnection -LocalPort $port -State Listen -ErrorAction SilentlyContinue |
    ForEach-Object {
        $p = Get-CimInstance Win32_Process -Filter "ProcessId=$($_.OwningProcess)"
        if ($p.CommandLine -match "streamlit") { Stop-Process -Id $_.OwningProcess -Force -ErrorAction SilentlyContinue }
    }
