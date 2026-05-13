# ─────────────────────────────────────────────────────────────────
#  Proxy local — Closum API
#  Corre este script antes de usar o dashboard.
#  Podes minimizar a janela; fecha-a quando acabares.
# ─────────────────────────────────────────────────────────────────

$port   = 3001
$prefix = "http://localhost:$port/"

$listener = [System.Net.HttpListener]::new()
$listener.Prefixes.Add($prefix)

try {
    $listener.Start()
} catch {
    Write-Host ""
    Write-Host "  ERRO: Nao foi possivel iniciar o proxy na porta $port." -ForegroundColor Red
    Write-Host "  Outra instancia pode ja estar a correr." -ForegroundColor Yellow
    Write-Host ""
    Read-Host "  Prima Enter para sair"
    exit 1
}

Write-Host ""
Write-Host "  ================================================" -ForegroundColor Cyan
Write-Host "   Proxy Closum activo em http://localhost:$port" -ForegroundColor Green
Write-Host "  ================================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "  Podes minimizar esta janela." -ForegroundColor Gray
Write-Host "  Nao a feches enquanto o dashboard estiver aberto." -ForegroundColor Gray
Write-Host ""

while ($listener.IsListening) {
    try {
        $ctx  = $listener.GetContext()
        $req  = $ctx.Request
        $resp = $ctx.Response

        # CORS headers — necessarios para o browser aceitar a resposta
        $resp.Headers.Add("Access-Control-Allow-Origin",  "*")
        $resp.Headers.Add("Access-Control-Allow-Methods", "GET, OPTIONS")
        $resp.Headers.Add("Access-Control-Allow-Headers", "Content-Type")

        if ($req.HttpMethod -eq "OPTIONS") {
            $resp.StatusCode = 200
            $resp.Close()
            continue
        }

        # Extrair o URL alvo do parametro ?url=
        $raw = $req.RawUrl
        if ($raw -match '[?&]url=([^&]+)') {
            $targetUrl = [System.Uri]::UnescapeDataString($matches[1])
        } else {
            $targetUrl = $null
        }

        $allowedPrefixes = @("https://api.closum.com/", "https://graph.facebook.com/")
        $isAllowed = $allowedPrefixes | Where-Object { $targetUrl -and $targetUrl.StartsWith($_) }
        if ($isAllowed) {
            try {
                $wc = [System.Net.WebClient]::new()
                $wc.Encoding = [System.Text.Encoding]::UTF8
                $data  = $wc.DownloadString($targetUrl)
                $bytes = [System.Text.Encoding]::UTF8.GetBytes($data)
                $resp.ContentType      = "application/json; charset=utf-8"
                $resp.ContentLength64  = $bytes.Length
                $resp.StatusCode       = 200
                $resp.OutputStream.Write($bytes, 0, $bytes.Length)
                Write-Host "  [OK] $targetUrl" -ForegroundColor Green
            } catch {
                $errJson = '{"status":false,"errors":["Erro ao contactar Closum: ' + $_.Exception.Message.Replace('"','\"') + '"]}'
                $bytes   = [System.Text.Encoding]::UTF8.GetBytes($errJson)
                $resp.ContentType     = "application/json; charset=utf-8"
                $resp.ContentLength64 = $bytes.Length
                $resp.StatusCode      = 502
                $resp.OutputStream.Write($bytes, 0, $bytes.Length)
                Write-Host "  [ERRO] $($_.Exception.Message)" -ForegroundColor Red
            }
        } else {
            $errJson = '{"status":false,"errors":["URL invalido"]}'
            $bytes   = [System.Text.Encoding]::UTF8.GetBytes($errJson)
            $resp.ContentType     = "application/json; charset=utf-8"
            $resp.ContentLength64 = $bytes.Length
            $resp.StatusCode      = 400
            $resp.OutputStream.Write($bytes, 0, $bytes.Length)
        }

        $resp.Close()

    } catch [System.Net.HttpListenerException] {
        # Listener foi fechado — sair normalmente
        break
    } catch {
        Write-Host "  [AVISO] $($_.Exception.Message)" -ForegroundColor Yellow
    }
}

$listener.Stop()
Write-Host ""
Write-Host "  Proxy terminado." -ForegroundColor Gray
