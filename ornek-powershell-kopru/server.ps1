# ============================================================
#  server.ps1  -  Modern tarayici <-> VBScript koprusu
#
#  - Windows ile gelen PowerShell'in HttpListener'i ile yerel HTTP sunucusu.
#  - "web" klasorundeki modern HTML/CSS/JS dosyalarini servis eder (ayni origin -> CORS yok).
#  - /api istekleri  ->  cscript //nologo backend.vbs <tmp> <action> [p1] [p2]
#  - VBScript sonucu (UTF-8 JSON) okunup tarayiciya dondurulur.
#
#  Node YOK.  WebView2 YOK.  Durdurmak icin bu pencerede Ctrl+C.
# ============================================================
$ErrorActionPreference = 'Stop'

$root    = Split-Path -Parent $MyInvocation.MyCommand.Path
$webRoot = Join-Path $root 'web'
$vbs     = Join-Path $root 'backend.vbs'
$port    = 8765
$prefix  = "http://localhost:$port/"
$cscript = 'cscript'    # 32-bit Jet/ACE gerekirse: "$env:WINDIR\SysWOW64\cscript.exe"

$mime = @{
    '.html' = 'text/html; charset=utf-8'
    '.css'  = 'text/css; charset=utf-8'
    '.js'   = 'application/javascript; charset=utf-8'
    '.json' = 'application/json; charset=utf-8'
    '.svg'  = 'image/svg+xml'
    '.png'  = 'image/png'
    '.ico'  = 'image/x-icon'
}

function Send-Bytes($ctx, [byte[]]$bytes, [string]$type, [int]$status = 200) {
    $ctx.Response.StatusCode      = $status
    $ctx.Response.ContentType     = $type
    $ctx.Response.ContentLength64 = $bytes.Length
    $ctx.Response.OutputStream.Write($bytes, 0, $bytes.Length)
    $ctx.Response.OutputStream.Close()
}
function Send-Text($ctx, [string]$text, [string]$type, [int]$status = 200) {
    Send-Bytes $ctx ([Text.Encoding]::UTF8.GetBytes($text)) $type $status
}

$listener = [System.Net.HttpListener]::new()
$listener.Prefixes.Add($prefix)          # "localhost" on eki yonetici gerektirmez
$listener.Start()
Write-Host ""
Write-Host "  Kopru calisiyor:  $prefix" -ForegroundColor Green
Write-Host "  Durdurmak icin :  Ctrl+C"  -ForegroundColor DarkGray
Write-Host ""

try {
    while ($listener.IsListening) {
        $ctx  = $listener.GetContext()
        $req  = $ctx.Request
        $path = $req.Url.AbsolutePath
        try {
            if ($path -eq '/api') {
                # ---------- KOPRU: HTTP istegi -> VBScript ----------
                $action = $req.QueryString['action']
                if (-not $action) { Send-Text $ctx '{"error":"action eksik"}' $mime['.json'] 400; continue }
                $p1 = [string]$req.QueryString['p1']
                $p2 = [string]$req.QueryString['p2']

                $tmp = [IO.Path]::GetTempFileName()
                # cscript'i kabuk araya girmeden dogrudan calistiririz (komut enjeksiyonu yok)
                (& $cscript '//nologo' $vbs $tmp $action $p1 $p2 2>&1 | Out-String).Trim() |
                    ForEach-Object { if ($_) { Write-Host "  [vbs] $_" -ForegroundColor DarkGray } }

                $out = if (Test-Path $tmp) { [IO.File]::ReadAllText($tmp, [Text.Encoding]::UTF8) } else { '' }
                Remove-Item $tmp -ErrorAction SilentlyContinue
                if (-not $out.Trim()) {
                    $out = '{"error":"VBScript bos sonuc dondurdu - yukaridaki [vbs] satirlarina bakin"}'
                }
                Write-Host ("  {0,-4} /api?action={1}" -f $req.HttpMethod, $action) -ForegroundColor Cyan
                Send-Text $ctx $out $mime['.json']
                continue
            }

            # ---------- Statik dosya servisi (modern HTML/CSS/JS) ----------
            if ($path -eq '/') { $path = '/index.html' }
            $full = [IO.Path]::GetFullPath((Join-Path $webRoot ($path.TrimStart('/'))))
            if (-not $full.StartsWith([IO.Path]::GetFullPath($webRoot))) {   # path traversal koruma
                Send-Text $ctx 'Forbidden' 'text/plain; charset=utf-8' 403; continue
            }
            if (Test-Path $full -PathType Leaf) {
                $ext  = [IO.Path]::GetExtension($full).ToLower()
                $type = if ($mime.ContainsKey($ext)) { $mime[$ext] } else { 'application/octet-stream' }
                Send-Bytes $ctx ([IO.File]::ReadAllBytes($full)) $type
            } else {
                Send-Text $ctx 'Bulunamadi' 'text/plain; charset=utf-8' 404
            }
        }
        catch {
            Write-Host "  HATA: $($_.Exception.Message)" -ForegroundColor Red
            try { Send-Text $ctx ('{"error":"' + ($_.Exception.Message -replace '"','\"') + '"}') $mime['.json'] 500 } catch {}
        }
    }
}
finally {
    $listener.Stop(); $listener.Close()
}
