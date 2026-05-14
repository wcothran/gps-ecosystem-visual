# Tiny static-file server for local development of the GPS Ecosystem Visual.
# Serves the directory this script lives in over HTTP.
#
# Usage:
#   powershell -ExecutionPolicy Bypass -File .\serve.ps1
#   (then open http://localhost:8766/ in a browser)

param([int]$Port = 8766)

$root = $PSScriptRoot
$prefix = "http://localhost:$Port/"

$mime = @{
    '.html'='text/html; charset=utf-8'
    '.css'='text/css; charset=utf-8'
    '.js'='application/javascript; charset=utf-8'
    '.json'='application/json'
    '.xlsx'='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    '.png'='image/png'; '.jpg'='image/jpeg'; '.jpeg'='image/jpeg'; '.gif'='image/gif'; '.svg'='image/svg+xml'
}

$listener = New-Object System.Net.HttpListener
$listener.Prefixes.Add($prefix)
$listener.Start()
Write-Host "Serving $root at $prefix"
Write-Host "Press Ctrl+C to stop."

try {
    while ($listener.IsListening) {
        $ctx = $listener.GetContext()
        try {
            $rel = [System.Uri]::UnescapeDataString($ctx.Request.Url.AbsolutePath.TrimStart('/'))
            if ([string]::IsNullOrEmpty($rel)) { $rel = 'index.html' }
            $file = Join-Path $root $rel
            if (Test-Path $file -PathType Leaf) {
                $ext = [System.IO.Path]::GetExtension($file).ToLower()
                $ct = $mime[$ext]
                if (-not $ct) { $ct = 'application/octet-stream' }
                $bytes = [System.IO.File]::ReadAllBytes($file)
                $ctx.Response.ContentType = $ct
                $ctx.Response.ContentLength64 = $bytes.Length
                $ctx.Response.AddHeader('Cache-Control', 'no-store')
                $ctx.Response.OutputStream.Write($bytes, 0, $bytes.Length)
            } else {
                $ctx.Response.StatusCode = 404
                $msg = [System.Text.Encoding]::UTF8.GetBytes("Not found: $rel")
                $ctx.Response.OutputStream.Write($msg, 0, $msg.Length)
            }
        } catch {
            $ctx.Response.StatusCode = 500
        } finally {
            $ctx.Response.Close()
        }
    }
} finally {
    $listener.Stop()
}
