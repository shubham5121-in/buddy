$port = 8080
$listener = New-Object System.Net.HttpListener
try {
    $listener.Prefixes.Add("http://localhost:$port/")
    $listener.Start()
} catch {
    $port = 8081
    $listener = New-Object System.Net.HttpListener
    $listener.Prefixes.Add("http://localhost:$port/")
    $listener.Start()
}

Write-Host "==========================================" -ForegroundColor Cyan
Write-Host "      SBE Web App Local Server" -ForegroundColor Cyan
Write-Host "==========================================" -ForegroundColor Cyan
Write-Host "Serving from: $PSScriptRoot"
Write-Host "URL: http://localhost:$port/" -ForegroundColor Green
Write-Host "STATUS: Running... Please do not close this window." -ForegroundColor Yellow
Write-Host ""

Start-Process "http://localhost:$port/index.html"

while ($listener.IsListening) {
    try {
        $context = $listener.GetContext()
        $request = $context.Request
        $response = $context.Response

        $localPath = $request.Url.LocalPath.TrimStart('/')
        if ([string]::IsNullOrEmpty($localPath)) {
            $localPath = "index.html"
        }

        $localPath = [uri]::UnescapeDataString($localPath)
        $filePath = Join-Path $PSScriptRoot $localPath
        
        # Ensure it's safe (prevent directory traversal)
        $fullPath = [System.IO.Path]::GetFullPath($filePath)
        $rootPath = [System.IO.Path]::GetFullPath($PSScriptRoot)
        
        if ($fullPath.StartsWith($rootPath) -and (Test-Path $fullPath -PathType Leaf)) {
            $content = [System.IO.File]::ReadAllBytes($fullPath)
            
            $ext = [System.IO.Path]::GetExtension($fullPath).ToLower()
            switch ($ext) {
                ".html" { $response.ContentType = "text/html" }
                ".js"   { $response.ContentType = "application/javascript" }
                ".css"  { $response.ContentType = "text/css" }
                ".png"  { $response.ContentType = "image/png" }
                ".jpg"  { $response.ContentType = "image/jpeg" }
                ".svg"  { $response.ContentType = "image/svg+xml" }
                ".json" { $response.ContentType = "application/json" }
                default { $response.ContentType = "application/octet-stream" }
            }
            
            $response.ContentLength64 = $content.Length
            $response.OutputStream.Write($content, 0, $content.Length)
        } else {
            $response.StatusCode = 404
        }
    } catch {
        # Ignore client disconnect errors
    } finally {
        if ($null -ne $context) {
            try { $context.Response.Close() } catch {}
        }
    }
}
