$url = "https://xxbfwvlqixnmonxytdxq.supabase.co/rest/v1/visits?period=is.null"
$headers = @{
    "apikey" = "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6Inh4YmZ3dmxxaXhubW9ueHl0ZHhxIiwicm9sZSI6InNlcnZpY2Vfcm9sZSIsImlhdCI6MTc4Mjc1NjE2NSwiZXhwIjoyMDk4MzMyMTY1fQ.PSk6RyFmg_OFTcCtYO74AeJj6wT4FGZS2K2JT9GEJ_A"
    "Authorization" = "Bearer eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6Inh4YmZ3dmxxaXhubW9ueHl0ZHhxIiwicm9sZSI6InNlcnZpY2Vfcm9sZSIsImlhdCI6MTc4Mjc1NjE2NSwiZXhwIjoyMDk4MzMyMTY1fQ.PSk6RyFmg_OFTcCtYO74AeJj6wT4FGZS2K2JT9GEJ_A"
    "Content-Type" = "application/json"
    "Prefer" = "return=minimal"
}
$body = '{"period": "July 2026"}'

try {
    Invoke-RestMethod -Uri $url -Method Patch -Headers $headers -Body $body
    Write-Host "Successfully patched NULL periods to 'July 2026' in visits table."
} catch {
    Write-Host "Error patching visits: $_"
}

$url_sum = "https://xxbfwvlqixnmonxytdxq.supabase.co/rest/v1/summaries?period=is.null"
try {
    Invoke-RestMethod -Uri $url_sum -Method Patch -Headers $headers -Body $body
    Write-Host "Successfully patched NULL periods in summaries table."
} catch {
    Write-Host "Error patching summaries: $_"
}

$url_spec = "https://xxbfwvlqixnmonxytdxq.supabase.co/rest/v1/specialty_classification?period=is.null"
try {
    Invoke-RestMethod -Uri $url_spec -Method Patch -Headers $headers -Body $body
    Write-Host "Successfully patched NULL periods in specialty_classification table."
} catch {
    Write-Host "Error patching specialty: $_"
}

$url_prod = "https://xxbfwvlqixnmonxytdxq.supabase.co/rest/v1/product_calls?period=is.null"
try {
    Invoke-RestMethod -Uri $url_prod -Method Patch -Headers $headers -Body $body
    Write-Host "Successfully patched NULL periods in product_calls table."
} catch {
    Write-Host "Error patching product_calls: $_"
}
