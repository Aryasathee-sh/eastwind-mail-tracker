# mcp_bridge.ps1
# Native PowerShell MCP stdio-to-SSE bridge for Claude Desktop
[Console]::InputEncoding = [System.Text.Encoding]::UTF8
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

$cloudUrl = "https://eastwind-mail-tracker.onrender.com/sse"
$postUrl = "https://eastwind-mail-tracker.onrender.com/message"

$client = New-Object System.Net.Http.HttpClient
$client.Timeout = [TimeSpan]::FromMinutes(10)

# Connect to SSE stream in background
$streamTask = [System.Threading.Tasks.Task]::Run([Action]{
    try {
        $stream = $client.GetStreamAsync($cloudUrl).Result
        $reader = New-Object System.IO.StreamReader($stream, [System.Text.Encoding]::UTF8)
        while (-not $reader.EndOfStream) {
            $line = $reader.ReadLine()
            if ($line -and $line.StartsWith("data: ")) {
                $data = $line.Substring(6).Trim()
                if ($data) {
                    [Console]::Out.WriteLine($data)
                    [Console]::Out.Flush()
                }
            }
        }
    } catch {}
})

# Read stdio from Claude Desktop and forward as HTTP POST
while ($true) {
    $line = [Console]::In.ReadLine()
    if ($null -eq $line) { break }
    if ($line.Trim()) {
        try {
            $content = New-Object System.Net.Http.StringContent($line, [System.Text.Encoding]::UTF8, "application/json")
            $resp = $client.PostAsync($postUrl, $content).Result
        } catch {}
    }
}
