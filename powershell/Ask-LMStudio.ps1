<#
.SYNOPSIS
Sends a query to a local LM Studio API and returns the response.

.DESCRIPTION
This script interacts with a local LM Studio server's API, sending a user query and optional context. It supports developer mode for raw output, logging with timestamps, and verbose/debug output. Context can be piped in or passed as a parameter, or provided from a text file. You can also control the maximum number of tokens sent to the LLM. The script checks if LM Studio is running before sending the query. Usage statistics can be optionally displayed. The -Context and -File parameters are mutually exclusive.

.PARAMETER Query
The user question or prompt to send to the LLM. Required. Can be provided positionally.

.PARAMETER Context
Optional. Additional context for the LLM, can be piped in or passed as a parameter. Objects are formatted for readability. Accepts string arrays or objects. Cannot be used with -File.

.PARAMETER File
Optional. Path to a text-like file to use as context for the LLM. The file must exist and have a recognized text extension (e.g., .txt, .md, .json, etc.). Cannot be used with -Context.

.PARAMETER Developer
If specified, outputs the raw API response as JSON instead of just the LLM's answer.

.PARAMETER LogFile
Optional. If specified, all output is also written to this file with timestamps.

.PARAMETER MaxTokens
Optional. The maximum number of tokens to send to the LLM (default: 100000). Used to limit the context length.

.PARAMETER Usage
If specified, displays usage statistics (model, prompt, completion, and total tokens) from the LM Studio API response.

.EXAMPLE
PS> .\Ask-LMStudio.ps1 "What is the weather in Omaha tomorrow?"

.EXAMPLE
PS> Get-Service | .\Ask-LMStudio.ps1 -Query "Which services are stopped?"

.EXAMPLE
PS> .\Ask-LMStudio.ps1 -Query "Summarize this text" -Context (Get-Content .\file.txt) -Verbose -LogFile .\log.txt -Usage

.EXAMPLE
PS> .\Ask-LMStudio.ps1 -Query "Summarize this file" -File .\file.txt -Verbose -LogFile .\log.txt -Usage

.NOTES
Requires LM Studio server running at http://localhost:1234
#>
param(
    [Parameter(Mandatory=$true, Position=0)]
    [string]$Query,
    [Parameter(ValueFromPipeline=$true, ParameterSetName='Context')]
    [string[]]$Context,
    [Parameter(ParameterSetName='File')]
    [string]$File,
    [switch]$Developer,
    [string]$LogFile,
    [int]$MaxTokens = 100000,
    [switch]$Usage
)

# Write-Log function for timestamped output and optional log file
function Write-Log {
    param(
        [string]$Message,
        [switch]$VerboseOnly
    )
    $timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $logLine = "[$timestamp] $Message"
    $isVerbose = $false
    if ($PSCmdlet -and $PSCmdlet.MyInvocation.BoundParameters.ContainsKey('Verbose')) {
        $isVerbose = $PSCmdlet.MyInvocation.BoundParameters['Verbose']
    } elseif ($global:VerbosePreference -eq 'Continue') {
        $isVerbose = $true
    }
    if ($isVerbose -or -not $VerboseOnly) {
        Write-Host $logLine
    }
    if ($LogFile) {
        Add-Content -Path $LogFile -Value $logLine
    }
}

# Define the API endpoint
$apiUrl = "http://localhost:1234/v1/chat/completions"


# Prepare the request body as JSON
$messages = @()
$contextString = $null
if ($PSCmdlet.ParameterSetName -eq 'File' -and $File) {
    Write-Log "File parameter detected: $File" -VerboseOnly
    if (-not (Test-Path $File)) {
        Write-Log "File '$File' does not exist. Aborting."
        return
    }
    $fileItem = Get-Item $File
    if ($fileItem.PSIsContainer) {
        Write-Log "'$File' is a directory, not a file. Aborting."
        return
    }
    $fileType = $fileItem.Extension.ToLower()
    $textExtensions = @(".txt", ".md", ".csv", ".json", ".xml", ".log", ".ps1", ".py", ".js", ".ts", ".html", ".css", ".ini", ".conf", ".yaml", ".yml")
    if ($textExtensions -notcontains $fileType) {
        Write-Log "File '$File' does not appear to be a text file (extension: $fileType). Aborting."
        return
    }
    $contextString = Get-Content -Path $File -Raw
    # Escape problematic characters
    $contextString = $contextString -replace '`', '``'           # Escape backticks
    $contextString = $contextString -replace '\u0000', ''        # Remove null chars
    $contextString = $contextString -replace '\u001b', ''        # Remove ESC
    $contextString = $contextString -replace '\r', ''            # Remove carriage returns (optional)
    $contextString = $contextString -replace '\0', ''            # Remove literal nulls
    $contextString = $contextString -replace '\$', '`$'          # Escape PowerShell variable sigil
    $contextString = $contextString -replace '"', '`"'           # Escape double quotes
    # Check context length (approximate MaxTokens tokens as MaxTokens*4 characters)
    $maxTokenChars = $MaxTokens * 4
    if ($contextString.Length -gt $maxTokenChars) {
        Write-Log "File content exceeds $MaxTokens tokens (approx. $maxTokenChars chars). Truncating context sent to LLM."
        $contextString = $contextString.Substring(0, $maxTokenChars)
    }
    Write-Log "Final context string from file: $contextString" -VerboseOnly
    $messages += @{ role = "system"; content = $contextString }
} elseif ($PSCmdlet.ParameterSetName -eq 'Context' -and $Context) {
    Write-Log "Context parameter detected." -VerboseOnly
    if ($Context -is [string[]]) {
        Write-Log "Context is string array. Joining with newlines." -VerboseOnly
        $contextString = $Context -join "`n"
    } elseif ($Context -is [array]) {
        Write-Log "Context is an array. Using Format-List on each item." -VerboseOnly
        $formattedList = @()
        foreach ($item in $Context) {
            $formattedList += ($item | Format-List | Out-String).Trim()
        }
        $contextString = $formattedList -join "`n---`n"
    } elseif ($Context) {
        Write-Log "Context is not string array or array. Attempting to format as table." -VerboseOnly
        $contextString = ($Context | Format-Table -AutoSize | Out-String).Trim()
    }
    if (-not $contextString) {
        Write-Log "Context string is empty after previous attempts. Using Out-String fallback." -VerboseOnly
        $contextString = $Context | Out-String
    }
    # Check context length (approximate MaxTokens tokens as MaxTokens*4 characters)
    $maxTokenChars = $MaxTokens * 4
    if ($contextString.Length -gt $maxTokenChars) {
        Write-Log "Context exceeds $MaxTokens tokens (approx. $maxTokenChars chars). Truncating context sent to LLM."
        $contextString = $contextString.Substring(0, $maxTokenChars)
    }
    Write-Log "Final context string: $contextString" -VerboseOnly
    $messages += @{ role = "system"; content = $contextString }
} else {
    Write-Log "No context or file parameter provided." -VerboseOnly
}
$messages += @{ role = "user"; content = $Query }

$body = @{
    model = "local-model"
    messages = $messages
} | ConvertTo-Json -Depth 4

# Check if LM Studio is running before sending the query (just check if port is open)
Write-Log "Checking if LM Studio API port is open at $apiUrl ..." -VerboseOnly
$uri = [System.Uri]$apiUrl
$tcpClient = $null
try {
    $tcpClient = New-Object System.Net.Sockets.TcpClient
    $tcpClient.Connect($uri.Host, $uri.Port)
    $tcpClient.Close()
} catch {
    Write-Log "LM Studio API port $($uri.Port) is not open on $($uri.Host). Please start LM Studio and try again."
    return
}

# Send the POST request
Write-Log "Sending POST request to $apiUrl with body: $body" -VerboseOnly
try {
    $response = Invoke-RestMethod -Uri $apiUrl -Method Post -Body $body -ContentType 'application/json'
    Write-Log "Received response from API." -VerboseOnly
    if ($developer) {
        Write-Log "Developer mode: outputting raw response." -VerboseOnly
        $response | ConvertTo-Json -Depth 4 | ForEach-Object { Write-Log $_ }
    } else {
        Write-Log "User mode: outputting choices[0].message.content." -VerboseOnly
        if ($response.choices -and $response.choices.Count -gt 0 -and $response.choices[0].message -and $response.choices[0].message.content) {
            Write-Log $response.choices[0].message.content
        } else {
            Write-Log "No message content found in response."
        }
        # Output usage stats if available and -Usage is set
        if ($Usage -and $response.usage) {
            $usageSummary = "Usage:"
            if ($response.model) {
                $usageSummary += " model=$($response.model)"
            }
            if ($response.usage.prompt_tokens) {
                $usageSummary += ", prompt=$($response.usage.prompt_tokens)"
            }
            if ($response.usage.completion_tokens) {
                $usageSummary += ", completion=$($response.usage.completion_tokens)"
            }
            if ($response.usage.total_tokens) {
                $usageSummary += ", total=$($response.usage.total_tokens)"
            }
            Write-Log $usageSummary
        }
    }
} catch {
    Write-Log "Failed to contact LM Studio API: $_"
}
