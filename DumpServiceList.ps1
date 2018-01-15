[CmdletBinding()]
param (
    [Parameter(Mandatory = $false)]
    [string] $OutputFileName
)

function Get-OutputFileName
{
    [CmdletBinding()]
    [OutputType([string])]
    param ()

    ('ServiceList_{0}.{1}.{2}_{3}_{4}.tsv' -f $PSVersionTable.BuildVersion.Major, $PSVersionTable.BuildVersion.Minor, $PSVersionTable.BuildVersion.Build, (Get-WinSystemLocale).Name, $env:COMPUTERNAME)
}

function Get-DescriptionLabel
{
    [CmdletBinding()]
    [OutputType([string])]
    param ()

    $systemLocale = (Get-WinSystemLocale).Name
    if ($systemLocale -eq 'ja-JP')
    {
        '説明'
    }
    else
    {
        'DESCRIPTION'
    }
}

function Invoke-ServiceControlCommand
{
    [OutputType([PSCustomObject])]
    param (
        [Parameter(Mandatory = $true)]
        [string] $Arguments
    )

    # Prepare the process information.
    $processInfo = New-Object -TypeName 'System.Diagnostics.ProcessStartInfo'
    $processInfo.UseShellExecute = $false
    $processInfo.RedirectStandardError = $true
    $processInfo.RedirectStandardOutput = $true
    $processInfo.FileName = 'C:\Windows\System32\sc.exe'
    $processInfo.Arguments = $Arguments

    # Create, execute and wait to the process.
    $process = New-Object -TypeName 'System.Diagnostics.Process'
    $process.StartInfo = $processInfo
    [void] $process.Start()
    $process.WaitForExit()

    # Retrieve the results of invoke command.
    [PSCustomObject] @{
        ExitCode = $process.ExitCode
        StdOut = $process.StandardOutput.ReadToEnd()
        StdErr = $process.StandardError.ReadToEnd()
    }
}

function Get-ServiceStartType
{
    [OutputType([string])]
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [string] $ServiceName
    )

    # Invoke the sc.exe command.
    $results = Invoke-ServiceControlCommand -Arguments ('qc "{0}"' -f $ServiceName)

    if ($results.ExitCode -eq 0)
    {
        # Extract the start type of the service from the stdout.
        $pattern = '^\s+START_TYPE\s+:\s+([0-9])\s+([A-Z_]+).+$'
        $match = [System.Text.RegularExpressions.regex]::Match($results.StdOut, $pattern, [System.Text.RegularExpressions.RegexOptions]::Multiline)
        $startTypeAsNum = $match.Captures[0].Groups[1].Value.Trim()
        $startTypeAsString = $match.Captures[0].Groups[2].Value.Trim()

        # Return the start type of the service.
        if (($startTypeAsNum -eq 2) -and ($startTypeAsString -eq 'AUTO_START'))
        {
            'Automatic'
        }
        elseif (($startTypeAsNum -eq 3) -and ($startTypeAsString -eq 'DEMAND_START'))
        {
            'Manual'
        }
        elseif (($startTypeAsNum -eq 4) -and ($startTypeAsString -eq 'DISABLED'))
        {
            'Disabled'
        }
        else
        {
            ('Unknown ({0}: {1})' -f $startTypeAsNum, $startTypeAsString)
        }
    }
    else
    {
        throw ('Failed sc.exe for {0} with {1}. {2}' -f $ServiceName, $results.ExitCode, $results.StdErr)
    }
}

function Get-ServiceDescription
{
    [OutputType([string])]
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [string] $ServiceName,

        [Parameter(Mandatory = $false)]
        [string] $DescriptionLabel = 'DESCRIPTION'
    )

    # Invoke the sc.exe command.
    $results = Invoke-ServiceControlCommand -Arguments ('qdescription "{0}" 4096' -f $ServiceName)

    if ($results.ExitCode -eq 0)
    {
        # Extract the description of the service from stdout.
        $pattern = ('^.+{0}:(.*)$' -f $DescriptionLabel)
        $match = [System.Text.RegularExpressions.regex]::Match($results.StdOut, $pattern, [System.Text.RegularExpressions.RegexOptions]::Singleline)
        $description = $match.Captures[0].Groups[1].Value.Trim()

        $description
    }
    else
    {
        if ($results.ExitCode -eq 15100)
        {
            ('Cannot read description. Error Code: {0}' -f $results.ExitCode)
        }
        else
        {
            throw ('Failed sc.exe for {0} with {1}. {2}' -f $ServiceName, $results.ExitCode, $results.StdErr)
        }
    }
}


# Generate the output file name if it does not specified as parameter.
if (-not $PSBoundParameters.ContainsKey('OutputFileName'))
{
    $OutputFileName = Get-OutputFileName
}

# Get the description label in the sc.exe command result.
$descriptionLabel = Get-DescriptionLabel

# Export the service informations as TSV.
Get-Service |
    Select-Object -Property 'ServiceName','DisplayName','StartType' |
    ForEach-Object -Process {

        $service = $_
        $description = Get-ServiceDescription -ServiceName $service.ServiceName -DescriptionLabel $descriptionLabel

        [pscustomobject] @{
            ServiceName = $service.ServiceName
            DisplayName = $service.DisplayName
            StartType = $service.StartType
            Description = $description
        }
    } |
    Export-Csv -LiteralPath $OutputFileName -NoTypeInformation -Delimiter "`t" -Encoding UTF8
