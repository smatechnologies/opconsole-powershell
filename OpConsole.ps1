param(
    $opconmodule = ((($MyInvocation.MyCommand).Path | Split-Path -Parent) + "\OpConModule.psm1")
    ,$opconsolemodule = ((($MyInvocation.MyCommand).Path | Split-Path -Parent) + "\OpConsoleModule.psm1")
    ,$consoleConfig = ((($MyInvocation.MyCommand).Path | Split-Path -Parent) + "\OpConsole.ini")
)

if((Test-Path $opconModule) -and (Test-Path $opconsoleModule))
{
    Import-Module -Name $opconmodule -Force
    Import-Module -Name $opconsolemodule -Force

    #Verify PS version is at least 7.0 
    if($PSVersionTable.PSVersion.Major -lt 7)
    {
        Write-Host "OpConsole only supports PowerShell 7.0+"
        Exit
    }
}
else
{
    Write-Host "Unable to import SMA API modules!"
    Exit
}

# Clear the console screen to start
Clear-Host

# Force TLS 1.2
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
OpCon_SkipCerts # Skips any self signed certificates

$logins = @()   # Array to manage OpCon env's
$logins += [pscustomobject]@{"id"=$logins.Count;"name"="Create New";"url"="";"user"="";"token"="";"expiration"="";"release"="";"active"=""}
$logs = @()     # Array of log files
$users = @()    # Array of api users
$cmdArray = @() # Array to store commands entered

Write-Host "==================================================================================================================="
Write-Host "                                   Welcome to OpConsole v0.5 for OpCon v19.1.1"
Write-Host "===================================================================================================================`n"

# Load any saved configurations
if(test-path $consoleConfig)
{
    Get-Content $consoleConfig | ForEach-Object { 
                                                    #if($_ -like "LOG*")
                                                    #{ $logs += [pscustomobject]@{"id"=$logs.Count;"Log"=(($_.Split("="))[1] | Split-Path -Leaf);"Location"=(($_.Split("="))[1] | Split-Path -Parent) } }
                                                    
                                                    if($_ -like "USER*")
                                                    { $users += [pscustomobject]@{ "id"=$users.Count;"user"=$_.Substring(5,$_.IndexOf("=")-5);"environment"=($_.Split("="))[1] } }

                                                    if($_ -like "CONNECT*")
                                                    { 
                                                        $version = ""
                                                        if($logins.Count -ge 1)
                                                        { $version = (OpCon_APIVersion -url ($_.Split("="))[1]).opConRestApiProductVersion }
                                                        
                                                        $logins += [pscustomobject]@{"id"=$logins.Count;"name"=$_.Substring($_.IndexOf("_")+1,$_.IndexOf("=")-($_.IndexOf("_")+1));"url"=($_.Split("="))[1];"user"="";"token"="";"expiration"="";"release"=$version;"active"="false" } 
                                                    }
                                                }
}
else
{ return Write-Host "No OpConsole.ini file found at: "$path }

# Add the users to the appropriate environment
$suppress = $logins | ForEach-Object{ 
                                        $tempName = $_.name
                                        $tempId = $_.id
                                        $users | Where-Object{ 
                                                                if($_.environment -eq $tempName)
                                                                { $logins[$tempId].user = $_.user }
                                                            }
                                    }

# Display logins
if($logins.Count -gt 1)
{ 
    Write-Host "Imported OpCon Environments:"
    $logins | Where-Object{ $_.name -ne "Create New" } | Format-Table Name,URL,Release | Out-Host 
}

Write-Host "For help use 'opc-help', to connect to OpCon use 'opc-connect'`n"

$rerun = ""
$command = ""
While($command -ne "exit" -and $command -ne "quit" -and $command -ne "opc-exit")
{
    $prompt = "<|"

    # Handles rerunning a command
    if($rerun -eq "")
    { $command = Read-Host $prompt }
    else 
    { 
        $command = $rerun
        $rerun = "" 
    }

    # Adds command to history
    try
    { $cmdArray += [pscustomobject]@{"id"=$cmdArray.Count;"command"=$command} }
    catch
    { Write-Host $_ }

    try
    {
        if($command -eq "opc-connect")
        { $logins = opc-connect -logins $logins }
        elseif($command -eq "opc-connect-view")
        { $logins | Format-Table Id,Name,URL,User,Expiration,Release,Active | Out-Host }
        #elseif($command -eq "opc-logs")
        #{ $logs = opc-logs -logs $logs } # Add option to search for lines that contain X
        elseif($command -eq "opc-history")
        { $rerun = opc-history -cmdArray $cmdArray }
        elseif($command -eq "opc-services")
        { opc-services }
        elseif(($command -notlike "*exit") -and ($command.StartsWith("opc") -and ($command -ne "opc-help")))
        { 
            $activeConnection = $logins | Where-Object{ $_.active -eq "true" }
            if($activeConnection)
            {
                if(OpConsole_CheckExpiration -connection $activeConnection)
                {
                    $session = $logins | Where-Object{ $_.active -eq "true" }
                    Invoke-Expression -Command ("$command -url " + $session.url + " -token '" + $session.token + "'") | Out-Host
                }
                else 
                {
                    $suppress = $logins | Where-Object{ $_.active -eq "true";$_.token = "" }
                    $suppress = opc-connect -logins $logins
                }
            }
            else
            {
                Write-Host "Must connect to an OpCon environment first!"
                $suppress = opc-connect -logins $logins
            }
        }
        elseif($command -notlike "*exit")
        { Invoke-Expression -Command $command | Out-Host }
    }
    catch
    { Write-Host $_ }

    Write-Host "==================================================================================`r`n"    
}

opc-exit
