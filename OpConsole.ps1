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
        MsgBox -Title "Error" -Message "OpConsole only supports Powershell 7+" 
        Exit
    }
}
else
{
    MsgBox -Title "Error" -Message "Unable to import SMA API modules!" 
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

    # Add the users to the appropriate environment
    if($users.Count -gt 0)
    {
        $logins | ForEach-Object{ 
                                    $loginURL = $_.URL
                                    $loginName = $_.name
                                    $loginRelease = $_.release

                                    $allUsers = $users | Where-Object{ $_.environment -eq $loginName}
                                    if($allUsers.Count -gt 1)
                                    {
                                        For($x=0;$x -lt $allUsers.Count;$x++)
                                        {
                                            if($x -eq 0)
                                            { ($logins | Where-Object{$_.name -eq $allUsers[$x].environment}).user = $allUsers[$x].user }
                                            else
                                            { $logins += [pscustomobject]@{"id"=$logins.Count;"name"=$loginName;"url"=$loginURL;"user"=$allUsers[$x].user;"token"="";"expiration"="";"release"=$loginRelease;"active"="false" } }
                                        }
                                    }
                                    elseif($allUsers.Count -eq 1)
                                    { ($logins | Where-Object{$_.name -eq $allUsers.environment}).user = $allUsers.user }
                                }
    }
}
else
{ return Write-Host "No OpConsole.ini file found at: "$path }

# Display logins
if($logins.Count -gt 1)
{ 
    Write-Host "Imported OpCon Environments:"
    $logins | Where-Object{ $_.name -ne "Create New" } | Format-Table Name,URL,User,Release | Out-Host 
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
        Switch -Wildcard ($command)
        {
            "opc-clear"         { Clear-Host; break }
            "opc-connect"       { $logins = opc-connect -logins $logins; break }
            "opc-connect-view"  { $logins | Format-Table Id,Name,URL,User,Expiration,Release,Active | Out-Host; break }
            "opc-exit"          { opc-exit; break }
            "opc-help"          { opc-help; break }
            "opc-history"       { $rerun = opc-history -cmdArray $cmdArray; break }
            "opc-logs"          { $logs = opc-logs -logs $logs; break }
            "opc-reload"        {
                                    Clear-Host
                                    Import-Module -Name $opconmodule -Force
                                    Import-Module -Name $opconsolemodule -Force
                                    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
                                    OpCon_SkipCerts # Skips any self signed certificates
                                    break
                                }
            "opc-services"      { opc-services; break }
            "opc-*"             {   
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
                                    break
                                }
            Default             { Invoke-Expression -Command $command | Out-Host; break }
        }
    }
    catch
    { Write-Host $_ }

    Write-Host "==================================================================================`r`n"    
}

opc-exit
