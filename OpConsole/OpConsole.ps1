param(
    $opconModule = ((($MyInvocation.MyCommand).Path | Split-Path -Parent) + "\Modules\OpConModule.psm1")
    ,$opconsoleModule = ((($MyInvocation.MyCommand).Path | Split-Path -Parent) + "\Modules\OpConsoleModule.psm1")
    ,$customModule = ((($MyInvocation.MyCommand).Path | Split-Path -Parent) + "\Modules\CustomModule.psm1")
    ,$consoleConfig = ((($MyInvocation.MyCommand).Path | Split-Path -Parent) + "\OpConsole.ini")
)

if((Test-Path $opconModule) -and (Test-Path $opconsoleModule))
{
    Import-Module -Name $opconmodule -Force
    Import-Module -Name $opconsolemodule -Force

    #Verify PS version is at least 7.0 
    if($PSVersionTable.PSVersion.Major -lt 7)
    {
        Write-Host "OpConsole only supports Powershell 7+" 
        MsgBox -Title "Error" -Message "OpConsole only supports Powershell 7+" 
        Exit
    }
}
else
{
    Write-Host "Unable to import SMA API modules!" 
    MsgBox -Title "Error" -Message "Unable to import SMA API modules!" 
    Exit
}

# Import Custom module
if($customModule -ne "")
{ 
    if(Test-Path $customModule)
    { Import-Module -Name $customModule -Force }
    else
    {
        Write-Host "Unable to import custom module!"
        MsgBox -Title "Error" -Message "Unable to import custom module!" 
        Exit
    }
}

# Clear the console screen to start
Clear-Host

# Force TLS 1.2
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
OpCon_SkipCerts # Skips any self signed certificates

$logins = @()     # Array to manage OpCon env's
$logs = @()       # Array of log files
$users = @()      # Array of api users
$sqlLogins = @()  # Array of sql servers and users
$sqlusers = @()   # Array of sql users 
$cmdArray = @()   # Array to store commands entered
$lastModified = "2020-08-10" # Last tested
$opconVersion = "19.1.1"     # Last tested OpCon Version

# Initialize arrays with "create new" option
$logins += [pscustomobject]@{"id"=$logins.Count;"name"="Create New";"url"="";"user"="";"token"="";"expiration"="";"release"="";"active"=""}
$logs += [pscustomobject]@{"id"=$logs.Count;"Location"="Create New" }
$sqlLogins += [PSCustomObject]@{ "id"=$sqlLogins.Count;"server"="";"sqlname"="Create New";"user"= "";db="";"password"="";"active"=""}

Write-Host "============================================================================="
Write-Host "      Welcome to OpConsole v0.75.$lastModified for OpCon v$opconVersion"
Write-Host "=============================================================================`n"

# Load any saved configurations
if(test-path $consoleConfig)
{
    Get-Content $consoleConfig | ForEach-Object { 
                                                    if($_ -like "LOG*")
                                                    { $logs += [pscustomobject]@{ "id"=$logs.Count;"Location"=($_.Split("="))[1] } }
                                                    
                                                    if($_ -like "OPCON-USER_*")
                                                    { $users += [pscustomobject]@{ "id"=$users.Count;"user"=($_.Split("="))[1];"environment"=$_.Substring($_.IndexOf("USER_")+5,$_.IndexOf("=")-($_.IndexOf("USER_")+5)) } }
                                                    
                                                    if($_ -like "OPCON-SERVER_*")
                                                    { 
                                                        $version = ""
                                                        if($logins.Count -ge 1)
                                                        { $version = (OpCon_APIVersion -url ($_.Split("="))[1]).opConRestApiProductVersion }
                                                        
                                                        $logins += [pscustomobject]@{"id"=$logins.Count;"name"=$_.Substring($_.IndexOf("_")+1,$_.IndexOf("=")-($_.IndexOf("_")+1));"url"=($_.Split("="))[1];"user"="";"token"="";"expiration"="";"release"=$version;"active"="false" } 
                                                    }

                                                    if($_ -like "SQL-SERVER_*")
                                                    { $sqlLogins += [pscustomobject]@{ "id"=$sqlLogins.Count;"server"=$_.Substring($_.IndexOf("=")+1,$_.IndexOf(",")-($_.IndexOf("=")+1));"sqlname"=$_.Substring($_.IndexOf("_")+1,$_.IndexOf("=")-($_.IndexOf("_")+1));"user"= "";db=$_.Substring($_.IndexOf(",")+1);"password"="";"active"="false" } }

                                                    if($_ -like "SQL-USER_*")
                                                    { $sqlusers += [pscustomobject]@{ "id"=$sqlusers.Count;"user"=$_.Substring($_.IndexOf("=")+1);"sqlname"= $_.Substring($_.IndexOf("_")+1,$_.IndexOf("=")-($_.IndexOf("_")+1)) } }
                                                }

    # Add the users to the appropriate environment
    if($users.Count -gt 0)
    {
        $logins | ForEach-Object{ 
                                    $loginURL = $_.URL
                                    $loginName = $_.name
                                    $loginRelease = $_.release

                                    $allUsers = $users | Where-Object{ $_.environment -eq $loginName}
                                    if($allUsers.Count -gt 0)
                                    {
                                        For($x=0;$x -lt $allUsers.Count;$x++)
                                        {
                                            if($x -eq 0)
                                            { ($logins | Where-Object{$_.name -eq $allUsers[$x].environment}).user = $allUsers[$x].user }
                                            else
                                            { $logins += [pscustomobject]@{"id"=$logins.Count;"name"=$loginName;"url"=$loginURL;"user"=$allUsers[$x].user;"token"="";"expiration"="";"release"=$loginRelease;"active"="false" } }
                                        }
                                    }
                                }
    }

    if($sqlusers.Count -gt 0 -and $sqlLogins.Count -gt 0)
    {
        $sqlLogins | ForEach-Object{
                                        $sqlserver = $_.sqlname
                                        $sqlserverConnection = $_.server
                                        $sqldb = $_.db
                                        $allsqlUsers = $sqlusers | Where-Object{ $_.sqlname -eq $sqlserver }
                                        if($allsqlUsers.Count -gt 0)
                                        {
                                            For($x=0;$x -lt $allsqlUsers.Count;$x++)
                                            {
                                                if($x -eq 0)
                                                { ($sqlLogins | Where-Object{ $_.sqlname -eq $allsqlUsers[$x].sqlname }).user = $allsqlUsers[$x].user }
                                                else
                                                { $sqlLogins += [pscustomobject]@{id=$sqlLogins.Count;server=$sqlserverConnection;db=$sqlDB;sqlname=$allsqlUsers[$x].sqlname;user=$allsqlUsers[$x].user;password="";active="false"} }
                                            }
                                        }
                                    }
    }
}
else
{ Write-Host "No OpConsole.ini file found at: "$consoleConfig }

# Display logins
if($logins.Count -gt 1)
{ 
    Write-Host "Imported Environments:"
    $logins.Where({ $_.name -ne "Create New" }) | Format-Table Name,URL,User,Release | Out-Host 
}

if($sqlLogins.Count -gt 0)
{ $sqlLogins.Where({$_.sqlname -ne "Create New"}) | Format-Table SQLName,Server,DB,User | Out-Host }

Write-Host "=============================================================================`r`n"
Write-Host "For help use 'opc-help' or 'opc-listall' `n"

$rerun = ""
$command = ""
While($command -ne "exit" -and $command -ne "quit" -and $command -ne "opc-exit")
{
    $prompt = "X"

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
            "opc-connect-view"  { $logins | Format-Table Id,Name,URL,User,Expiration,Release,Active | Out-Host; break }
            "opc-exit"          { Clear-Variable *; opc-exit; break }
            "opc-help"          { opc-help; break }
            "opc-history"       { $rerun = opc-history -cmdArray $cmdArray; break }
            "opc-logs"          { $logs = opc-logs -logs $logs; break }
            "opc-reload"        {
                                    Clear-Host
                                    Import-Module -Name $opconModule -Force
                                    Import-Module -Name $opconsoleModule -Force
                                    Import-Module -Name $customModule -Force
                                    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
                                    OpCon_SkipCerts # Skips any self signed certificates
                                    
                                    break
                                }
            "opc-services"      { opc-services; break }
            "opc-connect"       { 
                                    $menu = @()
                                    $menu += [pscustomobject]@{Id=$menu.Count;Option="Exit"}
                                    $menu += [pscustomobject]@{Id=$menu.Count;Option="OpCon"}
                                    
                                    if($sqlLogins.Count -gt 0)
                                    { $menu += [pscustomobject]@{Id=$menu.Count;Option="MS SQL"} }
                                
                                    $menu | Format-Table Id,Option | Out-Host
                                
                                    $selection = Read-Host "Enter a connection option <id>"
                                    if($menu[$selection].Option -eq "OpCon")
                                    { $logins = OpConsole_OpConConnect -logins $logins -configPath $consoleConfig }
                                    elseif($menu[$selection].Option -eq "MS SQL")
                                    { $sqlLogins = OpConsole_SQLConnect -sqlLogins $sqlLogins -configPath $consoleConfig }
    
                                    break 
                                }
            "*sql-*"            {
                                    try 
                                    {
                                        $sqlServer = $true
                                        Import-Module SqlServer -Force
                                    }
                                    catch 
                                    { 
                                        Write-Host "Unable to import 'SqlServer' module!"
                                        $sqlServer = $false
                                    }

                                    if($sqlServer)
                                    {
                                        $activeConnection = $sqlLogins.Where({ $_.active -eq "true" })
                                        if($activeConnection)
                                        {
                                            $session = $sqlLogins.Where({ $_.active -eq "true" })
                                            Invoke-Expression -Command ("$command -server " + $session.server + " -user '" + $session.user + "' -userPassword " + $session.password + " -db '" + $session.db + "'") | Out-Host
                                        }
                                        else
                                        {
                                            Write-Host "`r`n*****Must connect to a SQL environment first!*****"
                                            $sqlLogins = OpConsole_SQLConnect -sqlLogins $sqlLogins -configPath $consoleConfig
                                        }
                                    }
                                    break
                                }
            "opc-*"             {   
                                    $activeConnection = $logins.Where({ $_.active -eq "true" })
                                    if($activeConnection)
                                    {
                                        if(OpConsole_CheckExpiration -connection $activeConnection)
                                        {
                                            $session = $logins.Where({ $_.active -eq "true" })
                                            Invoke-Expression -Command ("$command -url " + $session.url + " -token '" + $session.token + "'") | Out-Host
                                        }
                                        else 
                                        {
                                            $logins.Where({ $_.active -eq "true";$_.token = "" }) | Out-Null
                                            $logins = OpConsole_OpConConnect -logins $logins -configPath $consoleConfig
                                        }
                                    }
                                    else
                                    {
                                        Write-Host "`r`n*****Must connect to an OpCon environment first!*****"
                                        $logins = OpConsole_OpConConnect -logins $logins -configPath $consoleConfig
                                    }
                                    break
                                }
            "custom-*"          { Invoke-Expression -Command $command | Out-Host; break }
            "$logins*"          { Write-Host "That command is not allowed."; break }
            "$sqlLogins*"       { Write-Host "That command is not allowed."; break }
            Default             { Invoke-Expression -Command $command | Out-Host; break }
        }
    }
    catch [Exception]
    { Write-Host $_ }

    Write-Host "`n"   
}

opc-exit
