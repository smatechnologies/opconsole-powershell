param(
    $opconModule = ((($MyInvocation.MyCommand).Path | Split-Path -Parent) + "\Modules\OpConModule.psm1")
    ,$opconsoleModule = ((($MyInvocation.MyCommand).Path | Split-Path -Parent) + "\Modules\OpConsoleModule.psm1")
    ,$reportModule = ((($MyInvocation.MyCommand).Path | Split-Path -Parent) + "\Modules\ReportModule.psm1")
    ,$customModule = ((($MyInvocation.MyCommand).Path | Split-Path -Parent) + "\Modules\CustomModule.psm1")
    ,$consoleConfig = ((($MyInvocation.MyCommand).Path | Split-Path -Parent) + "\OpConsole.ini")
)
# Stop console program on error
$ErrorActionPreference = "Stop"

if((Test-Path $opconModule) -and (Test-Path $opconsoleModule) -and (Test-Path $reportModule))
{
    Import-Module -Name $opconModule -Force
    Import-Module -Name $opconsoleModule -Force
    Import-Module -Name $reportModule -Force

    #Verify PS version is at least 7.0 
    if($PSVersionTable.PSVersion.Major -lt 7)
    { Write-Host "OpConsole only supports Powershell 7+";Exit }
}
else
{ Write-Host "Unable to import modules!";Exit }

# Import Custom module
if($customModule -ne "")
{ 
    if(Test-Path $customModule)
    { Import-Module -Name $customModule -Force }
    else
    { Read-Host "Unable to import custom module!"}
}

# Force TLS 1.2
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
OpCon_SkipCerts # Skips any self signed certificates

$logins = New-Object System.Collections.ArrayList     # Array to manage OpCon env's
$users = New-Object System.Collections.ArrayList      # Array of api users
$cmdArray = New-Object System.Collections.ArrayList   # Array to store commands entered

# Initialize arrays with "create new" option
$logins.Add( [pscustomobject]@{"id"=$logins.Count;"name"="Create New";"url"="";"user"="";"token"="";"expiration"="";"release"="";"active"=""} ) | Out-Null

# Load any saved configurations
if(test-path $consoleConfig)
{
    Get-Content $consoleConfig | ForEach-Object {         
        if($_ -like "OPCON-USER_*")
        { $users.Add( [pscustomobject]@{ "id"=$users.Count;"user"=($_.Split("="))[1];"environment"=$_.Substring($_.IndexOf("USER_")+5,$_.IndexOf("=")-($_.IndexOf("USER_")+5)) } ) | Out-Null }
        
        if($_ -like "OPCON-SERVER_*")
        { 
            # Consider a custom OpCon Login function to also grab version
            $version = ""
            $logins.Add( [pscustomobject]@{"id"=$logins.Count;"name"=$_.Substring($_.IndexOf("_")+1,$_.IndexOf("=")-($_.IndexOf("_")+1));"url"=($_.Split("="))[1];"user"="";"token"="";"expiration"="";"release"=$version;"active"="false" } ) | Out-Null 
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
            if($allUsers.Count -gt 0)
            {
                For($x=0;$x -lt $allUsers.Count;$x++)
                {
                    if($x -eq 0)
                    { ($logins | Where-Object{$_.name -eq $allUsers[$x].environment}).user = $allUsers[$x].user }
                    else
                    { $logins.Add( [pscustomobject]@{"id"=$logins.Count;"name"=$loginName;"url"=$loginURL;"user"=$allUsers[$x].user;"token"="";"expiration"="";"release"=$loginRelease;"active"="false" } ) | Out-Null }
                }
            }
        }
    }
}
else
{ Read-Host "No OpConsole.ini file found at: "$consoleConfig }

function OpConsole_Prep($currentLogins,$connectionPath)
{
    $activeConnection = $currentLogins.Where({ $_.active -eq "true" })
    if($activeConnection)
    {
        if(OpConsole_CheckExpiration -connection $activeConnection)
        { return $currentLogins }
        else 
        {
            $currentLogins.Where({ $_.active -eq "true";$_.token = "" }) | Out-Null
            $currentLogins = OpConsole_OpConConnect -logins $currentLogins -configPath $connectionPath
            OpConsole_Prep -currentLogins $currentLogins -configPath $connectionPath
        }
    }
    else
    {
        $currentLogins = OpConsole_OpConConnect -logins $currentLogins -configPath $connectionPath
        OpConsole_Prep -currentLogins $currentLogins -configPath $connectionPath
    }
}

function OpConsole_CheckExpiration($connection)
{
    if($connection.token -ne "")
    {
        $date1 = Get-Date -date $connection.expiration | Get-Date -Format "MM/dd/yyyy HH:mm"
        $date2 = Get-Date -Format "MM/dd/yyyy HH:mm"
        if($date2 -gt $date1)
        { 
            Read-Host "`r`n*****API token expired for "$connection.Name" with user: "$connection.user"*****"
            return $false
        }
        else 
        { return $true }
    }
    else
    { return $false }
}

<#
.SYNOPSIS

Connects to an OpCon environment

.OUTPUTS

Object with the id,url,token,expiration,release.
#>
Function OpConsole_OpConConnect($logins,$configPath)
{ 
    Clear-Host
    Write-Host "-----------------------------------------------"
    Write-Host "        Connect to OpCon environment"
    Write-Host "-----------------------------------------------"

    $logins | Format-Table Id,Name,URL,User,Expiration,Release,Active | Out-Host
    Write-Host "-----------------------------------------------"

    $opconEnv = Read-Host "Enter Option <id>"
    
    Switch ($opconEnv)
    {
        0   { 
            Write-Host "-----------------------------------------------"
            Write-Host "Creating new OpCon connection....."

            # TLS
            $connection = New-Object System.Collections.ArrayList
            $connection.Add( [pscustomobject]@{Id=$connection.Count;Option="TLS"} ) | Out-Null
            $connection.Add( [pscustomobject]@{Id=$connection.Count;Option="Non-TLS"} ) | Out-Null
            $tls = OpConsole_Menu -items $connection

            Switch ($tls)
            {
                0   { $tlsSetting = "https://" }
                1   { $tlsSetting = "http://" }
                default { $tlsSetting = "https://" }
            }
        
            # OpCon hostname/ip
            $hostname = Read-Host "Enter OpCon hostname or ip"
            if($hostname -eq "")
            { $hostname = "localhost" }

            # Port
            $portMenu = New-Object System.Collections.ArrayList
            $portMenu.Add( [PSCustomObject]@{Id=$portMenu.Count;Option="443"})
            $portMenu.Add( [PSCustomObject]@{Id=$portMenu.Count;Option="Custom"})
            $port = OpConsole_Menu -items $portMenu

            Switch ($port)
            {
                0   {$portSetting = "443"}
                1   {$portSetting = Read-Host "Enter custom port"}
                default {$portSetting = "443"}
            }

            # OpCon username or Windows Auth
            $user = Read-Host "Enter Username (blank for Windows Auth)"
            if($user -eq "")
            { 
                $user = "Windows Auth" 
                $auth = OpCon_Login -url $url -user $user
            }
            else
            {
                $password = Read-Host "Enter Password" -AsSecurestring
                $auth = OpCon_Login -url $url -user $user -password ((New-Object PSCredential "user",$password).GetNetworkCredential().Password)
            }
            
            if($auth.id)
            {
                $password = "" # Clear out password variable
                $logins.Add([pscustomobject]@{"id"=$logins.Count;"name"=$name;"url"=$url;"user"=$user;"token"=("Token " + $auth.id);"expiration"=($auth.validUntil);"release"=((OpCon_APIVersion -url $url).opConRestApiProductVersion);"active"="true"}) | Out-Null
                Clear-Host     # Clears console
                Read-Host "Connected to"($logins | Where-Object{ $_.active -eq "true"}).name", expires at"($logins | Where-Object{ $_.active -eq "true"}).expiration
            }
            else
            { OpConsole_OpConConnect -logins $logins -configPath $configPath | Out-Null }

            $menu = New-Object System.Collections.ArrayList
            $menu.Add([pscustomobject]@{"Id"=$menu.Count;"Option"="Save connection"}) | Out-Null
            $menu.Add([pscustomobject]@{"Id"=$menu.Count;"Option"="Don't save"}) | Out-Null
            $save = OpConsole_Menu -items $menu

            Switch ($save)
            {
                0   {             
                        # OpCon environment name
                        $url = $tlsSetting + $hostname + ":" + $portSetting
                        $name = Read-Host "Enter environment name (optional)"
                        
                        if($name -eq "")
                        { $name = $url.Substring($url.IndexOf("//")+2) }

                        "`r`n# Additional OpCon Connection" | Out-File -Append -FilePath $configPath
                        ("OPCON-SERVER_" + $logins[$logins.Count-1].name + "=" + $logins[$logins.Count-1].url) | Out-File -Append -FilePath $configPath
                        ("OPCON-USER_" + $logins[$logins.Count-1].name + "=" + $logins[$logins.Count-1].user) | Out-File -Append -FilePath $configPath

                        Read-Host "$name OpCon connection saved!"
                        return $logins
                    }
                1   { return $logins}
                default { return $logins }
            }
        }
        {1..999 -contains $_} {
            $logins | ForEach-Object -Parallel { if($_.active -eq "true" -or $_.active -eq ""){ $_.active = "false" }}
            if($logins[$opconEnv].user -eq "")
            { $logins[$opconEnv].user = Read-Host "Enter Username" }

            if($logins[$opconEnv].token -ne "")
            {
                if(!(OpConsole_CheckExpiration -connection $logins[$opconEnv]))
                { $logins[$opconEnv].token = "" }
            }

            if($logins[$opconEnv].token -eq "")
            {
                if($logins[$opconEnv].user -ne "Windows Auth")
                {
                    $password = Read-Host "Password" -AsSecurestring #("Password for User: "+$logins[$opconEnv].User+" for Environment: "+$logins[$opconEnv].Name)
                    $auth = OpCon_Login -url $logins[$opconEnv].url -user $logins[$opconEnv].user -password ((New-Object PSCredential "user",$password).GetNetworkCredential().Password)

                    if($auth.id)
                    {
                        $password = "" # Clear out password variable
                        $logins[$opconEnv].token = ("Token " + $auth.id)
                        $logins[$opconEnv].expiration = ($auth.validUntil)
                        $logins[$opconEnv].release = ((OpCon_APIVersion -url $logins[$opconEnv].url).opConRestApiProductVersion)
                        $logins[$opconEnv].active = "true"
                        Clear-Host
                        Read-Host "Active connection is"($logins | Where-Object{ $_.active -eq "true"}).name", expires at"($logins | Where-Object{ $_.active -eq "true"}).expiration
                    }
                    else
                    { 
                        Read-Host "Problem authenticating to OpCon..."
                        OpConsole_OpConConnect -logins $logins -configPath $configPath | Out-Null 
                    }
                }
                else
                {
                    $auth = OpCon_Login -url $logins[$opconEnv].url -user $logins[$opconEnv].user

                    if($auth.id)
                    {
                        $logins[$opconEnv].token = ("Token " + $auth.id)
                        $logins[$opconEnv].expiration = ($auth.validUntil)
                        $logins[$opconEnv].release = ((OpCon_APIVersion -url $logins[$opconEnv].url).opConRestApiProductVersion)
                        $logins[$opconEnv].active = "true"
                        Clear-Host
                        Read-Host "Active connection is"($logins | Where-Object{ $_.active -eq "true"}).name", expires at"($logins | Where-Object{ $_.active -eq "true"}).expiration
                    }
                    else
                    { OpConsole_OpConConnect -logins $logins -configPath $configPath | Out-Null }
                }
            }
            else 
            { 
                $logins[$opconEnv].active = "true" 
                Clear-Host
                Read-Host "Active connection is"($logins | Where-Object{ $_.active -eq "true"}).name"-"($logins | Where-Object{ $_.active -eq "true"}).user ", expires at"($logins | Where-Object{ $_.active -eq "true"}).expiration
            }
            return $logins
        }
        default { OpConsole_Exit }
    }
}
New-Alias "opc-opconconnect" OpConsole_OpConConnect
New-Alias "opc-connect" OpConsole_OpConConnect

function OpConsole_OpConConnections($logins,$configPath)
{
    Clear-Host
    Write-Host "-----------------------------------------------"
    Write-Host "           View OpCon connections"
    Write-Host "-----------------------------------------------"
    $logins.Where( { $_.name -ne "Create New"} ) | Format-Table Id,Name,URL,User,Expiration,Release,Active | Out-Host
    Write-Host "-----------------------------------------------"

    $menu = New-Object System.Collections.ArrayList
    $menu.Add( [PSCustomObject]@{Id=$menu.Count;Option="Manage OpCon connections"} ) | Out-Null
    $menu.Add( [PSCustomObject]@{Id=$menu.Count;Option="Main Menu"} ) | Out-Null
    $answer = OpConsole_Menu -items $menu

    Switch ($answer)
    {
        0   { OpConsole_OpConConnect -logins $logins -configPath $configPath; break }
        default { return }
    }
}

# Quit the console if "exit" or "quit" is entered
function OpConsole_Exit
{
    Clear-History
    Clear-Host
    $PSDefaultParameterValues.Remove("Invoke-RestMethod:SkipCertificateCheck")
    Exit
}
New-Alias "opc-exit" OpConsole_Exit


# OpConsole menu loop
While($true)
{
    Clear-Host
    Write-Host "==============================================="
    Write-Host "                 OPCONSOLE"
    Write-Host "==============================================="

    $menu = New-object System.Collections.ArrayList
    $menu.Add([PSCustomObject]@{ "Id" = $menu.Count;"Option" = "Batch Users" }) | Out-Null
    $menu.Add([PSCustomObject]@{ "Id" = $menu.Count;"Option" = "Expressions" }) | Out-Null
    $menu.Add([PSCustomObject]@{ "Id" = $menu.Count;"Option" = "Global Properties" }) | Out-Null
    $menu.Add([PSCustomObject]@{ "Id" = $menu.Count;"Option" = "OpCon connections" }) | Out-Null
    $menu.Add([PSCustomObject]@{ "Id" = $menu.Count;"Option" = "OpConsole changelog" }) | Out-Null
    $menu.Add([PSCustomObject]@{ "Id" = $menu.Count;"Option" = "Quit" }) | Out-Null
    $option = OpConsole_Menu -items $menu
    
    switch ($option)
    {
        0 { 
            $logins = OpConsole_Prep -currentLogins $logins
            OpConsole_BatchUsers -url (($logins.Where({ $_.active -eq "true" })).url) -token (($logins.Where({ $_.active -eq "true" })).token)
            break 
        }
        1 { 
            $logins = OpConsole_Prep -currentLogins $logins 
            OpConsole_Expression -url ($logins.Where({ $_.active -eq "true" }).url) -token ($logins.Where({ $_.active -eq "true" }).token); 
            break 
        }
        2 { 
            $logins = OpConsole_Prep -currentLogins $logins
             OpConsole_Properties -url ($logins.Where({ $_.active -eq "true" }).url) -token ($logins.Where({ $_.active -eq "true" }).token)
             break 
        }
        3 { OpConsole_OpConConnections -logins $logins -configPath $configPath;break }
        4 { OpConsole_OpConConnect -logins $logins -configPath $configPath;break }
        5 { OpConsole_Version; break }
        6 { return OpConsole_Exit }
        default { return; break }
    }
}

if(($MyInvocation.MyCommand).Path -like "*\*") # Windows
{ $cmdArray | Format-Table Id,Command,Time | Out-File -FilePath (($MyInvocation.MyCommand).Path + "\OpConsole.log") }
elseif(($MyInvocation.MyCommand).Path -like "*/*") # Unix
{ $cmdArray | Format-Table Id,Command,Time | Out-File -FilePath (($MyInvocation.MyCommand).Path + "/OpConsole.log") }

OpConsole_Exit -logins $logins 