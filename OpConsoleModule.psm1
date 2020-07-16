function OpConsole_Help
{
    $menu = @()
    $menu += [pscustomobject]@{"Command" = "opc-connect";"Description" = "Connects/Selects an OpCon environment"}
    $menu += [pscustomobject]@{"Command" = "opc-batchuser";"Description" = "Lets you manage OpCon batch users"}
    $menu += [pscustomobject]@{"Command" = "opc-eval";"Description" = "Lets you evaluate OpCon property expressions"}
    $menu += [pscustomobject]@{"Command" = "opc-reports";"Description" = "View various reports in OpConsole"}
    #$menu += [pscustomobject]@{"command" = "opc-scripts";"description" = "Lets you view or run scripts"}
    #$menu += [pscustomobject]@{"command" = "opc-ss";"description" = "Lets you manage Self Service"}
    $menu += [pscustomobject]@{"Command" = "opc-property" ;"Description" = "Lets you manage OpCon global properties"}
    $menu += [pscustomobject]@{"Command" = "opc-listall";"Description" = "Lists all commands"}
    $menu += [pscustomobject]@{"Command" = "opc-history";"Description" = "Shows previous commands and allows for rerurns"}
    $menu | Format-Table Command,Description | Out-Host
}
New-Alias "opc-help" OpConsole_Help

function OpConsole_Services
{         
    $server = Read-Host -Prompt "Server (blank if local)"
    $name = Read-Host -Prompt "Service/s (* wildcards)"

    if($server -and $name)
    { $tempServices = Invoke-Command -ComputerName "$server" -Script { Get-Service | Where-Object{ $_.Name -like "$name" } } }
    elseif($name)
    { $tempServices = Get-Service | Where-Object{ $_.Name -like "$name" } }

    if($tempServices.Count -gt 0)
    { 
        $services = @()
        $tempServices | ForEach-Object{ $services += [pscustomobject]@{"id"=$services.Count;"Status" = $_.status;"Name"=$_.Name;"DisplayName"=$_.DisplayName} } 

        $menu = @()
        $menu += [pscustomobject]@{"id"=$menu.Count;"Option"="Exit"}
        $menu += [pscustomobject]@{"id"=$menu.Count;"Option"="Start"}
        $menu += [pscustomobject]@{"id"=$menu.Count;"Option"="Stop"}

        $services | Format-Table Id,DisplayName,Status
        $svcOption = Read-Host -Prompt "Enter a service <id> (comma separated)"

        $menu | Format-Table Id,Option | Out-Host
        $option = Read-Host -Prompt "Enter an option <id>"

        if($option -gt 0)
        {
            $selectedOptions = $svcOption.Split(",")
            For($x=0;$x -lt $selectedOptions.Count;$x++)
            {
                if($server -ne "")
                { Invoke-Command -ComputerName "$server" -Script { (Get-Service -ComputerName $server).Name -like "*$services[($selectedOptions[$x])].Name*" | foreach-object{ sc.exe \\$server $status $_ } | Out-Host } }
                else
                { (Get-Service).Name -like "*$services[($selectedOptions[$x])].name*" | foreach-object{ sc.exe $status $_ } | Out-Host }
            }
                
            $subMenu = @()
            $subMenu += [pscustomobject]@{"id"=$subMenu.Count;"Option"="Exit"}
            $subMenu += [pscustomobject]@{"id"=$subMenu.Count;"Option"="Check service status"}
            $subMenu += [pscustomobject]@{"id"=$subMenu.Count;"Option"="Start over"}
            $subMenu | Format-Table Id,Option | Out-Host
            
            $checkStatus = Read-Host -Prompt "Enter an option <id>"
            if($checkStatus -eq 1)
            { 
                While($checkStatus -eq 1)
                {
                    For($x=0;$x -lt $selectedOptions.Count;$x++)
                    {
                        if($server -ne "")
                        { Invoke-Command -ComputerName "$server" -Script { Get-Service | Where-Object{ $_.Name -like ("*" + $services[($selectedOptions[$x])].Name + "*") } } | Out-Host }
                        else
                        { Get-Service | Where-Object{ $_.Name -like ("*" + $services[($selectedOptions[$x])].Name +"*") } | Out-Host }
                    }
                    $subMenu | Format-Table Id,Option | Out-Host
                    $checkStatus = Read-Host -Prompt "Enter an option <id>"
                }
            }
            
            if($checkStatus -eq 2)
            { opc-services }           
        }
    }
    else 
    {
        Write-Host "No services found with name "$name
        opc-services   
    }
}
New-Alias "opc-services" OpConsole_Services

function OpConsole_Properties($url,$token)
{
    $menu = @()
    $menu += [pscustomobject]@{"id"=$menu.Count;"Option" = "Exit"}
    $menu += [pscustomobject]@{"id"=$menu.Count;"Option" = "Create"}
    $menu += [pscustomobject]@{"id"=$menu.Count;"Option" = "View"}
    $menu += [pscustomobject]@{"id"=$menu.Count;"Option" = "Update"}
    
    $option = 999
    while($option -ne 0)
    { 
        $properties = OpCon_GetGlobalProperty -url $url -token $token

        $menu | Format-Table Id,Option | Out-Host
        $option = Read-Host -Prompt "Enter an option <id>)"

        If($option -eq 2)
        {
            $property = Read-Host -Prompt "Global Property name (blank if all)"
            
            if($property -eq "")
            { $properties | Out-Host }
            else
            { 
                $result = $properties | Where-Object{ $_.name -like "$property" } 
                $result | Out-Host

                If($result.Count -eq 0)
                { Write-Host "No properties with that name." }
            }
        }
        ElseIf($option -eq 1)
        {
            $propertyName = Read-Host -Prompt "Enter the new property name"

            $result = $properties | Where-Object{ $_.name -eq "$propertyName" }
            
            If($result)
            { Write-Host "Property named $propertyName already exists!" }
            else 
            {  
                $propertyValue = Read-Host -Prompt "Enter the property value"
                $encrypted = Read-Host -Prompt "Encrypted? (y/n, blank for no)"

                if($encrypted -match ("yes","y"))
                { $encrypted = $true }
                else
                { $encrypted = $false }

                OpCon_CreateGlobalProperty -url $url -token $token -name $propertyName -value $propertyValue -encrypt $encrypted 
            }
        }
        Elseif($option -eq 3)
        {
            $propertyName = Read-Host -Prompt "Enter the property name to update"
            $result = $properties | Where-Object{ $_.name -eq "$propertyName" }
            
            If($result)
            { 
                Write-Host "Current property value"
                Write-Host "----------------------"
                $result | Out-Host

                $propertyValue = Read-Host -Prompt "Enter the new property value"
                OpCon_SetGlobalProperty -url $url -token $token -id $result.id -value $propertyValue
            }
            else 
            { Write-Host "Property named $propertyName not found!" } 
        }
    }
}
New-Alias "opc-property" OpConsole_Properties
New-Alias "opc-properties" OpConsole_Properties

Function OpConsole_SelfService($url,$token)
{
    $menu = @()
    $menu += [pscustomobject]@{"id"=$menu.Count;"Option" = "Exit"}
    $menu += [pscustomobject]@{"id"=$menu.Count;"Option" = "Create"}
    $menu += [pscustomobject]@{"id"=$menu.Count;"Option" = "Edit"}
    $menu += [pscustomobject]@{"id"=$menu.Count;"Option" = "View"}

    $option = 999
    while($option -ne 0)
    {
        $menu | Format-Table Id,Option | Out-Host
        $option = Read-Host -Prompt "Enter an option <id>"

        if($option -eq 3)
        {
            $buttons = OpCon_GetAllServiceRequests -url $url -token $token
            $buttons | Format-Table -Property id,name,@{Label="category"; Expression={$_.serviceRequestCategory.name}},disableRule,hideRule,documentation | Out-Host
        }
        elseif($option -eq 2)
        {
            $buttons = OpCon_GetAllServiceRequests -url $url -token $token
            $buttons | Format-Table -Property id,name,documentation | Out-Host
            $buttonOption = Read-Host -Prompt "Enter a button <id>"

            $editMenu = @()
            $editMenu += [pscustomobject]@{"Id"=$editMenu.Count;"Option" = "Name"}
            $editMenu += [pscustomobject]@{"Id"=$editMenu.Count;"Option" = "Choice Dropdown"}
            $editMenu += [pscustomobject]@{"Id"=$editMenu.Count;"Option" = "Documentation"}
            $editMenu += [pscustomobject]@{"Id"=$editMenu.Count;"Option" = "Disable Rule"}
            $editMenu += [pscustomobject]@{"Id"=$editMenu.Count;"Option" = "Hide Rule"}
            #$editMenu += [pscustomobject]@{"Id"=$editMenu.Count;"Option" = "HTML"}
            
            $editMenu | Format-Table -Property id,option | Out-Host
            $editOption = Read-Host -Prompt "Enter an option <id> to edit"

            if($editOption -eq 0)
            {
                $newValue = Read-Host "Enter a new name"
                OpCon_UpdateSSButton -url $url -token $token -button $buttons[$buttonOption].name -field "name" -value $newValue | Out-Host
            }
            elseif($editOption -eq 1)
            { 
                $choiceDropdowns = OpCon_GetAllServiceRequestChoice -url $url -token $token -button ($buttons | Where-Object{ $_.id -eq $buttonOption }).name
                $choiceDropdowns | Format-Table -Property id,name,caption,value | Out-Host

                $choiceOptions = @()
                $choiceOptions += [pscustomobject]@{"id"=$choiceOptions.Count;"Option" = "Exit"}
                $choiceOptions += [pscustomobject]@{"id"=$choiceOptions.Count;"Option" = "Add"}
                $choiceOptions += [pscustomobject]@{"id"=$choiceOptions.Count;"Option" = "Delete"}
                $choiceOptions += [pscustomobject]@{"id"=$choiceOptions.Count;"Option" = "Edit (future enhancement)"}
                $choiceOptions | Format-Table -Property Id,Option | Out-Host
                $choicePrompt = Read-Host -Prompt "Enter an option <id>"
                
                if($choicePrompt -eq 3)
                {
                    $choiceDropdowns | Format-Table -Property id,name,caption,value | Out-Host
                    $dropdownPrompt = Read-Host "Enter an option to edit <id>"
                    $newCaption = Read-Host "Enter a new caption (blank to not change)"
                    $newValue = Read-Host "Enter a new value (blank to not change)"
                }
                elseif($choicePrompt -eq 1)
                {
                    $breakLoop = $false
                    while(!$breakLoop)
                    {
                        $choiceDropdowns = OpCon_GetAllServiceRequestChoice -url $url -token $token -button ($buttons | Where-Object{ $_.id -eq $buttonOption }).name
                        $dropDownAdd = Read-Host "Enter the <name> of the dropdown to add a choice (blank to exit)"

                        if($choiceDropdowns | Where-Object{ $_.name -eq $dropdownAdd })
                        {
                            $newCaption = Read-Host "Enter a new caption (blank to not change)"
                            $newValue = Read-Host "Enter a new value (blank to not change)"
                            $result = opc-addsschoice -url $url -token $token -id $buttonOption -addname $newCaption -addvalue $newValue -getdropdown $dropdownAdd
                            $breakLoop = $true
                        }
                        elseif($dropdownAdd -eq "")
                        { $breakLoop = $true }
                        else
                        { Write-Host "Invalid option entered" }
                    }
                }
                elseif($choicePrompt -eq 2)
                {
                    $breakLoop = $false
                    while(!$breakLoop)
                    {
                        $choiceDropdowns = OpCon_GetAllServiceRequestChoice -url $url -token $token -button ($buttons | Where-Object{ $_.id -eq $buttonOption }).name
                        $choiceDropdowns | Format-Table -Property id,name,caption,value | Out-Host
                        $deleteChoice = Read-Host "Enter the <name> of the dropdown to delete a choice (blank to exit)"
                        if($choiceDropdowns | Where-Object{ $_.name -eq $deleteChoice })
                        {
                            $deleteChoice = Read-Host "Enter an <id> to delete"
                            opc-deletesschoice -url $url -token $token -id $buttonOption -removeitem $choiceDropdowns[$deleteChoice].caption -getdropdown $choiceDropdowns[$deleteChoice].name
                        }
                        elseif($deleteChoice -eq "")
                        { $breakLoop = $true }
                        else
                        { Write-Host "Invalid option entered" }
                    }
                }
            }
            if($editOption -eq 2)
            {
                $newValue = Read-Host "Enter new documentation"
                OpCon_UpdateSSButton -url $url -token $token -button $buttons[$buttonOption].name -field "documentation" -value $newValue | Out-Host
            }
            if($editOption -eq 3)
            {
                $newValue = Read-Host "Enter a new disable rule"
                OpCon_UpdateSSButton -url $url -token $token -button $buttons[$buttonOption].name -field "disableRule" -value $newValue | Out-Host
            }
            if($editOption -eq 4)
            {
                $newValue = Read-Host "Enter a new hide rule"
                OpCon_UpdateSSButton -url $url -token $token -button $buttons[$buttonOption].name -field "hideRule" -value $newValue | Out-Host
            }
        }
    }
}
New-Alias "opc-ss" OpConsole_SelfService
New-Alias "opc-buttons" OpConsole_SelfService

# Quit the console if "exit" or "quit" is entered
function OpConsole_Exit
{
    Clear-History
    Clear-Host 
    $PSDefaultParameterValues.Remove("Invoke-RestMethod:SkipCertificateCheck")
    Exit
}
New-Alias "opc-exit" OpConsole_Exit

function OpConsole_ListAll
{
    Get-Alias | where-object{ $_.name.StartsWith("opc") } | Out-Host
    Write-Host "To quit simply type 'exit'"
}
New-Alias "opc-listall" OpConsole_ListAll

function OpConsole_History($cmdArray)
{
    $cmdArray | Format-Table Id,Command | Out-Host

    $selection = Read-Host "Rerun command (id, blank to skip)"
    if($selection -ne "")
    {
        try 
        {
            $cmdArray += [pscustomobject]@{"id"=$cmdArray.Count;"command"=$cmdArray[$selection].command}
            $rerun = $cmdArray[$selection].command  
        }
        catch 
        { Write-Host $_ }
    }
    else 
    { $rerun = "" }

    return $rerun
}
New-Alias "opc-history" OpConsole_History

function OpConsole_Logs($logs)
{
    if($logs.Count -eq 0)
    {
        $tempLogPath = Read-Host "Enter the path to Agent or SAM logs (can be UNC)"
        if(Test-Path $tempLogPath)
        {
            Get-ChildItem -Path $tempLogPath -Filter *.log | ForEach-Object{ 
            $logs += [PSCustomObject]@{"Id"=$logs.Count+1;"Log"=$_.Name;"Location"=$tempLogPath}
        }
            Write-Host $logs.Count"log files found!`r`n"
        }
        else 
        {
            Write-Host "Could not access $tempLogPath"    
        }
    }
    else 
    {
        $logs | Format-Table Id,Location,Log | Out-Host
        $selection = Read-Host "Enter <id> to view, ENTER for new, exit to go back"
        While($selection -ne "exit")
        {
            if($selection -eq "select")
            {
                $logs | Format-Table Id,Location,Log | Out-Host
                $selection = Read-Host "Enter <id> to view, ENTER for new, exit to go back"    
            }

            If($selection -eq "")
            {
                $tempLogPath = Read-Host "Enter the path to Windows Agent or SAM logs (can be UNC)"
                if(Test-Path $tempLogPath)
                {
                    Get-ChildItem -Path $tempLogPath -Filter *.log | ForEach-Object{ 
                        $logs += [PSCustomObject]@{"Id"=$logs.Count+1;"Log"=$_.Name;"Location"=$tempLogPath}
                    }
                    Write-Host $logs.Count"log files found!`r`n"
                }
                else 
                { Write-Host "Could not access $tempLogPath" }                    
            }
            
            Get-Content -Path ($logs[$selection-1].location + "\" + $logs[$selection-1].log) -Raw | Out-Host
            $selection = Read-Host "Press Enter to refresh, 'exit' to go back, 'select' to choose another log"
            While(($selection -ne "exit") -and ($selection -ne "select"))
            {
                Get-Content -Path ($logs[$selection-1].location + "\" + $logs[$selection-1].log) -Raw | Out-Host
                $selection = Read-Host "Press Enter to refresh, 'exit' to go back, 'select' to choose another log"
            }
        }
    }   
    
    return $logs
}
New-Alias "opc-logs" OpConsole_Logs

function OpConsole_Scripts($url,$token)
{
    #$allScripts = OpCon_GetScripts -url $url -token $token
     

    $scriptName = Read-Host "Script name"
    $scriptVersion = Read-Host "Script version (blank for latest)"
    $execute = Read-Host "Run $scriptName (y for yes, or n/blank for view)"

    $versionArray = @()
    $scriptId = (OpCon_GetScripts -url $url -token $token -scriptname $scriptName).id
    (OpCon_GetScriptVersions -url $url -token $token -id $scriptId).versions | ForEach-Object{ $versionArray += $_.version }
    if($scriptVersion)
    { 
        if($execute -eq "y")
        { Invoke-Expression (OpCon_GetScript -url $url -token $token -scriptId $scriptId -versionId $scriptVersion).Content | Out-Host }
        else
        { (OpCon_GetScript -url $url -token $token -scriptId $scriptId -versionId $scriptVersion).Content | Out-Host }
    }
    else 
    { 
        if($execute -eq "y")
        { Invoke-Expression (OpCon_GetScript -url $url -token $token -scriptId $scriptId -versionId (($versionArray | Measure-Object -Maximum).Maximum)).Content | Out-Host }
        else
        { (OpCon_GetScript -url $url -token $token -scriptId $scriptId -versionId (($versionArray | Measure-Object -Maximum).Maximum)).Content | Out-Host }
    }  
}


<#
.SYNOPSIS

Connects to an OpCon environment

.OUTPUTS

Object with the id,url,token,expiration,release.

.EXAMPLE

C:\PS> opconnect"
#>
Function OpConsole_Connect($logins)
{ 
    $logins | ForEach-Object -Parallel { if($_.active -eq "true" -or $_.active -eq "")
                                            { $_.active = "false" }
                                        }
    $logins | Format-Table Id,Name,User,Expiration,Release,Active | Out-Host
    $opconEnv = Read-Host "Enter OpCon environment <id> (blank to go back)"

    if($opconEnv -gt 0)
    {
        if($logins[$opconEnv].user -eq "")
        { $logins[$opconEnv].user = Read-Host "Enter Username" }

        if($logins[$opconEnv].url -eq "")
        { 
            Write-Host "Blank URL for "$logins[$opconEnv].name
            Break
        }

        if($logins[$opconEnv].token -ne "")
        {
            if(!(OpConsole_CheckExpiration -connection $logins[$opconEnv]))
            { $logins[$opconEnv].token = "" }
        }

        if($logins[$opconEnv].token -eq "")
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
                Write-Host "Connected to"($logins | Where-Object{ $_.active -eq "true"}).name", expires at"($logins | Where-Object{ $_.active -eq "true"}).expiration
            }
            else
            { $suppress = opc-connect -logins $logins }
        }
        else 
        { 
            $logins[$opconEnv].active = "true" 
            Clear-Host
            Write-Host "Connected to"($logins | Where-Object{ $_.active -eq "true"}).name", expires at"($logins | Where-Object{ $_.active -eq "true"}).expiration
        }
    }
    elseif(($logins[$opconEnv].name -eq "Create New") -and ($opconEnv -ne ""))
    {
        Write-Host "Creating new OpCon connection....."
        $tls = Read-Host "TLS (y/n, or blank for TLS)"
    
        if($tls -eq "y" -or $tls -eq "")
        { $tls = "https://" }
        elseif($tls -eq "n") 
        { $tls = "http://" }
        else
        { 
            Write-Host "Invalid TLS option, using https"
            $tls = "https://"
        }
    
        $hostname = Read-Host "OpCon API hostname or ip"
        $port = Read-Host "API Port (blank for 9010)"
        
        if($port -eq "")
        { $port = "9010" }
    
        $url = $tls + $hostname + ":" + $port
        $name = Read-Host "Environment name (optional)"
        
        if($name -eq "")
        { $name = $url.Substring($url.IndexOf("//")+2) }
    
        $user = Read-Host "Enter Username" #-AsSecureString 
        $password = Read-Host "Enter Password" -AsSecurestring          
        $auth = OpCon_Login -url $url -user $user -password ((New-Object PSCredential "user",$password).GetNetworkCredential().Password)
        
        if($auth.id)
        {
            $password = "" # Clear out password variable
            $logins += [pscustomobject]@{"id"=$logins.Count;"name"=$name;"url"=$url;"user"=$user;"token"=("Token " + $auth.id);"expiration"=($auth.validUntil);"release"=((OpCon_APIVersion -url $url).opConRestApiProductVersion);"active"="true"}
            Clear-Host # Clears console
            Write-Host "Connected to"($logins | Where-Object{ $_.active -eq "true"}).name", expires at"($logins | Where-Object{ $_.active -eq "true"}).expiration
        }
        else
        { $suppress = opc-connect -logins $logins }
    }

    return $logins
}
New-Alias "opc-connect" OpConsole_Connect
New-Alias "opc-select" OpConsole_Connect

Function OpConsole_ReadLogErrors($path)
{
    If(test-path $path)
    {
        $fileObj = @()
        $contents = Get-Content -Path $path
        For($x=0;$x -lt $contents.Count;$x++)
        {
            if(($contents[$x] -like "*failed*" -or $contents[$x] -like "*unable*") -and $contents[$x] -notlike "*processing event*")
            {
                $fileObj += [pscustomobject]@{"Date/Time"=$contents[$x].Substring(0,23);"Reason"=$contents[$x].Substring(27).Trim()} 
            }
        }
        return $fileObj 
    }
    Else
    {
        Write-Host "Could not access $path"
    }
}
New-Alias "opc-readlog" OpConsole_ReadLogErrors

Function OpConsole_ReadSAMLogEvents($path)
{
    If(test-path $path)
    {
        $fileObj = @()
        $contents = Get-Content -Path $path
        For($x=0;$x -lt $contents.Count;$x++)
        {
            if($contents[$x] -like "*processing event*" -and $contents[$x] -notlike "*processing events*")
            {
                $fileObj += [pscustomobject]@{"Date/Time"=$contents[$x].Substring(0,23);"Event"=$contents[$x].Substring(43,$contents[$x].IndexOf("Received") - 44).Trim();"Location"=$contents[$x].Substring($contents[$x].IndexOf("Received")).Trim()} 
            }
        }
        return $fileObj 
    }
    Else
    {
        Write-Host "Could not access $path"
    }
}
New-Alias "opc-readsamlog" OpConsole_ReadSAMLogEvents

function Opconsole_StressTest($url,$token)
{
    $job = Read-Host "Enter the job name to test"
    $frequency = Read-Host "Enter the frequency name (blank for 'onrequest')"

    if($frequency -eq "")
    { $frequency = "OnRequest" }

    $schedule = Read-Host "Enter the schedule name"
    $executions = Read-Host "How many job executions should be tested? (blank for 10)"

    if($executions -eq "")
    { $executions = 10 }

    for($x=0;$x -lt $executions;$x++)
    {
        $jobs = $jobs + ";$job"
        $freqs = $freqs + ";$frequency"
    }

    $sid = (OpCon_GetSchedule -url $url -token $token -sname "$schedule").id
    OpCon_ScheduleAction -url $url -token $token -sid $sid -jname $jobs.TrimStart(";") -frequency $freqs.TrimStart(";") -action "JOB:ADD" | Out-Host
}
New-Alias "opc-stress" OpConsole_StressTest
New-Alias "opc-stresstest" OpConsole_StressTest

function OpConsole_CheckExpiration($connection)
{
    if($connection.token -ne "")
    {
        $date1 = Get-Date -date $connection.expiration | Get-Date -Format "MM/dd/yyyy HH:mm"
        $date2 = Get-Date -Format "MM/dd/yyyy HH:mm"
        if($date2 -gt $date1)
        { 
            Write-Host "API token expired for "$connection.Name" with user: "$connection.user
            return $false
        }
        else 
        { return $true }
    }
    else
    { return $false }
}

function OpConsole_BatchUsers($url,$token)
{
    $menu = @()
    $menu += [pscustomobject]@{"id"=$menu.Count;"Option" = "Exit"}
    $menu += [pscustomobject]@{"id"=$menu.Count;"Option" = "Create"}
    $menu += [pscustomobject]@{"id"=$menu.Count;"Option" = "View"}
    $menu += [pscustomobject]@{"id"=$menu.Count;"Option" = "Update"}

    $platforms = @()
    $platforms += [pscustomobject]@{"id"=$platforms.Count;"Option"="IBMi"}
    $platforms += [pscustomobject]@{"id"=$platforms.Count;"Option"="MCP"}
    $platforms += [pscustomobject]@{"id"=$platforms.Count;"Option"="OpenVMS"}
    $platforms += [pscustomobject]@{"id"=$platforms.Count;"Option"="SQL"}
    $platforms += [pscustomobject]@{"id"=$platforms.Count;"Option"="Unix"}
    $platforms += [pscustomobject]@{"id"=$platforms.Count;"Option"="Windows"}

    $option = 999
    while($option -ne 0)
    {
        $batchUsers = OpCon_GetBatchUser -url $url -token $token #| Out-Host

        $menu | Format-Table Id,Option | Out-Host
        $option = Read-Host -Prompt "Enter an option <id>"

        If($option -eq 1)
        {
            $roles = OpCon_GetRoles -url $url -token $token

            $loginName = Read-Host "Batch User login"
            $platforms | Format-Table Id,Option | Out-Host
            $platform = Read-Host -Prompt "Enter Platform <id>"
            $userPassword = Read-Host "Batch User password" -AsSecureString
            $roles | Format-Table Id,Name | Out-Host
            $roleIds = Read-Host "Enter role <id> (seperate by comma for multiple)"
            $roleArray = $roleIds.Split(",")

            $roleIdArray = @()
            for($x=0;$x -lt $roleArray.Count;$x++)
            { 
                $currentRole = $roles | Where-Object{ $_.id -eq $roleArray[$x] }
                $roleIdArray += [pscustomobject]@{ "id"=$currentRole.id;"name"=$currentRole.name }  
            }
            
            OpCon_CreateBatchUser -url $url -token $token -platformName $platforms[$platform].Option -loginName $loginName -password ((New-Object PSCredential "user",$userPassword).GetNetworkCredential().Password) -roleIds $roleIdArray | Out-Host
        }
        ElseIf($option -eq 2)
        {
            $batchUser = Read-Host -Prompt "Batch User login (blank if all)"
            $platforms | Format-Table Id,Option | Out-Host
            $os = Read-Host -Prompt "Platform (blank if all)"
            
            if($batchUser -eq "")
            { $batchUsers | Out-Host }
            else
            { 
                if($os -eq "")
                { $result = $batchUsers | Where-Object{ $_.loginName -like "*$batchUser*" } }
                else 
                { $result = $batchUsers | Where-Object{ ($_.loginName -like "*$batchUser*") -and ($_.platform.name -eq $platforms[$os].Option) } }

                If($result.Count -eq 0)
                { Write-Host "No batch users with that login on the selected platform." }
                else 
                { $result | Out-Host }
            }
        }
        ElseIf($option -eq 3)
        {
            $batchUser = Read-Host -Prompt "Batch User login"

            $batchUserFields = @()
            $batchUserFields += [pscustomobject]@{"id"=$batchUserFields.Count;"Option"="loginName"}
            $batchUserFields += [pscustomobject]@{"id"=$batchUserFields.Count;"Option"="password"}

            $batchUserFields | Format-Table Id,Option | Out-Host
            $field = Read-Host -Prompt "Enter a field to update <id>"

            $specificUser = OpCon_GetBatchUser -url $url -token $token -loginName $batchUser
            if($field -eq 1)
            { 
                $value = Read-Host -Prompt "Enter the new value" -AsSecureString 
                OpCon_UpdateBatchUser -url $url -token $token -id $specificUser[0].id -field $batchUserFields[$field].Option -value ((New-Object PSCredential "user",$value).GetNetworkCredential().Password) | Out-Host
            }
            else 
            { 
                $value = Read-Host -Prompt "Enter the new value" 
                OpCon_UpdateBatchUser -url $url -token $token -id $specificUser[0].id -field $batchUserFields[$field].Option -value $value | Out-Host
            }
        }
    }    
}
New-Alias "opc-batchuser" OpConsole_BatchUsers

function OpConsole_Expression($url,$token)
{
    $expression = Read-Host "Enter an expression to evaluate"
    $evaluation = OpCon_PropertyExpression -url $url -token $token -expression $expression 
    $evaluation | Out-Host
    $global:evalresult = $evaluation.result
    Write-Host "Result saved to variable `$evalresult"
}
New-Alias "opc-eval" OpConsole_Expression

function msgbox {
    param (
        [string]$Message,
        [string]$Title = 'Message box title',   
        [string]$buttons = 'ok'
    )   
    # Load the assembly
    Add-Type -AssemblyName System.Windows.Forms | Out-Null
     
    # Define the button types
    switch ($buttons) {
       'ok' {$btn = [System.Windows.Forms.MessageBoxButtons]::OK; break}
       #'okcancel' {$btn = [System.Windows.Forms.MessageBoxButtons]::OKCancel; break}
       #'YesNoCancel' {$btn = [System.Windows.Forms.MessageBoxButtons]::YesNoCancel; break}
       #'YesNo' {$btn = [System.Windows.Forms.MessageBoxButtons]::yesno; break}
       #'RetryCancel'{$btn = [System.Windows.Forms.MessageBoxButtons]::RetryCancel; break}
       #default {$btn = [System.Windows.Forms.MessageBoxButtons]::RetryCancel; break}
    }
     
    # Display the message box
    $Return=[System.Windows.Forms.MessageBox]::Show($Message,$Title,$btn)
    
    # Display the option chosen by the user:
    #$Return
}

function OpConsole_Reports($url,$token)
{
    $menu = @()
    $menu += [PSCustomObject]@{ id = $menu.Count; Option = "Exit" }
    $menu += [PSCustomObject]@{ id = $menu.Count; Option = "Job Status Report" }
    $menu += [PSCustomObject]@{ id = $menu.Count; Option = "Job Count By Status" }
    $menu += [PSCustomObject]@{ id = $menu.Count; Option = "Jobs Running by Platform" }
    $menu += [PSCustomObject]@{ id = $menu.Count; Option = "Jobs waiting on Threshold/Resource" }
    $menu += [PSCustomObject]@{ id = $menu.Count; Option = "User Report" }
    $menu | Format-Table Id,Option | Out-Host
    $report = Read-Host "Enter an option <id>"

    if($menu[$report].Option -eq "Job Status Report")
    { opc-jobstatus -url $url -token $token }
    elseif($menu[$report].Option -eq "Job Count By Status")
    { opc-jobcountbystatus -url $url -token $token }
    elseif($menu[$report].Option -eq "Jobs Running by Platform")
    { opc-jobsbyplatform -url $url -token $token }
    elseif($menu[$report].Option -eq "Jobs waiting on Threshold/Resource")
    { opc-jobswaiting -url $url -token $token }
    elseif($menu[$report].Option -eq "User Report")
    { opc-userreport -url $url -token $token }
}
New-Alias "opc-reports" OpConsole_Reports


function OpConsole_UserReport($url,$token)
{
    $users = OpCon_GetUser -url $url -token $token
    $users | Format-Table Id,LoginName,Email,LastLoggedIn | Out-Host
}
New-Alias "opc-userreport" OpConsole_UserReport

function OpConsole_JobsRunningByPlatformReport($url,$token)
{
    $agents = OpCon_GetAgent -url $url -token $token
    $agents | Format-Table Id,Name,@{Label="Platform";Expression={$_.type.description} },CurrentJobs | Out-Host

    $selection = Read-Host "Enter an id to view running jobs on platform"
    OpCon_GetDailyJob -url $url -token $token | Where-Object{ 
                                                                ($_.startMachine.name -eq ($agents | Where-Object{$_.id -eq $selection }).name) -and ($_.status.category -eq "Running")
                                                            } | Format-Table @{Label="Schedule";Expression={$_.schedule.name} },@{Label="Date";Expression={$_.schedule.date.ToString().SubString(0,9)} },Name -Wrap | Out-Host

}
New-Alias "opc-jobsbyplatform" OpConsole_JobsRunningByPlatformReport

function OpConsole_JobsWaitingReport($url,$token)
{
    $date = Read-Host "Enter schedule date (yyyy-MM-dd, blank for today)"
    if($date -eq "")
    { $date = Get-Date -Format "yyyy-MM-dd" }
    else
    { $date = $date | Get-Date -Format "yyyy-MM-dd" }

    $jobs = OpCon_GetDailyJob -url $url -token $token -date $date | Where-Object{$_.status.description -eq "Wait Threshold/Resource Dependency"} 
    
    if($jobs.Count -gt 0)
    { $jobs | Format-Table UId,@{Label="Date";Expression={$_.schedule.date.ToString().SubString(0,9)} },@{Label="Schedule";Expression={$_.schedule.name} },Name | Out-Host }
    else
    { Write-Host "No jobs found waiting on threshold/resources" }
}
New-Alias "opc-jobswaiting" OpConsole_JobsWaitingReport

function OpConsole_JobCountByStatusReport($url,$token)
{
    $subMenu =@()
    $subMenu += [PSCustomObject]@{Id=$subMenu.Count;Option="Go Back"}
    $subMenu += [PSCustomObject]@{Id=$subMenu.Count;Option="Machine"}
    $subMenu += [PSCustomObject]@{Id=$subMenu.Count;Option="Tags"}
    $subMenu += [PSCustomObject]@{Id=$subMenu.Count;Option="All"}

    $subMenu | Format-Table Id,Option | Out-Host
    $selectSubMenu = Read-Host "Enter a filter option <id>"

    if($subMenu[$selectSubMenu].Option -eq "Go Back")
    { $suppress = opc-reports -url $url -token $token }
    elseif($subMenu[$selectSubMenu].Option -eq "Machine")
    { 
        $machines = OpCon_GetAgent -url $url -token $token
        $machines | Format-Table Id,Name,@{Label="Type";Expression={ $_.type.description } }
        $selectMachine = Read-Host "Enter a machine <id>"

        OpCon_GetDailyJobsCountByStatus -url $url -token $token -machine ($machines | Where-Object{$_.id -eq $selectMachine} ).name  | Out-Host
    }
    elseif($subMenu[$selectSubMenu].Option -eq "Tags")
    {
        $tags = Read-Host "Enter a tag to filter"
        OpCon_GetDailyJobsCountByStatus -url $url -token $token -tags $tags | Out-Host
    }
    elseif($subMenu[$selectSubMenu].Option -eq "All")
    { OpCon_GetDailyJobsCountByStatus -url $url -token $token | Out-Host }
}
New-Alias "opc-jobcountbystatus" OpConsole_JobCountByStatusReport

function OpConsole_JobStatusReport($url,$token)
{
    $subMenu = @()
        $subMenu += [PSCustomObject]@{ id = $subMenu.Count; Option = "Exit" }
        $subMenu += [PSCustomObject]@{ id = $subMenu.Count; Option = "All"; OpCon = "*" }
        $subMenu += [PSCustomObject]@{ id = $subMenu.Count; Option = "Failed"; OpCon = "failed" }
        #$subMenu += [PSCustomObject]@{ id = $subMenu.Count; Option = "Waiting"; OpCon = "waiting" }
        #$subMenu += [PSCustomObject]@{ id = $subMenu.Count; Option = "On Hold"; OpCon = "held" }
        #$subMenu += [PSCustomObject]@{ id = $subMenu.Count; Option = "Cancelled"; OpCon = "cancelled" }
        #$subMenu += [PSCustomObject]@{ id = $subMenu.Count; Option = "Under Review"; OpCon = "underReview" }
        #$subMenu += [PSCustomObject]@{ id = $subMenu.Count; Option = "Skipped"; OpCon = "skipped" }
        $subMenu | Format-Table Id,Option | Out-Host

        $status = Read-Host "Enter a status <id>" # (comma seperated)"
        $date = Read-Host "Enter a schedule date (yyyy-mm-dd, blank for today)"
        
        if($date -eq "")
        { $date = Get-Date -Format "yyyy-MM-dd" }
        else
        { $date = $date | Get-Date -Format "yyyy-MM-dd" }

        $result = OpCon_Reports -url $url -token $token -status $subMenu[$status].OpCon
        
        if($result.Count -gt 0)
        { 
            $result = $result | Where-Object{$date -eq ( $_.schedule.date | Get-Date -Format "yyyy-MM-dd") } | Where-Object{ $_.jobType.description -ne "Container" } 

            $resultMenu = @()
            $result | ForEach-Object{ $resultMenu += [pscustomobject]@{Id=$resultMenu.Count;Date=$date;JobId=$_.id;"Path"=$_.uniqueJobId;"Machine"=$_.startMachine.name;Status=$_.status.description} }
            $resultMenu | Format-Table Id,Date,Machine,Status,@{Label="Schedule|Path";Expression={$_.Path} } -Wrap | Out-Host

            $outputSelection = Read-Host "Enter an option to view job output <id> (blank to go back)"
            if($outputSelection -ne "")
            {
                $jobNumber = (OpCon_GetDailyJob -url $url -token $token -id ($resultMenu[$outputSelection].JobId)).jobNumber
                $jobOutput = (OpCon_GetJobOutput -url $url -token $token -jobNumber $jobnumber).jobInstanceActionItems[0].data 
                $global:jobOutput | Out-Host
                Write-Host "Job output saved to variable `$jobOutput"
            }
            $suppress = opc-reports -url $url -token $token
        }
        else
        { Write-Host "No jobs found with that status on $date" }
}
New-Alias "opc-jobstatus" OpConsole_JobStatusReport