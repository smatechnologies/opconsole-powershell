function OpConsole_Version()
{
    Clear-Host
    $changelog = "
     - Added support for Windows Authentication
     - Reworked menu systems for ease of use
     - Removed non-OpCon central functions
     - Various bug fixes"

    Write-Host "-----------------------------------------------"
    Write-Host " OpConsole v0.80.20220516 tested on OpCon 21.4"
    Write-Host $changelog
    Write-Host "-----------------------------------------------"

    $menu = New-Object System.Collections.ArrayList
    $menu.Add([PSCustomObject]@{ "Id" = $menu.Count;"Option" = "Main Menu" }) | Out-Null
    $option = OpConsole_Menu -items $menu

    Switch ($option)
    {
        default   { return }
    }
}

function OpConsole_Menu($items)
{
    $items | Format-Table Id,Option -Wrap | Out-Host
    $answer = Read-Host -Prompt "Enter Option <Id>"
    Write-Host "-----------------------------------------------"
    if($answer)
    { return $answer }
    else 
    { return ""}
}

function OpConsole_Options()
{
    Clear-Host
    $menu = New-Object System.Collections.ArrayList
    $menu.Add( [PSCustomObject]@{ "Id" = $menu.Count;"Option" = "OpConsole Changelog" }) | Out-Null
    $menu.Add( [PSCustomObject]@{ "Id" = $menu.Count;"Option" = "Main Menu" }) | Out-Null
    $option = OpConsole_Menu -items $menu

    Switch ($option.answer)
    {
        0   { OpConsole_Version; break }
        default { return; break }
    }
}

function OpConsole_Properties($url,$token)
{
    Clear-Host
    Write-Host "-----------------------------------------------"
    Write-Host "        Manage OpCon Global Properties"
    Write-Host "-----------------------------------------------"
    
    $menu = New-Object System.Collections.ArrayList
    $menu.Add([pscustomobject]@{"Id"=$menu.Count;"Option" = "Create"}) | Out-Null
    $menu.Add([pscustomobject]@{"Id"=$menu.Count;"Option" = "View"}) | Out-Null
    $menu.Add([pscustomobject]@{"Id"=$menu.Count;"Option" = "Update"}) | Out-Null
    $menu.Add([pscustomobject]@{"Id"=$menu.Count;"Option" = "Main Menu"}) | Out-Null
    $option = OpConsole_Menu -items $menu

    $properties = OpCon_GetGlobalProperty -url $url -token $token

    Switch ($option)
    {
        0   {
                $propertyName = Read-Host -Prompt "Enter new property name"
                $result = $properties | Where-Object{ $_.name -eq "$propertyName" }
                
                If($result)
                { Read-Host "Property named $propertyName already exists!" }
                else 
                {  
                    # Review requirements for property length and add check
                    $propertyValue = Read-Host -Prompt "Enter the property value"
    
                    $encryptMenu = New-Object System.Collections.ArrayList
                    $encryptMenu.Add([pscustomobject]@{"Id" = $encryptMenu.Count;"Option" = "Encrypt"}) | Out-Null
                    $encryptMenu.Add([pscustomobject]@{"Id" = $encryptMenu.Count;"Option" = "Do not Encrypt (default)"}) | Out-Null
                    $encrypted = OpConsole_Menu -items $encryptMenu
    
                    Switch ($encrypted)
                    {
                        0   {  
                            $newProperty = OpCon_CreateGlobalProperty -url $url -token $token -name $propertyName -value $propertyValue -encrypt $true
                            $newProperty | Format-Table Name,Value,Documentation,Encrypted -Wrap | Out-Host
                            break 
                        }
                        default { 
                            $newProperty = OpCon_CreateGlobalProperty -url $url -token $token -name $propertyName -value $propertyValue -encrypt $false
                            $newProperty | Format-Table Name,Value,Documentation,Encrypted -Wrap | Out-Host
                            break 
                        }
                    }
                }
                Break
            }
        1   {
                $property = Read-Host -Prompt "Global Property name (blank if all)"
            
                if($property -eq "")
                { $properties | Out-Host }
                else
                { 
                    $result = $properties.Where({ $_.name -like "*$property*" } )
        
                    If($result.Count -eq 0)
                    { Read-Host "No properties found with the name -"$property }
                    Else
                    { $result | Format-Table Name,Value,Documentation,encrypted -Wrap | Out-Host }
                }
                Break
            }
        2   {
                $propertyName = Read-Host -Prompt "Property name to update"
                $result = $properties | Where-Object{ $_.name -eq "$propertyName" }
                
                If($result)
                { 
                    Write-Host "Current property information:"
                    $result | Format-Table Name,Value,Documentation,encrypted -Wrap | Out-Host
        
                    $propertyValue = Read-Host -Prompt "New property value"
                    $newValue = OpCon_SetGlobalProperty -url $url -token $token -id $result.id -value $propertyValue
                    $newValue | Format-Table Name,Value,Documentation,encrypted -Wrap | Out-Host
                }
                else 
                { Read-Host "Property named $propertyName not found!" } 
                Break
            }
        default {return}
    }
    
    $menuEnd = New-Object System.Collections.ArrayList
    $menuEnd.Add([pscustomobject]@{"Id"=$menuEnd.Count;"Option" = "Manage another property"}) | Out-Null
    $menuEnd.Add([pscustomobject]@{"Id"=$menuEnd.Count;"Option" = "Main Menu"}) | Out-Null
    $option = OpConsole_Menu -items $menuEnd

    Switch ($option)
    {
        0   {OpConsole_Properties -url $url -token $token; break}
        default {return}
    }
}
New-Alias "opc-property" OpConsole_Properties
New-Alias "opc-properties" OpConsole_Properties

Function OpConsole_SelfService($url,$token)
{
    $menu = @()
    $menu += [pscustomobject]@{"Id"=$menu.Count;"Option" = "Exit"}
    $menu += [pscustomobject]@{"Id"=$menu.Count;"Option" = "Create"}
    $menu += [pscustomobject]@{"Id"=$menu.Count;"Option" = "Edit"}
    $menu += [pscustomobject]@{"Id"=$menu.Count;"Option" = "View"}

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
                        $choiceDropdowns = OpCon_GetAllServiceRequestChoice -url $url -token $token -button ($buttons.Where({ $_.id -eq $buttonOption })).name
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


function OpConsole_Scripts($url,$token)
{
    $scriptsArray = New-Object System.Collections.ArrayList
    $scriptName = Read-Host "Enter script name (wildcards supported *, blank to exit)"
    if($scriptName -ne "")
    { $getScripts = OpCon_GetScripts -url $url -token $token -scriptname $scriptName }
          
    if($getScripts.Count -gt 0)
    {
        $getScripts | ForEach-Object{ $scriptsArray.Add([pscustomobject]@{"Id"=$scriptsArray.Count;"Name"=$_.name;"Type"=$_.type.name;"ScriptId"=$_.id }) | Out-Null } 

        $scriptsArray | Format-Table Id,Name,Type | Out-Host
        $script = Read-Host "Enter a script <id> (blank to go back)"

        if($script -ne "")
        {
            $menu = @()
            $menu += [pscustomobject]@{Id=$menu.Count;Option="Exit"}
            $menu += [pscustomobject]@{Id=$menu.Count;Option="View script versions"}
            $menu += [pscustomobject]@{Id=$menu.Count;Option="Start over"}
        
            $menu | Format-Table Id,Option | Out-Host
            $selection = Read-Host "Enter an option <id>"

            if($menu[$selection].Option -eq "View script versions")
            {
                $versionArray = New-Object System.Collections.ArrayList
                (OpCon_GetScriptVersions -url $url -token $token -id $scriptsArray[$script].ScriptId).versions | ForEach-Object{ $versionArray.Add([pscustomobject]@{Version=$_.version;Comment=$_.message }) | Out-Null }

                $versionArray | Format-Table Version,Comment | Out-Host
                $scriptVersion = Read-Host "Enter a script version <id> (blank to exit)"
                
                if($scriptVersion -ne "")
                {
                    $subMenu = @()
                    $subMenu += [pscustomobject]@{Id=$subMenu.Count;Option="Exit"}
                    $subMenu += [pscustomobject]@{Id=$subMenu.Count;Option="Start over"}
                    $subMenu += [pscustomobject]@{Id=$subMenu.Count;Option="Run script version"}
                    $subMenu += [pscustomobject]@{Id=$subMenu.Count;Option="View script version"}
                    $subMenu | Format-Table Id,Option | Out-Host

                    $execute = Read-Host "Enter an option <id>"
                    Write-Host "================================"
                    
                    if($subMenu[$execute].Option -eq "Run script version")
                    { Invoke-Expression (OpCon_GetScript -url $url -token $token -scriptId $scriptsArray[$script].ScriptId -versionId $scriptVersion).Content | Out-Host }
                    elseif($subMenu[$execute].Option -eq "View script version")
                    { (OpCon_GetScript -url $url -token $token -scriptId $scriptsArray[$script].ScriptId -versionId $scriptVersion).Content | Out-Host }
                    elseif($subMenu[$execute].Option -eq "Start over")
                    { opc-scripts -url $url -token $token }
                }
                else
                { opc-scripts -url $url -token $token }
            }
            elseif($menu[$selection].Option -eq "Start over")
            { opc-scripts -url $url -token $token }
        }
        else
        { opc-scripts -url $url -token $token }
    }
    else
    { 
        if($scriptName -ne "")
        { Write-Host "No scripts found with that name"} 
    }
}
New-Alias "opc-scripts" OpConsole_Scripts


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

function OpConsole_BatchUsers($url,$token)
{
    Clear-Host
    $menu = New-Object System.Collections.ArrayList
    $menu.Add( [pscustomobject]@{"id"=$menu.Count;"Option" = "Create"} ) | Out-Null
    $menu.Add( [pscustomobject]@{"id"=$menu.Count;"Option" = "View"} ) | Out-Null
    $menu.Add( [pscustomobject]@{"id"=$menu.Count;"Option" = "Update"} ) | Out-Null
    $menu.Add( [pscustomobject]@{"id"=$menu.Count;"Option" = "Main Menu"} ) | Out-Null

    $platforms = New-Object System.Collections.ArrayList
    $platforms.Add( [pscustomobject]@{"id"=$platforms.Count;"Option"="IBMi"}) | Out-Null
    $platforms.Add( [pscustomobject]@{"id"=$platforms.Count;"Option"="MCP"}) | Out-Null
    $platforms.Add( [pscustomobject]@{"id"=$platforms.Count;"Option"="OpenVMS"}) | Out-Null
    $platforms.Add( [pscustomobject]@{"id"=$platforms.Count;"Option"="SQL"}) | Out-Null 
    $platforms.Add( [pscustomobject]@{"id"=$platforms.Count;"Option"="Unix"}) | Out-Null
    $platforms.Add( [pscustomobject]@{"id"=$platforms.Count;"Option"="Windows"}) | Out-Null

    $option = 999
    while($menu[$option].Option -ne "Main Menu")
    {
        $batchUsers = OpCon_GetBatchUser -url $url -token $token

        Write-Host "-----------------------------------------------"
        Write-Host "          Manage OpCon Batch Users"
        Write-Host "-----------------------------------------------"
        $option = OpConsole_Menu -items $menu
        Switch ($option)
        {
            0   {
                    # Batch user login
                    $loginName = Read-Host "Batch User login"
                    Write-Host "-----------------------------------------------"

                    # Batch user platform
                    Write-Host "Platform:"
                    $platform = OpConsole_Menu -items $platforms

                    # Batch user password
                    $userPassword = Read-Host "Batch User password" -AsSecureString
                    Write-Host "-----------------------------------------------"

                    # Roles assigned to batch user
                    Write-Host "Roles:"
                    $roles = OpCon_GetRoles -url $url -token $token
                    $roles | Format-Table Id,Name | Out-Host
                    $roleIds = Read-Host "Enter Option <id> (seperate by comma for multiple)"
                    $roleArray = $roleIds.Split(",")
        
                    $roleIdArray = New-Object System.Collections.ArrayList
                    for($x=0;$x -lt $roleArray.Count;$x++)
                    { 
                        $currentRole = $roles.Where({ $_.id -eq $roleArray[$x] })
                        $roleIdArray.Add([pscustomobject]@{ "id"=$currentRole.id;"name"=$currentRole.name }) | Out-Null
                    }
                    
                    $createUser = OpCon_CreateBatchUser -url $url -token $token -platformName $platforms[$platform].Option -loginName $loginName -password ((New-Object PSCredential "user",$userPassword).GetNetworkCredential().Password) -roleIds $roleIdArray
                    if($createUser.loginName)
                    {   
                        Write-Host "-----------------------------------------------"
                        Write-Host "Batch user created!"
                        $createUser | Format-Table LoginName,@{Label="Platform";Expression={$_.platform.name}} -Wrap | Out-Host 
                    }
                }
            1   {
                    $batchUser = Read-Host -Prompt "Batch User login (blank if all)"
                    
                    if($batchUser -eq "")
                    { $batchUsers | Format-Table LoginName,@{Label="Platform";Expression={$_.platform.name}} -Wrap | Out-Host }
                    else
                    { 
                        $result = $batchUsers.Where({ ($_.loginName -like "*$batchUser*")})
        
                        If($result.Count -eq 0)
                        { Read-Host "No batch users with login"$batchUser }
                        else 
                        { $result | Format-Table LoginName,@{Label="Platform";Expression={$_.platform.name}} -Wrap | Out-Host }
                    }
                }
            2   {
                    $batchUser = Read-Host -Prompt "Batch User login to change"
        
                    $specificUser = OpCon_GetBatchUser -url $url -token $token -loginName $batchUser
                    if($specificUser)
                    {
                        $batchUserFields = New-Object System.Collection.ArrayList
                        $batchUserFields.Add( [pscustomobject]@{"id"=$batchUserFields.Count;"Option"="Change loginName"} )
                        $batchUserFields.Add( [pscustomobject]@{"id"=$batchUserFields.Count;"Option"="Change password"} )
                        $field = OpConsole_Menu -items $batchUserFields

                        Switch ($field)
                        {
                            0   {
                                    $value = Read-Host -Prompt "Enter the new user login" 
                                    $updateUser = OpCon_UpdateBatchUser -url $url -token $token -id $specificUser[0].id -field "loginName" -value $value
                                    $updateUser | Format-Table LoginName,@{Label="Platform";Expression={$_.platform.name}} -Wrap | Out-Host 
                                }
                            1   {
                                    $value = Read-Host -Prompt "Enter the new password" -AsSecureString 
                                    $updateUser = OpCon_UpdateBatchUser -url $url -token $token -id $specificUser[0].id -field "password" -value ((New-Object PSCredential "user",$value).GetNetworkCredential().Password)
                                    $updateUser | Format-Table LoginName,@{Label="Platform";Expression={$_.platform.name}} -Wrap | Out-Host 
                                }
                        }
                    }
                    else{ Read-Host "No batch user found with login $batchuser" }
                }
            default { return }    
        }
    }    
}
New-Alias "opc-batchuser" OpConsole_BatchUsers

function OpConsole_Expression($url,$token)
{
    Clear-Host
    Write-Host "-----------------------------------------------"
    Write-Host "              Evaluate Expression"
    Write-Host "-----------------------------------------------"

    $expression = Read-Host "Enter an expression (1 == 1)"
    $evaluation = OpCon_PropertyExpression -url $url -token $token -expression $expression 
    Write-Host "Result ="$evaluation.result 

    $menu = New-Object System.Collections.ArrayList
    $menu.Add([pscustomobject]@{"Id"=$menu.Count;"Option" = "Evaluate another expression"}) | Out-Null
    $menu.Add([pscustomobject]@{"Id"=$menu.Count;"Option" = "Main Menu"}) | Out-Null
    $option = OpConsole_Menu -items $menu

    Switch ($option)
    {
        0   {OpConsole_Expression -url $url -token $token; break}
        default {return}
    }
}
New-Alias "opc-eval" OpConsole_Expression


function OpConsole_DailyJobs($url,$token)
{
    $menu = @()
    $menu += [pscustomobject]@{"id"=$menu.Count;"Option" = "Exit"}
    $menu += [pscustomobject]@{"id"=$menu.Count;"Option" = "Daily Jobs by name/schedule"}
    $menu += [pscustomobject]@{"id"=$menu.Count;"Option" = "Failed Jobs for date"}
    $menu += [pscustomobject]@{"id"=$menu.Count;"Option" = "Running Jobs by server"}

    $option = 999
    while($option -ne 0)
    {
        $menu | Format-Table Id,Option | Out-Host
        $option = Read-Host -Prompt "Enter an option <id>"

        If($option -eq 1)
        {
            $date = Read-Host -Prompt "Enter date (MM/DD/YY) or blank for today"
            
            if($date -eq "")
            { $date = Get-Date -format "yyyy-MM-dd" }

            $jobsCounter = 0
            $dailyJobs = OpCon_GetDailyJobsByStatus -url $url -token $token -date $date -status "Failed"
            if($dailyJobs)
            { $dailyJobs | Sort-Object -Descending -Property name | Format-Table @{Label="Id";Expression={$jobsCounter++}},@{Label="Schedule Name";Expression={$_.schedule.name}},@{Label="Job Name";Expression={$_.name}},@{Label="Type";Expression={$_.jobType.description}},@{Label="Start Time";Expression={$_.computedStartTime.time}},@{Label="End Time";Expression={$_.computedEndTime.time}},@{Label="Status";Expression={$_.status.description}},@{Label="Exit Code";Expression={$_.terminationDescription}} -AutoSize -RepeatHeader -Wrap | Out-Host }
            else {
                Write-Host "No failed jobs found for $date"
            }
        }
    }
}
New-Alias "opc-dailyjobs" OpConsole_DailyJobs