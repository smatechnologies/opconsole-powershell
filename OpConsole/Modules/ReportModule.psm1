function OpConsole_Reports($url,$token)
{
    $menu = New-Object System.Collections.ArrayList
    $menu.Add([PSCustomObject]@{ id = $menu.Count; Option = "Exit" }) | Out-Null
    $menu.Add([PSCustomObject]@{ id = $menu.Count; Option = "Job Count By Status" }) | Out-Null
    $menu.Add([PSCustomObject]@{ id = $menu.Count; Option = "Jobs Running by Platform" }) | Out-Null
    $menu.Add([PSCustomObject]@{ id = $menu.Count; Option = "Job Status Report" }) | Out-Null
    $menu.Add([PSCustomObject]@{ id = $menu.Count; Option = "Jobs waiting on Threshold/Resource" }) | Out-Null
    $menu.Add([PSCustomObject]@{ id = $menu.Count; Option = "User Report" }) | Out-Null
    $menu.Add([PSCustomObject]@{ id = $menu.Count; Option = "Agent Report" }) | Out-Null
    $menu | Format-Table Id,Option | Out-Host
    $report = Read-Host "Enter an option <id>"

    Switch(($menu[$report].Option))
    {
        "Exit"                               {}
        "Job Count By Status"                { opc-jobcountbystatus -url $url -token $token; break }
        "Jobs Running by Platform"           { opc-jobsbyplatform -url $url -token $token; break }
        "Job Status Report"                  { opc-jobstatus -url $url -token $token; break }
        "Jobs waiting on Threshold/Resource" { opc-jobswaiting -url $url -token $token; break }
        "User Report"                        { opc-userreport -url $url -token $token; break }
        "Agent Report"                       { opc-agentreport -url $url -token $token; break }
        Default                              { Write-Host "Invalid report selection"; opc-reports -url $url -token $token; break }
    }
}
New-Alias "opc-reports" OpConsole_Reports

function OpConsole_AgentReport($url,$token)
{
    $agents = OpCon_GetAgent -url $url -token $token
    Write-Host "OpCon Agent report"
    Write-Host "--------------------"
    $agents | Format-Table Id,Name,@{Label="status";Expression={
                                                                    Switch ($_.status.state)
                                                                    {
                                                                        "U" { "Up"; break }
                                                                        "D" { "Down"; break }
                                                                        "L" { "Limited"; break }
                                                                        "E" { "Error"; break }
                                                                        "W" { "Waiting"; break }
                                                                    }
                                                                }
                                    },CurrentJobs,LastUpdate
}
New-Alias "opc-agentreport" OpConsole_AgentReport

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
            $result = $result.Where({ $date -eq ( $_.schedule.date | Get-Date -Format "yyyy-MM-dd")} ).Where({ $_.jobType.description -ne "Container" })

            $resultMenu = New-Object System.Collections.ArrayList
            $result | ForEach-Object{ $resultMenu.Add([pscustomobject]@{Id=$resultMenu.Count;Date=$date;JobId=$_.id;"Path"=$_.uniqueJobId;"Machine"=$_.startMachine.name;Status=$_.status.description}) } | Out-Null
            $resultMenu | Format-Table Id,Date,Machine,Status,@{Label="Schedule|Path";Expression={$_.Path} } -Wrap | Out-Host

            $outputSelection = Read-Host "Enter an option to view job output <id> (blank to go back)"
            if($outputSelection -ne "")
            {
                $jobNumber = (OpCon_GetDailyJob -url $url -token $token -id ($resultMenu[$outputSelection].JobId)).jobNumber
                $global:jobOutput = (OpCon_GetJobOutput -url $url -token $token -jobNumber $jobnumber).jobInstanceActionItems[0].data 
                $global:jobOutput | Out-Host
                Write-Host "Job output saved to variable `$global:jobOutput"
            }
            opc-reports -url $url -token $token | Out-Null
        }
        else
        { Write-Host "No jobs found with that status on $date" }
}
New-Alias "opc-jobstatus" OpConsole_JobStatusReport

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
    (OpCon_GetDailyJobFiltered -url $url -token $token -filter ("startMachine=" + ($agents.Where({$_.id -eq $selection })).name)).Where({ 
                                                        ($_.status.category -eq "Running")
                                                    }) | Format-Table @{Label="Schedule";Expression={$_.schedule.name} },@{Label="Date";Expression={$_.schedule.date.ToString().SubString(0,9)} },Name -Wrap | Out-Host

}
New-Alias "opc-jobsbyplatform" OpConsole_JobsRunningByPlatformReport

function OpConsole_JobsWaitingReport($url,$token)
{
    $date = Read-Host "Enter schedule date (yyyy-MM-dd, blank for today)"
    if($date -eq "")
    { $date = Get-Date -Format "yyyy-MM-dd" }
    else
    { $date = $date | Get-Date -Format "yyyy-MM-dd" }

    $jobs = OpCon_GetDailyJob -url $url -token $token -date $date.Where({$_.status.description -eq "Wait Threshold/Resource Dependency"})
    
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
    { opc-reports -url $url -token $token | Out-Null }
    elseif($subMenu[$selectSubMenu].Option -eq "Machine")
    { 
        $machines = OpCon_GetAgent -url $url -token $token
        $machines | Format-Table Id,Name,@{Label="Type";Expression={ $_.type.description } } | Out-Host
        $selectMachine = Read-Host "Enter a machine <id>"

        OpCon_GetDailyJobsCountByStatus -url $url -token $token -machine ($machines.Where({$_.id -eq $selectMachine})).name  | Out-Host
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