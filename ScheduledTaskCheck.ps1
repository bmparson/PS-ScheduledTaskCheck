#Global Variable Declaration
$taskarray = @()
$failedTasks = @()
$taskReport = @()
$servername = Get-Content "c:\scripts\ScheduledTaskCheck\servers.txt"
$previousFailedTasks = $null
$msg = $null
$shutdownMsg = $null
$threshold = $null
$reboot = $false

#Function to get all Folders and SubFolders in the Scheduler Service
function Get-AllTaskSubFolders {
    [cmdletbinding()]
    param (
        # Set to use $Schedule as default parameter so it automatically list all files
        # For current schedule object if it exists.
        $FolderRef = $Schedule.getfolder("\")
    )
    if ($FolderRef.Path -eq '\') {
        $FolderRef
    }
    if (-not $RootFolder) {
        $ArrFolders = @()
        if(($Folders = $folderRef.getfolders(1))) {
            $Folders | ForEach-Object {
                $ArrFolders += $_
                if($_.getfolders(1)) {
                    Get-AllTaskSubFolders -FolderRef $_
                }
            }
        }
        $ArrFolders
    }
}
#Import Previous Failed Tasks and Threshold File for Record keeping and Messaging logic
if(Test-Path "c:\scripts\ScheduledTaskCheck\failedtaskreport.csv")
{
    $previousFailedTasks = Import-Csv "c:\scripts\ScheduledTaskCheck\failedtaskreport.csv"
}    
if(Test-Path "c:\scripts\ScheduledTaskCheck\taskfailurethreshold.csv")
{
    $threshold = Import-Csv "c:\scripts\ScheduledTaskCheck\taskfailurethreshold.csv"
}    
#Loop through each server defined in servers.txt
foreach($server in $servername)
{
    $boot=(get-date) - (gcim Win32_OperatingSystem -computername $server).LastBootUpTime
    if($boot.days -eq 0 -and $boot.hours -eq 0 -and $boot.minutes -le 5)
    {
        $reboot = $true
    }
    Else
    {
        $reboot = $false
    }
    
    if(-not $reboot)
    {
#Connect to Scheduler Service on Server
        $schedule = new-object -com("Schedule.Service")
        $schedule.connect($server)
#Get all Enabled Tasks
        $AllFolders = Get-AllTaskSubFolders
        foreach($folder in $AllFolders)
        {
#Exclude Microsoft Created Tasks        
            if($folder.path -notlike "\Microsoft*")
            {
                $tasks = $schedule.getfolder($folder.path).gettasks(0)
                $ourtasks = $tasks |select Name,Path,LastRunTime,LastTaskResult,State,@{Name="RunAs";Expression={[xml]$xml = $_.xml ; $xml.Task.Principals.principal.userID}},@{Name="Server";Expression={$server}}
                $reportTasks = $tasks |select Name,State,XML,@{Name="RunAs";Expression={[xml]$xml = $_.xml ; $xml.Task.Principals.principal.userID}},@{Name="Command";Expression={[xml]$xml = $_.xml ; $xml.Task.Actions.exec.command}},@{Name="Arguments";Expression={[xml]$xml = $_.xml ; $xml.Task.Actions.exec.arguments}},@{Name="Triggers";Expression={[xml]$xml = $_.xml ; $xml.Task.triggers}},@{Name="Server";Expression={$server}}
                $enabledTasks = $ourtasks | where {$_.state -ne 1}
                $enabledReportTasks = $reportTasks | where {$_.state -ne 1}

#Filter all Enabled Tasks for unimportant tasknames and put into an array
                foreach($task in $enabledtasks)
                {
                    $object = New-Object system.object
                    $object | add-member -type NoteProperty -name Name -Value $task.name
                    $object | add-member -type NoteProperty -name Path -Value $task.path
                    $object | add-member -type NoteProperty -name LastRunTime -Value $task.lastruntime
                    $object | add-member -type NoteProperty -name LastTaskResult -Value $task.LastTaskResult
                    $object | add-member -type NoteProperty -name State -Value $task.state
                    $object | add-member -type NoteProperty -name RunAs -Value ($task | where {$_.runas -ne $null} | foreach-object {$objSID = New-Object System.Security.Principal.SecurityIdentifier ($_.runas)
                    $objUser = $objSID.Translate( [System.Security.Principal.NTAccount])
                    $objuser.value})
                    $object | add-member -type NoteProperty -name Server -Value $task.server
                    $object | add-member -type NoteProperty -name TimesFailed -Value $null
                    $object | add-member -type NoteProperty -name MsgSent -Value $null
                    $object | add-member -type NoteProperty -name ShutdownStateCount -Value $null
                    if($object.name -notlike "Adobe*" -and $object.name -notlike "Optimize Start Menu*" -and $object.name -notlike "GoogleUpdate*")
                    {
                        $taskarray += $object
                    }
                }
                foreach($task in $enabledReportTasks)
                {
                    [xml]$xml=$task.xml
                    $triggerType=$xml.task.Triggers.ChildNodes.name
                    $trigger= $null
                    $months = $null
                    $days = $null
                    $weeks = $null
                    if($triggerType -eq "CalendarTrigger")
                        {
                            $StartBoundary = $xml.task.Triggers.CalendarTrigger.StartBoundary
                            $StartBoundary = $StartBoundary.Split("T")
                            if($xml.task.Triggers.CalendarTrigger.ChildNodes.name.Contains("Repetition"))
                                {
                                    $Interval = $xml.task.Triggers.CalendarTrigger.Repetition.Interval
                                    $Interval = $Interval -replace "PT"
                                    $Interval = $Interval -replace "H", " Hours"
                                    $Interval = $Interval -replace "D", " Days"
                                    $Interval = $Interval -replace "M", " Minutes"
                                    $Interval = " Repeat every " + $Interval
                                    if($xml.task.Triggers.CalendarTrigger.Repetition.ChildNodes.name.Contains("Duration"))
                                        {
                                            $Duration = $xml.task.Triggers.CalendarTrigger.Repetition.Duration
                                            $Duration = $Duration -replace "PT"
                                            $Duration = $Duration -replace "P"
                                            $Duration = $Duration -replace "H", " Hours"
                                            $Duration = $Duration -replace "D", " Days"
                                            $Duration = $Duration -replace "M", " Minutes"
                                            $Duration = "for " + $Duration
                                        }
                                        Else 
                                            {
                                                $Duration="Indefinitely"
                                            }
                                }
                            if($xml.task.Triggers.CalendarTrigger.ChildNodes.name.Contains("ScheduleByDay"))
                                {
                                    $daysinterval = $xml.task.Triggers.CalendarTrigger.ScheduleByDay.DaysInterval + " days"
                                    $daysinterval = $daysinterval -replace "1 days", "day" 
                                    $trigger = "At " + $StartBoundary[1] + " every " + $daysinterval
                                    if($interval)
                                            {
                                                $trigger += " - After triggered repeat every " + $Interval + " for a duration of " + $Duration
                                            } 
                                }
                            Elseif($xml.task.Triggers.CalendarTrigger.ChildNodes.name.Contains("ScheduleByWeek"))
                                {
                                    $weeksinterval = $xml.task.Triggers.CalendarTrigger.ScheduleByWeek.WeeksInterval + " weeks"
                                    $weeksinterval = $weeksinterval -replace "1 weeks", "week"
                                    $dayarray = $xml.task.Triggers.CalendarTrigger.ScheduleByWeek.DaysOfWeek.ChildNodes.name
                                    if($dayarray[0].count -gt 1)
                                        {
                                            for($i=0;$i -lt $dayarray.count;$i++)
                                                {
                                                    if($i -eq 0 -or $i -eq ($dayarray.count-1))
                                                        {
                                                            $days += $dayarray[$i]
                                                        }
                                                    else
                                                        {
                                                            $days += ", " + $dayarray[$i]
                                                        }
                                                }
                                        }
                                        else
                                        {
                                            $days = $dayarray
                                        }
                                    $trigger = "At " + $StartBoundary[1] + " on " + $days + " every " + $weeksinterval + ", starting " + $StartBoundary[0]
                                    if($interval)
                                            {
                                                $trigger += " - After triggered repeat every " + $Interval + " for a duration of " + $Duration
                                            }
                                }
                            Elseif($xml.task.Triggers.CalendarTrigger.ChildNodes.name.Contains("ScheduleByMonth") -and -not($xml.task.Triggers.CalendarTrigger.ChildNodes.name.Contains("ScheduleByMonthDaysOfWeek")))
                                {
                                    $dayarray = $xml.task.Triggers.CalendarTrigger.ScheduleByMonth.DaysOfMonth.Day
                                    $montharray = $xml.task.Triggers.CalendarTrigger.ScheduleByMonth.months.childnodes.name
                                    for($i=0;$i -lt $dayarray.count;$i++)
                                        {
                                            if($i -eq 0 -or $i -eq ($dayarray.count-1))
                                                {
                                                    $days += $dayarray[$i]
                                                }
                                            else
                                                {
                                                    $days += ", " + $dayarray[$i]
                                                }
                                        }
                                    If($montharray[0].count -gt "1")
                                        {
                                            for($i=0;$i -lt $montharray.count;$i++)
                                                {
                                                    if($i -eq 0 -or $i -eq ($montharray.count-1))
                                                        {
                                                            $months += $montharray[$i]
                                                        }
                                                    else
                                                        {
                                                            $months += ", " + $montharray[$i]
                                                        }
                                                }
                                        }
                                      else
                                        {
                                                $months = $montharray
                                        }
                                    $trigger = "At " + $StartBoundary[1] + " on day " + $days + " of " + $months + " starting " + $StartBoundary[0]
                                    if($interval)
                                            {
                                                $trigger += " - After triggered repeat every " + $Interval + " for a duration of " + $Duration
                                            }
                                }
                            Elseif($xml.task.Triggers.CalendarTrigger.ChildNodes.name.Contains("ScheduleByMonthDayOfWeek"))
                                {
                                    $weekarray = $xml.task.Triggers.CalendarTrigger.ScheduleByMonthDayOfWeek.weeks.week
                                    $dayarray = $xml.task.Triggers.CalendarTrigger.ScheduleByMonthDayOfWeek.DaysOfWeek.childnodes.name
                                    $montharray = $xml.task.Triggers.CalendarTrigger.ScheduleByMonthDayOfWeek.months.childnodes.name
                                    for($i=0;$i -lt $weekarray.count;$i++)
                                        {
                                            if($i -eq 0 -or $i -eq ($weekarray.count-1))
                                                {
                                                    $weeks += $weekarray[$i]
                                                }
                                            else
                                                {
                                                    $weeks += ", " + $weekarray[$i]
                                                }
                                        }
                                        $weeks = $weeks -replace "1","First"
                                        $weeks = $weeks -replace "2","Second"
                                        $weeks = $weeks -replace "3","Third"
                                        $weeks = $weeks -replace "4","Fourth"
                                    if($dayarray[0].count -gt 1)
                                        {
                                            for($i=0;$i -lt $dayarray.count;$i++)
                                                {
                                                    if($i -eq 0 -or $i -eq ($dayarray.count-1))
                                                        {
                                                            $days += $dayarray[$i]
                                                        }
                                                    else
                                                        {
                                                            $days += ", " + $dayarray[$i]
                                                        }
                                                }
                                        }
                                    else
                                        {
                                            $days = $dayarray
                                        }
                                    If($montharray[0].count -gt "1")
                                        {
                                            for($i=0;$i -lt $montharray.count;$i++)
                                                {
                                                    if($i -eq 0 -or $i -eq ($montharray.count-1))
                                                        {
                                                            $months += $montharray[$i]
                                                        }
                                                    else
                                                        {
                                                            $months += ", " + $montharray[$i]
                                                        }
                                                }
                                        }
                                    else
                                        {
                                                $months = $montharray
                                        }
                                }
                        }
                    elseif($triggerType -eq "TimeTrigger")
                        {
                            $StartBoundary = $xml.task.Triggers.TimeTrigger.StartBoundary
                            $StartBoundary = $StartBoundary.Split("T")
                            if($xml.task.Triggers.TimeTrigger.ChildNodes.name.Contains("Repetition"))
                                {
                                    $Interval = $xml.task.Triggers.TimeTrigger.Repetition.Interval
                                    $Interval = $Interval -replace "PT"
                                    $Interval = $Interval -replace "H", " Hours"
                                    $Interval = $Interval -replace "D", " Days"
                                    $Interval = $Interval -replace "M", " Minutes"
                                    $Interval = " Repeat every " + $Interval
                                    if($xml.task.Triggers.TimeTrigger.Repetition.ChildNodes.name.Contains("Duration"))
                                        {
                                            $Duration = $xml.task.Triggers.CalendarTrigger.Repetition.Duration
                                            $Duration = $Duration -replace "PT"
                                            $Duration = $Duration -replace "H", " Hours"
                                            $Duration = $Duration -replace "D", " Days"
                                            $Duration = $Duration -replace "M", " Minutes"
                                            $Duration = "for " + $Duration
                                        }
                                        Else 
                                        {
                                            $Duration="Indefinitely"
                                        }
                                }
                                $trigger = "At " + $StartBoundary[1] + " on " + $StartBoundary[0]
                                if($interval)
                                    {
                                        $trigger += " - After triggered repeat every " + $Interval + " for a duration of " + $Duration
                                    }
                        }
                    else 
                        {
                            $trigger=$triggerType
                        }
                    $object2 = New-Object system.object
                    $object2 | add-member -type NoteProperty -name Name -Value $task.name
                    #$object2 | add-member -type NoteProperty -name Path -Value $task.path
                    #$object2 | add-member -type NoteProperty -name LastRunTime -Value $task.lastruntime
                    #$object2 | add-member -type NoteProperty -name LastTaskResult -Value $task.LastTaskResult
                    #$object2 | add-member -type NoteProperty -name State -Value $task.state
                    $object2 | add-member -type NoteProperty -name RunAs -Value ($task | where {$_.runas -ne $null} | foreach-object {$objSID2 = New-Object System.Security.Principal.SecurityIdentifier ($_.runas)
                    $objUser2 = $objSID.Translate( [System.Security.Principal.NTAccount])
                    $objuser2.value})
                    $object2 | add-member -type NoteProperty -name Command -Value $task.command
                    $object2 | add-member -type NoteProperty -name Arguments -Value $task.arguments
                    $object2 | add-member -type NoteProperty -name Triggers -Value $trigger
                    $object2 | add-member -type NoteProperty -name Server -Value $task.server
                    #$object2 | add-member -type NoteProperty -name TimesFailed -Value $null
                    #$object2 | add-member -type NoteProperty -name MsgSent -Value $null
                    if($object2.name -notlike "Adobe*" -and $object2.name -notlike "Optimize Start Menu*" -and $object2.name -notlike "GoogleUpdate*")
                    {
                        $taskreport += $object2
                    }
                }
            }
        }
    }
}
 #Check each task in the array for failure
    foreach($task in $taskarray)
    {
 
 #Check to see if failed ignoring never run and currently running
        if(($task.LastTaskResult -ne 0) -and ($task.LastTaskResult -ne 267011) -and ($task.LastTaskResult -ne 267009) -and ($task.State -ne 4))
        {
            $failedTasks += $task
        }
    }


 #If there are failed tasks update previous failed task report and create message based on threshold and whether a message has already been sent
 if($failedTasks.count -gt 0)
 {
 #Generate an alert message to send
    foreach($task in $failedTasks)
    {
        if($previousFailedTasks.name -contains $task.name)
        {
            $index = [array]::indexof($previousFailedTasks.name,$task.name)
            if($task.LastRunTime -ne $previousFailedTasks[$index].LastRunTime)
            {
                $task.TimesFailed = ($previousFailedTasks[$index].TimesFailed -as [int]) + 1
                $task.MsgSent = $previousFailedTasks[$index].MsgSent
                $task.ShutdownStateCount = $previousFailedTasks[$index].ShutdownStateCount
            }
            else
            {
                $task.TimesFailed = $previousFailedTasks[$index].TimesFailed
                $task.MsgSent = $previousFailedTasks[$index].MsgSent
                $task.ShutdownStateCount = $previousFailedTasks[$index].ShutdownStateCount
            }
        }
        else
        {
            $task.TimesFailed = "1"
        }
        if($task.LastTaskResult -eq -2147023781)
        {
            $task.ShutdownStateCount += 1
        }
        if(($task.ShutdownStateCount % 6) -eq 1)
        {
            $shutdownMsg += $task.name + " " + $task.path + " on " + $task.server + " is in System Shutdown in Progress state." + "<br>`n"
        }
        if($threshold.taskname -contains $task.name)
        {
            $thresholdindex = [array]::indexof($threshold.taskname,$task.name)
            if($task.TimesFailed -ge $threshold[$thresholdindex].Threshold)
            {
                if($task.MsgSent -ne $true)
                {
                    $msg += $task.name + " " + $task.path +" failed on " + $task.server + " at " + $task.LastRunTime + "<br>`n"
                    $task.MsgSent = $true
                }
            }
            Else
            {
                $task.msgsent = $false
            }
        }
        Elseif($task.MsgSent -ne $true)
        {
            $msg += $task.name + " " + $task.path +" failed on " + $task.server + " at " + $task.LastRunTime + "<br>`n"
            $task.msgsent = $true
        }
    }
    
 }
 #Create and save report of failed tasks
 $failedTasks | Export-Csv "c:\scripts\ScheduledTaskCheck\failedtaskreport.csv" -NoTypeInformation
 
 #Create and save report of tasks
 $taskreport | Export-CSV "c:\scripts\ScheduledTaskCheck\taskreport.csv" -NoTypeInformation
 
 #Send message if there is one to send
 if($msg -ne $null)
 {
    send-mailmessage -From "email@domain.com" -to @("email1@domain.com", "email2@domain.com") -subject "Scheduled Task(s) Failed" -BodyAsHtml $msg -smtpserver "smtpserver.domain.com"
 }
 if($shutdownMsg -ne $null)
 {
    send-mailmessage -From "email@domain.com" -to @("email1@domain.com", "email2@domain.com") -subject "Scheduled Task(s) in System Shutdown in Progress State" -BodyAsHtml $msg -smtpserver "smtpserver.domain.com"
 }
