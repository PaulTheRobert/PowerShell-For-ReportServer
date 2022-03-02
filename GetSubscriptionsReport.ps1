$reportsMasterListPath = 'C:\Users\paul.davis\source\repos\Report Migration\Automation\Report Manager Master List.csv'
$reportList = Import-Csv -Path $reportsMasterListPath

$ReportServer = Read-Host -Prompt "which server? NEW or OLD"

$WhichReports=''

Switch($ReportServer){
    "New"{
        # NEW
        $serverUri = 'http://dcids-bi-prod1/Reports'
    }
    "Old"{
        # OLD
        $serverUri = 'http://dcids-ssrs01/ReportServer'

        $WhichReports = Read-Host -Prompt "do you want ALL, MIGRATE, or RETIRE reports?"
    }
}

$ssrsProxy = New-RsWebServiceProxy -ReportServerUri $serverUri

Switch($WhichReports){
    "ALL"{
        $list = $ssrsProxy.ListSubscriptions("/")
    }
    "MIGRATE"{
        $list = $reportList| Where-Object -Property Action -EQ 'Migrate'
    }
    "RETIRE"{
        $list = $reportList| Where-Object -Property Action -EQ 'Retire'
    }
    DEFAULT{
       # $list = $ssrsProxy.ListSubscriptions("/")
       $list = $ssrsProxy.ListChildren("/", $true) | Where-Object {$_.TypeName -eq "Report"}
    }
}

$outputArray = @()
$parameterArray = @()

foreach($l in $list){
    $subs = Get-RsSubscription -Proxy $ssrsProxy -RsItem $l.Path

    foreach($sub in $subs){ 



        $id = $l.SubscriptionId
        [xml]$Schedule = $sub.MatchData
        
        $Recurrence = $Schedule.ScheduleDefinition.LastChild.Name
        $StartDateTime = $Schedule.ScheduleDefinition.StartDateTime.InnerText

        try{$EndDateTime = $Schedule.ScheduleDefinition.EndDateTime.InnerText}catch{$EndDateTime = ''}

        $weeksInterval = $Schedule.ScheduleDefinition.WeeklyRecurrence.WeeksInterval
        $Days = $Schedule.ScheduleDefinition.WeeklyRecurrence.DaysOfWeek

        foreach($Day in $Days){
            try{ $Monday    = $Day.Monday   } catch {$Monday    = $False}
            try{ $Tuesday   = $Day.Tuesday  } catch {$Tuesday   = $False}
            try{ $Wednesday = $Day.Wednesday} catch {$Wednesday = $False}
            try{ $Thursday  = $Day.Thursday } catch {$Thursday  = $False}
            try{ $Friday    = $Day.Friday   } catch {$Friday    = $False}
            try{ $Saturday  = $Day.Saturday } catch {$Saturday  = $False}
            try{ $Sunday    = $Day.Sunday   } catch {$Sunday    = $False}
        }
        
        $Subject       = $sub.DeliverySettings.ParameterValues | Where-Object {$_.Name -eq "Subject"}
        $IncludeReport = $sub.DeliverySettings.ParameterValues | Where-Object {$_.Name -eq "IncludeReport"}
        $RenderFormat  = $sub.DeliverySettings.ParameterValues | Where-Object {$_.Name -eq "RenderFormat"}
        $IncludeLink   = $sub.DeliverySettings.ParameterValues | Where-Object {$_.Name -eq "IncludeLink"}
        $Priority      = $sub.DeliverySettings.ParameterValues | Where-Object {$_.Name -eq "ToPriority"}

        $To            = $sub.DeliverySettings.ParameterValues | Where-Object {$_.Name -eq "To"}
        $CC            = $sub.DeliverySettings.ParameterValues | Where-Object {$_.Name -eq "CC"}
        $BCC           = $sub.DeliverySettings.ParameterValues | Where-Object {$_.Name -eq "BCC"}

        $Params = ''

        foreach($v in $sub.Values){
            $Params += "$($v.Name):$($v.Value); "
        }
            
        $ReportsSubscriptionParam = [PSCustomObject]@{
            ID            = $sub.SubscriptionID
            Path          = $l.Path
            Report        = $sub.Report
            Desc          = $sub.Description
            LastExecuted  = $sub.LastExecuted
            Status        = $sub.Status
            Params        = $Params
            To            = $To.Value
            CC            = $CC.Value
            BCC           = $BCC.Value
            StartDate     = $StartDateTime.Substring(0,10)
            StartTime     = $StartDateTime.Substring(12,17)
            Monday        = $Monday
            Tuesday       = $Tuesday
            Wednesday     = $Wednesday
            Thursday      = $Thursday
            Friday        = $Friday
            Saturday      = $Saturday
            Sunday        = $Sunday
            
            Recurrence    = $Recurrence

            EndDateTime   = $EndDateTime
            WeeksInterval = $weeksInterval

            Subject       = $Subject.Value
            IncludeReport = $IncludeReport.Value
            RenderFormat  = $RenderFormat.Value
            IncludeLink   = $IncludeLink.Value
            Priority      = $Priority.Value
        }
        $parameterArray += $ReportsSubscriptionParam 
    }
}


# CHECK FOR ACTIVE USERS
$Users = @()

foreach($item in $parameterArray){
    # take all the values in the To, remove spaces, and split into an array
    # TO
    try{
        $To = $item.To.ToString().Split(';') -replace '\s'
    }catch{}
    
    foreach($T in $To){
        try{
            $Users += $T
        }catch{}
    }
    # CC
    try{
        $CC = $item.CC.Split(';') -replace '\s'
    }catch{}
        foreach($C in $CC){
            try{
                $Users += $C
            }
            catch{}
        }
    #BCC
    try{
        $BCC = $item.BCC.Split(';') -replace '\s'
    }catch{}
        foreach($B in $BCC){
            try{
                $Users += $B
            }
            catch{}
        }
}

#get rid of blank and duplicate users
$UniqueUsers = $Users | Where-Object {$_ -ne ''} | Select -Unique

$activeUsersArray = @()
# loop through the list of unique users and check AD 
foreach($user in $UniqueUsers){
    switch($user.Substring(0,4)){
        "dist"{
            $activeUsers = [PSCustomObject]@{
            User = $user
            Active = "Distro"
            }
        }
        Default{
            try{
                $identity = $user -replace '@dcids.org'
                $value = Get-ADUser $identity | Select-Object {$_.Enabled}
                #$activeUsers.Add("$($user), $($value.'$_.Enabled')")
                $activeUsers = [PSCustomObject]@{
                    User = $user
                    Active = $value.'$_.Enabled'
                }
            }
            catch{
                #this will set active to fales if not record is found in AD for the user
                #$activeUsers += ("$($user), False")
                $activeUsers = [PSCustomObject]@{
                    User = $user
                    Active = "False"
                }
            }
        }
    }
    $activeUsersArray += $activeUsers
}

Switch($ReportServer){
    "New"{
        $fileNameSubReport = "C:\Users\paul.davis\desktop\PowerBI ALL Subscriptions.csv"
        $parameterArray  | Export-Csv -Path $fileNameSubReport -Force -NoTypeInformation -Append

        $fileNameUserReport = "C:\Users\paul.davis\source\repos\Powershell\PowerBI Report Server\Reports\PowerBI ALL Subscriptions Active Users.csv"
        $activeUsersArray | Export-Csv $fileNameUserReport -Force -NoTypeInformation -Append
    }
    "Old"{
        $fileNameSubReport = "C:\Users\paul.davis\source\repos\Powershell\SSRS01 Report Server\Reports\SSRS01 $($WhichReports) Subscriptions.csv"
        $parameterArray  | Export-Csv -Path $fileNameSubReport -Force -NoTypeInformation -Append

        $fileNameUserReport = "C:\Users\paul.davis\source\repos\Powershell\SSRS01 Report Server\Reports\SSRS01 $($WhichReports) Subscriptions Active Users.csv"
        $activeUsersArray | Export-Csv $fileNameUserReport -Force -NoTypeInformation -Append
     }
}