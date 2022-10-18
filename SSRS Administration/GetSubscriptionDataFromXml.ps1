$path = 'C:\Users\pdavis\Desktop\SSRS ADMINISTATION\All Subscriptions\VAR-RPT1\'
# $path = 'C:\Users\pdavis\Desktop\SSRS ADMINISTATION\All Subscriptions\VAR-RPT3\'
$outputPath = 'C:\Users\pdavis\Desktop\SSRS ADMINISTATION\Reports on Reports\SubscriptionsData\SubscriptionsData_VAR-RPT1.csv'
# $outputPath = 'C:\Users\pdavis\Desktop\SSRS ADMINISTATION\Reports on Reports\SubscriptionsData\SubscriptionsData_VAR-RPT3.csv'
$items = Get-ChildItem -path $path -Recurse -File


foreach($item in $items){
        Clear-Variable -Name EventType, Report, Status, Active, IsDataDriven, DataRetrievalPlan, ConnectString, DataSet, Query, CommandText, CommandType
    
        [xml]$sub = Get-Content $item.FullName

        
        #use .'#text' to extract the value
        try{$EventType  = $sub.Objs.Obj.MS.S | Where-Object {$_.N -Like "EventType"}}catch{Write-Host "miss"}
        try{$Report     = $sub.Objs.Obj.MS.S | Where-Object {$_.N -Like "Report"}}catch{Write-Host "miss"}
        try{$Status     = $sub.Objs.Obj.MS.S | Where-Object {$_.N -Like "Status"}}catch{Write-Host "miss"}
        try{$Path       = $sub.Objs.Obj.MS.S | Where-Object {$_.N -Like "Path"}}catch{Write-Host "miss"}
        try{$Active     = $sub.Objs.Obj.MS.Obj.RefId[0]}catch{Write-Host "miss"}
        try{$IsDataDriven = $sub.Objs.Obj.MS.B | Where-Object {$_.N -Like "IsDataDriven"}}catch{Write-Host "miss"}
        try{$subscriptionType =  $sub.Objs.Obj.MS.Obj.Props.S | Where-Object {$_.N -Like "Extension"}}catch{Write-Host "miss"}

        try{$DataRetrievalPlan = $sub.Objs.Obj.MS.Obj | Where-Object {$_.N -Like "DataRetrievalPlan"}}catch{Write-Host "miss"}
        try{$ConnectString =    $DataRetrievalPlan.Props.Obj[0].Props.S | Where-Object {$_.N -Like "ConnectString"}}catch{Write-Host "miss"}

        try{$DataSet = $DataRetrievalPlan.Props.Obj | Where-Object {$_.N -Like "DataSet"}}catch{Write-Host "miss"}
        try{$Query  =  $DataSet.Props.Obj | Where-Object {$_.N -Like "Query"}}catch{Write-Host "miss"}

        try{$CommandType = $Query.Props.S | Where-Object {$_.N -Like "CommandType"}}catch{Write-Host "miss"}
        try{$CommandText = $Query.Props.S | Where-Object {$_.N -Like "CommandText"}}catch{Write-Host "miss"}
        
        #Store the information from this run into the array  
        $Output = [PSCustomObject]@{
            EventType   =   $EventType.'#text'
            Report      =   $Report.'#text'
            Status      =   $Status.'#text'
            Active      =   $Active
            Path        =   $Path.'#text'
            IsDataDriven =  $IsDataDriven.'#text'
            Extension       = $subscriptionType.'#text'
            ConnectionString = $ConnectString.'#text'
            CommandType     = $CommandType.'#text'
            CommandText     = $CommandText.'#text'
        }
        $Output | Export-Csv -Path $outputPath -Force -Append
    }

