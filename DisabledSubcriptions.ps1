$serverUri = 'http://dcids-bi-prod1/Reports'
$ssrsProxy = New-RsWebServiceProxy -ReportServerUri $serverUri

$list = $ssrsProxy.ListSubscriptions("/")

$outputArray = @()

foreach($l in $list){

    $Output = [PSCustomObject]@{
        Path         = $l.Path
        Report       = $l.Report
        Description  = $l.Description
        DisabledByUser          = $l.Active.DisabledByUser
        DisabledByUserSpecified = $l.Active.DisabledByUserSpecified
    }

    $outputArray += $Output
}

$outputArray | Where-Object {$_.DisabledByUser -eq $false}