# Where do you want this report saved
$exportPath = 'C:\Users\paul.davis\source\repos\reports.csv'


# Connect to SSRS 2017 on DCIDS-SSRS02
$serverUri = 'http://dcids-ssrs02/ReportServer'

$ssrsProxy = New-RsWebServiceProxy -ReportServerUri $serverUri

$ssrsReports = $ssrsProxy.ListChildren("/", $true) | Where-Object {$_.TypeName -eq "Report"}

#reports in test
#$reportsTest = $ssrsReports | where-object {$_.Path -LIKE '*test*'}
$reportsTest = $ssrsReports 


$reportsTest| Export-Csv -Path $exportPath -append -Force