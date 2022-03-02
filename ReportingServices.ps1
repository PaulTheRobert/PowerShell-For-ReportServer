
$serverUri = 'http://dcids-ssrs01/ReportServer'
$reportUri = 'http://dcids-ssrs01/ReportServer'

$item = '/Family After Care/Donor Family Contacts'

# get the folders from the report server
$Folders = Get-RsCatalogItems -RsFolder '/' -ReportServerUri $uri

# get only the names from the folder object 
# foreach ($folder in $folders) {Write-Host $folder.Name}


$reports = foreach($folder in $Folders){Get-RsCatalogItems -RsFolder "/$($folder.Name)" -ReportServerUri $uri}

$reports.count