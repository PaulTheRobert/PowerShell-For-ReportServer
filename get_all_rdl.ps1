# Get all rdl files off server 

#$downloadFolder = "C:\Users\paul.davis\source\repos\SSRS01"
$downloadFolder = "C:\Users\paul.davis\desktop\DCIDS-BI-DEV1"

#$serverUri = 'http://dcids-ssrs01/ReportServer'
$serverUri = 'http://dcids-bi-dev1/Reports'

$ssrsProxy = New-RsWebServiceProxy -ReportServerUri $serverUri

$ssrsItems = $ssrsProxy.ListChildren("/", $true) | Where-Object {$_.TypeName -eq "DataSource" -or $_.TypeName -eq "Report"}

Foreach($ssrsItem in $ssrsItems)
{
    # Determine extension for Reports and DataSources
    if ($ssrsItem.TypeName -eq "Report")
    {
        $extension = ".rdl"
    }
    else
    {
        $extension = ".rds"
    }
     
    # Write path to screen for debug purposes
    Write-Host "Downloading $($ssrsItem.Path)$($extension)";
 
    # Create download folder if it doesn't exist (concatenate: "c:\temp\ssrs\" and "/SSRSFolder/")
    $downloadFolderSub = $downloadFolder.Trim('\') + $ssrsItem.Path.Replace($ssrsItem.Name,"").Replace("/","\").Trim() 
    New-Item -ItemType Directory -Path $downloadFolderSub -Force > $null
 
    # Get SSRS file bytes in a variable
    $ssrsFile = New-Object System.Xml.XmlDocument
    [byte[]] $ssrsDefinition = $null
    $ssrsDefinition = $ssrsProxy.GetItemDefinition($ssrsItem.Path)
 
    # Download the actual bytes
    [System.IO.MemoryStream] $memoryStream = New-Object System.IO.MemoryStream(@(,$ssrsDefinition))
    $ssrsFile.Load($memoryStream)
    $fullDataSourceFileName = $downloadFolderSub + "\" + $ssrsItem.Name +  $extension;
    $ssrsFile.Save($fullDataSourceFileName);
}