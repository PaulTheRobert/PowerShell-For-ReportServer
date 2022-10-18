# paul davis
# implimented at VARIS LLC 03/25/2022

# Get all rdl files off server 

# To run this script you need to have ReporetingServicesTools installed
# Install-Module -Name ReportingServicesTools

# Where do you want the rdls downloaded to
$downloadFolder = 'C:\Users\pdavis\Desktop\SSRS ADMINISTATION\All RDLs\VAR-RPT3\'


# $serverUri = 'http://var-rpt1/ReportServer/ReportService2010.asmx?wsdl'
$serverUri = 'http://var-rpt3/ReportServer/ReportService2010.asmx?wsdl'

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