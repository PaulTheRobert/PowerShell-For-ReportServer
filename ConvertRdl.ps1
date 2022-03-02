# input old .rdl directory

# Old Reports
#$in = 'C:\Users\paul.davis\source\repos\Report Migration\TestAuto\Original'
#$in = 'C:\Users\paul.davis\source\repos\Report Migration\TestAuto\Original\Migrate\Old Reports\B2LC'
$in = 'C:\Users\paul.davis\source\repos\SSRS2019_Updated'

# ouput revised .rdl directory
#$out = 'C:\Users\paul.davis\source\repos\Report Migration\TestAuto\New'
$out = 'C:\Users\paul.davis\source\repos\SSRS2019_Migrated'

# get files from input dir
$items = Get-ChildItem -Path $in -Recurse -File

#datasource accumulator
$datasources = @()

#error accumulator
$DSErrorReports = @()

foreach($item in $items){
    Clear-Variable datasources
    $path = 'C:\Users\paul.davis\source\repos\SSRS2019_Updated\'
    $fullpath = $item.Directory.ToString()
    $dirPath = ($fullpath -replace [regex]::Escape($path), '')

    #build the dir structure with placeholders
    $destination = "$($out)\$($dirPath)\$($item.Name)"

    # Create a placeholder file so the dir structure is built
    New-Item -ItemType File -Path $destination -Force
    
    # get file contents as xml
    $rdl = [System.Xml.XmlDocument](Get-Content $item.FullName)    
   
    #$xmlns = 'http://schemas.microsoft.com/sqlserver/reporting/2016/01/reportdefinition'
    $xmlns = $rdl.Report.xmlns
    
    #set the new value of xmlns
    $rdl.Report.SetAttribute("xmlns",$xmlns)
    
    #$rd = "http://schemas.microsoft.com/SQLServer/reporting/reportdesigner"
    $rd = $rdl.Report.rd[0]
    #set the new value of rd
    $rdl.Report.SetAttribute("rd",$rd)

    #iterate over the data sources
    foreach($datasource in $rdl.Report.DataSources){
   #    write-host $datasource.DataSource.DataSourceReference
   # }
   # }
        #Store the new dataset info  
        $OldDataSource = [PSCustomObject]@{
            DataSourceName          = $datasource.DataSource.Name
            DataSourceReference     = $datasource.DataSource.DataSourceReference
            DatSourceId             = $datasource.DataSource.DataSourceID
        }

        #build the new data source object
        Switch($datasource.DataSource.DataSourceReference){      
            'DCI_REPORTING' {                    
                $Name = 'DCIReporting'
                $Catalog = 'DCIReporting'
                }
            'DCI_DataWarehouse' {
                $Name = 'iTxLegacy'
                $Catalog = 'iTxLegacy'
                }
            'DCI_Warehouse' {
                $Name = 'iTxLegacy'
                $Catalog = 'iTxLegacy'
                }
            }

            $NewDataSource = [PSCustomObject]@{
                OldName = $datasource.DataSource.Name
                Name = $Name
                Catalog = $Catalog
                SecurityType = $datasource.DataSource.SecurityType
                DataSourceId = $datasource.DataSource.DataSourceID
                }
            $datasources += $NewDataSource
        }
 <#   # remove the body node
  try{
        $body = $rdl.Report.Body
        $rdl.Report.RemoveChild($body)
      }
     catch{write-host 'No Body'}
  try{      
        $width = $rdl.Width
        $rdl.RemoveChild($width)
    }
  catch{write-host 'no-width'}
#>

    # remove the dataSourceNode 
    #$rdl.Report.DataSources.DataSource.RemoveAll()
    $datasource.RemoveAll()

    #make a node for each data source
        foreach($source in $datasources){
        #build the ds xml node
        $xmlDataSource = $rdl.CreateElement("DataSource", $datasource.NamespaceURI)
                
        # set ds name attribute
        $xmlDataSource.SetAttribute("Name", $source.Name)
        $xmlDS = $datasource.AppendChild($xmlDataSource)

        $xmlConectionProperties = $rdl.CreateElement("ConnectionProperties", $datasource.NamespaceURI)
        $xmlCP = $xmlDS.AppendChild($xmlConectionProperties)

        $xmlDataProvider = $rdl.CreateElement("DataProvider", $datasource.NamespaceURI)
        $xmlDataProvider.InnerText = "SQL"
        $xmlCP.AppendChild($xmlDataProvider)

        $xmlConnectString = $rdl.CreateElement("ConnectString", $datasource.NamespaceURI)
        $xmlConnectString.InnerText = "Data Source=DCIDS-SQL-PROD1; Initial Catalog=$($source.Catalog)"
        $xmlCP.AppendChild($xmlConnectString)

        $xmlIntegratedSecurity = $rdl.CreateElement("IntegratedSecurity", $datasource.NamespaceURI)
        $xmlIntegratedSecurity.InnerText = "true"
        $xmlCP.AppendChild($xmlIntegratedSecurity)

        $xmlSecurityType = $rdl.CreateElement("rd:SecurityType", $rd)
        $xmlSecurityType.InnerText = $source.SecurityType
        $xmlDS.AppendChild($xmlSecurityType)

        $xmlDataSourceId = $rdl.CreateElement("rd:DataSourceID", $rd)
        $xmlDataSourceId.InnerText = $source.DataSourceId
        $xmlDS.AppendChild($xmlDataSourceId)

    }


    #iterate over the data sets and change the source names as needed
    foreach($dataset in $rdl.Report.DataSets){
       try{
            $DataSourceName = $datasources|where-object {$_.OldName -eq $dataset.DataSet.Query.DataSourceName}|Select-Object -Property Name
            $dataset.DataSet.Query.DataSourceName =  $DataSourceName.Name
        }
        catch{
            write-host "error with $($item)"
            $DSErrorReports += $item
        }
    }

    # build file name
    #$outString = "$($out)\$($item)"

    #output new file
    $rdl.Save($destination)
}