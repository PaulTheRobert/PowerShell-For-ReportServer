
# Where do you want this report saved
$exportPath = 'C:\Users\paul.davis\desktop\DataSourcesReport_DCIDS_BI_DEV1_09212021.csv'

# where do all of the rdl files live that you want to parse
#$path = 'C:\Users\paul.davis\source\repos\SSRS01'
$path = 'C:\Users\paul.davis\desktop\DCIDS-BI-DEV1'

#loop through all the items files in this directory and sub directories 
$items = Get-ChildItem -Path $path -Recurse -File

foreach($item in $items){

    [xml]$rdl = Get-Content $item.FullName

    foreach($datasource in $rdl.Report.DataSources){
        foreach($dataset in $rdl.Report.DataSets.DataSet){
            #Write-Host "Data Source Reference: $($datasource.DataSource.DataSourceReference)"
            #Write-Host "Data Source Name: $($datasource.DataSource.Name)"
            #Write-Host "Data Set Data Source: $($dataset.Query.DataSourceName)"
            #Write-Host "Data Set Command Type: $($dataset.Query.CommandType)"
            #Write-Host "Data Set Command Text: $($dataset.Query.CommandText)"

            #Store the information from this run into the array  
            $Output = [PSCustomObject]@{
                Folder = $item.DirectoryName
                Name = $item
                DataSourceReference     = $datasource.DataSource.DataSourceReference
                DataSourceName          = $datasource.DataSource.Name
                DataSetDataSourceName   = $dataset.Query.DataSourceName
                DataSetCommandType      = $dataset.Query.CommandType
                DataSetCommandText      = $dataset.Query.CommandText
            }
                $Output | Export-Csv -Path $exportPath -append -Force
        }

    }

}

