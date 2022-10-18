# created by: Paul Davis
# Implimented at VARIS LLC on 03/25/2022
<#
    This script should parse the data sources out of a collection of RDLs and save the data
    in a csv report. 

    TODO
      [] Get the Connection string out of the data sets so you can see which server each points to
#>

# Where do you want this report saved
# you could add a folder picker here
$exportPath ='C:\Users\pdavis\Desktop\SSRS ADMINISTATION\Reports on Reports\Report Inventory\DEV-SSRS1.csv'

# where do all of the rdl files live that you want to parse
$path = 'C:\Users\pdavis\Desktop\SSRS ADMINISTATION\Migration\Test'

#loop through all the items files in this directory and sub directories 
$items = Get-ChildItem -Path $path -Recurse -File

foreach($item in $items){

    [xml]$rdl = Get-Content $item.FullName

    foreach($datasource in $rdl.Report.DataSources){
        foreach($dataset in $rdl.Report.DataSets.DataSet){
            #Store the information from this run into the array  
            $Output = [PSCustomObject]@{
                Folder                  = $item.DirectoryName
                Name                    = $item
                DataSourceReference     = $datasource.DataSource.DataSourceReference
                DataSourceName          = $datasource.DataSource.Name
                DataSetDataSourceName   = $dataset.Query.DataSourceName
                DataSetCommandType      = $dataset.Query.CommandType
                DataSetCommandText     = $dataset.Query.CommandText.Replace("`n", '').Replace("`r", '')                              
            }
                $Output | Export-Csv -Path $exportPath -append -Force
        }
    }
}
