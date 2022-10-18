# created by: Paul Davis
# Implimented at VARIS LLC on 03/25/2022
<#
    This script should parse the data sources out of a collection of RDLs and save the data
    in a csv report. 

    TODO
      [] get the subreport nodes
#>

# Where do you want this report saved
$exportPath ='C:\Users\pdavis\Desktop\SSRS ADMINISTATION\SubReportsReport_VAR-RPT3.csv'

# where do all of the rdl files live that you want to parse
$path = 'C:\Users\pdavis\Desktop\SSRS ADMINISTATION\ALL RDLs\VAR-RPT3\'

#loop through all the items files in this directory and sub directories 
$items = Get-ChildItem -Path $path -Recurse -File

# array to store data before output
$Collection = @()

# loop through each rdl in the folder
foreach($item in $items){
    <#
        Desired output
            ReportItemId
            ReportPath
            ReportName
              TablixName
              TextBoxName
              SubReportName
              SubReportParameters

    #>
    
    # get the xml from the rdl
    [xml]$rdl = Get-Content $item.FullName
        
    # Loop through each tablix
    foreach($tablix in $rdl.Report.ReportSections.ReportSection.Body.ReportItems.Tablix){
        #get the tablix name
        $tablixName = $tablix.Name

        #loop through the rows
        foreach($row in $tablix.TablixBody.TablixRows){
            #loop through the cells
            foreach($cell in $row.TablixRow){
                #loop through the tablix cells
                foreach($tablixCells in $cell.TablixCells){
                    foreach($tablixCell in $tablixCells.TablixCell){
                        #loop through the cell contents
                        foreach($cellContents in $tablixCell.CellContents){
                            #loop through the text boxes
                            foreach($textbox in $cellContents.Textbox){
                                # loop throug actions
                                if($textbox.ActionInfo -ne $null){
                                    foreach($action in $textbox.ActionInfo.Actions){
                                        $Drillthrough = $action.action.Drillthrough
                                        $DrillthroughReportName = $Drillthrough.ReportName
                                        $Parameters = ''
                                        foreach($parameter in $Drillthrough.Parameters.Parameter){
                                            $Parameters = $parameters + $parameter.Name +': ' + $parameter.value +', '
                                        }
                                        
                                        write-host $DrillthroughReportName
                                        # Store the information from this run into the array  
                                        $Output = [PSCustomObject]@{
                                            ReportId    = $rdl.Report.ReportID
                                            Folder      = $item.DirectoryName
                                            Name        = $item.Name
                                            Tablix      = $tablixName
                                            TextBox     = $textbox.Name
                                            DrillThroughReportName = $Drillthrough.ReportName
                                            Parameters  = $Parameters
                                        }
                                        $Collection += $Output
                                    }                                        
                                }

                            }
                        }
                    }
                }
            }
        }            
    }
}

$Collection | Export-Csv -Path $exportPath -Force
    