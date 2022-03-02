# master directory of all .rdl files to be handled
$masterDir = 'C:\users\paul.davis\source\repos\SSRS01'

#import the Master Report Migration csv file 
$file = 'C:\users\paul.davis\source\repos\Report Migration\Automation\Report Manager Master List.csv'

$reports = Import-Csv -Path $file

# initialize an empy array to store the objects for Retire and Migrate reports
$OldReports = @()
$LegacyReports = @()
$Retire  = @()

#split the csv into three objects - retire, Migrate Old, and Migrate Legacy
foreach($report in $reports){
   
        Switch($report.Action)
        {
            'Retire' {
                $item = [PSCustomObject]@{
                    Name = $report.'Report Name '
                    Path = $report.Path
                }
            # add the object to the Retire array
            $Retire += $item
            }
             'Migrate' {
                Switch($report.'Report Manager Folder' -eq 'Legacy Data Reports' ) 
                {
                    True {
                    $item = [PSCustomObject]@{
                        Name = $report.'Report Name '
                        Path = $report.Path
                        }
                    $LegacyReports += $item
                    }
                    False {
                    $item = [PSCustomObject]@{
                        Name = $report.'Report Name '
                        Path = $report.Path
                        }
                    $OldReports += $item
                    }
                }
            }
        }
    }

# define where to output the Old reports
$outOld =  'C:\users\paul.davis\source\repos\Report Migration\Automation\Migrate\Old Reports'

#loop through the migrate Old files and move them to the migrate dir
foreach($report in $OldReports){
    
    $source = "$($masterDir)$($report.Path).rdl" 
    $destination = "$($outOld)$($report.Path).rdl"

    # Create a placeholder file so the dir structure is built
    New-Item -ItemType File -Path $destination -Force

    #copy the actual rdl to the new destination
    Copy-Item -Path $source -Destination $destination

}

# define where to output the csv with the Migrate reports
$outLegacy =  'C:\users\paul.davis\source\repos\Report Migration\Automation\Migrate\Legacy Reports'

#loop through the migrate Old files and move them to the migrate dir
foreach($report in $LegacyReports){
    
    $source = "$($masterDir)$($report.Path).rdl" 
    $destination = "$($outLegacy)$($report.Path).rdl"

    # Create a placeholder file so the dir structure is built
    New-Item -ItemType File -Path $destination -Force

    #copy the actual rdl to the new destination
    Copy-Item -Path $source -Destination $destination

}
    

#define where to output the csv with the Retire Reports
$outRetire = 'C:\users\paul.davis\source\repos\Report Migration\Automation\Retire'

#loop through the retire files and move them to the retire dir
foreach($report in $Retire){
    
    $source = "$($masterDir)$($report.Path).rdl" 
    $destination = "$($outRetire)$($report.Path).rdl"

    # Create a placeholder file so the dir structure is built
    New-Item -ItemType File -Path $destination -Force

    #copy the actual rdl to the new destination
    Copy-Item -Path $source -Destination $destination

}

# Output CSV files for each group
$Retire | Export-Csv -Path "$($outRetire)\Retire.csv"

$OldReports | Export-Csv -Path "$($outOld)\OldReports.csv" -Force

$LegacyReports | Export-CSV -Path "$($outLegacy)\Legacy.csv" -Force
