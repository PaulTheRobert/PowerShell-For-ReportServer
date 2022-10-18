# SSRS User Permissions Audit
<#
    created by: paul davis
    implimented on: 04/06/2022

    This scripts gets a list of all folders from the Report Server 
    and then loops through that list checking the security

    It also collects a list of users on subscriptions 

    1 report is output at the end

#>
$ReportServer = "VAR-RPT1"

$serverUri = "http://$($ReportServer)/ReportServer/ReportService2010.asmx?wsdl"

$InheritParent = $true

$ssrsProxy = New-RsWebServiceProxy -ReportServerUri $serverUri
$folders = $ssrsProxy.ListChildren("/", $true) | Where-Object {$_.TypeName -eq "Folder"}

# store the permission results
$permissions = @();

#get the security on each folder
foreach($folder in $folders){

    $Policies = $ssrsProxy.GetPolicies($folder.Path, [ref] $InheritParent)

    foreach($policy in $Policies){
        
        $Path = $folder.Path;
        $GroupUserName = $policy.GroupUserName;

        # initilaze empty roles then stuff roles into one field
        $Roles = "";                        
        foreach($role in $policy.Roles){$Roles += $role.Name + ", " }
            

        $GroupUserName =  $GroupUserName -replace 'VARIS\\'
        try{$GroupMembers = Get-ADGroupMember -Identity $GroupUserName}
        catch{$GroupMembers =  Get-ADUser -Identity $GroupUserName}

        #loop through the policies and get ad users from group
        foreach($member in $GroupMembers){

            $user = Get-ADUser $member
          
            $Data = New-Object psobject -Property @{
                "Path" = $Path;
                "GroupUserName" = $GroupUserName; 
                "UserFullName" = $member.name;
                "User" = $member.SamAccountName;
                "Enabled" = $user.Enabled
                "Roles" = $Roles;
            }
            $permissions += $Data
        }       
    }

    
    
    

   
}

$permissions | Export-Csv  -Path "C:\Users\pdavis\Desktop\SSRS ADMINISTATION\Reports on Reports\Security Audit\$($ReportServer) SecurityAudit.csv"