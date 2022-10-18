#requires -version 5.0
function Get-DataDrivenSubscriptionProperties 
{
  param([object] $Subscription,
   [object]$ssrsproxy)
    $ssrsobject = [SSRSObject]::New()
    $sid = $Subscription.SubscriptionID
    $ddextensionSettings = $ddDataRetrievalPlan = $ddDescription = $ddactive = $ddstatus = $ddeventtype = $ddmatchdata = $ddparameters = $Null
    $ddOwner = $ssrsproxy.GetDataDrivenSubscriptionProperties($sid,[ref]$ddextensionSettings,[ref]$ddDataRetrievalPlan`
    ,[ref]$ddDescription,[ref]$ddactive,[ref]$ddstatus,[ref]$ddeventtype,[ref]$ddmatchdata,[ref]$ddparameters)
      
      $ssrsobject.subscription = $Subscription
      $ssrsobject.Owner = $ddOwner
      $ssrsobject.ExtensionSettings = $ddextensionSettings
      $ssrsobject.Description = $ddDescription
      $ssrsobject.DataRetrievalPlan = $ddDataRetrievalPlan
      $ssrsobject.Active = $ddactive
      $ssrsobject.Status = $ddstatus
      $ssrsobject.EventType = $ddeventtype
      $ssrsobject.MatchData = $ddmatchdata
      $ssrsobject.Parameters = $ddparameters
      $ssrsobject
   <# [PSCustomObject]@{
        'Owner' = $ddOwner
        'extensionSettings' = $ddextensionSettings
        'DataRetrievalPlan' = $ddDataRetrievalPlan
        'Description' = $ddDescription
        'active' = $ddactive
        'status' =$ddstatus
        'eventtype' =$ddeventtype 
        'matchdata' = $ddmatchdata
        'parameters' = $ddparameters
      } #>
}

function Get-SubscriptionProperties
{
  param([string]$Subscription,
  [object]$ssrsproxy)
  $subextensionSettings = $subDataRetrievalPlan = $subDescription = $subactive = $substatus = $subeventtype = $submatchdata = $subparameters = $Null
  
  $subOwner = $ssrsproxy.GetSubscriptionProperties($subscription.SubscriptionID,[ref]$subextensionSettings,[ref]$subDescription,[ref]$subactive,[ref]$substatus,[ref]$subeventtype,[ref]$submatchdata,[ref]$subparameters)
      $ssrsobject = [SSRSObject]::New()
      $ssrsobject.subscription = $Subscription
      $ssrsobject.Owner = $subOwner
      $ssrsobject.ExtensionSettings = $subextensionSettings
      $ssrsobject.Description = $subDescription
      $ssrsobject.DataRetrievalPlan = $subDataRetrievalPlan
      $ssrsobject.Active = $subactive
      $ssrsobject.Status = $substatus
      $ssrsobject.EventType = $subeventtype
      $ssrsobject.MatchData = $submatchdata
      $ssrsobject.Parameters = $subparameters
      <#
      [PSCustomObject]@{
        'Owner' = $subOwner
        'extensionSettings' = $subextensionSettings
        'Description' = $subDescription
        'active' = $subactive
        'status' =$substatus
        'eventtype' =$subeventtype 
        'matchdata' = $submatchdata
        'parameters' = $subparameters
      }
    #>
}

function Get-Subscriptions
{
  #Returns a nested object with each 
  param([object]$ssrsproxy, [string]$site, [switch]$DataDriven)
  #write-verbose 'Path to where the reports are must be specified to get the subscriptions you want.. Root (/) does not seem to get everything'
  $items = $ssrsproxy.ListChildren($site,$true) | Where-Object{$_.typename -eq 'report'}
  $subprops = $ddProps= @()
  foreach($item in $items)
  {
    $subs = $ssrsproxy.ListSubscriptions($item.path)
    write-verbose "found $($subs.count) subscriptions for $($item.Name)"
    
    if($subs)
    {
      foreach($sub in $subs)
      {
        
        if($sub.isdatadriven -eq 'true')
        {
          $ddProps += Get-DataDrivenSubscriptionProperties -subscription $sub -ssrsproxy $ssrsproxy
        }
        elseif(-not $DataDriven)
        {
          $subProps +=  Get-SubscriptionProperties -subscriptionid $sub -ssrsproxy $ssrsproxy
        }

      }
    }
  }
  if($DataDriven)
  {$ddProps}
  else {
    $subprops
  }
}

function New-XMLSubscriptionfile
{
    [CmdletBinding()]
    [Alias()]
    param([psobject]$subscriptionObject, [string]$path)
    if(test-path $path -PathType Leaf)
    {
      $path = split-path $path

    }
    if(-not(test-path $path))
    {
      mkdir $path
    }
    foreach($sub in $subscriptionObject)
    {
      $reportName = (($sub.subscription.report).split('.'))[0]
      $filename = "$path\$reportName.xml"
      $sub | Export-Clixml -Depth 100  -path $filename
    }
}

function New-JsonSubscriptionFile
{
    [CmdletBinding()]
    [Alias()]
    param([psobject]$subscriptionObject, [string]$path)
    if(test-path $path -PathType Leaf)
    {
      $path = split-path $path

    }
    if(-not(test-path $path))
    {
      mkdir $path
    }
    foreach($sub in $subscriptionObject)
    {
      $reportName = (($sub.subscription.report).split('.'))[0]
      $filename = "$path\$reportName.json"
      $sub | convertto-json -Depth 100  | out-file $filename
    }
}

Function Export-DataDrivenSubscriptions
{
    param([object]$ssrsproxy, [string]$site, [string]$path, [switch]$json)
    $subs = get-subscriptions -ssrsproxy $ssrsproxy -site $site -DataDriven
    if($json)
    {New-JsonSubscriptionFile -subscriptionObject $subs -path $path}
    else
    {New-XMLSubscriptionfile -subscriptionObject $subs -path $Path}

}

class SSRSObject
{
    [SSRSProxy.Subscription[]]$subscription
    [string]$Owner
    [SSRSProxy.ExtensionSettings]$ExtensionSettings
    [SSRSProxy.DataRetrievalPlan]$DataRetrievalPlan
    [string]$Description
    [SSRSProxy.ActiveState]$Active
    [string]$Status
    [string]$EventType
    [string]$MatchData
    [SSRSProxy.ParameterValueOrFieldReference[]]$Parameters

}