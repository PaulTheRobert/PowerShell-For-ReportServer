# Migrate subscriptions from one server to another

$fromServer =  'http://dcids-bi-prod1/Reports'
# $toServer = 'http://dcids-bi-prod1/Reports'

# save the xml files locally
$path = 'C:\Users\paul.davis\source\repos\Subscriptions_12152021\'

$fromProxy = New-RsWebServiceProxy -ReportServerUri $fromServer
# $toProxy = New-RsWebServiceProxy -ReportServerUri $toServer

$list = $fromProxy.ListSubscriptions("/")


foreach($l in $list){
    #export the subscriptions as xml from the old server to your computer 
     foreach($sub in Get-RsSubscription -Proxy $fromProxy -RsItem $l.path){
    
         Export-RsSubscriptionXml -Path "$($path)$($sub.SubscriptionID).xml" -Subscription $sub  

        #import the subsctiption from your computer to the new server
    
        #$importSub =Import-RsSubscriptionXml -Path "$($path)$($sub.SubscriptionID).xml" -proxy $toProxy 

        #Copy-RsSubscription -Subscription $importSub -RsItem $importSub.Path -Proxy $toProxy 

        }
    }
