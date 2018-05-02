    <#
.SYNOPSIS
  Quickly update the owner of recurrences in BusinessObjects from one user to another.
.DESCRIPTION
  Script will prompt for the username of the current owner and who to switch ownership over to.
.INPUTS
  Old owner username & new owner username
.OUTPUTS
  None
.NOTES
  Version:        1.0
  Author:         Cameron Ezell
  Creation Date:  2/22/2017
  Purpose/Change: Initial script development
  
.EXAMPLE
  .\change_owner.ps1
#>
    
    [System.Collections.ArrayList] $reportArray = @()
    $Username = 'Administrator'
    $Password = 'AdminPassword'
    [string]$CMS = "cms_server"
    $objSessionMgr = New-Object -com ("CrystalEnterprise.SessionMgr")            
    $objEnterpriseSession = $objSessionMgr.Logon($Username,$Password,$CMS,"Enterprise")                     
    $objInfoStore = $objEnterpriseSession.Service("","InfoStore")
    $ErrorActionPreference = 'silentlycontinue'

    
    
   
    
    
        function ChangeOwner
    {
        param($infoObject, $newOwner, $newOwnerID)
        
        $reportName = $infoObject.Title
        
        #Comment the block below for testing purposes as these make the changes to the reports
        $infoObject.Properties["SI_OWNER"].Value = "$newOwner"
        $infoObject.Properties["SI_DOC_SENDER"].Value = "$newOwner"
        $infoObject.SchedulingInfo.Properties["SI_SUBMITTER"].Value = "$newOwner"
        $infoObject.Properties["SI_OWNERID"].Value = "$newOwnerID"
        $infoObject.SchedulingInfo.Properties["SI_SUBMITTERID"].Value = "$newOwnerID"
        
        #commit the changes to the report object
        $infoObject.Save()
        

        Write-Host("Changed owner of $reportName")
        
    }
    
    #Read the selection for the Enterprise account being used to reschedule instances
    $selectedOwner = Read-Host -Prompt "Which username would you like to transfer ownership of reports from? (ex. OriginalUser)"

    #Query for all recurring instances that are owned by $selectedOwner
    $infoObjectsQuery = "Select * From CI_INFOOBJECTS WHERE SI_RECURRING = '1' AND SI_OWNER = `'$selectedOwner`'"
    $infoObjects = $objInfoStore.Query($InfoObjectsQuery);

    #Check to make sure there are any reports owned by this user. If not, exit.
    If($infoObjects.Count -gt 0)
    {
        #Put each report into an array for displaying the titles and count
        foreach($infoObject in $infoObjects)
        {
            [void]$reportArray.Add($infoObject.Title)
        }
        $reportCount = $infoObjects.Count

        Write-Host("Reports owned by $selectedOwner`:")

        #This creates a numbered list of reports that will be modified
        $i = 1
        foreach($report in $reportArray)
        {
            Write-Host("$i. $report")
            $i = $i + 1
        }


        Write-Host("$reportCount report(s) will be transferred to a new owner.") -BackgroundColor Yellow -ForegroundColor Black

        
        #Receive input for the new AD user that will become owner of the recurring instance
        $newOwner = Read-Host -Prompt "What's the user ID you'd like to transfer ownership to? (ex. NewUser)"

        #Query for reports owned by the AD user. We need this to find the numerical User ID since the CMS creates this for all users.
        $ownerQuery = "SELECT * FROM CI_INFOOBJECTS WHERE SI_OWNER = `'$newOwner`'"
        $ownerObjects = $objInfoStore.Query($ownerQuery);

        #I'm not sure if there's a way to get the numerical user ID if they don't own anything we can query on. We can look into this if necessary.
        If($ownerObjects.Count -gt 0)
        {
            $newOwnerID = $ownerObjects[1].Properties["SI_OWNERID"].Value
        }
        #If we can't find any reports owned by this user, we're going to exit.
        else
        {
            Write-Host("This owner wasn't found as owning any objects, so no owner ID could be collected. You will need to reach out to this customer and have them reschedule their instance(s).")
            exit
        }
        
        #Confirm the change
        Write-Host("Confirm: do you want to transfer ownership of $reportCount recurring instances from $selectedOwner to $newOwner`?") -BackgroundColor Yellow -ForegroundColor Red
        $confirmation = Read-Host -Prompt ">"


        #If they typed yes or Yes we'll go ahead and call the ChangeOwner function. Otherwise, exit.
        if($confirmation -eq "yes" -or $confirmation -eq "Yes" -or $confirmation -eq "y")
        {
            #Iterate through each object owned by our account
            foreach($infoObject in $infoObjects)
            {
                ChangeOwner -infoObject $infoObject -newOwner $newOwner -newOwnerID $newOwnerID
            }
        }
     else
        {
        Write-Host("Exiting script. Bye.")
        exit
        }
    }


    else {
           Write-Host "No recurring instances found with this username. Exiting..."
           exit
        }
