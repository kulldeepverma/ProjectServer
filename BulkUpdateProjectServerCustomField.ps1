<#
Author          : Kuldeep Verma
Created Date    : 16th Jan, 2018.
Title           : Update Resource Level Custom Field Project Server and Project Server Online 
Description     : Bulk Update Project Server Resource Level Custom Field 
#>

<# Celow is CSV sample which I am using for this code. First Column is headers. #>
<#
ResourceName,ResourceEmailAddress,EmpRate
Kuldeep, kuldeep.verma@xyz.com, 10000
Rajdeep, rajdeep.sardesai@xyz.com, 20000
Jaydeep, Jaydeep.prajapat@xyz.com, 40000
Amrita Japtap, Amrita.jagtap@xyz.com, 20000
#>

#Download and install Microsoft Client Component from https://www.microsoft.com/en-in/download/details.aspx?id=42038 
#Microsoft SharePoint Client Component is mandatory to move further. 
Add-Type -Path "C:\Program Files\Common Files\microsoft shared\Web Server Extensions\16\ISAPI\Microsoft.SharePoint.Client.dll"  
Add-Type -Path "C:\Program Files\Common Files\microsoft shared\Web Server Extensions\16\ISAPI\Microsoft.ProjectServer.Client.dll"  

#replace CSV path 
$csvpath="C:\FakePath\FileName.csv" 
$siteURL = "<Site URL>"  
$loginname = "<User Name>"  
$pwd = "<Password>"  
$securePassword = ConvertTo-SecureString $pwd -AsPlainText -Force 
$creds = New-Object System.Management.Automation.PsCredential $loginname,$securePassword
$projContext = New-Object Microsoft.ProjectServer.Client.ProjectContext($siteURL)
$projContext.Credentials = $creds
$customfields = $projContext.CustomFields
$projContext.Load($customfields)
$projContext.ExecuteQuery() 

#get all custom field 
$customfields | select InternalName, Name, IsRequired

#get custom field by anem 
$customfield = $customfields  | select InternalName, Name | WHERE {$_.Name -eq "<Field Display Name>"} 


Import-Csv $csvpath | ForEach-Object {
    Try {
        Write-Host "Updateding Resource $($_.ResourceName)..." -ForegroundColor Green
        $resourceEmail = $_.ResourceEmailAddress
        $resource = $projContext.EnterpriseResources | select Id, Name, Email, UserPrincipalName | where { $_.Email -eq $resourceEmail }
        if ($resource -ne $null) {
            $res = $projContext.EnterpriseResources.GetByGuid($resource.Id)
            #update $_.Value with your field name of CSV and Change the Internal field name 
            $res["<Internal Field Name>"] = $_.EmpRate
            $projContext.EnterpriseResources.Update()
            $projContext.ExecuteQuery()
            Write-Host "Updated Resource $($_.ResourceName) successfully!" -ForegroundColor Green
        }
        else {
            Write-Host "Resource $($_.ResourceName) not found!" -ForegroundColor DarkCyan
        }
    }
    catch {
        $_ | select -expandproperty invocationinfo
        Write-Host "Error occurred while updating $($_.ResourceName)..." -ForegroundColor Red
        write-host "$($_.Exception.Message)" -foregroundcolor DarkRed
    }
}
Write-Host "Script executed successfully" -ForegroundColor Yellow

