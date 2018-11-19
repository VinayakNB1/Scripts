Import-Module Microsoft.Online.SharePoint.PowerShell -DisableNameChecking
Connect-SPOService -Url https://xyz-admin.sharepoint.com/
Add-Type -Path "C:\\Users\\Microsoft.SharePoint.Client.dll"          
Add-Type -Path "C:\\Users\\Microsoft.SharePoint.Client.Runtime.dll"

$sSiteUrl= "https://xyz.sharepoint.com"
$sUserName="abc@xyz.onmicrosoft.com"
$sPassword=ConvertTo-SecureString "test" -AsPlainText -Force

$spoCtx = New-Object Microsoft.SharePoint.Client.ClientContext($sSiteUrl) 
$spoCredentials = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($sUserName, $sPassword)  
$spoCtx.Credentials = $spoCredentials  


$sListName= "abc";

$spoList = $spoCtx.Web.Lists.GetByTitle($sListName)
$spoCtx.Load($spoList)
$sItem="xyz"
       
Write-Host "Adding " $sItem.SPOListItem " to $sListName" 
$spoListItemCreationInformation = New-Object Microsoft.SharePoint.Client.ListItemCreationInformation 
$spoListItem=$spoList.AddItem($spoListItemCreationInformation) 
$spoListItem["Title"]="xyz"; 
$spoListItem.Update() 
$spoCtx.ExecuteQuery()             
      
