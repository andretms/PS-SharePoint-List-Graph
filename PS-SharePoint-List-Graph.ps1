## Get SharePoint List Items via the Microsoft Graph API
## By Andre Teixeira
## For Instructions: https://github.com/andretms/PS-SharePoint-List-Graph


##Please replace the three values below
$appId = "YourAppIdHere" 
$SiteRelativeUrl = "YourSiteRelativeUrlHere"
$ListName = "YourListNameHere"

## Script Starts Below
$msalPath = "$($MyInvocation.MyCommand.Path)\packages\Microsoft.Identity.Client.1.1.0-preview\lib\net45\Microsoft.Identity.Client.dll"
$graphURI = "https://graph.microsoft.com"
$graphScopes = @("Sites.Read.All")

$GraphBetaBaseUrl = "https://graph.microsoft.com/beta"
$GetSiteRootEndpoint = "$($GraphBetaBaseUrl)/sites/root"
$GetSitesEndpoint = "$($GraphBetaBaseUrl)/sites/{hostname}:/{server-relative-path}"
$GetListsEndpoint = "$($GraphBetaBaseUrl)/sites/{site-id}/lists"
$GetListItemsEndpoint = "$($GraphBetaBaseUrl)/sites/{site-id}/lists/{list-id}/items?expand=fields"

function InitializeMsal
{
	<#
	.SYNOPSIS
	Initialize MSAL Librarty
	
	.PARAMETER appId
	The application Id obtained by registering an app under https://apps.dev.microsoft.com
	#>
	
	param ( [Parameter(Mandatory=$true)]
            $appId )
  
	if ($publicClientApp.ClientId -eq $null)
	{
       [System.Reflection.Assembly]::LoadFrom($msalPath) | Out-Null
       $publicClientApp = New-Object "Microsoft.Identity.Client.PublicClientApplication" -ArgumentList $appId 
	}
	
	return $publicClientApp
}

function Get-AccessToken
{
	<#
	.SYNOPSIS
	Use MSAL library to get an Access Token to access a resource
	
	.PARAMETER scopes
	An array of permission scopes
	#>

	param ( [Parameter(Mandatory=$true)]
            [System.Collections.Generic.List[string]]  $scopes )

	$authResultTask = $publicClientApp.AcquireTokenSilentAsync($scopes, ($publicClientApp.Users | Select-Object -index 0))
	
	if ($authResultTask.Exception -ne $null)
	{
		if ($authResultTask.Exception.InnerException -is  [Microsoft.Identity.Client.MsalUiRequiredException])
		{
			$authResultTask = $publicClientApp.AcquireTokenAsync($scopes)
		}
		else
		{
			Write-Error $authResultTask.Exception
		}
	}

	if ($authResultTask.Result)
	{
		return $authResultTask.Result.AccessToken
	}
}

#Initialize MSAL
$publicClientApp = InitializeMsal -appId $appId 

if ($publicClientApp.ClientId -ne $null)
{
	#Get the Access Token using MSAL Library and add it to the HTTP Authorization Header
	$HttpAuthHeader = @{ Authorization = ("Bearer " + (Get-AccessToken $graphScopes)) }
	
	#Get the SharePoint Site Root for the user (i.e. contoso.sharepoint.com)
	$UserSiteRoot = Invoke-RestMethod -Uri $GetSiteRootEndpoint -Headers $HttpAuthHeader
	
	if ($UserSiteRoot -ne $null)
	{
		"Your Site Collection Root is: $($UserSiteRoot.webUrl). HostName is [$($UserSiteRoot.siteCollection.hostname)]" | Out-Host
		#Get the user site
		$UserSite = Invoke-RestMethod -Uri $GetSitesEndpoint.Replace("{hostname}", $UserSiteRoot.siteCollection[0].hostname).Replace("{server-relative-path}",$SiteRelativeUrl) -Headers $HttpAuthHeader
		
		#Now enumerate all site lists
		$SiteLists =  Invoke-RestMethod -Uri $GetListsEndpoint.Replace("{site-id}", $UserSite.id) -Headers $HttpAuthHeader
		
		#Find the list that correspond to the indicated
		$UserList = $SiteLists.value | Where-Object {$_.Name -eq $ListName}
		
		#Finally, enumerate the items
		$UserListItems =  Invoke-RestMethod -Uri $GetListItemsEndpoint.Replace("{site-id}", $UserSite.id).Replace("{list-id}", $UserList.id) -Headers $HttpAuthHeader
		
		$UserListItemsFields =  $UserListItems.value | Select-Object -ExpandProperty fields
		
		"You are done. Below is the Field Values for [$ListName] List:"
		
		$UserListItemsFields | Out-Host
	}
}