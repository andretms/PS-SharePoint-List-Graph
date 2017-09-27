## Get SharePoint List Items via the Microsoft Graph API
## By Andre Teixeira
## For Instructions: https://github.com/andretms/PS-SharePoint-List-Graph

##Please replace the three values below
$appId = "YourAppIdHere" # <-- Get it from https://apps.dev.microsoft.com/portal/register-app
$SiteRelativeUrl = "YourSiteRelativeUrlHere" # <-- Example: '/teams/MySharePointSite'
$ListName = "YourListNameHere" # <-- Type the name of your SharePoint list here

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

	"Obtaining access token for scopes: [$([string]::Join(`", `", $scopes))]" | Out-Host

	#Try to acquire the token silently first
	$authResultTask = $publicClientApp.AcquireTokenSilentAsync($scopes, ($publicClientApp.Users | Select-Object -index 0))
	
	if ($authResultTask.Exception -ne $null)
	{
		# If it fails - perhaps because the user has never signed-in before, 
		# check if the exception is MsalUiRequiredException - which indicates that a UI interaction is needed
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
		"Your Site Collection Root is: [$($UserSiteRoot.webUrl)]." | Out-Host
		
		#Get the user site
		$UserSite = Invoke-RestMethod -Uri $GetSitesEndpoint.Replace("{hostname}", $UserSiteRoot.siteCollection[0].hostname).Replace("{server-relative-path}",$SiteRelativeUrl) -Headers $HttpAuthHeader
		"Your Site Url is [$($UserSite.webUrl)]" | Out-Host
		
		#Now enumerate all site lists
		$SiteLists =  Invoke-RestMethod -Uri $GetListsEndpoint.Replace("{site-id}", $UserSite.id) -Headers $HttpAuthHeader
		"Site [$($UserSite.webUrl)] has $($SiteLists.value.Count) lists" | Out-Host
		
		
		#Find the list that correspond to the indicated
		$UserList = $SiteLists.value | Where-Object {$_.Name -eq $ListName}
		"The URL for your list is [$($UserList.webUrl)]" | Out-Host
		
		#Enumerate the List items
		$UserListItems =  Invoke-RestMethod -Uri $GetListItemsEndpoint.Replace("{site-id}", $UserSite.id).Replace("{list-id}", $UserList.id) -Headers $HttpAuthHeader
		
		#Finally, enumerate the items. I am only interested to the 'fields' property
		$UserListItemsFields =  $UserListItems.value | Select-Object -ExpandProperty fields
		"List [$($UserList.displayName)] has $($UserListItemsFields.Count) items" 
				
		"You are done. Below are the first 5 items for [$ListName] list:"
		
		$UserListItemsFields | Select-Object -First 5 | FT | Out-Host
	}
}