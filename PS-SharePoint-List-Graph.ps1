$msalPath = "C:\Temp\PSSharePointGraph\packages\Microsoft.Identity.Client.1.1.0-preview\lib\net45\Microsoft.Identity.Client.dll"

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

function Invoke-RestMethod-WithToken
{
	param ( [Parameter(Mandatory=$true)]
            [System.Uri] $uri, 
			[Parameter(Mandatory=$true)]
            [array] $scopes, 
			[Microsoft.PowerShell.Commands.WebRequestMethod] $method = [Microsoft.PowerShell.Commands.WebRequestMethod]::Get)

		"Quering [$uri]" | Out-Host
	 
		Invoke-RestMethod -Uri $uri -Headers @{ Authorization = ("Bearer " + (Get-AccessToken $scopes)) }
}

function Get-AuthHeaderForScopes
{
	param ( [Parameter(Mandatory=$true)]
            [array] $scopes )
	return  @{ Authorization = ("Bearer " + (Get-AccessToken $scopes)) } `
}


#Initialize MSAL
$publicClientApp = InitializeMsal -clientId $appId 

#Now Acquire the Token

if ($publicClientApp.ClientId -ne $null)
{
	$UserSiteRoot = Invoke-RestMethod-WithToken -Uri $GetSiteRootEndpoint -scopes $graphScopes
	if ($UserSiteRoot -ne $null)
	{
		"Your Site Collection Root is: $($UserSiteRoot.webUrl). HostName is [$($UserSiteRoot.siteCollection.hostname)]" | Out-Host
		#Get the user site
		$UserSite = Invoke-RestMethod-WithToken -Uri $GetSitesEndpoint.Replace("{hostname}", $UserSiteRoot.siteCollection[0].hostname).Replace("{server-relative-path}",$SiteRelativeUrl) -scopes $graphScopes
		
		#Now enumerate all site lists
		$SiteLists =  Invoke-RestMethod-WithToken -Uri $GetListsEndpoint.Replace("{site-id}", $UserSite.id) -scopes $graphScopes
		
		#Find the list that correspond to the indicated
		$UserList = $SiteLists.value | Where-Object {$_.Name -eq $ListName}
		
		#Finally, enumerate the items
		$UserListItems =  Invoke-RestMethod-WithToken -Uri $GetListItemsEndpoint.Replace("{site-id}", $UserSite.id).Replace("{list-id}", $UserList.id) -scopes $graphScopes
		
		$UserListItemsFields =  $UserListItems.value | Select-Object -ExpandProperty fields
		
		"You are done. Below is the Field Values:"
		
		$UserListItemsFields | Out-Host
	}
}