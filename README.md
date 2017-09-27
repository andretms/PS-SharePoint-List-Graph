# Get a SharePoint List via PowerShell using Microsoft Graph API
This sample PowerShell script demonstrate how to get  SharePoint List Items/fieds via PowerShell using the Microsoft Graph API.

## 1. Download MSAL Library
This sample requires MSAL Library for obtaining an Access Token to query the Microsoft Graph API. Because PowerShell does not have direct integration with nuget, below is instructions for manual download:

1. Go to https://www.nuget.org/packages/Microsoft.Identity.Client and select 'Manual Download' to download the MSAL `nupkg` file.
2. Rename the file extension to .zip 
3. Extract the zip file to `{root}\packages\Microsoft.Identity.Client.1.1.0-preview` - where `{root}` is the folder where your script is saved

## 2. Register an Application
You need to register an application to be able to access the Microsoft Graph API. In order to do this:

1. Go to https://apps.dev.microsoft.com/portal/register-app
2. Add a name for the application and make sure the *Guided Setup* option is **unchecked**
3. Click `Create`
4. Now configure the new application to be a `Native App` by clicking `Add Platforms` and selecting `Native Application`
5. Copy the Guid under `Application Id` to the clipboard

## 3. Configure your PowerShell Script

1. Open your PowerShell script and replace the `YourAppIdHere` with the Application Id for your application you just registered
2. Now add your Site Relative URL and the List Name by replacing the values of `YourSiteRelativeUrlHere` and `YourListNameHere`

```PowerShell
##Please replace the three values below
$appId = "YourAppIdHere" 
$SiteRelativeUrl = "YourSiteRelativeUrlHere"
$ListName = "YourListNameHere"
```

> #### What is my _Site Relative URL_ and my _List Name_?
> Access your list in a browser, then check the image below:
> ![Relative Site](./_pictures/Site-Relative-URL.PNG)
