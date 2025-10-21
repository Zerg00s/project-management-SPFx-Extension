# Project management SPFX solution

## Node.js Version 22.14.0

## Summary

SPFX extension that adds the "Create a project" button to the command bar of a document library. This button allows users to create a new project site directly from the document library interface, streamlining project management and collaboration within SharePoint. The button is only available in the document library called "Projects". This can be changed in the code as needed.

The project folders are copied based on the document library called "Project Templates". Each folder there is a project template that can be used to create a new project folder. The entire structure of the template folder is copied to the new project folder, including all subfolders and files.

When a user clicks the "Create a project" button, a dialog box appears prompting them to enter the name of the new project. Once the user submits the form, a new folder is created in the "Projects" document library with the specified name, and the contents of the selected template folder are copied into it.

## Screenshots

![alt text](image.png)

![alt text](image-1.png)

![alt text](image-2.png)

![alt text](image-3.png)

## Used SharePoint Framework Version

![version](https://img.shields.io/badge/version-1.21.1-green.svg)

## Applies to

- [SharePoint Framework](https://aka.ms/spfx)
- [Microsoft 365 tenant](https://docs.microsoft.com/en-us/sharepoint/dev/spfx/set-up-your-developer-tenant)

> Get your own free development tenant by subscribing to [Microsoft 365 developer program](http://aka.ms/o365devprogram)

## Prerequisites

> Node.js v22.14.0
> Create 2 document libraries in your SharePoint site:
> - "Projects" - where new project folders will be created
> - "Project Templates" - where project templates are stored as folders

## Minimal Path to Awesome

- Clone this repository
- Ensure that you are at the solution folder
- in the command-line run:
  - **npm install**
  - **gulp serve**

## Build the project

```cmd
gulp clean ; gulp bundle --ship ; gulp package-solution --ship
```

## Deploy the package

- Create site collection app catalog if you don't have one already:

```powershell
Add-SPOSiteCollectionAppCatalog -Site https://contoso.sharepoint.com/sites/project_site_of_your_choice
```
- upload the generated .sppkg file from the `sharepoint/solution` folder to the site collection app catalog's "Apps for SharePoint" library

![alt text](image-4.png)

- Site contents -> Add an app -> select the  app

![alt text](image-5.png)


## References

- [Getting started with SharePoint Framework](https://docs.microsoft.com/en-us/sharepoint/dev/spfx/set-up-your-developer-tenant)
- [Building for Microsoft teams](https://docs.microsoft.com/en-us/sharepoint/dev/spfx/build-for-teams-overview)
- [Use Microsoft Graph in your solution](https://docs.microsoft.com/en-us/sharepoint/dev/spfx/web-parts/get-started/using-microsoft-graph-apis)
- [Publish SharePoint Framework applications to the Marketplace](https://docs.microsoft.com/en-us/sharepoint/dev/spfx/publish-to-marketplace-overview)
- [Microsoft 365 Patterns and Practices](https://aka.ms/m365pnp) - Guidance, tooling, samples and open-source controls for your Microsoft 365 development