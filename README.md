# people-directory

## Summary

A SPFx web part to list people, locations and statuses from the Microsoft Graph API 

## Used SharePoint Framework Version

![version](https://img.shields.io/badge/version-1.19.0-green.svg)

## Applies to

- [SharePoint Framework](https://aka.ms/spfx)
- [Microsoft 365 tenant](https://docs.microsoft.com/en-us/sharepoint/dev/spfx/set-up-your-developer-tenant)

> Get your own free development tenant by subscribing to [Microsoft 365 developer program](http://aka.ms/o365devprogram)

## Prerequisites

  "dependencies": {
    "@fluentui/react": "^8.106.4",
    "@microsoft/sp-component-base": "1.19.0",
    "@microsoft/sp-core-library": "1.19.0",
    "@microsoft/sp-http": "latest",
    "@microsoft/sp-lodash-subset": "1.19.0",
    "@microsoft/sp-office-ui-fabric-core": "1.19.0",
    "@microsoft/sp-property-pane": "1.19.0",
    "@microsoft/sp-webpart-base": "1.19.0",
    "office-ui-fabric-react": "latest",
    "react": "17.0.1",
    "react-dom": "17.0.1",
    "tslib": "2.3.1"
  },
  "devDependencies": {
    "@microsoft/eslint-config-spfx": "1.20.1",
    "@microsoft/eslint-plugin-spfx": "1.20.1",
    "@microsoft/microsoft-graph-types": "2.40.0",
    "@microsoft/rush-stack-compiler-4.7": "0.1.0",
    "@microsoft/sp-build-web": "1.20.1",
    "@microsoft/sp-module-interfaces": "1.20.1",
    "@rushstack/eslint-config": "2.5.1",
    "@types/react": "17.0.45",
    "@types/react-dom": "17.0.17",
    "@types/webpack-env": "~1.15.2",
    "ajv": "^6.12.5",
    "eslint": "8.7.0",
    "eslint-plugin-react-hooks": "4.3.0",
    "gulp": "4.0.2",
    "typescript": 

## Disclaimer

**THIS CODE IS PROVIDED _AS IS_ WITHOUT WARRANTY OF ANY KIND, EITHER EXPRESS OR IMPLIED, INCLUDING ANY IMPLIED WARRANTIES OF FITNESS FOR A PARTICULAR PURPOSE, MERCHANTABILITY, OR NON-INFRINGEMENT.**

---

## Minimal Path to Awesome

- Clone this repository
- Ensure that you are at the solution folder
- in the command-line run:
  - **npm install**
  - **gulp serve**

> Include any additional steps as needed.

## Features

The People Directory web part can be placed on any Sharepoint page and will list all the individuals in an organisation, along with their job title, department and their prescence. It will also list their current status message OR their current "Out of Office" message.

You can filter by name or department.

All details are taken from their Microsoft 365 profile and obtained via the Graph API.

## References

- [Getting started with SharePoint Framework](https://docs.microsoft.com/en-us/sharepoint/dev/spfx/set-up-your-developer-tenant)
- [Building for Microsoft teams](https://docs.microsoft.com/en-us/sharepoint/dev/spfx/build-for-teams-overview)
- [Use Microsoft Graph in your solution](https://docs.microsoft.com/en-us/sharepoint/dev/spfx/web-parts/get-started/using-microsoft-graph-apis)
- [Publish SharePoint Framework applications to the Marketplace](https://docs.microsoft.com/en-us/sharepoint/dev/spfx/publish-to-marketplace-overview)
- [Microsoft 365 Patterns and Practices](https://aka.ms/m365pnp) - Guidance, tooling, samples and open-source controls for your Microsoft 365 development
