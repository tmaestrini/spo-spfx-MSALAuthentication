# sp-fx-msal-authentication

## Summary

The SPFX MSAL Authentication Demo webpart provides a practical example of implementing Microsoft Authentication Library (MSAL) within a SharePoint Framework environment, allowing users to authenticate against Entra ID and obtain access tokens for secure API calls. 

It should serve as a template for developers to understand and implement Entra ID authentication in their own SharePoint Framework solutions.

## Used SharePoint Framework Version

![version](https://img.shields.io/badge/version-1.20.0-green.svg)

## Applies to

- [SharePoint Framework](https://aka.ms/spfx)
- [Microsoft 365 tenant](https://docs.microsoft.com/en-us/sharepoint/dev/spfx/set-up-your-developer-tenant)

> Get your own free development tenant by subscribing to [Microsoft 365 developer program](http://aka.ms/o365devprogram)

## Solution

| Solution         | Author(s)                                               |
| ---------------- | ------------------------------------------------------- |
| spFxMsalAuthDemo | Tobias Maestrini (tobias.maestrini@gmail.com, https://bsky.app/profile/tmaestrini.bsky.social) |

## Version history

| Version | Date             | Comments        |
| ------- | ---------------- | --------------- |
| 1.0     | January 03, 2025 | Initial release |

## Disclaimer

**THIS CODE IS PROVIDED _AS IS_ WITHOUT WARRANTY OF ANY KIND, EITHER EXPRESS OR IMPLIED, INCLUDING ANY IMPLIED WARRANTIES OF FITNESS FOR A PARTICULAR PURPOSE, MERCHANTABILITY, OR NON-INFRINGEMENT.**

---

## Minimal Path to Awesome

- Clone this repository
- Ensure that you are at the solution folder
- in the command-line run:
  - **npm install**
  - **gulp serve**

Alternatively, you can use `npm run serve` by making use of the `fast-serve` module.

## Description

The SPFX MSAL Authentication Demo webpart provides a practical example of implementing Microsoft Authentication Library (MSAL) within a SharePoint Framework environment, allowing users to authenticate against Entra ID and obtain access tokens for secure API calls. It serves as a template for developers to understand and implement Entra ID authentication in their own SharePoint Framework solutions.

### Technical Details

- Framework: SharePoint Framework (SPFx)
- Authentication: MSAL.js v2
- Language: TypeScript/React
- Key Components used in the solution:
  - `AuthenticationContext`: Manages auth state
  - `PublicClientApplication`: MSAL instance
  - `TokenAcquisition`: Silent and interactive flows
  - `StateManagement`: React Context API

### Implementation

- Configurable client ID and tenant ID
- Custom scope management
- Token caching in sessionStorage
- Silent token refresh
- Interactive login fallback
- Error handling for auth failures


### Configuration Requirements

- Entra ID
  - App registration
  - Authentication type: Single-page application (SPA)
  - Implicit authorisation and hybrid flows: ID token
  - Application type: single tenant application
  - Redirect URI: https://{tenant}.sharepoint.com/_layouts/15/workbench.aspx
  - Required API Permissions: Microsoft Graph (User.Read), any further permission(s) along your (demo) needs

## References

- Microsoft Learn: [Acquiring and Using an (MSAL) Access Token](https://github.com/AzureAD/microsoft-authentication-library-for-js/blob/dev/lib/msal-browser/docs/acquire-token.md)
