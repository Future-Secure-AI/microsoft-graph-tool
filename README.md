# Microsoft Graph Tool
A basic CLI tool for interacting with the Microsoft Graph API. This tool allows you to list Microsoft 365 sites and drives using Azure AD application credentials.

## Installation

You can install this tool globally from npm:

```sh
npm install -g microsoft-graph-tool
```

This will make the `graph` command available globally.

## Usage

### Get SiteID from a Sharepoint URL

```sh
graph get-site <url> [--tenantId <tenantId>] [--clientId <clientId>] [--clientSecret <clientSecret>]
```

Resolves a SharePoint URL to its corresponding site ID and drive ID. Provide the full SharePoint URL as the positional argument. Credentials can be provided as options or via environment variables.

### List all sites

```sh
graph list-sites [search] [--tenantId <tenantId>] [--clientId <clientId>] [--clientSecret <clientSecret>]
```

Lists all sites in your tenant, with an optional `search` string. If credentials are not provided as options, the tool will use the corresponding environment variables.

### List all drives in a site

```sh
graph list-drives <siteId> [--tenantId <tenantId>] [--clientId <clientId>] [--clientSecret <clientSecret>]
```

Lists all drives for the specified site ID. You must provide the `siteId` as a positional argument. Credentials can be provided as options or via environment variables.


