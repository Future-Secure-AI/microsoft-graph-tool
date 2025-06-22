# Microsoft Graph Tool
A basic CLI tool for interacting with the Microsoft Graph API. This tool allows you to list Microsoft 365 sites and drives using Azure AD application credentials.

## Installation

You can install this tool globally from npm:

```sh
npm install -g microsoft-graph-tool
```

This will make the `microsoft-graph` command available globally.

## Usage

### List all sites

```
microsoft-graph list-sites [search] [--tenantId <tenantId>] [--clientId <clientId>] [--clientSecret <clientSecret>]
```

Lists all sites in your tenant, with an optional `search` string. If credentials are not provided as options, the tool will use the corresponding environment variables.


### List all drives in a site

```
microsoft-graph list-drives <siteId> [--tenantId <tenantId>] [--clientId <clientId>] [--clientSecret <clientSecret>]
```

Lists all drives for the specified site ID. You must provide the `siteId` as a positional argument. Credentials can be provided as options or via environment variables.
