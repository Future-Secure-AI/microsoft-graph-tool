#!/usr/bin/env NODE_NO_WARNINGS=1 node

// Register tsx to allow running TypeScript directly
await import("tsx");

import Table from "cli-table3";
import type { AzureClientId, AzureClientSecret, AzureTenantId } from "microsoft-graph/AzureApplicationCredentials";
import { createClientSecretContext } from "microsoft-graph/context";
import { getEnvironmentVariable } from "microsoft-graph/environmentVariable";
import iterateDrives from "microsoft-graph/iterateDrives";
import iterateSiteSearch from "microsoft-graph/iterateSiteSearch";
import { iterateToArray } from "microsoft-graph/iteration";
import type { SiteId } from "microsoft-graph/Site";
import { createSiteRef } from "microsoft-graph/site";
import process from "node:process";
import yargs from "yargs";
import { hideBin } from "yargs/helpers";

type BaseArgs = {
	tenantId: AzureTenantId;
	clientId: AzureClientId;
	clientSecret: AzureClientSecret;
};
type ListSitesArgs = BaseArgs & {};

type ListDrivesArgs = BaseArgs & {
	siteId: SiteId;
};

yargs(hideBin(process.argv))
	.options({
		tenantId: {
			type: "string",
			describe: "Azure Tenant ID (defaults to AZURE_TENANT_ID env)",
			default: getEnvironmentVariable("AZURE_TENANT_ID"),
		},
		clientId: {
			type: "string",
			describe: "Azure Client ID (defaults to AZURE_CLIENT_ID env)",
			default: getEnvironmentVariable("AZURE_CLIENT_ID"),
		},
		clientSecret: {
			type: "string",
			describe: "Azure Client Secret (defaults to AZURE_CLIENT_SECRET env)",
			default: getEnvironmentVariable("AZURE_CLIENT_SECRET"),
		},
	})
	.command<ListSitesArgs>(
		"list-sites",
		"List all sites.",
		(yargs) => yargs,
		async ({ tenantId, clientId, clientSecret }: ListSitesArgs) => {
			const contextRef = createClientSecretContext(tenantId, clientId, clientSecret);

			const sites = await iterateToArray(iterateSiteSearch(contextRef, "*"));
			if (sites.length === 0) {
				process.stdout.write("No sites found.\n");
				return;
			}

			const table = new Table({ head: ["id", "name"] });
			sites.map((site) => table.push([site.id ?? "", site.name ?? ""]));
			process.stdout.write(`${table.toString()}\n`);
		},
	)
	.command<ListDrivesArgs>(
		"list-drives <siteId>",
		"List all drives in a site.",
		(yargs) =>
			yargs.positional("siteId", {
				describe: "Site ID to list drives for",
				type: "string",
			}),
		async ({ tenantId, clientId, clientSecret, siteId }: ListDrivesArgs) => {
			const contextRef = createClientSecretContext(tenantId, clientId, clientSecret);
			const siteRef = createSiteRef(contextRef, siteId);
			const drives = await iterateToArray(iterateDrives(siteRef));
			if (drives.length === 0) {
				process.stdout.write("No drives found.\n");
				return;
			}

			const table = new Table({ head: ["id", "name"] });
			drives.map((drive) => table.push([drive.id ?? "", drive.name ?? ""]));
			process.stdout.write(`${table.toString()}\n`);
		},
	)
	.demandCommand(1, "You need at least one command before moving on.")
	.help()
	.strict()
	.parse();
