#!/usr/bin/env node

// Register tsx to allow running TypeScript directly
await import("tsx");

import Table from "cli-table3";
import type { AzureClientId, AzureClientSecret, AzureTenantId } from "microsoft-graph/AzureApplicationCredentials";
import { createClientSecretContext } from "microsoft-graph/context";
import { getEnvironmentVariable } from "microsoft-graph/environmentVariable";
import iterateSites from "microsoft-graph/iterateSites";
import { iterateToArray } from "microsoft-graph/iteration";
import type { SiteId } from "microsoft-graph/Site";
import { createSiteRef } from "microsoft-graph/site";
import yargs from "yargs";
import { hideBin } from "yargs/helpers";
import iterateDrives from "microsoft-graph/iterateDrives";

yargs(hideBin(process.argv))
	.command(
		"list-sites",
		"List all sites in your company geography",
		(yargs) =>
			yargs
				.option("tenantId", {
					type: "string",
					describe: "Azure Tenant ID (defaults to AZURE_TENANT_ID env)",
				})
				.option("clientId", {
					type: "string",
					describe: "Azure Client ID (defaults to AZURE_CLIENT_ID env)",
				})
				.option("clientSecret", {
					type: "string",
					describe: "Azure Client Secret (defaults to AZURE_CLIENT_SECRET env)",
				}),
		async (argv: any) => {
			const tenantId = (argv.tenantId || getEnvironmentVariable("AZURE_TENANT_ID")) as AzureTenantId;
			const clientId = (argv.clientId || getEnvironmentVariable("AZURE_CLIENT_ID")) as AzureClientId;
			const clientSecret = (argv.clientSecret || getEnvironmentVariable("AZURE_CLIENT_SECRET")) as AzureClientSecret;

			const contextRef = createClientSecretContext(tenantId, clientId, clientSecret);

			const sites = await iterateToArray(iterateSites(contextRef));
			if (sites.length === 0) {
				process.stdout.write("No sites found.\n");
				return;
			}

			const table = new Table({ head: ["id", "name"] });
			sites.map((site) => table.push([site.id ?? "", site.name ?? ""]));
			process.stdout.write(`${table.toString()}\n`);
		},
	)
	.command(
		"list-drives <siteId>",
		"List all drives in a site",
		(yargs) =>
			yargs.positional("siteId", {
				describe: "Site ID to list drives for",
				type: "string",
			}),
		async (argv: { siteId: string; tenantId?: string; clientId?: string; clientSecret?: string }) => {
			const tenantId = (argv.tenantId || getEnvironmentVariable("AZURE_TENANT_ID")) as AzureTenantId;
			const clientId = (argv.clientId || getEnvironmentVariable("AZURE_CLIENT_ID")) as AzureClientId;
			const clientSecret = (argv.clientSecret || getEnvironmentVariable("AZURE_CLIENT_SECRET")) as AzureClientSecret;
			const siteId = argv.siteId as SiteId;

			if (!siteId) {
				process.stderr.write("siteId is required.\n");
				process.exit(1);
			}

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
