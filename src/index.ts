#!/usr/bin/env node

// Register tsx to allow running TypeScript directly
await import("tsx");

import Table from "cli-table3";
import type { AzureClientId, AzureClientSecret, AzureTenantId } from "microsoft-graph/AzureApplicationCredentials";
import { createClientSecretContext } from "microsoft-graph/context";
import { getEnvironmentVariable } from "microsoft-graph/environmentVariable";
import iterateSites from "microsoft-graph/iterateSites";
import { iterateToArray } from "microsoft-graph/iteration";
import yargs from "yargs";
import { hideBin } from "yargs/helpers";

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

			// Use cli-table3 for table output without (index) and without quotes
			const table = new Table({ head: ["id", "name"] });
			sites.map((site) => table.push([site.id ?? "", site.name ?? ""]));
			process.stdout.write(`${table.toString()}\n`);
		},
	)
	.demandCommand(1, "You need at least one command before moving on.")
	.help()
	.strict()
	.parse();
