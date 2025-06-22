#!/usr/bin/env NODE_NO_WARNINGS=1 node

// Register tsx to allow running TypeScript directly
await import("tsx");

import type { AzureClientId, AzureClientSecret, AzureTenantId } from "microsoft-graph/AzureApplicationCredentials";
import { createClientSecretContext } from "microsoft-graph/context";
import { getEnvironmentVariable } from "microsoft-graph/environmentVariable";
import iterateDrives from "microsoft-graph/iterateDrives";
import iterateSiteSearch from "microsoft-graph/iterateSiteSearch";
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

			const iterator = iterateSiteSearch(contextRef, "*");
			const head = ["id", "name"];
			let colWidths: number[] = [];
			let found = false;
			for await (const site of iterator) {
				const row = [site.id ?? "", site.name ?? ""];
				if (!found) {
					colWidths = head.map((h, i) => Math.max(String(h).length, String(row[i]).length, 10));
					process.stdout.write(`${head.map((h, i) => h.padEnd(colWidths[i] ?? 10)).join(" | ")}\n`);
					process.stdout.write(`${"-".repeat(colWidths.reduce((a, b) => a + b + 3, -3))}\n`);
					found = true;
				}
				process.stdout.write(`${row.map((v, i) => String(v).padEnd(colWidths[i] ?? 10)).join(" | ")}\n`);
			}
			if (!found) {
				process.stdout.write("No sites found.\n");
			}
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
			const iterator = iterateDrives(siteRef);
			const head = ["id", "name"];
			let colWidths: number[] = [];
			let found = false;
			for await (const drive of iterator) {
				const row = [drive.id ?? "", drive.name ?? ""];
				if (!found) {
					colWidths = head.map((h, i) => Math.max(String(h).length, String(row[i]).length, 10));
					process.stdout.write(`${head.map((h, i) => h.padEnd(colWidths[i] ?? 10)).join(" | ")}\n`);
					process.stdout.write(`${"-".repeat(colWidths.reduce((a, b) => a + b + 3, -3))}\n`);
					found = true;
				}
				process.stdout.write(`${row.map((v, i) => String(v).padEnd(colWidths[i] ?? 10)).join(" | ")}\n`);
			}
			if (!found) {
				process.stdout.write("No drives found.\n");
			}
		},
	)
	.demandCommand(1, "You need at least one command before moving on.")
	.help()
	.strict()
	.parse();
