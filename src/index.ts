#!/usr/bin/env node

import { cac } from "cac";
import type { AzureClientId, AzureClientSecret, AzureTenantId } from "microsoft-graph/AzureApplicationCredentials";
import { createClientSecretContext } from "microsoft-graph/context";
import { getEnvironmentVariable } from "microsoft-graph/environmentVariable";
import getDriveFromUrl from "microsoft-graph/getDriveFromUrl";
import iterateDrives from "microsoft-graph/iterateDrives";
import iterateSiteSearch from "microsoft-graph/iterateSiteSearch";
import { parseSharepointUrl } from "microsoft-graph/sharepointUrl";
import type { SiteId } from "microsoft-graph/Site";
import { createSiteRef } from "microsoft-graph/site";
import process from "node:process";

type BaseArgs = {
	tenantId: AzureTenantId;
	clientId: AzureClientId;
	clientSecret: AzureClientSecret;
};

const cli = cac("graph-tool")
	.option("--tenantId <id>", "Azure Tenant ID", {
		default: getEnvironmentVariable("AZURE_TENANT_ID"),
	})
	.option("--clientId <id>", "Azure Client ID", {
		default: getEnvironmentVariable("AZURE_CLIENT_ID"),
	})
	.option("--clientSecret <secret>", "Azure Client Secret", {
		default: getEnvironmentVariable("AZURE_CLIENT_SECRET"),
	});

cli.command("list-sites [search]", "List all sites.").action(async (search: string | undefined, { tenantId, clientId, clientSecret }: BaseArgs) => {
	const searchTerm = search ?? "*";
	const contextRef = createClientSecretContext(tenantId, clientId, clientSecret);

	const iterator = iterateSiteSearch(contextRef, searchTerm);
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
});

cli.command("list-drives <siteId>", "List all drives in a site.").action(async (siteId: SiteId, { tenantId, clientId, clientSecret }: BaseArgs) => {
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
});

cli.command("resolve-url <url>", "Resolve a SharePoint URL to siteId and driveId.").action(async (url: string, { tenantId, clientId, clientSecret }: BaseArgs) => {
	const { hostName, siteName, driveName } = parseSharepointUrl(url);

	if (!hostName) {
		process.stdout.write("Invalid SharePoint URL: Host name is missing.");
		return;
	}
	if (!siteName) {
		process.stdout.write("Invalid SharePoint URL: Site name is missing.");
		return;
	}
	if (!driveName) {
		process.stdout.write("Invalid SharePoint URL: Drive name is missing.");
		return;
	}

	const contextRef = createClientSecretContext(tenantId, clientId, clientSecret);

	const drive = await getDriveFromUrl(contextRef, url);
	process.stdout.write(`Hostname: ${hostName}\n`);
	process.stdout.write(`Site Name: ${siteName}\n`);
	process.stdout.write(`Drive Name: ${driveName}\n`);
	process.stdout.write(`Site ID: ${drive.siteId}\n`);
	process.stdout.write(`Drive ID: ${drive.id}\n`);
});

cli.help();
cli.parse();
