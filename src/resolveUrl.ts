import type { AzureClientId, AzureClientSecret, AzureTenantId } from "microsoft-graph/AzureApplicationCredentials";

export async function resolveUrl(url: string, tenantId: AzureTenantId, clientId: AzureClientId, clientSecret: AzureClientSecret): Promise<{ siteId: string; driveId: string } | null> {
	try {
	} catch {
		return null;
	}
}
