import type {
  NormalizedProperty,
  PropertySyncInput,
  PropertySyncProvider,
  ProviderResult,
  ProviderStatus,
} from "@/lib/integrations/types";

const REQUIRED_ENV = [
  "POWERBI_CLIENT_ID",
  "POWERBI_CLIENT_SECRET",
  "POWERBI_TENANT_ID",
  "POWERBI_DATASET_ID",
] as const;

/**
 * STUB: implements the full PropertySyncProvider interface and returns
 * `not_configured` until Power BI API access is granted.
 *
 * When credentials land, the only change is inside `fetchProperties`: acquire a
 * token against the tenant, execute the dataset query, and map each row into
 * `NormalizedProperty`. Nothing downstream — schema, import service, UI —
 * changes, because the import service only ever sees NormalizedProperty.
 */
export class PowerBIPropertyProvider implements PropertySyncProvider {
  readonly key = "powerbi";
  readonly source = "POWER_BI" as const;

  private missingEnv() {
    return REQUIRED_ENV.filter((name) => !process.env[name]);
  }

  status(): ProviderStatus {
    const missing = this.missingEnv();
    return {
      key: this.key,
      label: "Power BI dataset",
      description:
        "Pull properties directly from the Power BI dataset on a schedule instead of uploading a file.",
      configured: missing.length === 0,
      reason:
        missing.length > 0
          ? `Not configured — awaiting Power BI API access (${missing.join(", ")} not set).`
          : undefined,
      requires: [...REQUIRED_ENV],
      capabilities: { upload: false, pull: true, incremental: true },
    };
  }

  async fetchProperties(_input: PropertySyncInput): Promise<ProviderResult<NormalizedProperty[]>> {
    const status = this.status();
    if (!status.configured) {
      return {
        ok: false,
        error: {
          code: "not_configured",
          message: "The Power BI provider is not configured yet.",
          remedy:
            "Use CSV upload for now. Once Power BI API access is granted, set the required environment variables and switch the active provider — no other changes are needed.",
        },
      };
    }

    // Wiring point: token acquisition + dataset query + row mapping.
    return {
      ok: false,
      error: {
        code: "provider_error",
        message: "Power BI credentials are present but the client has not been implemented yet.",
        remedy: "Implement PowerBIPropertyProvider.fetchProperties against the Power BI REST API.",
      },
    };
  }
}
