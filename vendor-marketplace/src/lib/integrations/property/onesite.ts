import type {
  NormalizedProperty,
  PropertySyncInput,
  PropertySyncProvider,
  ProviderResult,
  ProviderStatus,
} from "@/lib/integrations/types";

const REQUIRED_ENV = ["ONESITE_BASE_URL", "ONESITE_API_KEY"] as const;

/**
 * STUB: implements the full PropertySyncProvider interface and returns
 * `not_configured` until OneSite API access is granted.
 *
 * When access lands, implement `fetchProperties` to call the OneSite property
 * endpoint and map each record into `NormalizedProperty`. Whichever of Power BI
 * or OneSite arrives first can be activated by changing
 * PROPERTY_SYNC_PROVIDER — the schema, UI, and import service stay as they are.
 */
export class OneSitePropertyProvider implements PropertySyncProvider {
  readonly key = "onesite";
  readonly source = "ONESITE" as const;

  private missingEnv() {
    return REQUIRED_ENV.filter((name) => !process.env[name]);
  }

  status(): ProviderStatus {
    const missing = this.missingEnv();
    return {
      key: this.key,
      label: "OneSite",
      description: "Sync properties and units directly from OneSite, the system of record.",
      configured: missing.length === 0,
      reason:
        missing.length > 0
          ? `Not configured — awaiting OneSite API access (${missing.join(", ")} not set).`
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
          message: "The OneSite provider is not configured yet.",
          remedy:
            "Use CSV upload for now. Once OneSite API access is granted, set the required environment variables and switch the active provider — no other changes are needed.",
        },
      };
    }

    // Wiring point: authenticated OneSite property fetch + row mapping.
    return {
      ok: false,
      error: {
        code: "provider_error",
        message: "OneSite credentials are present but the client has not been implemented yet.",
        remedy: "Implement OneSitePropertyProvider.fetchProperties against the OneSite API.",
      },
    };
  }
}
