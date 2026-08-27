import type { PropertySyncProvider } from "@/lib/integrations/types";
import { CsvPropertySyncProvider } from "./csv";
import { OneSitePropertyProvider } from "./onesite";
import { PowerBIPropertyProvider } from "./powerbi";

/**
 * Provider registry. Activating a different source is a one-line env change
 * (PROPERTY_SYNC_PROVIDER=powerbi|onesite) — callers resolve providers by key
 * and only ever handle NormalizedProperty.
 */
export const propertyProviders: Record<string, PropertySyncProvider> = {
  csv: new CsvPropertySyncProvider(),
  powerbi: new PowerBIPropertyProvider(),
  onesite: new OneSitePropertyProvider(),
};

export const DEFAULT_PROPERTY_PROVIDER_KEY = process.env.PROPERTY_SYNC_PROVIDER ?? "csv";

export function getPropertyProvider(key?: string): PropertySyncProvider {
  return propertyProviders[key ?? DEFAULT_PROPERTY_PROVIDER_KEY] ?? propertyProviders.csv;
}

export function listPropertyProviderStatuses() {
  return Object.values(propertyProviders).map((p) => ({
    ...p.status(),
    active: p.key === DEFAULT_PROPERTY_PROVIDER_KEY,
  }));
}

export { CsvPropertySyncProvider, PowerBIPropertyProvider, OneSitePropertyProvider };
