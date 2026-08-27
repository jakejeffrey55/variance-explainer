import Papa from "papaparse";
import type {
  NormalizedProperty,
  PropertySyncInput,
  PropertySyncProvider,
  ProviderIssue,
  ProviderResult,
  ProviderStatus,
} from "@/lib/integrations/types";

/**
 * ACTIVE provider: an admin uploads a CSV export (today, out of Power BI).
 *
 * Column names are matched case-insensitively and accept the aliases that the
 * Power BI and OneSite exports actually use, so an admin can upload the file
 * they already have rather than reshaping it by hand.
 */

const COLUMN_ALIASES: Record<keyof NormalizedProperty | "isActive", string[]> = {
  externalId: ["external_id", "property_id", "propertyid", "property code", "property_code", "site id", "id"],
  name: ["name", "property", "property_name", "property name", "community", "community name"],
  addressLine1: ["address", "address1", "address_line1", "address line 1", "street", "street address"],
  addressLine2: ["address2", "address_line2", "address line 2", "suite", "unit"],
  city: ["city", "municipality"],
  state: ["state", "st", "province", "region"],
  postalCode: ["zip", "zipcode", "zip_code", "postal", "postal_code", "postal code"],
  latitude: ["lat", "latitude", "geo_lat"],
  longitude: ["lng", "long", "lon", "longitude", "geo_lng", "geo_long"],
  unitCount: ["units", "unit_count", "unit count", "total units", "# units", "number of units"],
  propertyManagerName: ["manager", "property_manager", "property manager", "pm_name", "manager name"],
  propertyManagerEmail: ["manager_email", "property manager email", "pm_email", "manager email", "email"],
  propertyManagerPhone: ["manager_phone", "property manager phone", "pm_phone", "manager phone", "phone"],
  isActive: ["active", "is_active", "status", "property status"],
};

/** Columns an admin must supply; everything else is optional. */
export const REQUIRED_COLUMNS = [
  "externalId",
  "name",
  "addressLine1",
  "city",
  "state",
  "postalCode",
  "latitude",
  "longitude",
] as const;

export const CSV_TEMPLATE_HEADERS = [
  "property_id",
  "name",
  "address",
  "address2",
  "city",
  "state",
  "zip",
  "latitude",
  "longitude",
  "units",
  "property_manager",
  "manager_email",
  "manager_phone",
  "active",
];

export const CSV_TEMPLATE_ROWS = [
  [
    "CRT-DAL-0148",
    "Cortland Uptown Dallas",
    "2801 Cedar Springs Rd",
    "",
    "Dallas",
    "TX",
    "75201",
    "32.7997",
    "-96.8065",
    "312",
    "Alicia Gomez",
    "alicia.gomez@example.com",
    "214-555-0201",
    "true",
  ],
];

function normalizeHeader(header: string) {
  return header.trim().toLowerCase().replace(/\s+/g, " ");
}

/** Maps the uploaded file's headers onto our canonical field names. */
export function buildHeaderMap(headers: string[]) {
  const map = new Map<string, string>(); // canonical field -> actual header
  const unmatched: string[] = [];

  for (const header of headers) {
    const normalized = normalizeHeader(header);
    const canonical = (Object.keys(COLUMN_ALIASES) as (keyof typeof COLUMN_ALIASES)[]).find((field) =>
      COLUMN_ALIASES[field].some((alias) => normalizeHeader(alias) === normalized),
    );
    if (canonical && !map.has(canonical)) map.set(canonical, header);
    else if (!canonical) unmatched.push(header);
  }

  const missingRequired = REQUIRED_COLUMNS.filter((field) => !map.has(field));
  return { map, unmatched, missingRequired };
}

function parseNumber(value: string | undefined) {
  if (value === undefined) return undefined;
  const cleaned = value.replace(/[$,\s]/g, "");
  if (cleaned === "") return undefined;
  const n = Number(cleaned);
  return Number.isFinite(n) ? n : undefined;
}

function parseActive(value: string | undefined) {
  if (value === undefined || value.trim() === "") return true;
  const v = value.trim().toLowerCase();
  if (["false", "0", "no", "n", "inactive", "disposed", "sold", "closed"].includes(v)) return false;
  return true;
}

export class CsvPropertySyncProvider implements PropertySyncProvider {
  readonly key = "csv";
  readonly source = "CSV" as const;

  status(): ProviderStatus {
    return {
      key: this.key,
      label: "CSV upload",
      description:
        "Upload a property export (for example, a Power BI export) and import it directly. Active today.",
      configured: true,
      capabilities: { upload: true, pull: false, incremental: false },
    };
  }

  async fetchProperties(input: PropertySyncInput): Promise<ProviderResult<NormalizedProperty[]>> {
    if (!input.fileContent || input.fileContent.trim() === "") {
      return {
        ok: false,
        error: {
          code: "invalid_input",
          message: "No file contents were received.",
          remedy: "Choose a .csv file and upload it again.",
        },
      };
    }

    const parsed = Papa.parse<Record<string, string>>(input.fileContent.trim(), {
      header: true,
      skipEmptyLines: "greedy",
      transformHeader: (h) => h.trim(),
    });

    const headers = parsed.meta.fields ?? [];
    const { map, missingRequired } = buildHeaderMap(headers);

    if (missingRequired.length > 0) {
      return {
        ok: false,
        error: {
          code: "invalid_input",
          message: `The file is missing required column${missingRequired.length > 1 ? "s" : ""}: ${missingRequired.join(", ")}.`,
          remedy:
            "Download the template below to see the expected headers. Column names are matched loosely, so common export names also work.",
          issues: missingRequired.map((field) => ({ field, message: "Required column not found." })),
        },
      };
    }

    const get = (row: Record<string, string>, field: string) => {
      const header = map.get(field);
      const value = header ? row[header] : undefined;
      return value === undefined || value === null ? undefined : String(value).trim();
    };

    const warnings: ProviderIssue[] = [];
    const seenExternalIds = new Set<string>();
    const properties: NormalizedProperty[] = [];

    parsed.data.forEach((row, index) => {
      const rowNumber = index + 2; // +1 for header, +1 for 1-based
      const rowIssues: ProviderIssue[] = [];

      const externalId = get(row, "externalId");
      const name = get(row, "name");
      const addressLine1 = get(row, "addressLine1");
      const city = get(row, "city");
      const state = get(row, "state");
      const postalCode = get(row, "postalCode");
      const latitude = parseNumber(get(row, "latitude"));
      const longitude = parseNumber(get(row, "longitude"));

      const require = (value: string | undefined, field: string) => {
        if (!value) rowIssues.push({ row: rowNumber, field, message: `${field} is required.` });
        return value;
      };

      require(externalId, "externalId");
      require(name, "name");
      require(addressLine1, "addressLine1");
      require(city, "city");
      require(state, "state");
      require(postalCode, "postalCode");

      if (latitude === undefined || longitude === undefined) {
        rowIssues.push({
          row: rowNumber,
          field: "latitude/longitude",
          message: "Latitude and longitude are required — they drive vendor radius matching and the map view.",
        });
      } else if (latitude < -90 || latitude > 90 || longitude < -180 || longitude > 180) {
        rowIssues.push({
          row: rowNumber,
          field: "latitude/longitude",
          message: "Coordinates are out of range.",
          value: `${latitude}, ${longitude}`,
        });
      }

      if (externalId && seenExternalIds.has(externalId)) {
        rowIssues.push({
          row: rowNumber,
          field: "externalId",
          message: "Duplicate property id in this file — only the first row was imported.",
          value: externalId,
        });
      }

      if (rowIssues.length > 0) {
        warnings.push(...rowIssues);
        return;
      }

      seenExternalIds.add(externalId!);
      properties.push({
        externalId: externalId!,
        name: name!,
        addressLine1: addressLine1!,
        addressLine2: get(row, "addressLine2") || null,
        city: city!,
        state: state!.toUpperCase().slice(0, 32),
        postalCode: postalCode!,
        latitude: latitude!,
        longitude: longitude!,
        unitCount: parseNumber(get(row, "unitCount")) ?? null,
        propertyManagerName: get(row, "propertyManagerName") || null,
        propertyManagerEmail: get(row, "propertyManagerEmail")?.toLowerCase() || null,
        propertyManagerPhone: get(row, "propertyManagerPhone") || null,
        isActive: parseActive(get(row, "isActive")),
      });
    });

    if (properties.length === 0) {
      return {
        ok: false,
        error: {
          code: "invalid_input",
          message: "No valid rows were found in the file.",
          remedy: "Fix the issues listed below and upload again.",
          issues: warnings.slice(0, 100),
        },
      };
    }

    return { ok: true, data: properties, warnings };
  }
}
