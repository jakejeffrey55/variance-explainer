import type { ImportStatus, PropertySource } from "@prisma/client";
import { prisma } from "@/lib/db";
import { getPropertyProvider } from "@/lib/integrations/property";
import type {
  NormalizedProperty,
  PropertySyncInput,
  ProviderError,
  ProviderIssue,
} from "@/lib/integrations/types";
import { logActivity } from "@/lib/services/activity";

/**
 * Provider-agnostic property import.
 *
 * This service knows nothing about CSV, Power BI, or OneSite — it asks a
 * provider for `NormalizedProperty[]` and upserts. That is the seam that makes
 * "whichever API access lands first" a configuration change: only the provider
 * differs, everything from here down is identical.
 */

export type ImportPreviewRow = {
  externalId: string;
  name: string;
  city: string;
  state: string;
  action: "create" | "update" | "unchanged";
  changes?: string[];
};

export type ImportResult = {
  dryRun: boolean;
  providerKey: string;
  batchId?: string;
  rowsTotal: number;
  created: number;
  updated: number;
  unchanged: number;
  failed: number;
  issues: ProviderIssue[];
  preview: ImportPreviewRow[];
};

const COMPARED_FIELDS = [
  "name",
  "addressLine1",
  "addressLine2",
  "city",
  "state",
  "postalCode",
  "latitude",
  "longitude",
  "unitCount",
  "propertyManagerName",
  "propertyManagerEmail",
  "propertyManagerPhone",
  "isActive",
] as const;

function diffFields(existing: Record<string, unknown>, incoming: NormalizedProperty) {
  const changes: string[] = [];
  for (const field of COMPARED_FIELDS) {
    const before = existing[field] ?? null;
    const after = (incoming as Record<string, unknown>)[field] ?? null;
    const changed =
      typeof before === "number" || typeof after === "number"
        ? Number(before) !== Number(after)
        : String(before) !== String(after);
    if (changed) changes.push(field);
  }
  return changes;
}

export async function importProperties(opts: {
  providerKey: string;
  input: PropertySyncInput;
  adminUserId: string;
  dryRun: boolean;
}): Promise<{ ok: true; result: ImportResult } | { ok: false; error: ProviderError }> {
  const provider = getPropertyProvider(opts.providerKey);
  const fetched = await provider.fetchProperties(opts.input);

  if (!fetched.ok) {
    // A failed import is still worth recording when it was a real attempt.
    if (!opts.dryRun) {
      await prisma.importBatch.create({
        data: {
          kind: "PROPERTY",
          providerKey: provider.key,
          status: "FAILED",
          fileName: opts.input.fileName ?? null,
          adminUserId: opts.adminUserId,
          errors: JSON.parse(JSON.stringify(fetched.error)),
          finishedAt: new Date(),
        },
      });
    }
    return { ok: false, error: fetched.error };
  }

  const incoming = fetched.data;
  const source = provider.source as PropertySource;

  const existing = await prisma.property.findMany({
    where: { source, externalId: { in: incoming.map((p) => p.externalId) } },
  });
  const existingByExternalId = new Map(existing.map((p) => [p.externalId ?? "", p]));

  const preview: ImportPreviewRow[] = [];
  let created = 0;
  let updated = 0;
  let unchanged = 0;

  for (const property of incoming) {
    const match = existingByExternalId.get(property.externalId);
    if (!match) {
      created += 1;
      preview.push({
        externalId: property.externalId,
        name: property.name,
        city: property.city,
        state: property.state,
        action: "create",
      });
      continue;
    }
    const changes = diffFields(match as unknown as Record<string, unknown>, property);
    if (changes.length === 0) {
      unchanged += 1;
      preview.push({
        externalId: property.externalId,
        name: property.name,
        city: property.city,
        state: property.state,
        action: "unchanged",
      });
    } else {
      updated += 1;
      preview.push({
        externalId: property.externalId,
        name: property.name,
        city: property.city,
        state: property.state,
        action: "update",
        changes,
      });
    }
  }

  const failedRows = new Set(fetched.warnings.map((w) => w.row).filter(Boolean)).size;

  if (opts.dryRun) {
    return {
      ok: true,
      result: {
        dryRun: true,
        providerKey: provider.key,
        rowsTotal: incoming.length + failedRows,
        created,
        updated,
        unchanged,
        failed: failedRows,
        issues: fetched.warnings,
        preview,
      },
    };
  }

  const batch = await prisma.importBatch.create({
    data: {
      kind: "PROPERTY",
      providerKey: provider.key,
      status: "PROCESSING" as ImportStatus,
      fileName: opts.input.fileName ?? null,
      rowsTotal: incoming.length + failedRows,
      adminUserId: opts.adminUserId,
    },
  });

  const syncedAt = new Date();
  await prisma.$transaction(
    incoming.map((property) =>
      prisma.property.upsert({
        where: { property_source_external_id: { source, externalId: property.externalId } },
        create: {
          ...property,
          source,
          sourceBatchId: batch.id,
          lastSyncedAt: syncedAt,
        },
        update: {
          ...property,
          sourceBatchId: batch.id,
          lastSyncedAt: syncedAt,
        },
      }),
    ),
  );

  await prisma.importBatch.update({
    where: { id: batch.id },
    data: {
      status: "COMPLETED",
      rowsImported: created + updated,
      rowsSkipped: unchanged,
      rowsFailed: failedRows,
      errors: fetched.warnings.length > 0 ? JSON.parse(JSON.stringify(fetched.warnings)) : undefined,
      finishedAt: new Date(),
    },
  });

  await logActivity({
    entityType: "PROPERTY",
    entityId: batch.id,
    action: "property.import",
    actorType: "ADMIN",
    actorAdminId: opts.adminUserId,
    summary: `Imported ${created + updated} propert${created + updated === 1 ? "y" : "ies"} via ${provider.key} (${created} new, ${updated} updated, ${unchanged} unchanged${failedRows ? `, ${failedRows} skipped` : ""}).`,
    metadata: { providerKey: provider.key, fileName: opts.input.fileName ?? null },
  });

  return {
    ok: true,
    result: {
      dryRun: false,
      providerKey: provider.key,
      batchId: batch.id,
      rowsTotal: incoming.length + failedRows,
      created,
      updated,
      unchanged,
      failed: failedRows,
      issues: fetched.warnings,
      preview,
    },
  };
}
