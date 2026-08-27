import type { ActivityEntityType, ActorType, Prisma } from "@prisma/client";
import { prisma } from "@/lib/db";

/**
 * Every job/bid/contract/property state change is written here. The job detail
 * page renders these rows as its timeline (phase 10), so the log is the record
 * of what happened — not a debugging aid.
 */
export type ActivityInput = {
  entityType: ActivityEntityType;
  entityId: string;
  jobId?: string | null;
  action: string;
  fromStatus?: string | null;
  toStatus?: string | null;
  actorType: ActorType;
  actorAdminId?: string | null;
  actorVendorId?: string | null;
  summary?: string | null;
  metadata?: Prisma.InputJsonValue;
};

export async function logActivity(input: ActivityInput, tx: Prisma.TransactionClient | typeof prisma = prisma) {
  return tx.activityLog.create({
    data: {
      entityType: input.entityType,
      entityId: input.entityId,
      jobId: input.jobId ?? null,
      action: input.action,
      fromStatus: input.fromStatus ?? null,
      toStatus: input.toStatus ?? null,
      actorType: input.actorType,
      actorAdminId: input.actorAdminId ?? null,
      actorVendorId: input.actorVendorId ?? null,
      summary: input.summary ?? null,
      metadata: input.metadata,
    },
  });
}
