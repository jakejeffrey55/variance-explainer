import { Prisma, type JobStatus } from "@prisma/client";
import { prisma } from "@/lib/db";
import { logActivity } from "@/lib/services/activity";
import { ALLOWED_STATUS_ACTIONS, type JobInput, type JobStatusAction } from "@/lib/validation/job";
import { formatCurrency } from "@/lib/utils";

export class JobRuleError extends Error {
  constructor(
    readonly code: string,
    message: string,
  ) {
    super(message);
    this.name = "JobRuleError";
  }
}

/**
 * Next sequential job number (J-1001, J-1002, …). Retried on collision so two
 * admins creating jobs at the same moment cannot both take the same number.
 */
async function nextJobNumber(tx: Prisma.TransactionClient) {
  const latest = await tx.job.findFirst({
    orderBy: { jobNumber: "desc" },
    select: { jobNumber: true },
  });
  const current = latest ? Number.parseInt(latest.jobNumber.replace(/\D/g, ""), 10) : 1000;
  return `J-${(Number.isFinite(current) ? current : 1000) + 1}`;
}

function toDecimal(value: number) {
  return new Prisma.Decimal(value.toFixed(2));
}

export async function createJob(input: JobInput, adminUserId: string) {
  const property = await prisma.property.findUnique({ where: { id: input.propertyId } });
  if (!property) throw new JobRuleError("property_not_found", "That property no longer exists.");
  if (!property.isActive) {
    throw new JobRuleError("property_inactive", "That property is inactive — reactivate it before posting jobs.");
  }

  const isEmergency = input.priority === "EMERGENCY";
  // An emergency is dispatched the moment it is created; a standard job is a
  // draft until the admin publishes it.
  const status: JobStatus = isEmergency ? "OPEN" : (input.status ?? "DRAFT");

  if (!isEmergency && status === "OPEN" && input.bidDeadline && input.bidDeadline.getTime() <= Date.now()) {
    throw new JobRuleError("deadline_in_past", "The bid deadline has already passed. Choose a future date.");
  }

  const settings = await prisma.approvalSettings.findUnique({ where: { id: "default" } });

  for (let attempt = 0; attempt < 5; attempt++) {
    try {
      return await prisma.$transaction(async (tx) => {
        const jobNumber = await nextJobNumber(tx);
        const job = await tx.job.create({
          data: {
            jobNumber,
            propertyId: input.propertyId,
            unitNumber: input.unitNumber || null,
            title: input.title,
            description: input.description,
            category: input.category,
            status,
            budgetMin: toDecimal(input.budgetMin),
            budgetMax: toDecimal(input.budgetMax),
            // A cap is only meaningful where bids exist.
            enforceBudgetCap: isEmergency ? false : input.enforceBudgetCap,
            bidDeadline: isEmergency ? null : (input.bidDeadline ?? null),
            scheduledStart: input.scheduledStart ?? null,
            dueDate: input.dueDate ?? null,
            priority: input.priority,
            emergencyCategory: isEmergency ? (input.emergencyCategory ?? null) : null,
            responseDeadlineMinutes: isEmergency
              ? (input.responseDeadlineMinutes ?? settings?.emergencyResponseMinutes ?? 15)
              : null,
            dispatchedAt: isEmergency ? new Date() : null,
            inviteOnly: isEmergency ? false : input.inviteOnly,
            createdByAdminId: adminUserId,
          },
          include: { property: true },
        });

        await logActivity(
          {
            entityType: "JOB",
            entityId: job.id,
            jobId: job.id,
            action: isEmergency ? "emergency.created" : "job.created",
            toStatus: job.status,
            actorType: "ADMIN",
            actorAdminId: adminUserId,
            summary: isEmergency
              ? `Emergency job ${job.jobNumber} created (${job.emergencyCategory}) with a ${job.responseDeadlineMinutes}-minute response window.`
              : `Job ${job.jobNumber} created as ${job.status === "OPEN" ? "open for bidding" : "a draft"} (${formatCurrency(input.budgetMin)}–${formatCurrency(input.budgetMax)}${input.enforceBudgetCap ? ", cap enforced" : ""}).`,
            metadata: { propertyId: job.propertyId, category: job.category },
          },
          tx,
        );

        return job;
      });
    } catch (err) {
      const collision =
        err instanceof Prisma.PrismaClientKnownRequestError &&
        err.code === "P2002" &&
        String(err.meta?.target ?? "").includes("job_number");
      if (!collision) throw err;
    }
  }
  throw new JobRuleError("job_number_conflict", "Could not allocate a job number. Please try again.");
}

export async function updateJob(jobId: string, input: Partial<JobInput>, adminUserId: string) {
  const existing = await prisma.job.findUnique({ where: { id: jobId } });
  if (!existing) throw new JobRuleError("job_not_found", "Job not found.");
  if (existing.status === "COMPLETED" || existing.status === "CANCELLED") {
    throw new JobRuleError("job_locked", `A ${existing.status.toLowerCase()} job can no longer be edited.`);
  }

  const bidCount = await prisma.bid.count({ where: { jobId, status: { not: "WITHDRAWN" } } });
  const budgetMin = input.budgetMin ?? Number(existing.budgetMin);
  const budgetMax = input.budgetMax ?? Number(existing.budgetMax);
  if (budgetMax < budgetMin) {
    throw new JobRuleError("budget_range", "Maximum budget must be at least the minimum.");
  }
  // Tightening the rules mid-bidding would invalidate bids vendors already
  // placed in good faith.
  if (bidCount > 0) {
    if (input.enforceBudgetCap === true && !existing.enforceBudgetCap) {
      throw new JobRuleError(
        "cap_after_bids",
        `Cannot turn on the budget cap after ${bidCount} bid${bidCount === 1 ? " has" : "s have"} been submitted.`,
      );
    }
    if (input.budgetMax !== undefined && existing.enforceBudgetCap && budgetMax < Number(existing.budgetMax)) {
      throw new JobRuleError(
        "cap_lowered_after_bids",
        "Cannot lower an enforced budget cap while bids are open.",
      );
    }
  }

  const changed: string[] = [];
  const data: Prisma.JobUpdateInput = {};
  const assign = <K extends keyof Prisma.JobUpdateInput>(key: K, value: Prisma.JobUpdateInput[K], label = key as string) => {
    if (value !== undefined) {
      data[key] = value;
      changed.push(label);
    }
  };

  assign("title", input.title);
  assign("description", input.description);
  assign("category", input.category);
  assign("unitNumber", input.unitNumber === undefined ? undefined : input.unitNumber || null);
  if (input.budgetMin !== undefined) assign("budgetMin", toDecimal(input.budgetMin));
  if (input.budgetMax !== undefined) assign("budgetMax", toDecimal(input.budgetMax));
  assign("enforceBudgetCap", input.enforceBudgetCap);
  assign("bidDeadline", input.bidDeadline === undefined ? undefined : input.bidDeadline);
  assign("scheduledStart", input.scheduledStart === undefined ? undefined : input.scheduledStart);
  assign("dueDate", input.dueDate === undefined ? undefined : input.dueDate);
  assign("inviteOnly", input.inviteOnly);
  if (input.propertyId && input.propertyId !== existing.propertyId) {
    data.property = { connect: { id: input.propertyId } };
    changed.push("property");
  }

  const job = await prisma.$transaction(async (tx) => {
    const updated = await tx.job.update({ where: { id: jobId }, data, include: { property: true } });
    if (changed.length > 0) {
      await logActivity(
        {
          entityType: "JOB",
          entityId: jobId,
          jobId,
          action: "job.updated",
          actorType: "ADMIN",
          actorAdminId: adminUserId,
          summary: `Job ${updated.jobNumber} updated (${changed.join(", ")}).`,
          metadata: { changed },
        },
        tx,
      );
    }
    return updated;
  });

  return job;
}

const ACTION_TARGET: Record<JobStatusAction, JobStatus> = {
  publish: "OPEN",
  close_bidding: "BIDDING_CLOSED",
  reopen: "OPEN",
  cancel: "CANCELLED",
  start: "IN_PROGRESS",
  complete: "COMPLETED",
};

export async function transitionJob(
  jobId: string,
  action: JobStatusAction,
  adminUserId: string,
  reason?: string,
) {
  const job = await prisma.job.findUnique({
    where: { id: jobId },
    include: { contract: { select: { status: true, externalId: true } } },
  });
  if (!job) throw new JobRuleError("job_not_found", "Job not found.");

  // Emergencies never carry bids, so the bidding transitions do not apply.
  if (job.priority === "EMERGENCY" && (action === "close_bidding" || action === "reopen")) {
    throw new JobRuleError(
      "not_a_bidding_job",
      "Emergency jobs are claimed rather than bid on — there is no bidding to open or close.",
    );
  }

  if (!ALLOWED_STATUS_ACTIONS[job.status].includes(action)) {
    throw new JobRuleError(
      "invalid_transition",
      `A ${job.status.replace("_", " ").toLowerCase()} job cannot be ${action.replace("_", " ")}ed.`,
    );
  }

  // Phase 7 rule, enforced from the start: a GC job whose contract has not at
  // least reached signature cannot move into work.
  if (action === "start" && job.contract && job.contract.status === "DRAFT") {
    throw new JobRuleError(
      "contract_not_ready",
      "This job cannot start until its contract reaches pending signature.",
    );
  }

  if (action === "publish" && job.bidDeadline && job.bidDeadline.getTime() <= Date.now()) {
    throw new JobRuleError("deadline_in_past", "Set a future bid deadline before publishing.");
  }

  const target = ACTION_TARGET[action];
  const now = new Date();

  return prisma.$transaction(async (tx) => {
    const updated = await tx.job.update({
      where: { id: jobId },
      data: {
        status: target,
        ...(action === "cancel" ? { cancelledAt: now } : {}),
        ...(action === "start" ? { startedAt: now } : {}),
        ...(action === "complete" ? { completedAt: now } : {}),
      },
    });

    await logActivity(
      {
        entityType: "JOB",
        entityId: jobId,
        jobId,
        action: `job.${action}`,
        fromStatus: job.status,
        toStatus: target,
        actorType: "ADMIN",
        actorAdminId: adminUserId,
        summary: reason
          ? `Job ${job.jobNumber} → ${target.replace("_", " ").toLowerCase()}: ${reason}`
          : `Job ${job.jobNumber} → ${target.replace("_", " ").toLowerCase()}.`,
      },
      tx,
    );

    return updated;
  });
}
