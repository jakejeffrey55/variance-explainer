import { z } from "zod";
import { EmergencyCategory, JobPriority, JobStatus, ServiceCategory } from "@prisma/client";

const optionalDate = z
  .union([z.string(), z.date(), z.null()])
  .optional()
  .transform((v) => {
    if (v === null || v === undefined || v === "") return null;
    const d = v instanceof Date ? v : new Date(v);
    return Number.isNaN(d.getTime()) ? null : d;
  });

export const jobBaseSchema = z.object({
  propertyId: z.string().min(1, "Choose a property."),
  unitNumber: z.string().trim().max(32).optional().nullable(),
  title: z.string().trim().min(4, "Give the job a title.").max(160),
  description: z.string().trim().min(10, "Describe the scope of work.").max(5000),
  category: z.enum(ServiceCategory),
  budgetMin: z.number({ error: "Enter a minimum budget." }).min(0).max(10_000_000),
  budgetMax: z.number({ error: "Enter a maximum budget." }).min(0).max(10_000_000),
  enforceBudgetCap: z.boolean().default(false),
  bidDeadline: optionalDate,
  scheduledStart: optionalDate,
  dueDate: optionalDate,
  priority: z.enum(JobPriority).default("STANDARD"),
  emergencyCategory: z.enum(EmergencyCategory).optional().nullable(),
  responseDeadlineMinutes: z.number().int().min(5).max(240).optional().nullable(),
  inviteOnly: z.boolean().default(false),
  status: z.enum(JobStatus).optional(),
});

export type JobInput = z.infer<typeof jobBaseSchema>;

/**
 * Cross-field rules. Emergency jobs skip bidding entirely, so a bid deadline or
 * an invite list on one would be meaningless — reject rather than silently ignore.
 */
function checkRules(value: Partial<JobInput>, ctx: z.RefinementCtx) {
  if (value.budgetMin !== undefined && value.budgetMax !== undefined && value.budgetMax < value.budgetMin) {
    ctx.addIssue({
      code: "custom",
      path: ["budgetMax"],
      message: "Maximum budget must be at least the minimum.",
    });
  }

  if (value.priority === "EMERGENCY") {
    if (!value.emergencyCategory) {
      ctx.addIssue({ code: "custom", path: ["emergencyCategory"], message: "Choose an emergency category." });
    }
    if (value.bidDeadline) {
      ctx.addIssue({
        code: "custom",
        path: ["bidDeadline"],
        message: "Emergency jobs skip bidding — they are claimed, not bid on.",
      });
    }
    if (value.inviteOnly) {
      ctx.addIssue({
        code: "custom",
        path: ["inviteOnly"],
        message: "Emergency jobs go to every eligible vendor and cannot be invite-only.",
      });
    }
  } else if (value.priority === "STANDARD" && value.emergencyCategory) {
    ctx.addIssue({
      code: "custom",
      path: ["emergencyCategory"],
      message: "Only emergency jobs carry an emergency category.",
    });
  }
}

export const jobCreateSchema = jobBaseSchema.superRefine(checkRules);
export const jobUpdateSchema = jobBaseSchema.partial().superRefine(checkRules);

export const jobStatusActionSchema = z.object({
  action: z.enum(["publish", "close_bidding", "reopen", "cancel", "start", "complete"]),
  reason: z.string().trim().max(500).optional(),
});

export type JobStatusAction = z.infer<typeof jobStatusActionSchema>["action"];

/** Which manual transitions an admin may drive, by current status. */
export const ALLOWED_STATUS_ACTIONS: Record<JobStatus, JobStatusAction[]> = {
  DRAFT: ["publish", "cancel"],
  OPEN: ["close_bidding", "cancel"],
  BIDDING_CLOSED: ["reopen", "cancel"],
  AWAITING_APPROVAL: ["cancel"],
  AWARDED: ["start", "cancel"],
  IN_PROGRESS: ["complete", "cancel"],
  COMPLETED: [],
  CANCELLED: [],
};
