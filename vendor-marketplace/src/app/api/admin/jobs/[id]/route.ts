import { NextResponse } from "next/server";
import { apiError, withAdmin } from "@/lib/auth/route";
import { prisma } from "@/lib/db";
import { JobRuleError, updateJob } from "@/lib/services/job";
import { jobUpdateSchema } from "@/lib/validation/job";

export const dynamic = "force-dynamic";

export const GET = withAdmin(async (_req, { params }) => {
  const job = await prisma.job.findUnique({
    where: { id: params.id },
    include: {
      property: true,
      createdBy: { select: { name: true } },
      // Admins see every bid on a job; the comparison table is built from this.
      bids: {
        orderBy: { amount: "asc" },
        include: {
          vendor: {
            select: {
              id: true,
              companyName: true,
              trustScore: true,
              complianceStatus: true,
              emergencyEligible: true,
            },
          },
        },
      },
      contract: true,
      requisition: true,
      approvalFlags: true,
      invitations: { include: { vendor: { select: { id: true, companyName: true } } } },
      activityLogs: {
        orderBy: { createdAt: "desc" },
        include: {
          actorAdmin: { select: { name: true } },
          actorVendor: { select: { companyName: true } },
        },
      },
      claimedBy: { select: { id: true, companyName: true, phone: true } },
      awardedVendor: { select: { id: true, companyName: true } },
    },
  });
  if (!job) return apiError(404, "not_found", "Job not found.");
  return NextResponse.json({ job });
});

export const PATCH = withAdmin(async (req, { params, actor }) => {
  const input = jobUpdateSchema.parse(await req.json());
  try {
    const job = await updateJob(params.id, input, actor.adminUserId);
    return NextResponse.json({ job });
  } catch (err) {
    if (err instanceof JobRuleError) {
      return apiError(err.code === "job_not_found" ? 404 : 422, err.code, err.message);
    }
    throw err;
  }
});
