import { NextResponse } from "next/server";
import { apiError, withAdmin } from "@/lib/auth/route";
import { JobRuleError, transitionJob } from "@/lib/services/job";
import { jobStatusActionSchema } from "@/lib/validation/job";

export const dynamic = "force-dynamic";

export const POST = withAdmin(async (req, { params, actor }) => {
  const { action, reason } = jobStatusActionSchema.parse(await req.json());
  try {
    const job = await transitionJob(params.id, action, actor.adminUserId, reason);
    return NextResponse.json({ job });
  } catch (err) {
    if (err instanceof JobRuleError) {
      return apiError(err.code === "job_not_found" ? 404 : 422, err.code, err.message);
    }
    throw err;
  }
});
