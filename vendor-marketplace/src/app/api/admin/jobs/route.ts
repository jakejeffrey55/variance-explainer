import { NextResponse } from "next/server";
import { Prisma, type JobStatus } from "@prisma/client";
import { apiError, withAdmin } from "@/lib/auth/route";
import { prisma } from "@/lib/db";
import { JobRuleError, createJob } from "@/lib/services/job";
import { jobCreateSchema } from "@/lib/validation/job";

export const dynamic = "force-dynamic";

export const GET = withAdmin(async (req) => {
  const url = new URL(req.url);
  const status = url.searchParams.get("status");
  const propertyId = url.searchParams.get("propertyId");
  const priority = url.searchParams.get("priority");
  const q = url.searchParams.get("q")?.trim();

  const where: Prisma.JobWhereInput = {
    ...(status && status !== "all" ? { status: status as JobStatus } : {}),
    ...(propertyId && propertyId !== "all" ? { propertyId } : {}),
    ...(priority === "EMERGENCY" || priority === "STANDARD" ? { priority } : {}),
    ...(q
      ? {
          OR: [
            { title: { contains: q, mode: "insensitive" } },
            { jobNumber: { contains: q, mode: "insensitive" } },
            { unitNumber: { contains: q, mode: "insensitive" } },
            { property: { name: { contains: q, mode: "insensitive" } } },
          ],
        }
      : {}),
  };

  const jobs = await prisma.job.findMany({
    where,
    orderBy: [{ priority: "desc" }, { createdAt: "desc" }],
    include: {
      property: { select: { id: true, name: true, city: true, state: true } },
      _count: { select: { bids: true } },
    },
    take: 200,
  });

  return NextResponse.json({ jobs });
});

export const POST = withAdmin(async (req, { actor }) => {
  const input = jobCreateSchema.parse(await req.json());
  try {
    const job = await createJob(input, actor.adminUserId);
    return NextResponse.json({ job }, { status: 201 });
  } catch (err) {
    if (err instanceof JobRuleError) return apiError(422, err.code, err.message);
    throw err;
  }
});
