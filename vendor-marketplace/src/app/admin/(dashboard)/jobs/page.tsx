import Link from "next/link";
import type { Metadata } from "next";
import { PageHeader } from "@/components/admin/admin-shell";
import { JobsClient, type JobRow } from "@/components/admin/jobs-client";
import { Button } from "@/components/ui/button";
import { requireAdmin } from "@/lib/auth/session";
import { prisma } from "@/lib/db";

export const metadata: Metadata = { title: "Jobs" };
export const dynamic = "force-dynamic";

export default async function JobsPage({ searchParams }: { searchParams: { status?: string } }) {
  await requireAdmin();

  const [jobs, properties] = await Promise.all([
    prisma.job.findMany({
      orderBy: [{ createdAt: "desc" }],
      include: {
        property: { select: { id: true, name: true } },
        _count: { select: { bids: true } },
      },
      take: 300,
    }),
    prisma.property.findMany({ orderBy: { name: "asc" }, select: { id: true, name: true } }),
  ]);

  const rows: JobRow[] = jobs.map((job) => ({
    id: job.id,
    jobNumber: job.jobNumber,
    title: job.title,
    status: job.status,
    priority: job.priority,
    emergencyCategory: job.emergencyCategory,
    category: job.category,
    unitNumber: job.unitNumber,
    propertyName: job.property.name,
    propertyId: job.property.id,
    budgetMin: job.budgetMin.toString(),
    budgetMax: job.budgetMax.toString(),
    enforceBudgetCap: job.enforceBudgetCap,
    bidDeadline: job.bidDeadline?.toISOString() ?? null,
    dispatchedAt: job.dispatchedAt?.toISOString() ?? null,
    claimedAt: job.claimedAt?.toISOString() ?? null,
    responseDeadlineMinutes: job.responseDeadlineMinutes,
    bidCount: job._count.bids,
    createdAt: job.createdAt.toISOString(),
  }));

  return (
    <>
      <PageHeader title="Jobs" description={`${rows.length} job${rows.length === 1 ? "" : "s"} across all properties.`}>
        <Button asChild>
          <Link href="/admin/jobs/new">New job</Link>
        </Button>
      </PageHeader>
      <JobsClient jobs={rows} properties={properties} initialStatus={searchParams.status ?? null} />
    </>
  );
}
