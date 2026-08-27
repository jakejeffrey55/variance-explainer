import Link from "next/link";
import { notFound } from "next/navigation";
import type { Metadata } from "next";
import { ArrowLeft } from "lucide-react";
import { PageHeader } from "@/components/admin/admin-shell";
import { JobForm } from "@/components/admin/job-form";
import type { JobFormValues } from "@/lib/forms/job";
import { Button } from "@/components/ui/button";
import { requireAdmin } from "@/lib/auth/session";
import { prisma } from "@/lib/db";

export const metadata: Metadata = { title: "Edit job" };
export const dynamic = "force-dynamic";

/** datetime-local and date inputs want local wall-clock strings, not ISO. */
function toLocalInput(date: Date | null, withTime: boolean) {
  if (!date) return "";
  const offset = date.getTimezoneOffset() * 60000;
  const local = new Date(date.getTime() - offset).toISOString();
  return withTime ? local.slice(0, 16) : local.slice(0, 10);
}

export default async function EditJobPage({ params }: { params: { id: string } }) {
  await requireAdmin();

  const [job, properties] = await Promise.all([
    prisma.job.findUnique({ where: { id: params.id } }),
    prisma.property.findMany({
      orderBy: { name: "asc" },
      select: { id: true, name: true, city: true, state: true, isActive: true },
    }),
  ]);
  if (!job) notFound();

  const initial: JobFormValues = {
    id: job.id,
    propertyId: job.propertyId,
    unitNumber: job.unitNumber ?? "",
    title: job.title,
    description: job.description,
    category: job.category,
    budgetMin: job.budgetMin.toString(),
    budgetMax: job.budgetMax.toString(),
    enforceBudgetCap: job.enforceBudgetCap,
    bidDeadline: toLocalInput(job.bidDeadline, true),
    scheduledStart: toLocalInput(job.scheduledStart, false),
    dueDate: toLocalInput(job.dueDate, false),
    priority: job.priority,
    emergencyCategory: job.emergencyCategory ?? "",
    responseDeadlineMinutes: String(job.responseDeadlineMinutes ?? 15),
    inviteOnly: job.inviteOnly,
  };

  return (
    <>
      <Button asChild variant="ghost" size="sm" className="mb-2 -ml-2">
        <Link href={`/admin/jobs/${job.id}`}>
          <ArrowLeft /> {job.jobNumber}
        </Link>
      </Button>
      <PageHeader title={`Edit ${job.jobNumber}`} description={job.title} />
      <JobForm properties={properties} initial={initial} mode="edit" />
    </>
  );
}
