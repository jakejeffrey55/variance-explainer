import Link from "next/link";
import type { Metadata } from "next";
import { ArrowLeft } from "lucide-react";
import { PageHeader } from "@/components/admin/admin-shell";
import { JobForm } from "@/components/admin/job-form";
import { blankJob } from "@/lib/forms/job";
import { EmptyState } from "@/components/empty-state";
import { Building2 } from "lucide-react";
import { Button } from "@/components/ui/button";
import { requireAdmin } from "@/lib/auth/session";
import { prisma } from "@/lib/db";

export const metadata: Metadata = { title: "New job" };
export const dynamic = "force-dynamic";

export default async function NewJobPage({ searchParams }: { searchParams: { propertyId?: string } }) {
  await requireAdmin();

  const [properties, settings] = await Promise.all([
    prisma.property.findMany({
      orderBy: { name: "asc" },
      select: { id: true, name: true, city: true, state: true, isActive: true },
    }),
    prisma.approvalSettings.findUnique({ where: { id: "default" } }),
  ]);

  const activeProperties = properties.filter((p) => p.isActive);

  return (
    <>
      <Button asChild variant="ghost" size="sm" className="mb-2 -ml-2">
        <Link href="/admin/jobs">
          <ArrowLeft /> Jobs
        </Link>
      </Button>

      <PageHeader title="New job" description="Post make-ready or contracting work, or dispatch an emergency." />

      {activeProperties.length === 0 ? (
        <EmptyState
          icon={Building2}
          title="Add a property first"
          description="Jobs are always posted against a property — import a CSV or add one by hand."
          action={
            <Button asChild size="sm">
              <Link href="/admin/import">Import properties</Link>
            </Button>
          }
        />
      ) : (
        <JobForm
          properties={properties}
          initial={blankJob(searchParams.propertyId ?? "")}
          mode="create"
          defaultResponseMinutes={settings?.emergencyResponseMinutes ?? 15}
        />
      )}
    </>
  );
}
