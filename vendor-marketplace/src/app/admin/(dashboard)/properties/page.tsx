import type { Metadata } from "next";
import { PageHeader } from "@/components/admin/admin-shell";
import { PropertiesClient, type PropertyRow } from "@/components/admin/properties-client";
import { Button } from "@/components/ui/button";
import Link from "next/link";
import { requireAdmin } from "@/lib/auth/session";
import { prisma } from "@/lib/db";

export const metadata: Metadata = { title: "Properties" };
export const dynamic = "force-dynamic";

export default async function PropertiesPage() {
  await requireAdmin();

  const properties = await prisma.property.findMany({
    orderBy: { name: "asc" },
    include: { _count: { select: { jobs: true, rentRolls: true } } },
  });

  const rows: PropertyRow[] = properties.map((p) => ({
    id: p.id,
    externalId: p.externalId,
    name: p.name,
    addressLine1: p.addressLine1,
    addressLine2: p.addressLine2,
    city: p.city,
    state: p.state,
    postalCode: p.postalCode,
    latitude: p.latitude,
    longitude: p.longitude,
    unitCount: p.unitCount,
    propertyManagerName: p.propertyManagerName,
    propertyManagerEmail: p.propertyManagerEmail,
    propertyManagerPhone: p.propertyManagerPhone,
    isActive: p.isActive,
    source: p.source,
    jobCount: p._count.jobs,
    unitRollCount: p._count.rentRolls,
    lastSyncedAt: p.lastSyncedAt ? p.lastSyncedAt.toISOString() : null,
  }));

  return (
    <>
      <PageHeader
        title="Properties"
        description={`${rows.filter((r) => r.isActive).length} active of ${rows.length} total.`}
      >
        <Button asChild variant="outline">
          <Link href="/admin/import">Import CSV</Link>
        </Button>
      </PageHeader>
      <PropertiesClient properties={rows} />
    </>
  );
}
