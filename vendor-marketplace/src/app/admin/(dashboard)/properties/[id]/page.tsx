import Link from "next/link";
import { notFound } from "next/navigation";
import type { Metadata } from "next";
import { ArrowLeft, Building2, ClipboardList, MapPin, Users } from "lucide-react";
import { PageHeader } from "@/components/admin/admin-shell";
import { EmptyState } from "@/components/empty-state";
import { EmergencyBadge, JobStatusBadge } from "@/components/status-badges";
import { Badge } from "@/components/ui/badge";
import { Button } from "@/components/ui/button";
import { Card, CardContent, CardDescription, CardHeader, CardTitle } from "@/components/ui/card";
import { Table, TableBody, TableCell, TableHead, TableHeader, TableRow } from "@/components/ui/table";
import { requireAdmin } from "@/lib/auth/session";
import { prisma } from "@/lib/db";
import { formatCurrency, formatDate } from "@/lib/utils";

export const metadata: Metadata = { title: "Property" };
export const dynamic = "force-dynamic";

export default async function PropertyDetailPage({ params }: { params: { id: string } }) {
  await requireAdmin();

  const property = await prisma.property.findUnique({
    where: { id: params.id },
    include: {
      jobs: { orderBy: { createdAt: "desc" }, take: 25 },
      rentRolls: { orderBy: { unitNumber: "asc" } },
    },
  });
  if (!property) notFound();

  return (
    <>
      <Button asChild variant="ghost" size="sm" className="mb-2 -ml-2">
        <Link href="/admin/properties">
          <ArrowLeft /> Properties
        </Link>
      </Button>

      <PageHeader
        title={property.name}
        description={`${property.addressLine1}, ${property.city}, ${property.state} ${property.postalCode}`}
      >
        <Button asChild>
          <Link href={`/admin/jobs/new?propertyId=${property.id}`}>New job here</Link>
        </Button>
      </PageHeader>

      <div className="grid gap-4 sm:grid-cols-2 lg:grid-cols-4">
        <Card>
          <CardContent className="space-y-1 p-4">
            <p className="text-xs uppercase tracking-wide text-muted-foreground">Units</p>
            <p className="text-2xl font-semibold tabular-nums">{property.unitCount ?? "—"}</p>
          </CardContent>
        </Card>
        <Card>
          <CardContent className="space-y-1 p-4">
            <p className="text-xs uppercase tracking-wide text-muted-foreground">Jobs</p>
            <p className="text-2xl font-semibold tabular-nums">{property.jobs.length}</p>
          </CardContent>
        </Card>
        <Card>
          <CardContent className="space-y-1 p-4">
            <p className="text-xs uppercase tracking-wide text-muted-foreground">Source</p>
            <p className="text-sm font-medium">{property.source}</p>
            <p className="text-xs text-muted-foreground">
              {property.lastSyncedAt ? `Synced ${formatDate(property.lastSyncedAt)}` : "Entered manually"}
            </p>
          </CardContent>
        </Card>
        <Card>
          <CardContent className="space-y-1 p-4">
            <p className="text-xs uppercase tracking-wide text-muted-foreground">Coordinates</p>
            <p className="flex items-center gap-1 text-sm font-medium tabular-nums">
              <MapPin className="h-3.5 w-3.5 text-muted-foreground" />
              {property.latitude.toFixed(4)}, {property.longitude.toFixed(4)}
            </p>
            <p className="text-xs text-muted-foreground">Drives vendor radius matching</p>
          </CardContent>
        </Card>
      </div>

      <div className="mt-6 grid min-w-0 gap-6 lg:grid-cols-3">
        <Card className="min-w-0 lg:col-span-2">
          <CardHeader>
            <CardTitle>Jobs at this property</CardTitle>
            <CardDescription>Most recent 25.</CardDescription>
          </CardHeader>
          <CardContent className="p-0">
            {property.jobs.length === 0 ? (
              <EmptyState
                icon={ClipboardList}
                title="No jobs here yet"
                description="Post a make-ready or contracting job for this property."
                action={
                  <Button asChild size="sm">
                    <Link href={`/admin/jobs/new?propertyId=${property.id}`}>New job</Link>
                  </Button>
                }
                className="m-5 border-0"
              />
            ) : (
              <Table>
                <TableHeader>
                  <TableRow>
                    <TableHead>Job</TableHead>
                    <TableHead>Budget</TableHead>
                    <TableHead>Status</TableHead>
                  </TableRow>
                </TableHeader>
                <TableBody>
                  {property.jobs.map((job) => (
                    <TableRow key={job.id}>
                      <TableCell>
                        <div className="flex flex-col">
                          <Link href={`/admin/jobs/${job.id}`} className="font-medium hover:underline">
                            {job.jobNumber} · {job.title}
                          </Link>
                          <span className="text-xs text-muted-foreground">
                            {job.unitNumber ? `Unit ${job.unitNumber} · ` : ""}
                            {formatDate(job.createdAt)}
                          </span>
                        </div>
                      </TableCell>
                      <TableCell className="whitespace-nowrap tabular-nums">
                        {formatCurrency(job.budgetMin)}–{formatCurrency(job.budgetMax)}
                      </TableCell>
                      <TableCell>
                        <div className="flex items-center gap-2">
                          {job.priority === "EMERGENCY" && <EmergencyBadge />}
                          <JobStatusBadge status={job.status} />
                        </div>
                      </TableCell>
                    </TableRow>
                  ))}
                </TableBody>
              </Table>
            )}
          </CardContent>
        </Card>

        <div className="min-w-0 space-y-6">
          <Card>
            <CardHeader>
              <CardTitle className="flex items-center gap-2">
                <Users className="h-4 w-4" /> Property manager
              </CardTitle>
            </CardHeader>
            <CardContent className="space-y-1 text-sm">
              <p className="font-medium">{property.propertyManagerName ?? "Not recorded"}</p>
              <p className="text-muted-foreground">{property.propertyManagerEmail ?? "—"}</p>
              <p className="text-muted-foreground">{property.propertyManagerPhone ?? "—"}</p>
            </CardContent>
          </Card>

          <Card>
            <CardHeader>
              <CardTitle className="flex items-center gap-2">
                <Building2 className="h-4 w-4" /> Rent roll
              </CardTitle>
              <CardDescription>Imported separately; informs due-date suggestions only.</CardDescription>
            </CardHeader>
            <CardContent className="p-0">
              {property.rentRolls.length === 0 ? (
                <p className="px-5 pb-5 text-sm text-muted-foreground">No rent roll imported for this property.</p>
              ) : (
                <Table>
                  <TableHeader>
                    <TableRow>
                      <TableHead>Unit</TableHead>
                      <TableHead>Status</TableHead>
                      <TableHead>Move-out</TableHead>
                    </TableRow>
                  </TableHeader>
                  <TableBody>
                    {property.rentRolls.map((unit) => (
                      <TableRow key={unit.id}>
                        <TableCell className="font-medium">{unit.unitNumber}</TableCell>
                        <TableCell>
                          <Badge variant="secondary">{unit.status ?? "—"}</Badge>
                        </TableCell>
                        <TableCell className="whitespace-nowrap text-muted-foreground">
                          {formatDate(unit.moveOutDate)}
                        </TableCell>
                      </TableRow>
                    ))}
                  </TableBody>
                </Table>
              )}
            </CardContent>
          </Card>
        </div>
      </div>
    </>
  );
}
