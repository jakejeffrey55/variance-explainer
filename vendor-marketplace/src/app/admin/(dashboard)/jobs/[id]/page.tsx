import Link from "next/link";
import { notFound } from "next/navigation";
import type { Metadata } from "next";
import { ArrowLeft, Building2, CalendarClock, FileSignature, Flag, Gavel, Siren, Timer } from "lucide-react";
import { ActivityTimeline } from "@/components/admin/activity-timeline";
import { PageHeader } from "@/components/admin/admin-shell";
import { JobStatusActions } from "@/components/admin/job-status-actions";
import { EmptyState } from "@/components/empty-state";
import {
  BidStatusBadge,
  CategoryBadge,
  ContractStatusBadge,
  EmergencyBadge,
  JobStatusBadge,
} from "@/components/status-badges";
import { Alert, AlertDescription, AlertTitle } from "@/components/ui/alert";
import { Badge } from "@/components/ui/badge";
import { Button } from "@/components/ui/button";
import { Card, CardContent, CardDescription, CardHeader, CardTitle } from "@/components/ui/card";
import { Separator } from "@/components/ui/separator";
import { Table, TableBody, TableCell, TableHead, TableHeader, TableRow } from "@/components/ui/table";
import { requireAdmin } from "@/lib/auth/session";
import { prisma } from "@/lib/db";
import { formatCurrency, formatDate, formatRelative, titleCase } from "@/lib/utils";

export const metadata: Metadata = { title: "Job" };
export const dynamic = "force-dynamic";

export default async function JobDetailPage({ params }: { params: { id: string } }) {
  await requireAdmin();

  const job = await prisma.job.findUnique({
    where: { id: params.id },
    include: {
      property: true,
      createdBy: { select: { name: true } },
      // Admin-only: every bid on the job. Vendors never receive this shape.
      bids: {
        orderBy: [{ status: "asc" }, { amount: "asc" }],
        include: { vendor: { select: { id: true, companyName: true, trustScore: true } } },
      },
      contract: true,
      requisition: true,
      approvalFlags: true,
      invitations: { include: { vendor: { select: { companyName: true } } } },
      claimedBy: { select: { companyName: true, phone: true } },
      awardedVendor: { select: { companyName: true } },
      activityLogs: {
        orderBy: { createdAt: "desc" },
        include: { actorAdmin: { select: { name: true } }, actorVendor: { select: { companyName: true } } },
      },
    },
  });
  if (!job) notFound();

  const liveBids = job.bids.filter((b) => b.status !== "WITHDRAWN");
  const averageBid =
    liveBids.length > 0 ? liveBids.reduce((sum, b) => sum + Number(b.amount), 0) / liveBids.length : null;

  const isEmergency = job.priority === "EMERGENCY";
  const overdue =
    isEmergency &&
    !job.claimedAt &&
    job.dispatchedAt &&
    Date.now() - job.dispatchedAt.getTime() > (job.responseDeadlineMinutes ?? 15) * 60000;

  const facts = [
    { label: "Property", value: job.property.name, href: `/admin/properties/${job.propertyId}` },
    { label: "Unit / area", value: job.unitNumber ?? "—" },
    { label: "Budget", value: `${formatCurrency(job.budgetMin)}–${formatCurrency(job.budgetMax)}` },
    {
      label: "Budget cap",
      value: job.enforceBudgetCap ? "Enforced — bids above max rejected" : "Not enforced",
    },
    {
      label: isEmergency ? "Dispatched" : "Bid deadline",
      value: isEmergency ? formatDate(job.dispatchedAt, true) : formatDate(job.bidDeadline, true),
    },
    { label: "Scheduled start", value: formatDate(job.scheduledStart) },
    { label: "Due date", value: formatDate(job.dueDate) },
    { label: "Created by", value: `${job.createdBy.name} · ${formatDate(job.createdAt)}` },
  ];

  return (
    <>
      <Button asChild variant="ghost" size="sm" className="mb-2 -ml-2">
        <Link href="/admin/jobs">
          <ArrowLeft /> Jobs
        </Link>
      </Button>

      <PageHeader title={job.title} description={`${job.jobNumber} · ${job.property.name}`}>
        <JobStatusActions jobId={job.id} status={job.status} priority={job.priority} />
        {job.status !== "COMPLETED" && job.status !== "CANCELLED" && (
          <Button asChild variant="outline" size="sm">
            <Link href={`/admin/jobs/${job.id}/edit`}>Edit</Link>
          </Button>
        )}
      </PageHeader>

      <div className="mb-4 flex flex-wrap items-center gap-2">
        <JobStatusBadge status={job.status} />
        {isEmergency && <EmergencyBadge />}
        {isEmergency && job.emergencyCategory && (
          <Badge variant="emergency">{titleCase(job.emergencyCategory)}</Badge>
        )}
        <CategoryBadge category={job.category} />
        {job.inviteOnly && <Badge variant="secondary">Invite only</Badge>}
      </div>

      {overdue && (
        <Alert variant="emergency" className="mb-6">
          <Siren />
          <AlertTitle>Unclaimed past its response window</AlertTitle>
          <AlertDescription>
            Dispatched {formatRelative(job.dispatchedAt)} with a {job.responseDeadlineMinutes}-minute window and no
            vendor has claimed it. Assign someone directly or widen the search.
          </AlertDescription>
        </Alert>
      )}

      {job.approvalFlags.length > 0 && (
        <Alert variant="warning" className="mb-6">
          <Flag />
          <AlertTitle>
            {job.approvalFlags.length} approval flag{job.approvalFlags.length === 1 ? "" : "s"} on this job
          </AlertTitle>
          <AlertDescription>
            <ul className="mt-1 space-y-1">
              {job.approvalFlags.map((flag) => (
                <li key={flag.id}>
                  {flag.type === "ABOVE_AVERAGE_THRESHOLD"
                    ? `Approved ${formatCurrency(flag.approvedAmount)} — ${flag.deltaPct?.toFixed(1)}% above the ${formatCurrency(flag.averageBidAmount)} average of ${flag.bidCount} bids.`
                    : `Approved ${formatCurrency(flag.approvedAmount)} — above the ${formatCurrency(flag.budgetMax)} budget maximum.`}
                  {flag.note ? ` ${flag.note}` : ""}
                </li>
              ))}
            </ul>
          </AlertDescription>
        </Alert>
      )}

      <div className="grid min-w-0 gap-6 lg:grid-cols-3">
        <div className="min-w-0 space-y-6 lg:col-span-2">
          <Card>
            <CardHeader>
              <CardTitle>Scope of work</CardTitle>
            </CardHeader>
            <CardContent className="space-y-4">
              <p className="whitespace-pre-wrap text-sm leading-relaxed">{job.description}</p>
              <Separator />
              <dl className="grid gap-x-6 gap-y-3 sm:grid-cols-2">
                {facts.map((fact) => (
                  <div key={fact.label}>
                    <dt className="text-xs uppercase tracking-wide text-muted-foreground">{fact.label}</dt>
                    <dd className="text-sm font-medium">
                      {fact.href ? (
                        <Link href={fact.href} className="hover:underline">
                          {fact.value}
                        </Link>
                      ) : (
                        fact.value
                      )}
                    </dd>
                  </div>
                ))}
              </dl>
            </CardContent>
          </Card>

          <Card>
            <CardHeader>
              <CardTitle className="flex items-center gap-2">
                <Gavel className="h-4 w-4" /> Bids
              </CardTitle>
              <CardDescription>
                Admin-only view. Vendors only ever see their own bid and its status.
                {averageBid !== null && ` Average of ${liveBids.length} live bids: ${formatCurrency(averageBid)}.`}
              </CardDescription>
            </CardHeader>
            <CardContent className="p-0">
              {job.bids.length === 0 ? (
                <EmptyState
                  icon={Gavel}
                  title={isEmergency ? "Emergency jobs are claimed, not bid on" : "No bids yet"}
                  description={
                    isEmergency
                      ? "The first eligible vendor to claim this job is assigned instantly."
                      : "Matched vendors will appear here as they submit."
                  }
                  className="m-5 border-0"
                />
              ) : (
                <Table>
                  <TableHeader>
                    <TableRow>
                      <TableHead>Vendor</TableHead>
                      <TableHead className="text-right">Amount</TableHead>
                      <TableHead className="text-right">vs. average</TableHead>
                      <TableHead>Submitted</TableHead>
                      <TableHead>Status</TableHead>
                    </TableRow>
                  </TableHeader>
                  <TableBody>
                    {job.bids.map((bid) => {
                      const delta =
                        averageBid && bid.status !== "WITHDRAWN"
                          ? ((Number(bid.amount) - averageBid) / averageBid) * 100
                          : null;
                      return (
                        <TableRow key={bid.id}>
                          <TableCell>
                            <div className="flex flex-col">
                              <span className="font-medium">{bid.vendor.companyName}</span>
                              {bid.vendor.trustScore !== null && (
                                <span className="text-xs text-muted-foreground">
                                  Trust {bid.vendor.trustScore.toFixed(0)}
                                </span>
                              )}
                            </div>
                          </TableCell>
                          <TableCell className="text-right font-medium tabular-nums">
                            {formatCurrency(bid.amount)}
                          </TableCell>
                          <TableCell className="text-right tabular-nums text-muted-foreground">
                            {delta === null ? "—" : `${delta > 0 ? "+" : ""}${delta.toFixed(1)}%`}
                          </TableCell>
                          <TableCell className="whitespace-nowrap text-muted-foreground">
                            {formatDate(bid.submittedAt)}
                          </TableCell>
                          <TableCell>
                            <BidStatusBadge status={bid.status} />
                          </TableCell>
                        </TableRow>
                      );
                    })}
                  </TableBody>
                </Table>
              )}
            </CardContent>
          </Card>

          <Card>
            <CardHeader>
              <CardTitle>Timeline</CardTitle>
              <CardDescription>Every state change on this job.</CardDescription>
            </CardHeader>
            <CardContent>
              <ActivityTimeline
                entries={job.activityLogs.map((entry) => ({
                  id: entry.id,
                  action: entry.action,
                  summary: entry.summary,
                  fromStatus: entry.fromStatus,
                  toStatus: entry.toStatus,
                  actorLabel: entry.actorAdmin?.name ?? entry.actorVendor?.companyName ?? "System",
                  createdAt: entry.createdAt,
                }))}
              />
            </CardContent>
          </Card>
        </div>

        <div className="min-w-0 space-y-6">
          {isEmergency && (
            <Card className="border-emergency-border">
              <CardHeader>
                <CardTitle className="flex items-center gap-2">
                  <Timer className="h-4 w-4 text-emergency" /> Dispatch
                </CardTitle>
              </CardHeader>
              <CardContent className="space-y-2 text-sm">
                <Row label="Response window" value={`${job.responseDeadlineMinutes ?? 15} minutes`} />
                <Row label="Dispatched" value={formatDate(job.dispatchedAt, true)} />
                <Row
                  label="Claimed"
                  value={
                    job.claimedAt && job.dispatchedAt
                      ? `${job.claimedBy?.companyName} · ${Math.round((job.claimedAt.getTime() - job.dispatchedAt.getTime()) / 60000)} min`
                      : "Not yet claimed"
                  }
                />
                <Row
                  label="On site"
                  value={
                    job.onSiteAt && job.dispatchedAt
                      ? `${Math.round((job.onSiteAt.getTime() - job.dispatchedAt.getTime()) / 60000)} min after dispatch`
                      : "Not yet on site"
                  }
                />
                {job.escalatedAt && (
                  <Row
                    label="Escalated"
                    value={`${formatRelative(job.escalatedAt)} · radius ${job.escalationRadiusMiles ?? "—"} mi`}
                  />
                )}
              </CardContent>
            </Card>
          )}

          <Card>
            <CardHeader>
              <CardTitle className="flex items-center gap-2">
                <Building2 className="h-4 w-4" /> Property
              </CardTitle>
            </CardHeader>
            <CardContent className="space-y-1 text-sm">
              <Link href={`/admin/properties/${job.propertyId}`} className="font-medium hover:underline">
                {job.property.name}
              </Link>
              <p className="text-muted-foreground">
                {job.property.addressLine1}
                <br />
                {job.property.city}, {job.property.state} {job.property.postalCode}
              </p>
              <p className="pt-1 text-xs text-muted-foreground">
                {job.property.propertyManagerName ?? "No manager recorded"}
              </p>
            </CardContent>
          </Card>

          {(job.contract || job.requisition) && (
            <Card>
              <CardHeader>
                <CardTitle className="flex items-center gap-2">
                  <FileSignature className="h-4 w-4" /> Downstream
                </CardTitle>
                <CardDescription>Procurement and contracting records.</CardDescription>
              </CardHeader>
              <CardContent className="space-y-3 text-sm">
                {job.requisition && (
                  <div className="space-y-1">
                    <p className="text-xs uppercase tracking-wide text-muted-foreground">Requisition</p>
                    <p className="font-mono text-xs">{job.requisition.externalId ?? "—"}</p>
                    <p className="text-muted-foreground">
                      {titleCase(job.requisition.status)} · {formatCurrency(job.requisition.amount)} ·{" "}
                      {job.requisition.providerKey}
                    </p>
                  </div>
                )}
                {job.contract && (
                  <div className="space-y-1">
                    <p className="text-xs uppercase tracking-wide text-muted-foreground">Contract</p>
                    <p className="font-mono text-xs">{job.contract.externalId ?? "—"}</p>
                    <div className="flex items-center gap-2">
                      <ContractStatusBadge status={job.contract.status} />
                      <span className="text-muted-foreground">{formatCurrency(job.contract.amount)}</span>
                    </div>
                  </div>
                )}
              </CardContent>
            </Card>
          )}

          {job.invitations.length > 0 && (
            <Card>
              <CardHeader>
                <CardTitle>Invited vendors</CardTitle>
              </CardHeader>
              <CardContent className="space-y-2 text-sm">
                {job.invitations.map((invite) => (
                  <div key={invite.id} className="flex items-center justify-between gap-2">
                    <span>{invite.vendor.companyName}</span>
                    <Badge variant={invite.viewedAt ? "secondary" : "muted"}>
                      {invite.viewedAt ? "viewed" : "sent"}
                    </Badge>
                  </div>
                ))}
              </CardContent>
            </Card>
          )}

          <Card>
            <CardHeader>
              <CardTitle className="flex items-center gap-2">
                <CalendarClock className="h-4 w-4" /> Award
              </CardTitle>
            </CardHeader>
            <CardContent className="space-y-2 text-sm">
              <Row label="Awarded to" value={job.awardedVendor?.companyName ?? "Not awarded"} />
              <Row label="Awarded" value={formatDate(job.awardedAt, true)} />
              <Row label="Started" value={formatDate(job.startedAt, true)} />
              <Row label="Completed" value={formatDate(job.completedAt, true)} />
            </CardContent>
          </Card>
        </div>
      </div>
    </>
  );
}

function Row({ label, value }: { label: string; value: string }) {
  return (
    <div className="flex items-baseline justify-between gap-3">
      <span className="text-xs uppercase tracking-wide text-muted-foreground">{label}</span>
      <span className="text-right font-medium">{value}</span>
    </div>
  );
}
