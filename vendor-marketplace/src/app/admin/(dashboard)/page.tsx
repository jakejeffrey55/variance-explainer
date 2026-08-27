import Link from "next/link";
import type { Metadata } from "next";
import {
  AlertTriangle,
  ArrowRight,
  Building2,
  ClipboardList,
  FileSignature,
  Flag,
  Gavel,
  Plug,
  ShieldAlert,
  Siren,
  UserCheck,
} from "lucide-react";
import { PageHeader } from "@/components/admin/admin-shell";
import { EmptyState } from "@/components/empty-state";
import { EmergencyBadge, JobStatusBadge } from "@/components/status-badges";
import { Alert, AlertDescription, AlertTitle } from "@/components/ui/alert";
import { Badge } from "@/components/ui/badge";
import { Button } from "@/components/ui/button";
import { Card, CardContent, CardDescription, CardHeader, CardTitle } from "@/components/ui/card";
import { requireAdmin } from "@/lib/auth/session";
import { prisma } from "@/lib/db";
import { listPropertyProviderStatuses } from "@/lib/integrations/property";
import { formatCurrency, formatDate, formatRelative } from "@/lib/utils";

export const metadata: Metadata = { title: "Dashboard" };
export const dynamic = "force-dynamic";

export default async function AdminDashboardPage() {
  const actor = await requireAdmin();
  const now = new Date();

  const settings = await prisma.approvalSettings.findUnique({ where: { id: "default" } });
  const complianceWindowDays = settings?.complianceExpiryWarningDays ?? 30;
  const complianceCutoff = new Date(now.getTime() + complianceWindowDays * 86400000);

  const [
    openJobs,
    draftJobs,
    awaitingApproval,
    flaggedApprovals,
    contractsPending,
    pendingVendors,
    complianceExpiring,
    activeProperties,
    emergencies,
    recentJobs,
    recentActivity,
  ] = await Promise.all([
    prisma.job.count({ where: { status: "OPEN", priority: "STANDARD" } }),
    prisma.job.count({ where: { status: "DRAFT" } }),
    prisma.job.count({ where: { status: "AWAITING_APPROVAL" } }),
    prisma.approvalFlag.count({ where: { acknowledgedAt: null } }),
    prisma.contract.count({ where: { status: "PENDING_SIGNATURE" } }),
    prisma.vendor.count({ where: { accountStatus: "PENDING" } }),
    prisma.vendor.count({
      where: {
        accountStatus: "ACTIVE",
        OR: [
          { complianceStatus: "EXPIRED" },
          { complianceExpiresAt: { lte: complianceCutoff, gte: now } },
        ],
      },
    }),
    prisma.property.count({ where: { isActive: true } }),
    prisma.job.findMany({
      where: { priority: "EMERGENCY", status: { in: ["OPEN", "AWARDED", "IN_PROGRESS"] } },
      include: { property: { select: { name: true, city: true } }, claimedBy: { select: { companyName: true } } },
      orderBy: { dispatchedAt: "desc" },
      take: 8,
    }),
    prisma.job.findMany({
      orderBy: { createdAt: "desc" },
      take: 6,
      include: {
        property: { select: { name: true } },
        _count: { select: { bids: true } },
      },
    }),
    prisma.activityLog.findMany({
      orderBy: { createdAt: "desc" },
      take: 8,
      include: { actorAdmin: { select: { name: true } }, actorVendor: { select: { companyName: true } } },
    }),
  ]);

  // An emergency is "overdue" once its response window has elapsed with nobody
  // claiming it — this is the thing that must never be missed on this page.
  const overdueEmergencies = emergencies.filter(
    (job) =>
      !job.claimedByVendorId &&
      job.dispatchedAt &&
      now.getTime() - job.dispatchedAt.getTime() > (job.responseDeadlineMinutes ?? 15) * 60000,
  );

  const providers = listPropertyProviderStatuses();

  const stats = [
    { label: "Open jobs", value: openJobs, href: "/admin/jobs?status=OPEN", icon: ClipboardList },
    { label: "Drafts", value: draftJobs, href: "/admin/jobs?status=DRAFT", icon: ClipboardList },
    { label: "Awaiting approval", value: awaitingApproval, href: "/admin/jobs?status=AWAITING_APPROVAL", icon: Gavel },
    { label: "Flagged approvals", value: flaggedApprovals, href: "/admin/jobs", icon: Flag },
    { label: "Contracts to sign", value: contractsPending, href: "/admin/jobs", icon: FileSignature },
    { label: "Vendors pending", value: pendingVendors, href: "/admin/jobs", icon: UserCheck },
    { label: "Compliance expiring", value: complianceExpiring, href: "/admin/jobs", icon: ShieldAlert },
    { label: "Active properties", value: activeProperties, href: "/admin/properties", icon: Building2 },
  ];

  return (
    <>
      <PageHeader title={`Good day, ${actor.name.split(" ")[0]}`} description="Everything that needs your attention right now.">
        <Button asChild>
          <Link href="/admin/jobs/new">New job</Link>
        </Button>
        <Button asChild variant="outline">
          <Link href="/admin/import">Import properties</Link>
        </Button>
      </PageHeader>

      {overdueEmergencies.length > 0 && (
        <Alert variant="emergency" className="mb-6 animate-emergency-pulse">
          <Siren />
          <AlertTitle>
            {overdueEmergencies.length} emergency job{overdueEmergencies.length === 1 ? "" : "s"} unclaimed past the
            response window
          </AlertTitle>
          <AlertDescription>
            <ul className="mt-2 space-y-1">
              {overdueEmergencies.map((job) => (
                <li key={job.id}>
                  <Link href={`/admin/jobs/${job.id}`} className="font-medium text-foreground hover:underline">
                    {job.jobNumber} — {job.title}
                  </Link>{" "}
                  · {job.property.name} · dispatched {formatRelative(job.dispatchedAt)}
                </li>
              ))}
            </ul>
          </AlertDescription>
        </Alert>
      )}

      <div className="grid grid-cols-2 gap-3 sm:grid-cols-3 lg:grid-cols-4">
        {stats.map((stat) => {
          const Icon = stat.icon;
          const emphasise = stat.value > 0 && ["Flagged approvals", "Compliance expiring", "Vendors pending"].includes(stat.label);
          return (
            <Link key={stat.label} href={stat.href}>
              <Card className="h-full transition-shadow hover:shadow-md">
                <CardContent className="flex items-start justify-between gap-2 p-4">
                  <div className="space-y-1">
                    <p className="text-xs font-medium uppercase tracking-wide text-muted-foreground">{stat.label}</p>
                    <p className={`text-2xl font-semibold tabular-nums ${emphasise ? "text-warning" : ""}`}>
                      {stat.value}
                    </p>
                  </div>
                  <Icon className="h-4 w-4 shrink-0 text-muted-foreground" />
                </CardContent>
              </Card>
            </Link>
          );
        })}
      </div>

      <div className="mt-6 grid min-w-0 gap-6 lg:grid-cols-3">
        <Card className="min-w-0 lg:col-span-2">
          <CardHeader className="flex-row items-center justify-between space-y-0">
            <div>
              <CardTitle>Recent jobs</CardTitle>
              <CardDescription>Newest first, emergencies included.</CardDescription>
            </div>
            <Button asChild variant="ghost" size="sm">
              <Link href="/admin/jobs">
                All jobs <ArrowRight className="h-4 w-4" />
              </Link>
            </Button>
          </CardHeader>
          <CardContent className="p-0">
            {recentJobs.length === 0 ? (
              <EmptyState
                icon={ClipboardList}
                title="No jobs yet"
                description="Create your first make-ready or contracting job to start collecting bids."
                action={
                  <Button asChild size="sm">
                    <Link href="/admin/jobs/new">New job</Link>
                  </Button>
                }
                className="m-5 border-0"
              />
            ) : (
              <ul className="divide-y">
                {recentJobs.map((job) => (
                  <li key={job.id}>
                    <Link
                      href={`/admin/jobs/${job.id}`}
                      className="flex flex-col gap-2 px-5 py-3 transition-colors hover:bg-accent/50 sm:flex-row sm:items-center sm:justify-between"
                    >
                      <div className="min-w-0 space-y-1">
                        <div className="flex min-w-0 flex-wrap items-center gap-2">
                          <span className="font-mono text-xs text-muted-foreground">{job.jobNumber}</span>
                          <span className="min-w-0 truncate font-medium">{job.title}</span>
                          {job.priority === "EMERGENCY" && <EmergencyBadge />}
                        </div>
                        <p className="truncate text-sm text-muted-foreground">
                          {job.property.name}
                          {job.unitNumber ? ` · Unit ${job.unitNumber}` : ""} ·{" "}
                          {formatCurrency(job.budgetMin)}–{formatCurrency(job.budgetMax)}
                        </p>
                      </div>
                      <div className="flex shrink-0 items-center gap-2">
                        <Badge variant="muted">
                          {job._count.bids} bid{job._count.bids === 1 ? "" : "s"}
                        </Badge>
                        <JobStatusBadge status={job.status} />
                      </div>
                    </Link>
                  </li>
                ))}
              </ul>
            )}
          </CardContent>
        </Card>

        <div className="min-w-0 space-y-6">
          <Card>
            <CardHeader>
              <CardTitle className="flex items-center gap-2">
                <Plug className="h-4 w-4" /> Integrations
              </CardTitle>
              <CardDescription>Property sync adapters.</CardDescription>
            </CardHeader>
            <CardContent className="space-y-3">
              {providers.map((provider) => (
                <div key={provider.key} className="flex items-start justify-between gap-3">
                  <div className="min-w-0">
                    <p className="text-sm font-medium">{provider.label}</p>
                    <p className="text-xs text-muted-foreground">
                      {provider.configured ? provider.description : provider.reason}
                    </p>
                  </div>
                  <Badge variant={provider.active ? "success" : provider.configured ? "secondary" : "muted"}>
                    {provider.active ? "Active" : provider.configured ? "Ready" : "Stubbed"}
                  </Badge>
                </div>
              ))}
            </CardContent>
          </Card>

          <Card>
            <CardHeader>
              <CardTitle>Recent activity</CardTitle>
              <CardDescription>Every state change is logged.</CardDescription>
            </CardHeader>
            <CardContent>
              {recentActivity.length === 0 ? (
                <p className="text-sm text-muted-foreground">Nothing logged yet.</p>
              ) : (
                <ol className="space-y-3">
                  {recentActivity.map((entry) => (
                    <li key={entry.id} className="flex gap-3">
                      <div className="mt-1.5 h-1.5 w-1.5 shrink-0 rounded-full bg-primary" />
                      <div className="min-w-0 space-y-0.5">
                        <p className="text-sm leading-snug">{entry.summary ?? entry.action}</p>
                        <p className="text-xs text-muted-foreground">
                          {entry.actorAdmin?.name ?? entry.actorVendor?.companyName ?? "System"} ·{" "}
                          {formatDate(entry.createdAt, true)}
                        </p>
                      </div>
                    </li>
                  ))}
                </ol>
              )}
            </CardContent>
          </Card>
        </div>
      </div>

      {emergencies.length > 0 && overdueEmergencies.length === 0 && (
        <Card className="mt-6">
          <CardHeader>
            <CardTitle className="flex items-center gap-2">
              <AlertTriangle className="h-4 w-4 text-emergency" /> Active emergencies
            </CardTitle>
            <CardDescription>All within their response window or already claimed.</CardDescription>
          </CardHeader>
          <CardContent className="space-y-2">
            {emergencies.map((job) => (
              <Link
                key={job.id}
                href={`/admin/jobs/${job.id}`}
                className="flex items-center justify-between rounded-lg border border-emergency-border bg-emergency-soft px-3 py-2 text-sm"
              >
                <span className="font-medium">
                  {job.jobNumber} · {job.title}
                </span>
                <span className="text-muted-foreground">
                  {job.claimedBy ? `Claimed by ${job.claimedBy.companyName}` : `Dispatched ${formatRelative(job.dispatchedAt)}`}
                </span>
              </Link>
            ))}
          </CardContent>
        </Card>
      )}
    </>
  );
}
