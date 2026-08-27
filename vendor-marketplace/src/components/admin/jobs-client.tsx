"use client";

import * as React from "react";
import Link from "next/link";
import { ClipboardList, Search, Siren } from "lucide-react";
import type { JobStatus } from "@prisma/client";
import { EmptyState } from "@/components/empty-state";
import { EmergencyBadge, JobStatusBadge } from "@/components/status-badges";
import { Badge } from "@/components/ui/badge";
import { Button } from "@/components/ui/button";
import { Card, CardContent } from "@/components/ui/card";
import { Input } from "@/components/ui/input";
import { Table, TableBody, TableCell, TableHead, TableHeader, TableRow } from "@/components/ui/table";
import { cn, formatCurrency, formatDate, formatRelative, titleCase } from "@/lib/utils";

export type JobRow = {
  id: string;
  jobNumber: string;
  title: string;
  status: JobStatus;
  priority: "STANDARD" | "EMERGENCY";
  emergencyCategory: string | null;
  category: string;
  unitNumber: string | null;
  propertyName: string;
  propertyId: string;
  budgetMin: string;
  budgetMax: string;
  enforceBudgetCap: boolean;
  bidDeadline: string | null;
  dispatchedAt: string | null;
  claimedAt: string | null;
  responseDeadlineMinutes: number | null;
  bidCount: number;
  createdAt: string;
};

const STATUS_FILTERS: (JobStatus | "all")[] = [
  "all",
  "DRAFT",
  "OPEN",
  "BIDDING_CLOSED",
  "AWAITING_APPROVAL",
  "AWARDED",
  "IN_PROGRESS",
  "COMPLETED",
  "CANCELLED",
];

export function JobsClient({
  jobs,
  properties,
  initialStatus,
}: {
  jobs: JobRow[];
  properties: { id: string; name: string }[];
  initialStatus?: string | null;
}) {
  const [status, setStatus] = React.useState<string>(initialStatus ?? "all");
  const [propertyId, setPropertyId] = React.useState("all");
  const [query, setQuery] = React.useState("");
  const [emergencyOnly, setEmergencyOnly] = React.useState(false);

  const filtered = React.useMemo(() => {
    const q = query.trim().toLowerCase();
    return jobs.filter((job) => {
      if (status !== "all" && job.status !== status) return false;
      if (propertyId !== "all" && job.propertyId !== propertyId) return false;
      if (emergencyOnly && job.priority !== "EMERGENCY") return false;
      if (!q) return true;
      return [job.title, job.jobNumber, job.propertyName, job.unitNumber ?? ""].some((f) =>
        f.toLowerCase().includes(q),
      );
    });
  }, [jobs, status, propertyId, query, emergencyOnly]);

  const emergencies = filtered.filter((j) => j.priority === "EMERGENCY");
  const standard = filtered.filter((j) => j.priority === "STANDARD");

  return (
    <>
      <div className="mb-4 space-y-3">
        <div className="flex flex-col gap-3 sm:flex-row sm:items-center">
          <div className="relative w-full sm:max-w-xs">
            <Search className="pointer-events-none absolute left-3 top-1/2 h-4 w-4 -translate-y-1/2 text-muted-foreground" />
            <Input
              value={query}
              onChange={(e) => setQuery(e.target.value)}
              placeholder="Search job, unit, or property"
              className="pl-9"
              aria-label="Search jobs"
            />
          </div>
          <select
            value={propertyId}
            onChange={(e) => setPropertyId(e.target.value)}
            aria-label="Filter by property"
            className="h-9 rounded-md border border-input bg-card px-3 text-sm shadow-sm focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-ring"
          >
            <option value="all">All properties</option>
            {properties.map((p) => (
              <option key={p.id} value={p.id}>
                {p.name}
              </option>
            ))}
          </select>
          <Button
            type="button"
            variant={emergencyOnly ? "emergency" : "outline"}
            size="sm"
            onClick={() => setEmergencyOnly((v) => !v)}
          >
            <Siren /> Emergencies only
          </Button>
        </div>

        {/* Single-tap status filters rather than a dropdown. */}
        <div className="flex flex-wrap gap-1.5">
          {STATUS_FILTERS.map((value) => {
            const count = value === "all" ? jobs.length : jobs.filter((j) => j.status === value).length;
            return (
              <button
                key={value}
                type="button"
                onClick={() => setStatus(value)}
                className={cn(
                  "rounded-full border px-3 py-1 text-xs font-medium transition-colors",
                  status === value
                    ? "border-primary bg-primary text-primary-foreground"
                    : "bg-card text-muted-foreground hover:bg-accent",
                )}
              >
                {value === "all" ? "All" : titleCase(value)}
                <span className="ml-1.5 tabular-nums opacity-70">{count}</span>
              </button>
            );
          })}
        </div>
      </div>

      {filtered.length === 0 ? (
        <EmptyState
          icon={ClipboardList}
          title={jobs.length === 0 ? "No jobs yet" : "No jobs match these filters"}
          description={
            jobs.length === 0
              ? "Post your first make-ready or contracting job."
              : "Clear a filter or search for something else."
          }
          action={
            jobs.length === 0 ? (
              <Button asChild size="sm">
                <Link href="/admin/jobs/new">New job</Link>
              </Button>
            ) : undefined
          }
        />
      ) : (
        <div className="space-y-6">
          {emergencies.length > 0 && (
            <section>
              <h2 className="mb-2 flex items-center gap-2 text-sm font-semibold uppercase tracking-wide text-emergency">
                <Siren className="h-4 w-4" /> Emergency
              </h2>
              <Card className="border-emergency-border">
                <CardContent className="p-0">
                  <JobTable rows={emergencies} emergency />
                </CardContent>
              </Card>
            </section>
          )}

          {standard.length > 0 && (
            <section>
              {emergencies.length > 0 && (
                <h2 className="mb-2 text-sm font-semibold uppercase tracking-wide text-muted-foreground">
                  Standard
                </h2>
              )}
              <Card>
                <CardContent className="p-0">
                  <JobTable rows={standard} />
                </CardContent>
              </Card>
            </section>
          )}
        </div>
      )}
    </>
  );
}

function JobTable({ rows, emergency = false }: { rows: JobRow[]; emergency?: boolean }) {
  return (
    <Table>
      <TableHeader>
        <TableRow>
          <TableHead>Job</TableHead>
          <TableHead>Property</TableHead>
          <TableHead>Budget</TableHead>
          <TableHead>{emergency ? "Dispatched" : "Bid deadline"}</TableHead>
          <TableHead className="text-right">{emergency ? "Claimed" : "Bids"}</TableHead>
          <TableHead>Status</TableHead>
        </TableRow>
      </TableHeader>
      <TableBody>
        {rows.map((job) => {
          const deadlinePassed = job.bidDeadline ? new Date(job.bidDeadline).getTime() < Date.now() : false;
          const overdue =
            emergency &&
            !job.claimedAt &&
            job.dispatchedAt &&
            Date.now() - new Date(job.dispatchedAt).getTime() > (job.responseDeadlineMinutes ?? 15) * 60000;

          return (
            <TableRow key={job.id} className={cn(overdue && "bg-emergency-soft")}>
              <TableCell>
                <div className="flex flex-col">
                  <Link href={`/admin/jobs/${job.id}`} className="font-medium hover:underline">
                    {job.title}
                  </Link>
                  <span className="font-mono text-xs text-muted-foreground">
                    {job.jobNumber}
                    {job.unitNumber ? ` · Unit ${job.unitNumber}` : ""} · {titleCase(job.category)}
                  </span>
                </div>
              </TableCell>
              <TableCell className="whitespace-nowrap text-muted-foreground">{job.propertyName}</TableCell>
              <TableCell className="whitespace-nowrap tabular-nums">
                {formatCurrency(job.budgetMin)}–{formatCurrency(job.budgetMax)}
                {job.enforceBudgetCap && (
                  <Badge variant="secondary" className="ml-2">
                    cap
                  </Badge>
                )}
              </TableCell>
              <TableCell className="whitespace-nowrap text-muted-foreground">
                {emergency ? (
                  job.dispatchedAt ? (
                    <span className={cn(overdue && "font-medium text-emergency")}>
                      {formatRelative(job.dispatchedAt)}
                      {overdue ? " · overdue" : ""}
                    </span>
                  ) : (
                    "—"
                  )
                ) : job.bidDeadline ? (
                  <span className={cn(deadlinePassed && "text-muted-foreground line-through")}>
                    {formatDate(job.bidDeadline, true)}
                  </span>
                ) : (
                  "—"
                )}
              </TableCell>
              <TableCell className="text-right tabular-nums">
                {emergency ? (job.claimedAt ? formatRelative(job.claimedAt) : "—") : job.bidCount}
              </TableCell>
              <TableCell>
                <div className="flex items-center gap-1.5">
                  {job.priority === "EMERGENCY" && <EmergencyBadge />}
                  <JobStatusBadge status={job.status} />
                </div>
              </TableCell>
            </TableRow>
          );
        })}
      </TableBody>
    </Table>
  );
}
