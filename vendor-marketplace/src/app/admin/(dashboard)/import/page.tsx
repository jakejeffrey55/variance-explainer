import type { Metadata } from "next";
import { History } from "lucide-react";
import { PageHeader } from "@/components/admin/admin-shell";
import { ImportClient } from "@/components/admin/import-client";
import { Badge } from "@/components/ui/badge";
import { Card, CardContent, CardDescription, CardHeader, CardTitle } from "@/components/ui/card";
import { requireAdmin } from "@/lib/auth/session";
import { prisma } from "@/lib/db";
import { listPropertyProviderStatuses } from "@/lib/integrations/property";
import { formatDate } from "@/lib/utils";

export const metadata: Metadata = { title: "Import" };
export const dynamic = "force-dynamic";

export default async function ImportPage() {
  await requireAdmin();

  const providers = listPropertyProviderStatuses();
  const batches = await prisma.importBatch.findMany({
    where: { kind: "PROPERTY" },
    orderBy: { startedAt: "desc" },
    take: 10,
    include: { adminUser: { select: { name: true } } },
  });

  return (
    <>
      <PageHeader
        title="Import properties"
        description="CSV upload is active today. Power BI and OneSite are wired to the same interface and switch on once API access lands."
      />

      <ImportClient providers={providers} />

      <Card className="mt-6">
        <CardHeader>
          <CardTitle className="flex items-center gap-2">
            <History className="h-4 w-4" /> Import history
          </CardTitle>
          <CardDescription>Last 10 property imports.</CardDescription>
        </CardHeader>
        <CardContent>
          {batches.length === 0 ? (
            <p className="text-sm text-muted-foreground">No imports run yet.</p>
          ) : (
            <ul className="divide-y">
              {batches.map((batch) => (
                <li key={batch.id} className="flex flex-wrap items-center justify-between gap-2 py-3 first:pt-0 last:pb-0">
                  <div className="min-w-0">
                    <p className="truncate text-sm font-medium">{batch.fileName ?? `${batch.providerKey} sync`}</p>
                    <p className="text-xs text-muted-foreground">
                      {formatDate(batch.startedAt, true)} · {batch.adminUser?.name ?? "System"} ·{" "}
                      {batch.rowsImported} imported, {batch.rowsSkipped} unchanged, {batch.rowsFailed} skipped
                    </p>
                  </div>
                  <Badge
                    variant={
                      batch.status === "COMPLETED" ? "success" : batch.status === "FAILED" ? "destructive" : "secondary"
                    }
                  >
                    {batch.status.toLowerCase()}
                  </Badge>
                </li>
              ))}
            </ul>
          )}
        </CardContent>
      </Card>
    </>
  );
}
