"use client";

import * as React from "react";
import { useRouter } from "next/navigation";
import {
  AlertCircle,
  CheckCircle2,
  Download,
  FileSpreadsheet,
  Loader2,
  Lock,
  Upload,
} from "lucide-react";
import { Alert, AlertDescription, AlertTitle } from "@/components/ui/alert";
import { Badge } from "@/components/ui/badge";
import { Button } from "@/components/ui/button";
import { Card, CardContent, CardDescription, CardHeader, CardTitle } from "@/components/ui/card";
import { Table, TableBody, TableCell, TableHead, TableHeader, TableRow } from "@/components/ui/table";
import { cn } from "@/lib/utils";

type ProviderStatus = {
  key: string;
  label: string;
  description: string;
  configured: boolean;
  reason?: string;
  requires?: string[];
  active: boolean;
  capabilities: { upload: boolean; pull: boolean; incremental: boolean };
};

type ImportIssue = { row?: number; field?: string; message: string; value?: string };

type ImportResult = {
  dryRun: boolean;
  providerKey: string;
  rowsTotal: number;
  created: number;
  updated: number;
  unchanged: number;
  failed: number;
  issues: ImportIssue[];
  preview: { externalId: string; name: string; city: string; state: string; action: string; changes?: string[] }[];
};

export function ImportClient({ providers }: { providers: ProviderStatus[] }) {
  const router = useRouter();
  const [selected, setSelected] = React.useState(providers.find((p) => p.active)?.key ?? "csv");
  const [file, setFile] = React.useState<File | null>(null);
  const [pending, setPending] = React.useState<"preview" | "import" | null>(null);
  const [error, setError] = React.useState<{ message: string; remedy?: string; issues?: ImportIssue[] } | null>(null);
  const [result, setResult] = React.useState<ImportResult | null>(null);
  const [dragging, setDragging] = React.useState(false);

  const provider = providers.find((p) => p.key === selected)!;

  async function submit(dryRun: boolean) {
    if (!file && provider.capabilities.upload) {
      setError({ message: "Choose a CSV file first." });
      return;
    }
    setPending(dryRun ? "preview" : "import");
    setError(null);

    const form = new FormData();
    form.set("providerKey", selected);
    form.set("dryRun", String(dryRun));
    if (file) form.set("file", file);

    const res = await fetch("/api/admin/properties/import", { method: "POST", body: form });
    const data = await res.json().catch(() => null);

    if (!res.ok) {
      setResult(null);
      setError({
        message: data?.error?.message ?? "The import failed.",
        remedy: data?.error?.remedy,
        issues: data?.error?.issues,
      });
      setPending(null);
      return;
    }

    setResult(data.result as ImportResult);
    setPending(null);
    if (!dryRun) router.refresh();
  }

  return (
    <div className="grid min-w-0 gap-6 lg:grid-cols-3">
      <div className="min-w-0 space-y-6 lg:col-span-2">
        <Card>
          <CardHeader>
            <CardTitle>Import properties</CardTitle>
            <CardDescription>
              Upload an export (for example, from Power BI). Column names are matched loosely, so the file you
              already have will usually work as-is.
            </CardDescription>
          </CardHeader>
          <CardContent className="space-y-4">
            {!provider.configured ? (
              <Alert variant="warning">
                <Lock />
                <AlertTitle>{provider.label} is not configured yet</AlertTitle>
                <AlertDescription>
                  <p>{provider.reason}</p>
                  {provider.requires && (
                    <p className="mt-2 font-mono text-xs">{provider.requires.join(", ")}</p>
                  )}
                  <p className="mt-2">
                    Use CSV upload in the meantime — the import path and everything downstream is identical, so
                    switching later needs no other changes.
                  </p>
                </AlertDescription>
              </Alert>
            ) : (
              <>
                <label
                  onDragOver={(e) => {
                    e.preventDefault();
                    setDragging(true);
                  }}
                  onDragLeave={() => setDragging(false)}
                  onDrop={(e) => {
                    e.preventDefault();
                    setDragging(false);
                    const dropped = e.dataTransfer.files?.[0];
                    if (dropped) {
                      setFile(dropped);
                      setResult(null);
                      setError(null);
                    }
                  }}
                  className={cn(
                    "flex cursor-pointer flex-col items-center justify-center gap-2 rounded-xl border-2 border-dashed px-6 py-10 text-center transition-colors",
                    dragging ? "border-primary bg-primary/5" : "border-input hover:border-primary/50",
                  )}
                >
                  <input
                    type="file"
                    accept=".csv,text/csv"
                    className="sr-only"
                    onChange={(e) => {
                      setFile(e.target.files?.[0] ?? null);
                      setResult(null);
                      setError(null);
                    }}
                  />
                  <FileSpreadsheet className="h-8 w-8 text-muted-foreground" />
                  {file ? (
                    <div>
                      <p className="font-medium">{file.name}</p>
                      <p className="text-xs text-muted-foreground">{(file.size / 1024).toFixed(1)} KB</p>
                    </div>
                  ) : (
                    <div>
                      <p className="font-medium">Drop a CSV here, or click to choose</p>
                      <p className="text-xs text-muted-foreground">Up to 5 MB</p>
                    </div>
                  )}
                </label>

                <div className="flex flex-wrap items-center gap-2">
                  <Button onClick={() => submit(true)} variant="outline" disabled={pending !== null || !file}>
                    {pending === "preview" && <Loader2 className="animate-spin" />}
                    Preview changes
                  </Button>
                  <Button onClick={() => submit(false)} disabled={pending !== null || !file}>
                    {pending === "import" ? <Loader2 className="animate-spin" /> : <Upload />}
                    Import
                  </Button>
                  <Button asChild variant="ghost" size="sm">
                    <a href="/api/admin/properties/template" download>
                      <Download /> Template
                    </a>
                  </Button>
                </div>
              </>
            )}

            {error && (
              <Alert variant="destructive">
                <AlertCircle />
                <AlertTitle>{error.message}</AlertTitle>
                <AlertDescription>
                  {error.remedy && <p>{error.remedy}</p>}
                  {error.issues && error.issues.length > 0 && (
                    <ul className="mt-2 space-y-1 text-xs">
                      {error.issues.slice(0, 8).map((issue, i) => (
                        <li key={i}>
                          {issue.row ? `Row ${issue.row}: ` : ""}
                          {issue.message}
                        </li>
                      ))}
                      {error.issues.length > 8 && <li>…and {error.issues.length - 8} more.</li>}
                    </ul>
                  )}
                </AlertDescription>
              </Alert>
            )}

            {result && (
              <Alert variant={result.dryRun ? "info" : "success"}>
                <CheckCircle2 />
                <AlertTitle>
                  {result.dryRun
                    ? "Preview — nothing has been saved yet"
                    : `Imported ${result.created + result.updated} of ${result.rowsTotal} rows`}
                </AlertTitle>
                <AlertDescription>
                  {result.created} new · {result.updated} updated · {result.unchanged} unchanged
                  {result.failed > 0 ? ` · ${result.failed} skipped` : ""}
                </AlertDescription>
              </Alert>
            )}
          </CardContent>
        </Card>

        {result && result.preview.length > 0 && (
          <Card>
            <CardHeader>
              <CardTitle>{result.dryRun ? "Preview" : "Imported rows"}</CardTitle>
              <CardDescription>{result.preview.length} rows parsed successfully.</CardDescription>
            </CardHeader>
            <CardContent className="p-0">
              <Table>
                <TableHeader>
                  <TableRow>
                    <TableHead>Property ID</TableHead>
                    <TableHead>Name</TableHead>
                    <TableHead>Location</TableHead>
                    <TableHead>Action</TableHead>
                  </TableRow>
                </TableHeader>
                <TableBody>
                  {result.preview.slice(0, 100).map((row) => (
                    <TableRow key={row.externalId}>
                      <TableCell className="font-mono text-xs">{row.externalId}</TableCell>
                      <TableCell className="font-medium">{row.name}</TableCell>
                      <TableCell className="text-muted-foreground">
                        {row.city}, {row.state}
                      </TableCell>
                      <TableCell>
                        <Badge
                          variant={
                            row.action === "create" ? "success" : row.action === "update" ? "warning" : "muted"
                          }
                        >
                          {row.action}
                        </Badge>
                        {row.changes && row.changes.length > 0 && (
                          <span className="ml-2 text-xs text-muted-foreground">{row.changes.join(", ")}</span>
                        )}
                      </TableCell>
                    </TableRow>
                  ))}
                </TableBody>
              </Table>
            </CardContent>
          </Card>
        )}

        {result && result.issues.length > 0 && (
          <Card>
            <CardHeader>
              <CardTitle>Skipped rows</CardTitle>
              <CardDescription>These rows were not imported. Fix them and upload again.</CardDescription>
            </CardHeader>
            <CardContent className="p-0">
              <Table>
                <TableHeader>
                  <TableRow>
                    <TableHead className="w-20">Row</TableHead>
                    <TableHead>Field</TableHead>
                    <TableHead>Problem</TableHead>
                  </TableRow>
                </TableHeader>
                <TableBody>
                  {result.issues.slice(0, 50).map((issue, i) => (
                    <TableRow key={i}>
                      <TableCell className="tabular-nums">{issue.row ?? "—"}</TableCell>
                      <TableCell className="font-mono text-xs">{issue.field ?? "—"}</TableCell>
                      <TableCell className="text-muted-foreground">{issue.message}</TableCell>
                    </TableRow>
                  ))}
                </TableBody>
              </Table>
            </CardContent>
          </Card>
        )}
      </div>

      <Card className="h-fit min-w-0">
        <CardHeader>
          <CardTitle>Data source</CardTitle>
          <CardDescription>
            All three sources produce identical property records. Whichever API access arrives first can be turned
            on without changing anything else.
          </CardDescription>
        </CardHeader>
        <CardContent className="space-y-2">
          {providers.map((p) => (
            <button
              key={p.key}
              type="button"
              onClick={() => {
                setSelected(p.key);
                setResult(null);
                setError(null);
              }}
              className={cn(
                "w-full rounded-lg border p-3 text-left transition-colors",
                selected === p.key ? "border-primary bg-primary/5" : "hover:bg-accent",
              )}
            >
              <div className="flex items-center justify-between gap-2">
                <span className="font-medium">{p.label}</span>
                <Badge variant={p.active ? "success" : p.configured ? "secondary" : "muted"}>
                  {p.active ? "Active" : p.configured ? "Ready" : "Stubbed"}
                </Badge>
              </div>
              <p className="mt-1 text-xs text-muted-foreground">
                {p.configured ? p.description : p.reason}
              </p>
            </button>
          ))}
        </CardContent>
      </Card>
    </div>
  );
}
