"use client";

import * as React from "react";
import Link from "next/link";
import { Building2, Pencil, Plus, Search } from "lucide-react";
import { PropertyDialog, type PropertyFormValues } from "@/components/admin/property-dialog";
import { EmptyState } from "@/components/empty-state";
import { Badge } from "@/components/ui/badge";
import { Button } from "@/components/ui/button";
import { Card, CardContent } from "@/components/ui/card";
import { Input } from "@/components/ui/input";
import { Table, TableBody, TableCell, TableHead, TableHeader, TableRow } from "@/components/ui/table";

export type PropertyRow = PropertyFormValues & {
  id: string;
  source: string;
  jobCount: number;
  unitRollCount: number;
  lastSyncedAt: string | null;
};

const SOURCE_LABEL: Record<string, string> = {
  CSV: "CSV import",
  POWER_BI: "Power BI",
  ONESITE: "OneSite",
  MANUAL: "Manual",
};

export function PropertiesClient({ properties }: { properties: PropertyRow[] }) {
  const [query, setQuery] = React.useState("");
  const [showInactive, setShowInactive] = React.useState(false);
  const [dialogOpen, setDialogOpen] = React.useState(false);
  const [editing, setEditing] = React.useState<PropertyFormValues | null>(null);

  const filtered = React.useMemo(() => {
    const q = query.trim().toLowerCase();
    return properties.filter((p) => {
      if (!showInactive && !p.isActive) return false;
      if (!q) return true;
      return [p.name, p.city, p.state, p.addressLine1, p.externalId ?? ""].some((field) =>
        field.toLowerCase().includes(q),
      );
    });
  }, [properties, query, showInactive]);

  const inactiveCount = properties.filter((p) => !p.isActive).length;

  return (
    <>
      <div className="mb-4 flex flex-col gap-3 sm:flex-row sm:items-center sm:justify-between">
        <div className="relative w-full sm:max-w-xs">
          <Search className="pointer-events-none absolute left-3 top-1/2 h-4 w-4 -translate-y-1/2 text-muted-foreground" />
          <Input
            value={query}
            onChange={(e) => setQuery(e.target.value)}
            placeholder="Search name, city, or ID"
            className="pl-9"
            aria-label="Search properties"
          />
        </div>
        <div className="flex items-center gap-2">
          {inactiveCount > 0 && (
            <Button variant="outline" size="sm" onClick={() => setShowInactive((v) => !v)}>
              {showInactive ? "Hide" : "Show"} inactive ({inactiveCount})
            </Button>
          )}
          <Button
            size="sm"
            onClick={() => {
              setEditing(null);
              setDialogOpen(true);
            }}
          >
            <Plus /> Add property
          </Button>
        </div>
      </div>

      {filtered.length === 0 ? (
        <EmptyState
          icon={Building2}
          title={properties.length === 0 ? "No properties yet" : "No properties match that search"}
          description={
            properties.length === 0
              ? "Import a CSV export or add a property by hand to start posting jobs."
              : "Try a different name, city, or property ID."
          }
          action={
            properties.length === 0 ? (
              <div className="flex gap-2">
                <Button asChild size="sm">
                  <Link href="/admin/import">Import CSV</Link>
                </Button>
                <Button
                  size="sm"
                  variant="outline"
                  onClick={() => {
                    setEditing(null);
                    setDialogOpen(true);
                  }}
                >
                  Add manually
                </Button>
              </div>
            ) : undefined
          }
        />
      ) : (
        <Card>
          <CardContent className="p-0">
            <Table>
              <TableHeader>
                <TableRow>
                  <TableHead>Property</TableHead>
                  <TableHead>Location</TableHead>
                  <TableHead className="text-right">Units</TableHead>
                  <TableHead className="text-right">Jobs</TableHead>
                  <TableHead>Source</TableHead>
                  <TableHead className="w-24" />
                </TableRow>
              </TableHeader>
              <TableBody>
                {filtered.map((property) => (
                  <TableRow key={property.id}>
                    <TableCell>
                      <div className="flex flex-col">
                        <Link href={`/admin/properties/${property.id}`} className="font-medium hover:underline">
                          {property.name}
                        </Link>
                        <span className="text-xs text-muted-foreground">
                          {property.externalId ?? "—"} · {property.addressLine1}
                        </span>
                      </div>
                    </TableCell>
                    <TableCell className="whitespace-nowrap text-muted-foreground">
                      {property.city}, {property.state} {property.postalCode}
                    </TableCell>
                    <TableCell className="text-right tabular-nums">{property.unitCount ?? "—"}</TableCell>
                    <TableCell className="text-right tabular-nums">{property.jobCount}</TableCell>
                    <TableCell>
                      <div className="flex flex-wrap items-center gap-1">
                        <Badge variant="secondary">{SOURCE_LABEL[property.source] ?? property.source}</Badge>
                        {!property.isActive && <Badge variant="muted">Inactive</Badge>}
                      </div>
                    </TableCell>
                    <TableCell className="text-right">
                      <Button
                        variant="ghost"
                        size="sm"
                        onClick={() => {
                          setEditing(property);
                          setDialogOpen(true);
                        }}
                      >
                        <Pencil /> Edit
                      </Button>
                    </TableCell>
                  </TableRow>
                ))}
              </TableBody>
            </Table>
          </CardContent>
        </Card>
      )}

      <PropertyDialog open={dialogOpen} onOpenChange={setDialogOpen} property={editing} />
    </>
  );
}
