"use client";

import * as React from "react";
import { useRouter } from "next/navigation";
import { AlertCircle, Loader2 } from "lucide-react";
import { Alert, AlertDescription } from "@/components/ui/alert";
import { Button } from "@/components/ui/button";
import {
  Dialog,
  DialogContent,
  DialogDescription,
  DialogFooter,
  DialogHeader,
  DialogTitle,
} from "@/components/ui/dialog";
import { Input } from "@/components/ui/input";
import { Label } from "@/components/ui/label";
import { Switch } from "@/components/ui/switch";

export type PropertyFormValues = {
  id?: string;
  externalId?: string | null;
  name: string;
  addressLine1: string;
  addressLine2?: string | null;
  city: string;
  state: string;
  postalCode: string;
  latitude: number | string;
  longitude: number | string;
  unitCount?: number | string | null;
  propertyManagerName?: string | null;
  propertyManagerEmail?: string | null;
  propertyManagerPhone?: string | null;
  isActive: boolean;
};

const BLANK: PropertyFormValues = {
  name: "",
  addressLine1: "",
  addressLine2: "",
  city: "",
  state: "",
  postalCode: "",
  latitude: "",
  longitude: "",
  unitCount: "",
  propertyManagerName: "",
  propertyManagerEmail: "",
  propertyManagerPhone: "",
  isActive: true,
};

export function PropertyDialog({
  open,
  onOpenChange,
  property,
}: {
  open: boolean;
  onOpenChange: (open: boolean) => void;
  property?: PropertyFormValues | null;
}) {
  const router = useRouter();
  const [values, setValues] = React.useState<PropertyFormValues>(property ?? BLANK);
  const [error, setError] = React.useState<string | null>(null);
  const [pending, setPending] = React.useState(false);

  React.useEffect(() => {
    if (open) {
      setValues(property ?? BLANK);
      setError(null);
    }
  }, [open, property]);

  const set = <K extends keyof PropertyFormValues>(key: K, value: PropertyFormValues[K]) =>
    setValues((prev) => ({ ...prev, [key]: value }));

  async function onSubmit(event: React.FormEvent) {
    event.preventDefault();
    setPending(true);
    setError(null);

    const payload = {
      externalId: values.externalId || null,
      name: values.name,
      addressLine1: values.addressLine1,
      addressLine2: values.addressLine2 || null,
      city: values.city,
      state: values.state,
      postalCode: values.postalCode,
      latitude: Number(values.latitude),
      longitude: Number(values.longitude),
      unitCount: values.unitCount === "" || values.unitCount === null ? null : Number(values.unitCount),
      propertyManagerName: values.propertyManagerName || null,
      propertyManagerEmail: values.propertyManagerEmail || null,
      propertyManagerPhone: values.propertyManagerPhone || null,
      isActive: values.isActive,
    };

    const res = await fetch(property?.id ? `/api/admin/properties/${property.id}` : "/api/admin/properties", {
      method: property?.id ? "PATCH" : "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(payload),
    });

    if (!res.ok) {
      const data = (await res.json().catch(() => null)) as
        | { error?: { message?: string; issues?: { path: string; message: string }[] } }
        | null;
      const issue = data?.error?.issues?.[0];
      setError(issue ? `${issue.path}: ${issue.message}` : (data?.error?.message ?? "Could not save the property."));
      setPending(false);
      return;
    }

    setPending(false);
    onOpenChange(false);
    router.refresh();
  }

  return (
    <Dialog open={open} onOpenChange={onOpenChange}>
      <DialogContent className="max-w-2xl">
        <DialogHeader>
          <DialogTitle>{property?.id ? "Edit property" : "Add property"}</DialogTitle>
          <DialogDescription>
            Latitude and longitude are required — vendor radius matching and the map view read them directly.
          </DialogDescription>
        </DialogHeader>

        <form onSubmit={onSubmit} className="space-y-4">
          {error && (
            <Alert variant="destructive">
              <AlertCircle />
              <AlertDescription>{error}</AlertDescription>
            </Alert>
          )}

          <div className="grid gap-4 sm:grid-cols-2">
            <div className="space-y-2 sm:col-span-2">
              <Label htmlFor="name">Property name</Label>
              <Input id="name" required value={values.name} onChange={(e) => set("name", e.target.value)} />
            </div>

            <div className="space-y-2 sm:col-span-2">
              <Label htmlFor="addressLine1">Street address</Label>
              <Input
                id="addressLine1"
                required
                value={values.addressLine1}
                onChange={(e) => set("addressLine1", e.target.value)}
              />
            </div>

            <div className="space-y-2">
              <Label htmlFor="city">City</Label>
              <Input id="city" required value={values.city} onChange={(e) => set("city", e.target.value)} />
            </div>

            <div className="grid grid-cols-2 gap-3">
              <div className="space-y-2">
                <Label htmlFor="state">State</Label>
                <Input id="state" required value={values.state} onChange={(e) => set("state", e.target.value)} />
              </div>
              <div className="space-y-2">
                <Label htmlFor="postalCode">ZIP</Label>
                <Input
                  id="postalCode"
                  required
                  value={values.postalCode}
                  onChange={(e) => set("postalCode", e.target.value)}
                />
              </div>
            </div>

            <div className="space-y-2">
              <Label htmlFor="latitude">Latitude</Label>
              <Input
                id="latitude"
                type="number"
                step="any"
                required
                value={values.latitude}
                onChange={(e) => set("latitude", e.target.value)}
                placeholder="32.7997"
              />
            </div>

            <div className="space-y-2">
              <Label htmlFor="longitude">Longitude</Label>
              <Input
                id="longitude"
                type="number"
                step="any"
                required
                value={values.longitude}
                onChange={(e) => set("longitude", e.target.value)}
                placeholder="-96.8065"
              />
            </div>

            <div className="space-y-2">
              <Label htmlFor="unitCount">Units</Label>
              <Input
                id="unitCount"
                type="number"
                min={0}
                value={values.unitCount ?? ""}
                onChange={(e) => set("unitCount", e.target.value)}
              />
            </div>

            <div className="space-y-2">
              <Label htmlFor="externalId">External ID</Label>
              <Input
                id="externalId"
                value={values.externalId ?? ""}
                onChange={(e) => set("externalId", e.target.value)}
                placeholder="CRT-DAL-0148"
              />
            </div>

            <div className="space-y-2">
              <Label htmlFor="propertyManagerName">Property manager</Label>
              <Input
                id="propertyManagerName"
                value={values.propertyManagerName ?? ""}
                onChange={(e) => set("propertyManagerName", e.target.value)}
              />
            </div>

            <div className="space-y-2">
              <Label htmlFor="propertyManagerEmail">Manager email</Label>
              <Input
                id="propertyManagerEmail"
                type="email"
                value={values.propertyManagerEmail ?? ""}
                onChange={(e) => set("propertyManagerEmail", e.target.value)}
              />
            </div>

            <div className="space-y-2">
              <Label htmlFor="propertyManagerPhone">Manager phone</Label>
              <Input
                id="propertyManagerPhone"
                value={values.propertyManagerPhone ?? ""}
                onChange={(e) => set("propertyManagerPhone", e.target.value)}
              />
            </div>

            <div className="flex items-center justify-between rounded-lg border p-3 sm:col-span-2">
              <div>
                <Label htmlFor="isActive">Active</Label>
                <p className="text-xs text-muted-foreground">Inactive properties cannot have new jobs posted.</p>
              </div>
              <Switch id="isActive" checked={values.isActive} onCheckedChange={(v) => set("isActive", v)} />
            </div>
          </div>

          <DialogFooter>
            <Button type="button" variant="outline" onClick={() => onOpenChange(false)} disabled={pending}>
              Cancel
            </Button>
            <Button type="submit" disabled={pending}>
              {pending && <Loader2 className="animate-spin" />}
              {property?.id ? "Save changes" : "Add property"}
            </Button>
          </DialogFooter>
        </form>
      </DialogContent>
    </Dialog>
  );
}
