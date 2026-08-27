"use client";

import * as React from "react";
import { useRouter } from "next/navigation";
import { AlertCircle, Loader2, Siren } from "lucide-react";
import type { EmergencyCategory, ServiceCategory } from "@prisma/client";
import { Alert, AlertDescription, AlertTitle } from "@/components/ui/alert";
import { Button } from "@/components/ui/button";
import { Card, CardContent, CardDescription, CardHeader, CardTitle } from "@/components/ui/card";
import { Input } from "@/components/ui/input";
import { Label } from "@/components/ui/label";
import { Select, SelectContent, SelectItem, SelectTrigger, SelectValue } from "@/components/ui/select";
import { Switch } from "@/components/ui/switch";
import { Textarea } from "@/components/ui/textarea";
import { type JobFormValues } from "@/lib/forms/job";
import { cn, titleCase } from "@/lib/utils";

const CATEGORIES: ServiceCategory[] = [
  "MAKE_READY",
  "GENERAL_CONTRACTING",
  "PLUMBING",
  "ELECTRICAL",
  "HVAC",
  "FLOORING",
  "PAINTING",
  "DRYWALL",
  "APPLIANCE",
  "CLEANING",
  "LANDSCAPING",
  "ROOFING",
  "PEST_CONTROL",
  "WATER_MITIGATION",
  "OTHER",
];

const EMERGENCY_CATEGORIES: EmergencyCategory[] = [
  "AC_HVAC",
  "LEAK",
  "WATER_EXTRACTION",
  "ELECTRICAL",
  "OTHER",
];


function toIsoOrNull(localValue: string) {
  if (!localValue) return null;
  const d = new Date(localValue);
  return Number.isNaN(d.getTime()) ? null : d.toISOString();
}

export function JobForm({
  properties,
  initial,
  mode,
  defaultResponseMinutes = 15,
}: {
  properties: { id: string; name: string; city: string; state: string; isActive: boolean }[];
  initial: JobFormValues;
  mode: "create" | "edit";
  defaultResponseMinutes?: number;
}) {
  const router = useRouter();
  const [values, setValues] = React.useState<JobFormValues>(initial);
  const [error, setError] = React.useState<string | null>(null);
  const [pending, setPending] = React.useState<"draft" | "publish" | null>(null);

  const set = <K extends keyof JobFormValues>(key: K, value: JobFormValues[K]) =>
    setValues((prev) => ({ ...prev, [key]: value }));

  const isEmergency = values.priority === "EMERGENCY";

  async function submit(publish: boolean) {
    setPending(publish ? "publish" : "draft");
    setError(null);

    const payload: Record<string, unknown> = {
      propertyId: values.propertyId,
      unitNumber: values.unitNumber || null,
      title: values.title,
      description: values.description,
      category: values.category,
      budgetMin: Number(values.budgetMin),
      budgetMax: Number(values.budgetMax),
      enforceBudgetCap: isEmergency ? false : values.enforceBudgetCap,
      bidDeadline: isEmergency ? null : toIsoOrNull(values.bidDeadline),
      scheduledStart: toIsoOrNull(values.scheduledStart),
      dueDate: toIsoOrNull(values.dueDate),
      priority: values.priority,
      emergencyCategory: isEmergency ? values.emergencyCategory || null : null,
      responseDeadlineMinutes: isEmergency ? Number(values.responseDeadlineMinutes) : null,
      inviteOnly: isEmergency ? false : values.inviteOnly,
    };
    if (mode === "create") payload.status = publish ? "OPEN" : "DRAFT";

    const res = await fetch(mode === "create" ? "/api/admin/jobs" : `/api/admin/jobs/${initial.id}`, {
      method: mode === "create" ? "POST" : "PATCH",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(payload),
    });

    const data = await res.json().catch(() => null);
    if (!res.ok) {
      const issue = data?.error?.issues?.[0];
      setError(issue ? `${issue.path}: ${issue.message}` : (data?.error?.message ?? "Could not save the job."));
      setPending(null);
      return;
    }

    setPending(null);
    router.push(`/admin/jobs/${data.job.id}`);
    router.refresh();
  }

  return (
    <form
      onSubmit={(e) => {
        e.preventDefault();
        submit(mode === "edit" || isEmergency);
      }}
      className="grid min-w-0 gap-6 lg:grid-cols-3"
    >
      <div className="min-w-0 space-y-6 lg:col-span-2">
        {error && (
          <Alert variant="destructive">
            <AlertCircle />
            <AlertDescription>{error}</AlertDescription>
          </Alert>
        )}

        <Card>
          <CardHeader>
            <CardTitle>Scope</CardTitle>
            <CardDescription>What needs doing, and where.</CardDescription>
          </CardHeader>
          <CardContent className="space-y-4">
            <div className="grid gap-4 sm:grid-cols-2">
              <div className="space-y-2">
                <Label htmlFor="propertyId">Property</Label>
                <Select value={values.propertyId} onValueChange={(v) => set("propertyId", v)}>
                  <SelectTrigger id="propertyId">
                    <SelectValue placeholder="Choose a property" />
                  </SelectTrigger>
                  <SelectContent>
                    {properties
                      .filter((p) => p.isActive || p.id === values.propertyId)
                      .map((p) => (
                        <SelectItem key={p.id} value={p.id}>
                          {p.name} — {p.city}, {p.state}
                        </SelectItem>
                      ))}
                  </SelectContent>
                </Select>
              </div>

              <div className="space-y-2">
                <Label htmlFor="unitNumber">Unit / area</Label>
                <Input
                  id="unitNumber"
                  value={values.unitNumber}
                  onChange={(e) => set("unitNumber", e.target.value)}
                  placeholder="214, Clubhouse, Building C"
                />
              </div>
            </div>

            <div className="space-y-2">
              <Label htmlFor="title">Title</Label>
              <Input
                id="title"
                required
                value={values.title}
                onChange={(e) => set("title", e.target.value)}
                placeholder="Full make-ready turn — Unit 214"
              />
            </div>

            <div className="space-y-2">
              <Label htmlFor="description">Scope of work</Label>
              <Textarea
                id="description"
                required
                rows={5}
                value={values.description}
                onChange={(e) => set("description", e.target.value)}
                placeholder="Paint throughout, replace LVP in living room, deep clean, punch appliances…"
              />
            </div>

            <div className="space-y-2">
              <Label htmlFor="category">Trade</Label>
              <Select value={values.category} onValueChange={(v) => set("category", v as ServiceCategory)}>
                <SelectTrigger id="category">
                  <SelectValue />
                </SelectTrigger>
                <SelectContent>
                  {CATEGORIES.map((c) => (
                    <SelectItem key={c} value={c}>
                      {titleCase(c)}
                    </SelectItem>
                  ))}
                </SelectContent>
              </Select>
            </div>
          </CardContent>
        </Card>

        <Card>
          <CardHeader>
            <CardTitle>Budget</CardTitle>
            <CardDescription>Vendors see this range on the job board.</CardDescription>
          </CardHeader>
          <CardContent className="space-y-4">
            <div className="grid gap-4 sm:grid-cols-2">
              <div className="space-y-2">
                <Label htmlFor="budgetMin">Minimum</Label>
                <Input
                  id="budgetMin"
                  type="number"
                  min={0}
                  step="0.01"
                  required
                  value={values.budgetMin}
                  onChange={(e) => set("budgetMin", e.target.value)}
                  placeholder="1800"
                />
              </div>
              <div className="space-y-2">
                <Label htmlFor="budgetMax">Maximum</Label>
                <Input
                  id="budgetMax"
                  type="number"
                  min={0}
                  step="0.01"
                  required
                  value={values.budgetMax}
                  onChange={(e) => set("budgetMax", e.target.value)}
                  placeholder="3200"
                />
              </div>
            </div>

            {!isEmergency && (
              <div className="flex items-start justify-between gap-4 rounded-lg border p-3">
                <div>
                  <Label htmlFor="enforceBudgetCap">Enforce budget cap</Label>
                  <p className="text-xs text-muted-foreground">
                    Bids above the maximum are rejected at submission. Leave off to allow over-budget bids — those
                    get flagged at approval instead.
                  </p>
                </div>
                <Switch
                  id="enforceBudgetCap"
                  checked={values.enforceBudgetCap}
                  onCheckedChange={(v) => set("enforceBudgetCap", v)}
                />
              </div>
            )}
          </CardContent>
        </Card>

        <Card>
          <CardHeader>
            <CardTitle>Timing</CardTitle>
            <CardDescription>
              {isEmergency
                ? "Emergency jobs are claimed on a first-response basis — no bid deadline."
                : "Bidding closes automatically at the deadline."}
            </CardDescription>
          </CardHeader>
          <CardContent className="grid gap-4 sm:grid-cols-3">
            {!isEmergency && (
              <div className="space-y-2">
                <Label htmlFor="bidDeadline">Bid deadline</Label>
                <Input
                  id="bidDeadline"
                  type="datetime-local"
                  value={values.bidDeadline}
                  onChange={(e) => set("bidDeadline", e.target.value)}
                />
              </div>
            )}
            <div className="space-y-2">
              <Label htmlFor="scheduledStart">Scheduled start</Label>
              <Input
                id="scheduledStart"
                type="date"
                value={values.scheduledStart}
                onChange={(e) => set("scheduledStart", e.target.value)}
              />
            </div>
            <div className="space-y-2">
              <Label htmlFor="dueDate">Due date</Label>
              <Input
                id="dueDate"
                type="date"
                value={values.dueDate}
                onChange={(e) => set("dueDate", e.target.value)}
              />
            </div>
          </CardContent>
        </Card>
      </div>

      <div className="min-w-0 space-y-6">
        <Card className={cn(isEmergency && "border-emergency-border bg-emergency-soft")}>
          <CardHeader>
            <CardTitle className="flex items-center gap-2">
              {isEmergency && <Siren className="h-4 w-4 text-emergency" />}
              Priority
            </CardTitle>
            <CardDescription>
              Emergency jobs skip bidding entirely and go straight to eligible vendors.
            </CardDescription>
          </CardHeader>
          <CardContent className="space-y-4">
            <div className="grid grid-cols-2 gap-2">
              <Button
                type="button"
                variant={values.priority === "STANDARD" ? "default" : "outline"}
                onClick={() => set("priority", "STANDARD")}
              >
                Standard
              </Button>
              <Button
                type="button"
                variant={isEmergency ? "emergency" : "outline"}
                onClick={() => {
                  set("priority", "EMERGENCY");
                  set("responseDeadlineMinutes", String(defaultResponseMinutes));
                }}
              >
                Emergency
              </Button>
            </div>

            {isEmergency && (
              <>
                <div className="space-y-2">
                  <Label htmlFor="emergencyCategory">Category</Label>
                  <Select
                    value={values.emergencyCategory}
                    onValueChange={(v) => set("emergencyCategory", v as EmergencyCategory)}
                  >
                    <SelectTrigger id="emergencyCategory">
                      <SelectValue placeholder="Choose a category" />
                    </SelectTrigger>
                    <SelectContent>
                      {EMERGENCY_CATEGORIES.map((c) => (
                        <SelectItem key={c} value={c}>
                          {titleCase(c)}
                        </SelectItem>
                      ))}
                    </SelectContent>
                  </Select>
                </div>

                <div className="space-y-2">
                  <Label htmlFor="responseDeadlineMinutes">Response window (minutes)</Label>
                  <Input
                    id="responseDeadlineMinutes"
                    type="number"
                    min={5}
                    max={240}
                    value={values.responseDeadlineMinutes}
                    onChange={(e) => set("responseDeadlineMinutes", e.target.value)}
                  />
                  <p className="text-xs text-muted-foreground">
                    If nobody claims it in this window, the radius widens and an admin is alerted.
                  </p>
                </div>

                <Alert variant="emergency">
                  <Siren />
                  <AlertTitle>Dispatched immediately</AlertTitle>
                  <AlertDescription>
                    Only emergency-eligible, compliant, active vendors in range will see or be able to claim this job.
                  </AlertDescription>
                </Alert>
              </>
            )}
          </CardContent>
        </Card>

        {!isEmergency && (
          <Card>
            <CardHeader>
              <CardTitle>Visibility</CardTitle>
            </CardHeader>
            <CardContent>
              <div className="flex items-start justify-between gap-4">
                <div>
                  <Label htmlFor="inviteOnly">Invite specific vendors</Label>
                  <p className="text-xs text-muted-foreground">
                    Hides the job from the open board — useful for GC work with a shortlist.
                  </p>
                </div>
                <Switch
                  id="inviteOnly"
                  checked={values.inviteOnly}
                  onCheckedChange={(v) => set("inviteOnly", v)}
                />
              </div>
            </CardContent>
          </Card>
        )}

        <div className="flex flex-col gap-2">
          {mode === "create" ? (
            <>
              <Button type="submit" disabled={pending !== null}>
                {pending === "publish" && <Loader2 className="animate-spin" />}
                {isEmergency ? "Create & dispatch" : "Publish job"}
              </Button>
              {!isEmergency && (
                <Button type="button" variant="outline" disabled={pending !== null} onClick={() => submit(false)}>
                  {pending === "draft" && <Loader2 className="animate-spin" />}
                  Save as draft
                </Button>
              )}
            </>
          ) : (
            <Button type="submit" disabled={pending !== null}>
              {pending !== null && <Loader2 className="animate-spin" />}
              Save changes
            </Button>
          )}
        </div>
      </div>
    </form>
  );
}
