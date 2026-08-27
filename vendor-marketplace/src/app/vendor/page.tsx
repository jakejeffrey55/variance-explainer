import { redirect } from "next/navigation";
import type { Metadata } from "next";
import { Clock, HardHat, ShieldAlert, ShieldCheck, Siren } from "lucide-react";
import { SignOutButton } from "@/components/sign-out-button";
import { AccountStatusBadge, ComplianceBadge } from "@/components/status-badges";
import { Alert, AlertDescription, AlertTitle } from "@/components/ui/alert";
import { Badge } from "@/components/ui/badge";
import { Card, CardContent, CardDescription, CardHeader, CardTitle } from "@/components/ui/card";
import { getVendorActor } from "@/lib/auth/session";
import { formatDate, titleCase } from "@/lib/utils";

export const metadata: Metadata = { title: "Vendor" };
export const dynamic = "force-dynamic";

/**
 * Vendor home. The full vendor experience — profile, filtered job board,
 * bidding — is built in the next phase; what exists today is the account gate
 * itself, which is the part that has to be right before any job data is shown.
 */
export default async function VendorHomePage() {
  const actor = await getVendorActor();
  if (!actor) redirect("/vendor/login");

  const { vendor } = actor;
  const pending = vendor.accountStatus === "PENDING";
  const complianceBlocked = vendor.complianceStatus === "EXPIRED" || vendor.complianceStatus === "NOT_SUBMITTED";

  return (
    <main className="mx-auto min-h-dvh w-full max-w-3xl px-4 py-10 sm:px-6">
      <div className="mb-8 flex items-start justify-between gap-4">
        <div className="flex items-center gap-3">
          <div className="flex h-10 w-10 items-center justify-center rounded-lg bg-primary text-primary-foreground">
            <HardHat className="h-5 w-5" />
          </div>
          <div>
            <h1 className="text-xl font-semibold tracking-tight">{vendor.companyName}</h1>
            <p className="text-sm text-muted-foreground">Signed in as {actor.name}</p>
          </div>
        </div>
        <SignOutButton surface="vendor" />
      </div>

      {pending && (
        <Alert variant="warning" className="mb-6">
          <Clock />
          <AlertTitle>Your account is pending approval</AlertTitle>
          <AlertDescription>
            A property management administrator reviews every new vendor. You will be notified once your account is
            approved — until then you cannot view or bid on jobs.
          </AlertDescription>
        </Alert>
      )}

      {!pending && complianceBlocked && (
        <Alert variant="destructive" className="mb-6">
          <ShieldAlert />
          <AlertTitle>Compliance documents need attention</AlertTitle>
          <AlertDescription>
            Your credentialing is {titleCase(vendor.complianceStatus).toLowerCase()}. You can view jobs, but you
            cannot submit new bids until your documents are current.
          </AlertDescription>
        </Alert>
      )}

      <div className="grid gap-4 sm:grid-cols-2">
        <Card>
          <CardHeader>
            <CardTitle className="text-base">Account</CardTitle>
            <CardDescription>Set by property management.</CardDescription>
          </CardHeader>
          <CardContent className="space-y-3 text-sm">
            <div className="flex items-center justify-between">
              <span className="text-muted-foreground">Status</span>
              <AccountStatusBadge status={vendor.accountStatus} />
            </div>
            <div className="flex items-center justify-between">
              <span className="text-muted-foreground">Compliance</span>
              <ComplianceBadge status={vendor.complianceStatus} />
            </div>
            <div className="flex items-center justify-between">
              <span className="text-muted-foreground">Expires</span>
              <span className="font-medium">{formatDate(vendor.complianceExpiresAt)}</span>
            </div>
            <div className="flex items-center justify-between">
              <span className="text-muted-foreground">Emergency dispatch</span>
              {vendor.emergencyEligible ? (
                <Badge variant="emergency">
                  <Siren className="h-3 w-3" /> Eligible
                </Badge>
              ) : (
                <Badge variant="muted">Not enabled</Badge>
              )}
            </div>
          </CardContent>
        </Card>

        <Card>
          <CardHeader>
            <CardTitle className="text-base">Service profile</CardTitle>
            <CardDescription>Used to match you to jobs.</CardDescription>
          </CardHeader>
          <CardContent className="space-y-3 text-sm">
            <div className="flex items-center justify-between">
              <span className="text-muted-foreground">Service radius</span>
              <span className="font-medium">{vendor.serviceRadiusMiles} miles</span>
            </div>
            <div className="space-y-1.5">
              <span className="text-muted-foreground">Trades</span>
              <div className="flex flex-wrap gap-1.5">
                {vendor.serviceCategories.length === 0 ? (
                  <span className="text-muted-foreground">None selected</span>
                ) : (
                  vendor.serviceCategories.map((c) => (
                    <Badge key={c} variant="secondary">
                      {titleCase(c)}
                    </Badge>
                  ))
                )}
              </div>
            </div>
            {vendor.trustScore !== null && (
              <div className="flex items-center justify-between">
                <span className="text-muted-foreground">Trust score</span>
                <span className="font-medium tabular-nums">{vendor.trustScore.toFixed(0)}</span>
              </div>
            )}
          </CardContent>
        </Card>
      </div>

      <Card className="mt-4">
        <CardHeader>
          <CardTitle className="flex items-center gap-2 text-base">
            <ShieldCheck className="h-4 w-4" /> Job board
          </CardTitle>
        </CardHeader>
        <CardContent className="text-sm text-muted-foreground">
          {pending
            ? "Jobs matched to your trades and service area will appear here once your account is approved."
            : "Your matched job board, bidding, and emergency claims are being built next. Your account gates above are already enforced server-side."}
        </CardContent>
      </Card>
    </main>
  );
}
