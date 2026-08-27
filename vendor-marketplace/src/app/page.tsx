import Link from "next/link";
import { Building2, HardHat, ShieldCheck, Zap } from "lucide-react";
import { Button } from "@/components/ui/button";
import { Card, CardContent, CardDescription, CardHeader, CardTitle } from "@/components/ui/card";

export default function HomePage() {
  return (
    <main className="mx-auto flex min-h-dvh max-w-5xl flex-col justify-center gap-10 px-6 py-16">
      <div className="space-y-3">
        <div className="inline-flex items-center gap-2 rounded-full border bg-card px-3 py-1 text-xs font-medium text-muted-foreground">
          <Zap className="h-3.5 w-3.5 text-primary" />
          Make-ready · General contracting · Emergency dispatch
        </div>
        <h1 className="text-4xl font-semibold tracking-tight sm:text-5xl">Vendor Marketplace</h1>
        <p className="max-w-2xl text-lg text-muted-foreground">
          Post make-ready and contracting work, match it to credentialed vendors by service radius, and move
          approved jobs into procurement — with emergencies dispatched in minutes.
        </p>
      </div>

      <div className="grid gap-4 sm:grid-cols-2">
        <Card>
          <CardHeader>
            <div className="flex h-10 w-10 items-center justify-center rounded-lg bg-primary/10 text-primary">
              <Building2 className="h-5 w-5" />
            </div>
            <CardTitle>Property management</CardTitle>
            <CardDescription>
              Import properties, post jobs, review bids, approve work, and dispatch emergencies.
            </CardDescription>
          </CardHeader>
          <CardContent>
            <Button asChild className="w-full">
              <Link href="/admin/login">Admin sign in</Link>
            </Button>
          </CardContent>
        </Card>

        <Card>
          <CardHeader>
            <div className="flex h-10 w-10 items-center justify-center rounded-lg bg-primary/10 text-primary">
              <HardHat className="h-5 w-5" />
            </div>
            <CardTitle>Vendors</CardTitle>
            <CardDescription>
              See jobs matched to your trades and service area, bid, claim emergencies, and track your work.
            </CardDescription>
          </CardHeader>
          <CardContent>
            <Button asChild variant="outline" className="w-full">
              <Link href="/vendor/login">Vendor sign in</Link>
            </Button>
          </CardContent>
        </Card>
      </div>

      <p className="flex items-center gap-2 text-sm text-muted-foreground">
        <ShieldCheck className="h-4 w-4" />
        Admin and vendor accounts are entirely separate — different credentials, different sessions.
      </p>
    </main>
  );
}
