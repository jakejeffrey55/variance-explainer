import { Suspense } from "react";
import Link from "next/link";
import { redirect } from "next/navigation";
import type { Metadata } from "next";
import { Building2 } from "lucide-react";
import { LoginForm } from "@/components/login-form";
import { Card, CardContent, CardDescription, CardHeader, CardTitle } from "@/components/ui/card";
import { Skeleton } from "@/components/ui/skeleton";
import { getAdminActor } from "@/lib/auth/session";

export const metadata: Metadata = { title: "Admin sign in" };
export const dynamic = "force-dynamic";

export default async function AdminLoginPage() {
  if (await getAdminActor()) redirect("/admin");

  return (
    <main className="flex min-h-dvh items-center justify-center bg-secondary/40 px-4 py-12">
      <div className="w-full max-w-sm space-y-6">
        <div className="flex items-center gap-2">
          <div className="flex h-9 w-9 items-center justify-center rounded-lg bg-primary text-primary-foreground">
            <Building2 className="h-5 w-5" />
          </div>
          <span className="text-lg font-semibold tracking-tight">Vendor Marketplace</span>
        </div>

        <Card>
          <CardHeader>
            <CardTitle>Admin sign in</CardTitle>
            <CardDescription>Property management staff only.</CardDescription>
          </CardHeader>
          <CardContent>
            <Suspense fallback={<Skeleton className="h-48 w-full" />}>
              <LoginForm surface="admin" defaultPath="/admin" />
            </Suspense>
          </CardContent>
        </Card>

        <p className="text-center text-sm text-muted-foreground">
          Vendor?{" "}
          <Link href="/vendor/login" className="font-medium text-primary hover:underline">
            Sign in here
          </Link>
        </p>
      </div>
    </main>
  );
}
