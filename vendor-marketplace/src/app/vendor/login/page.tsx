import { Suspense } from "react";
import Link from "next/link";
import { redirect } from "next/navigation";
import type { Metadata } from "next";
import { HardHat } from "lucide-react";
import { LoginForm } from "@/components/login-form";
import { Card, CardContent, CardDescription, CardHeader, CardTitle } from "@/components/ui/card";
import { Skeleton } from "@/components/ui/skeleton";
import { getVendorActor } from "@/lib/auth/session";

export const metadata: Metadata = { title: "Vendor sign in" };
export const dynamic = "force-dynamic";

export default async function VendorLoginPage() {
  if (await getVendorActor()) redirect("/vendor");

  return (
    <main className="flex min-h-dvh items-center justify-center bg-secondary/40 px-4 py-12">
      <div className="w-full max-w-sm space-y-6">
        <div className="flex items-center gap-2">
          <div className="flex h-9 w-9 items-center justify-center rounded-lg bg-primary text-primary-foreground">
            <HardHat className="h-5 w-5" />
          </div>
          <span className="text-lg font-semibold tracking-tight">Vendor Marketplace</span>
        </div>

        <Card>
          <CardHeader>
            <CardTitle>Vendor sign in</CardTitle>
            <CardDescription>Suppliers and contractors.</CardDescription>
          </CardHeader>
          <CardContent>
            <Suspense fallback={<Skeleton className="h-48 w-full" />}>
              <LoginForm surface="vendor" defaultPath="/vendor" />
            </Suspense>
          </CardContent>
        </Card>

        <p className="text-center text-sm text-muted-foreground">
          Property management staff?{" "}
          <Link href="/admin/login" className="font-medium text-primary hover:underline">
            Sign in here
          </Link>
        </p>
      </div>
    </main>
  );
}
