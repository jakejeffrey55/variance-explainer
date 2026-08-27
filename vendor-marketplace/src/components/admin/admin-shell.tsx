"use client";

import * as React from "react";
import Link from "next/link";
import { usePathname, useRouter } from "next/navigation";
import { Building2, ClipboardList, LayoutDashboard, LogOut, Menu, Upload, X } from "lucide-react";
import { Button } from "@/components/ui/button";
import { signOutOfSurface } from "@/lib/auth/client";
import { cn } from "@/lib/utils";

const NAV = [
  { href: "/admin", label: "Dashboard", icon: LayoutDashboard, exact: true },
  { href: "/admin/jobs", label: "Jobs", icon: ClipboardList },
  { href: "/admin/properties", label: "Properties", icon: Building2 },
  { href: "/admin/import", label: "Import", icon: Upload },
];

export function AdminShell({
  admin,
  children,
}: {
  admin: { name: string; email: string; role: string };
  children: React.ReactNode;
}) {
  const pathname = usePathname();
  const router = useRouter();
  const [mobileOpen, setMobileOpen] = React.useState(false);

  const isActive = (href: string, exact?: boolean) =>
    exact ? pathname === href : pathname === href || pathname.startsWith(`${href}/`);

  async function handleSignOut() {
    await signOutOfSurface("admin");
    router.replace("/admin/login");
    router.refresh();
  }

  const nav = (
    <nav className="flex flex-col gap-1">
      {NAV.map((item) => {
        const Icon = item.icon;
        const active = isActive(item.href, item.exact);
        return (
          <Link
            key={item.href}
            href={item.href}
            onClick={() => setMobileOpen(false)}
            className={cn(
              "flex items-center gap-3 rounded-lg px-3 py-2 text-sm font-medium transition-colors",
              active
                ? "bg-primary/10 text-primary"
                : "text-muted-foreground hover:bg-accent hover:text-foreground",
            )}
            aria-current={active ? "page" : undefined}
          >
            <Icon className="h-4 w-4" />
            {item.label}
          </Link>
        );
      })}
    </nav>
  );

  return (
    <div className="min-h-dvh bg-secondary/30">
      {/* Desktop sidebar */}
      <aside className="fixed inset-y-0 left-0 hidden w-60 flex-col border-r bg-card px-4 py-5 lg:flex">
        <Link href="/admin" className="mb-6 flex items-center gap-2">
          <div className="flex h-8 w-8 items-center justify-center rounded-lg bg-primary text-primary-foreground">
            <Building2 className="h-4 w-4" />
          </div>
          <span className="font-semibold tracking-tight">Marketplace</span>
        </Link>
        {nav}
        <div className="mt-auto space-y-3 border-t pt-4">
          <div className="px-1">
            <p className="truncate text-sm font-medium">{admin.name}</p>
            <p className="truncate text-xs text-muted-foreground">{admin.role.toLowerCase()}</p>
          </div>
          <Button variant="ghost" size="sm" className="w-full justify-start" onClick={handleSignOut}>
            <LogOut className="h-4 w-4" />
            Sign out
          </Button>
        </div>
      </aside>

      {/* Mobile top bar */}
      <header className="sticky top-0 z-40 flex items-center justify-between border-b bg-card px-4 py-3 lg:hidden">
        <Link href="/admin" className="flex items-center gap-2">
          <div className="flex h-8 w-8 items-center justify-center rounded-lg bg-primary text-primary-foreground">
            <Building2 className="h-4 w-4" />
          </div>
          <span className="font-semibold tracking-tight">Marketplace</span>
        </Link>
        <Button variant="ghost" size="icon" onClick={() => setMobileOpen((v) => !v)} aria-label="Toggle navigation">
          {mobileOpen ? <X className="h-5 w-5" /> : <Menu className="h-5 w-5" />}
        </Button>
      </header>

      {mobileOpen && (
        <div className="border-b bg-card px-4 py-3 lg:hidden">
          {nav}
          <Button variant="ghost" size="sm" className="mt-2 w-full justify-start" onClick={handleSignOut}>
            <LogOut className="h-4 w-4" />
            Sign out ({admin.name})
          </Button>
        </div>
      )}

      <div className="lg:pl-60">
        <div className="mx-auto w-full max-w-7xl px-4 py-6 sm:px-6 lg:px-8">{children}</div>
      </div>
    </div>
  );
}

export function PageHeader({
  title,
  description,
  children,
}: {
  title: string;
  description?: string;
  children?: React.ReactNode;
}) {
  return (
    <div className="mb-6 flex flex-col gap-3 sm:flex-row sm:items-start sm:justify-between">
      <div className="space-y-1">
        <h1 className="text-2xl font-semibold tracking-tight">{title}</h1>
        {description && <p className="text-sm text-muted-foreground">{description}</p>}
      </div>
      {children && <div className="flex flex-wrap items-center gap-2">{children}</div>}
    </div>
  );
}
