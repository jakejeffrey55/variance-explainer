import { redirect } from "next/navigation";
import { AdminShell } from "@/components/admin/admin-shell";
import { getAdminActor } from "@/lib/auth/session";

export const dynamic = "force-dynamic";

/**
 * Server-side gate for the whole admin surface. Middleware already bounces
 * unauthenticated navigation, but this is the check that actually decides —
 * it re-reads the admin from the database on every request.
 */
export default async function AdminLayout({ children }: { children: React.ReactNode }) {
  const actor = await getAdminActor();
  if (!actor) redirect("/admin/login");

  return (
    <AdminShell admin={{ name: actor.name, email: actor.email, role: actor.role }}>{children}</AdminShell>
  );
}
