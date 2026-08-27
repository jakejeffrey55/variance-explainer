import { getServerSession } from "next-auth";
import type { AdminRole, ComplianceStatus, Vendor } from "@prisma/client";
import { adminAuthOptions, vendorAuthOptions } from "@/lib/auth/options";
import { prisma } from "@/lib/db";

/**
 * Server-side authorization primitives.
 *
 * Rule for the whole codebase: no route handler, server action, or server
 * component may read the database for a vendor without going through
 * `requireVendor*` and filtering on the returned `vendorId`. The UI is never
 * the enforcement point — these functions are.
 */

export class AuthError extends Error {
  constructor(
    readonly status: 401 | 403,
    readonly code: string,
    message: string,
  ) {
    super(message);
    this.name = "AuthError";
  }
}

export type AdminActor = {
  scope: "admin";
  adminUserId: string;
  email: string;
  name: string;
  role: AdminRole;
};

export type VendorActor = {
  scope: "vendor";
  vendorUserId: string;
  vendorId: string;
  email: string;
  name: string;
  vendor: Vendor;
};

export async function getAdminActor(): Promise<AdminActor | null> {
  const session = await getServerSession(adminAuthOptions);
  if (!session?.user || session.user.scope !== "admin") return null;

  // Re-read from the database so deactivation takes effect immediately rather
  // than at token expiry.
  const admin = await prisma.adminUser.findUnique({
    where: { id: session.user.id },
  });
  if (!admin || !admin.isActive) return null;

  return {
    scope: "admin",
    adminUserId: admin.id,
    email: admin.email,
    name: admin.name,
    role: admin.role,
  };
}

export async function requireAdmin(): Promise<AdminActor> {
  const actor = await getAdminActor();
  if (!actor) throw new AuthError(401, "admin_auth_required", "Admin sign-in required.");
  return actor;
}

export async function requireAdminRole(...roles: AdminRole[]): Promise<AdminActor> {
  const actor = await requireAdmin();
  if (roles.length > 0 && !roles.includes(actor.role)) {
    throw new AuthError(403, "insufficient_admin_role", "You do not have access to this action.");
  }
  return actor;
}

export async function getVendorActor(): Promise<VendorActor | null> {
  const session = await getServerSession(vendorAuthOptions);
  if (!session?.user || session.user.scope !== "vendor" || !session.user.vendorId) return null;

  const vendorUser = await prisma.vendorUser.findUnique({
    where: { id: session.user.id },
    include: { vendor: true },
  });
  if (!vendorUser || !vendorUser.isActive) return null;
  // Defence in depth: the vendor id in the token must still match the row.
  if (vendorUser.vendorId !== session.user.vendorId) return null;

  return {
    scope: "vendor",
    vendorUserId: vendorUser.id,
    vendorId: vendorUser.vendorId,
    email: vendorUser.email,
    name: vendorUser.name,
    vendor: vendorUser.vendor,
  };
}

export async function requireVendor(): Promise<VendorActor> {
  const actor = await getVendorActor();
  if (!actor) throw new AuthError(401, "vendor_auth_required", "Vendor sign-in required.");
  if (actor.vendor.accountStatus === "SUSPENDED" || actor.vendor.accountStatus === "REJECTED") {
    throw new AuthError(403, "vendor_account_blocked", "This vendor account is not active.");
  }
  return actor;
}

/** Vendor whose signup an admin has approved — the gate for seeing any job. */
export async function requireApprovedVendor(): Promise<VendorActor> {
  const actor = await requireVendor();
  if (actor.vendor.accountStatus !== "ACTIVE") {
    throw new AuthError(
      403,
      "vendor_pending_approval",
      "Your account is pending approval. You will be notified once an administrator approves it.",
    );
  }
  return actor;
}

const BID_BLOCKING_COMPLIANCE: ComplianceStatus[] = ["EXPIRED", "NOT_SUBMITTED"];

/** Approved vendor whose credentialing is current — the gate for submitting bids. */
export async function requireBiddingVendor(): Promise<VendorActor> {
  const actor = await requireApprovedVendor();
  if (BID_BLOCKING_COMPLIANCE.includes(actor.vendor.complianceStatus)) {
    throw new AuthError(
      403,
      "vendor_compliance_expired",
      "Your compliance documents are expired or missing. Update them before submitting new bids.",
    );
  }
  return actor;
}

/** Bidding-eligible vendor that an admin has additionally flagged for emergencies. */
export async function requireEmergencyVendor(): Promise<VendorActor> {
  const actor = await requireBiddingVendor();
  if (!actor.vendor.emergencyEligible) {
    throw new AuthError(
      403,
      "vendor_not_emergency_eligible",
      "Your account is not enabled for emergency dispatch.",
    );
  }
  return actor;
}
