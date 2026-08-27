import { NextResponse } from "next/server";
import { withVendor } from "@/lib/auth/route";

export const dynamic = "force-dynamic";

export const GET = withVendor(async (_req, { actor }) =>
  NextResponse.json({
    scope: actor.scope,
    vendorUserId: actor.vendorUserId,
    vendorId: actor.vendorId,
    email: actor.email,
    companyName: actor.vendor.companyName,
    accountStatus: actor.vendor.accountStatus,
    complianceStatus: actor.vendor.complianceStatus,
    emergencyEligible: actor.vendor.emergencyEligible,
  }),
);
