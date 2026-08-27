import { NextResponse } from "next/server";
import { withAdmin } from "@/lib/auth/route";

export const dynamic = "force-dynamic";

export const GET = withAdmin(async (_req, { actor }) =>
  NextResponse.json({
    scope: actor.scope,
    adminUserId: actor.adminUserId,
    email: actor.email,
    name: actor.name,
    role: actor.role,
  }),
);
