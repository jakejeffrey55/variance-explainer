import { NextResponse } from "next/server";
import { withAdmin } from "@/lib/auth/route";
import { listPropertyProviderStatuses } from "@/lib/integrations/property";

export const dynamic = "force-dynamic";

/** Which integration adapters are live and which are still stubbed. */
export const GET = withAdmin(async () =>
  NextResponse.json({ property: listPropertyProviderStatuses() }),
);
