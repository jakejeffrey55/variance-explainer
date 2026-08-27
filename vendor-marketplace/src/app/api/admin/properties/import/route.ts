import { NextResponse } from "next/server";
import { apiError, withAdmin } from "@/lib/auth/route";
import { prisma } from "@/lib/db";
import { importProperties } from "@/lib/services/property-import";

export const dynamic = "force-dynamic";

const MAX_UPLOAD_BYTES = 5 * 1024 * 1024;

/** Import history, for the panel under the upload form. */
export const GET = withAdmin(async () => {
  const batches = await prisma.importBatch.findMany({
    where: { kind: "PROPERTY" },
    orderBy: { startedAt: "desc" },
    take: 15,
    include: { adminUser: { select: { name: true } } },
  });
  return NextResponse.json({ batches });
});

export const POST = withAdmin(async (req, { actor }) => {
  const form = await req.formData();
  const providerKey = String(form.get("providerKey") ?? "csv");
  const dryRun = String(form.get("dryRun") ?? "false") === "true";
  const file = form.get("file");

  let fileContent: string | undefined;
  let fileName: string | undefined;

  if (file instanceof File) {
    if (file.size > MAX_UPLOAD_BYTES) {
      return apiError(413, "file_too_large", "That file is larger than 5 MB. Split it and import in parts.");
    }
    fileContent = await file.text();
    fileName = file.name;
  }

  const outcome = await importProperties({
    providerKey,
    input: { fileContent, fileName },
    adminUserId: actor.adminUserId,
    dryRun,
  });

  if (!outcome.ok) {
    // A provider that is not wired up yet is a 501, not a client error.
    const status = outcome.error.code === "not_configured" ? 501 : 422;
    return NextResponse.json({ error: outcome.error }, { status });
  }

  return NextResponse.json({ result: outcome.result });
});
