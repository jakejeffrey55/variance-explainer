import { NextResponse } from "next/server";
import { apiError, withAdmin } from "@/lib/auth/route";
import { prisma } from "@/lib/db";
import { logActivity } from "@/lib/services/activity";
import { propertyUpdateSchema } from "@/lib/validation/property";

export const dynamic = "force-dynamic";

export const GET = withAdmin(async (_req, { params }) => {
  const property = await prisma.property.findUnique({
    where: { id: params.id },
    include: {
      jobs: { orderBy: { createdAt: "desc" }, take: 50 },
      rentRolls: { orderBy: { unitNumber: "asc" } },
    },
  });
  if (!property) return apiError(404, "not_found", "Property not found.");
  return NextResponse.json({ property });
});

export const PATCH = withAdmin(async (req, { params, actor }) => {
  const existing = await prisma.property.findUnique({ where: { id: params.id } });
  if (!existing) return apiError(404, "not_found", "Property not found.");

  const input = propertyUpdateSchema.parse(await req.json());
  const property = await prisma.property.update({
    where: { id: params.id },
    data: { ...input, propertyManagerEmail: input.propertyManagerEmail || null },
  });

  await logActivity({
    entityType: "PROPERTY",
    entityId: property.id,
    action: "property.updated",
    actorType: "ADMIN",
    actorAdminId: actor.adminUserId,
    summary: `Property ${property.name} updated.`,
    metadata: { fields: Object.keys(input) },
  });

  return NextResponse.json({ property });
});

/**
 * Properties with history are deactivated rather than deleted — jobs, bids and
 * approval flags reference them.
 */
export const DELETE = withAdmin(async (_req, { params, actor }) => {
  const property = await prisma.property.findUnique({
    where: { id: params.id },
    include: { _count: { select: { jobs: true } } },
  });
  if (!property) return apiError(404, "not_found", "Property not found.");

  if (property._count.jobs > 0) {
    const deactivated = await prisma.property.update({
      where: { id: params.id },
      data: { isActive: false },
    });
    await logActivity({
      entityType: "PROPERTY",
      entityId: property.id,
      action: "property.deactivated",
      actorType: "ADMIN",
      actorAdminId: actor.adminUserId,
      summary: `Property ${property.name} deactivated (${property._count.jobs} job${property._count.jobs === 1 ? "" : "s"} in history).`,
    });
    return NextResponse.json({ property: deactivated, deactivated: true });
  }

  await prisma.property.delete({ where: { id: params.id } });
  await logActivity({
    entityType: "PROPERTY",
    entityId: property.id,
    action: "property.deleted",
    actorType: "ADMIN",
    actorAdminId: actor.adminUserId,
    summary: `Property ${property.name} deleted.`,
  });
  return NextResponse.json({ deleted: true });
});
