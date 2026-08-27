import { NextResponse } from "next/server";
import { Prisma } from "@prisma/client";
import { withAdmin } from "@/lib/auth/route";
import { prisma } from "@/lib/db";
import { logActivity } from "@/lib/services/activity";
import { propertyInputSchema } from "@/lib/validation/property";

export const dynamic = "force-dynamic";

export const GET = withAdmin(async (req) => {
  const url = new URL(req.url);
  const q = url.searchParams.get("q")?.trim();
  const activeParam = url.searchParams.get("active");

  const where: Prisma.PropertyWhereInput = {
    ...(activeParam === "true" ? { isActive: true } : activeParam === "false" ? { isActive: false } : {}),
    ...(q
      ? {
          OR: [
            { name: { contains: q, mode: "insensitive" } },
            { city: { contains: q, mode: "insensitive" } },
            { externalId: { contains: q, mode: "insensitive" } },
            { addressLine1: { contains: q, mode: "insensitive" } },
          ],
        }
      : {}),
  };

  const properties = await prisma.property.findMany({
    where,
    orderBy: { name: "asc" },
    include: { _count: { select: { jobs: true, rentRolls: true } } },
  });

  return NextResponse.json({ properties });
});

export const POST = withAdmin(async (req, { actor }) => {
  const body = await req.json();
  const input = propertyInputSchema.parse(body);

  const property = await prisma.property.create({
    data: {
      ...input,
      propertyManagerEmail: input.propertyManagerEmail || null,
      source: "MANUAL",
      externalId: input.externalId || null,
    },
  });

  await logActivity({
    entityType: "PROPERTY",
    entityId: property.id,
    action: "property.created",
    actorType: "ADMIN",
    actorAdminId: actor.adminUserId,
    summary: `Property ${property.name} added manually.`,
  });

  return NextResponse.json({ property }, { status: 201 });
});
