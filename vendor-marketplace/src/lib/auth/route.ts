import { NextResponse } from "next/server";
import { ZodError } from "zod";
import {
  AuthError,
  type AdminActor,
  type VendorActor,
  requireAdmin,
  requireApprovedVendor,
  requireBiddingVendor,
  requireEmergencyVendor,
  requireVendor,
} from "@/lib/auth/session";

type RouteContext = { params: Record<string, string> };

type Handler<A> = (
  req: Request,
  ctx: RouteContext & { actor: A },
) => Promise<Response> | Response;

export function apiError(status: number, code: string, message: string, extra?: object) {
  return NextResponse.json({ error: { code, message, ...extra } }, { status });
}

function toResponse(err: unknown) {
  if (err instanceof AuthError) return apiError(err.status, err.code, err.message);
  if (err instanceof ZodError) {
    return apiError(422, "invalid_request", "Invalid request payload.", {
      issues: err.issues.map((i) => ({ path: i.path.join("."), message: i.message })),
    });
  }
  console.error("[api] unhandled error", err);
  return apiError(500, "internal_error", "Something went wrong. Please try again.");
}

function wrap<A>(getActor: () => Promise<A>, handler: Handler<A>) {
  return async (req: Request, ctx: RouteContext = { params: {} }) => {
    try {
      const actor = await getActor();
      return await handler(req, { ...ctx, actor });
    } catch (err) {
      return toResponse(err);
    }
  };
}

/** Admin-only API route. */
export const withAdmin = (handler: Handler<AdminActor>) => wrap(requireAdmin, handler);

/** Any signed-in vendor, including one still pending approval. */
export const withVendor = (handler: Handler<VendorActor>) => wrap(requireVendor, handler);

/** Vendor whose account an admin has approved. */
export const withApprovedVendor = (handler: Handler<VendorActor>) =>
  wrap(requireApprovedVendor, handler);

/** Approved vendor with current compliance — required to submit a bid. */
export const withBiddingVendor = (handler: Handler<VendorActor>) =>
  wrap(requireBiddingVendor, handler);

/** Emergency-eligible vendor — required to see or claim emergency dispatch. */
export const withEmergencyVendor = (handler: Handler<VendorActor>) =>
  wrap(requireEmergencyVendor, handler);
