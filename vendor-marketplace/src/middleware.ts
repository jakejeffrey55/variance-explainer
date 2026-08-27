import { NextResponse, type NextRequest } from "next/server";
import { getToken } from "next-auth/jwt";

/**
 * Edge-level gate. This is a convenience redirect layer only — the real
 * enforcement lives in src/lib/auth/session.ts, which re-reads the actor from
 * the database on every request. Middleware never grants access; it only
 * bounces obviously unauthenticated navigation to the right login screen.
 */

const isProd = process.env.NODE_ENV === "production";
const cookiePrefix = isProd ? "__Secure-" : "";
const ADMIN_COOKIE = `${cookiePrefix}vm.admin.session-token`;
const VENDOR_COOKIE = `${cookiePrefix}vm.vendor.session-token`;

const PUBLIC_ADMIN_PATHS = ["/admin/login"];
const PUBLIC_VENDOR_PATHS = ["/vendor/login", "/vendor/signup", "/vendor/pending"];

export async function middleware(req: NextRequest) {
  const { pathname } = req.nextUrl;

  const surface = pathname.startsWith("/admin")
    ? ("admin" as const)
    : pathname.startsWith("/vendor")
      ? ("vendor" as const)
      : null;
  if (!surface) return NextResponse.next();

  const publicPaths = surface === "admin" ? PUBLIC_ADMIN_PATHS : PUBLIC_VENDOR_PATHS;
  if (publicPaths.some((p) => pathname === p || pathname.startsWith(`${p}/`))) {
    return NextResponse.next();
  }

  const token = await getToken({
    req,
    secret: process.env.NEXTAUTH_SECRET,
    cookieName: surface === "admin" ? ADMIN_COOKIE : VENDOR_COOKIE,
  });

  if (!token || token.scope !== surface) {
    const url = req.nextUrl.clone();
    url.pathname = surface === "admin" ? "/admin/login" : "/vendor/login";
    url.search = `?callbackUrl=${encodeURIComponent(pathname)}`;
    return NextResponse.redirect(url);
  }

  return NextResponse.next();
}

export const config = {
  matcher: ["/admin/:path*", "/vendor/:path*"],
};
