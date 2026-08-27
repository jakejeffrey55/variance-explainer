import type { NextAuthOptions } from "next-auth";
import CredentialsProvider from "next-auth/providers/credentials";
import bcrypt from "bcryptjs";
import { prisma } from "@/lib/db";

/**
 * Two entirely separate authentication surfaces.
 *
 *  - Admins authenticate at /api/auth/admin/*  and receive the cookie
 *    `vm.admin.session-token`, carrying scope "admin".
 *  - Vendors authenticate at /api/auth/vendor/* and receive the cookie
 *    `vm.vendor.session-token`, carrying scope "vendor".
 *
 * The two credential stores are different tables (admin_users / vendor_users),
 * so an email that exists on one surface cannot be used on the other, and a
 * cookie issued by one surface is invisible to the other (different names,
 * different signing payloads). Nothing about a request's *route* grants access:
 * every guard in src/lib/auth/session.ts re-reads the actor from the database.
 */

export type SessionScope = "admin" | "vendor";

const AUTH_SECRET = process.env.NEXTAUTH_SECRET;

const isProd = process.env.NODE_ENV === "production";
const cookiePrefix = isProd ? "__Secure-" : "";

function cookiesFor(scope: SessionScope): NextAuthOptions["cookies"] {
  return {
    sessionToken: {
      name: `${cookiePrefix}vm.${scope}.session-token`,
      options: {
        httpOnly: true,
        sameSite: "lax",
        path: "/",
        secure: isProd,
      },
    },
    csrfToken: {
      name: `${isProd ? "__Host-" : ""}vm.${scope}.csrf-token`,
      options: { httpOnly: true, sameSite: "lax", path: "/", secure: isProd },
    },
    callbackUrl: {
      name: `${cookiePrefix}vm.${scope}.callback-url`,
      options: { sameSite: "lax", path: "/", secure: isProd },
    },
  };
}

const credentialsSchema = {
  email: { label: "Email", type: "email" },
  password: { label: "Password", type: "password" },
} as const;

function normalizeEmail(email: unknown): string | null {
  if (typeof email !== "string") return null;
  const trimmed = email.trim().toLowerCase();
  return trimmed.length > 0 ? trimmed : null;
}

/**
 * Constant-ish work regardless of whether the account exists, so response time
 * does not disclose which emails are registered.
 */
const DUMMY_HASH = "$2a$10$N9qo8uLOickgx2ZMRZoMyeIjZAgcfl7p92ldGxad68LJZdL17lhWy";

async function verifyPassword(password: unknown, hash: string | undefined) {
  if (typeof password !== "string" || password.length === 0) {
    await bcrypt.compare("x", DUMMY_HASH);
    return false;
  }
  return bcrypt.compare(password, hash ?? DUMMY_HASH);
}

export const adminAuthOptions: NextAuthOptions = {
  secret: AUTH_SECRET,
  session: { strategy: "jwt", maxAge: 60 * 60 * 12 },
  cookies: cookiesFor("admin"),
  pages: { signIn: "/admin/login", error: "/admin/login" },
  providers: [
    CredentialsProvider({
      id: "admin-credentials",
      name: "Admin",
      credentials: credentialsSchema,
      async authorize(credentials) {
        const email = normalizeEmail(credentials?.email);
        if (!email) return null;

        const admin = await prisma.adminUser.findUnique({ where: { email } });
        const ok = await verifyPassword(credentials?.password, admin?.passwordHash);
        if (!admin || !ok || !admin.isActive) return null;

        await prisma.adminUser.update({
          where: { id: admin.id },
          data: { lastLoginAt: new Date() },
        });

        return {
          id: admin.id,
          email: admin.email,
          name: admin.name,
          scope: "admin" as const,
          adminRole: admin.role,
        };
      },
    }),
  ],
  callbacks: {
    async jwt({ token, user }) {
      if (user) {
        token.scope = "admin";
        token.adminRole = (user as { adminRole?: string }).adminRole;
      }
      // Reject any token that was not minted by this surface.
      if (token.scope !== "admin") return {};
      return token;
    },
    async session({ session, token }) {
      if (token.scope !== "admin") {
        // Never hand back a usable session for a foreign scope.
        return { ...session, user: undefined, expires: session.expires } as typeof session;
      }
      session.user = {
        id: token.sub as string,
        email: token.email as string,
        name: token.name as string,
        scope: "admin",
        adminRole: token.adminRole as string | undefined,
      };
      return session;
    },
  },
};

export const vendorAuthOptions: NextAuthOptions = {
  secret: AUTH_SECRET,
  session: { strategy: "jwt", maxAge: 60 * 60 * 24 * 7 },
  cookies: cookiesFor("vendor"),
  pages: { signIn: "/vendor/login", error: "/vendor/login" },
  providers: [
    CredentialsProvider({
      id: "vendor-credentials",
      name: "Vendor",
      credentials: credentialsSchema,
      async authorize(credentials) {
        const email = normalizeEmail(credentials?.email);
        if (!email) return null;

        const vendorUser = await prisma.vendorUser.findUnique({
          where: { email },
          include: { vendor: { select: { id: true, accountStatus: true } } },
        });
        const ok = await verifyPassword(credentials?.password, vendorUser?.passwordHash);
        if (!vendorUser || !ok || !vendorUser.isActive) return null;
        // SUSPENDED/REJECTED vendors cannot hold a session at all. PENDING
        // vendors *can* sign in — they land on a "waiting for approval" state —
        // but every job/bid guard denies them (see requireApprovedVendor).
        if (
          vendorUser.vendor.accountStatus === "SUSPENDED" ||
          vendorUser.vendor.accountStatus === "REJECTED"
        ) {
          return null;
        }

        await prisma.vendorUser.update({
          where: { id: vendorUser.id },
          data: { lastLoginAt: new Date() },
        });

        return {
          id: vendorUser.id,
          email: vendorUser.email,
          name: vendorUser.name,
          scope: "vendor" as const,
          vendorId: vendorUser.vendorId,
        };
      },
    }),
  ],
  callbacks: {
    async jwt({ token, user }) {
      if (user) {
        token.scope = "vendor";
        token.vendorId = (user as { vendorId?: string }).vendorId;
      }
      if (token.scope !== "vendor") return {};
      return token;
    },
    async session({ session, token }) {
      if (token.scope !== "vendor" || !token.vendorId) {
        return { ...session, user: undefined, expires: session.expires } as typeof session;
      }
      session.user = {
        id: token.sub as string,
        email: token.email as string,
        name: token.name as string,
        scope: "vendor",
        // The vendor id is bound to the session at sign-in and is never taken
        // from a request parameter. Every vendor-scoped query filters on it.
        vendorId: token.vendorId as string,
      };
      return session;
    },
  },
};

export const authOptionsByScope: Record<SessionScope, NextAuthOptions> = {
  admin: adminAuthOptions,
  vendor: vendorAuthOptions,
};
