import type { DefaultSession } from "next-auth";

declare module "next-auth" {
  interface Session {
    user?: {
      id: string;
      email: string;
      name: string;
      scope: "admin" | "vendor";
      adminRole?: string;
      vendorId?: string;
    } & DefaultSession["user"];
  }

  interface User {
    id: string;
    scope: "admin" | "vendor";
    adminRole?: string;
    vendorId?: string;
  }
}

declare module "next-auth/jwt" {
  interface JWT {
    scope?: "admin" | "vendor";
    adminRole?: string;
    vendorId?: string;
  }
}
