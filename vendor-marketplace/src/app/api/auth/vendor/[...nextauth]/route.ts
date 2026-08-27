import NextAuth from "next-auth";
import { vendorAuthOptions } from "@/lib/auth/options";

const handler = NextAuth(vendorAuthOptions);

export { handler as GET, handler as POST };
