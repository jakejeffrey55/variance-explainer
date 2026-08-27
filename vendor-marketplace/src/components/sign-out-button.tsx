"use client";

import { useRouter } from "next/navigation";
import { LogOut } from "lucide-react";
import { Button } from "@/components/ui/button";
import { signOutOfSurface, type AuthSurface } from "@/lib/auth/client";

export function SignOutButton({ surface }: { surface: AuthSurface }) {
  const router = useRouter();
  return (
    <Button
      variant="outline"
      size="sm"
      onClick={async () => {
        await signOutOfSurface(surface);
        router.replace(`/${surface}/login`);
        router.refresh();
      }}
    >
      <LogOut /> Sign out
    </Button>
  );
}
