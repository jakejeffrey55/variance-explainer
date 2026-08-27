"use client";

import * as React from "react";
import { useRouter, useSearchParams } from "next/navigation";
import { AlertCircle, Loader2 } from "lucide-react";
import { Alert, AlertDescription } from "@/components/ui/alert";
import { Button } from "@/components/ui/button";
import { Input } from "@/components/ui/input";
import { Label } from "@/components/ui/label";
import { signInWithCredentials, type AuthSurface } from "@/lib/auth/client";

export function LoginForm({ surface, defaultPath }: { surface: AuthSurface; defaultPath: string }) {
  const router = useRouter();
  const searchParams = useSearchParams();
  const [email, setEmail] = React.useState("");
  const [password, setPassword] = React.useState("");
  const [error, setError] = React.useState<string | null>(null);
  const [pending, setPending] = React.useState(false);

  async function onSubmit(event: React.FormEvent) {
    event.preventDefault();
    setPending(true);
    setError(null);

    const result = await signInWithCredentials(surface, email, password);
    if (!result.ok) {
      setError(result.message);
      setPending(false);
      return;
    }

    const callbackUrl = searchParams.get("callbackUrl");
    const target = callbackUrl?.startsWith(`/${surface}`) ? callbackUrl : defaultPath;
    router.replace(target);
    router.refresh();
  }

  return (
    <form onSubmit={onSubmit} className="space-y-4">
      {error && (
        <Alert variant="destructive">
          <AlertCircle />
          <AlertDescription>{error}</AlertDescription>
        </Alert>
      )}

      <div className="space-y-2">
        <Label htmlFor="email">Email</Label>
        <Input
          id="email"
          type="email"
          autoComplete="username"
          required
          value={email}
          onChange={(e) => setEmail(e.target.value)}
          placeholder={surface === "admin" ? "you@company.com" : "you@yourcompany.com"}
        />
      </div>

      <div className="space-y-2">
        <Label htmlFor="password">Password</Label>
        <Input
          id="password"
          type="password"
          autoComplete="current-password"
          required
          value={password}
          onChange={(e) => setPassword(e.target.value)}
        />
      </div>

      <Button type="submit" className="w-full" disabled={pending}>
        {pending && <Loader2 className="animate-spin" />}
        {pending ? "Signing in…" : "Sign in"}
      </Button>
    </form>
  );
}
