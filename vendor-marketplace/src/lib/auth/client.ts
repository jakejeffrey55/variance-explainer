"use client";

/**
 * Sign-in/out against a specific auth surface.
 *
 * The two surfaces are separate NextAuth mounts with separate cookies, so the
 * surface is part of the URL rather than a parameter on a shared endpoint.
 * (next-auth/react's helpers assume a single global mount, which is exactly the
 * shared-session model this app deliberately does not use.)
 */
export type AuthSurface = "admin" | "vendor";

async function csrfToken(surface: AuthSurface) {
  const res = await fetch(`/api/auth/${surface}/csrf`, { credentials: "same-origin" });
  const data = (await res.json()) as { csrfToken: string };
  return data.csrfToken;
}

export async function signInWithCredentials(
  surface: AuthSurface,
  email: string,
  password: string,
): Promise<{ ok: true } | { ok: false; message: string }> {
  const token = await csrfToken(surface);
  const body = new URLSearchParams({
    csrfToken: token,
    email,
    password,
    json: "true",
    callbackUrl: window.location.origin,
  });

  const res = await fetch(`/api/auth/${surface}/callback/${surface}-credentials`, {
    method: "POST",
    headers: { "Content-Type": "application/x-www-form-urlencoded" },
    body,
    credentials: "same-origin",
    redirect: "follow",
  });

  const data = (await res.json().catch(() => ({}))) as { url?: string };
  // NextAuth signals a failed credential check by redirecting back to the
  // sign-in page with ?error=, rather than by status code.
  if (!res.ok || (data.url && data.url.includes("error"))) {
    return {
      ok: false,
      message:
        surface === "admin"
          ? "Those credentials do not match an active admin account."
          : "Those credentials do not match an active vendor account.",
    };
  }
  return { ok: true };
}

export async function signOutOfSurface(surface: AuthSurface) {
  const token = await csrfToken(surface);
  await fetch(`/api/auth/${surface}/signout`, {
    method: "POST",
    headers: { "Content-Type": "application/x-www-form-urlencoded" },
    body: new URLSearchParams({ csrfToken: token, json: "true", callbackUrl: window.location.origin }),
    credentials: "same-origin",
  });
}
