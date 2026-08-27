import { clsx, type ClassValue } from "clsx";
import { twMerge } from "tailwind-merge";

export function cn(...inputs: ClassValue[]) {
  return twMerge(clsx(inputs));
}

export function formatCurrency(value: number | string | { toString(): string } | null | undefined) {
  if (value === null || value === undefined) return "—";
  const n = typeof value === "number" ? value : Number(value.toString());
  if (Number.isNaN(n)) return "—";
  return n.toLocaleString("en-US", {
    style: "currency",
    currency: "USD",
    maximumFractionDigits: n % 1 === 0 ? 0 : 2,
  });
}

export function formatDate(value: Date | string | null | undefined, withTime = false) {
  if (!value) return "—";
  const d = typeof value === "string" ? new Date(value) : value;
  if (Number.isNaN(d.getTime())) return "—";
  return d.toLocaleString("en-US", {
    month: "short",
    day: "numeric",
    year: "numeric",
    ...(withTime ? { hour: "numeric", minute: "2-digit" } : {}),
  });
}

/** "in 3 days" / "6 min ago" — used for deadlines and dispatch timers. */
export function formatRelative(value: Date | string | null | undefined) {
  if (!value) return "—";
  const d = typeof value === "string" ? new Date(value) : value;
  const diffMs = d.getTime() - Date.now();
  const abs = Math.abs(diffMs);
  const units: [Intl.RelativeTimeFormatUnit, number][] = [
    ["day", 86400000],
    ["hour", 3600000],
    ["minute", 60000],
  ];
  const rtf = new Intl.RelativeTimeFormat("en-US", { numeric: "auto" });
  for (const [unit, ms] of units) {
    if (abs >= ms || unit === "minute") {
      return rtf.format(Math.round(diffMs / ms), unit);
    }
  }
  return "just now";
}

export function titleCase(value: string) {
  return value
    .toLowerCase()
    .split("_")
    .map((w) => w.charAt(0).toUpperCase() + w.slice(1))
    .join(" ");
}
