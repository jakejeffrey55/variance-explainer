import { formatDate } from "@/lib/utils";

export type TimelineEntry = {
  id: string;
  action: string;
  summary: string | null;
  fromStatus: string | null;
  toStatus: string | null;
  actorLabel: string;
  createdAt: Date;
};

/** Renders the ActivityLog for one job — the record of everything that happened. */
export function ActivityTimeline({ entries }: { entries: TimelineEntry[] }) {
  if (entries.length === 0) {
    return <p className="text-sm text-muted-foreground">Nothing has happened on this job yet.</p>;
  }

  return (
    <ol className="relative space-y-5 border-l pl-5">
      {entries.map((entry) => (
        <li key={entry.id} className="relative">
          <span
            className="absolute -left-[1.4rem] top-1.5 h-2 w-2 rounded-full bg-primary ring-4 ring-card"
            aria-hidden
          />
          <p className="text-sm leading-snug">{entry.summary ?? entry.action}</p>
          <p className="mt-0.5 text-xs text-muted-foreground">
            {entry.actorLabel} · {formatDate(entry.createdAt, true)}
            {entry.fromStatus && entry.toStatus ? ` · ${entry.fromStatus} → ${entry.toStatus}` : ""}
          </p>
        </li>
      ))}
    </ol>
  );
}
