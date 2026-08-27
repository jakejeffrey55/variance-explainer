"use client";

import * as React from "react";
import { useRouter } from "next/navigation";
import { AlertCircle, Loader2 } from "lucide-react";
import type { JobPriority, JobStatus } from "@prisma/client";
import { Alert, AlertDescription } from "@/components/ui/alert";
import { Button } from "@/components/ui/button";
import {
  Dialog,
  DialogContent,
  DialogDescription,
  DialogFooter,
  DialogHeader,
  DialogTitle,
} from "@/components/ui/dialog";

type Action = "publish" | "close_bidding" | "reopen" | "cancel" | "start" | "complete";

const ACTIONS: Record<JobStatus, Action[]> = {
  DRAFT: ["publish", "cancel"],
  OPEN: ["close_bidding", "cancel"],
  BIDDING_CLOSED: ["reopen", "cancel"],
  AWAITING_APPROVAL: ["cancel"],
  AWARDED: ["start", "cancel"],
  IN_PROGRESS: ["complete", "cancel"],
  COMPLETED: [],
  CANCELLED: [],
};

const COPY: Record<Action, { label: string; title: string; description: string; destructive?: boolean }> = {
  publish: {
    label: "Publish",
    title: "Publish this job?",
    description: "Matched vendors will see it on their job board and can start bidding immediately.",
  },
  close_bidding: {
    label: "Close bidding",
    title: "Close bidding now?",
    description: "No further bids will be accepted. Bids already submitted stay available for review.",
  },
  reopen: {
    label: "Reopen bidding",
    title: "Reopen bidding?",
    description: "The job returns to the board. Make sure the bid deadline is still in the future.",
  },
  start: {
    label: "Mark in progress",
    title: "Mark this job in progress?",
    description: "Use this once the awarded vendor is mobilised on site.",
  },
  complete: {
    label: "Mark complete",
    title: "Mark this job complete?",
    description: "Completed jobs can no longer be edited, and become ready for rating.",
  },
  cancel: {
    label: "Cancel job",
    title: "Cancel this job?",
    description: "The job closes to vendors and cannot be reopened. This cannot be undone.",
    destructive: true,
  },
};

export function JobStatusActions({
  jobId,
  status,
  priority = "STANDARD",
}: {
  jobId: string;
  status: JobStatus;
  priority?: JobPriority;
}) {
  const router = useRouter();
  const [openAction, setOpenAction] = React.useState<Action | null>(null);
  const [pending, setPending] = React.useState(false);
  const [error, setError] = React.useState<string | null>(null);

  // Emergencies skip bidding entirely, so those transitions never apply to them.
  const available = ACTIONS[status].filter(
    (action) => priority !== "EMERGENCY" || (action !== "close_bidding" && action !== "reopen"),
  );
  if (available.length === 0) return null;

  async function run(action: Action) {
    setPending(true);
    setError(null);
    const res = await fetch(`/api/admin/jobs/${jobId}/status`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ action }),
    });
    if (!res.ok) {
      const data = await res.json().catch(() => null);
      setError(data?.error?.message ?? "Could not update the job.");
      setPending(false);
      return;
    }
    setPending(false);
    setOpenAction(null);
    router.refresh();
  }

  return (
    <>
      {available.map((action) => (
        <Button
          key={action}
          variant={COPY[action].destructive ? "outline" : action === "publish" ? "default" : "secondary"}
          size="sm"
          onClick={() => {
            setError(null);
            setOpenAction(action);
          }}
          className={COPY[action].destructive ? "text-destructive" : undefined}
        >
          {COPY[action].label}
        </Button>
      ))}

      <Dialog open={openAction !== null} onOpenChange={(open) => !open && setOpenAction(null)}>
        <DialogContent className="max-w-md">
          {openAction && (
            <>
              <DialogHeader>
                <DialogTitle>{COPY[openAction].title}</DialogTitle>
                <DialogDescription>{COPY[openAction].description}</DialogDescription>
              </DialogHeader>

              {error && (
                <Alert variant="destructive">
                  <AlertCircle />
                  <AlertDescription>{error}</AlertDescription>
                </Alert>
              )}

              <DialogFooter>
                <Button variant="outline" onClick={() => setOpenAction(null)} disabled={pending}>
                  Never mind
                </Button>
                <Button
                  variant={COPY[openAction].destructive ? "destructive" : "default"}
                  onClick={() => run(openAction)}
                  disabled={pending}
                >
                  {pending && <Loader2 className="animate-spin" />}
                  {COPY[openAction].label}
                </Button>
              </DialogFooter>
            </>
          )}
        </DialogContent>
      </Dialog>
    </>
  );
}
