import type {
  AccountStatus,
  BidStatus,
  ComplianceStatus,
  ContractStatus,
  JobStatus,
  ServiceCategory,
} from "@prisma/client";
import { AlertTriangle, Siren } from "lucide-react";
import { Badge, type BadgeProps } from "@/components/ui/badge";
import { cn, titleCase } from "@/lib/utils";

const JOB_STATUS_VARIANT: Record<JobStatus, BadgeProps["variant"]> = {
  DRAFT: "muted",
  OPEN: "default",
  BIDDING_CLOSED: "secondary",
  AWAITING_APPROVAL: "warning",
  AWARDED: "success",
  IN_PROGRESS: "default",
  COMPLETED: "success",
  CANCELLED: "muted",
};

export function JobStatusBadge({ status }: { status: JobStatus }) {
  return <Badge variant={JOB_STATUS_VARIANT[status]}>{titleCase(status)}</Badge>;
}

const BID_STATUS_VARIANT: Record<BidStatus, BadgeProps["variant"]> = {
  SUBMITTED: "default",
  WITHDRAWN: "muted",
  APPROVED: "success",
  REJECTED: "destructive",
  EXPIRED: "muted",
};

export function BidStatusBadge({ status }: { status: BidStatus }) {
  return <Badge variant={BID_STATUS_VARIANT[status]}>{titleCase(status)}</Badge>;
}

const ACCOUNT_STATUS_VARIANT: Record<AccountStatus, BadgeProps["variant"]> = {
  PENDING: "warning",
  ACTIVE: "success",
  SUSPENDED: "destructive",
  REJECTED: "destructive",
};

export function AccountStatusBadge({ status }: { status: AccountStatus }) {
  return <Badge variant={ACCOUNT_STATUS_VARIANT[status]}>{titleCase(status)}</Badge>;
}

const COMPLIANCE_VARIANT: Record<ComplianceStatus, BadgeProps["variant"]> = {
  COMPLIANT: "success",
  EXPIRING_SOON: "warning",
  EXPIRED: "destructive",
  NOT_SUBMITTED: "muted",
};

export function ComplianceBadge({ status }: { status: ComplianceStatus }) {
  return (
    <Badge variant={COMPLIANCE_VARIANT[status]}>
      {status === "EXPIRED" && <AlertTriangle className="h-3 w-3" />}
      {titleCase(status)}
    </Badge>
  );
}

const CONTRACT_VARIANT: Record<ContractStatus, BadgeProps["variant"]> = {
  NOT_REQUIRED: "muted",
  DRAFT: "secondary",
  PENDING_SIGNATURE: "warning",
  SIGNED: "success",
  VOID: "destructive",
};

export function ContractStatusBadge({ status }: { status: ContractStatus }) {
  return <Badge variant={CONTRACT_VARIANT[status]}>{titleCase(status)}</Badge>;
}

export function EmergencyBadge({ className }: { className?: string }) {
  return (
    <Badge variant="emergency" className={cn("uppercase tracking-wide", className)}>
      <Siren className="h-3 w-3" />
      Emergency
    </Badge>
  );
}

export function CategoryBadge({ category }: { category: ServiceCategory }) {
  return <Badge variant="secondary">{titleCase(category)}</Badge>;
}
