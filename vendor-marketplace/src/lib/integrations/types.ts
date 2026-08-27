import type {
  ContractStatus,
  EmergencyCategory,
  PropertySource,
  RequisitionStatus,
  ServiceCategory,
} from "@prisma/client";

/**
 * Integration contracts.
 *
 * Every external system is reached through one of these interfaces. Application
 * code imports the *interface*, never a concrete client, so swapping a mock or
 * CSV implementation for a live API is a registry change — no schema, UI, or
 * downstream code changes.
 *
 * Each provider reports its own configuration state. A provider that is not
 * wired up yet (Power BI, OneSite) implements the full interface and returns a
 * `not_configured` result, so callers handle it as data rather than as a crash.
 */

export type ProviderIssue = {
  /** 1-based row number for file-based providers; omitted for API providers. */
  row?: number;
  field?: string;
  message: string;
  /** The offending value, echoed back for the admin to correct. */
  value?: string;
};

export type ProviderErrorCode =
  | "not_configured"
  | "auth_failed"
  | "unreachable"
  | "invalid_input"
  | "rate_limited"
  | "provider_error";

export type ProviderError = {
  code: ProviderErrorCode;
  message: string;
  /** What an admin should do about it — surfaced directly in the UI. */
  remedy?: string;
  issues?: ProviderIssue[];
};

export type ProviderResult<T> =
  | { ok: true; data: T; warnings: ProviderIssue[] }
  | { ok: false; error: ProviderError };

export type ProviderStatus = {
  key: string;
  label: string;
  description: string;
  configured: boolean;
  /** Present when `configured` is false: what is missing. */
  reason?: string;
  /** Env vars that must be set before this provider can be activated. */
  requires?: string[];
  capabilities: {
    /** Admin uploads a file (CSV). */
    upload: boolean;
    /** Provider can be polled for data on demand. */
    pull: boolean;
    /** Provider supports fetching only rows changed since a timestamp. */
    incremental: boolean;
  };
};

// ---------------------------------------------------------------------------
// PropertySyncProvider — CSV active today; Power BI and OneSite stubbed.
// ---------------------------------------------------------------------------

/**
 * The canonical property shape. All three implementations produce exactly this,
 * which is what makes the swap free: the import service and everything
 * downstream only ever sees `NormalizedProperty`.
 */
export type NormalizedProperty = {
  externalId: string;
  name: string;
  addressLine1: string;
  addressLine2?: string | null;
  city: string;
  state: string;
  postalCode: string;
  latitude: number;
  longitude: number;
  unitCount?: number | null;
  propertyManagerName?: string | null;
  propertyManagerEmail?: string | null;
  propertyManagerPhone?: string | null;
  isActive: boolean;
};

export type PropertySyncInput = {
  /** Raw file contents, for upload-based providers. */
  fileContent?: string;
  fileName?: string;
  /** Incremental cutoff, for pull-based providers. */
  since?: Date;
};

export interface PropertySyncProvider {
  readonly key: string;
  readonly source: PropertySource;
  status(): ProviderStatus;
  fetchProperties(input: PropertySyncInput): Promise<ProviderResult<NormalizedProperty[]>>;
}

// ---------------------------------------------------------------------------
// VendorSyncProvider — Vendorply credentialing (CSV today).
// ---------------------------------------------------------------------------

export type NormalizedVendorCredential = {
  externalId: string;
  companyName?: string | null;
  email: string;
  phone?: string | null;
  licenseNumber?: string | null;
  w9OnFile?: boolean | null;
  complianceExpiresAt?: Date | null;
  insuranceExpiresAt?: Date | null;
  /** Provider's own verdict, if it publishes one. */
  compliant?: boolean | null;
};

export interface VendorSyncProvider {
  readonly key: string;
  status(): ProviderStatus;
  fetchCredentials(input: PropertySyncInput): Promise<ProviderResult<NormalizedVendorCredential[]>>;
}

// ---------------------------------------------------------------------------
// RequisitionProvider — Coupa (mock until API access lands).
// ---------------------------------------------------------------------------

export type RequisitionRequest = {
  jobId: string;
  jobNumber: string;
  bidId?: string;
  vendorId: string;
  vendorName: string;
  vendorExternalId?: string | null;
  propertyExternalId?: string | null;
  propertyName: string;
  description: string;
  amount: number;
  category: ServiceCategory;
  emergency: boolean;
  emergencyCategory?: EmergencyCategory | null;
  needByDate?: Date | null;
};

export type RequisitionResponse = {
  externalId: string;
  status: RequisitionStatus;
  raw: unknown;
};

export interface RequisitionProvider {
  readonly key: string;
  status(): ProviderStatus;
  createRequisition(req: RequisitionRequest): Promise<ProviderResult<RequisitionResponse>>;
  getRequisition(externalId: string): Promise<ProviderResult<RequisitionResponse>>;
}

// ---------------------------------------------------------------------------
// ContractProvider — GC jobs above the contract threshold.
// ---------------------------------------------------------------------------

export type ContractRequest = {
  jobId: string;
  jobNumber: string;
  bidId: string;
  vendorId: string;
  vendorName: string;
  vendorEmail: string;
  propertyName: string;
  scopeOfWork: string;
  amount: number;
};

export type ContractResponse = {
  externalId: string;
  status: ContractStatus;
  documentUrl?: string | null;
  expiresAt?: Date | null;
  raw: unknown;
};

export interface ContractProvider {
  readonly key: string;
  status(): ProviderStatus;
  createContract(req: ContractRequest): Promise<ProviderResult<ContractResponse>>;
  getContract(externalId: string): Promise<ProviderResult<ContractResponse>>;
}

// ---------------------------------------------------------------------------
// ReviewProvider — Google Places (real).
// ---------------------------------------------------------------------------

export type ExternalReviewSummary = {
  placeId: string;
  rating: number | null;
  reviewCount: number | null;
  fetchedAt: Date;
};

export interface ReviewProvider {
  readonly key: string;
  status(): ProviderStatus;
  findPlaceId(query: { companyName: string; address?: string | null }): Promise<ProviderResult<string | null>>;
  getRating(placeId: string): Promise<ProviderResult<ExternalReviewSummary>>;
}

// ---------------------------------------------------------------------------
// NotificationProvider — SMS (primary for emergencies), push, email.
// ---------------------------------------------------------------------------

export type NotificationRecipient = {
  vendorId?: string;
  adminUserId?: string;
  phone?: string | null;
  email?: string | null;
  pushSubscriptions?: { endpoint: string; p256dh: string; auth: string }[];
};

export type NotificationMessage = {
  template: string;
  jobId?: string;
  subject?: string;
  body: string;
  url?: string;
  /** Emergency dispatch bypasses quiet hours and always tries SMS first. */
  urgent?: boolean;
};

export type NotificationSendResult = {
  channel: "SMS" | "PUSH" | "EMAIL";
  destination: string;
  delivered: boolean;
  providerMessageId?: string | null;
  error?: string;
};

export interface NotificationProvider {
  readonly key: string;
  status(): ProviderStatus;
  send(
    recipient: NotificationRecipient,
    message: NotificationMessage,
    channels: ("SMS" | "PUSH" | "EMAIL")[],
  ): Promise<NotificationSendResult[]>;
}
