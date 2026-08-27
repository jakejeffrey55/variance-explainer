import type { Prisma } from "@prisma/client";
import { AuthError } from "@/lib/auth/session";
import { prisma } from "@/lib/db";

/**
 * Query-level data isolation.
 *
 * Every vendor-facing read builds its `where` clause here, so the vendor id
 * from the session is structurally impossible to omit. A vendor can never see
 * another vendor's bid amount, chat, rating, or document — not because the UI
 * hides it, but because the query cannot match those rows.
 */

/** A vendor's own bids only. */
export function vendorBidWhere(vendorId: string, extra?: Prisma.BidWhereInput): Prisma.BidWhereInput {
  return { ...extra, vendorId };
}

/** Fields a vendor is allowed to read about a bid — always its own. */
export const vendorBidSelect = {
  id: true,
  jobId: true,
  amount: true,
  laborCost: true,
  materialCost: true,
  status: true,
  notes: true,
  estimatedStartDate: true,
  estimatedCompletionDate: true,
  submittedAt: true,
  withdrawnAt: true,
  approvedAt: true,
  rejectedAt: true,
  rejectionReason: true,
} satisfies Prisma.BidSelect;

/**
 * Job fields a vendor may read. Deliberately omits every competitive signal:
 * no bids relation, no awarded bid, no internal notes, no other vendors.
 */
export const vendorJobSelect = {
  id: true,
  jobNumber: true,
  title: true,
  description: true,
  category: true,
  status: true,
  priority: true,
  emergencyCategory: true,
  budgetMin: true,
  budgetMax: true,
  enforceBudgetCap: true,
  bidDeadline: true,
  scheduledStart: true,
  dueDate: true,
  responseDeadlineMinutes: true,
  dispatchedAt: true,
  claimedByVendorId: true,
  claimedAt: true,
  onSiteAt: true,
  awardedVendorId: true,
  unitNumber: true,
  createdAt: true,
  property: {
    select: {
      id: true,
      name: true,
      addressLine1: true,
      addressLine2: true,
      city: true,
      state: true,
      postalCode: true,
      latitude: true,
      longitude: true,
    },
  },
} satisfies Prisma.JobSelect;

/** Loads a bid, asserting it belongs to the calling vendor. */
export async function getOwnBidOrThrow(vendorId: string, bidId: string) {
  const bid = await prisma.bid.findFirst({
    where: { id: bidId, vendorId },
    select: vendorBidSelect,
  });
  // 404-as-403: never reveal whether a foreign bid id exists.
  if (!bid) throw new AuthError(403, "bid_not_accessible", "Bid not found for this account.");
  return bid;
}

/** Loads a chat thread, asserting it belongs to the calling vendor. */
export async function getOwnThreadOrThrow(vendorId: string, threadId: string) {
  const thread = await prisma.chatThread.findFirst({ where: { id: threadId, vendorId } });
  if (!thread) throw new AuthError(403, "thread_not_accessible", "Conversation not found.");
  return thread;
}

/**
 * A vendor may read a job only when it is open to them: an open standard job,
 * an emergency they are eligible for, or a job they are already engaged on
 * (bid, invitation, claim, or award). Enforced with a database predicate.
 */
export function vendorVisibleJobWhere(vendorId: string): Prisma.JobWhereInput {
  return {
    OR: [
      { status: { in: ["OPEN", "BIDDING_CLOSED", "AWAITING_APPROVAL"] }, inviteOnly: false },
      { bids: { some: { vendorId } } },
      { invitations: { some: { vendorId } } },
      { claimedByVendorId: vendorId },
      { awardedVendorId: vendorId },
    ],
  };
}
