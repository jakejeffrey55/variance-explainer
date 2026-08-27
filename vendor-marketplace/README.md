# Vendor Marketplace

Installable PWA connecting property management make-ready / general contracting work to
credentialed vendors. Sits between OneSite (property management), Coupa (procurement) and
Vendorply (vendor credentialing) — none of which we have API access to yet, so every one of
them is behind an adapter interface with a working implementation we control today.

**Stack:** Next.js 14 (App Router, TypeScript) · PostgreSQL + Prisma · Tailwind + shadcn/ui ·
NextAuth · PWA (manifest + service worker).

## Build status

| Phase | Scope | State |
| --- | --- | --- |
| 1 | Scaffold, Prisma schema, separate admin/vendor auth, seed data | **done** |
| 2 | Property/job CRUD, CSV import (Power BI + OneSite stubbed), admin dashboard | **done** |
| 3 | Vendor profile/preferences, onboarding approval gate, filtered job board | pending |
| 4 | Bidding + approval, data isolation, budget cap, withdrawal, ApprovalFlag | pending |
| 5 | Emergency dispatch: SMS/push/email, claim flow, auto-escalation | pending |
| 6 | Calendar views, vendor blackout dates | pending |
| 7 | Auto-contract flow for GC jobs > $5,000 | pending |
| 8 | JobRating, TrustScoreSnapshot, Google ReviewProvider | pending |
| 9 | Chat threads per job | pending |
| 10 | ActivityLog timelines, dashboards, bid comparison, docs, map, CSV export | pending |
| 11 | PWA manifest, service worker, push notifications | pending |
| 12 | Polish: empty/loading states, mobile QA, error handling | pending |

## Getting started

```bash
cp .env.example .env          # fill in DATABASE_URL + NEXTAUTH_SECRET at minimum
npm install
npm run db:migrate            # applies prisma/migrations
npm run db:seed               # destructive: wipes and reseeds dev data
npm run dev
```

Against a running dev server:

```bash
bash scripts/auth-smoke.sh      # 16 checks — cross-scope denial, credential separation, account gates
bash scripts/phase2-smoke.sh    # 45 checks — property CRUD, CSV import, job rules, transitions, pages
```

Do not run `npm run build` while `npm run dev` is running — they share `.next` and the dev
server will start serving a half-written build.

## Authentication model

Two authentication surfaces that share nothing:

| | Admin | Vendor |
| --- | --- | --- |
| Login route | `/admin/login` | `/vendor/login` |
| NextAuth mount | `/api/auth/admin/*` | `/api/auth/vendor/*` |
| Credential table | `admin_users` | `vendor_users` |
| Session cookie | `vm.admin.session-token` | `vm.vendor.session-token` |
| Session scope | `"admin"` | `"vendor"` |

An admin cookie presented to a vendor endpoint is not a downgraded session — it is no session
at all, and vice versa. The two cookies can coexist in one browser without interfering.

Enforcement lives in `src/lib/auth/session.ts` and is applied by the route wrappers in
`src/lib/auth/route.ts`:

- `withAdmin` — admin scope, re-read from the database each request (deactivation is immediate).
- `withVendor` — any signed-in vendor, including one pending approval.
- `withApprovedVendor` — `account_status = ACTIVE`; the gate for seeing any job.
- `withBiddingVendor` — approved **and** compliance not expired/missing; the gate for bidding.
- `withEmergencyVendor` — bidding-eligible **and** `emergency_eligible`; the gate for emergency dispatch.

`src/lib/security/scope.ts` holds the query-level isolation helpers (`vendorBidWhere`,
`vendorJobSelect`, `getOwnBidOrThrow`, `vendorVisibleJobWhere`). Vendor reads build their
`where` clause there so the session's vendor id cannot be left out; a vendor id is never
accepted from a request parameter. Middleware (`src/middleware.ts`) only redirects
unauthenticated navigation — it never grants access.

## Admin surface

| Route | What it does |
| --- | --- |
| `/admin` | Dashboard: attention counters, overdue-emergency banner, recent jobs, activity, adapter status |
| `/admin/jobs` | Filterable job board — emergencies pinned in their own section, single-tap status filters |
| `/admin/jobs/new`, `/admin/jobs/[id]`, `/admin/jobs/[id]/edit` | Create (standard or emergency), detail with bid table + timeline, edit |
| `/admin/properties`, `/admin/properties/[id]` | Property list/search, manual add/edit, detail with jobs and rent roll |
| `/admin/import` | CSV import with dry-run preview, per-row errors, template download, import history |

Job rules enforced in `src/lib/services/job.ts` (not in the form):

- Emergency jobs open and dispatch immediately; standard jobs start as drafts.
- An emergency cannot carry a bid deadline, a budget cap, an invite list, or the
  bidding transitions — it is claimed, not bid on.
- A job cannot be published with a bid deadline in the past.
- The budget cap cannot be switched on, or an enforced maximum lowered, once bids exist.
- A GC job whose contract has not reached pending signature cannot move to in progress.
- Completed and cancelled jobs are immutable.

## Integration adapters

| Interface | Implementation today | Stubbed for later |
| --- | --- | --- |
| `PropertySyncProvider` | `CsvPropertySyncProvider` (admin uploads a Power BI export) | `PowerBIPropertyProvider`, `OneSiteProvider` — same interface, return "not configured" |
| `RequisitionProvider` | mock (no Coupa API access yet) | real Coupa client |
| `VendorSyncProvider` | CSV (Vendorply export) | Vendorply API |
| `ContractProvider` | mock | real contract system |
| `ReviewProvider` | Google Places (real) | — |
| `NotificationProvider` | real: Twilio SMS (primary), web push, email | — |

Nothing in the schema names a source system: integration rows carry `provider_key` +
`external_id`, so whichever API access lands first — Power BI or OneSite — is a provider swap,
not a migration.

## Data model

See `prisma/schema.prisma`. Notable invariants that live in application code rather than
constraints:

- At most one non-`WITHDRAWN` bid per (job, vendor). Withdrawn bids stay in history, which is
  why this is not a DB unique constraint.
- `ApprovalSettings` is a singleton row (`id = "default"`) holding the admin-configurable
  approval threshold (7.5%), contract threshold ($5,000), emergency response window (15 min)
  and trust-score weights.
- Emergency jobs are excluded from average-bid comparison entirely — they never carry bids.
