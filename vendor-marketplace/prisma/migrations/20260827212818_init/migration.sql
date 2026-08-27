-- CreateEnum
CREATE TYPE "AdminRole" AS ENUM ('OWNER', 'MANAGER', 'STAFF');

-- CreateEnum
CREATE TYPE "AccountStatus" AS ENUM ('PENDING', 'ACTIVE', 'SUSPENDED', 'REJECTED');

-- CreateEnum
CREATE TYPE "ComplianceStatus" AS ENUM ('COMPLIANT', 'EXPIRING_SOON', 'EXPIRED', 'NOT_SUBMITTED');

-- CreateEnum
CREATE TYPE "ServiceCategory" AS ENUM ('MAKE_READY', 'GENERAL_CONTRACTING', 'PLUMBING', 'ELECTRICAL', 'HVAC', 'FLOORING', 'PAINTING', 'DRYWALL', 'APPLIANCE', 'CLEANING', 'LANDSCAPING', 'ROOFING', 'PEST_CONTROL', 'WATER_MITIGATION', 'OTHER');

-- CreateEnum
CREATE TYPE "JobStatus" AS ENUM ('DRAFT', 'OPEN', 'BIDDING_CLOSED', 'AWAITING_APPROVAL', 'AWARDED', 'IN_PROGRESS', 'COMPLETED', 'CANCELLED');

-- CreateEnum
CREATE TYPE "JobPriority" AS ENUM ('STANDARD', 'EMERGENCY');

-- CreateEnum
CREATE TYPE "EmergencyCategory" AS ENUM ('AC_HVAC', 'LEAK', 'WATER_EXTRACTION', 'ELECTRICAL', 'OTHER');

-- CreateEnum
CREATE TYPE "BidStatus" AS ENUM ('SUBMITTED', 'WITHDRAWN', 'APPROVED', 'REJECTED', 'EXPIRED');

-- CreateEnum
CREATE TYPE "RequisitionStatus" AS ENUM ('PENDING', 'SUBMITTED', 'ACKNOWLEDGED', 'FAILED');

-- CreateEnum
CREATE TYPE "ContractStatus" AS ENUM ('NOT_REQUIRED', 'DRAFT', 'PENDING_SIGNATURE', 'SIGNED', 'VOID');

-- CreateEnum
CREATE TYPE "ApprovalFlagType" AS ENUM ('ABOVE_AVERAGE_THRESHOLD', 'ABOVE_BUDGET_CAP');

-- CreateEnum
CREATE TYPE "PropertySource" AS ENUM ('CSV', 'POWER_BI', 'ONESITE', 'MANUAL');

-- CreateEnum
CREATE TYPE "ImportKind" AS ENUM ('PROPERTY', 'VENDOR', 'RENT_ROLL');

-- CreateEnum
CREATE TYPE "ImportStatus" AS ENUM ('PENDING', 'PROCESSING', 'COMPLETED', 'FAILED');

-- CreateEnum
CREATE TYPE "AvailabilityType" AS ENUM ('BLACKOUT', 'REDUCED_CAPACITY');

-- CreateEnum
CREATE TYPE "ActorType" AS ENUM ('ADMIN', 'VENDOR', 'SYSTEM');

-- CreateEnum
CREATE TYPE "ActivityEntityType" AS ENUM ('JOB', 'BID', 'CONTRACT', 'REQUISITION', 'VENDOR', 'PROPERTY', 'RATING', 'APPROVAL_FLAG');

-- CreateEnum
CREATE TYPE "NotificationChannel" AS ENUM ('SMS', 'PUSH', 'EMAIL');

-- CreateEnum
CREATE TYPE "NotificationStatus" AS ENUM ('QUEUED', 'SENT', 'DELIVERED', 'FAILED');

-- CreateEnum
CREATE TYPE "DocumentType" AS ENUM ('COI', 'W9', 'LICENSE', 'CONTRACT', 'PHOTO_BEFORE', 'PHOTO_AFTER', 'INVOICE', 'SCOPE_OF_WORK', 'OTHER');

-- CreateTable
CREATE TABLE "admin_users" (
    "id" TEXT NOT NULL,
    "email" TEXT NOT NULL,
    "password_hash" TEXT NOT NULL,
    "name" TEXT NOT NULL,
    "phone" TEXT,
    "role" "AdminRole" NOT NULL DEFAULT 'STAFF',
    "is_active" BOOLEAN NOT NULL DEFAULT true,
    "last_login_at" TIMESTAMP(3),
    "created_at" TIMESTAMP(3) NOT NULL DEFAULT CURRENT_TIMESTAMP,
    "updated_at" TIMESTAMP(3) NOT NULL,

    CONSTRAINT "admin_users_pkey" PRIMARY KEY ("id")
);

-- CreateTable
CREATE TABLE "vendor_users" (
    "id" TEXT NOT NULL,
    "vendor_id" TEXT NOT NULL,
    "email" TEXT NOT NULL,
    "password_hash" TEXT NOT NULL,
    "name" TEXT NOT NULL,
    "phone" TEXT,
    "is_primary" BOOLEAN NOT NULL DEFAULT false,
    "is_active" BOOLEAN NOT NULL DEFAULT true,
    "last_login_at" TIMESTAMP(3),
    "created_at" TIMESTAMP(3) NOT NULL DEFAULT CURRENT_TIMESTAMP,
    "updated_at" TIMESTAMP(3) NOT NULL,

    CONSTRAINT "vendor_users_pkey" PRIMARY KEY ("id")
);

-- CreateTable
CREATE TABLE "properties" (
    "id" TEXT NOT NULL,
    "external_id" TEXT,
    "source" "PropertySource" NOT NULL DEFAULT 'CSV',
    "source_batch_id" TEXT,
    "name" TEXT NOT NULL,
    "address_line1" TEXT NOT NULL,
    "address_line2" TEXT,
    "city" TEXT NOT NULL,
    "state" TEXT NOT NULL,
    "postal_code" TEXT NOT NULL,
    "latitude" DOUBLE PRECISION NOT NULL,
    "longitude" DOUBLE PRECISION NOT NULL,
    "unit_count" INTEGER,
    "property_manager_name" TEXT,
    "property_manager_email" TEXT,
    "property_manager_phone" TEXT,
    "is_active" BOOLEAN NOT NULL DEFAULT true,
    "last_synced_at" TIMESTAMP(3),
    "created_at" TIMESTAMP(3) NOT NULL DEFAULT CURRENT_TIMESTAMP,
    "updated_at" TIMESTAMP(3) NOT NULL,

    CONSTRAINT "properties_pkey" PRIMARY KEY ("id")
);

-- CreateTable
CREATE TABLE "rent_rolls" (
    "id" TEXT NOT NULL,
    "property_id" TEXT NOT NULL,
    "import_batch_id" TEXT,
    "unit_number" TEXT NOT NULL,
    "unit_type" TEXT,
    "square_feet" INTEGER,
    "market_rent" DECIMAL(12,2),
    "current_rent" DECIMAL(12,2),
    "status" TEXT,
    "lease_end_date" TIMESTAMP(3),
    "move_out_date" TIMESTAMP(3),
    "move_in_date" TIMESTAMP(3),
    "imported_at" TIMESTAMP(3) NOT NULL DEFAULT CURRENT_TIMESTAMP,

    CONSTRAINT "rent_rolls_pkey" PRIMARY KEY ("id")
);

-- CreateTable
CREATE TABLE "import_batches" (
    "id" TEXT NOT NULL,
    "kind" "ImportKind" NOT NULL,
    "provider_key" TEXT NOT NULL,
    "status" "ImportStatus" NOT NULL DEFAULT 'PENDING',
    "file_name" TEXT,
    "rows_total" INTEGER NOT NULL DEFAULT 0,
    "rows_imported" INTEGER NOT NULL DEFAULT 0,
    "rows_skipped" INTEGER NOT NULL DEFAULT 0,
    "rows_failed" INTEGER NOT NULL DEFAULT 0,
    "errors" JSONB,
    "admin_user_id" TEXT,
    "started_at" TIMESTAMP(3) NOT NULL DEFAULT CURRENT_TIMESTAMP,
    "finished_at" TIMESTAMP(3),

    CONSTRAINT "import_batches_pkey" PRIMARY KEY ("id")
);

-- CreateTable
CREATE TABLE "vendors" (
    "id" TEXT NOT NULL,
    "company_name" TEXT NOT NULL,
    "contact_name" TEXT NOT NULL,
    "email" TEXT NOT NULL,
    "phone" TEXT NOT NULL,
    "sms_opt_in" BOOLEAN NOT NULL DEFAULT true,
    "address_line1" TEXT,
    "city" TEXT,
    "state" TEXT,
    "postal_code" TEXT,
    "latitude" DOUBLE PRECISION NOT NULL,
    "longitude" DOUBLE PRECISION NOT NULL,
    "service_radius_miles" INTEGER NOT NULL DEFAULT 25,
    "service_categories" "ServiceCategory"[],
    "account_status" "AccountStatus" NOT NULL DEFAULT 'PENDING',
    "emergency_eligible" BOOLEAN NOT NULL DEFAULT false,
    "compliance_status" "ComplianceStatus" NOT NULL DEFAULT 'NOT_SUBMITTED',
    "vendorply_id" TEXT,
    "compliance_expires_at" TIMESTAMP(3),
    "insurance_expires_at" TIMESTAMP(3),
    "license_number" TEXT,
    "w9_on_file" BOOLEAN NOT NULL DEFAULT false,
    "last_credential_sync_at" TIMESTAMP(3),
    "google_place_id" TEXT,
    "google_rating" DOUBLE PRECISION,
    "google_rating_count" INTEGER,
    "google_rating_fetched_at" TIMESTAMP(3),
    "years_in_business" INTEGER,
    "trust_score" DOUBLE PRECISION,
    "trust_score_at" TIMESTAMP(3),
    "approved_at" TIMESTAMP(3),
    "approved_by_admin_id" TEXT,
    "rejection_reason" TEXT,
    "internal_notes" TEXT,
    "created_at" TIMESTAMP(3) NOT NULL DEFAULT CURRENT_TIMESTAMP,
    "updated_at" TIMESTAMP(3) NOT NULL,

    CONSTRAINT "vendors_pkey" PRIMARY KEY ("id")
);

-- CreateTable
CREATE TABLE "vendor_availability" (
    "id" TEXT NOT NULL,
    "vendor_id" TEXT NOT NULL,
    "type" "AvailabilityType" NOT NULL DEFAULT 'BLACKOUT',
    "start_date" TIMESTAMP(3) NOT NULL,
    "end_date" TIMESTAMP(3) NOT NULL,
    "reason" TEXT,
    "created_at" TIMESTAMP(3) NOT NULL DEFAULT CURRENT_TIMESTAMP,

    CONSTRAINT "vendor_availability_pkey" PRIMARY KEY ("id")
);

-- CreateTable
CREATE TABLE "jobs" (
    "id" TEXT NOT NULL,
    "job_number" TEXT NOT NULL,
    "property_id" TEXT NOT NULL,
    "unit_number" TEXT,
    "title" TEXT NOT NULL,
    "description" TEXT NOT NULL,
    "category" "ServiceCategory" NOT NULL,
    "status" "JobStatus" NOT NULL DEFAULT 'DRAFT',
    "budget_min" DECIMAL(12,2) NOT NULL,
    "budget_max" DECIMAL(12,2) NOT NULL,
    "enforce_budget_cap" BOOLEAN NOT NULL DEFAULT false,
    "bid_deadline" TIMESTAMP(3),
    "invite_only" BOOLEAN NOT NULL DEFAULT false,
    "scheduled_start" TIMESTAMP(3),
    "due_date" TIMESTAMP(3),
    "priority" "JobPriority" NOT NULL DEFAULT 'STANDARD',
    "emergency_category" "EmergencyCategory",
    "response_deadline_minutes" INTEGER,
    "dispatched_at" TIMESTAMP(3),
    "claimed_by_vendor_id" TEXT,
    "claimed_at" TIMESTAMP(3),
    "on_site_at" TIMESTAMP(3),
    "escalated_at" TIMESTAMP(3),
    "escalation_radius_miles" INTEGER,
    "awarded_vendor_id" TEXT,
    "awarded_bid_id" TEXT,
    "awarded_at" TIMESTAMP(3),
    "started_at" TIMESTAMP(3),
    "completed_at" TIMESTAMP(3),
    "cancelled_at" TIMESTAMP(3),
    "created_by_admin_id" TEXT NOT NULL,
    "created_at" TIMESTAMP(3) NOT NULL DEFAULT CURRENT_TIMESTAMP,
    "updated_at" TIMESTAMP(3) NOT NULL,

    CONSTRAINT "jobs_pkey" PRIMARY KEY ("id")
);

-- CreateTable
CREATE TABLE "job_invitations" (
    "id" TEXT NOT NULL,
    "job_id" TEXT NOT NULL,
    "vendor_id" TEXT NOT NULL,
    "invited_by_admin_id" TEXT NOT NULL,
    "invited_at" TIMESTAMP(3) NOT NULL DEFAULT CURRENT_TIMESTAMP,
    "viewed_at" TIMESTAMP(3),
    "responded_at" TIMESTAMP(3),

    CONSTRAINT "job_invitations_pkey" PRIMARY KEY ("id")
);

-- CreateTable
CREATE TABLE "bids" (
    "id" TEXT NOT NULL,
    "job_id" TEXT NOT NULL,
    "vendor_id" TEXT NOT NULL,
    "amount" DECIMAL(12,2) NOT NULL,
    "labor_cost" DECIMAL(12,2),
    "material_cost" DECIMAL(12,2),
    "status" "BidStatus" NOT NULL DEFAULT 'SUBMITTED',
    "notes" TEXT,
    "estimated_start_date" TIMESTAMP(3),
    "estimated_completion_date" TIMESTAMP(3),
    "above_budget_max" BOOLEAN NOT NULL DEFAULT false,
    "submitted_at" TIMESTAMP(3) NOT NULL DEFAULT CURRENT_TIMESTAMP,
    "withdrawn_at" TIMESTAMP(3),
    "approved_at" TIMESTAMP(3),
    "approved_by_admin_id" TEXT,
    "rejected_at" TIMESTAMP(3),
    "rejection_reason" TEXT,
    "updated_at" TIMESTAMP(3) NOT NULL,

    CONSTRAINT "bids_pkey" PRIMARY KEY ("id")
);

-- CreateTable
CREATE TABLE "requisitions" (
    "id" TEXT NOT NULL,
    "job_id" TEXT NOT NULL,
    "bid_id" TEXT,
    "vendor_id" TEXT NOT NULL,
    "provider_key" TEXT NOT NULL,
    "external_id" TEXT,
    "status" "RequisitionStatus" NOT NULL DEFAULT 'PENDING',
    "amount" DECIMAL(12,2) NOT NULL,
    "request_payload" JSONB,
    "response_body" JSONB,
    "error_message" TEXT,
    "submitted_at" TIMESTAMP(3),
    "acknowledged_at" TIMESTAMP(3),
    "created_at" TIMESTAMP(3) NOT NULL DEFAULT CURRENT_TIMESTAMP,
    "updated_at" TIMESTAMP(3) NOT NULL,

    CONSTRAINT "requisitions_pkey" PRIMARY KEY ("id")
);

-- CreateTable
CREATE TABLE "contracts" (
    "id" TEXT NOT NULL,
    "job_id" TEXT NOT NULL,
    "bid_id" TEXT,
    "vendor_id" TEXT NOT NULL,
    "provider_key" TEXT NOT NULL,
    "external_id" TEXT,
    "status" "ContractStatus" NOT NULL DEFAULT 'DRAFT',
    "amount" DECIMAL(12,2) NOT NULL,
    "document_url" TEXT,
    "payload" JSONB,
    "sent_at" TIMESTAMP(3),
    "signed_at" TIMESTAMP(3),
    "expires_at" TIMESTAMP(3),
    "error_message" TEXT,
    "created_at" TIMESTAMP(3) NOT NULL DEFAULT CURRENT_TIMESTAMP,
    "updated_at" TIMESTAMP(3) NOT NULL,

    CONSTRAINT "contracts_pkey" PRIMARY KEY ("id")
);

-- CreateTable
CREATE TABLE "job_ratings" (
    "id" TEXT NOT NULL,
    "job_id" TEXT NOT NULL,
    "vendor_id" TEXT NOT NULL,
    "admin_user_id" TEXT NOT NULL,
    "score" INTEGER NOT NULL,
    "no_show" BOOLEAN NOT NULL DEFAULT false,
    "missed_deadline" BOOLEAN NOT NULL DEFAULT false,
    "quality_score" INTEGER,
    "communication_score" INTEGER,
    "timeliness_score" INTEGER,
    "comment" TEXT,
    "created_at" TIMESTAMP(3) NOT NULL DEFAULT CURRENT_TIMESTAMP,

    CONSTRAINT "job_ratings_pkey" PRIMARY KEY ("id")
);

-- CreateTable
CREATE TABLE "trust_score_snapshots" (
    "id" TEXT NOT NULL,
    "vendor_id" TEXT NOT NULL,
    "score" DOUBLE PRECISION NOT NULL,
    "avg_job_rating" DOUBLE PRECISION,
    "compliance_uptime_pct" DOUBLE PRECISION,
    "google_rating" DOUBLE PRECISION,
    "experience_score" DOUBLE PRECISION,
    "reliability_penalty" DOUBLE PRECISION NOT NULL DEFAULT 0,
    "emergency_response_score" DOUBLE PRECISION,
    "no_show_count" INTEGER NOT NULL DEFAULT 0,
    "missed_deadline_count" INTEGER NOT NULL DEFAULT 0,
    "completed_job_count" INTEGER NOT NULL DEFAULT 0,
    "avg_emergency_response_mins" DOUBLE PRECISION,
    "weights" JSONB NOT NULL,
    "components" JSONB NOT NULL,
    "reason" TEXT,
    "computed_at" TIMESTAMP(3) NOT NULL DEFAULT CURRENT_TIMESTAMP,

    CONSTRAINT "trust_score_snapshots_pkey" PRIMARY KEY ("id")
);

-- CreateTable
CREATE TABLE "approval_flags" (
    "id" TEXT NOT NULL,
    "job_id" TEXT NOT NULL,
    "bid_id" TEXT NOT NULL,
    "property_id" TEXT NOT NULL,
    "type" "ApprovalFlagType" NOT NULL,
    "threshold_pct" DOUBLE PRECISION,
    "average_bid_amount" DECIMAL(12,2),
    "approved_amount" DECIMAL(12,2) NOT NULL,
    "delta_pct" DOUBLE PRECISION,
    "budget_max" DECIMAL(12,2),
    "bid_count" INTEGER,
    "note" TEXT,
    "acknowledged_by_admin_id" TEXT,
    "acknowledged_at" TIMESTAMP(3),
    "created_at" TIMESTAMP(3) NOT NULL DEFAULT CURRENT_TIMESTAMP,

    CONSTRAINT "approval_flags_pkey" PRIMARY KEY ("id")
);

-- CreateTable
CREATE TABLE "approval_settings" (
    "id" TEXT NOT NULL DEFAULT 'default',
    "above_average_threshold_pct" DOUBLE PRECISION NOT NULL DEFAULT 7.5,
    "contract_threshold_amount" DECIMAL(12,2) NOT NULL DEFAULT 5000,
    "emergency_response_minutes" INTEGER NOT NULL DEFAULT 15,
    "emergency_radius_expansion_miles" INTEGER NOT NULL DEFAULT 25,
    "default_service_radius_miles" INTEGER NOT NULL DEFAULT 25,
    "compliance_expiry_warning_days" INTEGER NOT NULL DEFAULT 30,
    "trust_score_weights" JSONB NOT NULL,
    "updated_by_admin_id" TEXT,
    "updated_at" TIMESTAMP(3) NOT NULL,

    CONSTRAINT "approval_settings_pkey" PRIMARY KEY ("id")
);

-- CreateTable
CREATE TABLE "chat_threads" (
    "id" TEXT NOT NULL,
    "job_id" TEXT NOT NULL,
    "vendor_id" TEXT NOT NULL,
    "admin_unread_count" INTEGER NOT NULL DEFAULT 0,
    "vendor_unread_count" INTEGER NOT NULL DEFAULT 0,
    "last_message_at" TIMESTAMP(3),
    "created_at" TIMESTAMP(3) NOT NULL DEFAULT CURRENT_TIMESTAMP,

    CONSTRAINT "chat_threads_pkey" PRIMARY KEY ("id")
);

-- CreateTable
CREATE TABLE "chat_messages" (
    "id" TEXT NOT NULL,
    "thread_id" TEXT NOT NULL,
    "senderType" "ActorType" NOT NULL,
    "sender_admin_id" TEXT,
    "sender_vendor_user_id" TEXT,
    "body" TEXT NOT NULL,
    "attachment_url" TEXT,
    "read_by_admin_at" TIMESTAMP(3),
    "read_by_vendor_at" TIMESTAMP(3),
    "created_at" TIMESTAMP(3) NOT NULL DEFAULT CURRENT_TIMESTAMP,

    CONSTRAINT "chat_messages_pkey" PRIMARY KEY ("id")
);

-- CreateTable
CREATE TABLE "job_documents" (
    "id" TEXT NOT NULL,
    "job_id" TEXT NOT NULL,
    "vendor_id" TEXT,
    "type" "DocumentType" NOT NULL DEFAULT 'OTHER',
    "file_name" TEXT NOT NULL,
    "url" TEXT NOT NULL,
    "mime_type" TEXT,
    "size_bytes" INTEGER,
    "uploaded_by_type" "ActorType" NOT NULL,
    "uploaded_by_id" TEXT,
    "created_at" TIMESTAMP(3) NOT NULL DEFAULT CURRENT_TIMESTAMP,

    CONSTRAINT "job_documents_pkey" PRIMARY KEY ("id")
);

-- CreateTable
CREATE TABLE "activity_logs" (
    "id" TEXT NOT NULL,
    "entity_type" "ActivityEntityType" NOT NULL,
    "entity_id" TEXT NOT NULL,
    "job_id" TEXT,
    "action" TEXT NOT NULL,
    "from_status" TEXT,
    "to_status" TEXT,
    "actor_type" "ActorType" NOT NULL,
    "actor_admin_id" TEXT,
    "actor_vendor_id" TEXT,
    "summary" TEXT,
    "metadata" JSONB,
    "created_at" TIMESTAMP(3) NOT NULL DEFAULT CURRENT_TIMESTAMP,

    CONSTRAINT "activity_logs_pkey" PRIMARY KEY ("id")
);

-- CreateTable
CREATE TABLE "notification_logs" (
    "id" TEXT NOT NULL,
    "channel" "NotificationChannel" NOT NULL,
    "status" "NotificationStatus" NOT NULL DEFAULT 'QUEUED',
    "recipient_type" "ActorType" NOT NULL,
    "vendor_id" TEXT,
    "admin_user_id" TEXT,
    "job_id" TEXT,
    "template" TEXT NOT NULL,
    "destination" TEXT NOT NULL,
    "body" TEXT,
    "provider_key" TEXT,
    "provider_message_id" TEXT,
    "error_message" TEXT,
    "sent_at" TIMESTAMP(3),
    "created_at" TIMESTAMP(3) NOT NULL DEFAULT CURRENT_TIMESTAMP,

    CONSTRAINT "notification_logs_pkey" PRIMARY KEY ("id")
);

-- CreateTable
CREATE TABLE "push_subscriptions" (
    "id" TEXT NOT NULL,
    "endpoint" TEXT NOT NULL,
    "p256dh" TEXT NOT NULL,
    "auth" TEXT NOT NULL,
    "vendor_user_id" TEXT,
    "admin_user_id" TEXT,
    "user_agent" TEXT,
    "last_used_at" TIMESTAMP(3),
    "created_at" TIMESTAMP(3) NOT NULL DEFAULT CURRENT_TIMESTAMP,

    CONSTRAINT "push_subscriptions_pkey" PRIMARY KEY ("id")
);

-- CreateIndex
CREATE UNIQUE INDEX "admin_users_email_key" ON "admin_users"("email");

-- CreateIndex
CREATE UNIQUE INDEX "vendor_users_email_key" ON "vendor_users"("email");

-- CreateIndex
CREATE INDEX "vendor_users_vendor_id_idx" ON "vendor_users"("vendor_id");

-- CreateIndex
CREATE INDEX "properties_is_active_idx" ON "properties"("is_active");

-- CreateIndex
CREATE UNIQUE INDEX "properties_source_external_id_key" ON "properties"("source", "external_id");

-- CreateIndex
CREATE INDEX "rent_rolls_move_out_date_idx" ON "rent_rolls"("move_out_date");

-- CreateIndex
CREATE UNIQUE INDEX "rent_rolls_property_id_unit_number_key" ON "rent_rolls"("property_id", "unit_number");

-- CreateIndex
CREATE INDEX "import_batches_kind_started_at_idx" ON "import_batches"("kind", "started_at");

-- CreateIndex
CREATE UNIQUE INDEX "vendors_email_key" ON "vendors"("email");

-- CreateIndex
CREATE INDEX "vendors_account_status_idx" ON "vendors"("account_status");

-- CreateIndex
CREATE INDEX "vendors_emergency_eligible_account_status_compliance_status_idx" ON "vendors"("emergency_eligible", "account_status", "compliance_status");

-- CreateIndex
CREATE INDEX "vendor_availability_vendor_id_start_date_end_date_idx" ON "vendor_availability"("vendor_id", "start_date", "end_date");

-- CreateIndex
CREATE UNIQUE INDEX "jobs_job_number_key" ON "jobs"("job_number");

-- CreateIndex
CREATE UNIQUE INDEX "jobs_awarded_bid_id_key" ON "jobs"("awarded_bid_id");

-- CreateIndex
CREATE INDEX "jobs_status_priority_idx" ON "jobs"("status", "priority");

-- CreateIndex
CREATE INDEX "jobs_property_id_status_idx" ON "jobs"("property_id", "status");

-- CreateIndex
CREATE INDEX "jobs_bid_deadline_idx" ON "jobs"("bid_deadline");

-- CreateIndex
CREATE INDEX "jobs_priority_status_claimed_by_vendor_id_idx" ON "jobs"("priority", "status", "claimed_by_vendor_id");

-- CreateIndex
CREATE UNIQUE INDEX "job_invitations_job_id_vendor_id_key" ON "job_invitations"("job_id", "vendor_id");

-- CreateIndex
CREATE INDEX "bids_job_id_status_idx" ON "bids"("job_id", "status");

-- CreateIndex
CREATE INDEX "bids_vendor_id_status_idx" ON "bids"("vendor_id", "status");

-- CreateIndex
CREATE UNIQUE INDEX "requisitions_job_id_key" ON "requisitions"("job_id");

-- CreateIndex
CREATE UNIQUE INDEX "requisitions_bid_id_key" ON "requisitions"("bid_id");

-- CreateIndex
CREATE UNIQUE INDEX "contracts_job_id_key" ON "contracts"("job_id");

-- CreateIndex
CREATE UNIQUE INDEX "contracts_bid_id_key" ON "contracts"("bid_id");

-- CreateIndex
CREATE UNIQUE INDEX "job_ratings_job_id_key" ON "job_ratings"("job_id");

-- CreateIndex
CREATE INDEX "job_ratings_vendor_id_idx" ON "job_ratings"("vendor_id");

-- CreateIndex
CREATE INDEX "trust_score_snapshots_vendor_id_computed_at_idx" ON "trust_score_snapshots"("vendor_id", "computed_at");

-- CreateIndex
CREATE INDEX "approval_flags_created_at_idx" ON "approval_flags"("created_at");

-- CreateIndex
CREATE INDEX "approval_flags_property_id_created_at_idx" ON "approval_flags"("property_id", "created_at");

-- CreateIndex
CREATE INDEX "chat_threads_vendor_id_last_message_at_idx" ON "chat_threads"("vendor_id", "last_message_at");

-- CreateIndex
CREATE UNIQUE INDEX "chat_threads_job_id_vendor_id_key" ON "chat_threads"("job_id", "vendor_id");

-- CreateIndex
CREATE INDEX "chat_messages_thread_id_created_at_idx" ON "chat_messages"("thread_id", "created_at");

-- CreateIndex
CREATE INDEX "job_documents_job_id_type_idx" ON "job_documents"("job_id", "type");

-- CreateIndex
CREATE INDEX "activity_logs_job_id_created_at_idx" ON "activity_logs"("job_id", "created_at");

-- CreateIndex
CREATE INDEX "activity_logs_entity_type_entity_id_created_at_idx" ON "activity_logs"("entity_type", "entity_id", "created_at");

-- CreateIndex
CREATE INDEX "notification_logs_job_id_channel_idx" ON "notification_logs"("job_id", "channel");

-- CreateIndex
CREATE INDEX "notification_logs_vendor_id_created_at_idx" ON "notification_logs"("vendor_id", "created_at");

-- CreateIndex
CREATE UNIQUE INDEX "push_subscriptions_endpoint_key" ON "push_subscriptions"("endpoint");

-- AddForeignKey
ALTER TABLE "vendor_users" ADD CONSTRAINT "vendor_users_vendor_id_fkey" FOREIGN KEY ("vendor_id") REFERENCES "vendors"("id") ON DELETE CASCADE ON UPDATE CASCADE;

-- AddForeignKey
ALTER TABLE "rent_rolls" ADD CONSTRAINT "rent_rolls_property_id_fkey" FOREIGN KEY ("property_id") REFERENCES "properties"("id") ON DELETE CASCADE ON UPDATE CASCADE;

-- AddForeignKey
ALTER TABLE "rent_rolls" ADD CONSTRAINT "rent_rolls_import_batch_id_fkey" FOREIGN KEY ("import_batch_id") REFERENCES "import_batches"("id") ON DELETE SET NULL ON UPDATE CASCADE;

-- AddForeignKey
ALTER TABLE "import_batches" ADD CONSTRAINT "import_batches_admin_user_id_fkey" FOREIGN KEY ("admin_user_id") REFERENCES "admin_users"("id") ON DELETE SET NULL ON UPDATE CASCADE;

-- AddForeignKey
ALTER TABLE "vendors" ADD CONSTRAINT "vendors_approved_by_admin_id_fkey" FOREIGN KEY ("approved_by_admin_id") REFERENCES "admin_users"("id") ON DELETE SET NULL ON UPDATE CASCADE;

-- AddForeignKey
ALTER TABLE "vendor_availability" ADD CONSTRAINT "vendor_availability_vendor_id_fkey" FOREIGN KEY ("vendor_id") REFERENCES "vendors"("id") ON DELETE CASCADE ON UPDATE CASCADE;

-- AddForeignKey
ALTER TABLE "jobs" ADD CONSTRAINT "jobs_property_id_fkey" FOREIGN KEY ("property_id") REFERENCES "properties"("id") ON DELETE RESTRICT ON UPDATE CASCADE;

-- AddForeignKey
ALTER TABLE "jobs" ADD CONSTRAINT "jobs_created_by_admin_id_fkey" FOREIGN KEY ("created_by_admin_id") REFERENCES "admin_users"("id") ON DELETE RESTRICT ON UPDATE CASCADE;

-- AddForeignKey
ALTER TABLE "jobs" ADD CONSTRAINT "jobs_claimed_by_vendor_id_fkey" FOREIGN KEY ("claimed_by_vendor_id") REFERENCES "vendors"("id") ON DELETE SET NULL ON UPDATE CASCADE;

-- AddForeignKey
ALTER TABLE "jobs" ADD CONSTRAINT "jobs_awarded_vendor_id_fkey" FOREIGN KEY ("awarded_vendor_id") REFERENCES "vendors"("id") ON DELETE SET NULL ON UPDATE CASCADE;

-- AddForeignKey
ALTER TABLE "jobs" ADD CONSTRAINT "jobs_awarded_bid_id_fkey" FOREIGN KEY ("awarded_bid_id") REFERENCES "bids"("id") ON DELETE SET NULL ON UPDATE CASCADE;

-- AddForeignKey
ALTER TABLE "job_invitations" ADD CONSTRAINT "job_invitations_job_id_fkey" FOREIGN KEY ("job_id") REFERENCES "jobs"("id") ON DELETE CASCADE ON UPDATE CASCADE;

-- AddForeignKey
ALTER TABLE "job_invitations" ADD CONSTRAINT "job_invitations_vendor_id_fkey" FOREIGN KEY ("vendor_id") REFERENCES "vendors"("id") ON DELETE CASCADE ON UPDATE CASCADE;

-- AddForeignKey
ALTER TABLE "job_invitations" ADD CONSTRAINT "job_invitations_invited_by_admin_id_fkey" FOREIGN KEY ("invited_by_admin_id") REFERENCES "admin_users"("id") ON DELETE RESTRICT ON UPDATE CASCADE;

-- AddForeignKey
ALTER TABLE "bids" ADD CONSTRAINT "bids_job_id_fkey" FOREIGN KEY ("job_id") REFERENCES "jobs"("id") ON DELETE CASCADE ON UPDATE CASCADE;

-- AddForeignKey
ALTER TABLE "bids" ADD CONSTRAINT "bids_vendor_id_fkey" FOREIGN KEY ("vendor_id") REFERENCES "vendors"("id") ON DELETE RESTRICT ON UPDATE CASCADE;

-- AddForeignKey
ALTER TABLE "bids" ADD CONSTRAINT "bids_approved_by_admin_id_fkey" FOREIGN KEY ("approved_by_admin_id") REFERENCES "admin_users"("id") ON DELETE SET NULL ON UPDATE CASCADE;

-- AddForeignKey
ALTER TABLE "requisitions" ADD CONSTRAINT "requisitions_job_id_fkey" FOREIGN KEY ("job_id") REFERENCES "jobs"("id") ON DELETE CASCADE ON UPDATE CASCADE;

-- AddForeignKey
ALTER TABLE "requisitions" ADD CONSTRAINT "requisitions_bid_id_fkey" FOREIGN KEY ("bid_id") REFERENCES "bids"("id") ON DELETE SET NULL ON UPDATE CASCADE;

-- AddForeignKey
ALTER TABLE "requisitions" ADD CONSTRAINT "requisitions_vendor_id_fkey" FOREIGN KEY ("vendor_id") REFERENCES "vendors"("id") ON DELETE RESTRICT ON UPDATE CASCADE;

-- AddForeignKey
ALTER TABLE "contracts" ADD CONSTRAINT "contracts_job_id_fkey" FOREIGN KEY ("job_id") REFERENCES "jobs"("id") ON DELETE CASCADE ON UPDATE CASCADE;

-- AddForeignKey
ALTER TABLE "contracts" ADD CONSTRAINT "contracts_bid_id_fkey" FOREIGN KEY ("bid_id") REFERENCES "bids"("id") ON DELETE SET NULL ON UPDATE CASCADE;

-- AddForeignKey
ALTER TABLE "contracts" ADD CONSTRAINT "contracts_vendor_id_fkey" FOREIGN KEY ("vendor_id") REFERENCES "vendors"("id") ON DELETE RESTRICT ON UPDATE CASCADE;

-- AddForeignKey
ALTER TABLE "job_ratings" ADD CONSTRAINT "job_ratings_job_id_fkey" FOREIGN KEY ("job_id") REFERENCES "jobs"("id") ON DELETE CASCADE ON UPDATE CASCADE;

-- AddForeignKey
ALTER TABLE "job_ratings" ADD CONSTRAINT "job_ratings_vendor_id_fkey" FOREIGN KEY ("vendor_id") REFERENCES "vendors"("id") ON DELETE RESTRICT ON UPDATE CASCADE;

-- AddForeignKey
ALTER TABLE "job_ratings" ADD CONSTRAINT "job_ratings_admin_user_id_fkey" FOREIGN KEY ("admin_user_id") REFERENCES "admin_users"("id") ON DELETE RESTRICT ON UPDATE CASCADE;

-- AddForeignKey
ALTER TABLE "trust_score_snapshots" ADD CONSTRAINT "trust_score_snapshots_vendor_id_fkey" FOREIGN KEY ("vendor_id") REFERENCES "vendors"("id") ON DELETE CASCADE ON UPDATE CASCADE;

-- AddForeignKey
ALTER TABLE "approval_flags" ADD CONSTRAINT "approval_flags_job_id_fkey" FOREIGN KEY ("job_id") REFERENCES "jobs"("id") ON DELETE CASCADE ON UPDATE CASCADE;

-- AddForeignKey
ALTER TABLE "approval_flags" ADD CONSTRAINT "approval_flags_bid_id_fkey" FOREIGN KEY ("bid_id") REFERENCES "bids"("id") ON DELETE CASCADE ON UPDATE CASCADE;

-- AddForeignKey
ALTER TABLE "approval_flags" ADD CONSTRAINT "approval_flags_property_id_fkey" FOREIGN KEY ("property_id") REFERENCES "properties"("id") ON DELETE RESTRICT ON UPDATE CASCADE;

-- AddForeignKey
ALTER TABLE "approval_flags" ADD CONSTRAINT "approval_flags_acknowledged_by_admin_id_fkey" FOREIGN KEY ("acknowledged_by_admin_id") REFERENCES "admin_users"("id") ON DELETE SET NULL ON UPDATE CASCADE;

-- AddForeignKey
ALTER TABLE "approval_settings" ADD CONSTRAINT "approval_settings_updated_by_admin_id_fkey" FOREIGN KEY ("updated_by_admin_id") REFERENCES "admin_users"("id") ON DELETE SET NULL ON UPDATE CASCADE;

-- AddForeignKey
ALTER TABLE "chat_threads" ADD CONSTRAINT "chat_threads_job_id_fkey" FOREIGN KEY ("job_id") REFERENCES "jobs"("id") ON DELETE CASCADE ON UPDATE CASCADE;

-- AddForeignKey
ALTER TABLE "chat_threads" ADD CONSTRAINT "chat_threads_vendor_id_fkey" FOREIGN KEY ("vendor_id") REFERENCES "vendors"("id") ON DELETE CASCADE ON UPDATE CASCADE;

-- AddForeignKey
ALTER TABLE "chat_messages" ADD CONSTRAINT "chat_messages_thread_id_fkey" FOREIGN KEY ("thread_id") REFERENCES "chat_threads"("id") ON DELETE CASCADE ON UPDATE CASCADE;

-- AddForeignKey
ALTER TABLE "chat_messages" ADD CONSTRAINT "chat_messages_sender_admin_id_fkey" FOREIGN KEY ("sender_admin_id") REFERENCES "admin_users"("id") ON DELETE SET NULL ON UPDATE CASCADE;

-- AddForeignKey
ALTER TABLE "chat_messages" ADD CONSTRAINT "chat_messages_sender_vendor_user_id_fkey" FOREIGN KEY ("sender_vendor_user_id") REFERENCES "vendor_users"("id") ON DELETE SET NULL ON UPDATE CASCADE;

-- AddForeignKey
ALTER TABLE "job_documents" ADD CONSTRAINT "job_documents_job_id_fkey" FOREIGN KEY ("job_id") REFERENCES "jobs"("id") ON DELETE CASCADE ON UPDATE CASCADE;

-- AddForeignKey
ALTER TABLE "job_documents" ADD CONSTRAINT "job_documents_vendor_id_fkey" FOREIGN KEY ("vendor_id") REFERENCES "vendors"("id") ON DELETE SET NULL ON UPDATE CASCADE;

-- AddForeignKey
ALTER TABLE "activity_logs" ADD CONSTRAINT "activity_logs_job_id_fkey" FOREIGN KEY ("job_id") REFERENCES "jobs"("id") ON DELETE CASCADE ON UPDATE CASCADE;

-- AddForeignKey
ALTER TABLE "activity_logs" ADD CONSTRAINT "activity_logs_actor_admin_id_fkey" FOREIGN KEY ("actor_admin_id") REFERENCES "admin_users"("id") ON DELETE SET NULL ON UPDATE CASCADE;

-- AddForeignKey
ALTER TABLE "activity_logs" ADD CONSTRAINT "activity_logs_actor_vendor_id_fkey" FOREIGN KEY ("actor_vendor_id") REFERENCES "vendors"("id") ON DELETE SET NULL ON UPDATE CASCADE;

-- AddForeignKey
ALTER TABLE "notification_logs" ADD CONSTRAINT "notification_logs_vendor_id_fkey" FOREIGN KEY ("vendor_id") REFERENCES "vendors"("id") ON DELETE SET NULL ON UPDATE CASCADE;

-- AddForeignKey
ALTER TABLE "notification_logs" ADD CONSTRAINT "notification_logs_admin_user_id_fkey" FOREIGN KEY ("admin_user_id") REFERENCES "admin_users"("id") ON DELETE SET NULL ON UPDATE CASCADE;

-- AddForeignKey
ALTER TABLE "notification_logs" ADD CONSTRAINT "notification_logs_job_id_fkey" FOREIGN KEY ("job_id") REFERENCES "jobs"("id") ON DELETE SET NULL ON UPDATE CASCADE;

-- AddForeignKey
ALTER TABLE "push_subscriptions" ADD CONSTRAINT "push_subscriptions_vendor_user_id_fkey" FOREIGN KEY ("vendor_user_id") REFERENCES "vendor_users"("id") ON DELETE CASCADE ON UPDATE CASCADE;

-- AddForeignKey
ALTER TABLE "push_subscriptions" ADD CONSTRAINT "push_subscriptions_admin_user_id_fkey" FOREIGN KEY ("admin_user_id") REFERENCES "admin_users"("id") ON DELETE CASCADE ON UPDATE CASCADE;
