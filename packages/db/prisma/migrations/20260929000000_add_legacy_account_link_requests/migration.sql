CREATE TYPE "LegacyAccountLinkStatus" AS ENUM ('pending', 'approved', 'rejected', 'linked');

CREATE TABLE "legacy_account_link_requests" (
    "id" UUID NOT NULL,
    "provider_id" TEXT NOT NULL,
    "subject" TEXT NOT NULL,
    "email" TEXT NOT NULL,
    "user_id" TEXT NOT NULL,
    "status" "LegacyAccountLinkStatus" NOT NULL DEFAULT 'pending',
    "reviewed_by_id" TEXT,
    "reviewed_at" TIMESTAMPTZ(3),
    "created_at" TIMESTAMPTZ(3) NOT NULL DEFAULT CURRENT_TIMESTAMP,
    "updated_at" TIMESTAMPTZ(3) NOT NULL DEFAULT CURRENT_TIMESTAMP,

    CONSTRAINT "legacy_account_link_requests_pkey" PRIMARY KEY ("id")
);

CREATE UNIQUE INDEX "legacy_account_link_requests_provider_id_subject_key"
ON "legacy_account_link_requests"("provider_id", "subject");

CREATE INDEX "legacy_account_link_requests_status_id_idx"
ON "legacy_account_link_requests"("status", "id");

CREATE INDEX "legacy_account_link_requests_user_id_status_idx"
ON "legacy_account_link_requests"("user_id", "status");

ALTER TABLE "legacy_account_link_requests"
ADD CONSTRAINT "legacy_account_link_requests_user_id_fkey"
FOREIGN KEY ("user_id") REFERENCES "user"("id") ON DELETE CASCADE ON UPDATE CASCADE;

ALTER TABLE "legacy_account_link_requests"
ADD CONSTRAINT "legacy_account_link_requests_reviewed_by_id_fkey"
FOREIGN KEY ("reviewed_by_id") REFERENCES "user"("id") ON DELETE SET NULL ON UPDATE CASCADE;
