CREATE SEQUENCE "legacy_sso_generation_seq" AS BIGINT;

ALTER TABLE "legacy_account_link_requests"
ADD COLUMN "reviewed_generation" BIGINT;
