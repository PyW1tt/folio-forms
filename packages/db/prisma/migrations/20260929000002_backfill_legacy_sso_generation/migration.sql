UPDATE "legacy_account_link_requests"
SET "reviewed_generation" = nextval('"legacy_sso_generation_seq"')
WHERE "status" IN ('approved', 'linked');
