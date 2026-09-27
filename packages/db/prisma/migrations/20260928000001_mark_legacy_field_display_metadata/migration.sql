ALTER TABLE "field_manifests"
  ADD COLUMN "display_metadata_version" INTEGER NOT NULL DEFAULT 0;

ALTER TABLE "field_manifests"
  ALTER COLUMN "display_metadata_version" SET DEFAULT 1;
