ALTER TABLE "manifest_fields"
  ADD COLUMN "label" TEXT,
  ADD COLUMN "placeholder" TEXT,
  ADD COLUMN "position" INTEGER;

WITH ranked_fields AS (
  SELECT
    "id",
    (ROW_NUMBER() OVER (PARTITION BY "manifest_id" ORDER BY "tag") - 1)::integer AS "position"
  FROM "manifest_fields"
)
UPDATE "manifest_fields" AS field
SET
  "label" = field."tag",
  "position" = ranked_fields."position"
FROM ranked_fields
WHERE field."id" = ranked_fields."id";

ALTER TABLE "manifest_fields"
  ALTER COLUMN "label" SET NOT NULL,
  ALTER COLUMN "position" SET NOT NULL;

CREATE UNIQUE INDEX "manifest_fields_manifest_id_position_key"
  ON "manifest_fields"("manifest_id", "position");
