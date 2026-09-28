CREATE TYPE "FillMethod" AS ENUM ('onlyoffice', 'native');

ALTER TABLE "forms"
ADD COLUMN "fill_method" "FillMethod" NOT NULL DEFAULT 'onlyoffice';
