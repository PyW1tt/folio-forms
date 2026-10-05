import { strToU8, zipSync } from "fflate";

export const maxParagraphs = 20;

export const maxFields = 40;

const xmlNamespace =
  "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

export interface GeneratedField {
  label: string;
  placeholder: string;
  tag: string;
}

export interface GeneratedTemplate {
  description: string;
  document: Uint8Array;
  fields: GeneratedField[];
  paragraphs: string[];
  title: string;
}

export interface RevisionEdits {
  editedParagraphs?: unknown;
  removedFieldTags?: unknown;
}

const xmlEscape = (value: string): string =>
  value
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&apos;");

const xmlText = (value: string): string => {
  if (
    [...value].some((character) => {
      const codePoint = character.codePointAt(0) ?? 0;
      return !(
        codePoint === 0x09 ||
        codePoint === 0x0a ||
        codePoint === 0x0d ||
        (codePoint >= 0x20 && codePoint <= 0xd7_ff) ||
        (codePoint >= 0xe0_00 && codePoint <= 0xff_fd) ||
        (codePoint >= 0x1_00_00 && codePoint <= 0x10_ff_ff)
      );
    })
  ) {
    throw new Error("Generated text contains an invalid XML character");
  }
  return xmlEscape(value);
};

const requiredText = (
  value: unknown,
  name: string,
  maxLength: number
): string => {
  if (
    typeof value !== "string" ||
    !value.trim() ||
    value.trim().length > maxLength
  ) {
    throw new Error(`${name} is invalid`);
  }
  return value.trim();
};

export const validateGeneratedTemplate = (value: {
  description: unknown;
  fields: unknown;
  paragraphs: unknown;
  title: unknown;
}): Omit<GeneratedTemplate, "document"> => {
  const title = requiredText(value.title, "title", 200);
  const description =
    typeof value.description === "string" ? value.description.trim() : null;
  if (description === null || description.length > 2000) {
    throw new Error("description is invalid");
  }
  if (
    !Array.isArray(value.paragraphs) ||
    value.paragraphs.length > maxParagraphs ||
    value.paragraphs.some(
      (paragraph) => typeof paragraph !== "string" || paragraph.length > 2000
    )
  ) {
    throw new Error("paragraphs are invalid");
  }
  if (
    !Array.isArray(value.fields) ||
    value.fields.length === 0 ||
    value.fields.length > maxFields
  ) {
    throw new Error("fields are invalid");
  }
  const tags = new Set<string>();
  const fields = value.fields.map((field) => {
    if (!field || typeof field !== "object" || Array.isArray(field)) {
      throw new Error("field is invalid");
    }
    const input = field as Record<string, unknown>;
    const label = requiredText(input.label, "field label", 120);
    const placeholder = requiredText(
      input.placeholder,
      "field placeholder",
      120
    );
    const tag = requiredText(input.tag, "field tag", 64);
    if (!/^[a-z][a-z0-9_]*$/u.test(tag) || tags.has(tag)) {
      throw new Error("field tag is invalid or duplicated");
    }
    tags.add(tag);
    return { label, placeholder, tag };
  });
  const paragraphs = value.paragraphs.map((paragraph) =>
    (paragraph as string).trim()
  );
  return { description, fields, paragraphs, title };
};

const paragraphXml = (text: string): string =>
  `<w:p><w:r><w:t xml:space="preserve">${xmlText(text)}</w:t></w:r></w:p>`;

export const createTemplateDocx = (input: {
  description: unknown;
  fields: unknown;
  paragraphs: unknown;
  title: unknown;
}): GeneratedTemplate => {
  const content = validateGeneratedTemplate(input);
  const fieldsXml = content.fields
    .map(
      (field) =>
        `${paragraphXml(`${field.label}:`)}<w:p><w:sdt><w:sdtPr><w:alias w:val="${xmlText(field.label)}"/><w:tag w:val="${xmlText(field.tag)}"/><w:text/><w:showingPlcHdr/></w:sdtPr><w:sdtContent><w:r><w:t>${xmlText(field.placeholder)}</w:t></w:r></w:sdtContent></w:sdt></w:p>`
    )
    .join("");
  const descriptionXml = content.description
    ? paragraphXml(content.description)
    : "";
  const documentXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document xmlns:w="${xmlNamespace}"><w:body>${paragraphXml(content.title)}${descriptionXml}${content.paragraphs.map(paragraphXml).join("")}${fieldsXml}<w:sectPr/></w:body></w:document>`;
  const document = zipSync(
    {
      "[Content_Types].xml": strToU8(
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>'
      ),
      "_rels/.rels": strToU8(
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>'
      ),
      "word/document.xml": strToU8(documentXml),
    },
    { level: 6 }
  );
  return { ...content, document };
};

export const preserveUnchangedContent = (
  current: GeneratedTemplate | undefined,
  candidate: GeneratedTemplate,
  edits: RevisionEdits
): void => {
  const editedParagraphs =
    edits.editedParagraphs === undefined ? [] : edits.editedParagraphs;
  const removedFieldTags =
    edits.removedFieldTags === undefined ? [] : edits.removedFieldTags;
  if (
    !Array.isArray(editedParagraphs) ||
    editedParagraphs.length > maxParagraphs ||
    editedParagraphs.some(
      (paragraph: unknown) =>
        typeof paragraph !== "string" ||
        !current?.paragraphs.includes(paragraph)
    )
  ) {
    throw new Error(
      "Revision edited paragraphs must reference current paragraphs"
    );
  }
  if (
    !Array.isArray(removedFieldTags) ||
    removedFieldTags.length > maxFields ||
    removedFieldTags.some(
      (tag: unknown) =>
        typeof tag !== "string" ||
        !current?.fields.some((field) => field.tag === tag)
    )
  ) {
    throw new Error(
      "Revision removed field tags must reference current controls"
    );
  }
  if (!current) {
    return;
  }
  for (const paragraph of current.paragraphs) {
    if (
      !candidate.paragraphs.includes(paragraph) &&
      !editedParagraphs.includes(paragraph)
    ) {
      throw new Error("Revision unexpectedly removed an existing paragraph");
    }
  }
  const nextTags = new Set(candidate.fields.map((field) => field.tag));
  for (const field of current.fields) {
    if (!nextTags.has(field.tag) && !removedFieldTags.includes(field.tag)) {
      throw new Error(
        "Revision unexpectedly removed an existing tagged control"
      );
    }
  }
};
