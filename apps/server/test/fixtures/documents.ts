import { zipSync, strToU8 } from "fflate";

export const maxTemplateUploadBytes = 25 * 1024 * 1024;

export const templateContentTypesXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="bin" ContentType="application/octet-stream"/><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>`;

const templateRelationshipsXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>`;

export const docxXmlFixture = ({
  additionalParts = {},
  contentTypes = templateContentTypesXml,
  document,
  paddingBytes = 0,
  relationships = templateRelationshipsXml,
}: {
  additionalParts?: Record<string, Uint8Array>;
  contentTypes?: string;
  document: string;
  paddingBytes?: number;
  relationships?: string;
}): Uint8Array =>
  zipSync(
    {
      "[Content_Types].xml": strToU8(contentTypes),
      "_rels/.rels": strToU8(relationships),
      "word/document.xml": strToU8(document),
      "word/media/padding.bin": new Uint8Array(paddingBytes),
      ...additionalParts,
    },
    { level: 0 }
  );

export const docxFixture = (label: string, paddingBytes = 0): Uint8Array =>
  docxXmlFixture({
    document: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:t>${label}</w:t></w:r></w:p><w:sectPr/></w:body></w:document>`,
    paddingBytes,
  });

export const contentControlDocument = (controls: string): string =>
  `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" xmlns:word="http://purl.oclc.org/ooxml/wordprocessingml/main" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"><w:body>${controls}<w:sectPr/></w:body></w:document>`;

export const contentControl = ({
  alias,
  placeholderText,
  tag,
  type,
}: {
  alias?: string;
  placeholderText?: string;
  tag: string;
  type: string;
}): string => {
  const aliasProperty =
    alias === undefined ? "" : `<w:alias w:val="${alias}"/>`;
  const placeholderProperty =
    placeholderText === undefined ? "" : "<w:showingPlcHdr/>";
  const content = placeholderText ?? "fixture";
  return `<w:sdt><w:sdtPr>${aliasProperty}<w:tag w:val="${tag}"/>${type}${placeholderProperty}</w:sdtPr><w:sdtContent><w:r><w:t>${content}</w:t></w:r></w:sdtContent></w:sdt>`;
};

const pictureDrawing = (relationshipId: string): string =>
  `<w:drawing xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><a:blip r:embed="${relationshipId}"/></w:drawing>`;

const onePixelPng = Uint8Array.from(
  Buffer.from(
    "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+A8AAQUBAScY42YAAAAASUVORK5CYII=",
    "base64"
  )
);

export const pngFixture = (
  width = 1,
  height = 1,
  byteLength = onePixelPng.byteLength
): Uint8Array<ArrayBuffer> => {
  const bytes = new Uint8Array(byteLength);
  bytes.set(onePixelPng.subarray(0, Math.min(onePixelPng.length, byteLength)));
  const writeUint32 = (value: number, offset: number) => {
    bytes[offset] = Math.floor(value / 0x1_00_00_00) % 0x1_00;
    bytes[offset + 1] = Math.floor(value / 0x1_00_00) % 0x1_00;
    bytes[offset + 2] = Math.floor(value / 0x1_00) % 0x1_00;
    bytes[offset + 3] = value % 0x1_00;
  };
  writeUint32(width, 16);
  writeUint32(height, 20);
  return bytes;
};

export const jpegFixture = (width = 1, height = 1): Uint8Array<ArrayBuffer> =>
  Uint8Array.from([
    0xff,
    0xd8,
    0xff,
    0xc0,
    0x00,
    0x11,
    0x08,
    Math.floor(height / 0x1_00),
    height % 0x1_00,
    Math.floor(width / 0x1_00),
    width % 0x1_00,
    0x03,
    0x01,
    0x11,
    0x00,
    0x02,
    0x11,
    0x00,
    0x03,
    0x11,
    0x00,
    0xff,
    0xd9,
  ]);

export const gifFixture = Uint8Array.from(
  Buffer.from("R0lGODlhAQABAIAAAAAAAP///ywAAAAAAQABAAACAUwAOw==", "base64")
);

export const pictureDocumentFixture = ({
  images = [],
  includeStaticImage = false,
  showingPlaceholder = false,
}: {
  images?: { bytes: Uint8Array; extension: string }[];
  includeStaticImage?: boolean;
  showingPlaceholder?: boolean;
} = {}): Uint8Array => {
  const imageEntries = images.map((image, index) => ({
    bytes: image.bytes,
    extension: image.extension,
    relationshipId: `rIdPicture${index + 1}`,
  }));
  if (includeStaticImage) {
    imageEntries.push({
      bytes: pngFixture(),
      extension: "png",
      relationshipId: "rIdStatic",
    });
  }
  const relationships = imageEntries.map(
    ({ extension, relationshipId }, index) =>
      `<Relationship Id="${relationshipId}" Target="media/image${index + 1}.${extension}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"/>`
  );
  const staticDrawing = includeStaticImage ? pictureDrawing("rIdStatic") : "";
  const pictureDrawings = imageEntries
    .slice(0, images.length)
    .map(({ relationshipId }) => pictureDrawing(relationshipId))
    .join("");
  const additionalParts: Record<string, Uint8Array> = {
    "word/_rels/document.xml.rels": strToU8(
      `<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${relationships.join("")}</Relationships>`
    ),
  };
  for (const [index, image] of imageEntries.entries()) {
    additionalParts[`word/media/image${index + 1}.${image.extension}`] =
      image.bytes;
  }
  const extensions = new Set(imageEntries.map(({ extension }) => extension));
  const contentTypes = templateContentTypesXml.replace(
    "</Types>",
    `${[...extensions]
      .map(
        (extension) =>
          `<Default Extension="${extension}" ContentType="image/${extension === "jpg" ? "jpeg" : extension}"/>`
      )
      .join("")}</Types>`
  );
  const pictureProperties = showingPlaceholder ? "<w:showingPlcHdr/>" : "";
  return docxXmlFixture({
    additionalParts,
    contentTypes,
    document: contentControlDocument(
      `${staticDrawing}<w:sdt><w:sdtPr><w:tag w:val="photo"/><w:picture/>${pictureProperties}</w:sdtPr><w:sdtContent><w:r>${pictureDrawings || "<w:t>empty</w:t>"}</w:r></w:sdtContent></w:sdt>`
    ),
  });
};

export const strictDocxFixture = (label: string): Uint8Array =>
  docxXmlFixture({
    contentTypes: `<?xml version="1.0"?><ct:Types xmlns:ct="http://schemas.openxmlformats.org/package/2006/content-types"><ct:Override ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml" PartName="/word/document.xml"/></ct:Types>`,
    document: `<?xml version="1.0"?><word:document xmlns:word="http://purl.oclc.org/ooxml/wordprocessingml/main"><word:body><word:p><word:r><word:t>${label}</word:t></word:r></word:p></word:body></word:document>`,
    relationships: `<?xml version="1.0"?><pkg:Relationships xmlns:pkg="http://schemas.openxmlformats.org/package/2006/relationships"><pkg:Relationship Id="rId1" Target="/word/document.xml" Type="http://purl.oclc.org/ooxml/officeDocument/relationships/officeDocument"/></pkg:Relationships>`,
  });

const utf16Xml = (value: string, byteOrder: "be" | "le"): Uint8Array => {
  const littleEndian = Buffer.from(value, "utf16le");
  const bytes = new Uint8Array(littleEndian.byteLength + 2);
  bytes[0] = byteOrder === "le" ? 0xff : 0xfe;
  bytes[1] = byteOrder === "le" ? 0xfe : 0xff;
  for (let index = 0; index < littleEndian.byteLength; index += 2) {
    const target = index + 2;
    const firstByte = littleEndian[index] ?? 0;
    const secondByte = littleEndian[index + 1] ?? 0;
    bytes[target] = byteOrder === "le" ? firstByte : secondByte;
    bytes[target + 1] = byteOrder === "le" ? secondByte : firstByte;
  }
  return bytes;
};

export const utf16DocxFixture = (): Uint8Array =>
  zipSync(
    {
      "[Content_Types].xml": utf16Xml(
        templateContentTypesXml.replace(
          'encoding="UTF-8"',
          'encoding="UTF-16"'
        ),
        "le"
      ),
      "_rels/.rels": utf16Xml(
        templateRelationshipsXml.replace(
          'encoding="UTF-8"',
          'encoding="UTF-16"'
        ),
        "be"
      ),
      "word/document.xml": utf16Xml(
        '<?xml version="1.0" encoding="UTF-16"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body/></w:document>',
        "le"
      ),
    },
    { level: 0 }
  );

export const sizedDocxFixture = (
  label: string,
  byteLength: number
): Uint8Array => {
  const emptyPadding = docxFixture(label);
  const paddingLength = byteLength - emptyPadding.byteLength;
  if (paddingLength < 0) {
    throw new Error("The requested DOCX fixture size is too small");
  }
  const fixture = docxFixture(label, paddingLength);
  if (fixture.byteLength !== byteLength) {
    throw new Error("The DOCX fixture did not reach the requested size");
  }
  return fixture;
};
