import { unzipSync } from "fflate";
import { SaxesParser } from "saxes";

// oxlint-disable func-style prefer-destructuring no-await-in-loop no-use-before-define no-nested-ternary complexity no-shadow -- Route modules keep declaration order and sequential persistence invariants.
import { fail } from "../http/errors";

export const maxTemplateArchiveExpandedBytes = 64 * 1024 * 1024;

const maxTemplateArchiveEntries = 2048;

export const templatePackageRelationshipNamespace =
  "http://schemas.openxmlformats.org/package/2006/relationships";

export const templatePackageContentTypesNamespace =
  "http://schemas.openxmlformats.org/package/2006/content-types";

const templateOfficeDocumentRelationships = new Set([
  "http://purl.oclc.org/ooxml/officeDocument/relationships/officeDocument",
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument",
]);

export const templateWordMainNamespace =
  "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

export const templateWordStrictNamespace =
  "http://purl.oclc.org/ooxml/wordprocessingml/main";

export const templateCheckboxNamespace =
  "http://schemas.microsoft.com/office/word/2010/wordml";

export const templateWord2012Namespace =
  "http://schemas.microsoft.com/office/word/2012/wordml";

export const templateMarkupCompatibilityNamespace =
  "http://schemas.openxmlformats.org/markup-compatibility/2006";

export const templateXmlnsNamespace = "http://www.w3.org/2000/xmlns/";

export const templateWordNamespaces = new Set([
  templateWordStrictNamespace,
  templateWordMainNamespace,
]);

export const templateDrawingMlNamespaces = new Set([
  "http://purl.oclc.org/ooxml/drawingml/main",
  "http://schemas.openxmlformats.org/drawingml/2006/main",
]);

export const templateOfficeRelationshipNamespaces = new Set([
  "http://purl.oclc.org/ooxml/officeDocument/relationships",
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
]);

export const templateGlossaryDocumentRelationships = new Set([
  "http://purl.oclc.org/ooxml/officeDocument/relationships/glossaryDocument",
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships/glossaryDocument",
]);

export const templateVmlNamespace = "urn:schemas-microsoft-com:vml";

export const templateControlNamespaces = new Set([
  ...templateWordNamespaces,
  templateCheckboxNamespace,
  templateWord2012Namespace,
]);

export const templateControlKey = (namespace: string, local: string): string =>
  `${namespace}#${local}`;

const templateDocumentContentType =
  "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml";

interface TemplateXmlAttribute {
  local: string;
  uri: string;
  value: string;
}

export interface TemplateXmlElement {
  attributes: TemplateXmlAttribute[];
  local: string;
  uri: string;
}

interface TemplateXmlVisitor {
  close?: (element: TemplateXmlElement) => void;
  open?: (element: TemplateXmlElement) => void;
  text?: (value: string) => void;
}

export function templateXmlElement(tag: {
  attributes: Record<string, { local: string; uri: string; value: string }>;
  local: string;
  uri: string;
}): TemplateXmlElement {
  return {
    attributes: Object.values(tag.attributes).map((attribute) => ({
      local: attribute.local,
      uri: attribute.uri,
      value: attribute.value,
    })),
    local: tag.local,
    uri: tag.uri,
  };
}

export function parseTemplateXml(
  xml: string,
  visitor: TemplateXmlVisitor = {}
): void {
  const parser = new SaxesParser({ position: false, xmlns: true });
  parser.on("doctype", () => {
    throw new Error("DOCX XML document types are not allowed");
  });
  parser.on("opentag", (tag) => {
    visitor.open?.(templateXmlElement(tag));
  });
  parser.on("closetag", (tag) => {
    visitor.close?.(templateXmlElement(tag));
  });
  parser.on("text", (value) => {
    visitor.text?.(value);
  });
  try {
    parser.write(xml).close();
  } catch {
    fail(422, "invalid_template", "The DOCX package contains invalid XML");
  }
}

export function templateAttribute(
  element: TemplateXmlElement,
  local: string,
  uri = ""
): string | undefined {
  return element.attributes.find(
    (attribute) => attribute.local === local && attribute.uri === uri
  )?.value;
}

function readTemplateArchive(
  bytes: Uint8Array,
  include: (archivePath: string) => boolean
): Record<string, Uint8Array> {
  let expandedBytes = 0;
  let entryCount = 0;
  const names = new Set<string>();
  try {
    return unzipSync(bytes, {
      filter: (file) => {
        entryCount += 1;
        const archivePath = file.name;
        const segments = archivePath.split("/");
        if (
          !archivePath ||
          archivePath.startsWith("/") ||
          archivePath.includes("\\") ||
          archivePath.includes("\0") ||
          segments.some(
            (segment, index) =>
              segment === "." ||
              segment === ".." ||
              (segment.length === 0 && index !== segments.length - 1)
          ) ||
          names.has(archivePath)
        ) {
          throw new Error("Unsafe DOCX archive path");
        }
        names.add(archivePath);
        if (
          !Number.isSafeInteger(file.originalSize) ||
          file.originalSize < 0 ||
          expandedBytes > maxTemplateArchiveExpandedBytes - file.originalSize
        ) {
          throw new Error("DOCX archive expands beyond the safety limit");
        }
        expandedBytes += file.originalSize;
        if (entryCount > maxTemplateArchiveEntries) {
          throw new Error("DOCX archive contains too many entries");
        }
        return include(archivePath);
      },
    });
  } catch {
    fail(422, "invalid_template", "The template is not a safe DOCX archive");
  }
}

export function templateArchiveText(
  archive: Record<string, Uint8Array>,
  archivePath: string
): string {
  const bytes = archive[archivePath];
  if (!bytes || bytes.byteLength === 0) {
    fail(422, "invalid_template", "The DOCX package is incomplete");
  }
  try {
    let decodedBytes = bytes;
    let encoding: "utf-16" | "utf-8" = "utf-8";
    if (
      (bytes[0] === 0xff && bytes[1] === 0xfe) ||
      (bytes[0] === 0x3c && bytes[1] === 0x00)
    ) {
      encoding = "utf-16";
    } else if (
      (bytes[0] === 0xfe && bytes[1] === 0xff) ||
      (bytes[0] === 0x00 && bytes[1] === 0x3c)
    ) {
      if (bytes.byteLength % 2 !== 0) {
        throw new Error("UTF-16 XML must contain complete code units");
      }
      decodedBytes = new Uint8Array(bytes);
      for (let index = 0; index < decodedBytes.byteLength; index += 2) {
        const firstByte = decodedBytes[index] ?? 0;
        decodedBytes[index] = decodedBytes[index + 1] ?? 0;
        decodedBytes[index + 1] = firstByte;
      }
      encoding = "utf-16";
    }
    return new TextDecoder(encoding, { fatal: true }).decode(decodedBytes);
  } catch {
    fail(422, "invalid_template", "The DOCX package contains invalid XML");
  }
}

const externalRelationshipTargetPattern = /^[a-z][a-z\d+.-]*:/iu;

function isExternalRelationshipTarget(target: string | undefined): boolean {
  const normalized = target?.trim() ?? "";
  return (
    normalized.startsWith("//") ||
    normalized.startsWith("\\\\") ||
    externalRelationshipTargetPattern.test(normalized)
  );
}

function validateOfficeRelationships(
  archive: Record<string, Uint8Array>
): void {
  for (const archivePath of Object.keys(archive)) {
    if (!archivePath.toLowerCase().endsWith(".rels")) {
      continue;
    }
    parseTemplateXml(templateArchiveText(archive, archivePath), {
      open: (element) => {
        if (
          element.local !== "Relationship" ||
          element.uri !== templatePackageRelationshipNamespace
        ) {
          return;
        }
        const targetMode = templateAttribute(element, "TargetMode");
        const target = templateAttribute(element, "Target");
        if (
          targetMode?.trim().toLowerCase() === "external" ||
          isExternalRelationshipTarget(target)
        ) {
          fail(
            422,
            "invalid_template",
            "External DOCX relationships are not allowed"
          );
        }
      },
    });
  }
}

export function validateOfficeRelationshipsBytes(bytes: Uint8Array): void {
  const archive = readTemplateArchive(bytes, (archivePath) =>
    archivePath.toLowerCase().endsWith(".rels")
  );
  validateOfficeRelationships(archive);
}

function isTemplateXmlContentType(contentType: string | undefined): boolean {
  const normalized = contentType?.split(";", 1)[0]?.trim().toLowerCase();
  return (
    normalized === "application/xml" ||
    normalized === "text/xml" ||
    normalized === "application/vnd.openxmlformats-officedocument.vmldrawing" ||
    normalized?.endsWith("+xml") === true
  );
}

export function templateRelationshipPartPath(partPath: string): string {
  const separator = partPath.lastIndexOf("/");
  const directory = separator === -1 ? "" : partPath.slice(0, separator + 1);
  const filename = partPath.slice(separator + 1);
  return `${directory}_rels/${filename}.rels`;
}

export function resolveTemplateRelationshipTarget(
  sourcePath: string,
  target: string | undefined
): string | null {
  const normalizedTarget = target?.trim().split("#", 1)[0] ?? "";
  if (!normalizedTarget || isExternalRelationshipTarget(normalizedTarget)) {
    return null;
  }
  const separator = sourcePath.lastIndexOf("/");
  const directory = separator === -1 ? "" : sourcePath.slice(0, separator + 1);
  const segments =
    `${normalizedTarget.startsWith("/") ? "" : directory}${normalizedTarget.replace(/^\/+/u, "")}`.split(
      "/"
    );
  const normalizedSegments: string[] = [];
  for (const segment of segments) {
    if (!segment || segment === ".") {
      continue;
    }
    if (segment === "..") {
      if (normalizedSegments.length === 0) {
        return null;
      }
      normalizedSegments.pop();
      continue;
    }
    normalizedSegments.push(segment);
  }
  return normalizedSegments.join("/");
}

function reachableTemplateParts(
  archive: Record<string, Uint8Array>,
  startPath: string
): Set<string> {
  const reachable = new Set<string>([startPath]);
  const queue = [startPath];
  while (queue.length > 0) {
    const sourcePath = queue.shift();
    if (!sourcePath) {
      continue;
    }
    const relationshipPath = templateRelationshipPartPath(sourcePath);
    const relationshipBytes = archive[relationshipPath];
    if (!relationshipBytes) {
      continue;
    }
    parseTemplateXml(templateArchiveText(archive, relationshipPath), {
      open: (element) => {
        if (
          element.local !== "Relationship" ||
          element.uri !== templatePackageRelationshipNamespace ||
          templateAttribute(element, "TargetMode") !== undefined
        ) {
          return;
        }
        const targetPath = resolveTemplateRelationshipTarget(
          sourcePath,
          templateAttribute(element, "Target")
        );
        if (
          targetPath &&
          Object.hasOwn(archive, targetPath) &&
          !reachable.has(targetPath)
        ) {
          reachable.add(targetPath);
          queue.push(targetPath);
        }
      },
    });
  }
  return reachable;
}

export function reachableTemplateControlParts(
  archive: Record<string, Uint8Array>,
  xmlPaths: Set<string>
): Set<string> {
  const reachable = reachableTemplateParts(archive, "word/document.xml");
  return new Set(
    [...xmlPaths].filter(
      (archivePath) =>
        reachable.has(archivePath) &&
        templateControlPartPattern.test(archivePath)
    )
  );
}

function validateTemplatePackageArchive(
  archive: Record<string, Uint8Array>
): Set<string> {
  const contentTypes = templateArchiveText(archive, "[Content_Types].xml");
  const relationships = templateArchiveText(archive, "_rels/.rels");
  const document = templateArchiveText(archive, "word/document.xml");
  const defaultContentTypes = new Map<string, string>();
  const overrideContentTypes = new Map<string, string>();
  let contentTypesRoot = false;
  let contentTypesRootSeen = false;
  parseTemplateXml(contentTypes, {
    open: (element) => {
      if (!contentTypesRootSeen) {
        contentTypesRootSeen = true;
        contentTypesRoot =
          element.local === "Types" &&
          element.uri === templatePackageContentTypesNamespace;
      }
      if (element.uri !== templatePackageContentTypesNamespace) {
        return;
      }
      const contentType = templateAttribute(element, "ContentType");
      if (element.local === "Default") {
        const extension = templateAttribute(
          element,
          "Extension"
        )?.toLowerCase();
        if (!extension || !contentType || defaultContentTypes.has(extension)) {
          fail(422, "invalid_template", "The DOCX content types are invalid");
        }
        defaultContentTypes.set(extension, contentType);
      } else if (element.local === "Override") {
        const partName = templateAttribute(element, "PartName");
        if (
          !partName?.startsWith("/") ||
          !contentType ||
          overrideContentTypes.has(partName)
        ) {
          fail(422, "invalid_template", "The DOCX content types are invalid");
        }
        overrideContentTypes.set(partName, contentType);
      }
    },
  });
  const documentOverride =
    overrideContentTypes.get("/word/document.xml") ===
    templateDocumentContentType;

  let relationshipsRoot = false;
  let relationshipsRootSeen = false;
  let officeDocumentRelationship = false;
  parseTemplateXml(relationships, {
    open: (element) => {
      if (!relationshipsRootSeen) {
        relationshipsRootSeen = true;
        relationshipsRoot =
          element.local === "Relationships" &&
          element.uri === templatePackageRelationshipNamespace;
      }
      const target = templateAttribute(element, "Target");
      if (
        element.local === "Relationship" &&
        element.uri === templatePackageRelationshipNamespace &&
        templateOfficeDocumentRelationships.has(
          templateAttribute(element, "Type") ?? ""
        ) &&
        templateAttribute(element, "TargetMode") === undefined &&
        (target === "word/document.xml" || target === "/word/document.xml")
      ) {
        officeDocumentRelationship = true;
      }
    },
  });

  let documentBody = false;
  let documentRoot = false;
  let documentRootSeen = false;
  parseTemplateXml(document, {
    open: (element) => {
      if (!documentRootSeen) {
        documentRootSeen = true;
        documentRoot =
          element.local === "document" &&
          templateWordNamespaces.has(element.uri);
      }
      if (element.local === "body" && templateWordNamespaces.has(element.uri)) {
        documentBody = true;
      }
    },
  });

  if (
    !contentTypesRoot ||
    !documentOverride ||
    !relationshipsRoot ||
    !officeDocumentRelationship ||
    !documentRoot ||
    !documentBody
  ) {
    fail(
      422,
      "invalid_template",
      "The DOCX package is missing required document parts"
    );
  }

  const xmlPaths = new Set([
    "[Content_Types].xml",
    "_rels/.rels",
    "word/document.xml",
  ]);
  for (const archivePath of Object.keys(archive)) {
    const lowercasePath = archivePath.toLowerCase();
    const extension = lowercasePath.slice(lowercasePath.lastIndexOf(".") + 1);
    const contentType =
      overrideContentTypes.get(`/${archivePath}`) ??
      defaultContentTypes.get(extension);
    if (
      lowercasePath.endsWith(".xml") ||
      lowercasePath.endsWith(".rels") ||
      lowercasePath.endsWith(".vml") ||
      isTemplateXmlContentType(contentType)
    ) {
      xmlPaths.add(archivePath);
    }
  }
  for (const archivePath of xmlPaths) {
    if (
      archivePath !== "[Content_Types].xml" &&
      archivePath !== "_rels/.rels" &&
      archivePath !== "word/document.xml"
    ) {
      parseTemplateXml(templateArchiveText(archive, archivePath));
    }
  }
  return xmlPaths;
}

export function safeTemplateArchive(bytes: Uint8Array): {
  archive: Record<string, Uint8Array>;
  xmlPaths: Set<string>;
} {
  const archive = readTemplateArchive(bytes, () => true);
  validateOfficeRelationships(archive);
  return {
    archive,
    xmlPaths: validateTemplatePackageArchive(archive),
  };
}

export function validateTemplatePackage(bytes: Uint8Array): void {
  safeTemplateArchive(bytes);
}

const templateControlPartPattern =
  /^word\/(?:document|endnotes|footnotes|footer\d+|header\d+)\.xml$/u;
