import path from "node:path";

import { FieldType } from "@onlyoffice/db";
import { strToU8, zipSync } from "fflate";
import { SaxesParser } from "saxes";

// oxlint-disable func-style prefer-destructuring no-await-in-loop no-use-before-define no-nested-ternary complexity no-shadow -- Route modules keep declaration order and sequential persistence invariants.
import { fail } from "../http/errors";
import { htmlEscape } from "../markup";
import type { NativePictureUpload } from "../responses/native-request";
import { templateControlAttribute } from "./fields";
import { nativeFlattenAlternateFieldsXml } from "./native-alternate";
import {
  parseTemplateXml,
  reachableTemplateControlParts,
  safeTemplateArchive,
  templateArchiveText,
  templateAttribute,
  templatePackageContentTypesNamespace,
  templatePackageRelationshipNamespace,
  templateRelationshipPartPath,
  templateWordNamespaces,
  templateWordStrictNamespace,
  templateXmlElement,
} from "./package";
import {
  invalidResponsePicture,
  responsePictureRelationshipId,
  responsePictureRelationships,
} from "./pictures";
import type {
  ResponsePictureDimensions,
  ResponsePictureManifestField,
} from "./pictures";

function appendTemplateXmlChild(
  xml: string,
  rootLocalName: string,
  rootNamespace: string,
  child: string
): string {
  const parser = new SaxesParser({ position: true, xmlns: true });
  const rootNames: string[] = [];
  const rootCloseStarts: number[] = [];
  let depth = 0;
  parser.on("doctype", () => {
    fail(422, "invalid_template", "DOCX XML document types are not allowed");
  });
  parser.on("opentag", (tag) => {
    if (depth === 0) {
      if (
        rootNames.length > 0 ||
        tag.local !== rootLocalName ||
        tag.uri !== rootNamespace
      ) {
        fail(422, "invalid_template", "The DOCX package XML root is invalid");
      }
      rootNames.push(tag.name);
    }
    if (!tag.isSelfClosing) {
      depth += 1;
    }
  });
  parser.on("closetag", (tag) => {
    if (
      depth === 1 &&
      tag.local === rootLocalName &&
      tag.uri === rootNamespace
    ) {
      rootCloseStarts.push(xml.lastIndexOf("</", parser.position - 1));
    }
    if (!tag.isSelfClosing) {
      depth -= 1;
    }
  });
  try {
    parser.write(xml).close();
  } catch {
    fail(422, "invalid_template", "The DOCX package contains invalid XML");
  }
  const rootName = rootNames[0];
  const rootCloseStart = rootCloseStarts[0];
  if (!rootName || rootCloseStart === undefined) {
    fail(422, "invalid_template", "The DOCX package XML root is invalid");
  }
  const prefix = rootName.includes(":")
    ? `${rootName.slice(0, rootName.indexOf(":"))}:`
    : "";
  return `${xml.slice(0, rootCloseStart)}${child.replaceAll(
    "{prefix}",
    prefix
  )}${xml.slice(rootCloseStart)}`;
}

function setNativePictureContentType(
  archive: Record<string, Uint8Array>,
  mediaPath: string,
  contentType: string
): void {
  const xml = templateArchiveText(archive, "[Content_Types].xml");
  const parser = new SaxesParser({ position: true, xmlns: true });
  const overrides: {
    contentType: string | undefined;
    end: number;
    start: number;
  }[] = [];
  parser.on("opentag", (tag) => {
    if (
      tag.local !== "Override" ||
      tag.uri !== templatePackageContentTypesNamespace ||
      templateAttribute(templateXmlElement(tag), "PartName") !== `/${mediaPath}`
    ) {
      return;
    }
    if (overrides.length > 0) {
      fail(422, "invalid_template", "The DOCX content types are invalid");
    }
    overrides.push({
      contentType: templateAttribute(templateXmlElement(tag), "ContentType"),
      end: parser.position,
      start: xml.lastIndexOf("<", parser.position - 1),
    });
  });
  parser.on("doctype", () => {
    fail(422, "invalid_template", "DOCX XML document types are not allowed");
  });
  try {
    parser.write(xml).close();
  } catch {
    fail(422, "invalid_template", "The DOCX package contains invalid XML");
  }
  const override = overrides[0];
  if (override) {
    if (override.contentType === contentType) {
      return;
    }
    const start = override.start;
    const end = override.end;
    const openingTag = xml.slice(start, end);
    const updatedTag = openingTag.replace(
      /(?<prefix>\sContentType\s*=\s*)(?<quote>["'])[^"']*\k<quote>/iu,
      `$<prefix>"${contentType}"`
    );
    if (updatedTag === openingTag) {
      fail(422, "invalid_template", "The DOCX content types are invalid");
    }
    archive["[Content_Types].xml"] = strToU8(
      `${xml.slice(0, start)}${updatedTag}${xml.slice(end)}`.replace(
        /encoding=(?<quote>["'])UTF-16(?:LE|BE)?\k<quote>/iu,
        'encoding="UTF-8"'
      )
    );
    return;
  }
  const updatedXml = appendTemplateXmlChild(
    xml,
    "Types",
    templatePackageContentTypesNamespace,
    `<{prefix}Override PartName="/${mediaPath}" ContentType="${contentType}"/>`
  );
  archive["[Content_Types].xml"] = strToU8(
    updatedXml.replace(
      /encoding=(?<quote>["'])UTF-16(?:LE|BE)?\k<quote>/iu,
      'encoding="UTF-8"'
    )
  );
}

function pictureRelationshipReferenceCount(
  xml: string,
  relationshipId: string
): number {
  let count = 0;
  parseTemplateXml(xml, {
    open: (element) => {
      if (responsePictureRelationshipId(element) === relationshipId) {
        count += 1;
      }
    },
  });
  return count;
}

function addNativePicturePackageParts(
  archive: Record<string, Uint8Array>,
  sourcePath: string,
  upload: NativePictureUpload,
  existingRelationshipIds: readonly string[]
): string {
  const existingRelationships = responsePictureRelationships(
    archive,
    sourcePath
  );
  const existingRelationshipId =
    existingRelationshipIds.length === 1 ? existingRelationshipIds[0] : null;
  const existingMediaPath = existingRelationshipId
    ? existingRelationships.get(existingRelationshipId)
    : null;
  if (
    existingRelationshipId &&
    existingMediaPath?.startsWith("word/media/picture-") &&
    archive[existingMediaPath] &&
    [...existingRelationships.values()].filter(
      (target) => target === existingMediaPath
    ).length === 1 &&
    pictureRelationshipReferenceCount(
      templateArchiveText(archive, sourcePath),
      existingRelationshipId
    ) === 1
  ) {
    archive[existingMediaPath] = upload.bytes;
    setNativePictureContentType(
      archive,
      existingMediaPath,
      upload.dimensions.format === "jpeg" ? "image/jpeg" : "image/png"
    );
    return existingRelationshipId;
  }

  const extension = upload.dimensions.format === "jpeg" ? "jpeg" : "png";
  const mediaName = `picture-${crypto.randomUUID()}.${extension}`;
  const mediaPath = `word/media/${mediaName}`;
  const relationshipPath = templateRelationshipPartPath(sourcePath);
  const relationshipId = (() => {
    let suffix = 1;
    while (existingRelationships.has(`rId${suffix}`)) {
      suffix += 1;
    }
    return `rId${suffix}`;
  })();
  const target = path.posix.relative(path.posix.dirname(sourcePath), mediaPath);
  const relationship = `<{prefix}Relationship Id="${relationshipId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="${target}"/>`;
  const relationshipsXml = archive[relationshipPath]
    ? appendTemplateXmlChild(
        templateArchiveText(archive, relationshipPath),
        "Relationships",
        templatePackageRelationshipNamespace,
        relationship
      )
    : `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="${templatePackageRelationshipNamespace}">${relationship.replaceAll("{prefix}", "")}</Relationships>`;
  archive[relationshipPath] = strToU8(
    relationshipsXml.replace(
      /encoding=(?<quote>["'])UTF-16(?:LE|BE)?\k<quote>/iu,
      'encoding="UTF-8"'
    )
  );
  setNativePictureContentType(
    archive,
    mediaPath,
    upload.dimensions.format === "jpeg" ? "image/jpeg" : "image/png"
  );
  archive[mediaPath] = upload.bytes;
  return relationshipId;
}

function nativePictureDrawing(
  tag: string,
  block: boolean,
  wordPrefix: string,
  contentNamespace: string,
  relationshipId: string,
  drawingId: number,
  dimensions: ResponsePictureDimensions
): string {
  const strict = contentNamespace === templateWordStrictNamespace;
  const drawingNamespace = strict
    ? "http://purl.oclc.org/ooxml/drawingml/main"
    : "http://schemas.openxmlformats.org/drawingml/2006/main";
  const wordDrawingNamespace = strict
    ? "http://purl.oclc.org/ooxml/drawingml/wordprocessingDrawing"
    : "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing";
  const pictureNamespace = strict
    ? "http://purl.oclc.org/ooxml/drawingml/picture"
    : "http://schemas.openxmlformats.org/drawingml/2006/picture";
  const relationshipNamespace = strict
    ? "http://purl.oclc.org/ooxml/officeDocument/relationships"
    : "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
  const cx = dimensions.width * 9525;
  const cy = dimensions.height * 9525;
  const drawing =
    `<${wordPrefix}drawing xmlns:a="${drawingNamespace}" xmlns:pic="${pictureNamespace}" xmlns:wp="${wordDrawingNamespace}" xmlns:r="${relationshipNamespace}">` +
    `<wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="${cx}" cy="${cy}"/>` +
    `<wp:docPr id="${drawingId}" name="Picture ${htmlEscape(tag)}"/>` +
    `<wp:cNvGraphicFramePr><a:graphicFrameLocks noChangeAspect="1"/></wp:cNvGraphicFramePr>` +
    `<a:graphic><a:graphicData uri="${pictureNamespace}"><pic:pic>` +
    `<pic:nvPicPr><pic:cNvPr id="${drawingId}" name="Picture ${htmlEscape(tag)}"/><pic:cNvPicPr><a:picLocks noChangeAspect="1"/></pic:cNvPicPr></pic:nvPicPr>` +
    `<pic:blipFill><a:blip r:embed="${relationshipId}"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>` +
    `<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr>` +
    `</pic:pic></a:graphicData></a:graphic></wp:inline></${wordPrefix}drawing>`;
  const run = `<${wordPrefix}r>${drawing}</${wordPrefix}r>`;
  return block ? `<${wordPrefix}p>${run}</${wordPrefix}p>` : run;
}

export function overlayNativeResponsePictures(
  documentBytes: Uint8Array,
  manifestFields: readonly ResponsePictureManifestField[],
  pictures: ReadonlyMap<string, NativePictureUpload>
): Uint8Array {
  if (pictures.size === 0) {
    return documentBytes;
  }
  const pictureFields = manifestFields.filter(
    ({ type }) => type === FieldType.picture
  );
  const fieldsByTag = new Map(pictureFields.map((field) => [field.tag, field]));
  const selectedTags = new Set(pictures.keys());
  for (const tag of selectedTags) {
    if (!fieldsByTag.has(tag)) {
      invalidResponsePicture(tag, "field is not a published Picture");
    }
  }
  const { archive, xmlPaths } = safeTemplateArchive(documentBytes);
  const controlParts = reachableTemplateControlParts(archive, xmlPaths);
  const fieldTags = new Set(manifestFields.map(({ tag }) => tag));
  const drawingIds = new Set<number>();
  for (const archivePath of xmlPaths) {
    parseTemplateXml(templateArchiveText(archive, archivePath), {
      open: (element) => {
        if (element.local !== "docPr") {
          return;
        }
        const id = Number(templateAttribute(element, "id"));
        if (Number.isSafeInteger(id) && id > 0) {
          drawingIds.add(id);
        }
      },
    });
  }
  const placedTags = new Set<string>();
  for (const archivePath of controlParts) {
    const originalXml = templateArchiveText(archive, archivePath);
    const flattened = nativeFlattenAlternateFieldsXml(originalXml, fieldTags);
    if (!flattened.supported) {
      fail(422, "invalid_template", "Native picture control is unsupported");
    }
    const xml = flattened.xml;
    const controls: {
      block: boolean;
      contentEnd: number | null;
      contentName: string | null;
      contentNamespace: string;
      contentPrefix: string;
      contentSelfClosing: boolean;
      contentStart: number | null;
      contentTokenStart: number | null;
      inPropertiesDepth: number;
      picture: boolean;
      relationshipIds: string[];
      showingPlaceholderRanges: { end: number; start: number }[];
      showingPlaceholderStarts: number[];
      tag: string | null;
    }[] = [];
    const elements: { local: string; uri: string }[] = [];
    const patches: { end: number; replacement: string; start: number }[] = [];
    const parser = new SaxesParser({ position: true, xmlns: true });
    parser.on("doctype", () => {
      fail(422, "invalid_template", "DOCX XML document types are not allowed");
    });
    parser.on("opentag", (tag) => {
      if (tag.local === "sdt" && templateWordNamespaces.has(tag.uri)) {
        controls.push({
          block: !elements.some(
            (element) =>
              element.local === "p" && templateWordNamespaces.has(element.uri)
          ),
          contentEnd: null,
          contentName: null,
          contentNamespace: "",
          contentPrefix: "",
          contentSelfClosing: false,
          contentStart: null,
          contentTokenStart: null,
          inPropertiesDepth: 0,
          picture: false,
          relationshipIds: [],
          showingPlaceholderRanges: [],
          showingPlaceholderStarts: [],
          tag: null,
        });
      } else {
        const frame = controls.at(-1);
        if (frame) {
          if (tag.local === "sdtPr" && templateWordNamespaces.has(tag.uri)) {
            frame.inPropertiesDepth = 1;
          } else if (frame.inPropertiesDepth > 0) {
            if (
              frame.inPropertiesDepth === 1 &&
              tag.local === "tag" &&
              templateWordNamespaces.has(tag.uri)
            ) {
              frame.tag =
                templateControlAttribute(templateXmlElement(tag), "val") ??
                null;
            }
            if (
              frame.inPropertiesDepth === 1 &&
              tag.local === "picture" &&
              templateWordNamespaces.has(tag.uri)
            ) {
              frame.picture = true;
            }
            if (
              frame.inPropertiesDepth === 1 &&
              tag.local === "showingPlcHdr" &&
              templateWordNamespaces.has(tag.uri)
            ) {
              frame.showingPlaceholderStarts.push(
                xml.lastIndexOf("<", parser.position - 1)
              );
            }
            frame.inPropertiesDepth += 1;
          } else if (
            tag.local === "sdtContent" &&
            templateWordNamespaces.has(tag.uri)
          ) {
            const start = xml.lastIndexOf("<", parser.position - 1);
            frame.contentName = tag.name;
            frame.contentNamespace = tag.uri;
            frame.contentPrefix = tag.name.includes(":")
              ? tag.name.slice(0, tag.name.indexOf(":") + 1)
              : "";
            frame.contentSelfClosing = tag.isSelfClosing;
            frame.contentStart = tag.isSelfClosing ? null : parser.position;
            frame.contentTokenStart = tag.isSelfClosing ? start : null;
          } else if (frame.picture) {
            const relationshipId = responsePictureRelationshipId(
              templateXmlElement(tag)
            );
            if (relationshipId) {
              frame.relationshipIds.push(relationshipId);
            }
          }
        }
      }
      elements.push({ local: tag.local, uri: tag.uri });
    });
    parser.on("closetag", (tag) => {
      const frame = controls.at(-1);
      if (frame) {
        if (tag.local === "sdtPr" && templateWordNamespaces.has(tag.uri)) {
          frame.inPropertiesDepth = 0;
        } else if (
          frame.inPropertiesDepth === 2 &&
          tag.local === "showingPlcHdr" &&
          templateWordNamespaces.has(tag.uri)
        ) {
          const start = frame.showingPlaceholderStarts.pop();
          if (start !== undefined) {
            frame.showingPlaceholderRanges.push({
              end: parser.position,
              start,
            });
          }
          frame.inPropertiesDepth -= 1;
        } else if (frame.inPropertiesDepth > 0) {
          frame.inPropertiesDepth -= 1;
        } else if (
          tag.local === "sdtContent" &&
          templateWordNamespaces.has(tag.uri)
        ) {
          frame.contentEnd = frame.contentSelfClosing
            ? parser.position
            : xml.lastIndexOf("</", parser.position - 1);
          const fieldTag = frame.tag?.trim();
          const upload = fieldTag ? pictures.get(fieldTag) : undefined;
          if (frame.picture && fieldTag && upload) {
            if (placedTags.has(fieldTag)) {
              invalidResponsePicture(fieldTag, "content control is duplicated");
            }
            placedTags.add(fieldTag);
            for (const range of frame.showingPlaceholderRanges) {
              patches.push({ ...range, replacement: "" });
            }
            let drawingId = 1;
            while (drawingIds.has(drawingId)) {
              drawingId += 1;
            }
            drawingIds.add(drawingId);
            const relationshipId = addNativePicturePackageParts(
              archive,
              archivePath,
              upload,
              frame.relationshipIds
            );
            if (!frame.contentName) {
              invalidResponsePicture(fieldTag, "content control is incomplete");
            }
            const content = nativePictureDrawing(
              fieldTag,
              frame.block,
              frame.contentPrefix,
              frame.contentNamespace,
              relationshipId,
              drawingId,
              upload.dimensions
            );
            if (
              frame.contentSelfClosing &&
              frame.contentTokenStart !== null &&
              frame.contentEnd !== null
            ) {
              const openingTag = xml.slice(
                frame.contentTokenStart,
                frame.contentEnd
              );
              patches.push({
                end: frame.contentEnd,
                replacement: `${openingTag.replace(/\/\s*>$/u, ">")}${content}</${frame.contentName}>`,
                start: frame.contentTokenStart,
              });
            } else if (
              frame.contentStart !== null &&
              frame.contentEnd !== null
            ) {
              patches.push({
                end: frame.contentEnd,
                replacement: content,
                start: frame.contentStart,
              });
            } else {
              invalidResponsePicture(fieldTag, "content control is incomplete");
            }
          }
        } else if (tag.local === "sdt" && templateWordNamespaces.has(tag.uri)) {
          controls.pop();
        }
      }
      elements.pop();
    });
    try {
      parser.write(xml).close();
    } catch {
      fail(422, "invalid_template", "The DOCX package contains invalid XML");
    }
    let updatedXml = xml;
    for (const patch of patches.toSorted(
      (left, right) => right.start - left.start
    )) {
      updatedXml =
        updatedXml.slice(0, patch.start) +
        patch.replacement +
        updatedXml.slice(patch.end);
    }
    if (updatedXml !== originalXml) {
      archive[archivePath] = strToU8(
        updatedXml.replace(
          /encoding=(?<quote>["'])UTF-16(?:LE|BE)?\k<quote>/iu,
          'encoding="UTF-8"'
        )
      );
    }
  }
  for (const tag of selectedTags) {
    if (!placedTags.has(tag)) {
      invalidResponsePicture(tag, "content control is missing");
    }
  }
  return zipSync(archive);
}
