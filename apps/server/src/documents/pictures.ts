import { FieldType } from "@onlyoffice/db";

// oxlint-disable func-style prefer-destructuring no-await-in-loop no-use-before-define no-nested-ternary complexity no-shadow -- Route modules keep declaration order and sequential persistence invariants.
import { fail } from "../http/errors";
import { templateControlAttribute } from "./fields";
import {
  parseTemplateXml,
  reachableTemplateControlParts,
  resolveTemplateRelationshipTarget,
  safeTemplateArchive,
  templateArchiveText,
  templateAttribute,
  templateDrawingMlNamespaces,
  templateMarkupCompatibilityNamespace,
  templateOfficeRelationshipNamespaces,
  templatePackageRelationshipNamespace,
  templateRelationshipPartPath,
  templateVmlNamespace,
  templateWordNamespaces,
} from "./package";
import type { TemplateXmlElement } from "./package";

const responsePictureJpegSofMarkers = new Set([
  0xc0, 0xc1, 0xc2, 0xc3, 0xc5, 0xc6, 0xc7, 0xc9, 0xca, 0xcb, 0xcd, 0xce, 0xcf,
]);

const responsePictureJpegStandaloneMarkers = new Set([
  0x01, 0xd0, 0xd1, 0xd2, 0xd3, 0xd4, 0xd5, 0xd6, 0xd7, 0xd8,
]);

const responsePicturePngSignature = [
  0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a,
] as const;

export interface ResponsePictureManifestField {
  pictureMaxBytes: number | null;
  pictureMaxHeight: number | null;
  pictureMaxWidth: number | null;
  required: boolean;
  tag: string;
  type: FieldType;
}

interface ResponsePictureControlFrame {
  inPropertiesDepth: number;
  picture: boolean;
  relationshipIds: string[];
  showingPlaceholder: boolean;
  tag: string | null;
}

interface ResponsePictureControl {
  relationshipIds: string[];
  tag: string;
}

export interface ResponsePictureDimensions {
  format: "jpeg" | "png";
  height: number;
  width: number;
}

export function invalidResponsePicture(tag: string, message: string): never {
  fail(422, "invalid_template", `Invalid picture field ${tag}: ${message}`);
}

export function responsePictureRelationshipId(
  element: TemplateXmlElement
): string | undefined {
  const local =
    element.local === "blip" && templateDrawingMlNamespaces.has(element.uri)
      ? "embed"
      : element.local === "imagedata" && element.uri === templateVmlNamespace
        ? "id"
        : null;
  if (!local) {
    return undefined;
  }
  return element.attributes
    .find(
      (attribute) =>
        attribute.local === local &&
        templateOfficeRelationshipNamespaces.has(attribute.uri)
    )
    ?.value.trim();
}

function responsePictureUint16(bytes: Uint8Array, offset: number): number {
  return (bytes[offset] ?? 0) * 0x1_00 + (bytes[offset + 1] ?? 0);
}

function responsePictureUint32(bytes: Uint8Array, offset: number): number {
  return (
    (bytes[offset] ?? 0) * 0x1_00_00_00 +
    (bytes[offset + 1] ?? 0) * 0x1_00_00 +
    (bytes[offset + 2] ?? 0) * 0x1_00 +
    (bytes[offset + 3] ?? 0)
  );
}

function responsePictureJpegDimensions(
  bytes: Uint8Array
): ResponsePictureDimensions | null {
  if (bytes.byteLength < 2 || bytes[0] !== 0xff || bytes[1] !== 0xd8) {
    return null;
  }
  let offset = 2;
  while (offset < bytes.byteLength) {
    if (bytes[offset] !== 0xff) {
      return null;
    }
    while (bytes[offset] === 0xff) {
      offset += 1;
    }
    const marker = bytes[offset];
    if (marker === undefined || marker === 0) {
      return null;
    }
    offset += 1;
    if (marker === 0xd9) {
      return null;
    }
    if (responsePictureJpegStandaloneMarkers.has(marker)) {
      continue;
    }
    if (offset + 2 > bytes.byteLength) {
      return null;
    }
    const segmentLength = responsePictureUint16(bytes, offset);
    if (segmentLength < 2 || offset + segmentLength > bytes.byteLength) {
      return null;
    }
    if (responsePictureJpegSofMarkers.has(marker)) {
      if (segmentLength < 7) {
        return null;
      }
      return {
        format: "jpeg",
        height: responsePictureUint16(bytes, offset + 3),
        width: responsePictureUint16(bytes, offset + 5),
      };
    }
    if (marker === 0xda) {
      return null;
    }
    offset += segmentLength;
  }
  return null;
}

export function responsePictureImageDimensions(
  bytes: Uint8Array
): ResponsePictureDimensions | null {
  const hasPngSignature =
    bytes.byteLength >= responsePicturePngSignature.length &&
    responsePicturePngSignature.every((byte, index) => bytes[index] === byte);
  if (hasPngSignature) {
    if (
      bytes.byteLength < 24 ||
      responsePictureUint32(bytes, 8) !== 13 ||
      bytes[12] !== 0x49 ||
      bytes[13] !== 0x48 ||
      bytes[14] !== 0x44 ||
      bytes[15] !== 0x52
    ) {
      return null;
    }
    return {
      format: "png",
      height: responsePictureUint32(bytes, 20),
      width: responsePictureUint32(bytes, 16),
    };
  }
  return responsePictureJpegDimensions(bytes);
}

export function validateResponsePictureMediaBytes(
  tag: string,
  bytes: Uint8Array,
  field: ResponsePictureManifestField,
  dimensions = responsePictureImageDimensions(bytes)
): void {
  if (
    field.pictureMaxBytes !== null &&
    bytes.byteLength > field.pictureMaxBytes
  ) {
    invalidResponsePicture(tag, "image bytes exceed the published limit");
  }
  if (!dimensions) {
    invalidResponsePicture(tag, "image must be a valid JPEG or PNG");
  }
  if (dimensions.width <= 0 || dimensions.height <= 0) {
    invalidResponsePicture(tag, "image dimensions must be positive");
  }
  if (
    field.pictureMaxWidth !== null &&
    dimensions.width > field.pictureMaxWidth
  ) {
    invalidResponsePicture(tag, "image width exceeds the published limit");
  }
  if (
    field.pictureMaxHeight !== null &&
    dimensions.height > field.pictureMaxHeight
  ) {
    invalidResponsePicture(tag, "image height exceeds the published limit");
  }
}

export function responsePictureRelationships(
  archive: Record<string, Uint8Array>,
  sourcePath: string
): Map<string, string | null> {
  const relationships = new Map<string, string | null>();
  const relationshipPath = templateRelationshipPartPath(sourcePath);
  if (!archive[relationshipPath]) {
    return relationships;
  }
  parseTemplateXml(templateArchiveText(archive, relationshipPath), {
    open: (element) => {
      if (
        element.local !== "Relationship" ||
        element.uri !== templatePackageRelationshipNamespace
      ) {
        return;
      }
      const id = templateAttribute(element, "Id")?.trim();
      if (!id) {
        return;
      }
      const target = resolveTemplateRelationshipTarget(
        sourcePath,
        templateAttribute(element, "Target")
      );
      relationships.set(id, relationships.has(id) ? null : target);
    },
  });
  return relationships;
}

export function validateResponsePictureControls(
  bytes: Uint8Array,
  manifestFields: readonly ResponsePictureManifestField[],
  enforceRequired: boolean
): Set<string> {
  const pictureFields = manifestFields.filter(
    (field) => field.type === FieldType.picture
  );
  const presentTags = new Set<string>();
  if (pictureFields.length === 0) {
    return presentTags;
  }
  const { archive, xmlPaths } = safeTemplateArchive(bytes);
  const fieldsByTag = new Map(pictureFields.map((field) => [field.tag, field]));
  const seenTags = new Set<string>();
  const relationshipsBySource = new Map<string, Map<string, string | null>>();
  for (const archivePath of reachableTemplateControlParts(archive, xmlPaths)) {
    const controls: ResponsePictureControlFrame[] = [];
    let alternateFallbackDepth = 0;
    const pictureControls: ResponsePictureControl[] = [];
    parseTemplateXml(templateArchiveText(archive, archivePath), {
      close: (element) => {
        if (
          element.local === "Fallback" &&
          element.uri === templateMarkupCompatibilityNamespace
        ) {
          alternateFallbackDepth -= 1;
          return;
        }
        if (alternateFallbackDepth > 0) {
          return;
        }
        const frame = controls.at(-1);
        if (!frame) {
          return;
        }
        if (
          element.local === "sdtPr" &&
          templateWordNamespaces.has(element.uri)
        ) {
          frame.inPropertiesDepth = 0;
          return;
        }
        if (frame.inPropertiesDepth > 0) {
          frame.inPropertiesDepth -= 1;
          return;
        }
        if (
          element.local === "sdt" &&
          templateWordNamespaces.has(element.uri)
        ) {
          controls.pop();
          const tag = frame.tag?.trim();
          if (frame.picture && tag) {
            pictureControls.push({
              relationshipIds: frame.showingPlaceholder
                ? []
                : frame.relationshipIds,
              tag,
            });
          }
        }
      },
      open: (element) => {
        if (
          element.local === "Fallback" &&
          element.uri === templateMarkupCompatibilityNamespace
        ) {
          alternateFallbackDepth += 1;
          return;
        }
        if (alternateFallbackDepth > 0) {
          return;
        }
        if (
          element.local === "sdt" &&
          templateWordNamespaces.has(element.uri)
        ) {
          controls.push({
            inPropertiesDepth: 0,
            picture: false,
            relationshipIds: [],
            showingPlaceholder: false,
            tag: null,
          });
          return;
        }
        const frame = controls.at(-1);
        if (!frame) {
          return;
        }
        if (
          element.local === "sdtPr" &&
          templateWordNamespaces.has(element.uri)
        ) {
          frame.inPropertiesDepth = 1;
          return;
        }
        if (frame.inPropertiesDepth > 0) {
          if (
            frame.inPropertiesDepth === 1 &&
            element.local === "tag" &&
            templateWordNamespaces.has(element.uri)
          ) {
            frame.tag = templateControlAttribute(element, "val") ?? null;
          }
          if (
            frame.inPropertiesDepth === 1 &&
            element.local === "showingPlcHdr" &&
            templateWordNamespaces.has(element.uri)
          ) {
            const value = templateControlAttribute(element, "val")
              ?.trim()
              .toLowerCase();
            frame.showingPlaceholder =
              value === undefined || !["0", "false", "off"].includes(value);
          }
          if (
            frame.inPropertiesDepth === 1 &&
            element.local === "picture" &&
            templateWordNamespaces.has(element.uri)
          ) {
            frame.picture = true;
          }
          frame.inPropertiesDepth += 1;
          return;
        }
        const relationshipId = responsePictureRelationshipId(element);
        if (relationshipId === undefined) {
          return;
        }
        for (const activeFrame of controls) {
          if (activeFrame.picture && activeFrame.inPropertiesDepth === 0) {
            activeFrame.relationshipIds.push(relationshipId);
          }
        }
      },
    });
    for (const control of pictureControls) {
      const field = fieldsByTag.get(control.tag);
      if (!field) {
        continue;
      }
      if (seenTags.has(control.tag)) {
        invalidResponsePicture(control.tag, "content control is duplicated");
      }
      seenTags.add(control.tag);
      if (control.relationshipIds.length === 0) {
        if (field.required && enforceRequired) {
          invalidResponsePicture(control.tag, "a required image is missing");
        }
        continue;
      }
      if (control.relationshipIds.length > 1) {
        invalidResponsePicture(
          control.tag,
          "content control references multiple images"
        );
      }
      const relationshipId = control.relationshipIds[0];
      const relationships =
        relationshipsBySource.get(archivePath) ??
        responsePictureRelationships(archive, archivePath);
      relationshipsBySource.set(archivePath, relationships);
      const mediaPath = relationshipId
        ? relationships.get(relationshipId)
        : undefined;
      if (!mediaPath) {
        invalidResponsePicture(
          control.tag,
          "image relationship or target media is missing"
        );
      }
      const media = archive[mediaPath];
      if (!media) {
        invalidResponsePicture(control.tag, "target media is missing");
      }
      validateResponsePictureMediaBytes(control.tag, media, field);
      presentTags.add(control.tag);
    }
  }
  for (const field of pictureFields) {
    if (!seenTags.has(field.tag)) {
      invalidResponsePicture(field.tag, "content control is missing");
    }
  }
  return presentTags;
}

export function responsePicturePresence(
  bytes: Uint8Array,
  pictureFields: readonly ResponsePictureManifestField[]
): Record<string, boolean> {
  const presentTags = validateResponsePictureControls(
    bytes,
    pictureFields,
    false
  );
  return Object.fromEntries(
    pictureFields.map(({ tag }) => [tag, presentTags.has(tag)])
  );
}
