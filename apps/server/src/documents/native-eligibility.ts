import { prisma } from "@onlyoffice/db";
import { SaxesParser } from "saxes";

// oxlint-disable func-style prefer-destructuring no-await-in-loop no-use-before-define no-nested-ternary complexity no-shadow -- Route modules keep declaration order and sequential persistence invariants.
import { fail } from "../http/errors";
import { parseTemplateFields, templateControlAttribute } from "./fields";
import {
  templateArchiveText,
  templateMarkupCompatibilityNamespace,
  templateWordNamespaces,
  templateXmlElement,
} from "./package";
import { validateResponsePictureControls } from "./pictures";

export async function validateResponseDocument(
  publishedTemplateId: string,
  bytes: Uint8Array,
  enforceRequired: boolean
): Promise<void> {
  const fields = parseTemplateFields(bytes);
  const manifest = await prisma.fieldManifest.findUnique({
    include: { fields: { orderBy: { tag: "asc" } } },
    where: { publishedTemplateId },
  });
  if (!manifest || manifest.fields.length !== fields.length) {
    fail(
      422,
      "invalid_template",
      "The response document does not match the published manifest"
    );
  }
  const fieldsByTag = new Map(fields.map((field) => [field.tag, field]));
  for (const manifestField of manifest.fields) {
    const field = fieldsByTag.get(manifestField.tag);
    const fieldOptions = field?.options;
    const manifestOptions = manifestField.options;
    const optionsMatch =
      (fieldOptions === null && manifestOptions === null) ||
      (Array.isArray(fieldOptions) &&
        Array.isArray(manifestOptions) &&
        manifestOptions.length === fieldOptions.length &&
        fieldOptions.every((option, index) => {
          const manifestOption = manifestOptions[index];
          return (
            typeof manifestOption === "object" &&
            manifestOption !== null &&
            !Array.isArray(manifestOption) &&
            manifestOption.displayText === option.displayText &&
            manifestOption.value === option.value
          );
        }));
    if (
      !field ||
      field.type !== manifestField.type ||
      !optionsMatch ||
      field.pictureMaxBytes !== manifestField.pictureMaxBytes ||
      field.pictureMaxHeight !== manifestField.pictureMaxHeight ||
      field.pictureMaxWidth !== manifestField.pictureMaxWidth
    ) {
      fail(
        422,
        "invalid_template",
        "The response document does not match the published manifest"
      );
    }
  }
  validateResponsePictureControls(bytes, manifest.fields, enforceRequired);
}

export function templatePartsHaveNestedControls(
  archive: Record<string, Uint8Array>,
  controlParts: ReadonlySet<string>
): boolean {
  for (const archivePath of controlParts) {
    let controlDepth = 0;
    let alternateFallbackDepth = 0;
    let nested = false;
    const parser = new SaxesParser({ xmlns: true });
    parser.on("opentag", (tag) => {
      if (
        tag.local === "Fallback" &&
        tag.uri === templateMarkupCompatibilityNamespace
      ) {
        alternateFallbackDepth += 1;
        return;
      }
      if (alternateFallbackDepth > 0) {
        return;
      }
      if (tag.local === "sdt" && templateWordNamespaces.has(tag.uri)) {
        nested ||= controlDepth > 0;
        controlDepth += 1;
      }
    });
    parser.on("closetag", (tag) => {
      if (
        tag.local === "Fallback" &&
        tag.uri === templateMarkupCompatibilityNamespace
      ) {
        alternateFallbackDepth -= 1;
        return;
      }
      if (
        alternateFallbackDepth === 0 &&
        tag.local === "sdt" &&
        templateWordNamespaces.has(tag.uri) &&
        controlDepth > 0
      ) {
        controlDepth -= 1;
      }
    });
    parser.write(templateArchiveText(archive, archivePath)).close();
    if (nested) {
      return true;
    }
  }
  return false;
}

function hasUnsupportedNativeDateFormatTokens(format: string): boolean {
  const unquoted = format.replaceAll(/'[^']*'|"[^"]*"/gu, "");
  const unrecognized = unquoted.replaceAll(
    /dddd|ddd|yyyy|MMMM|MMM|yy|MM|dd|M|d/gu,
    ""
  );
  return /[A-Za-z]/u.test(unrecognized);
}

export function templatePartsHaveUnsupportedDateSettings(
  archive: Record<string, Uint8Array>,
  controlParts: ReadonlySet<string>
): boolean {
  for (const archivePath of controlParts) {
    let alternateFallbackDepth = 0;
    let inDateControl = false;
    let propertiesDepth = 0;
    let unsupportedDateSettings = false;
    const parser = new SaxesParser({ xmlns: true });
    parser.on("opentag", (tag) => {
      if (
        tag.local === "Fallback" &&
        tag.uri === templateMarkupCompatibilityNamespace
      ) {
        alternateFallbackDepth += 1;
        return;
      }
      if (alternateFallbackDepth > 0) {
        return;
      }
      if (tag.local === "sdtPr" && templateWordNamespaces.has(tag.uri)) {
        propertiesDepth = 1;
        inDateControl = false;
        return;
      }
      if (propertiesDepth === 0) {
        return;
      }
      if (
        propertiesDepth === 1 &&
        tag.local === "date" &&
        templateWordNamespaces.has(tag.uri)
      ) {
        inDateControl = true;
      } else if (
        inDateControl &&
        tag.local === "calendar" &&
        templateWordNamespaces.has(tag.uri)
      ) {
        const calendar = templateControlAttribute(
          templateXmlElement(tag),
          "val"
        );
        unsupportedDateSettings ||=
          calendar !== undefined && calendar !== "gregorian";
      } else if (
        inDateControl &&
        tag.local === "dateFormat" &&
        templateWordNamespaces.has(tag.uri)
      ) {
        const format = templateControlAttribute(templateXmlElement(tag), "val");
        unsupportedDateSettings ||=
          format !== undefined && hasUnsupportedNativeDateFormatTokens(format);
      } else if (
        inDateControl &&
        tag.local === "lid" &&
        templateWordNamespaces.has(tag.uri)
      ) {
        const language = templateControlAttribute(
          templateXmlElement(tag),
          "val"
        );
        unsupportedDateSettings ||=
          language !== undefined && !/^en(?:-|$)/iu.test(language);
      }
      propertiesDepth += 1;
    });
    parser.on("closetag", (tag) => {
      if (
        tag.local === "Fallback" &&
        tag.uri === templateMarkupCompatibilityNamespace
      ) {
        alternateFallbackDepth -= 1;
        return;
      }
      if (alternateFallbackDepth > 0) {
        return;
      }
      if (tag.local === "sdtPr" && templateWordNamespaces.has(tag.uri)) {
        propertiesDepth = 0;
        inDateControl = false;
      } else if (propertiesDepth > 0) {
        if (tag.local === "date" && templateWordNamespaces.has(tag.uri)) {
          inDateControl = false;
        }
        propertiesDepth -= 1;
      }
    });
    parser.write(templateArchiveText(archive, archivePath)).close();
    if (unsupportedDateSettings) {
      return true;
    }
  }
  return false;
}
