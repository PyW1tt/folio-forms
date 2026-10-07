import { FieldType } from "@onlyoffice/db";
import { strToU8, zipSync } from "fflate";
import { SaxesParser } from "saxes";

// oxlint-disable func-style prefer-destructuring no-await-in-loop no-use-before-define no-nested-ternary complexity no-shadow -- Route modules keep declaration order and sequential persistence invariants.
import { fail } from "../http/errors";
import { htmlEscape } from "../markup";
import type { JsonRecord } from "../model-types";
import { nativeFlattenAlternateFieldsXml } from "./native-alternate";
import { nativeFieldControlXml } from "./native-controls";
import { templatePartsHaveNestedControls } from "./native-eligibility";
import {
  reachableTemplateControlParts,
  safeTemplateArchive,
  templateArchiveText,
  templateCheckboxNamespace,
  templateWordNamespaces,
  templateXmlnsNamespace,
} from "./package";

export function overlayNativeResponseDocument(
  callbackBytes: Uint8Array,
  fields: readonly { options: unknown; tag: string; type: FieldType }[],
  data: JsonRecord,
  selectedTags: ReadonlySet<string>
): Uint8Array {
  const { archive, xmlPaths } = safeTemplateArchive(callbackBytes);
  const controlParts = reachableTemplateControlParts(archive, xmlPaths);
  if (templatePartsHaveNestedControls(archive, controlParts)) {
    fail(422, "invalid_template", "Native fields cannot be nested");
  }
  if (controlParts.size === 0) {
    fail(422, "invalid_template", "The document has no fields");
  }
  const fieldsByTag = new Map(
    fields
      .filter(
        ({ tag, type }) => type !== FieldType.picture && selectedTags.has(tag)
      )
      .map(({ options, tag, type }) => [tag, { options, type }])
  );
  const fieldTags = new Set(fields.map(({ tag }) => tag));
  for (const archivePath of controlParts) {
    const originalXml = templateArchiveText(archive, archivePath);
    const flattened = nativeFlattenAlternateFieldsXml(
      nativeFieldControlXml(originalXml, data, fieldsByTag, selectedTags),
      fieldTags
    );
    if (!flattened.supported) {
      fail(
        422,
        "invalid_template",
        "AlternateContent field branches are unsupported"
      );
    }
    const xml = flattened.xml;
    if (xml !== originalXml) {
      archive[archivePath] = strToU8(
        xml.replace(
          /encoding=(?<quote>["'])UTF-16(?:LE|BE)?\k<quote>/iu,
          'encoding="UTF-8"'
        )
      );
    }
  }
  return zipSync(archive);
}

export function nativePdfScalarControlXml(xml: string): string {
  interface ScalarControlFrame {
    contentCloseStart: number | null;
    contentDepth: number | null;
    contentHasContent: boolean;
    contentIsBlock: boolean;
    contentOpenEnd: number | null;
    contentSelfClosing: boolean;
    namespaceAttributeNames: ReadonlySet<string>;
    namespaceDeclarations: { name: string; value: string }[];
    parentContentControls: ScalarControlFrame[] | null;
    parentIsBlockContainer: boolean;
    scalarControl: boolean;
    start: number;
    startEnd: number;
    wordNamespace: string;
    wordPrefix: string;
  }
  const controls: ScalarControlFrame[] = [];
  const elementStack: { local: string; uri: string }[] = [];
  const propertyControls: ScalarControlFrame[] = [];
  const patches: { end: number; replacement: string; start: number }[] = [];
  const parser = new SaxesParser({ position: true, xmlns: true });
  const preserveNamespaces = (
    start: number,
    end: number,
    declarations: readonly { name: string; value: string }[],
    existingAttributes: ReadonlySet<string>
  ) => {
    if (declarations.length === 0) {
      return;
    }
    const original = xml.slice(start, end);
    let replacement = original;
    for (const declaration of declarations) {
      if (existingAttributes.has(declaration.name)) {
        continue;
      }
      const insertionIndex = replacement.endsWith("/>")
        ? replacement.length - 2
        : replacement.length - 1;
      replacement = `${replacement.slice(0, insertionIndex)} ${declaration.name}="${htmlEscape(declaration.value)}"${replacement.slice(insertionIndex)}`;
    }
    if (replacement !== original) {
      patches.push({ end, replacement, start });
    }
  };
  const activeScalarNamespaceDeclarations = () => {
    const declarations = new Map<string, { name: string; value: string }>();
    for (const control of controls) {
      if (!control.scalarControl || control.contentDepth === null) {
        continue;
      }
      for (const declaration of control.namespaceDeclarations) {
        declarations.set(declaration.name, declaration);
      }
    }
    return [...declarations.values()];
  };
  parser.on("doctype", () => {
    fail(422, "invalid_template", "DOCX XML document types are not allowed");
  });
  parser.on("opentag", (tag) => {
    const start = xml.lastIndexOf("<", parser.position - 1);
    let hasContentControl = false;
    let parentContentControls: ScalarControlFrame[] | null = null;
    for (const control of controls) {
      if (control.contentDepth === null) {
        continue;
      }
      if (control.contentDepth === 0) {
        hasContentControl = true;
        control.contentHasContent = true;
        if (tag.local === "sdt" && templateWordNamespaces.has(tag.uri)) {
          (parentContentControls ??= []).push(control);
        }
      }
      if (
        templateWordNamespaces.has(tag.uri) &&
        (tag.local === "p" ||
          tag.local === "tbl" ||
          (control.contentDepth === 0 &&
            (tag.local === "altChunk" ||
              tag.local === "customXml" ||
              tag.local === "oMathPara")))
      ) {
        control.contentIsBlock = true;
      }
    }
    if (hasContentControl && tag.local !== "sdt") {
      preserveNamespaces(
        start,
        parser.position,
        activeScalarNamespaceDeclarations(),
        new Set(Object.keys(tag.attributes))
      );
    }
    for (const control of controls) {
      if (control.contentDepth !== null) {
        control.contentDepth += 1;
      }
    }
    if (tag.local === "sdt" && templateWordNamespaces.has(tag.uri)) {
      const parent = elementStack.at(-1);
      controls.push({
        contentCloseStart: null,
        contentDepth: null,
        contentHasContent: false,
        contentIsBlock: false,
        contentOpenEnd: null,
        contentSelfClosing: false,
        namespaceAttributeNames: new Set(Object.keys(tag.attributes)),
        namespaceDeclarations: Object.values(tag.attributes)
          .filter((attribute) => attribute.uri === templateXmlnsNamespace)
          .map((attribute) => ({
            name: attribute.name,
            value: attribute.value,
          })),
        parentContentControls,
        parentIsBlockContainer:
          parent !== undefined &&
          templateWordNamespaces.has(parent.uri) &&
          (parent.local === "body" ||
            parent.local === "comment" ||
            parent.local === "docPartBody" ||
            parent.local === "endnote" ||
            parent.local === "ftr" ||
            parent.local === "footnote" ||
            parent.local === "hdr" ||
            parent.local === "txbxContent"),
        scalarControl: false,
        start,
        startEnd: parser.position,
        wordNamespace: tag.uri,
        wordPrefix: tag.name.includes(":")
          ? tag.name.slice(0, tag.name.indexOf(":"))
          : "",
      });
    } else if (tag.local === "sdtPr" && templateWordNamespaces.has(tag.uri)) {
      const control = controls.at(-1);
      if (control) {
        propertyControls.push(control);
      }
    } else if (
      propertyControls.length > 0 &&
      ((tag.local === "checkbox" && tag.uri === templateCheckboxNamespace) ||
        (templateWordNamespaces.has(tag.uri) &&
          (tag.local === "comboBox" ||
            tag.local === "date" ||
            tag.local === "dropDownList")))
    ) {
      const control = propertyControls.at(-1);
      if (control) {
        control.scalarControl = true;
      }
    } else if (
      tag.local === "sdtContent" &&
      templateWordNamespaces.has(tag.uri)
    ) {
      const control = controls.at(-1);
      if (control) {
        control.contentOpenEnd = parser.position;
        control.contentSelfClosing = tag.isSelfClosing;
        control.contentDepth = 0;
        for (const attribute of Object.values(tag.attributes)) {
          if (attribute.uri === templateXmlnsNamespace) {
            control.namespaceDeclarations.push({
              name: attribute.name,
              value: attribute.value,
            });
          }
        }
      }
    }
    elementStack.push({ local: tag.local, uri: tag.uri });
  });
  parser.on("closetag", (tag) => {
    if (tag.local === "sdtContent" && templateWordNamespaces.has(tag.uri)) {
      const control = controls.at(-1);
      if (control && control.contentOpenEnd !== null) {
        control.contentCloseStart = xml.lastIndexOf("<", parser.position - 1);
        control.contentDepth = null;
      }
    } else if (tag.local === "sdtPr" && templateWordNamespaces.has(tag.uri)) {
      propertyControls.pop();
    } else if (tag.local === "sdt" && templateWordNamespaces.has(tag.uri)) {
      const control = controls.pop();
      if (control?.contentIsBlock && control.parentContentControls) {
        for (const parentControl of control.parentContentControls) {
          parentControl.contentIsBlock = true;
        }
      }
      if (
        control?.scalarControl &&
        control.contentOpenEnd !== null &&
        control.contentCloseStart !== null
      ) {
        if (control.contentSelfClosing) {
          patches.push({
            end: parser.position,
            replacement: "",
            start: control.start,
          });
        } else {
          const needsParagraph =
            control.parentIsBlockContainer &&
            control.contentHasContent &&
            !control.contentIsBlock;
          const paragraphName = control.wordPrefix
            ? `${control.wordPrefix}:p`
            : "p";
          const namespaceAttribute = control.wordPrefix
            ? `xmlns:${control.wordPrefix}`
            : "xmlns";
          const paragraphOpen = needsParagraph
            ? `<${paragraphName} ${namespaceAttribute}="${htmlEscape(control.wordNamespace)}">`
            : "";
          const paragraphClose = needsParagraph ? `</${paragraphName}>` : "";
          patches.push(
            {
              end: control.contentOpenEnd,
              replacement: paragraphOpen,
              start: control.start,
            },
            {
              end: parser.position,
              replacement: paragraphClose,
              start: control.contentCloseStart,
            }
          );
        }
      } else if (control) {
        preserveNamespaces(
          control.start,
          control.startEnd,
          activeScalarNamespaceDeclarations(),
          control.namespaceAttributeNames
        );
      }
    }
    for (const control of controls) {
      if (control.contentDepth !== null && control.contentDepth > 0) {
        control.contentDepth -= 1;
      }
    }
    elementStack.pop();
  });
  try {
    parser.write(xml).close();
  } catch {
    fail(422, "invalid_template", "The DOCX package contains invalid XML");
  }
  let transformed = xml;
  for (const patch of patches.toSorted(
    (left, right) => right.start - left.start
  )) {
    transformed =
      transformed.slice(0, patch.start) +
      patch.replacement +
      transformed.slice(patch.end);
  }
  return transformed;
}
