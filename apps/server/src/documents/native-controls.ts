import { FieldType } from "@onlyoffice/db";
import { SaxesParser } from "saxes";

// oxlint-disable func-style prefer-destructuring no-await-in-loop no-use-before-define no-nested-ternary complexity no-shadow -- Route modules keep declaration order and sequential persistence invariants.
import { fail } from "../http/errors";
import { htmlEscape } from "../markup";
import type { JsonRecord } from "../model-types";
import { templateControlAttribute } from "./fields";
import {
  nativeDateDisplayValue,
  nativeRewriteTextControlContentXml,
} from "./native-text";
import {
  applyNativeXmlPatches,
  nativeApplyTextControlFontXml,
  nativeCheckboxState,
  nativeNamespacedAttribute,
  nativeRunPropertiesXml,
  nativeXmlAttributePattern,
  nativeXmlTagEndPattern,
} from "./native-xml";
import type {
  NativeCheckboxState,
  NativeFontAttributeNames,
} from "./native-xml";
import {
  templateCheckboxNamespace,
  templateMarkupCompatibilityNamespace,
  templateWordNamespaces,
  templateXmlElement,
  templateXmlnsNamespace,
} from "./package";

interface NativeFieldControlFrame {
  block: boolean;
  checkboxElement: {
    attributeDeclaration: string;
    attributeName: string;
    end: number;
    name: string;
    prefix: string;
    start: number;
  } | null;
  checkboxHasChecked: boolean;
  checkedState: NativeCheckboxState | null;
  contentEnd: number | null;
  contentName: string | null;
  contentPrefix: string;
  contentNamespace: string;
  contentSelfClosing: boolean;
  contentStart: number | null;
  contentTokenStart: number | null;
  paragraphProperties: string | null;
  paragraphPropertiesStart: number | null;
  runProperties: string | null;
  runPropertiesStart: number | null;
  contentFontAttributeNames: (NativeFontAttributeNames | null)[];
  runPropertiesFontAttributeNames: NativeFontAttributeNames | null | undefined;
  inPropertiesDepth: number;
  dateFormat: string | null;
  tag: string | null;
  uncheckedState: NativeCheckboxState | null;
}

export function nativeFieldControlXml(
  xml: string,
  data: JsonRecord,
  fieldsByTag: Map<string, { options: unknown; type: FieldType }>,
  selectedTags?: ReadonlySet<string>
): string {
  const controlTags: (string | null)[] = [];
  let alternateTagFallbackDepth = 0;
  let controlTag: string | null = null;
  let inControlProperties = false;
  const tagParser = new SaxesParser({ xmlns: true });
  tagParser.on("doctype", () => {
    fail(422, "invalid_template", "DOCX XML document types are not allowed");
  });
  tagParser.on("opentag", (tag) => {
    if (
      tag.local === "Fallback" &&
      tag.uri === templateMarkupCompatibilityNamespace
    ) {
      alternateTagFallbackDepth += 1;
      return;
    }
    if (alternateTagFallbackDepth > 0) {
      return;
    }
    if (tag.local === "sdt" && templateWordNamespaces.has(tag.uri)) {
      controlTag = null;
    } else if (tag.local === "sdtPr" && templateWordNamespaces.has(tag.uri)) {
      inControlProperties = true;
    } else if (
      inControlProperties &&
      tag.local === "tag" &&
      templateWordNamespaces.has(tag.uri)
    ) {
      controlTag =
        templateControlAttribute(templateXmlElement(tag), "val") ?? null;
    }
  });
  tagParser.on("closetag", (tag) => {
    if (
      tag.local === "Fallback" &&
      tag.uri === templateMarkupCompatibilityNamespace
    ) {
      alternateTagFallbackDepth -= 1;
      return;
    }
    if (alternateTagFallbackDepth > 0) {
      return;
    }
    if (tag.local === "sdtPr" && templateWordNamespaces.has(tag.uri)) {
      inControlProperties = false;
    } else if (tag.local === "sdt" && templateWordNamespaces.has(tag.uri)) {
      controlTags.push(controlTag?.trim() ?? null);
    }
  });
  try {
    tagParser.write(xml).close();
  } catch {
    fail(422, "invalid_template", "DOCX XML is malformed");
  }
  let controlIndex = 0;
  const controls: NativeFieldControlFrame[] = [];
  const elements: {
    local: string;
    namespaces: Map<string, string>;
    uri: string;
  }[] = [];
  const patches: { end: number; replacement: string; start: number }[] = [];
  const showingPlcHdrStarts: { remove: boolean; start: number }[] = [];
  let alternateFallbackDepth = 0;
  const parser = new SaxesParser({ position: true, xmlns: true });
  parser.on("doctype", () => {
    fail(422, "invalid_template", "DOCX XML document types are not allowed");
  });
  parser.on("opentag", (tag) => {
    const namespaceBindings = new Map(elements.at(-1)?.namespaces);
    for (const [attributeName, attribute] of Object.entries(tag.attributes)) {
      if (attribute.uri === templateXmlnsNamespace) {
        namespaceBindings.set(
          attributeName === "xmlns"
            ? ""
            : attributeName.slice(attributeName.indexOf(":") + 1),
          attribute.value
        );
      }
    }
    const element = {
      local: tag.local,
      namespaces: namespaceBindings,
      uri: tag.uri,
    };
    if (
      tag.local === "Fallback" &&
      tag.uri === templateMarkupCompatibilityNamespace
    ) {
      alternateFallbackDepth += 1;
      elements.push(element);
      return;
    }
    if (alternateFallbackDepth > 0) {
      elements.push(element);
      return;
    }
    if (tag.local === "sdt" && templateWordNamespaces.has(tag.uri)) {
      if (controls.length > 0) {
        fail(422, "invalid_template", "Nested native fields are unsupported");
      }
      controls.push({
        block: !elements.some(
          (element) =>
            element.local === "p" && templateWordNamespaces.has(element.uri)
        ),
        checkboxElement: null,
        checkboxHasChecked: false,
        checkedState: null,
        contentEnd: null,
        contentFontAttributeNames: [],
        contentName: null,
        contentNamespace: "",
        contentPrefix: "",
        contentSelfClosing: false,
        contentStart: null,
        contentTokenStart: null,
        dateFormat: null,
        inPropertiesDepth: 0,
        paragraphProperties: null,
        paragraphPropertiesStart: null,
        runProperties: null,
        runPropertiesFontAttributeNames: undefined,
        runPropertiesStart: null,
        tag: controlTags[controlIndex] ?? null,
        uncheckedState: null,
      });
      controlIndex += 1;
    } else {
      const frame = controls.at(-1);
      if (frame) {
        if (
          tag.local === "rFonts" &&
          frame.contentStart !== null &&
          frame.contentEnd === null
        ) {
          let fontAttributeNames: NativeFontAttributeNames | null = null;
          if (templateWordNamespaces.has(tag.uri)) {
            fontAttributeNames = {
              ascii: null,
              asciiTheme: null,
              elementName: tag.name,
              hAnsi: null,
              hAnsiTheme: null,
            };
            for (const [attributeName, attribute] of Object.entries(
              tag.attributes
            )) {
              if (attribute.uri !== tag.uri) {
                continue;
              }
              if (attribute.local === "ascii") {
                fontAttributeNames.ascii = attributeName;
              } else if (attribute.local === "hAnsi") {
                fontAttributeNames.hAnsi = attributeName;
              } else if (attribute.local === "asciiTheme") {
                fontAttributeNames.asciiTheme = attributeName;
              } else if (attribute.local === "hAnsiTheme") {
                fontAttributeNames.hAnsiTheme = attributeName;
              }
            }
          }
          frame.contentFontAttributeNames.push(fontAttributeNames);
          if (
            frame.runPropertiesStart !== null &&
            frame.runPropertiesFontAttributeNames === undefined
          ) {
            frame.runPropertiesFontAttributeNames = fontAttributeNames;
          }
        }
        if (tag.local === "sdtPr" && templateWordNamespaces.has(tag.uri)) {
          frame.inPropertiesDepth = 1;
        } else if (frame.inPropertiesDepth > 0) {
          if (
            frame.inPropertiesDepth === 1 &&
            tag.local === "tag" &&
            templateWordNamespaces.has(tag.uri)
          ) {
            const start = xml.lastIndexOf("<", parser.position - 1);
            frame.tag =
              templateControlAttribute(templateXmlElement(tag), "val") ?? null;
            if (
              start !== -1 &&
              frame.tag !== null &&
              (!selectedTags || selectedTags.has(frame.tag.trim()))
            ) {
              const openingTag = xml.slice(start, parser.position);
              const normalizedTag = frame.tag.trim();
              if (normalizedTag !== frame.tag) {
                const tagValueAttributeName = Object.entries(
                  tag.attributes
                ).find(
                  ([, attribute]) =>
                    attribute.local === "val" && attribute.uri === tag.uri
                )?.[0];
                if (tagValueAttributeName) {
                  patches.push({
                    end: parser.position,
                    replacement: openingTag.replaceAll(
                      nativeXmlAttributePattern,
                      (match, whitespace, attributeName, assignment, quote) =>
                        attributeName === tagValueAttributeName
                          ? `${whitespace}${attributeName}${assignment}${quote}${htmlEscape(normalizedTag)}${quote}`
                          : match
                    ),
                    start,
                  });
                }
              }
            }
          }
          if (
            tag.local === "dateFormat" &&
            templateWordNamespaces.has(tag.uri)
          ) {
            frame.dateFormat =
              templateControlAttribute(templateXmlElement(tag), "val") ?? null;
          }
          if (
            tag.local === "showingPlcHdr" &&
            templateWordNamespaces.has(tag.uri)
          ) {
            const start = xml.lastIndexOf("<", parser.position - 1);
            const remove =
              !selectedTags || selectedTags.has(frame.tag?.trim() ?? "");
            if (tag.isSelfClosing && remove && start !== -1) {
              patches.push({ end: parser.position, replacement: "", start });
            }
            showingPlcHdrStarts.push({
              remove: remove && !tag.isSelfClosing,
              start,
            });
          }
          const field = frame.tag
            ? fieldsByTag.get(frame.tag.trim())
            : undefined;
          const propertyName =
            tag.local === "date" && field?.type === FieldType.date
              ? "fullDate"
              : (tag.local === "dropDownList" &&
                    field?.type === FieldType.dropdown) ||
                  (tag.local === "comboBox" && field?.type === FieldType.combo)
                ? "lastValue"
                : null;
          if (propertyName) {
            const start = xml.lastIndexOf("<", parser.position - 1);
            const openingTag =
              start === -1 ? "" : xml.slice(start, parser.position);
            const propertyAttributeName = Object.entries(tag.attributes).find(
              ([, attribute]) =>
                attribute.local === propertyName && attribute.uri === tag.uri
            )?.[0];
            const storedValue = frame.tag ? data[frame.tag.trim()] : undefined;
            const replacementValue =
              typeof storedValue === "string" &&
              !(propertyName === "fullDate" && storedValue === "")
                ? propertyName === "fullDate"
                  ? `${storedValue}T00:00:00Z`
                  : storedValue
                : null;
            if (
              start !== -1 &&
              (propertyAttributeName || replacementValue !== null)
            ) {
              let replacement: string;
              if (propertyAttributeName) {
                replacement = openingTag.replaceAll(
                  nativeXmlAttributePattern,
                  (match, whitespace, name, assignment, quote) =>
                    name === propertyAttributeName
                      ? replacementValue === null
                        ? ""
                        : `${whitespace}${name}${assignment}${quote}${htmlEscape(replacementValue)}${quote}`
                      : match
                );
              } else {
                let prefix = tag.name.includes(":")
                  ? tag.name.slice(0, tag.name.indexOf(":") + 1)
                  : (Object.entries(tag.attributes)
                      .find(
                        ([name, attribute]) =>
                          name.includes(":") && attribute.uri === tag.uri
                      )?.[0]
                      .replace(/[^:]+$/u, ":") ?? "");
                let namespaceDeclaration = "";
                if (!prefix) {
                  let localPrefix = "word";
                  for (
                    let suffix = 1;
                    openingTag.includes(`${localPrefix}:`);
                    suffix += 1
                  ) {
                    localPrefix = `word${suffix}`;
                  }
                  prefix = `${localPrefix}:`;
                  namespaceDeclaration = ` xmlns:${localPrefix}="${tag.uri}"`;
                }
                replacement = openingTag.replace(
                  nativeXmlTagEndPattern,
                  ` ${prefix}${propertyName}="${htmlEscape(replacementValue ?? "")}"${namespaceDeclaration}$<selfClosing>>`
                );
              }
              patches.push({
                end: parser.position,
                replacement,
                start,
              });
            }
          }
          if (
            tag.local === "checkbox" &&
            tag.uri === templateCheckboxNamespace &&
            frame.tag &&
            fieldsByTag.get(frame.tag.trim())?.type === FieldType.checkbox
          ) {
            const start = xml.lastIndexOf("<", parser.position - 1);
            const openingTag =
              start === -1 ? "" : xml.slice(start, parser.position);
            const prefix = tag.name.includes(":")
              ? tag.name.slice(0, tag.name.indexOf(":") + 1)
              : "";
            const valueAttribute = nativeNamespacedAttribute(
              tag.name,
              namespaceBindings,
              tag.uri,
              "val",
              openingTag
            );
            const checkedValue = data[frame.tag.trim()] === true ? "1" : "0";
            if (/\/\s*>$/u.test(openingTag)) {
              patches.push({
                end: parser.position,
                replacement: `${openingTag.replace(/\/\s*>$/u, ">")}<${prefix}checked ${valueAttribute.name}="${checkedValue}"${valueAttribute.declaration}/></${tag.name}>`,
                start,
              });
            } else {
              frame.checkboxElement = {
                attributeDeclaration: valueAttribute.declaration,
                attributeName: valueAttribute.name,
                end: parser.position,
                name: tag.name,
                prefix,
                start,
              };
              frame.checkboxHasChecked = false;
            }
          }
          if (
            tag.local === "checkedState" &&
            tag.uri === templateCheckboxNamespace &&
            frame.tag &&
            fieldsByTag.get(frame.tag.trim())?.type === FieldType.checkbox
          ) {
            frame.checkedState = nativeCheckboxState(templateXmlElement(tag));
          }
          if (
            tag.local === "uncheckedState" &&
            tag.uri === templateCheckboxNamespace &&
            frame.tag &&
            fieldsByTag.get(frame.tag.trim())?.type === FieldType.checkbox
          ) {
            frame.uncheckedState = nativeCheckboxState(templateXmlElement(tag));
          }
          if (
            tag.local === "checked" &&
            tag.uri === templateCheckboxNamespace &&
            frame.tag &&
            fieldsByTag.get(frame.tag.trim())?.type === FieldType.checkbox
          ) {
            frame.checkboxHasChecked = true;
            const start = xml.lastIndexOf("<", parser.position - 1);
            const checkedAttributeName = Object.entries(tag.attributes).find(
              ([, attribute]) =>
                attribute.local === "val" && attribute.uri === tag.uri
            )?.[0];
            const unqualifiedValueName = Object.entries(tag.attributes).find(
              ([, attribute]) =>
                attribute.local === "val" && attribute.uri === ""
            )?.[0];
            if (start !== -1) {
              const openingTag = xml.slice(start, parser.position);
              const checkedValue = data[frame.tag.trim()] === true ? "1" : "0";
              const qualifiedAttribute = checkedAttributeName
                ? { declaration: "", name: checkedAttributeName }
                : nativeNamespacedAttribute(
                    tag.name,
                    namespaceBindings,
                    tag.uri,
                    "val",
                    openingTag
                  );
              const replacement =
                checkedAttributeName || unqualifiedValueName
                  ? openingTag.replaceAll(
                      nativeXmlAttributePattern,
                      (match, whitespace, name, assignment, quote) =>
                        name === checkedAttributeName ||
                        name === unqualifiedValueName
                          ? `${whitespace}${qualifiedAttribute.name}${assignment}${quote}${checkedValue}${quote}${qualifiedAttribute.declaration}`
                          : match
                    )
                  : openingTag.replace(
                      nativeXmlTagEndPattern,
                      ` ${qualifiedAttribute.name}="${checkedValue}"${qualifiedAttribute.declaration}$<selfClosing>>`
                    );
              patches.push({
                end: parser.position,
                replacement,
                start,
              });
            }
          }
          frame.inPropertiesDepth += 1;
        } else if (
          frame.contentStart !== null &&
          frame.inPropertiesDepth === 0 &&
          tag.local === "pPr" &&
          templateWordNamespaces.has(tag.uri) &&
          frame.paragraphProperties === null &&
          frame.paragraphPropertiesStart === null
        ) {
          frame.paragraphPropertiesStart = xml.lastIndexOf(
            "<",
            parser.position - 1
          );
        } else if (
          frame.contentStart !== null &&
          frame.inPropertiesDepth === 0 &&
          tag.local === "rPr" &&
          templateWordNamespaces.has(tag.uri) &&
          frame.runProperties === null &&
          frame.runPropertiesStart === null
        ) {
          frame.runPropertiesStart = xml.lastIndexOf("<", parser.position - 1);
        } else if (
          tag.local === "sdtContent" &&
          templateWordNamespaces.has(tag.uri)
        ) {
          const start = xml.lastIndexOf("<", parser.position - 1);
          const openingTag =
            start === -1 ? "" : xml.slice(start, parser.position);
          frame.contentNamespace = tag.uri;
          frame.contentName = tag.name;
          frame.contentPrefix = tag.name.includes(":")
            ? tag.name.slice(0, tag.name.indexOf(":") + 1)
            : "";
          frame.contentSelfClosing = /\/\s*>$/u.test(openingTag);
          frame.contentStart = frame.contentSelfClosing
            ? null
            : parser.position;
          frame.contentTokenStart = frame.contentSelfClosing ? start : null;
        }
      }
    }
    elements.push(element);
  });
  parser.on("closetag", (tag) => {
    if (
      tag.local === "Fallback" &&
      tag.uri === templateMarkupCompatibilityNamespace
    ) {
      alternateFallbackDepth -= 1;
      elements.pop();
      return;
    }
    if (alternateFallbackDepth > 0) {
      elements.pop();
      return;
    }
    if (tag.local === "showingPlcHdr" && templateWordNamespaces.has(tag.uri)) {
      const placeholder = showingPlcHdrStarts.pop();
      if (placeholder?.remove && placeholder.start >= 0) {
        patches.push({
          end: parser.position,
          replacement: "",
          start: placeholder.start,
        });
      }
    }
    const frame = controls.at(-1);
    if (frame) {
      if (
        frame.contentStart !== null &&
        frame.paragraphPropertiesStart !== null &&
        tag.local === "pPr" &&
        templateWordNamespaces.has(tag.uri)
      ) {
        frame.paragraphProperties = xml.slice(
          frame.paragraphPropertiesStart,
          parser.position
        );
        frame.paragraphPropertiesStart = null;
      } else if (
        frame.contentStart !== null &&
        frame.runPropertiesStart !== null &&
        tag.local === "rPr" &&
        templateWordNamespaces.has(tag.uri)
      ) {
        frame.runProperties = xml.slice(
          frame.runPropertiesStart,
          parser.position
        );
        frame.runPropertiesStart = null;
      } else if (tag.local === "sdtPr" && templateWordNamespaces.has(tag.uri)) {
        frame.inPropertiesDepth = 0;
      } else if (frame.inPropertiesDepth > 0) {
        if (
          tag.local === "checkbox" &&
          tag.uri === templateCheckboxNamespace &&
          frame.checkboxElement
        ) {
          if (!frame.checkboxHasChecked) {
            const { attributeDeclaration, attributeName, end, prefix, start } =
              frame.checkboxElement;
            const checkedValue =
              data[frame.tag?.trim() ?? ""] === true ? "1" : "0";
            patches.push({
              end,
              replacement: `${xml.slice(start, end)}<${prefix}checked ${attributeName}="${checkedValue}"${attributeDeclaration}/>`,
              start,
            });
          }
          frame.checkboxElement = null;
        }
        frame.inPropertiesDepth -= 1;
      } else if (
        tag.local === "sdtContent" &&
        templateWordNamespaces.has(tag.uri)
      ) {
        frame.contentEnd = frame.contentSelfClosing
          ? parser.position
          : xml.lastIndexOf("</", parser.position - 1);
        const completed = controls.pop();
        if (!completed?.tag || !completed.contentName) {
          fail(422, "invalid_template", "A native field is incomplete");
        }
        if (selectedTags && !selectedTags.has(completed.tag.trim())) {
          elements.pop();
          return;
        }
        const tagName = completed.tag.trim();
        const field = fieldsByTag.get(tagName);
        const storedValue = data[tagName];
        const checkboxChecked = storedValue === true;
        const checkboxState = checkboxChecked
          ? completed.checkedState
          : completed.uncheckedState;
        const value =
          field?.type === FieldType.checkbox
            ? (checkboxState?.glyph ?? (checkboxChecked ? "☒" : "☐"))
            : typeof storedValue === "string"
              ? storedValue
              : "";
        const option =
          typeof storedValue === "string" &&
          (field?.type === FieldType.dropdown ||
            field?.type === FieldType.combo)
            ? (Array.isArray(field.options) ? field.options : []).find(
                (item) =>
                  item &&
                  typeof item === "object" &&
                  !Array.isArray(item) &&
                  (item as Record<string, unknown>).value === storedValue
              )
            : null;
        const displayValue =
          field?.type === FieldType.date
            ? nativeDateDisplayValue(value, completed.dateFormat)
            : option &&
                typeof (option as Record<string, unknown>).displayText ===
                  "string"
              ? ((option as Record<string, unknown>).displayText as string)
              : value;
        const existingContent =
          !completed.contentSelfClosing &&
          completed.contentStart !== null &&
          completed.contentEnd !== null
            ? xml.slice(completed.contentStart, completed.contentEnd)
            : null;
        let preservedContent =
          existingContent === null
            ? null
            : nativeRewriteTextControlContentXml(existingContent, displayValue);
        if (preservedContent !== null && checkboxState?.font) {
          preservedContent = nativeApplyTextControlFontXml(
            preservedContent,
            checkboxState.font,
            completed.contentFontAttributeNames,
            completed.contentNamespace
          );
        }
        if (preservedContent === null) {
          const lines = displayValue.split(/\r\n|\r|\n/u);
          const text = lines
            .map(
              (line, index) =>
                `${index > 0 ? `<${completed.contentPrefix}br/>` : ""}<${completed.contentPrefix}t xml:space="preserve">${htmlEscape(line)}</${completed.contentPrefix}t>`
            )
            .join("");
          const runProperties = nativeRunPropertiesXml(
            completed.runProperties,
            completed.contentPrefix,
            checkboxState?.font ?? null,
            completed.runPropertiesFontAttributeNames,
            completed.contentNamespace
          );
          const run = `<${completed.contentPrefix}r>${runProperties}${text}</${completed.contentPrefix}r>`;
          const content = completed.block
            ? `<${completed.contentPrefix}p>${completed.paragraphProperties ?? ""}${run}</${completed.contentPrefix}p>`
            : run;
          if (completed.contentSelfClosing) {
            if (
              completed.contentTokenStart === null ||
              completed.contentEnd === null
            ) {
              fail(422, "invalid_template", "A native field is incomplete");
            }
            const openingTag = xml.slice(
              completed.contentTokenStart,
              completed.contentEnd
            );
            const expandedOpeningTag = openingTag.replace(/\/\s*>$/u, ">");
            patches.push({
              end: completed.contentEnd,
              replacement: `${expandedOpeningTag}${content}</${completed.contentName}>`,
              start: completed.contentTokenStart,
            });
          } else if (
            completed.contentStart !== null &&
            completed.contentEnd !== null &&
            completed.contentEnd >= completed.contentStart
          ) {
            patches.push({
              end: completed.contentEnd,
              replacement: content,
              start: completed.contentStart,
            });
          } else {
            fail(422, "invalid_template", "A native field is incomplete");
          }
        } else if (preservedContent !== existingContent) {
          const contentStart = completed.contentStart;
          const contentEnd = completed.contentEnd;
          if (contentStart === null || contentEnd === null) {
            fail(422, "invalid_template", "A native field is incomplete");
          }
          patches.push({
            end: contentEnd,
            replacement: preservedContent,
            start: contentStart,
          });
        }
      }
    }
    elements.pop();
  });
  try {
    parser.write(xml).close();
  } catch {
    fail(422, "invalid_template", "The DOCX package contains invalid XML");
  }
  return applyNativeXmlPatches(xml, patches);
}
