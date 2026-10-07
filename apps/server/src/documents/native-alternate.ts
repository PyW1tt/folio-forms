import { SaxesParser } from "saxes";

// oxlint-disable func-style prefer-destructuring no-await-in-loop no-use-before-define no-nested-ternary complexity no-shadow -- Route modules keep declaration order and sequential persistence invariants.
import { fail } from "../http/errors";
import { htmlEscape } from "../markup";
import { templateControlAttribute } from "./fields";
import { applyNativeXmlPatches, nativeXmlTagEndPattern } from "./native-xml";
import {
  templateControlNamespaces,
  templateMarkupCompatibilityNamespace,
  templateWordNamespaces,
  templateXmlElement,
  templateXmlnsNamespace,
} from "./package";

function nativePreserveAlternateNamespaceDeclarations(
  xml: string,
  declarations: readonly { name: string; value: string }[]
): string {
  if (declarations.length === 0) {
    return xml;
  }
  let depth = 0;
  const patches: { end: number; replacement: string; start: number }[] = [];
  const parser = new SaxesParser({
    fragment: true,
    position: true,
    xmlns: false,
  });
  parser.on("doctype", () => {
    fail(422, "invalid_template", "DOCX XML document types are not allowed");
  });
  parser.on("opentag", (tag) => {
    if (depth === 0) {
      const start = xml.lastIndexOf("<", parser.position - 1);
      if (start !== -1) {
        const originalTag = xml.slice(start, parser.position);
        let replacement = originalTag;
        for (const declaration of declarations) {
          if (!Object.hasOwn(tag.attributes, declaration.name)) {
            replacement = replacement.replace(
              nativeXmlTagEndPattern,
              ` ${declaration.name}="${htmlEscape(declaration.value)}"$<selfClosing>>`
            );
          }
        }
        if (replacement !== originalTag) {
          patches.push({ end: parser.position, replacement, start });
        }
      }
    }
    if (!tag.isSelfClosing) {
      depth += 1;
    }
  });
  parser.on("closetag", (tag) => {
    if (!tag.isSelfClosing) {
      depth -= 1;
    }
  });
  try {
    parser.write(xml).close();
  } catch {
    fail(422, "invalid_template", "DOCX XML is malformed");
  }
  return applyNativeXmlPatches(xml, patches);
}

// PDF conversion omits AlternateContent fields; retain compatible Choice content and namespace bindings.
export function nativeFlattenAlternateFieldsXml(
  xml: string,
  fieldTags: ReadonlySet<string>,
  validateOnly = false
): { xml: string; supported: boolean } {
  let result = xml;
  while (result.includes("AlternateContent")) {
    const sourceXml = result;
    interface AlternateContentFrame {
      branch: "choice" | "fallback" | null;
      choiceContentStart: number | null;
      choiceFieldCount: number;
      choiceFieldEnd: number | null;
      choiceFieldStart: number | null;
      choiceHasField: boolean;
      choiceNamespaceDeclarations: Map<string, string> | null;
      choiceNamespaces: Map<string, string> | null;
      choiceRequiresSupported: boolean;
      fallbackHasField: boolean;
      hasField: boolean;
      hasNestedField: boolean;
      namespaceDeclarations: Map<string, string>;
      parent: AlternateContentFrame | null;
      start: number;
      unsupported: boolean;
      firstSupportedChoiceHasField: boolean | null;
    }
    interface AlternateControlFrame {
      alternate: AlternateContentFrame | null;
      branch: "choice" | "fallback" | null;
      inPropertiesDepth: number;
      start: number;
      tag: string | null;
    }
    const alternates: AlternateContentFrame[] = [];
    const controls: AlternateControlFrame[] = [];
    const patches: { end: number; replacement: string; start: number }[] = [];
    const namespaceBindings = new Map<string, string>([
      ["xml", "http://www.w3.org/XML/1998/namespace"],
    ]);
    const namespaceChanges: {
      prefix: string;
      previous: string | undefined;
    }[][] = [];
    const parser = new SaxesParser({ position: true, xmlns: true });
    let unsupported = false;
    parser.on("doctype", () => {
      fail(422, "invalid_template", "DOCX XML document types are not allowed");
    });
    parser.on("opentag", (tag) => {
      const changes: { prefix: string; previous: string | undefined }[] = [];
      for (const attribute of Object.values(tag.attributes)) {
        if (attribute.uri !== templateXmlnsNamespace) {
          continue;
        }
        const prefix = attribute.name === "xmlns" ? "" : attribute.local;
        changes.push({ prefix, previous: namespaceBindings.get(prefix) });
        namespaceBindings.set(prefix, attribute.value);
      }
      namespaceChanges.push(changes);
      if (
        tag.local === "AlternateContent" &&
        tag.uri === templateMarkupCompatibilityNamespace
      ) {
        const declarations = new Map<string, string>();
        for (const attribute of Object.values(tag.attributes)) {
          if (attribute.uri === templateXmlnsNamespace) {
            declarations.set(attribute.name, attribute.value);
          }
        }
        alternates.push({
          branch: null,
          choiceContentStart: null,
          choiceFieldCount: 0,
          choiceFieldEnd: null,
          choiceFieldStart: null,
          choiceHasField: false,
          choiceNamespaceDeclarations: null,
          choiceNamespaces: null,
          choiceRequiresSupported: false,
          fallbackHasField: false,
          firstSupportedChoiceHasField: null,
          hasField: false,
          hasNestedField: false,
          namespaceDeclarations: declarations,
          parent: alternates.at(-1) ?? null,
          start: sourceXml.lastIndexOf("<", parser.position - 1),
          unsupported: false,
        });
      } else if (
        tag.local === "Choice" &&
        tag.uri === templateMarkupCompatibilityNamespace
      ) {
        const alternate = alternates.at(-1);
        if (alternate) {
          const requires = Object.values(tag.attributes).find(
            (attribute) =>
              attribute.local === "Requires" && attribute.uri === ""
          )?.value;
          const requiredPrefixes =
            requires?.trim().split(/\s+/u).filter(Boolean) ?? [];
          alternate.choiceRequiresSupported =
            requiredPrefixes.length > 0 &&
            requiredPrefixes.every((prefix) =>
              templateControlNamespaces.has(namespaceBindings.get(prefix) ?? "")
            );
          alternate.choiceNamespaceDeclarations = new Map(
            alternate.namespaceDeclarations
          );
          for (const attribute of Object.values(tag.attributes)) {
            if (attribute.uri === templateXmlnsNamespace) {
              alternate.choiceNamespaceDeclarations.set(
                attribute.name,
                attribute.value
              );
            }
          }
          alternate.branch = "choice";
          alternate.choiceContentStart = parser.position;
          alternate.choiceHasField = false;
        }
      } else if (
        tag.local === "Fallback" &&
        tag.uri === templateMarkupCompatibilityNamespace
      ) {
        const alternate = alternates.at(-1);
        if (alternate) {
          alternate.branch = "fallback";
        }
      }
      const control = controls.at(-1);
      if (tag.local === "sdt" && templateWordNamespaces.has(tag.uri)) {
        const alternate = alternates.at(-1) ?? null;
        controls.push({
          alternate,
          branch: alternate?.branch ?? null,
          inPropertiesDepth: 0,
          start: sourceXml.lastIndexOf("<", parser.position - 1),
          tag: null,
        });
      } else if (control) {
        if (tag.local === "sdtPr" && templateWordNamespaces.has(tag.uri)) {
          control.inPropertiesDepth = 1;
        } else if (control.inPropertiesDepth > 0) {
          if (
            control.inPropertiesDepth === 1 &&
            tag.local === "tag" &&
            templateWordNamespaces.has(tag.uri)
          ) {
            control.tag =
              templateControlAttribute(templateXmlElement(tag), "val") ?? null;
          }
          control.inPropertiesDepth += 1;
        }
      }
    });
    parser.on("closetag", (tag) => {
      if (tag.local === "sdt" && templateWordNamespaces.has(tag.uri)) {
        const completed = controls.pop();
        const fieldTag = completed?.tag?.trim();
        if (completed?.alternate && fieldTag && fieldTags.has(fieldTag)) {
          completed.alternate.hasField = true;
          if (completed.branch === "choice") {
            completed.alternate.choiceHasField = true;
          } else if (completed.branch === "fallback") {
            completed.alternate.fallbackHasField = true;
          }
        }
      } else {
        const control = controls.at(-1);
        if (
          control &&
          tag.local === "sdtPr" &&
          templateWordNamespaces.has(tag.uri)
        ) {
          control.inPropertiesDepth = 0;
        } else if (control && control.inPropertiesDepth > 0) {
          control.inPropertiesDepth -= 1;
        }
      }
      const alternate = alternates.at(-1);
      if (
        alternate &&
        tag.local === "Choice" &&
        tag.uri === templateMarkupCompatibilityNamespace
      ) {
        if (alternate.choiceRequiresSupported) {
          if (alternate.firstSupportedChoiceHasField === null) {
            alternate.firstSupportedChoiceHasField = alternate.choiceHasField;
          } else if (
            !alternate.firstSupportedChoiceHasField &&
            alternate.choiceHasField
          ) {
            alternate.unsupported = true;
          }
        }
        if (alternate.choiceHasField) {
          alternate.choiceFieldCount += 1;
          alternate.unsupported ||= !alternate.choiceRequiresSupported;
          if (
            alternate.choiceFieldStart === null &&
            alternate.choiceContentStart !== null
          ) {
            alternate.choiceFieldStart = alternate.choiceContentStart;
            alternate.choiceFieldEnd = sourceXml.lastIndexOf(
              "</",
              parser.position - 1
            );
            alternate.choiceNamespaces = alternate.choiceNamespaceDeclarations;
          }
        }
        alternate.branch = null;
        alternate.choiceContentStart = null;
        alternate.choiceHasField = false;
        alternate.choiceNamespaceDeclarations = null;
      } else if (
        alternate &&
        tag.local === "Fallback" &&
        tag.uri === templateMarkupCompatibilityNamespace
      ) {
        alternate.branch = null;
      } else if (
        alternate &&
        tag.local === "AlternateContent" &&
        tag.uri === templateMarkupCompatibilityNamespace
      ) {
        const completed = alternates.pop();
        if (completed) {
          completed.unsupported ||= completed.choiceFieldCount > 1;
          completed.unsupported ||= completed.fallbackHasField;
          unsupported ||= completed.unsupported;
          if (completed.hasField && completed.parent) {
            completed.parent.hasNestedField = true;
            completed.parent.hasField = true;
            if (completed.parent.branch === "choice") {
              completed.parent.choiceHasField = true;
            } else if (completed.parent.branch === "fallback") {
              completed.parent.fallbackHasField = true;
            }
          }
          if (
            !validateOnly &&
            completed.hasField &&
            !completed.hasNestedField &&
            completed.choiceFieldStart !== null &&
            completed.choiceFieldEnd !== null
          ) {
            const declarations = Array.from(
              completed.choiceNamespaces ?? [],
              ([name, value]) => ({ name, value })
            );
            patches.push({
              end: parser.position,
              replacement: nativePreserveAlternateNamespaceDeclarations(
                sourceXml.slice(
                  completed.choiceFieldStart,
                  completed.choiceFieldEnd
                ),
                declarations
              ),
              start: completed.start,
            });
          }
        }
      }
      const changes = namespaceChanges.pop() ?? [];
      for (const { prefix, previous } of changes) {
        if (previous === undefined) {
          namespaceBindings.delete(prefix);
        } else {
          namespaceBindings.set(prefix, previous);
        }
      }
    });
    try {
      parser.write(sourceXml).close();
    } catch {
      fail(422, "invalid_template", "The DOCX package contains invalid XML");
    }
    if (unsupported) {
      return { supported: false, xml: result };
    }
    if (validateOnly) {
      return { supported: true, xml: result };
    }
    const flattened = applyNativeXmlPatches(sourceXml, patches);
    if (flattened === result) {
      return { supported: true, xml: result };
    }
    result = flattened;
  }
  return { supported: true, xml: result };
}
