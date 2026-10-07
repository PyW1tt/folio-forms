import { SaxesParser } from "saxes";

// oxlint-disable func-style prefer-destructuring no-await-in-loop no-use-before-define no-nested-ternary complexity no-shadow -- Route modules keep declaration order and sequential persistence invariants.
import { fail } from "../http/errors";
import { htmlEscape } from "../markup";
import { templateControlAttribute } from "./fields";
import { templateWordMainNamespace } from "./package";
import type { TemplateXmlElement } from "./package";

export interface NativeFontAttributeNames {
  ascii: string | null;
  asciiTheme: string | null;
  elementName: string;
  hAnsi: string | null;
  hAnsiTheme: string | null;
}

export interface NativeCheckboxState {
  font: string | null;
  glyph: string | null;
}

export function nativeCheckboxState(
  element: TemplateXmlElement
): NativeCheckboxState {
  const value = templateControlAttribute(element, "val");
  const codePoint =
    value && /^[\da-f]{1,6}$/iu.test(value)
      ? Number.parseInt(value, 16)
      : Number.NaN;
  return {
    font: templateControlAttribute(element, "font") ?? null,
    glyph:
      Number.isInteger(codePoint) &&
      codePoint <= 0x10_ff_ff &&
      !(codePoint >= 0xd8_00 && codePoint <= 0xdf_ff)
        ? String.fromCodePoint(codePoint)
        : null,
  };
}

export function nativeNamespacedAttribute(
  elementName: string,
  namespaceBindings: ReadonlyMap<string, string>,
  namespaceUri: string,
  localName: string,
  openingTag: string
): { declaration: string; name: string } {
  let prefix = elementName.includes(":")
    ? elementName.slice(0, elementName.indexOf(":"))
    : [...namespaceBindings].find(
        ([candidate, uri]) => candidate && uri === namespaceUri
      )?.[0];
  if (!prefix) {
    prefix = "word";
    for (
      let suffix = 1;
      namespaceBindings.has(prefix) || openingTag.includes(`${prefix}:`);
      suffix += 1
    ) {
      prefix = `word${suffix}`;
    }
    return {
      declaration: ` xmlns:${prefix}="${namespaceUri}"`,
      name: `${prefix}:${localName}`,
    };
  }
  return { declaration: "", name: `${prefix}:${localName}` };
}

export const nativeXmlAttributePattern =
  /(?<whitespace>\s)(?<name>[^\s="'<>/]+)(?<assignment>\s*=\s*)(?<quote>["'])[\s\S]*?\k<quote>/gu;

export const nativeXmlTagEndPattern = /(?<selfClosing>\/?)>$/u;

function nativeRemoveXmlAttribute(
  openingTag: string,
  attributeName: string
): string {
  return openingTag.replaceAll(
    nativeXmlAttributePattern,
    (match, _whitespace: string, name: string) =>
      name === attributeName ? "" : match
  );
}

export function applyNativeXmlPatches(
  xml: string,
  patches: readonly { end: number; replacement: string; start: number }[]
): string {
  let result = xml;
  for (const patch of patches.toSorted(
    (left, right) => right.start - left.start
  )) {
    result =
      result.slice(0, patch.start) +
      patch.replacement +
      result.slice(patch.end);
  }
  return result;
}

export function nativeTextOpeningTag(openingTag: string, text: string): string {
  const opening = openingTag.replace(/\/\s*>$/u, ">");
  return text.trim() === text || /\bxml:space\s*=/u.test(opening)
    ? opening
    : opening.replace(/>$/u, ' xml:space="preserve">');
}

export function nativeTextContentXml(
  name: string,
  openingTag: string,
  text: string
): string {
  const prefix = name.includes(":")
    ? name.slice(0, name.lastIndexOf(":") + 1)
    : "";
  const [firstPart = "", ...tabParts] = text.split("\t");
  let replacement = htmlEscape(firstPart);
  for (const part of tabParts) {
    replacement += `</${name}><${prefix}tab/>${nativeTextOpeningTag(openingTag, part)}${htmlEscape(part)}`;
  }
  return replacement;
}

export function nativeRunPropertiesXml(
  runProperties: string | null,
  contentPrefix: string,
  font: string | null,
  fontAttributeNames: NativeFontAttributeNames | null | undefined = undefined,
  wordNamespace = templateWordMainNamespace
): string {
  if (!font) {
    return runProperties ?? "";
  }
  const fontXml =
    contentPrefix.length > 0
      ? `<${contentPrefix}rFonts ${contentPrefix}ascii="${htmlEscape(font)}" ${contentPrefix}hAnsi="${htmlEscape(font)}"/>`
      : `<rFonts xmlns:word="${wordNamespace}" word:ascii="${htmlEscape(font)}" word:hAnsi="${htmlEscape(font)}"/>`;
  if (!runProperties) {
    return `<${contentPrefix}rPr>${fontXml}</${contentPrefix}rPr>`;
  }
  let fontElementStart: number | null = null;
  let fontElementOpenEnd: number | null = null;
  let fontElementEnd: number | null = null;
  let fontElementName: string | null = null;
  let fontOpeningTag = "";
  const parser = new SaxesParser({ position: true, xmlns: false });
  parser.on("opentag", (tag) => {
    const localName = tag.name.slice(tag.name.lastIndexOf(":") + 1);
    if (
      fontElementStart !== null ||
      localName !== "rFonts" ||
      fontAttributeNames === null ||
      (fontAttributeNames && tag.name !== fontAttributeNames.elementName)
    ) {
      return;
    }
    fontElementStart = runProperties.lastIndexOf("<", parser.position - 1);
    if (fontElementStart < 0) {
      return;
    }
    fontElementOpenEnd = parser.position;
    fontElementName = tag.name;
    const prefix = tag.name.includes(":")
      ? tag.name.slice(0, tag.name.lastIndexOf(":") + 1)
      : "";
    const setAttribute = (
      openingTag: string,
      attributeName: string
    ): string => {
      let replaced = false;
      const updated = openingTag.replaceAll(
        nativeXmlAttributePattern,
        (
          match,
          whitespace: string,
          name: string,
          assignment: string,
          quote: string
        ) => {
          if (name !== attributeName) {
            return match;
          }
          replaced = true;
          return `${whitespace}${name}${assignment}${quote}${htmlEscape(font)}${quote}`;
        }
      );
      return replaced
        ? updated
        : updated.replace(
            nativeXmlTagEndPattern,
            ` ${attributeName}="${htmlEscape(font)}"$<selfClosing>>`
          );
    };
    fontOpeningTag = runProperties.slice(fontElementStart, fontElementOpenEnd);
    let fontAttributePrefix =
      fontAttributeNames?.ascii?.slice(
        0,
        fontAttributeNames.ascii.lastIndexOf(":") + 1
      ) ??
      fontAttributeNames?.hAnsi?.slice(
        0,
        fontAttributeNames.hAnsi.lastIndexOf(":") + 1
      ) ??
      "";
    if (!fontAttributePrefix) {
      let prefix = "word";
      for (let suffix = 1; fontOpeningTag.includes(`${prefix}:`); suffix += 1) {
        prefix = `word${suffix}`;
      }
      fontAttributePrefix = `${prefix}:`;
      fontOpeningTag = fontOpeningTag.replace(
        nativeXmlTagEndPattern,
        ` xmlns:${prefix}="${wordNamespace}"$<selfClosing>>`
      );
    }
    fontOpeningTag = setAttribute(
      fontOpeningTag,
      fontAttributeNames?.ascii ?? `${fontAttributePrefix}ascii`
    );
    fontOpeningTag = setAttribute(
      fontOpeningTag,
      fontAttributeNames?.hAnsi ?? `${fontAttributePrefix}hAnsi`
    );
    if (fontAttributeNames) {
      for (const attributeName of [
        fontAttributeNames.asciiTheme,
        fontAttributeNames.hAnsiTheme,
      ]) {
        if (attributeName) {
          fontOpeningTag = nativeRemoveXmlAttribute(
            fontOpeningTag,
            attributeName
          );
        }
      }
    } else if (fontAttributeNames === undefined) {
      fontOpeningTag = nativeRemoveXmlAttribute(
        fontOpeningTag,
        `${prefix}asciiTheme`
      );
      fontOpeningTag = nativeRemoveXmlAttribute(
        fontOpeningTag,
        `${prefix}hAnsiTheme`
      );
    }
    if (tag.isSelfClosing) {
      fontElementEnd = parser.position;
    }
  });
  parser.on("closetag", (tag) => {
    if (
      fontElementStart !== null &&
      fontElementEnd === null &&
      tag.name === fontElementName
    ) {
      fontElementEnd = parser.position;
    }
  });
  try {
    parser.write(runProperties).close();
  } catch {
    fail(422, "invalid_template", "DOCX XML is malformed");
  }
  if (
    fontElementStart !== null &&
    fontElementOpenEnd !== null &&
    fontElementEnd !== null
  ) {
    return (
      runProperties.slice(0, fontElementStart) +
      fontOpeningTag +
      runProperties.slice(fontElementOpenEnd, fontElementEnd) +
      runProperties.slice(fontElementEnd)
    );
  }
  if (/\/\s*>$/u.test(runProperties)) {
    const propertiesName =
      runProperties.match(/^<(?<name>[^\s/>]+)/u)?.groups?.name;
    if (propertiesName) {
      return `${runProperties.replace(/\/\s*>$/u, ">")}${fontXml}</${propertiesName}>`;
    }
  }
  const closingTagStart = runProperties.lastIndexOf("</");
  if (closingTagStart === -1) {
    return runProperties;
  }
  return `${runProperties.slice(0, closingTagStart)}${fontXml}${runProperties.slice(closingTagStart)}`;
}

export function nativeApplyTextControlFontXml(
  xml: string,
  font: string,
  fontAttributeNamesByElement: readonly (NativeFontAttributeNames | null)[],
  wordNamespace = templateWordMainNamespace
): string {
  interface RunFrame {
    fontAttributeNames: NativeFontAttributeNames | null | undefined;
    hasText: boolean;
    name: string;
    openingEnd: number;
    propertiesEnd: number | null;
    propertiesName: string | null;
    propertiesStart: number | null;
  }
  const runStack: RunFrame[] = [];
  const patches: { end: number; replacement: string; start: number }[] = [];
  const parser = new SaxesParser({
    fragment: true,
    position: true,
    xmlns: false,
  });
  let fontElementIndex = 0;
  parser.on("opentag", (tag) => {
    const localName = tag.name.slice(tag.name.lastIndexOf(":") + 1);
    if (localName === "rFonts") {
      const fontAttributeNames =
        fontAttributeNamesByElement[fontElementIndex] ?? null;
      fontElementIndex += 1;
      const run = runStack.at(-1);
      if (run && run.fontAttributeNames === undefined) {
        run.fontAttributeNames = fontAttributeNames;
      }
    } else if (localName === "r") {
      runStack.push({
        fontAttributeNames: undefined,
        hasText: false,
        name: tag.name,
        openingEnd: parser.position,
        propertiesEnd: null,
        propertiesName: null,
        propertiesStart: null,
      });
    } else if (localName === "rPr" && runStack.length > 0) {
      const run = runStack.at(-1);
      if (run && run.propertiesStart === null) {
        run.propertiesName = tag.name;
        run.propertiesStart = xml.lastIndexOf("<", parser.position - 1);
        run.propertiesEnd = tag.isSelfClosing ? parser.position : null;
      }
    } else if (localName === "t" && runStack.length > 0) {
      const run = runStack.at(-1);
      if (run) {
        run.hasText = true;
      }
    }
  });
  parser.on("closetag", (tag) => {
    const localName = tag.name.slice(tag.name.lastIndexOf(":") + 1);
    if (localName === "rPr") {
      const run = runStack.at(-1);
      if (run?.propertiesName === tag.name && run.propertiesEnd === null) {
        run.propertiesEnd = parser.position;
      }
    } else if (localName === "r" && runStack.at(-1)?.name === tag.name) {
      const run = runStack.pop();
      if (!run?.hasText) {
        return;
      }
      if (
        run.propertiesStart !== null &&
        run.propertiesEnd !== null &&
        run.propertiesName
      ) {
        const properties = xml.slice(run.propertiesStart, run.propertiesEnd);
        const prefix = run.propertiesName.includes(":")
          ? run.propertiesName.slice(0, run.propertiesName.lastIndexOf(":") + 1)
          : "";
        const replacement = nativeRunPropertiesXml(
          properties,
          prefix,
          font,
          run.fontAttributeNames,
          wordNamespace
        );
        if (replacement !== properties) {
          patches.push({
            end: run.propertiesEnd,
            replacement,
            start: run.propertiesStart,
          });
        }
        return;
      }
      const prefix = run.name.includes(":")
        ? run.name.slice(0, run.name.lastIndexOf(":") + 1)
        : "";
      patches.push({
        end: run.openingEnd,
        replacement: nativeRunPropertiesXml(
          null,
          prefix,
          font,
          undefined,
          wordNamespace
        ),
        start: run.openingEnd,
      });
    }
  });
  parser.on("doctype", () => {
    fail(422, "invalid_template", "DOCX XML document types are not allowed");
  });
  try {
    parser.write(xml).close();
  } catch {
    fail(422, "invalid_template", "DOCX XML is malformed");
  }
  return applyNativeXmlPatches(xml, patches);
}
