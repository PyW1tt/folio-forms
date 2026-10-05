import { SaxesParser } from "saxes";

// oxlint-disable func-style prefer-destructuring no-await-in-loop no-use-before-define no-nested-ternary complexity no-shadow -- Route modules keep declaration order and sequential persistence invariants.
import { fail } from "../http/errors";
import {
  applyNativeXmlPatches,
  nativeTextContentXml,
  nativeTextOpeningTag,
} from "./native-xml";

const nativeDateMonths = [
  "January",
  "February",
  "March",
  "April",
  "May",
  "June",
  "July",
  "August",
  "September",
  "October",
  "November",
  "December",
] as const;

const nativeDateWeekdays = [
  "Sunday",
  "Monday",
  "Tuesday",
  "Wednesday",
  "Thursday",
  "Friday",
  "Saturday",
] as const;

export function nativeDateDisplayValue(
  value: string,
  format: string | null
): string {
  if (!format) {
    return value;
  }
  const [yearText, monthText, dayText] = value.split("-");
  const year = Number(yearText);
  const month = Number(monthText);
  const day = Number(dayText);
  if (
    !yearText ||
    !monthText ||
    !dayText ||
    !Number.isInteger(year) ||
    !Number.isInteger(month) ||
    !Number.isInteger(day)
  ) {
    return value;
  }
  const monthName = nativeDateMonths[month - 1];
  const weekday =
    nativeDateWeekdays[new Date(`${value}T00:00:00Z`).getUTCDay()];
  if (!monthName || !weekday) {
    return value;
  }
  const monthNumber = String(month);
  const dayNumber = String(day);
  return format.replaceAll(
    /'[^']*'|"[^"]*"|dddd|ddd|yyyy|MMMM|MMM|yy|MM|dd|M|d/gu,
    (token) => {
      if (token.startsWith("'") || token.startsWith('"')) {
        return token.slice(1, -1);
      }
      switch (token) {
        case "yyyy": {
          return yearText;
        }
        case "yy": {
          return yearText.slice(-2);
        }
        case "dddd": {
          return weekday;
        }
        case "ddd": {
          return weekday.slice(0, 3);
        }
        case "MMMM": {
          return monthName;
        }
        case "MMM": {
          return monthName.slice(0, 3);
        }
        case "MM": {
          return monthNumber.padStart(2, "0");
        }
        case "M": {
          return monthNumber;
        }
        case "dd": {
          return dayNumber.padStart(2, "0");
        }
        case "d": {
          return dayNumber;
        }
        default: {
          return token;
        }
      }
    }
  );
}

export function nativeRewriteTextControlContentXml(
  xml: string,
  value: string
): string | null {
  interface TextNode {
    contentEnd: number;
    line: number;
    name: string;
    openingStart: number;
    openingTag: string;
    paragraph: number;
    selfClosing: boolean;
    start: number;
    tagEnd: number;
    text: string;
  }
  interface TabNode {
    end: number;
    line: number;
    name: string;
    paragraph: number;
    start: number;
  }
  interface Paragraph {
    end: number;
    index: number;
    start: number;
    startLine: number;
  }
  interface LineBreak {
    end: number;
    line: number;
    name: string;
    paragraph: number;
    start: number;
  }
  const textNodes: TextNode[] = [];
  const tabNodes: TabNode[] = [];
  const paragraphs: Paragraph[] = [];
  const lineBreaks: LineBreak[] = [];
  const paragraphStack: Omit<Paragraph, "end">[] = [];
  const breakStack: Omit<LineBreak, "end">[] = [];
  const tabStack: Omit<TabNode, "end">[] = [];
  let activeTextNode: Omit<TextNode, "contentEnd" | "tagEnd"> | null = null;
  let currentLine = 0;
  let currentParagraph = -1;
  let visibleText = "";
  const parser = new SaxesParser({
    fragment: true,
    position: true,
    xmlns: false,
  });
  parser.on("doctype", () => {
    fail(422, "invalid_template", "DOCX XML document types are not allowed");
  });
  parser.on("opentag", (tag) => {
    const localName = tag.name.slice(tag.name.lastIndexOf(":") + 1);
    if (localName === "p") {
      if (currentParagraph >= 0) {
        currentLine += 1;
        visibleText += "\n";
      }
      currentParagraph += 1;
      paragraphStack.push({
        index: currentParagraph,
        start: xml.lastIndexOf("<", parser.position - 1),
        startLine: currentLine,
      });
    } else if (localName === "br" || localName === "cr") {
      const lineBreak = {
        line: currentLine,
        name: tag.name,
        paragraph: currentParagraph,
        start: xml.lastIndexOf("<", parser.position - 1),
      };
      if (tag.isSelfClosing) {
        lineBreaks.push({ ...lineBreak, end: parser.position });
      } else {
        breakStack.push(lineBreak);
      }
      currentLine += 1;
      visibleText += "\n";
    } else if (localName === "tab") {
      const tab = {
        line: currentLine,
        name: tag.name,
        paragraph: currentParagraph,
        start: xml.lastIndexOf("<", parser.position - 1),
      };
      if (tag.isSelfClosing) {
        tabNodes.push({ ...tab, end: parser.position });
      } else {
        tabStack.push(tab);
      }
      visibleText += "\t";
    } else if (localName === "t") {
      const openingStart = xml.lastIndexOf("<", parser.position - 1);
      const textNode = {
        line: currentLine,
        name: tag.name,
        openingStart,
        openingTag: xml.slice(openingStart, parser.position),
        paragraph: currentParagraph,
        selfClosing: tag.isSelfClosing,
        start: parser.position,
        text: "",
      };
      if (tag.isSelfClosing) {
        textNodes.push({
          ...textNode,
          contentEnd: parser.position,
          tagEnd: parser.position,
        });
      } else {
        activeTextNode = textNode;
      }
    }
  });
  parser.on("text", (text) => {
    if (activeTextNode) {
      activeTextNode.text += text;
      visibleText += text;
    }
  });
  parser.on("closetag", (tag) => {
    const localName = tag.name.slice(tag.name.lastIndexOf(":") + 1);
    if (localName === "t" && activeTextNode?.name === tag.name) {
      const contentEnd = xml.lastIndexOf("</", parser.position - 1);
      textNodes.push({
        ...activeTextNode,
        contentEnd,
        tagEnd: parser.position,
      });
      activeTextNode = null;
    } else if (
      (localName === "br" || localName === "cr") &&
      breakStack.at(-1)?.name === tag.name
    ) {
      const lineBreak = breakStack.pop();
      if (lineBreak) {
        lineBreaks.push({ ...lineBreak, end: parser.position });
      }
    } else if (localName === "tab" && tabStack.at(-1)?.name === tag.name) {
      const tab = tabStack.pop();
      if (tab) {
        tabNodes.push({ ...tab, end: parser.position });
      }
    } else if (localName === "p") {
      const paragraph = paragraphStack.pop();
      if (paragraph) {
        paragraphs.push({ ...paragraph, end: parser.position });
      }
    }
  });
  try {
    parser.write(xml).close();
  } catch {
    fail(422, "invalid_template", "DOCX XML is malformed");
  }
  if (textNodes.length === 0 || visibleText === value) {
    return textNodes.length === 0 ? null : xml;
  }
  const lines = value.split(/\r\n|\r|\n/u);
  const existingLineCount = currentLine + 1;
  const removedParagraphs = new Set(
    lines.length < existingLineCount
      ? paragraphs
          .filter(({ startLine }) => startLine >= lines.length)
          .map(({ index }) => index)
      : []
  );
  const removedParagraphRanges = paragraphs.filter(({ index }) =>
    removedParagraphs.has(index)
  );
  const keptTextNodes = textNodes.filter(
    ({ paragraph }) => !removedParagraphs.has(paragraph)
  );
  const keptTabNodes = tabNodes.filter(
    ({ paragraph }) => !removedParagraphs.has(paragraph)
  );
  type ContentSegment =
    | { kind: "text"; node: TextNode }
    | { kind: "tab"; node: TabNode };
  const segmentsByLine = new Map<number, ContentSegment[]>();
  for (const node of keptTextNodes) {
    const segments = segmentsByLine.get(node.line) ?? [];
    segments.push({ kind: "text", node });
    segmentsByLine.set(node.line, segments);
  }
  for (const node of keptTabNodes) {
    const segments = segmentsByLine.get(node.line) ?? [];
    segments.push({ kind: "tab", node });
    segmentsByLine.set(node.line, segments);
  }
  for (const segments of segmentsByLine.values()) {
    segments.sort((left, right) => {
      const leftStart =
        left.kind === "text" ? left.node.openingStart : left.node.start;
      const rightStart =
        right.kind === "text" ? right.node.openingStart : right.node.start;
      return leftStart - rightStart;
    });
  }
  for (
    let line = 0;
    line < Math.min(lines.length, existingLineCount);
    line += 1
  ) {
    if (lines[line] && !segmentsByLine.has(line)) {
      return null;
    }
  }
  const extraLines = lines.slice(existingLineCount);
  const appendSegment = extraLines.length
    ? segmentsByLine.get(existingLineCount - 1)?.at(-1)
    : undefined;
  if (extraLines.length > 0 && !appendSegment) {
    return null;
  }
  const assignments = new Map<TextNode, string>();
  const tabAssignments = new Map<TabNode, string>();
  for (const [line, segments] of segmentsByLine) {
    const characters = [...(lines[line] ?? "")];
    const originalLengths = segments.map((segment) =>
      segment.kind === "text" ? [...segment.node.text].length : 1
    );
    const originalLength = originalLengths.reduce(
      (total, length) => total + length,
      0
    );
    const assign = (segment: ContentSegment, text: string): void => {
      if (segment.kind === "text") {
        assignments.set(segment.node, text);
      } else {
        tabAssignments.set(segment.node, text);
      }
    };
    if (originalLength === 0) {
      const [first, ...rest] = segments;
      if (first) {
        assign(first, characters.join(""));
      }
      for (const segment of rest) {
        assign(segment, "");
      }
      continue;
    }
    let offset = 0;
    for (const [index, segment] of segments.entries()) {
      const end =
        index === segments.length - 1
          ? characters.length
          : Math.min(characters.length, offset + (originalLengths[index] ?? 0));
      assign(segment, characters.slice(offset, end).join(""));
      offset = end;
    }
  }
  const patches: { end: number; replacement: string; start: number }[] = [];
  for (const paragraph of removedParagraphRanges) {
    patches.push({
      end: paragraph.end,
      replacement: "",
      start: paragraph.start,
    });
  }
  if (lines.length < existingLineCount) {
    for (const lineBreak of lineBreaks) {
      if (
        lineBreak.line >= lines.length - 1 &&
        !removedParagraphs.has(lineBreak.paragraph)
      ) {
        patches.push({
          end: lineBreak.end,
          replacement: "",
          start: lineBreak.start,
        });
      }
    }
  }
  for (const node of keptTextNodes) {
    const assigned = assignments.get(node) ?? "";
    const appended =
      appendSegment?.kind === "text" && node === appendSegment.node
        ? extraLines
        : [];
    if (node.selfClosing && !assigned && appended.length === 0) {
      continue;
    }
    const opening = nativeTextOpeningTag(node.openingTag, assigned);
    const prefix = node.name.includes(":")
      ? node.name.slice(0, node.name.lastIndexOf(":") + 1)
      : "";
    let replacement = nativeTextContentXml(node.name, opening, assigned);
    for (const line of appended) {
      replacement += `</${node.name}><${prefix}br/>${nativeTextOpeningTag(node.openingTag, line)}${nativeTextContentXml(node.name, opening, line)}`;
    }
    if (node.selfClosing) {
      patches.push({
        end: node.tagEnd,
        replacement: `${opening}${replacement}</${node.name}>`,
        start: node.openingStart,
      });
    } else {
      if (opening !== node.openingTag) {
        patches.push({
          end: node.start,
          replacement: opening,
          start: node.openingStart,
        });
      }
      if (replacement !== xml.slice(node.start, node.contentEnd)) {
        patches.push({
          end: node.contentEnd,
          replacement,
          start: node.start,
        });
      }
    }
  }
  for (const node of keptTabNodes) {
    const assigned = tabAssignments.get(node) ?? "";
    const appended =
      appendSegment?.kind === "tab" && node === appendSegment.node
        ? extraLines
        : [];
    if (assigned === "\t" && appended.length === 0) {
      continue;
    }
    const prefix = node.name.includes(":")
      ? node.name.slice(0, node.name.lastIndexOf(":") + 1)
      : "";
    if (!assigned && appended.length === 0) {
      patches.push({ end: node.end, replacement: "", start: node.start });
      continue;
    }
    const name = `${prefix}t`;
    const opening = nativeTextOpeningTag(
      `<${name} xml:space="preserve">`,
      assigned
    );
    let replacement = `${opening}${nativeTextContentXml(name, opening, assigned)}</${name}>`;
    for (const line of appended) {
      const lineOpening = nativeTextOpeningTag(opening, line);
      replacement += `<${prefix}br/>${lineOpening}${nativeTextContentXml(name, lineOpening, line)}</${name}>`;
    }
    patches.push({ end: node.end, replacement, start: node.start });
  }
  return applyNativeXmlPatches(xml, patches);
}
