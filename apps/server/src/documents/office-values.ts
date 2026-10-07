import { FieldType } from "@onlyoffice/db";

// oxlint-disable func-style prefer-destructuring no-await-in-loop no-use-before-define no-nested-ternary complexity no-shadow -- Route modules keep declaration order and sequential persistence invariants.
import { jsonRecord } from "../http/input";
import type { JsonRecord, PrefillSnapshot } from "../model-types";
import type { OperationMetadata } from "../operations/model";
import { isValidDateFieldValue } from "../responses/date-value";
import { templateControlAttribute } from "./fields";
import { nativeCheckboxState } from "./native-xml";
import type { NativeCheckboxState } from "./native-xml";
import {
  parseTemplateXml,
  reachableTemplateControlParts,
  safeTemplateArchive,
  templateArchiveText,
  templateCheckboxNamespace,
  templateMarkupCompatibilityNamespace,
  templateWordNamespaces,
} from "./package";

export interface OnlyOfficeFieldDisplayMetadata {
  calendar: string | null;
  checkedState: NativeCheckboxState | null;
  dateFormat: string | null;
  dateLanguage: string | null;
  uncheckedState: NativeCheckboxState | null;
}

export function onlyOfficeFieldDisplayMetadata(
  documentBytes: Uint8Array
): Map<string, OnlyOfficeFieldDisplayMetadata> {
  const { archive, xmlPaths } = safeTemplateArchive(documentBytes);
  const metadataByTag = new Map<string, OnlyOfficeFieldDisplayMetadata>();
  for (const archivePath of reachableTemplateControlParts(archive, xmlPaths)) {
    const controls: {
      dateFormat: string | null;
      dateLanguage: string | null;
      calendar: string | null;
      checkedState: NativeCheckboxState | null;
      inPropertiesDepth: number;
      tag: string | null;
      uncheckedState: NativeCheckboxState | null;
    }[] = [];
    let alternateFallbackDepth = 0;
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
        } else if (frame.inPropertiesDepth > 0) {
          frame.inPropertiesDepth -= 1;
        } else if (
          element.local === "sdt" &&
          templateWordNamespaces.has(element.uri)
        ) {
          controls.pop();
          if (frame.tag) {
            metadataByTag.set(frame.tag, {
              calendar: frame.calendar,
              checkedState: frame.checkedState,
              dateFormat: frame.dateFormat,
              dateLanguage: frame.dateLanguage,
              uncheckedState: frame.uncheckedState,
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
            calendar: null,
            checkedState: null,
            dateFormat: null,
            dateLanguage: null,
            inPropertiesDepth: 0,
            tag: null,
            uncheckedState: null,
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
        if (frame.inPropertiesDepth === 0) {
          return;
        }
        if (
          frame.inPropertiesDepth === 1 &&
          element.local === "tag" &&
          templateWordNamespaces.has(element.uri)
        ) {
          frame.tag = templateControlAttribute(element, "val")?.trim() ?? null;
        } else if (
          element.local === "checkedState" &&
          element.uri === templateCheckboxNamespace
        ) {
          frame.checkedState = nativeCheckboxState(element);
        } else if (
          element.local === "uncheckedState" &&
          element.uri === templateCheckboxNamespace
        ) {
          frame.uncheckedState = nativeCheckboxState(element);
        } else if (
          element.local === "dateFormat" &&
          templateWordNamespaces.has(element.uri)
        ) {
          frame.dateFormat = templateControlAttribute(element, "val") ?? null;
        } else if (
          element.local === "lid" &&
          templateWordNamespaces.has(element.uri)
        ) {
          frame.dateLanguage = templateControlAttribute(element, "val") ?? null;
        } else if (
          element.local === "calendar" &&
          templateWordNamespaces.has(element.uri)
        ) {
          frame.calendar = templateControlAttribute(element, "val") ?? null;
        }
        frame.inPropertiesDepth += 1;
      },
    });
  }
  return metadataByTag;
}

export function onlyOfficeFieldTagAliases(
  documentBytes: Uint8Array
): Record<string, string> {
  const { archive, xmlPaths } = safeTemplateArchive(documentBytes);
  const aliases = new Map<string, string>();
  const ambiguousAliases = new Set<string>();
  const fieldTags = new Set<string>();

  for (const archivePath of reachableTemplateControlParts(archive, xmlPaths)) {
    let controlDepth = 0;
    let propertiesDepth = 0;
    parseTemplateXml(templateArchiveText(archive, archivePath), {
      close: (element) => {
        if (
          element.local === "sdtPr" &&
          templateWordNamespaces.has(element.uri)
        ) {
          propertiesDepth = 0;
          return;
        }
        if (propertiesDepth > 0) {
          propertiesDepth -= 1;
        }
        if (
          element.local === "sdt" &&
          templateWordNamespaces.has(element.uri)
        ) {
          controlDepth -= 1;
        }
      },
      open: (element) => {
        if (
          element.local === "sdt" &&
          templateWordNamespaces.has(element.uri)
        ) {
          controlDepth += 1;
          return;
        }
        if (controlDepth === 0) {
          return;
        }
        if (
          element.local === "sdtPr" &&
          templateWordNamespaces.has(element.uri)
        ) {
          propertiesDepth = 1;
          return;
        }
        if (propertiesDepth === 0) {
          return;
        }
        if (
          propertiesDepth === 1 &&
          element.local === "tag" &&
          templateWordNamespaces.has(element.uri)
        ) {
          const fieldTag = templateControlAttribute(element, "val")?.trim();
          if (fieldTag) {
            fieldTags.add(fieldTag);
            for (const attribute of element.attributes) {
              if (attribute.local !== "val" || attribute.uri === element.uri) {
                continue;
              }
              const alias = attribute.value.trim();
              if (!alias || alias === fieldTag) {
                continue;
              }
              const previous = aliases.get(alias);
              if (previous && previous !== fieldTag) {
                aliases.delete(alias);
                ambiguousAliases.add(alias);
              } else if (!ambiguousAliases.has(alias)) {
                aliases.set(alias, fieldTag);
              }
            }
          }
        }
        propertiesDepth += 1;
      },
    });
  }

  return Object.fromEntries(
    [...aliases].filter(
      ([alias]) => !fieldTags.has(alias) && !ambiguousAliases.has(alias)
    )
  );
}

interface OnlyOfficeDateNames {
  monthsLong: string[];
  monthsShort: string[];
  weekdaysLong: string[];
  weekdaysShort: string[];
}

const onlyOfficeDateNames = new Map<string, OnlyOfficeDateNames>();

function utcDate(year: number, month: number, day: number): Date {
  return new Date(Date.UTC(year, month, day));
}

function onlyOfficeDateNamesFor(locale: string): OnlyOfficeDateNames | null {
  const cached = onlyOfficeDateNames.get(locale);
  if (cached) {
    return cached;
  }
  try {
    const names = {
      monthsLong: Array.from({ length: 12 }, (_, month) =>
        new Intl.DateTimeFormat(locale, {
          month: "long",
          timeZone: "UTC",
        }).format(utcDate(2024, month, 1))
      ),
      monthsShort: Array.from({ length: 12 }, (_, month) =>
        new Intl.DateTimeFormat(locale, {
          month: "short",
          timeZone: "UTC",
        }).format(utcDate(2024, month, 1))
      ),
      weekdaysLong: Array.from({ length: 7 }, (_, day) =>
        new Intl.DateTimeFormat(locale, {
          timeZone: "UTC",
          weekday: "long",
        }).format(utcDate(2024, 0, 7 + day))
      ),
      weekdaysShort: Array.from({ length: 7 }, (_, day) =>
        new Intl.DateTimeFormat(locale, {
          timeZone: "UTC",
          weekday: "short",
        }).format(utcDate(2024, 0, 7 + day))
      ),
    };
    onlyOfficeDateNames.set(locale, names);
    return names;
  } catch {
    return null;
  }
}

function escapeDateMaskLiteral(value: string): string {
  return value.replaceAll(/[.*+?^${}()|[\]\\]/gu, "\\$&");
}

function onlyOfficeDateDisplayValue(
  value: string,
  metadata: OnlyOfficeFieldDisplayMetadata | undefined
): string | null {
  const format = metadata?.dateFormat;
  const calendar = metadata?.calendar;
  if (
    !format ||
    /(?:^|[^y])yy(?:[^y]|$)/u.test(
      format.replaceAll(/'[^']*'|"[^"]*"/gu, "")
    ) ||
    (calendar !== null &&
      calendar !== undefined &&
      calendar !== "gregorian" &&
      calendar !== "thai")
  ) {
    return null;
  }
  const locale = metadata?.dateLanguage?.split(/[-_]/u)[0] || "en";
  const names = onlyOfficeDateNamesFor(locale);
  if (!names) {
    return null;
  }
  const captures: string[] = [];
  const tokenPattern = /'[^']*'|"[^"]*"|dddd|ddd|yyyy|MMMM|MMM|yy|MM|dd|M|d/gu;
  let pattern = "^";
  let offset = 0;
  for (const match of format.matchAll(tokenPattern)) {
    const index = match.index ?? 0;
    pattern += escapeDateMaskLiteral(format.slice(offset, index));
    const token = match[0];
    offset = index + token.length;
    if (token.startsWith("'") || token.startsWith('"')) {
      pattern += escapeDateMaskLiteral(token.slice(1, -1));
      continue;
    }
    if (
      token === "yyyy" ||
      token === "yy" ||
      token === "dd" ||
      token === "d" ||
      token === "MM" ||
      token === "M"
    ) {
      captures.push(token);
      pattern +=
        token === "yyyy"
          ? "(\\d{4})"
          : token === "yy"
            ? "(\\d{2})"
            : token === "dd" || token === "MM"
              ? "(\\d{2})"
              : "(\\d{1,2})";
      continue;
    }
    const tokenNames =
      token === "MMMM"
        ? names.monthsLong
        : token === "MMM"
          ? names.monthsShort
          : token === "dddd"
            ? names.weekdaysLong
            : token === "ddd"
              ? names.weekdaysShort
              : [];
    if (tokenNames.length > 0) {
      captures.push(token);
      pattern += `(${tokenNames
        .toSorted((left, right) => right.length - left.length)
        .map(escapeDateMaskLiteral)
        .join("|")})`;
    } else {
      pattern += escapeDateMaskLiteral(token);
    }
  }
  pattern += `${escapeDateMaskLiteral(format.slice(offset))}$`;
  const result = new RegExp(pattern, "iu").exec(value);
  if (!result) {
    return null;
  }
  let year: number | null = null;
  let month: number | null = null;
  let day: number | null = null;
  const weekdayValues: { token: string; value: string }[] = [];
  for (const [index, token] of captures.entries()) {
    const capture = result[index + 1];
    if (capture === undefined) {
      return null;
    }
    if (token === "yyyy") {
      year = Number(capture);
    } else if (token === "yy") {
      year = 2000 + Number(capture);
    } else if (token === "MM" || token === "M") {
      month = Number(capture);
    } else if (token === "dd" || token === "d") {
      day = Number(capture);
    } else if (token === "MMMM" || token === "MMM") {
      const monthNames =
        token === "MMMM" ? names.monthsLong : names.monthsShort;
      month =
        monthNames.findIndex(
          (name) => name.toLowerCase() === capture.toLowerCase()
        ) + 1;
    } else {
      weekdayValues.push({ token, value: capture });
    }
  }
  if (year === null || month === null || day === null) {
    return null;
  }
  if (calendar === "thai") {
    year -= 543;
  }
  const isoValue = `${String(year).padStart(4, "0")}-${String(month).padStart(2, "0")}-${String(day).padStart(2, "0")}`;
  if (!isValidDateFieldValue(isoValue)) {
    return null;
  }
  if (weekdayValues.length > 0) {
    const weekday = new Date(`${isoValue}T00:00:00Z`).getUTCDay();
    if (
      weekdayValues.some(({ token, value: actual }) => {
        const weekdays =
          token === "dddd" ? names.weekdaysLong : names.weekdaysShort;
        return weekdays[weekday]?.toLowerCase() !== actual.toLowerCase();
      })
    ) {
      return null;
    }
  }
  return isoValue;
}

export function onlyOfficeFieldValueNeedsMetadata(
  field: { options: unknown; tag: string; type: FieldType },
  value: unknown
): boolean {
  if (typeof value !== "string") {
    return false;
  }
  if (field.type === FieldType.checkbox) {
    return true;
  }
  if (field.type === FieldType.date) {
    return true;
  }
  if (field.type !== FieldType.dropdown && field.type !== FieldType.combo) {
    return false;
  }
  const options = Array.isArray(field.options)
    ? field.options.flatMap((option) => {
        if (!option || typeof option !== "object" || Array.isArray(option)) {
          return [];
        }
        const record = option as Record<string, unknown>;
        return typeof record.displayText === "string" &&
          typeof record.value === "string"
          ? [{ displayText: record.displayText, value: record.value }]
          : [];
      })
    : [];
  return (
    options.some(
      (option) => option.displayText === value && option.value !== value
    ) || value === ""
  );
}

export function normalizeOnlyOfficeDisplayValue(
  field: { options: unknown; tag: string; type: FieldType },
  value: unknown,
  metadata: OnlyOfficeFieldDisplayMetadata | undefined
): unknown {
  if (typeof value !== "string") {
    return value;
  }
  if (field.type === FieldType.checkbox) {
    if (value === "") {
      return false;
    }
    const checkedGlyph = metadata?.checkedState?.glyph ?? "☒";
    const uncheckedGlyph = metadata?.uncheckedState?.glyph ?? "☐";
    if (checkedGlyph !== uncheckedGlyph) {
      if (value === checkedGlyph) {
        return true;
      }
      if (value === uncheckedGlyph) {
        return false;
      }
    }
    return value;
  }
  if (field.type === FieldType.date) {
    if (value === "") {
      return null;
    }
    return onlyOfficeDateDisplayValue(value, metadata) ?? value;
  }
  if (field.type !== FieldType.dropdown && field.type !== FieldType.combo) {
    return value;
  }
  const options = Array.isArray(field.options)
    ? field.options.flatMap((option) => {
        if (!option || typeof option !== "object" || Array.isArray(option)) {
          return [];
        }
        const record = option as Record<string, unknown>;
        return typeof record.displayText === "string" &&
          typeof record.value === "string"
          ? [{ displayText: record.displayText, value: record.value }]
          : [];
      })
    : [];
  if (value === "") {
    return null;
  }
  const selectedOption = options.find((option) => option.displayText === value);
  return selectedOption?.value ?? value;
}

export function callbackFieldMetadata(
  data: JsonRecord,
  suppliedData: unknown,
  snapshot: PrefillSnapshot | null
): Pick<OperationMetadata, "callbackEmptyFieldTags" | "serverHeldFieldTags"> {
  const suppliedFields = jsonRecord(suppliedData);
  const lockedFields = jsonRecord(snapshot?.lockedFields);
  const tags = Object.keys(data);
  return {
    callbackEmptyFieldTags: tags.filter((tag) => suppliedFields[tag] === ""),
    serverHeldFieldTags: tags.filter(
      (tag) => lockedFields[tag] === true || !Object.hasOwn(suppliedFields, tag)
    ),
  };
}
