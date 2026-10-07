import { FieldType } from "@onlyoffice/db";

// oxlint-disable func-style prefer-destructuring no-await-in-loop no-use-before-define no-nested-ternary complexity no-shadow -- Route modules keep declaration order and sequential persistence invariants.
import type { GeneratedTemplate } from "../ai-authoring/document";
import { fail } from "../http/errors";
import {
  parseTemplateXml,
  reachableTemplateControlParts,
  resolveTemplateRelationshipTarget,
  safeTemplateArchive,
  templateArchiveText,
  templateAttribute,
  templateCheckboxNamespace,
  templateControlKey,
  templateControlNamespaces,
  templateGlossaryDocumentRelationships,
  templateMarkupCompatibilityNamespace,
  templatePackageRelationshipNamespace,
  templateRelationshipPartPath,
  templateWord2012Namespace,
  templateWordMainNamespace,
  templateWordNamespaces,
  templateWordStrictNamespace,
} from "./package";
import type { TemplateXmlElement } from "./package";

export interface ParsedTemplateField {
  label: string;
  options: { displayText: string; value: string }[] | null;
  pictureMaxBytes: number | null;
  pictureMaxHeight: number | null;
  pictureMaxWidth: number | null;
  placeholder: string | null;
  tag: string;
  type: FieldType;
}

interface TemplateControlFrame {
  alias: string | null;
  documentOrder: number;
  duplicateMarker: boolean;
  duplicateTag: boolean;
  inContentDepth: number;
  inPropertiesDepth: number;
  markers: Set<FieldType>;
  options: { displayText: string; value: string }[];
  placeholderName: string | null;
  placeholderText: string;
  propertyStack: string[];
  showingPlaceholder: boolean;
  tag: string | null;
  unsupported: string | null;
}

const templateControlTypeMarkers = new Map<string, FieldType>([
  [templateControlKey(templateWordMainNamespace, "comboBox"), FieldType.combo],
  [templateControlKey(templateWordMainNamespace, "date"), FieldType.date],
  [
    templateControlKey(templateWordMainNamespace, "dropDownList"),
    FieldType.dropdown,
  ],
  [templateControlKey(templateWordMainNamespace, "picture"), FieldType.picture],
  [templateControlKey(templateWordMainNamespace, "text"), FieldType.text],
  [
    templateControlKey(templateWordStrictNamespace, "comboBox"),
    FieldType.combo,
  ],
  [templateControlKey(templateWordStrictNamespace, "date"), FieldType.date],
  [
    templateControlKey(templateWordStrictNamespace, "dropDownList"),
    FieldType.dropdown,
  ],
  [
    templateControlKey(templateWordStrictNamespace, "picture"),
    FieldType.picture,
  ],
  [templateControlKey(templateWordStrictNamespace, "text"), FieldType.text],
  [
    templateControlKey(templateCheckboxNamespace, "checkbox"),
    FieldType.checkbox,
  ],
]);

const templateUnsupportedControlMarkers = new Set([
  templateControlKey(templateWordMainNamespace, "citation"),
  templateControlKey(templateWordMainNamespace, "docPartGallery"),
  templateControlKey(templateWordMainNamespace, "docPartList"),
  templateControlKey(templateWordMainNamespace, "docPartObj"),
  templateControlKey(templateWordMainNamespace, "equation"),
  templateControlKey(templateWordMainNamespace, "group"),
  templateControlKey(templateWordMainNamespace, "richText"),
  templateControlKey(templateWord2012Namespace, "repeatingSection"),
  templateControlKey(templateWord2012Namespace, "repeatingSectionItem"),
]);

const templateControlTypeLocals = new Set([
  "checkbox",
  "comboBox",
  "date",
  "dropDownList",
  "picture",
  "text",
]);

const templateUnsupportedControlLocals = new Set([
  "citation",
  "docPartGallery",
  "docPartList",
  "docPartObj",
  "equation",
  "group",
  "richText",
  "repeatingSection",
  "repeatingSectionItem",
]);

const templateControlMetadataProperties = new Set([
  "alias",
  "appearance",
  "calendar",
  "checked",
  "checkedState",
  "color",
  "dataBinding",
  "dateFormat",
  "formPr",
  "id",
  "lock",
  "placeholder",
  "lid",
  "rPr",
  "showingPlcHdr",
  "tag",
  "temporary",
  "uncheckedState",
]);

export function templateControlAttribute(
  element: TemplateXmlElement,
  local: string
): string | undefined {
  return element.attributes.find(
    (attribute) => attribute.local === local && attribute.uri === element.uri
  )?.value;
}

function parsedTemplateField(
  frame: TemplateControlFrame,
  glossaryPlaceholders: ReadonlyMap<string, string>
): ParsedTemplateField {
  if (!frame.tag?.trim()) {
    fail(422, "invalid_template", "Every content control must have a tag");
  }
  if (frame.unsupported) {
    fail(
      422,
      "invalid_template",
      `Unsupported content control type: ${frame.unsupported}`
    );
  }
  if (frame.duplicateTag) {
    fail(
      422,
      "invalid_template",
      "A content control cannot repeat its tag property"
    );
  }
  if (frame.duplicateMarker) {
    fail(
      422,
      "invalid_template",
      "A content control cannot repeat a field type marker"
    );
  }
  if (frame.markers.size > 1) {
    fail(
      422,
      "invalid_template",
      "A content control cannot declare multiple field types"
    );
  }
  const type = frame.markers.values().next().value ?? FieldType.text;
  if (
    (type === FieldType.dropdown || type === FieldType.combo) &&
    frame.options.length === 0
  ) {
    fail(
      422,
      "invalid_template",
      "Dropdown and combo fields must define options"
    );
  }
  if (
    type !== FieldType.dropdown &&
    type !== FieldType.combo &&
    frame.options.length > 0
  ) {
    fail(
      422,
      "invalid_template",
      "Only dropdown and combo fields may define options"
    );
  }
  const uniqueOptionLabels = new Set(
    frame.options.map((option) => option.displayText)
  );
  const uniqueOptionValues = new Set(
    frame.options.map((option) => option.value)
  );
  if (
    uniqueOptionLabels.size !== frame.options.length ||
    uniqueOptionValues.size !== frame.options.length
  ) {
    fail(422, "invalid_template", "Dropdown and combo options must be unique");
  }
  const tag = frame.tag.trim();
  const label = frame.alias?.trim();
  const placeholder =
    glossaryPlaceholders.get(frame.placeholderName ?? "")?.trim() ||
    (frame.showingPlaceholder ? frame.placeholderText.trim() : "");
  return {
    label: label || tag,
    options: frame.options.length > 0 ? frame.options : null,
    pictureMaxBytes: type === FieldType.picture ? 10 * 1024 * 1024 : null,
    pictureMaxHeight: type === FieldType.picture ? 4096 : null,
    pictureMaxWidth: type === FieldType.picture ? 4096 : null,
    placeholder: placeholder || null,
    tag,
    type,
  };
}

function templateGlossaryPlaceholders(
  archive: Record<string, Uint8Array>
): Map<string, string> {
  const relationshipsPath = templateRelationshipPartPath("word/document.xml");
  if (!archive[relationshipsPath]) {
    return new Map();
  }
  let glossaryPath: string | null = null;
  parseTemplateXml(templateArchiveText(archive, relationshipsPath), {
    open: (element) => {
      if (
        element.local === "Relationship" &&
        element.uri === templatePackageRelationshipNamespace &&
        templateGlossaryDocumentRelationships.has(
          templateAttribute(element, "Type") ?? ""
        ) &&
        templateAttribute(element, "TargetMode") === undefined
      ) {
        glossaryPath = resolveTemplateRelationshipTarget(
          "word/document.xml",
          templateAttribute(element, "Target")
        );
      }
    },
  });
  if (!glossaryPath || !archive[glossaryPath]) {
    return new Map();
  }
  const placeholders = new Map<string, string>();
  const docParts: {
    bodyDepth: number;
    name: string | null;
    text: string;
  }[] = [];
  const elementStack: string[] = [];
  parseTemplateXml(templateArchiveText(archive, glossaryPath), {
    close: (element) => {
      const docPart = docParts.at(-1);
      if (
        docPart &&
        element.local === "docPartBody" &&
        templateWordNamespaces.has(element.uri)
      ) {
        docPart.bodyDepth -= 1;
      }
      if (
        element.local === "docPart" &&
        templateWordNamespaces.has(element.uri)
      ) {
        const completed = docParts.pop();
        if (
          completed?.name &&
          completed.text.trim() &&
          !placeholders.has(completed.name)
        ) {
          placeholders.set(completed.name, completed.text.trim());
        }
      }
      elementStack.pop();
    },
    open: (element) => {
      if (
        element.local === "docPart" &&
        templateWordNamespaces.has(element.uri)
      ) {
        docParts.push({ bodyDepth: 0, name: null, text: "" });
      }
      const docPart = docParts.at(-1);
      if (!docPart) {
        elementStack.push(element.local);
        return;
      }
      if (
        elementStack.at(-1) === "docPartPr" &&
        element.local === "docPartName" &&
        templateWordNamespaces.has(element.uri)
      ) {
        docPart.name = templateControlAttribute(element, "val") ?? null;
      }
      if (
        element.local === "docPartBody" &&
        templateWordNamespaces.has(element.uri)
      ) {
        docPart.bodyDepth += 1;
      }
      elementStack.push(element.local);
    },
    text: (value) => {
      const docPart = docParts.at(-1);
      if (docPart && docPart.bodyDepth > 0) {
        docPart.text += value;
      }
    },
  });
  return placeholders;
}

export function parseTemplateFields(
  bytes: Uint8Array,
  content?: { paragraphs: string[]; fieldText: Map<string, string> }
): ParsedTemplateField[] {
  const { archive, xmlPaths } = safeTemplateArchive(bytes);
  const controlPaths = reachableTemplateControlParts(archive, xmlPaths);
  const fields: { documentOrder: number; field: ParsedTemplateField }[] = [];
  const glossaryPlaceholders = templateGlossaryPlaceholders(archive);
  let documentOrderOffset = 0;
  for (const archivePath of controlPaths) {
    const documentOrderBase = documentOrderOffset;
    let documentOrder = 0;
    let alternateFallbackDepth = 0;
    const controls: TemplateControlFrame[] = [];
    const paragraphs: { text: string }[] = [];
    let textDepth = 0;
    let propertyDepth = 0;
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
        if (content && templateWordNamespaces.has(element.uri)) {
          if (element.local === "t") {
            textDepth -= 1;
          }
          if (element.local.endsWith("Pr")) {
            propertyDepth -= 1;
          }
          if (element.local === "p") {
            const text = paragraphs.pop()?.text.trim();
            if (text) {
              content.paragraphs.push(text);
            }
          }
        }
        const frame = controls.at(-1);
        if (!frame) {
          return;
        }
        if (
          element.local === "sdtContent" &&
          templateWordNamespaces.has(element.uri)
        ) {
          frame.inContentDepth -= 1;
          return;
        }
        if (
          element.local === "sdtPr" &&
          templateWordNamespaces.has(element.uri)
        ) {
          frame.inPropertiesDepth = 0;
          frame.propertyStack.length = 0;
          return;
        }
        if (frame.inPropertiesDepth > 0) {
          frame.propertyStack.pop();
          frame.inPropertiesDepth -= 1;
        }
        if (
          element.local === "sdt" &&
          templateWordNamespaces.has(element.uri)
        ) {
          controls.pop();
          if (content && frame.tag) {
            content.fieldText.set(
              frame.tag.trim(),
              frame.placeholderText.trim()
            );
          }
          fields.push({
            documentOrder: documentOrderBase + frame.documentOrder,
            field: parsedTemplateField(frame, glossaryPlaceholders),
          });
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
        if (content && templateWordNamespaces.has(element.uri)) {
          if (element.local === "t") {
            textDepth += 1;
          }
          if (element.local.endsWith("Pr")) {
            propertyDepth += 1;
          }
          if (element.local === "p") {
            paragraphs.push({ text: "" });
          }
        }
        if (
          element.local === "sdt" &&
          templateWordNamespaces.has(element.uri)
        ) {
          const controlOrder = documentOrderBase + documentOrder;
          documentOrder += 1;
          controls.push({
            alias: null,
            documentOrder: controlOrder,
            duplicateMarker: false,
            duplicateTag: false,
            inContentDepth: 0,
            inPropertiesDepth: 0,
            markers: new Set(),
            options: [],
            placeholderName: null,
            placeholderText: "",
            propertyStack: [],
            showingPlaceholder: false,
            tag: null,
            unsupported: null,
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
        if (
          element.local === "sdtContent" &&
          templateWordNamespaces.has(element.uri)
        ) {
          frame.inContentDepth += 1;
          return;
        }
        if (frame.inPropertiesDepth === 0) {
          return;
        }
        const isControlElement = templateControlNamespaces.has(element.uri);
        const isDirectProperty = frame.inPropertiesDepth === 1;
        const parentProperty = frame.propertyStack.at(-1);
        if (parentProperty === "listItem") {
          fail(422, "invalid_template", "Field options are malformed");
        }
        const controlKey = templateControlKey(element.uri, element.local);
        const marker = templateControlTypeMarkers.get(controlKey);
        const isUnsupportedControl =
          templateUnsupportedControlMarkers.has(controlKey);
        const isTag =
          templateWordNamespaces.has(element.uri) && element.local === "tag";
        const isListItem =
          templateWordNamespaces.has(element.uri) &&
          element.local === "listItem";
        const isAllowedMetadata =
          (templateWordNamespaces.has(element.uri) &&
            templateControlMetadataProperties.has(element.local)) ||
          (element.uri === templateWord2012Namespace &&
            element.local === "appearance") ||
          (element.uri === templateCheckboxNamespace &&
            ["checked", "checkedState", "uncheckedState"].includes(
              element.local
            ));
        const isWrongNamespaceSemantic =
          (element.local === "tag" && !isTag) ||
          (templateControlTypeLocals.has(element.local) &&
            marker === undefined) ||
          (templateUnsupportedControlLocals.has(element.local) &&
            !isUnsupportedControl);
        if (
          (isTag ||
            marker !== undefined ||
            isUnsupportedControl ||
            isWrongNamespaceSemantic) &&
          !isDirectProperty
        ) {
          fail(
            422,
            "invalid_template",
            `Misplaced content control property: ${element.local}`
          );
        }
        if (isTag) {
          if (frame.tag !== null) {
            frame.duplicateTag = true;
          }
          frame.tag = templateControlAttribute(element, "val") ?? null;
        }
        if (
          isDirectProperty &&
          templateWordNamespaces.has(element.uri) &&
          element.local === "alias"
        ) {
          frame.alias = templateControlAttribute(element, "val") ?? null;
        }
        if (
          frame.inPropertiesDepth === 2 &&
          parentProperty === "placeholder" &&
          templateWordNamespaces.has(element.uri) &&
          element.local === "docPart"
        ) {
          frame.placeholderName =
            templateControlAttribute(element, "val") ?? null;
        }
        if (
          isDirectProperty &&
          templateWordNamespaces.has(element.uri) &&
          element.local === "showingPlcHdr"
        ) {
          const value = templateControlAttribute(element, "val")
            ?.trim()
            .toLowerCase();
          frame.showingPlaceholder =
            value === undefined || !["0", "false", "off"].includes(value);
        }
        if (marker !== undefined) {
          if (frame.markers.has(marker)) {
            frame.duplicateMarker = true;
          }
          frame.markers.add(marker);
        }
        if (isUnsupportedControl) {
          frame.unsupported = element.local;
        }
        if (element.local === "listItem") {
          if (
            !isListItem ||
            frame.inPropertiesDepth !== 2 ||
            (parentProperty !== "comboBox" && parentProperty !== "dropDownList")
          ) {
            fail(422, "invalid_template", "Field options are malformed");
          }
          if (
            !frame.markers.has(FieldType.dropdown) &&
            !frame.markers.has(FieldType.combo)
          ) {
            fail(
              422,
              "invalid_template",
              "Only dropdown and combo fields may define options"
            );
          }
          const displayText = templateControlAttribute(element, "displayText");
          const value = templateControlAttribute(element, "value");
          if (!displayText?.trim() || value === undefined) {
            fail(422, "invalid_template", "Field options are malformed");
          }
          frame.options.push({
            displayText: displayText.trim(),
            value,
          });
        }
        if (
          isDirectProperty &&
          (!isControlElement ||
            (marker === undefined &&
              !isUnsupportedControl &&
              !isAllowedMetadata))
        ) {
          fail(
            422,
            "invalid_template",
            `Unknown content control type: ${element.local}`
          );
        }
        frame.propertyStack.push(element.local);
        frame.inPropertiesDepth += 1;
      },
      text: (value) => {
        const frame = controls.at(-1);
        if (content) {
          if (
            alternateFallbackDepth > 0 ||
            textDepth === 0 ||
            propertyDepth > 0
          ) {
            return;
          }
          if (
            frame &&
            frame.inContentDepth > 0 &&
            frame.inPropertiesDepth === 0
          ) {
            frame.placeholderText += value;
          } else if (controls.length === 0) {
            const paragraph = paragraphs.at(-1);
            if (paragraph) {
              paragraph.text += value;
            }
          }
        } else if (frame?.showingPlaceholder && frame.inContentDepth > 0) {
          frame.placeholderText += value;
        }
      },
    });
    documentOrderOffset += documentOrder;
  }

  if (fields.length === 0) {
    fail(
      422,
      "invalid_template",
      "The template must contain at least one tagged content control"
    );
  }
  const orderedFields = fields
    .toSorted((left, right) => left.documentOrder - right.documentOrder)
    .map(({ field }) => field);
  const duplicates = orderedFields.filter(
    (field, index) =>
      orderedFields.findIndex((candidate) => candidate.tag === field.tag) !==
      index
  );
  if (duplicates.length > 0) {
    fail(
      422,
      "invalid_template",
      `Content control tags must be unique: ${[...new Set(duplicates.map((field) => field.tag))].join(", ")}`
    );
  }
  return orderedFields;
}

export function responseTextControlValues(
  bytes: Uint8Array
): Map<string, string> {
  const { archive, xmlPaths } = safeTemplateArchive(bytes);
  const values = new Map<string, string>();
  for (const archivePath of reachableTemplateControlParts(archive, xmlPaths)) {
    const controls: {
      contentDepth: number;
      inPropertiesDepth: number;
      paragraphCount: number;
      showingPlaceholder: boolean;
      tag: string | null;
      text: string;
    }[] = [];
    let alternateFallbackDepth = 0;
    let textDepth = 0;
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
        if (element.local === "t" && templateWordNamespaces.has(element.uri)) {
          textDepth -= 1;
          return;
        }
        const frame = controls.at(-1);
        if (!frame) {
          return;
        }
        if (
          element.local === "sdtContent" &&
          templateWordNamespaces.has(element.uri)
        ) {
          frame.contentDepth -= 1;
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
          if (frame.tag && !frame.showingPlaceholder) {
            values.set(frame.tag, frame.text);
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
        if (element.local === "t" && templateWordNamespaces.has(element.uri)) {
          textDepth += 1;
          return;
        }
        if (
          element.local === "sdt" &&
          templateWordNamespaces.has(element.uri)
        ) {
          controls.push({
            contentDepth: 0,
            inPropertiesDepth: 0,
            paragraphCount: 0,
            showingPlaceholder: false,
            tag: null,
            text: "",
          });
          return;
        }
        const frame = controls.at(-1);
        if (!frame) {
          return;
        }
        if (element.local === "p" && templateWordNamespaces.has(element.uri)) {
          for (const control of controls) {
            if (
              control.contentDepth === 0 ||
              control.inPropertiesDepth > 0 ||
              control.showingPlaceholder
            ) {
              continue;
            }
            if (control.paragraphCount > 0) {
              control.text += "\n";
            }
            control.paragraphCount += 1;
          }
          return;
        }
        if (
          (element.local === "tab" ||
            element.local === "br" ||
            element.local === "cr") &&
          templateWordNamespaces.has(element.uri)
        ) {
          const separator = element.local === "tab" ? "\t" : "\n";
          for (const control of controls) {
            if (
              control.contentDepth > 0 &&
              control.inPropertiesDepth === 0 &&
              !control.showingPlaceholder
            ) {
              control.text += separator;
            }
          }
          return;
        }
        if (
          element.local === "sdtPr" &&
          templateWordNamespaces.has(element.uri)
        ) {
          frame.inPropertiesDepth = 1;
          return;
        }
        if (
          element.local === "sdtContent" &&
          templateWordNamespaces.has(element.uri)
        ) {
          frame.contentDepth += 1;
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
        frame.inPropertiesDepth += 1;
      },
      text: (value) => {
        if (textDepth === 0) {
          return;
        }
        for (const frame of controls) {
          if (frame.contentDepth > 0 && !frame.showingPlaceholder) {
            frame.text += value;
          }
        }
      },
    });
  }
  return values;
}

export function validateTemplateControls(
  bytes: Uint8Array,
  expected?: Omit<GeneratedTemplate, "document">
): string[] {
  const content = expected
    ? { fieldText: new Map<string, string>(), paragraphs: [] as string[] }
    : undefined;
  const fields = parseTemplateFields(bytes, content);
  if (expected && content) {
    const remaining = new Map<string, number>();
    for (const text of [
      expected.title,
      expected.description,
      ...expected.paragraphs,
    ]) {
      if (text) {
        remaining.set(text, (remaining.get(text) ?? 0) + 1);
      }
    }
    for (const text of content.paragraphs) {
      const count = remaining.get(text) ?? 0;
      if (count > 0) {
        remaining.set(text, count - 1);
      } else if (
        !expected.fields.some(
          (field) => text === field.label || text === `${field.label}:`
        )
      ) {
        throw new Error(
          "Generated DOCX static text does not match its metadata"
        );
      }
    }
    if ([...remaining.values()].some((count) => count > 0)) {
      throw new Error("Generated DOCX static text does not match its metadata");
    }
    if (
      fields.some((field) => {
        const declared = expected.fields.find(
          (candidate) => candidate.tag === field.tag
        );
        return (
          !declared ||
          field.type !== FieldType.text ||
          field.label !== declared.label ||
          content.fieldText.get(field.tag) !== declared.placeholder
        );
      })
    ) {
      throw new Error("Generated DOCX controls do not match its metadata");
    }
  }
  return fields.map((field) => field.tag);
}
