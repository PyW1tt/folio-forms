import { maxTemplateArchiveExpandedBytes } from "../documents/package";
import type {
  ResponsePictureDimensions,
  ResponsePictureManifestField,
} from "../documents/pictures";
import {
  responsePictureImageDimensions,
  invalidResponsePicture,
  validateResponsePictureMediaBytes,
} from "../documents/pictures";
import { fail } from "../http/errors";
import {
  maxTemplateMultipartOverheadBytes,
  readJsonRecord,
  readRequestBytes,
  asRecord,
} from "../http/input";
import type { JsonRecord } from "../model-types";
// oxlint-disable func-style prefer-destructuring no-await-in-loop no-use-before-define no-nested-ternary complexity no-shadow -- Route modules keep declaration order and sequential persistence invariants.
import { maxResponseDataBytes } from "./data";

const maxNativePictureMultipartOverheadBytes =
  maxResponseDataBytes + maxTemplateMultipartOverheadBytes;

export interface NativePictureUpload {
  bytes: Uint8Array;
  dimensions: ResponsePictureDimensions;
}
interface NativeResponseUploadedFile {
  arrayBuffer: () => Promise<ArrayBuffer>;
  size: number;
  type: string;
}
type NativeResponseMultipartEntry = string | NativeResponseUploadedFile;

interface NativeResponseRequest {
  input: JsonRecord;
  pictures: Map<string, NativePictureUpload>;
}

export async function nativeResponseRequest(
  request: Request,
  pictureFields: readonly ResponsePictureManifestField[]
): Promise<NativeResponseRequest> {
  const contentType = request.headers.get("content-type");
  if (!contentType || !/^multipart\/form-data(?:\s*;|$)/iu.test(contentType)) {
    return {
      input: await readJsonRecord(
        request,
        maxNativePictureMultipartOverheadBytes
      ),
      pictures: new Map(),
    };
  }
  const imageByteLimit = pictureFields.reduce(
    (total, field) => total + (field.pictureMaxBytes ?? 0),
    0
  );
  const maximumBodyBytes = Math.min(
    maxTemplateArchiveExpandedBytes + maxNativePictureMultipartOverheadBytes,
    imageByteLimit + maxNativePictureMultipartOverheadBytes
  );
  const requestBytes = await readRequestBytes(
    request,
    maximumBodyBytes,
    "Multipart form data is required"
  );
  const multipartRequest = new Request(request.url, {
    body: requestBytes,
    headers: { "content-type": contentType },
    method: "POST",
  });
  let entries: IterableIterator<[string, NativeResponseMultipartEntry]>;
  try {
    const multipartFields = await multipartRequest.formData();
    entries = multipartFields.entries();
  } catch {
    fail(400, "invalid_request", "Multipart form data is invalid");
  }
  const fieldsByTag = new Map(pictureFields.map((field) => [field.tag, field]));
  const pictures = new Map<string, NativePictureUpload>();
  let input: JsonRecord | null = null;
  for (const [key, value] of entries) {
    if (key === "payload") {
      if (input || typeof value !== "string") {
        fail(400, "invalid_request", "payload must be provided once");
      }
      try {
        input = asRecord(JSON.parse(value));
      } catch {
        fail(400, "invalid_request", "payload must be a valid JSON object");
      }
      continue;
    }
    if (!key.startsWith("picture:") || typeof value === "string") {
      fail(
        400,
        "invalid_request",
        "Only payload and picture files are accepted"
      );
    }
    const tag = key.slice("picture:".length);
    const field = fieldsByTag.get(tag);
    if (!field || pictures.has(tag)) {
      fail(400, "invalid_request", "Picture field is unknown or duplicated");
    }
    if (value.size > (field.pictureMaxBytes ?? 0)) {
      fail(
        413,
        "payload_too_large",
        `Picture field ${tag} exceeds its byte limit`
      );
    }
    const fileType = value.type.trim().toLowerCase();
    if (fileType && fileType !== "image/jpeg" && fileType !== "image/png") {
      fail(
        415,
        "invalid_file_type",
        `Picture field ${tag} must be JPEG or PNG`
      );
    }
    const bytes = new Uint8Array(await value.arrayBuffer());
    const dimensions = responsePictureImageDimensions(bytes);
    if (
      !dimensions ||
      (fileType && fileType !== `image/${dimensions.format}`)
    ) {
      invalidResponsePicture(tag, "image must be a valid JPEG or PNG");
    }
    validateResponsePictureMediaBytes(tag, bytes, field, dimensions);
    pictures.set(tag, { bytes, dimensions });
  }
  if (!input) {
    fail(400, "invalid_request", "payload is required");
  }
  return { input, pictures };
}
