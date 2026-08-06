const path = require("path")

const MAX_IMAGE_BYTES = 8 * 1024 * 1024
const IMAGE_DOWNLOAD_TIMEOUT_MS = 10000
const SUPPORTED_IMAGE_TYPES = new Set([
  "image/png",
  "image/jpeg",
  "image/gif",
  "image/webp"
])

class EventImageError extends Error {
  constructor(message) {
    super(message)
    this.name = "EventImageError"
  }
}

function normalizeContentType(value) {
  return String(value || "").split(";", 1)[0].trim().toLowerCase()
}

function validateImageSignature(contentType, data) {
  if (contentType === "image/png") {
    return data.length >= 8 && data.subarray(0, 8).equals(
      Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a])
    )
  }
  if (contentType === "image/jpeg") {
    return data.length >= 3 && data[0] === 0xff && data[1] === 0xd8 && data[2] === 0xff
  }
  if (contentType === "image/gif") {
    return data.length >= 6 && ["GIF87a", "GIF89a"].includes(data.subarray(0, 6).toString("ascii"))
  }
  if (contentType === "image/webp") {
    return data.length >= 12
      && data.subarray(0, 4).toString("ascii") === "RIFF"
      && data.subarray(8, 12).toString("ascii") === "WEBP"
  }
  return false
}

function validateAttachmentMetadata(attachment) {
  const contentType = normalizeContentType(attachment?.contentType)
  if (!SUPPORTED_IMAGE_TYPES.has(contentType)) {
    throw new EventImageError("Image must be PNG, JPEG, GIF or WebP.")
  }
  if (!Number.isInteger(attachment.size) || attachment.size <= 0) {
    throw new EventImageError("Image size is invalid.")
  }
  if (attachment.size > MAX_IMAGE_BYTES) {
    throw new EventImageError("Image must be 8 MB or smaller.")
  }
  let attachmentUrl
  try {
    attachmentUrl = new URL(attachment.url)
  } catch {
    throw new EventImageError("Image upload URL is invalid.")
  }
  const discordHosts = ["discord.com", "discordapp.com", "discordapp.net"]
  const trustedHost = discordHosts.some(host =>
    attachmentUrl.hostname === host || attachmentUrl.hostname.endsWith(`.${host}`)
  )
  if (attachmentUrl.protocol !== "https:" || !trustedHost) {
    throw new EventImageError("Image upload URL is not a trusted Discord attachment URL.")
  }

  const originalFilename = path.basename(String(attachment.name || "event-image"))
    .slice(0, 255)
  return { contentType, originalFilename, attachmentUrl: attachmentUrl.toString() }
}

async function readBoundedBody(response, maximumBytes) {
  if (!response.body?.getReader) {
    const data = Buffer.from(await response.arrayBuffer())
    if (data.length > maximumBytes) throw new EventImageError("Image exceeds the 8 MB limit.")
    return data
  }

  const reader = response.body.getReader()
  const chunks = []
  let total = 0
  while (true) {
    const { done, value } = await reader.read()
    if (done) break
    total += value.byteLength
    if (total > maximumBytes) {
      await reader.cancel().catch(() => {})
      throw new EventImageError("Image exceeds the 8 MB limit.")
    }
    chunks.push(Buffer.from(value))
  }
  return Buffer.concat(chunks, total)
}

async function downloadEventImage(
  attachment,
  {
    fetchImpl = globalThis.fetch,
    timeoutMs = IMAGE_DOWNLOAD_TIMEOUT_MS,
    maximumBytes = MAX_IMAGE_BYTES
  } = {}
) {
  const metadata = validateAttachmentMetadata(attachment)
  const controller = new AbortController()
  const timeout = setTimeout(() => controller.abort(), timeoutMs)

  try {
    const response = await fetchImpl(metadata.attachmentUrl, {
      signal: controller.signal,
      redirect: "error"
    })
    if (!response.ok) throw new EventImageError("Discord image download failed.")

    const responseType = normalizeContentType(response.headers.get("content-type"))
    if (responseType !== metadata.contentType || !SUPPORTED_IMAGE_TYPES.has(responseType)) {
      throw new EventImageError("Downloaded image type did not match the upload.")
    }

    const declaredLengthHeader = response.headers.get("content-length")
    const declaredLength = declaredLengthHeader === null
      ? null
      : Number(declaredLengthHeader)
    if (Number.isFinite(declaredLength) && declaredLength > maximumBytes) {
      throw new EventImageError("Image exceeds the 8 MB limit.")
    }
    if (Number.isFinite(declaredLength) && declaredLength !== attachment.size) {
      throw new EventImageError("Downloaded image size did not match the upload.")
    }

    const imageData = await readBoundedBody(response, maximumBytes)
    if (
      imageData.length === 0
      || imageData.length !== attachment.size
      || (Number.isFinite(declaredLength) && imageData.length !== declaredLength)
    ) {
      throw new EventImageError("Downloaded image size did not match the upload.")
    }
    if (!validateImageSignature(metadata.contentType, imageData)) {
      throw new EventImageError("Image content did not match its declared format.")
    }

    return {
      contentType: metadata.contentType,
      originalFilename: metadata.originalFilename,
      byteSize: imageData.length,
      imageData
    }
  } catch (error) {
    if (error instanceof EventImageError) throw error
    if (controller.signal.aborted || error?.name === "AbortError") {
      throw new EventImageError("Image download timed out.")
    }
    throw new EventImageError("Discord image download failed.")
  } finally {
    clearTimeout(timeout)
  }
}

module.exports = {
  MAX_IMAGE_BYTES,
  IMAGE_DOWNLOAD_TIMEOUT_MS,
  SUPPORTED_IMAGE_TYPES,
  EventImageError,
  normalizeContentType,
  validateImageSignature,
  validateAttachmentMetadata,
  readBoundedBody,
  downloadEventImage
}
