const axios = require("axios")
const {
  parseActiveCatalogueHtml,
  parseWosActiveCatalogueHtml
} = require("./sourceParsers")
const { safeSourceError } = require("./sourceIngestion")

const CATALOGUES = Object.freeze({
  wos: {
    name: "WOSRewards",
    url: "https://www.wosrewards.com/"
  },
  kingshot: {
    name: "KingshotRewards",
    url: "https://kingshotrewards.com/"
  }
})

function createCatalogueAdapter({
  gameProfile,
  transport = axios,
  timeoutMs = 10000,
  maximumBodyBytes = 1024 * 1024
}) {
  const catalogue = CATALOGUES[gameProfile]
  if (!catalogue) throw new Error("Unsupported game profile")
  return Object.freeze({
    ...catalogue,
    async fetchActiveCodes() {
      const response = await transport.get(catalogue.url, {
        timeout: timeoutMs,
        responseType: "text",
        maxContentLength: maximumBodyBytes,
        maxBodyLength: maximumBodyBytes,
        headers: {
          "User-Agent": "RACHIE-GiftCodeDiscovery/1.0 (+provider-neutral public catalogue poller)"
        },
        validateStatus: status => status >= 200 && status < 300
      })
      const body = String(response.data || "")
      if (Buffer.byteLength(body, "utf8") > maximumBodyBytes) {
        throw Object.assign(new Error("catalogue response exceeded configured limit"), {
          code: "SOURCE_BODY_TOO_LARGE"
        })
      }
      const codes = gameProfile === "wos"
        ? parseWosActiveCatalogueHtml(body)
        : parseActiveCatalogueHtml(body)
      if (!codes.length) {
        const contentType = String(response.headers?.["content-type"] || "")
          .slice(0, 100) || null
        throw Object.assign(
          new Error("active catalogue markup was not recognised"),
          {
            code: "SOURCE_MARKUP_UNRECOGNISED",
            sourceDiagnostics: {
              httpStatus: Number(response.status) || null,
              contentType,
              responseBytes: Buffer.byteLength(body, "utf8"),
              expectedStructure: gameProfile === "wos"
                ? "active_gift_codes_section"
                : "explicit_active_code_entry"
            }
          }
        )
      }
      return codes
    }
  })
}

function createCataloguePoller({
  gameProfile,
  sourceRepository,
  ingestion,
  adapter,
  enabled = false,
  logger = console,
  now = () => new Date()
}) {
  let activePoll = null

  async function poll() {
    if (!enabled) return { polled: false, reason: "disabled" }
    if (activePoll) return activePoll
    activePoll = (async () => {
      const polledAt = now()
      let source = null
      try {
        source = await sourceRepository.ensureSource({
          sourceType: "public_catalogue",
          sourceName: adapter.name,
          sourceReference: adapter.url,
          trusted: false
        })
        const codes = await adapter.fetchActiveCodes()
        let candidates = 0
        for (const code of codes) {
          const result = await ingestion.ingest({
            source: {
              sourceType: "public_catalogue",
              sourceName: adapter.name,
              sourceReference: adapter.url,
              trusted: false
            },
            code,
            observationKey: `catalogue:${code}`,
            observedAt: polledAt,
            provenance: { transport: "http_catalogue", catalogue: adapter.name }
          })
          if (!result.duplicate) candidates += 1
        }
        await sourceRepository.markCataloguePoll(source.id, {
          successful: true,
          observedCodes: codes,
          now: polledAt
        })
        return { polled: true, observed: codes.length, candidates }
      } catch (error) {
        const errorCode = safeSourceError(error)
        if (source) {
          await sourceRepository.markCataloguePoll(source.id, {
            successful: false,
            errorCode,
            now: polledAt
          }).catch(() => {})
        }
        logger.warn(JSON.stringify({
          event: "gift_code_catalogue_poll_failed",
          game_profile: gameProfile,
          source: adapter.name,
          error_code: errorCode,
          http_status: error.sourceDiagnostics?.httpStatus,
          content_type: error.sourceDiagnostics?.contentType,
          response_bytes: error.sourceDiagnostics?.responseBytes,
          expected_structure: error.sourceDiagnostics?.expectedStructure
        }))
        return { polled: true, failed: true, errorCode }
      }
    })().finally(() => { activePoll = null })
    return activePoll
  }

  return Object.freeze({ poll, enabled })
}

module.exports = { CATALOGUES, createCatalogueAdapter, createCataloguePoller }
