function boundedInteger(value, fallback, minimum, maximum) {
  const parsed = Number(value)
  return Number.isInteger(parsed) && parsed >= minimum && parsed <= maximum ? parsed : fallback
}

function normalizedHeaders(headers = {}) {
  const result = {}
  for (const [key, value] of Object.entries(headers || {})) {
    result[String(key).toLowerCase()] = Array.isArray(value) ? value[0] : value
  }
  return result
}

function retryAfterMilliseconds(headers, nowMs = Date.now()) {
  const value = normalizedHeaders(headers)["retry-after"]
  if (value === undefined || value === null || value === "") return null
  const seconds = Number(value)
  if (Number.isFinite(seconds) && seconds >= 0) return Math.ceil(seconds * 1000)
  const date = Date.parse(String(value))
  return Number.isFinite(date) ? Math.max(0, date - nowMs) : null
}

function rateLimitObservation({ gameProfile, endpointType, httpStatus, headers, requestedAt, respondedAt }) {
  const normalized = normalizedHeaders(headers)
  return Object.freeze({
    gameProfile,
    endpointType,
    httpStatus: Number(httpStatus) || null,
    limit: normalized["x-ratelimit-limit"] ?? null,
    remaining: normalized["x-ratelimit-remaining"] ?? null,
    reset: normalized["x-ratelimit-reset"] ?? null,
    retryAfter: normalized["retry-after"] ?? null,
    requestedAt: new Date(requestedAt),
    respondedAt: new Date(respondedAt)
  })
}

class ConservativeRateLimiter {
  constructor({
    gameProfile,
    minimumDelayMs = 1000,
    maximumRetries = 2,
    baseBackoffMs = 2000,
    maximumBackoffMs = 60000,
    maximumObservations = 100,
    now = () => Date.now(),
    sleep = delay => new Promise(resolve => setTimeout(resolve, delay))
  }) {
    this.gameProfile = gameProfile
    this.minimumDelayMs = boundedInteger(minimumDelayMs, 1000, 0, 60000)
    this.maximumRetries = boundedInteger(maximumRetries, 2, 0, 5)
    this.baseBackoffMs = boundedInteger(baseBackoffMs, 2000, 100, 60000)
    this.maximumBackoffMs = boundedInteger(maximumBackoffMs, 60000, 1000, 15 * 60000)
    this.maximumObservations = boundedInteger(maximumObservations, 100, 1, 1000)
    this.now = now
    this.sleep = sleep
    this.tail = Promise.resolve()
    this.lastStartedAt = null
    this.observations = []
  }

  recordObservation(result, endpointType, requestedAt, respondedAt) {
    this.observations.push(rateLimitObservation({
      gameProfile: this.gameProfile,
      endpointType,
      httpStatus: result.httpStatus,
      headers: result.headers,
      requestedAt,
      respondedAt
    }))
    if (this.observations.length > this.maximumObservations) this.observations.shift()
  }

  async execute(operation, { endpointType = "gift_code", shouldRetry = result => result.retryable } = {}) {
    let attempt = 0
    while (true) {
      if (this.lastStartedAt !== null) {
        const spacing = this.minimumDelayMs - (this.now() - this.lastStartedAt)
        if (spacing > 0) await this.sleep(spacing)
      }
      const requestedAt = this.now()
      this.lastStartedAt = requestedAt
      const result = await operation(attempt + 1)
      const respondedAt = this.now()
      this.recordObservation(result, endpointType, requestedAt, respondedAt)

      if (!shouldRetry(result) || attempt >= this.maximumRetries) {
        return { ...result, attempts: attempt + 1 }
      }
      const retryAfter = retryAfterMilliseconds(result.headers, respondedAt)
      const exponential = Math.min(
        this.maximumBackoffMs,
        this.baseBackoffMs * (2 ** attempt)
      )
      await this.sleep(Math.max(retryAfter ?? 0, exponential))
      attempt += 1
    }
  }

  schedule(operation, options) {
    const scheduled = this.tail.then(
      () => this.execute(operation, options),
      () => this.execute(operation, options)
    )
    this.tail = scheduled.then(() => undefined, () => undefined)
    return scheduled
  }

  getObservations() {
    return [...this.observations]
  }
}

function rateLimiterConfig(env = process.env) {
  return {
    minimumDelayMs: boundedInteger(env.CENTURY_MINIMUM_DELAY_MS, 1000, 0, 60000),
    maximumRetries: boundedInteger(env.CENTURY_MAXIMUM_RETRIES, 2, 0, 5),
    baseBackoffMs: boundedInteger(env.CENTURY_BACKOFF_BASE_MS, 2000, 100, 60000),
    maximumBackoffMs: boundedInteger(env.CENTURY_BACKOFF_CAP_MS, 60000, 1000, 15 * 60000)
  }
}

module.exports = {
  boundedInteger,
  normalizedHeaders,
  retryAfterMilliseconds,
  rateLimitObservation,
  ConservativeRateLimiter,
  rateLimiterConfig
}
