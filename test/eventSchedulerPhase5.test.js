const test = require("node:test")
const assert = require("node:assert/strict")

const {
  getEventDeliveryConfig
} = require("../src/eventDeliveryConfig")
const {
  deliveryWindow,
  buildDeliveryClaims,
  generateMissingDeliveryClaims
} = require("../src/eventDeliveryGeneration")
const {
  mapClaimPayload
} = require("../src/eventDeliveryRepository")
const {
  RETRY_BACKOFF_MINUTES,
  PermanentDeliveryError,
  sanitizeDeliveryError,
  retryDelayMinutes,
  workerStartReason,
  createEventDeliveryWorker
} = require("../src/eventDeliveryWorker")
const {
  shutdownEventSchedulerSubsystem
} = require("../src/eventSchedulerHealth")

const NOW = new Date("2026-08-06T12:00:00Z")

function event(overrides = {}) {
  return {
    id: "41",
    guild_id: "guild-wos",
    game_profile: "wos",
    alliance_name: "North",
    event_name: "Bear Hunt",
    first_occurrence_date: "2026-08-06",
    event_time_utc: "12:10:00",
    recurrence_days: 3,
    advance_reminder_minutes: 10,
    reminder_at_start: true,
    publish_to_alliance: true,
    status: "active",
    event_channel_id: "channel-wos",
    groups: [],
    ...overrides
  }
}

function claimKey(claim) {
  return [
    claim.eventId,
    claim.groupId || "",
    claim.gameProfile,
    claim.occurrenceAt.toISOString(),
    claim.deliveryKind,
    claim.targetKind,
    claim.targetChannelId
  ].join(":")
}

test("delivery configuration validates values and falls back safely", () => {
  assert.deepEqual(getEventDeliveryConfig({}), {
    lookaheadMinutes: 1440,
    graceMinutes: 60,
    pollIntervalMs: 30000,
    batchSize: 10,
    claimLeaseSeconds: 60,
    handlerTimeoutMs: 30000
  })
  assert.deepEqual(getEventDeliveryConfig({
    EVENT_SCHEDULER_LOOKAHEAD_MINUTES: "60",
    EVENT_SCHEDULER_GRACE_MINUTES: "0",
    EVENT_SCHEDULER_POLL_INTERVAL_MS: "5000",
    EVENT_SCHEDULER_BATCH_SIZE: "20",
    EVENT_SCHEDULER_CLAIM_LEASE_SECONDS: "120",
    EVENT_SCHEDULER_HANDLER_TIMEOUT_MS: "1000"
  }), {
    lookaheadMinutes: 60,
    graceMinutes: 0,
    pollIntervalMs: 5000,
    batchSize: 20,
    claimLeaseSeconds: 120,
    handlerTimeoutMs: 1000
  })
  assert.deepEqual(getEventDeliveryConfig({
    EVENT_SCHEDULER_LOOKAHEAD_MINUTES: "0",
    EVENT_SCHEDULER_GRACE_MINUTES: "-1",
    EVENT_SCHEDULER_POLL_INTERVAL_MS: "4999",
    EVENT_SCHEDULER_BATCH_SIZE: "lots",
    EVENT_SCHEDULER_CLAIM_LEASE_SECONDS: "9999",
    EVENT_SCHEDULER_HANDLER_TIMEOUT_MS: "0"
  }), getEventDeliveryConfig({}))
})

test("generation creates advance and one-minute final claims for ungrouped events", () => {
  const window = deliveryWindow(NOW, { lookaheadMinutes: 1440, graceMinutes: 60 })
  const both = buildDeliveryClaims([event()], {
    gameProfile: "wos",
    windowStart: window.start,
    windowEnd: window.end
  })
  assert.deepEqual(both.map(item => item.deliveryKind), ["advance_reminder", "final_reminder"])
  assert.equal(both[0].deliverAt.toISOString(), "2026-08-06T12:00:00.000Z")
  assert.equal(both[1].deliverAt.toISOString(), "2026-08-06T12:09:00.000Z")

  const advanceOnly = buildDeliveryClaims([event({ reminder_at_start: false })], {
    gameProfile: "wos", windowStart: window.start, windowEnd: window.end
  })
  assert.deepEqual(advanceOnly.map(item => item.deliveryKind), ["advance_reminder"])
  const finalOnly = buildDeliveryClaims([event({ advance_reminder_minutes: null })], {
    gameProfile: "wos", windowStart: window.start, windowEnd: window.end
  })
  assert.deepEqual(finalOnly.map(item => item.deliveryKind), ["final_reminder"])
  const none = buildDeliveryClaims([event({
    advance_reminder_minutes: null,
    reminder_at_start: false
  })], { gameProfile: "wos", windowStart: window.start, windowEnd: window.end })
  assert.deepEqual(none, [])
})

test("grouped generation creates separate deterministic claims per group", () => {
  const grouped = event({
    event_time_utc: null,
    groups: [
      { group_id: "2", group_name: "Beta", event_time_utc: "12:10:00", sort_order: 1 },
      { group_id: "1", group_name: "Alpha", event_time_utc: "12:10:00", sort_order: 0 }
    ]
  })
  const claims = buildDeliveryClaims([grouped], {
    gameProfile: "wos",
    windowStart: new Date("2026-08-06T11:00:00Z"),
    windowEnd: new Date("2026-08-06T13:00:00Z")
  })
  assert.deepEqual(
    claims.map(item => `${item.groupId}:${item.deliveryKind}`),
    ["1:advance_reminder", "1:final_reminder", "2:advance_reminder", "2:final_reminder"]
  )
})

test("generation excludes inactive, unconfigured, state and wrong-profile targets", () => {
  const options = {
    gameProfile: "wos",
    windowStart: new Date("2026-08-06T11:00:00Z"),
    windowEnd: new Date("2026-08-06T13:00:00Z")
  }
  const excluded = [
    event({ status: "paused" }),
    event({ status: "deleted" }),
    event({ event_channel_id: null }),
    event({ publish_to_alliance: false, publish_to_state: true }),
    event({ game_profile: "kingshot" })
  ]
  assert.deepEqual(buildDeliveryClaims(excluded, options), [])
  assert.ok(buildDeliveryClaims([event()], options).every(item => item.targetKind === "alliance"))
  assert.deepEqual(buildDeliveryClaims([event()], {
    ...options,
    gameProfile: "kingshot"
  }), [])
})

test("generation applies grace and lookahead as a half-open delivery window", () => {
  const events = [
    event({ id: "grace", event_time_utc: "11:30:00", advance_reminder_minutes: null }),
    event({ id: "old", event_time_utc: "10:59:00", advance_reminder_minutes: null }),
    event({ id: "inside", event_time_utc: "11:01:00", advance_reminder_minutes: null }),
    event({ id: "future", first_occurrence_date: "2026-08-07", event_time_utc: "12:01:00", advance_reminder_minutes: null })
  ]
  const window = deliveryWindow(NOW, { lookaheadMinutes: 1440, graceMinutes: 60 })
  const claims = buildDeliveryClaims(events, {
    gameProfile: "wos", windowStart: window.start, windowEnd: window.end
  })
  assert.deepEqual(claims.map(item => item.eventId), ["grace", "inside"])
})

test("historical generation jumps to the bounded window and insertion is idempotent", async () => {
  const stored = new Set()
  const repository = {
    async listActiveEventDefinitions() {
      return [event({
        first_occurrence_date: "1900-01-03",
        event_time_utc: "12:10:00",
        recurrence_days: 3
      })]
    },
    async insertMissingDeliveryClaims(claims) {
      let inserted = 0
      for (const claim of claims) {
        const key = claimKey(claim)
        if (!stored.has(key)) {
          stored.add(key)
          inserted += 1
        }
      }
      return inserted
    }
  }
  const config = { lookaheadMinutes: 1440, graceMinutes: 60 }
  const first = await generateMissingDeliveryClaims({ repository, gameProfile: "wos", now: NOW, config })
  const second = await generateMissingDeliveryClaims({ repository, gameProfile: "wos", now: NOW, config })
  assert.equal(first.generated, 2)
  assert.equal(first.inserted, 2)
  assert.equal(second.inserted, 0)
})

test("claim payload is structured, copied and does not expose image URLs", () => {
  const source = Buffer.from([1, 2, 3])
  const payload = mapClaimPayload({
    claim_id: "5",
    game_profile: "wos",
    attempt_count: 1,
    delivery_kind: "advance_reminder",
    target_kind: "alliance",
    target_guild_id: "guild",
    target_channel_id: "channel",
    occurrence_at: NOW,
    deliver_at: NOW,
    event_id: "41",
    guild_id: "guild",
    alliance_name: "North",
    event_name: "Bear Hunt",
    recurrence_days: 3,
    group_id: "2",
    group_name: "Alpha",
    group_event_time_utc: "12:10:00",
    group_sort_order: 0,
    image_filename: "event.png",
    image_content_type: "image/png",
    image_byte_size: 3,
    image_data: source
  })
  source[0] = 9
  assert.equal(payload.image.imageData[0], 1)
  assert.equal(payload.group.name, "Alpha")
  assert.equal(payload.alliance.name, "North")
  assert.equal(payload.image.url, undefined)
  assert.ok(Object.isFrozen(payload))
})

function fakeWorkerRepository({ claims = [], handlerPayload = {} } = {}) {
  const calls = []
  return {
    gameProfile: "wos",
    calls,
    async listActiveEventDefinitions() {
      calls.push({ method: "list" })
      return []
    },
    async insertMissingDeliveryClaims(items) {
      calls.push({ method: "insert", items })
      return 0
    },
    async claimDueDeliveries(input) {
      calls.push({ method: "claim", input })
      const result = claims.splice(0)
      return result
    },
    async getClaimPayload(input) {
      calls.push({ method: "payload", input })
      return Object.freeze({ claim: { id: input.claimId }, ...handlerPayload })
    },
    async markClaimSent(input) {
      calls.push({ method: "sent", input })
      return true
    },
    async markClaimFailed(input) {
      calls.push({ method: "failed", input })
      return true
    },
    async markClaimPermanentlyFailed(input) {
      calls.push({ method: "permanent", input })
      return true
    }
  }
}

function workerOptions(overrides = {}) {
  return {
    env: { EVENT_SCHEDULER_ENABLED: "true" },
    health: { available: true, gameProfile: "wos" },
    repository: fakeWorkerRepository(),
    gameProfile: "wos",
    botInstanceName: "rachie-wos",
    deliveryHandler: async () => ({ sentMessageId: "synthetic-test-id" }),
    logger: { error() {} },
    now: () => new Date(NOW),
    workerId: "test-worker",
    config: {
      lookaheadMinutes: 1440,
      graceMinutes: 60,
      pollIntervalMs: 5000,
      batchSize: 10,
      claimLeaseSeconds: 60,
      handlerTimeoutMs: 50
    },
    ...overrides
  }
}

test("worker success stores a supplied synthetic message ID", async () => {
  const repository = fakeWorkerRepository({ claims: [{ id: "7", attempt_count: 1 }] })
  const worker = createEventDeliveryWorker(workerOptions({ repository }))
  assert.equal(await worker.tick(), 1)
  const sent = repository.calls.find(call => call.method === "sent")
  assert.equal(sent.input.sentMessageId, "synthetic-test-id")
  assert.equal(sent.input.botInstanceName, "rachie-wos")
  assert.equal(sent.input.workerId, "test-worker")
})

test("retry policy is deterministic, sanitised and stops at five attempts", async () => {
  assert.deepEqual(RETRY_BACKOFF_MINUTES, [1, 5, 15, 30, 60])
  assert.deepEqual([1, 2, 3, 4, 5].map(retryDelayMinutes), [1, 5, 15, 30, 60])
  const sanitized = sanitizeDeliveryError(new Error(
    `DATABASE_URL=postgres://secret\nhttps://cdn.discordapp.com/attachment ${"x".repeat(600)}`
  ))
  assert.doesNotMatch(sanitized, /postgres|discordapp|secret/)
  assert.ok(sanitized.length <= 500)

  for (const [attempt, expectedMethod] of [[1, "failed"], [5, "permanent"]]) {
    const repository = fakeWorkerRepository({ claims: [{ id: String(attempt), attempt_count: attempt }] })
    const worker = createEventDeliveryWorker(workerOptions({
      repository,
      deliveryHandler: async () => { throw new Error("temporary") }
    }))
    await worker.tick()
    const failure = repository.calls.find(call => call.method === expectedMethod)
    assert.ok(failure)
    if (attempt === 1) {
      assert.equal(failure.input.nextAttemptAt.toISOString(), "2026-08-06T12:01:00.000Z")
    }
  }
})

test("permanent failures do not schedule retries and handler timeouts are retryable", async () => {
  const permanentRepository = fakeWorkerRepository({ claims: [{ id: "1", attempt_count: 1 }] })
  await createEventDeliveryWorker(workerOptions({
    repository: permanentRepository,
    deliveryHandler: async () => { throw new PermanentDeliveryError("invalid target") }
  })).tick()
  assert.ok(permanentRepository.calls.some(call => call.method === "permanent"))

  const timeoutRepository = fakeWorkerRepository({ claims: [{ id: "2", attempt_count: 1 }] })
  await createEventDeliveryWorker(workerOptions({
    repository: timeoutRepository,
    deliveryHandler: async () => new Promise(() => {}),
    config: { ...workerOptions().config, handlerTimeoutMs: 5 }
  })).tick()
  const failed = timeoutRepository.calls.find(call => call.method === "failed")
  assert.match(failed.input.lastError, /timed out/)
})

test("worker start gates reject incomplete or unavailable configurations", () => {
  const base = workerOptions()
  const cases = [
    [{ ...base, env: {} }, "disabled"],
    [{ ...base, health: { available: false } }, "database unavailable"],
    [{ ...base, gameProfile: "other" }, "invalid game profile"],
    [{ ...base, gameProfile: "kingshot" }, "health profile mismatch"],
    [{ ...base, botInstanceName: "" }, "missing bot instance"],
    [{ ...base, health: { ...base.health, botInstanceName: "another" } }, "health bot instance mismatch"],
    [{ ...base, repository: null }, "missing delivery repository"],
    [{ ...base, repository: { ...base.repository, gameProfile: "kingshot" } }, "repository profile mismatch"],
    [{ ...base, deliveryHandler: null }, "missing delivery handler"]
  ]
  for (const [input, expected] of cases) {
    assert.equal(workerStartReason(input), expected)
    const worker = createEventDeliveryWorker(input)
    assert.deepEqual(worker.start(), { started: false, reason: expected })
  }
})

test("worker prevents overlapping ticks and survives tick errors", async () => {
  let release
  let listCalls = 0
  const repository = fakeWorkerRepository()
  repository.listActiveEventDefinitions = async () => {
    listCalls += 1
    if (listCalls === 1) await new Promise(resolve => { release = resolve })
    return []
  }
  const worker = createEventDeliveryWorker(workerOptions({ repository }))
  const first = worker.tick()
  const overlapping = worker.tick()
  assert.equal(first, overlapping)
  assert.equal(listCalls, 1)
  release()
  await first
  await worker.tick()
  assert.equal(listCalls, 2)

  let failures = 1
  repository.listActiveEventDefinitions = async () => {
    if (failures-- > 0) throw new Error("database unavailable")
    return []
  }
  assert.equal(await worker.tick(), 0)
  assert.equal(await worker.tick(), 0)
})

test("worker stop clears polling and bounds active-tick shutdown", async () => {
  let intervalCallback
  let cleared = false
  let release
  const repository = fakeWorkerRepository()
  repository.listActiveEventDefinitions = async () => new Promise(resolve => { release = resolve })
  const worker = createEventDeliveryWorker(workerOptions({
    repository,
    setIntervalFn(callback) {
      intervalCallback = callback
      return { unref() {} }
    },
    clearIntervalFn() { cleared = true }
  }))
  assert.equal(worker.start().started, true)
  assert.equal(typeof intervalCallback, "function")
  const result = await worker.stop({ timeoutMs: 5 })
  assert.equal(result.drained, false)
  assert.equal(cleared, true)
  intervalCallback()
  assert.equal(repository.calls.filter(call => call.method === "claim").length, 0)
  release([])
})

test("scheduler shutdown stops the worker before closing the pool path", async () => {
  const calls = []
  const result = await shutdownEventSchedulerSubsystem({
    worker: {
      async stop(input) {
        calls.push(input)
        return { drained: true }
      }
    },
    timeoutMs: 25
  })
  assert.deepEqual(calls, [{ timeoutMs: 25 }])
  assert.deepEqual(result, { workerDrained: true })
})
