const test = require("node:test")
const assert = require("node:assert/strict")

const {
  DateTimeValidationError,
  parseUtcTime,
  parseIsoDate,
  parseTimeOrGroups
} = require("../src/timeParsing")
const {
  EventValidationError,
  validateEventDraft
} = require("../src/eventValidation")
const {
  MAX_IMAGE_BYTES,
  EventImageError,
  validateAttachmentMetadata,
  downloadEventImage
} = require("../src/eventImage")
const {
  InteractionSessionError,
  InteractionSessionStore
} = require("../src/interactionSessions")
const {
  createEventSchedulerRepository
} = require("../src/eventSchedulerRepository")
const {
  EVENTS_PER_PAGE,
  formatEventEntry,
  formatEventListPage
} = require("../src/eventSchedulerFormatting")

const validDraft = Object.freeze({
  allianceName: "North",
  eventName: "Bear Hunt",
  firstOccurrenceDate: "2026-08-01",
  firstDateIsPast: true,
  eventTimeUtc: "18:30",
  groups: [],
  grouped: false,
  recurrenceDays: 7,
  advanceReminderMinutes: 10,
  reminderAtStart: true,
  publishToAlliance: true,
  publishToState: false,
  includeInWeeklyRoundup: true,
  image: null
})

test("time parser accepts and normalizes documented examples", () => {
  const examples = new Map([
    ["18:30", "18:30"],
    ["1830", "18:30"],
    ["1800", "18:00"],
    ["6:30pm", "18:30"],
    ["6.30 PM", "18:30"],
    ["6 PM", "18:00"],
    ["18", "18:00"]
  ])
  for (const [input, expected] of examples) assert.equal(parseUtcTime(input), expected)
})

test("time parser rejects invalid and ambiguous values", () => {
  for (const input of ["", "24:00", "18:60", "630", "6:3", "noon", "18pm", "00am"]) {
    assert.throws(() => parseUtcTime(input), DateTimeValidationError, input)
  }
})

test("date parser validates calendar dates without locale parsing", () => {
  const now = new Date("2026-08-06T12:00:00Z")
  assert.equal(parseIsoDate("2026-08-07", now).value, "2026-08-07")
  assert.equal(parseIsoDate("2024-02-29", now).date.toISOString(), "2024-02-29T00:00:00.000Z")
  assert.equal(parseIsoDate("2026-01-01", now).isPast, true)
  assert.throws(() => parseIsoDate("2025-02-29", now), DateTimeValidationError)
  assert.throws(() => parseIsoDate("2026-04-31", now), DateTimeValidationError)
  assert.throws(() => parseIsoDate("06/08/2026", now), DateTimeValidationError)
})

test("time-or-groups parser supports ungrouped and unique grouped events", () => {
  assert.deepEqual(parseTimeOrGroups("18"), { eventTimeUtc: "18:00", groups: [] })
  assert.deepEqual(parseTimeOrGroups("Alpha = 18:30\nBeta = 6 PM"), {
    eventTimeUtc: null,
    groups: [
      { groupName: "Alpha", eventTimeUtc: "18:30", sortOrder: 0 },
      { groupName: "Beta", eventTimeUtc: "18:00", sortOrder: 1 }
    ]
  })
  assert.throws(() => parseTimeOrGroups("Alpha = 18:00\nalpha = 19:00"), DateTimeValidationError)
  assert.throws(() => parseTimeOrGroups("Alpha = 18:00\n19:00"), DateTimeValidationError)
})

test("event validation accepts valid grouped and ungrouped drafts", () => {
  assert.equal(validateEventDraft(validDraft).eventTimeUtc, "18:30")
  const grouped = validateEventDraft({
    ...validDraft,
    eventTimeUtc: null,
    grouped: true,
    groups: [{ groupName: "Alpha", eventTimeUtc: "18:30", sortOrder: 0 }]
  })
  assert.equal(grouped.groups.length, 1)
})

test("event validation rejects structural and option errors", () => {
  const invalidDrafts = [
    { ...validDraft, eventTimeUtc: null },
    { ...validDraft, eventTimeUtc: null, grouped: true, groups: [] },
    {
      ...validDraft,
      eventTimeUtc: null,
      grouped: true,
      groups: [{ groupName: "Alpha", eventTimeUtc: null }]
    },
    {
      ...validDraft,
      eventTimeUtc: null,
      grouped: true,
      groups: [
        { groupName: "Alpha", eventTimeUtc: "18:00" },
        { groupName: "alpha", eventTimeUtc: "19:00" }
      ]
    },
    { ...validDraft, recurrenceDays: 5 },
    { ...validDraft, advanceReminderMinutes: 15 },
    { ...validDraft, publishToState: true }
  ]
  for (const draft of invalidDrafts) {
    assert.throws(() => validateEventDraft(draft), EventValidationError)
  }
  assert.equal(
    validateEventDraft({ ...validDraft, publishToState: true }, { stateLinkEnabled: true })
      .publishToState,
    true
  )
})

test("image download validates metadata, signature and bounded response", async () => {
  const png = Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a])
  const attachment = {
    contentType: "image/png",
    size: png.length,
    name: "event.png",
    url: "https://cdn.discordapp.com/attachments/test/signed"
  }
  const image = await downloadEventImage(attachment, {
    fetchImpl: async () => new Response(png, {
      status: 200,
      headers: { "content-type": "image/png" }
    })
  })
  assert.equal(image.byteSize, png.length)
  assert.equal(image.contentType, "image/png")
  assert.ok(image.imageData.equals(png))
  assert.equal(validDraft.image, null)

  assert.throws(
    () => validateAttachmentMetadata({ ...attachment, contentType: "text/plain" }),
    EventImageError
  )
  assert.throws(
    () => validateAttachmentMetadata({ ...attachment, size: MAX_IMAGE_BYTES + 1 }),
    EventImageError
  )
  assert.throws(
    () => validateAttachmentMetadata({ ...attachment, url: "https://example.invalid/image.png" }),
    /trusted Discord/
  )
})

test("image download enforces timeout and actual response size", async () => {
  const attachment = {
    contentType: "image/png",
    size: 8,
    name: "event.png",
    url: "https://cdn.discordapp.com/attachments/test/signed"
  }
  await assert.rejects(
    downloadEventImage(attachment, {
      timeoutMs: 5,
      fetchImpl: async (_url, { signal }) => new Promise((_resolve, reject) => {
        signal.addEventListener("abort", () => {
          const error = new Error("aborted")
          error.name = "AbortError"
          reject(error)
        })
      })
    }),
    /timed out/
  )

  const oversized = Buffer.concat([
    Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]),
    Buffer.from([0x00])
  ])
  await assert.rejects(
    downloadEventImage(attachment, {
      fetchImpl: async () => new Response(oversized, {
        status: 200,
        headers: { "content-type": "image/png" }
      })
    }),
    /size did not match/
  )
  await assert.rejects(
    downloadEventImage(attachment, {
      fetchImpl: async () => new Response(
        Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]),
        {
          status: 200,
          headers: { "content-type": "image/png", "content-length": "9" }
        }
      )
    }),
    /size did not match/
  )
  await assert.rejects(
    downloadEventImage({ ...attachment, size: oversized.length }, {
      maximumBytes: 8,
      fetchImpl: async () => new Response(oversized, {
        status: 200,
        headers: { "content-type": "image/png" }
      })
    }),
    /8 MB limit/
  )
})

function makeTransactionPool({ failOn } = {}) {
  const calls = []
  const client = {
    async query(text, values) {
      calls.push({ text, values })
      if (failOn && text.includes(failOn)) throw new Error("injected failure")
      if (text.includes("INSERT INTO scheduled_events")) return { rows: [{ id: "41" }] }
      return { rows: [] }
    },
    release() {
      calls.push({ text: "RELEASE" })
    }
  }
  return {
    calls,
    query: client.query.bind(client),
    async connect() {
      return client
    }
  }
}

function createInput(overrides = {}) {
  return {
    guildId: "guild-1",
    createdByUserId: "user-1",
    createdByBotInstance: "rachie-wos",
    event: { ...validDraft, ...overrides }
  }
}

test("event creation is atomic and stores groups and optional image", async () => {
  const pool = makeTransactionPool()
  const image = {
    originalFilename: "event.png",
    contentType: "image/png",
    byteSize: 8,
    imageData: Buffer.alloc(8)
  }
  await createEventSchedulerRepository(pool, "wos").createEvent(createInput({
    eventTimeUtc: null,
    groups: [{ groupName: "Alpha", eventTimeUtc: "18:30", sortOrder: 0 }],
    image
  }))
  assert.equal(pool.calls[0].text, "BEGIN")
  assert.ok(pool.calls.some(call => call.text.includes("INSERT INTO scheduled_event_groups")))
  assert.ok(pool.calls.some(call => call.text.includes("INSERT INTO scheduled_event_images")))
  assert.ok(pool.calls.some(call => call.text === "COMMIT"))
  assert.ok(!pool.calls.some(call => call.text === "ROLLBACK"))
})

test("group and image failures roll back the complete event", async () => {
  for (const failOn of ["INSERT INTO scheduled_event_groups", "INSERT INTO scheduled_event_images"]) {
    const pool = makeTransactionPool({ failOn })
    await assert.rejects(
      createEventSchedulerRepository(pool, "wos").createEvent(createInput({
        eventTimeUtc: null,
        groups: [{ groupName: "Alpha", eventTimeUtc: "18:30", sortOrder: 0 }],
        image: {
          originalFilename: "event.png",
          contentType: "image/png",
          byteSize: 8,
          imageData: Buffer.alloc(8)
        }
      })),
      /injected failure/
    )
    assert.ok(pool.calls.some(call => call.text === "ROLLBACK"), failOn)
    assert.ok(!pool.calls.some(call => call.text === "COMMIT"), failOn)
  }
})

test("event writes and reads remain scoped to guild and game profile", async () => {
  const calls = []
  const pool = {
    async query(text, values) {
      calls.push({ text, values })
      if (text.includes("FROM scheduled_events e") && text.includes("COUNT(*)")) {
        return { rows: [] }
      }
      if (text.includes("LEFT JOIN scheduled_event_images")) return { rows: [] }
      return { rows: [] }
    }
  }
  await createEventSchedulerRepository(pool, "wos").listEvents("same-guild")
  await createEventSchedulerRepository(pool, "kingshot").listEvents("same-guild")
  await createEventSchedulerRepository(pool, "wos").getEvent("same-guild", "1")
  await createEventSchedulerRepository(pool, "kingshot").getEvent("same-guild", "1")

  assert.deepEqual(calls[0].values.slice(0, 2), ["same-guild", "wos"])
  assert.deepEqual(calls[1].values.slice(0, 2), ["same-guild", "kingshot"])
  assert.deepEqual(calls[2].values, ["1", "same-guild", "wos"])
  assert.deepEqual(calls[3].values, ["1", "same-guild", "kingshot"])
})

test("creation sessions reject wrong ownership and expiry", () => {
  let now = 1000
  const store = new InteractionSessionStore({ ttlMs: 100, now: () => now })
  const context = { userId: "u1", guildId: "g1", gameProfile: "wos" }
  const id = store.create(context, { eventName: "Test" })
  assert.throws(() => store.get(id, { ...context, userId: "u2" }), InteractionSessionError)
  assert.throws(() => store.get(id, { ...context, guildId: "g2" }), InteractionSessionError)
  assert.throws(() => store.get(id, { ...context, gameProfile: "kingshot" }), InteractionSessionError)
  now = 1100
  assert.throws(() => store.get(id, context), InteractionSessionError)
})

test("cancel and successful completion clear creation sessions", () => {
  const store = new InteractionSessionStore()
  const context = { userId: "u1", guildId: "g1", gameProfile: "wos" }
  const cancelled = store.create(context)
  store.cancel(cancelled, context)
  assert.throws(() => store.get(cancelled, context), InteractionSessionError)

  const completed = store.create(context)
  store.complete(completed, context)
  assert.throws(() => store.get(completed, context), InteractionSessionError)
})

test("event listing formats grouped, ungrouped, active and paused entries safely", () => {
  const ungrouped = {
    event_name: "Bear Hunt",
    status: "active",
    alliance_name: "North",
    first_occurrence_date: "2026-08-10",
    event_time_utc: "18:30:00",
    groups: [],
    recurrence_days: 7,
    advance_reminder_minutes: 10,
    reminder_at_start: true,
    publish_to_alliance: true,
    publish_to_state: false,
    include_in_weekly_roundup: true,
    has_image: true
  }
  const grouped = {
    ...ungrouped,
    event_name: "Foundry",
    status: "paused",
    event_time_utc: null,
    groups: [{ group_name: "Alpha", event_time_utc: "20:00:00" }]
  }
  assert.match(formatEventEntry(ungrouped), /18:30 UTC/)
  assert.match(formatEventEntry(grouped), /Alpha: 20:00 UTC/)
  assert.match(formatEventEntry(grouped), /paused/)

  const page = formatEventListPage([ungrouped, grouped], 0, EVENTS_PER_PAGE + 1)
  assert.match(page, /page 1 of 2/)
  assert.ok(page.length <= 1950)
  assert.doesNotMatch(page, /kingshot-only-event/)
})
