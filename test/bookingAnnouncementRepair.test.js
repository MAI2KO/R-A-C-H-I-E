const assert = require("node:assert/strict")
const test = require("node:test")

const { executeInspection, inspectAnnouncementRepair,
  runAnnouncementRepair } = require("../src/bookingAnnouncementRepair")
const { parseArguments } = require("../scripts/repairBookingAnnouncements")

const cutoff = "2026-08-30T12:00:00.000Z"
const token = "r".repeat(43)

function member(id, { admin = false, role = false } = {}) {
  return { id, user: { bot: false }, permissions: { has: () => admin },
    roles: { cache: { has: () => role } } }
}

function fixture({ profile = "wos", messageExists = true, sendFails = false } = {}) {
  const edits = []
  const sends = []
  const existing = { id: "600000000000000001", author: { id: "999" },
    createdTimestamp: new Date("2026-08-30T11:00:00.000Z").getTime(),
    content: `Minister sign-up is now open\nhttps://localhost:8080/booking/member`,
    components: [], async edit(payload) { edits.push(payload); this.content = payload.content
      this.components = payload.components; return this } }
  const channel = { id: "500000000000000001", messages: { async fetch(id) {
    if (typeof id === "object") return new Map(messageExists ? [[existing.id, existing]] : [])
    if (messageExists) return existing
    const error = new Error("deleted"); error.code = 10008; throw error
  } }, async send(payload) { sends.push(payload); if (sendFails) throw Object.assign(
    new Error("no permission"), { code: 50013 })
    return { id: "600000000000000002", ...payload } } }
  const guild = { ownerId: "1", channels: { fetch: async () => channel },
    members: { fetch: async () => new Map([
      ["1", member("1")], ["2", member("2", { admin: true })],
      ["3", member("3", { role: true })], ["4", member("4")]
    ]) } }
  const candidate = { profile, notificationId: "11111111-1111-4111-8111-111111111111",
    communityId: "22222222-2222-4222-8222-222222222222", communityCode: "1234",
    windowId: "33333333-3333-4333-8333-333333333333",
    guestLinkId: "44444444-4444-4444-8444-444444444444", guestLinkHint: "old…hint",
    sentAt: "2026-08-30T11:00:00.000Z", discordChannelId: channel.id,
    discordMessageId: existing.id, guilds: ["700000000000000001"], sentBefore: cutoff }
  return { candidate, channel, guild, edits, sends,
    client: { user: { id: "999" }, guilds: { fetch: async () => guild } },
    setupRepository: { async get() { return { minister_sign_up_channel_id: channel.id,
      bot_manager_role_id: "42" } } } }
}

test("repair arguments require an explicit mode and canonical pre-fix cutoff", () => {
  assert.deepEqual(parseArguments(["--dry-run", "--sent-before", cutoff]),
    { dryRun: true, sentBefore: cutoff })
  assert.deepEqual(parseArguments(["--sent-before", cutoff, "--execute"]),
    { dryRun: false, sentBefore: cutoff })
  assert.throws(() => parseArguments(["--execute"]), /usage/)
  assert.throws(() => parseArguments(["--dry-run", "--sent-before", "2026-08-30"]), /canonical/)
})

test("dry run resolves native managers and destinations without mutation", async () => {
  const f = fixture()
  let began = 0
  const results = await runAnnouncementRepair({ client: f.client,
    api: { announcementRepairCandidates: async value => {
      assert.equal(value, cutoff); return { candidates: [f.candidate] }
    }, async beginAnnouncementRepair() { began++ } }, setupRepository: f.setupRepository,
    dryRun: true, sentBefore: cutoff })
  assert.equal(began, 0)
  assert.equal(f.edits.length, 0)
  assert.equal(f.sends.length, 0)
  assert.equal(results[0].status, "planned")
  assert.equal(results[0].managerCount, 3)
  assert.deepEqual(results[0].destinations, [{ guildId: "700000000000000001",
    channelId: f.channel.id, messageId: "600000000000000001", botAccess: true,
    action: "edit" }])
  assert.equal(JSON.stringify(results).includes(token), false)
})

for (const [profile, origin, place] of [
  ["wos", "https://r-a-c-h-i-e.com", "State"],
  ["kingshot", "https://peggie.r-a-c-h-i-e.com", "Kingdom"],
]) {
  test(`${profile} execution rotates then edits the existing announcement with the public origin`, async () => {
    const f = fixture({ profile })
    const completed = []
    const inspection = await inspectAnnouncementRepair(f.client, f.setupRepository, f.candidate)
    const result = await executeInspection({ baseUrl: origin, allowLoopback: false,
      async beginAnnouncementRepair(id, suppliedCutoff) {
        assert.equal(id, f.candidate.notificationId); assert.equal(suppliedCutoff, cutoff)
        return { repair: { ...f.candidate, profile, closesAt: "2030-09-06T12:30:00Z",
          guestPath: `/book/${token}` } }
      }, async completeAnnouncementRepair(id, message) { completed.push([id, message]) } }, inspection)
    assert.equal(result.status, "completed")
    assert.equal(f.edits.length, 1)
    assert.match(f.edits[0].content, new RegExp(`Guest booking — ${place} 1234`))
    assert.match(f.edits[0].content, new RegExp(`${origin}/book/${token}`))
    assert.doesNotMatch(f.edits[0].content, /localhost|railway|staging|\/booking\/member/i)
    assert.deepEqual(f.edits[0].components, [])
    assert.equal(completed.length, 1)
    assert.equal(JSON.stringify(result).includes(token), false)
  })
}

test("a deleted message is recreated and its new reference is completed", async () => {
  const f = fixture({ messageExists: false, profile: "kingshot" })
  const inspection = await inspectAnnouncementRepair(f.client, f.setupRepository, f.candidate)
  let completed
  const result = await executeInspection({ baseUrl: "https://peggie.r-a-c-h-i-e.com",
    allowLoopback: false, async beginAnnouncementRepair() { return { repair: {
      ...f.candidate, profile: "kingshot", closesAt: "2030-09-06T12:30:00Z",
      guestPath: `/book/${token}` } } }, async completeAnnouncementRepair(_id, message) {
      completed = message
    } }, inspection)
  assert.equal(result.status, "completed")
  assert.equal(f.sends.length, 1)
  assert.equal(f.sends[0].enforceNonce, true)
  assert.deepEqual(completed, { discordChannelId: f.channel.id,
    discordMessageId: "600000000000000002" })
})

test("unsafe preflight skips rotation and a Discord failure remains resumable", async () => {
  const missing = fixture()
  missing.setupRepository.get = async () => ({ minister_sign_up_channel_id: "500" })
  missing.client.guilds.fetch = async () => { throw Object.assign(new Error("missing"), { code: 10004 }) }
  let began = 0
  const skipped = await runAnnouncementRepair({ client: missing.client,
    api: { announcementRepairCandidates: async () => ({ candidates: [missing.candidate] }),
      async beginAnnouncementRepair() { began++ } }, setupRepository: missing.setupRepository,
    dryRun: false, sentBefore: cutoff })
  assert.equal(began, 0)
  assert.equal(skipped[0].status, "skipped")

  const partial = fixture({ messageExists: false, sendFails: true })
  const inspection = await inspectAnnouncementRepair(partial.client, partial.setupRepository,
    partial.candidate)
  let completes = 0
  const result = await executeInspection({ baseUrl: "https://r-a-c-h-i-e.com",
    allowLoopback: false, async beginAnnouncementRepair() { return { repair: {
      ...partial.candidate, closesAt: "2030-09-06T12:30:00Z", guestPath: `/book/${token}` } } },
    async completeAnnouncementRepair() { completes++ } }, inspection)
  assert.equal(result.status, "partial_failure")
  assert.equal(completes, 0)
  assert.equal(JSON.stringify(result).includes(token), false)
})

test("one community failure does not prevent a later community repair", async () => {
  const good = fixture()
  const bad = { ...good.candidate, notificationId: "aaaaaaaa-aaaa-4aaa-8aaa-aaaaaaaaaaaa",
    communityId: "bbbbbbbb-bbbb-4bbb-8bbb-bbbbbbbbbbbb", communityCode: "9999",
    guilds: ["700000000000000099"] }
  const results = await runAnnouncementRepair({ client: { ...good.client, guilds: {
    async fetch(id) {
      if (id === "700000000000000099") throw Object.assign(new Error("missing"), { code: 10004 })
      return good.guild
    }
  } }, api: { baseUrl: "https://r-a-c-h-i-e.com", allowLoopback: false,
    async announcementRepairCandidates() { return { candidates: [bad, good.candidate] } },
    async beginAnnouncementRepair(_id, suppliedCutoff) {
      assert.equal(suppliedCutoff, cutoff)
      return { repair: { ...good.candidate, closesAt: "2030-09-06T12:30:00Z",
        guestPath: `/book/${token}` } }
    }, async completeAnnouncementRepair() {} }, setupRepository: good.setupRepository,
  dryRun: false, sentBefore: cutoff })
  assert.deepEqual(results.map(result => result.status), ["skipped", "completed"])
})
