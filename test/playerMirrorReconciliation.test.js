const assert = require("node:assert/strict")
const test = require("node:test")

const { createBookingWebsiteClient } = require("../src/bookingWebsiteClient")
const { createPlayerRepository } = require("../src/giftCodes/playerRepository")
const { runPlayerMirrorReconciliation } = require("../src/giftCodes/playerMirrorReconciliation")
const { parseArguments, runCli } = require("../scripts/reconcilePlayerMirrors")

function sourceAccount(playerId, overrides = {}) {
  return { game_profile: "wos", discord_user_id: "1234567", player_id: playerId,
    state_or_kingdom_number: "1001", in_game_name: `Player ${playerId}`,
    alliance_abbreviation: "TAG", is_primary: false, is_active: true,
    gift_redemption_enabled: true, ...overrides }
}

test("operator arguments require exactly one explicit mode", () => {
  assert.deepEqual(parseArguments(["--dry-run"]), { dryRun: true })
  assert.deepEqual(parseArguments(["--execute"]), { dryRun: false })
  assert.throws(() => parseArguments([]), /usage/)
  assert.throws(() => parseArguments(["--dry-run", "--execute"]), /usage/)
})

test("repository enumeration is exact-profile and active-only without updates", async () => {
  const calls = []
  const pool = { async query(sql, values) { calls.push([sql, values]); return { rows: [] } },
    async connect() { throw new Error("not used") } }
  const repository = createPlayerRepository(pool, "kingshot")
  assert.deepEqual(await repository.listActiveAccountsForReconciliation(), [])
  assert.match(calls[0][0], /game_profile=\$1 AND is_active=true/)
  assert.deepEqual(calls[0][1], ["kingshot"])
  assert.doesNotMatch(calls[0][0], /UPDATE|INSERT|DELETE/i)
})

test("dry-run groups owners, calls preview only, and preserves bot account state", async () => {
  const accounts = [sourceAccount("111111", { is_primary: true }),
    sourceAccount("222222"), sourceAccount("333333", { discord_user_id: "7654321",
      gift_redemption_enabled: false })]
  const before = structuredClone(accounts)
  const previews = []
  const api = { profile: "wos", async playerMirrorPreview(group) {
    previews.push(group)
    return { reconciliation: { mutations: 0, plannedCreates: 1, plannedUpdates: 1,
      results: group.map((value, index) => ({ profile: "wos",
        discordUserId: value.discordUserId, playerId: value.playerId,
        botIsPrimary: value.isPrimary,
        websiteMirrorStatus: index === 0 ? "primary mismatch" : "matching",
        communities: ["1001"], plannedAction: index === 0 ? "synchronize primary" : "none" })) } }
  }, async playerMirrorExecute() { throw new Error("must not execute") } }
  const result = await runPlayerMirrorReconciliation({ profile: "wos", dryRun: true, api,
    repository: { gameProfile: "wos", async listActiveAccountsForReconciliation() {
      return accounts
    } } })
  assert.equal(previews.length, 2)
  assert.equal(result.summary.activeBotAccounts, 3)
  assert.equal(result.summary.primaryMismatches, 2)
  assert.equal(result.summary.mutations, 0)
  assert.deepEqual(accounts, before)
})

test("execution uses execute only and rejects every profile crossover", async () => {
  const account = sourceAccount("111111", { is_primary: true })
  let executions = 0
  const api = { profile: "wos", async playerMirrorExecute(group) {
    executions += 1
    assert.equal(group[0].isPrimary, true)
    assert.equal("giftRedemptionEnabled" in group[0], false)
    return { reconciliation: { mutations: 2, plannedCreates: 1, plannedUpdates: 1,
      results: [{ profile: "wos", discordUserId: "1234567", playerId: "111111",
        botIsPrimary: true, websiteMirrorStatus: "missing", communities: [],
        plannedAction: "create mirror for 1001; synchronize primary" }] } }
  } }
  const repository = { gameProfile: "wos", async listActiveAccountsForReconciliation() {
    return [account]
  } }
  const result = await runPlayerMirrorReconciliation({ profile: "wos", repository, api,
    dryRun: false })
  assert.equal(executions, 1)
  assert.equal(result.summary.plannedCreates, 1)
  assert.equal(result.summary.mutations, 2)
  await assert.rejects(runPlayerMirrorReconciliation({ profile: "kingshot", repository,
    api: { ...api, profile: "kingshot" }, dryRun: false }), /profile mismatch/)
  await assert.rejects(runPlayerMirrorReconciliation({ profile: "wos", repository,
    api: { ...api, profile: "kingshot" }, dryRun: false }), /profile mismatch/)
})

test("invalid-metadata owner groups are counted as skipped without stopping later owners", async () => {
  const accounts = [sourceAccount("111111", { in_game_name: "", alliance_abbreviation: "" }),
    sourceAccount("222222", { discord_user_id: "7654321" })]
  const owners = []
  const api = { profile: "wos", async playerMirrorPreview(group) {
    owners.push(group[0].discordUserId)
    const invalid = group[0].inGameName === ""
    return { reconciliation: { conflict: invalid ? "invalid_account_metadata" : null,
      mutations: 0, results: group.map(value => ({ profile: "wos",
        discordUserId: value.discordUserId, playerId: value.playerId,
        botIsPrimary: value.isPrimary,
        websiteMirrorStatus: invalid ? "ambiguous/conflict" : "matching",
        communities: invalid ? [] : ["1001"],
        plannedAction: invalid
          ? "skip: owner group contains invalid account metadata" : "none" })) } }
  } }
  const result = await runPlayerMirrorReconciliation({ profile: "wos", dryRun: true, api,
    repository: { gameProfile: "wos", async listActiveAccountsForReconciliation() {
      return accounts
    } } })
  assert.equal(owners.length, 2)
  assert.equal(result.summary.activeBotAccounts, 2)
  assert.equal(result.summary.skipped, 1)
  assert.equal(result.summary.ambiguousConflicted, 1)
  assert.equal(result.summary.matching, 1)
})

test("CLI exits 2 for completed skips and 1 for infrastructure failure", async () => {
  const completedRuntime = { exitCode: undefined, stderr: { write() {
    assert.fail("completed reconciliation must not write an error")
  } } }
  await runCli({ operation: async () => ({ summary: { skipped: 1 } }),
    runtime: completedRuntime })
  assert.equal(completedRuntime.exitCode, 2)

  const messages = []
  const failedRuntime = { exitCode: undefined, stderr: { write(message) {
    messages.push(message)
  } } }
  await runCli({ operation: async () => { throw new Error("private account data") },
    runtime: failedRuntime })
  assert.equal(failedRuntime.exitCode, 1)
  assert.deepEqual(messages, ["Player mirror reconciliation failed safely.\n"])
})

test("signed website client uses distinct preview and execute endpoints", async () => {
  const paths = []
  const client = createBookingWebsiteClient({ config: { enabled: true, profile: "wos",
    baseUrl: "https://r-a-c-h-i-e.com", secret: "s".repeat(32), allowLoopback: false },
  createNonce: () => "11111111-1111-4111-8111-111111111111", now: () => 1760000000000,
  async fetchImplementation(url, options) {
    paths.push([new URL(url).pathname, JSON.parse(options.body), options.headers])
    return new Response(JSON.stringify({ ok: true, reconciliation: { results: [] } }),
      { status: 200, headers: { "content-type": "application/json" } })
  } })
  await client.playerMirrorPreview([{ playerId: "1" }])
  await client.playerMirrorExecute([{ playerId: "1" }])
  await client.playerAccountCleanupPreview([{ accountRef: "a" }])
  await client.playerAccountCleanupExecute([{ accountRef: "a" }])
  assert.deepEqual(paths.map(([path]) => path), [
    "/api/internal/v1/discord/player-mirrors/preview",
    "/api/internal/v1/discord/player-mirrors/execute",
    "/api/internal/v1/discord/player-account-cleanup/preview",
    "/api/internal/v1/discord/player-account-cleanup/execute"
  ])
  assert.ok(paths.every(([, body, headers]) => (Array.isArray(body.accounts)
      || Array.isArray(body.candidates))
    && headers["x-booking-profile"] === "wos"
    && /^v1=/.test(headers["x-booking-signature"])))
})
