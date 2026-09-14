const assert = require("node:assert/strict")
const test = require("node:test")

const { diagnoseActiveAccount, auditInvalidPlayerAccounts } = require("../src/giftCodes/playerAccountDiagnostics")
const { parseArguments } = require("../scripts/auditInvalidPlayerAccounts")

function account(overrides = {}) {
  return { id: "11111111-1111-4111-8111-111111111111", game_profile: "wos",
    discord_user_id: "12345", player_id: "67890", state_or_kingdom_number: "1001",
    in_game_name: "Player", alliance_abbreviation: "TAG", is_active: true, ...overrides }
}

test("diagnostic classifies each invalid field without returning private identities", () => {
  const result = diagnoseActiveAccount(account({ discord_user_id: null,
    state_or_kingdom_number: "bad", in_game_name: "", alliance_abbreviation: null }), "wos")
  assert.deepEqual(result.reasons, ["invalid_location", "invalid_in_game_name",
    "invalid_alliance", "invalid_owner"])
  assert.equal(result.discordUserId, null)
  assert.match(result.accountRef, /^[a-f0-9]{16}$/)
  assert.equal(diagnoseActiveAccount(account(), "wos"), null)
})

test("dry-run reports safe aggregate counts and website history without mutation", async () => {
  const source = account({ in_game_name: null })
  const repository = { gameProfile: "wos", async listActiveAccountsForDiagnostics() {
    return [source]
  } }
  const api = { profile: "wos", async playerAccountCleanupPreview(candidates) {
    return { cleanup: { results: [{ accountRef: candidates[0].accountRef,
      reasons: candidates[0].reasons, participantMirrors: 1, activeParticipantMirrors: 1,
      inactiveParticipantMirrors: 0, bookingReferences: 2, historyReferences: 3,
      cleanupStatus: "safe_noop", mutations: 0 }] } }
  } }
  const result = await auditInvalidPlayerAccounts({ profile: "wos", repository, api })
  assert.equal(result.summary.invalidActiveAccounts, 1)
  assert.equal(result.summary.invalidInGameName, 1)
  assert.equal(result.summary.withParticipantMirrors, 1)
  assert.equal(result.summary.withBookings, 1)
  assert.equal("playerId" in result.accounts[0], false)
})

test("targeted release cleans website first and uses guarded soft release", async () => {
  const source = account({ alliance_abbreviation: null })
  const order = []
  const repository = { gameProfile: "wos", async listActiveAccountsForDiagnostics() {
    return [source]
  }, async releaseAccount(input) { order.push("bot"); assert.equal(input.expectedAccountId, source.id)
    return { account: { ...source, is_active: false }, replacement: {
      discord_user_id: source.discord_user_id, player_id: "77777"
    } } } }
  const diagnostic = diagnoseActiveAccount(source, "wos")
  const api = { profile: "wos", async playerAccountCleanupExecute() { order.push("website")
    return { cleanup: { results: [{ cleanupStatus: "deactivated", mutations: 1 }] } } },
  async registration(input) { order.push("primary"); assert.equal(input.primarySyncOnly, true) } }
  const result = await auditInvalidPlayerAccounts({ profile: "wos", repository, api,
    releaseRef: diagnostic.accountRef, operatorDiscordUserId: "99999" })
  assert.deepEqual(order, ["website", "bot", "primary"])
  assert.equal(result.botAccountReleased, true)
  assert.equal(result.replacementPrimaryMirrored, true)
})

test("ownerless invalid account deactivates only after website confirms no active mirror", async () => {
  const source = account({ discord_user_id: null, in_game_name: null })
  const diagnostic = diagnoseActiveAccount(source, "wos")
  let deactivated = 0
  const repository = { gameProfile: "wos", async listActiveAccountsForDiagnostics() {
    return [source]
  }, async deactivateInvalidUnownedAccount(input) {
    assert.equal(input.accountId, source.id); deactivated += 1
    return { ...source, is_active: false }
  } }
  const api = { profile: "wos", async playerAccountCleanupExecute() {
    return { cleanup: { results: [{ cleanupStatus: "safe_noop", mutations: 0 }] } }
  } }
  const result = await auditInvalidPlayerAccounts({ profile: "wos", repository, api,
    releaseRef: diagnostic.accountRef, operatorDiscordUserId: "99999" })
  assert.equal(deactivated, 1)
  assert.equal(result.botAccountReleased, true)
})

test("retry completes after website deactivation when bot release initially fails", async () => {
  const source = account({ alliance_abbreviation: null })
  const diagnostic = diagnoseActiveAccount(source, "wos")
  let releaseAttempts = 0
  let websiteAttempts = 0
  const repository = { gameProfile: "wos",
    async listActiveAccountsForDiagnostics() { return [source] },
    async findCompletedPlayerCleanup() { return null },
    async releaseAccount() {
      releaseAttempts += 1
      if (releaseAttempts === 1) throw new Error("bot database unavailable")
      return { account: { ...source, is_active: false }, replacement: null }
    } }
  const api = { profile: "wos", async playerAccountCleanupExecute() {
    websiteAttempts += 1
    return { cleanup: { results: [{
      cleanupStatus: websiteAttempts === 1 ? "deactivated" : "already_completed",
      mutations: websiteAttempts === 1 ? 1 : 0
    }] } }
  } }
  const input = { profile: "wos", repository, api, releaseRef: diagnostic.accountRef,
    operatorDiscordUserId: "99999" }
  await assert.rejects(auditInvalidPlayerAccounts(input), /bot database unavailable/)
  const retried = await auditInvalidPlayerAccounts(input)
  assert.equal(retried.website.cleanupStatus, "already_completed")
  assert.equal(retried.botAccountReleased, true)
  assert.equal(releaseAttempts, 2)
})

test("retry rediscovers a committed release and only retries replacement MAIN mirroring", async () => {
  const source = account({ alliance_abbreviation: null, is_primary: true })
  const diagnostic = diagnoseActiveAccount(source, "wos")
  const replacement = { discord_user_id: source.discord_user_id, player_id: "77777",
    is_primary: true }
  let active = true
  let releaseCount = 0
  let registrationAttempts = 0
  const completed = { account: { ...source, discord_user_id: null, is_active: false,
      is_primary: false }, previousOwnerDiscordUserId: source.discord_user_id,
    replacement, cleanupMetadata: { reasons: diagnostic.reasons }, alreadyCompleted: true }
  const repository = { gameProfile: "wos",
    async listActiveAccountsForDiagnostics() { return active ? [source] : [] },
    async findCompletedPlayerCleanup(ref) {
      assert.equal(ref, diagnostic.accountRef)
      return active ? null : completed
    },
    async releaseAccount() {
      releaseCount += 1
      active = false
      return { ...completed, alreadyCompleted: false }
    } }
  let websiteAttempts = 0
  const api = { profile: "wos", async playerAccountCleanupExecute() {
    websiteAttempts += 1
    return { cleanup: { results: [{ cleanupStatus: websiteAttempts === 1
      ? "deactivated" : "already_completed", mutations: websiteAttempts === 1 ? 1 : 0 }] } }
  }, async registration(input) {
    registrationAttempts += 1
    assert.equal(input.primarySyncOnly, true)
    if (registrationAttempts === 1) throw new Error("MAIN mirror response lost")
  } }
  const input = { profile: "wos", repository, api, releaseRef: diagnostic.accountRef,
    operatorDiscordUserId: "99999" }
  await assert.rejects(auditInvalidPlayerAccounts(input), /MAIN mirror response lost/)
  const retried = await auditInvalidPlayerAccounts(input)
  assert.equal(retried.botReleaseAlreadyCompleted, true)
  assert.equal(retried.replacementPrimaryMirrored, true)
  assert.equal(releaseCount, 1)
  assert.equal(registrationAttempts, 2)
})

test("a lost bot commit response resolves through the durable cleanup marker", async () => {
  const source = account({ alliance_abbreviation: null })
  const diagnostic = diagnoseActiveAccount(source, "wos")
  let completed = null
  let releaseCount = 0
  const repository = { gameProfile: "wos",
    async listActiveAccountsForDiagnostics() { return [source] },
    async findCompletedPlayerCleanup() { return completed },
    async releaseAccount() {
      releaseCount += 1
      completed = { account: { ...source, discord_user_id: null, is_active: false },
        previousOwnerDiscordUserId: source.discord_user_id, replacement: null,
        cleanupMetadata: { reasons: diagnostic.reasons }, alreadyCompleted: true }
      throw new Error("connection lost after commit")
    } }
  const api = { profile: "wos", async playerAccountCleanupExecute() {
    return { cleanup: { results: [{ cleanupStatus: "deactivated", mutations: 1 }] } }
  } }
  const result = await auditInvalidPlayerAccounts({ profile: "wos", repository, api,
    releaseRef: diagnostic.accountRef, operatorDiscordUserId: "99999" })
  assert.equal(result.botAccountReleased, true)
  assert.equal(result.botReleaseAlreadyCompleted, true)
  assert.equal(releaseCount, 1)
})

test("concurrent cleanup attempts share one guarded release and one ownership transition", async () => {
  const source = account({ alliance_abbreviation: null })
  const diagnostic = diagnoseActiveAccount(source, "wos")
  let active = true
  let releaseCount = 0
  const completed = { account: { ...source, discord_user_id: null, is_active: false },
    previousOwnerDiscordUserId: source.discord_user_id, replacement: null,
    cleanupMetadata: { reasons: diagnostic.reasons }, alreadyCompleted: true }
  const repository = { gameProfile: "wos",
    async listActiveAccountsForDiagnostics() { return [source] },
    async findCompletedPlayerCleanup() { return active ? null : completed },
    async releaseAccount() {
      if (!active) return null
      active = false
      releaseCount += 1
      return { ...completed, alreadyCompleted: false }
    } }
  const api = { profile: "wos", async playerAccountCleanupExecute() {
    return { cleanup: { results: [{ cleanupStatus: "deactivated", mutations: 1 }] } }
  } }
  const input = { profile: "wos", repository, api, releaseRef: diagnostic.accountRef,
    operatorDiscordUserId: "99999" }
  const results = await Promise.all([
    auditInvalidPlayerAccounts(input), auditInvalidPlayerAccounts(input)
  ])
  assert.equal(releaseCount, 1)
  assert.equal(results.filter(result => result.botReleaseAlreadyCompleted).length, 1)
})

test("cleanup CLI requires preview or one exact guarded release", () => {
  assert.deepEqual(parseArguments(["--dry-run"]), { releaseRef: null, operator: null })
  assert.deepEqual(parseArguments(["--release=0123456789abcdef", "--operator=12345"]),
    { releaseRef: "0123456789abcdef", operator: "12345" })
  assert.throws(() => parseArguments([]), /usage/)
})
