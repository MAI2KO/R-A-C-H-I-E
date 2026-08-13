const test = require("node:test")
const assert = require("node:assert/strict")
const { Pool } = require("pg")

const { runMigrations } = require("../src/migrate")
const { createPlayerRepository } = require("../src/giftCodes/playerRepository")
const { createPlayerService } = require("../src/giftCodes/playerService")
const { createGiftCodeRepository } = require("../src/giftCodes/repository")
const { createGiftCodeService } = require("../src/giftCodes/service")
const { createGiftCodeCommunityRepository } = require("../src/giftCodes/communityRepository")

const databaseUrl = process.env.TEST_DATABASE_URL

test("soft-deactivated accounts leave current UX and safely reactivate without losing history", {
  skip: databaseUrl ? false : "TEST_DATABASE_URL is not configured"
}, async () => {
  const admin = new Pool({ connectionString: databaseUrl, max: 1 })
  const schema = `player_lifecycle_${process.pid}_${Date.now()}`
  let pool
  try {
    await admin.query(`CREATE SCHEMA "${schema}"`)
    pool = new Pool({
      connectionString: databaseUrl,
      max: 8,
      options: `-c search_path=${schema}`
    })
    await runMigrations({ pool, logger: { log() {}, error() {} } })

    for (const profile of ["wos", "kingshot"]) {
      const owner = profile === "wos" ? "81001" : "82001"
      const guildId = profile === "wos" ? "81002" : "82002"
      const playerId = profile === "wos" ? "93986200" : "83986200"
      const initialLocation = profile === "wos" ? "689" : "521"
      const nextLocation = profile === "wos" ? "700" : "540"
      const expectedLabel = profile === "wos" ? "State" : "Kingdom"
      const playerRepository = createPlayerRepository(pool, profile)
      const players = createPlayerService({ repository: playerRepository, gameProfile: profile })
      const giftRepository = createGiftCodeRepository(pool, profile)
      const gifts = createGiftCodeService({
        repository: giftRepository,
        gameProfile: profile,
        env: { GIFT_CODE_MAX_AUTO_REDEEM_ACCOUNTS_PER_USER: "2" }
      })
      const community = createGiftCodeCommunityRepository(pool, profile)

      const account = await players.register({
        discordUserId: owner,
        guildId,
        playerId,
        locationNumber: initialLocation
      })
      assert.equal(players.terms.locationLabel, expectedLabel)

      const submitted = await giftRepository.recordSubmission({ code: `${profile}Lifecycle1` })
      await pool.query(
        "UPDATE gift_codes SET status = 'active', verified_at_utc = now() WHERE id = $1",
        [submitted.giftCode.id]
      )
      await gifts.setAutomaticRedemption({
        discordUserId: owner,
        guildId,
        playerId,
        enabled: true
      })
      assert.equal((await giftRepository.activeAccountStatuses(owner, null, guildId)).length, 1)
      assert.equal((await community.communityStats(guildId)).enabled_accounts, 1)

      await gifts.setAutomaticRedemption({
        discordUserId: owner,
        guildId,
        playerId,
        enabled: false
      })
      const optedOut = await playerRepository.getOwnedAccount(owner, playerId)
      assert.deepEqual(optedOut.account_metadata.autoRedeemPreference, {
        enabled: false,
        explicit: true
      })

      const removal = await players.remove({ discordUserId: owner, playerId })
      assert.equal(removal.account.is_active, false)
      assert.equal(removal.replacement, null)
      assert.deepEqual(await playerRepository.listOwnedAccounts(owner), [])
      assert.equal((await playerRepository.listOwnedAccounts(owner, { includeInactive: true })).length, 1)
      assert.deepEqual(await giftRepository.activeAccountStatuses(owner, null, guildId), [])
      const historical = await giftRepository.historicalAccountStatuses(owner, playerId, guildId)
      assert.equal(historical.length, 1)
      assert.equal(historical[0].is_active, false)
      assert.equal(historical[0].gift_redemption_enabled, false)
      assert.equal(historical[0].guild_gift_code_enrolled, false)

      const statsAfterRemoval = await community.communityStats(guildId)
      assert.equal(statsAfterRemoval.registered_users, 0)
      assert.equal(statsAfterRemoval.registered_accounts, 0)
      assert.equal(statsAfterRemoval.auto_redeem_players, 0)
      assert.equal(statsAfterRemoval.enabled_accounts, 0)
      const retainedRedemptions = (await pool.query(
        `SELECT status FROM gift_code_redemptions
          WHERE game_profile = $1 AND player_account_id = $2`,
        [profile, account.id]
      )).rows
      assert.deepEqual(retainedRedemptions.map(row => row.status), ["disabled"])
      assert.equal((await pool.query(
        `SELECT COUNT(*)::integer AS count FROM player_location_history
          WHERE game_profile = $1 AND player_account_id = $2`,
        [profile, account.id]
      )).rows[0].count, 1)

      const reactivated = await players.register({
        discordUserId: owner,
        guildId,
        playerId,
        locationNumber: nextLocation
      })
      assert.equal(reactivated.id, account.id)
      assert.equal(reactivated.is_active, true)
      assert.equal(reactivated.is_primary, true)
      assert.equal(reactivated.gift_redemption_enabled, false)
      assert.deepEqual(reactivated.account_metadata.autoRedeemPreference, {
        enabled: false,
        explicit: true
      })
      assert.equal(reactivated.state_or_kingdom_number, nextLocation)
      assert.equal((await pool.query(
        `SELECT COUNT(*)::integer AS count FROM player_accounts
          WHERE game_profile = $1 AND player_id = $2`,
        [profile, playerId]
      )).rows[0].count, 1)
      const locationHistory = (await pool.query(
        `SELECT previous_number, new_number FROM player_location_history
          WHERE game_profile = $1 AND player_account_id = $2
          ORDER BY changed_at_utc, id`,
        [profile, account.id]
      )).rows
      assert.deepEqual(locationHistory, [
        { previous_number: null, new_number: initialLocation },
        { previous_number: initialLocation, new_number: nextLocation }
      ])
      assert.equal((await giftRepository.activeAccountStatuses(owner, playerId, guildId)).length, 1)
      assert.equal((await giftRepository.activeAccountStatuses(owner, playerId, guildId))[0]
        .guild_gift_code_enrolled, false)
    }
  } finally {
    await pool?.end().catch(() => {})
    await admin.query(`DROP SCHEMA IF EXISTS "${schema}" CASCADE`).catch(() => {})
    await admin.end()
  }
})
