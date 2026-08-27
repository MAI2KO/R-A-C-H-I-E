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
        locationNumber: initialLocation,
        inGameName: `${profile} Player`, allianceAbbreviation: "TST"
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

      await assert.rejects(
        players.register({
          discordUserId: `${owner}9`,
          guildId,
          playerId,
          locationNumber: nextLocation,
          inGameName: "Other Player", allianceAbbreviation: "TST"
        }),
        error => error.code === "PLAYER_ALREADY_REGISTERED"
          && /current owner needs to release it first/.test(error.message)
      )

      const reactivated = await players.register({
        discordUserId: owner,
        guildId,
        playerId,
        locationNumber: nextLocation,
        inGameName: `${profile} Player`, allianceAbbreviation: "TST"
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

test("explicit release preserves canonical history and safely establishes a new owner", {
  skip: databaseUrl ? false : "TEST_DATABASE_URL is not configured"
}, async () => {
  const admin = new Pool({ connectionString: databaseUrl, max: 1 })
  const schema = `player_release_${process.pid}_${Date.now()}`
  let pool
  try {
    await admin.query(`CREATE SCHEMA "${schema}"`)
    pool = new Pool({
      connectionString: databaseUrl,
      max: 8,
      options: `-c search_path=${schema}`
    })
    await runMigrations({ pool, logger: { log() {}, error() {} } })

    const playerId = "93986209"
    const oldOwner = "707866087248756736"
    const newOwner = "807866087248756736"
    const operator = "907866087248756736"
    const wosGuild = "717866087248756736"
    const kingshotGuild = "817866087248756736"
    const wosRepository = createPlayerRepository(pool, "wos")
    const kingshotRepository = createPlayerRepository(pool, "kingshot")
    const wosPlayers = createPlayerService({ repository: wosRepository, gameProfile: "wos" })
    const kingshotPlayers = createPlayerService({ repository: kingshotRepository, gameProfile: "kingshot" })
    const gifts = createGiftCodeRepository(pool, "wos")

    const wosAccount = await wosPlayers.register({
      discordUserId: oldOwner, guildId: wosGuild, playerId, locationNumber: "689",
      inGameName: "WOS Owner", allianceAbbreviation: "TST"
    })
    const kingshotAccount = await kingshotPlayers.register({
      discordUserId: oldOwner, guildId: kingshotGuild, playerId, locationNumber: "521",
      inGameName: "King Owner", allianceAbbreviation: "TST"
    })
    assert.notEqual(wosAccount.id, kingshotAccount.id)

    const redemptions = []
    for (const [suffix, status, notification] of [
      ["Success", "success", "sent"],
      ["Already", "already_redeemed", "pending"],
      ["Pending", "claimed", "sending"]
    ]) {
      const submission = await gifts.recordSubmission({ code: `Release${suffix}` })
      await pool.query(
        "UPDATE gift_codes SET status = 'active', verified_at_utc = now() WHERE id = $1",
        [submission.giftCode.id]
      )
      const queued = await gifts.queueRedemption({
        giftCodeId: submission.giftCode.id,
        playerAccountId: wosAccount.id
      })
      const updated = (await pool.query(
        `UPDATE gift_code_redemptions
            SET status = $2::varchar, notification_status = $3::varchar,
                claimed_by_worker = CASE WHEN $2::varchar = 'claimed' THEN 'test-worker' ELSE NULL END,
                claimed_at_utc = CASE WHEN $2::varchar = 'claimed' THEN now() ELSE NULL END,
                claimed_until_utc = CASE WHEN $2::varchar = 'claimed' THEN now() + interval '1 minute' ELSE NULL END,
                completed_at_utc = CASE WHEN $2::varchar IN ('success', 'already_redeemed') THEN now() ELSE NULL END
          WHERE id = $1 RETURNING *`,
        [queued.redemption.id, status, notification]
      )).rows[0]
      redemptions.push(updated)
    }

    await pool.query(
      `UPDATE player_accounts SET gift_redemption_enabled = true,
          account_metadata = '{"autoRedeemPreference":{"enabled":true,"explicit":true}}'::jsonb
        WHERE id = $1`,
      [wosAccount.id]
    )
    await pool.query(
      `UPDATE player_account_guilds SET gift_code_enrolled = true,
          gift_code_updated_at_utc = now()
        WHERE game_profile = 'wos' AND player_account_id = $1`,
      [wosAccount.id]
    )

    const released = await wosPlayers.release({ discordUserId: oldOwner, playerId })
    assert.equal(released.account.id, wosAccount.id)
    assert.equal(released.account.discord_user_id, null)
    assert.equal(released.account.is_active, false)
    assert.equal(released.account.gift_redemption_enabled, false)
    assert.equal(released.account.account_metadata.autoRedeemPreference.enabled, false)
    assert.deepEqual(released.guildIds, [wosGuild])
    assert.deepEqual(await wosRepository.listOwnedAccounts(oldOwner, { includeInactive: true }), [])
    assert.equal((await pool.query(
      "SELECT COUNT(*)::integer AS count FROM player_account_guilds WHERE game_profile = 'wos' AND player_account_id = $1",
      [wosAccount.id]
    )).rows[0].count, 0)

    const retained = (await pool.query(
      `SELECT id, status, notification_status, discord_owner_id_snapshot
         FROM gift_code_redemptions
        WHERE game_profile = 'wos' AND player_account_id = $1 ORDER BY created_at_utc, id`,
      [wosAccount.id]
    )).rows
    assert.equal(retained.length, 3)
    assert.equal(retained.find(row => row.id === redemptions[0].id).status, "success")
    assert.equal(retained.find(row => row.id === redemptions[0].id).notification_status, "sent")
    assert.equal(retained.find(row => row.id === redemptions[1].id).status, "already_redeemed")
    assert.equal(retained.find(row => row.id === redemptions[1].id).notification_status, "suppressed")
    assert.equal(retained.find(row => row.id === redemptions[2].id).status, "disabled")
    assert.equal(retained.find(row => row.id === redemptions[2].id).notification_status, "suppressed")
    assert.ok(retained.every(row => row.discord_owner_id_snapshot === oldOwner))
    assert.equal(await gifts.claimNotification(redemptions[1].id), null)

    const completedAfterRelease = await gifts.finishRedemption({
      claim: {
        ...redemptions[2],
        gift_code_id: redemptions[2].gift_code_id,
        player_account_id: wosAccount.id,
        player_id_snapshot: playerId,
        location_number_snapshot: "689",
        attempt_number: 1
      },
      workerId: "test-worker",
      result: {
        httpStatus: 200,
        classification: {
          state: "success",
          raw: { code: 1, errCode: 20000, message: "success" }
        },
        requestStartedAt: new Date("2026-08-13T10:00:00Z"),
        responseReceivedAt: new Date("2026-08-13T10:00:01Z")
      },
      now: new Date("2026-08-13T10:00:01Z"),
      status: "success"
    })
    assert.equal(completedAfterRelease.released_during_attempt, true)
    assert.equal(completedAfterRelease.status, "disabled")
    assert.equal(completedAfterRelease.notification_status, "suppressed")
    assert.equal((await pool.query(
      `SELECT COUNT(*)::integer AS count FROM gift_code_attempts
        WHERE game_profile = 'wos' AND redemption_id = $1`,
      [redemptions[2].id]
    )).rows[0].count, 1)

    const claimed = await wosPlayers.register({
      discordUserId: newOwner, guildId: wosGuild, playerId, locationNumber: "700",
      inGameName: "New Owner", allianceAbbreviation: "TST"
    })
    assert.equal(claimed.id, wosAccount.id)
    assert.equal(claimed.registration_status, "claimed")
    assert.equal(claimed.discord_user_id, newOwner)
    assert.equal(claimed.gift_redemption_enabled, false)
    assert.equal(claimed.account_metadata.autoRedeemPreference, undefined)
    assert.equal((await gifts.redemptionHistory(claimed.id, newOwner)).length, 0)
    assert.equal((await gifts.redemptionHistory(claimed.id, oldOwner)).length, 3)
    assert.equal((await pool.query(
      `SELECT gift_code_enrolled FROM player_account_guilds
        WHERE game_profile = 'wos' AND guild_id = $1 AND player_account_id = $2`,
      [wosGuild, claimed.id]
    )).rows[0].gift_code_enrolled, false)

    const ownership = (await pool.query(
      `SELECT action_type, previous_discord_user_id, new_discord_user_id,
              performed_by_discord_user_id
         FROM player_account_ownership_history
        WHERE game_profile = 'wos' AND player_account_id = $1
        ORDER BY changed_at_utc, id`,
      [wosAccount.id]
    )).rows
    assert.deepEqual(ownership.map(row => row.action_type), ["release", "claim"])
    assert.equal(ownership[0].previous_discord_user_id, oldOwner)
    assert.equal(ownership[0].performed_by_discord_user_id, oldOwner)
    assert.equal(ownership[1].new_discord_user_id, newOwner)

    const isolatedKingshot = await kingshotRepository.getAccountByPlayerId(playerId)
    assert.equal(isolatedKingshot.discord_user_id, oldOwner)
    assert.equal(isolatedKingshot.is_active, true)
    await assert.rejects(
      kingshotPlayers.operatorRelease({
        playerId,
        operatorDiscordUserId: operator,
        expectedAccountId: kingshotAccount.id,
        expectedOwnerDiscordUserId: newOwner
      }),
      error => error.code === "PLAYER_OWNERSHIP_CHANGED"
    )
    const operatorRelease = await kingshotPlayers.operatorRelease({
      playerId,
      operatorDiscordUserId: operator,
      expectedAccountId: kingshotAccount.id,
      expectedOwnerDiscordUserId: oldOwner
    })
    assert.equal(operatorRelease.previousOwnerDiscordUserId, oldOwner)
    const operatorAudit = (await pool.query(
      `SELECT * FROM player_account_ownership_history
        WHERE game_profile = 'kingshot' AND player_account_id = $1`,
      [kingshotAccount.id]
    )).rows[0]
    assert.equal(operatorAudit.action_type, "operator_release")
    assert.equal(operatorAudit.performed_by_discord_user_id, operator)
    assert.equal(operatorAudit.previous_discord_user_id, oldOwner)
  } finally {
    await pool?.end().catch(() => {})
    await admin.query(`DROP SCHEMA IF EXISTS "${schema}" CASCADE`).catch(() => {})
    await admin.end()
  }
})
