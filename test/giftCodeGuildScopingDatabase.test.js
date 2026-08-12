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
const logger = { log() {}, error() {} }

function redemptionResult(state) {
  return {
    httpStatus: 200,
    headers: {},
    classification: {
      state,
      raw: {
        code: state === "success" ? 0 : null,
        errCode: state === "success" ? 20000 : 40008,
        message: state === "success" ? "SUCCESS" : "RECEIVED"
      }
    }
  }
}

test("gift-code community activity is explicitly guild scoped while redemption stays global", {
  skip: databaseUrl ? false : "TEST_DATABASE_URL is not configured"
}, async () => {
  const admin = new Pool({ connectionString: databaseUrl, max: 1 })
  const schema = `gift_guild_scope_${process.pid}_${Date.now()}`
  let pool
  try {
    await admin.query(`CREATE SCHEMA "${schema}"`)
    pool = new Pool({
      connectionString: databaseUrl,
      max: 10,
      options: `-c search_path=${schema}`
    })
    const migrations = await runMigrations({ pool, logger })
    assert.equal(migrations.applied.at(-1), "016_gift_code_sources.sql")

    const guildA = "700000000000000001"
    const guildB = "700000000000000002"
    const guildC = "700000000000000003"
    const owner = "100000000000000001"
    const wosCommunity = createGiftCodeCommunityRepository(pool, "wos")
    for (const [guildId, channelId] of [
      [guildA, "800000000000000001"],
      [guildB, "800000000000000002"],
      [guildC, "800000000000000003"]
    ]) await wosCommunity.setChannel(guildId, channelId)

    const wosPlayers = createPlayerService({
      repository: createPlayerRepository(pool, "wos"),
      gameProfile: "wos"
    })
    const account = await wosPlayers.register({
      discordUserId: owner,
      playerId: "200000001",
      locationNumber: "689"
    })
    const wosRepository = createGiftCodeRepository(pool, "wos")
    const wosGifts = createGiftCodeService({ repository: wosRepository, gameProfile: "wos" })
    const submission = await wosRepository.recordSubmission({
      code: "GuildScopedCode",
      submittedByDiscordUserId: owner,
      metadata: { guildId: guildA, submissionKind: "user" }
    })
    await pool.query(
      `UPDATE gift_codes SET status = 'active', verification_state = 'complete',
          verified_at_utc = now() WHERE id = $1`,
      [submission.giftCode.id]
    )

    const guildAEnable = await wosGifts.setAutomaticRedemption({
      discordUserId: owner,
      guildId: guildA,
      playerId: account.player_id,
      enabled: true
    })
    assert.ok(guildAEnable.engagement_event)
    assert.equal(guildAEnable.engagement_event.guild_id, guildA)
    assert.equal((await wosCommunity.communityStats(guildA)).auto_redeem_players, 1)
    assert.equal((await wosCommunity.communityStats(guildB)).auto_redeem_players, 0)
    assert.equal((await wosCommunity.communityStats(guildB)).registered_accounts, 0)

    const guildCEnable = await wosGifts.setAutomaticRedemption({
      discordUserId: owner,
      guildId: guildC,
      playerId: account.player_id,
      enabled: true
    })
    assert.ok(guildCEnable.engagement_event)
    assert.equal(guildCEnable.engagement_event.guild_id, guildC)
    const enrolments = (await pool.query(
      `SELECT guild_id, gift_code_enrolled, gift_code_first_enabled_at_utc
         FROM player_account_guilds
        WHERE game_profile = 'wos' AND player_account_id = $1
        ORDER BY guild_id`,
      [account.id]
    )).rows
    assert.deepEqual(enrolments.map(row => row.guild_id), [guildA, guildC])
    assert.ok(enrolments.every(row => row.gift_code_enrolled && row.gift_code_first_enabled_at_utc))
    assert.equal((await wosCommunity.communityStats(guildC)).auto_redeem_players, 1)
    assert.equal((await pool.query(
      `SELECT COUNT(*)::integer AS count FROM gift_code_redemptions
        WHERE game_profile = 'wos' AND gift_code_id = $1 AND player_account_id = $2`,
      [submission.giftCode.id, account.id]
    )).rows[0].count, 1, "multi-guild enrolment duplicated the Century redemption")

    const engagement = await wosCommunity.prepareCodeEngagement(submission.giftCode.id, 99)
    assert.ok(engagement.length > 0)
    assert.ok(engagement.every(event => event.guild_id === guildA))
    assert.equal((await pool.query(
      `SELECT COUNT(*)::integer AS count FROM gift_code_engagement_events
        WHERE game_profile = 'wos' AND gift_code_id = $1 AND guild_id <> $2`,
      [submission.giftCode.id, guildA]
    )).rows[0].count, 0)
    const progressEvent = engagement.find(event => event.event_type === "code_progress")
    assert.equal(progressEvent.progress_remaining, 1)
    assert.equal(progressEvent.metadata.queuedCount, 1)
    await pool.query(
      `UPDATE gift_code_engagement_events
          SET status = 'completed', progress_remaining = 1,
              channel_id = '800000000000000001', message_id = '900000000000000001'
        WHERE id = $1`,
      [progressEvent.id]
    )

    const claim = await wosRepository.claimRedemption({
      workerId: "redeem-worker",
      now: new Date("2026-08-11T14:00:00Z"),
      leaseSeconds: 60
    })
    assert.equal(claim.player_account_id, account.id)
    await wosRepository.finishRedemption({
      claim,
      workerId: "redeem-worker",
      result: redemptionResult("success"),
      now: new Date("2026-08-11T14:00:01Z"),
      status: "success"
    })
    const dmClaims = await Promise.all([
      wosRepository.claimNotification(claim.id),
      wosRepository.claimNotification(claim.id)
    ])
    assert.equal(dmClaims.filter(Boolean).length, 1, "one result produced multiple player DMs")
    assert.equal((await pool.query(
      `SELECT COUNT(*)::integer AS count FROM gift_code_attempts
        WHERE game_profile = 'wos' AND redemption_id = $1`,
      [claim.id]
    )).rows[0].count, 1)
    const refresh = await wosCommunity.claimProgressRefresh(
      submission.giftCode.id,
      account.id,
      "success",
      "progress-worker",
      new Date("2026-08-11T14:00:02Z")
    )
    assert.equal(refresh.event.guild_id, guildA)
    assert.equal(refresh.progress.remaining, 0)
    assert.equal((await pool.query(
      `SELECT COUNT(*)::integer AS count FROM gift_code_engagement_events
        WHERE game_profile = 'wos' AND gift_code_id = $1
          AND event_type = 'code_progress' AND guild_id IN ($2, $3)`,
      [submission.giftCode.id, guildB, guildC]
    )).rows[0].count, 0)

    const alreadyCode = await wosRepository.recordSubmission({ code: "AlreadyClaimedCode" })
    await pool.query(
      `UPDATE gift_codes SET status = 'active', verification_state = 'complete',
          verified_at_utc = now() WHERE id = $1`,
      [alreadyCode.giftCode.id]
    )
    assert.equal((await wosRepository.fanOutActiveCode({
      giftCodeId: alreadyCode.giftCode.id
    })).length, 1)
    const alreadyClaim = await wosRepository.claimRedemption({
      workerId: "already-worker",
      now: new Date("2026-08-11T14:01:00Z"),
      leaseSeconds: 60
    })
    assert.equal(alreadyClaim.gift_code_id, alreadyCode.giftCode.id)
    await wosRepository.finishRedemption({
      claim: alreadyClaim,
      workerId: "already-worker",
      result: redemptionResult("already_redeemed"),
      now: new Date("2026-08-11T14:01:01Z"),
      status: "already_redeemed"
    })
    const guildAStats = await wosCommunity.communityStats(guildA)
    const guildCStats = await wosCommunity.communityStats(guildC)
    const guildBStats = await wosCommunity.communityStats(guildB)
    for (const stats of [guildAStats, guildCStats]) {
      assert.equal(stats.successful_redemptions, 1)
      assert.equal(stats.already_redeemed, 1)
      assert.equal(stats.successful_redemptions + stats.already_redeemed, 2)
    }
    assert.equal(guildBStats.successful_redemptions, 0)
    assert.equal(guildBStats.already_redeemed, 0)
    const accountStatus = (await wosRepository.accountStatuses(
      owner,
      account.player_id,
      guildA
    ))[0]
    assert.equal(accountStatus.successful_redemptions, 1)
    assert.equal(accountStatus.already_redeemed, 1)
    assert.equal(accountStatus.completed_redemption_checks, 2)
    const ownerStats = await wosCommunity.accountOwnerStats(owner, guildA)
    assert.equal(ownerStats.successfulRedemptions, 1)
    assert.equal(ownerStats.alreadyRedeemed, 1)

    await wosGifts.setAutomaticRedemption({
      discordUserId: owner,
      guildId: guildA,
      playerId: account.player_id,
      enabled: false
    })
    assert.equal((await wosCommunity.communityStats(guildA)).auto_redeem_players, 0)
    assert.equal((await wosCommunity.communityStats(guildC)).auto_redeem_players, 0)
    assert.equal((await pool.query(
      `SELECT COUNT(*)::integer AS count FROM player_account_guilds
        WHERE game_profile = 'wos' AND player_account_id = $1
          AND gift_code_enrolled = true`,
      [account.id]
    )).rows[0].count, 2, "global disable erased an explicit guild enrolment")
    const reenabled = await wosGifts.setAutomaticRedemption({
      discordUserId: owner,
      guildId: guildA,
      playerId: account.player_id,
      enabled: true
    })
    assert.equal(reenabled.engagement_event, null)
    assert.equal((await wosCommunity.communityStats(guildC)).auto_redeem_players, 1)
    assert.equal((await pool.query(
      `SELECT COUNT(*)::integer AS count FROM gift_code_redemptions
        WHERE game_profile = 'wos' AND gift_code_id = $1 AND player_account_id = $2`,
      [submission.giftCode.id, account.id]
    )).rows[0].count, 1)

    const ksPlayers = createPlayerService({
      repository: createPlayerRepository(pool, "kingshot"),
      gameProfile: "kingshot"
    })
    const ksAccount = await ksPlayers.register({
      discordUserId: owner,
      playerId: "300000001",
      locationNumber: "521"
    })
    const ksGifts = createGiftCodeService({
      repository: createGiftCodeRepository(pool, "kingshot"),
      gameProfile: "kingshot"
    })
    await ksGifts.setAutomaticRedemption({
      discordUserId: owner,
      guildId: guildB,
      playerId: ksAccount.player_id,
      enabled: true
    })
    assert.equal((await createGiftCodeCommunityRepository(pool, "kingshot")
      .communityStats(guildB)).auto_redeem_players, 1)
    const kingshotStats = await createGiftCodeCommunityRepository(pool, "kingshot")
      .communityStats(guildB)
    assert.equal(kingshotStats.successful_redemptions, 0)
    assert.equal(kingshotStats.already_redeemed, 0)
    assert.equal((await wosCommunity.communityStats(guildB)).auto_redeem_players, 0)
  } finally {
    await pool?.end().catch(() => {})
    await admin.query(`DROP SCHEMA IF EXISTS "${schema}" CASCADE`).catch(() => {})
    await admin.end()
  }
})
