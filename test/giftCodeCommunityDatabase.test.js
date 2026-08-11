const test = require("node:test")
const assert = require("node:assert/strict")
const { Pool } = require("pg")

const { runMigrations } = require("../src/migrate")
const { createPlayerRepository } = require("../src/giftCodes/playerRepository")
const { createPlayerService } = require("../src/giftCodes/playerService")
const { createGiftCodeRepository } = require("../src/giftCodes/repository")
const { createGiftCodeService, GiftCodeError } = require("../src/giftCodes/service")
const { createGiftCodeCommunityRepository } = require("../src/giftCodes/communityRepository")

const databaseUrl = process.env.TEST_DATABASE_URL
const logger = { log() {}, error() {} }

test("community configuration, auto-redemption cap and engagement state are durable and scoped", {
  skip: databaseUrl ? false : "TEST_DATABASE_URL is not configured"
}, async () => {
  const admin = new Pool({ connectionString: databaseUrl, max: 1 })
  const schema = `gift_community_${process.pid}_${Date.now()}`
  let pool
  try {
    await admin.query(`CREATE SCHEMA "${schema}"`)
    pool = new Pool({
      connectionString: databaseUrl,
      max: 12,
      options: `-c search_path=${schema}`
    })
    const first = await runMigrations({ pool, logger })
    assert.equal(first.applied.length, 13)
    assert.equal(first.applied.at(-1), "013_gift_code_community.sql")
    assert.deepEqual((await runMigrations({ pool, logger })).applied, [])

    const wosCommunity = createGiftCodeCommunityRepository(pool, "wos")
    const ksCommunity = createGiftCodeCommunityRepository(pool, "kingshot")
    await wosCommunity.setChannel("777777777777777777", "888888888888888888")
    assert.equal((await wosCommunity.getSettings("777777777777777777")).gift_code_channel_id, "888888888888888888")
    assert.equal(await ksCommunity.getSettings("777777777777777777"), null)

    const provisionClaims = await Promise.all([
      wosCommunity.claimContributorRoleProvision(
        "777777777777777777", "role-worker-a", new Date("2026-08-11T09:00:00Z")
      ),
      wosCommunity.claimContributorRoleProvision(
        "777777777777777777", "role-worker-b", new Date("2026-08-11T09:00:00Z")
      )
    ])
    assert.equal(provisionClaims.filter(Boolean).length, 1)
    const roleWorker = provisionClaims[0] ? "role-worker-a" : "role-worker-b"
    await wosCommunity.completeContributorRoleProvision(
      "777777777777777777",
      roleWorker,
      "999999999999999999",
      new Date("2026-08-11T09:00:01Z")
    )
    const roleSettings = await wosCommunity.getSettings("777777777777777777")
    assert.equal(roleSettings.contributor_role_id, "999999999999999999")
    assert.equal(roleSettings.contributor_role_status, "ready")

    const owner = "111111111111111111"
    const guild = "777777777777777777"
    const playerRepository = createPlayerRepository(pool, "wos")
    const players = createPlayerService({ repository: playerRepository, gameProfile: "wos" })
    const registered = []
    for (let index = 1; index <= 4; index += 1) {
      registered.push(await players.register({
        discordUserId: owner,
        guildId: guild,
        playerId: `1000${index}`,
        locationNumber: `${688 + index}`
      }))
    }
    const giftRepository = createGiftCodeRepository(pool, "wos")
    const gifts = createGiftCodeService({
      repository: giftRepository,
      gameProfile: "wos",
      env: { GIFT_CODE_MAX_AUTO_REDEEM_ACCOUNTS_PER_USER: "2" }
    })
    const enabled = await Promise.allSettled(registered.slice(0, 3).map(account =>
      gifts.setAutomaticRedemption({
        discordUserId: owner,
        guildId: guild,
        playerId: account.player_id,
        enabled: true
      })
    ))
    assert.equal(
      enabled.filter(result => result.status === "fulfilled").length,
      2,
      enabled.filter(result => result.status === "rejected")
        .map(result => `${result.reason?.code}:${result.reason?.message}`).join(" | ")
    )
    const rejected = enabled.find(result => result.status === "rejected")
    assert.ok(rejected.reason instanceof GiftCodeError)
    assert.equal(rejected.reason.code, "AUTO_REDEEM_ACCOUNT_LIMIT")
    assert.match(rejected.reason.message, /up to 2 Whiteout Survival accounts/)
    assert.equal((await pool.query(
      `SELECT COUNT(*)::integer AS count FROM player_accounts
        WHERE game_profile = 'wos' AND discord_user_id = $1
          AND is_active = true AND gift_redemption_enabled = true`,
      [owner]
    )).rows[0].count, 2)

    const currentlyEnabled = (await giftRepository.accountStatuses(owner))
      .filter(account => account.gift_redemption_enabled)
    await gifts.setAutomaticRedemption({
      discordUserId: owner,
      guildId: guild,
      playerId: currentlyEnabled[0].player_id,
      enabled: false
    })
    const third = registered.find(account => !currentlyEnabled.some(value => value.id === account.id))
    await gifts.setAutomaticRedemption({
      discordUserId: owner,
      guildId: guild,
      playerId: third.player_id,
      enabled: true
    })
    assert.equal((await giftRepository.accountStatuses(owner))
      .filter(account => account.is_active && account.gift_redemption_enabled).length, 2)

    const enabledForRemoval = (await giftRepository.accountStatuses(owner))
      .find(account => account.gift_redemption_enabled)
    await players.remove({ discordUserId: owner, playerId: enabledForRemoval.player_id })
    const fourth = registered.find(account => account.id !== third.id
      && account.id !== enabledForRemoval.id
      && !currentlyEnabled.some(value => value.id === account.id)) || registered[3]
    await gifts.setAutomaticRedemption({
      discordUserId: owner,
      guildId: guild,
      playerId: fourth.player_id,
      enabled: true
    })
    assert.equal((await giftRepository.accountStatuses(owner))
      .filter(account => account.is_active && account.gift_redemption_enabled).length, 2)

    const ksPlayers = createPlayerService({
      repository: createPlayerRepository(pool, "kingshot"),
      gameProfile: "kingshot"
    })
    const ksAccount = await ksPlayers.register({
      discordUserId: owner,
      guildId: guild,
      playerId: "10001",
      locationNumber: "521"
    })
    const ksGifts = createGiftCodeService({
      repository: createGiftCodeRepository(pool, "kingshot"),
      gameProfile: "kingshot",
      env: { GIFT_CODE_MAX_AUTO_REDEEM_ACCOUNTS_PER_USER: "2" }
    })
    await ksGifts.setAutomaticRedemption({
      discordUserId: owner,
      guildId: guild,
      playerId: ksAccount.player_id,
      enabled: true
    })
    assert.equal((await ksGifts.status({ discordUserId: owner }))[0].gift_redemption_enabled, true)

    assert.equal((await pool.query(
      `SELECT COUNT(*)::integer AS count FROM gift_code_engagement_events
        WHERE game_profile = 'wos' AND guild_id = $1
          AND event_type = 'auto_redeem_join'`,
      [guild]
    )).rows[0].count, 1, "first-enable announcement was not idempotent")

    const firstSubmission = await giftRepository.recordSubmission({
      code: "NewCode",
      submittedByDiscordUserId: owner,
      metadata: { guildId: guild, submissionKind: "user" }
    })
    await giftRepository.recordSubmission({
      code: "NewCode",
      submittedByDiscordUserId: "222222222222222222",
      metadata: { guildId: "666666666666666666", submissionKind: "user" }
    })
    await pool.query(
      `UPDATE gift_codes SET status = 'active', verification_state = 'complete',
          verified_at_utc = now() WHERE id = $1`,
      [firstSubmission.giftCode.id]
    )
    const prepared = await wosCommunity.prepareCodeEngagement(firstSubmission.giftCode.id, 6)
    assert.equal(prepared.length, 2)
    assert.ok(prepared.every(event => event.discord_user_id === owner))
    assert.ok(prepared.every(event => event.guild_id === guild))
    assert.deepEqual(await wosCommunity.prepareCodeEngagement(firstSubmission.giftCode.id, 6), [])

    for (const [code, status] of [["BadCode", "invalid"], ["OldCode", "expired"]]) {
      const submission = await giftRepository.recordSubmission({
        code,
        submittedByDiscordUserId: owner,
        metadata: { guildId: guild }
      })
      await pool.query("UPDATE gift_codes SET status = $2 WHERE id = $1", [submission.giftCode.id, status])
      assert.deepEqual(await wosCommunity.prepareCodeEngagement(submission.giftCode.id, 0), [])
    }

    const progressEvent = prepared.find(event => event.event_type === "code_progress")
    await pool.query(
      `UPDATE gift_code_engagement_events
          SET status = 'completed', channel_id = '888888888888888888',
              message_id = '555555555555555555', progress_remaining = 6,
              last_update_at_utc = $2
        WHERE id = $1`,
      [progressEvent.id, new Date("2026-08-11T10:00:00Z")]
    )
    for (let index = 1; index < 5; index += 1) {
      assert.equal(await wosCommunity.claimProgressRefresh(
        firstSubmission.giftCode.id,
        "success",
        "progress-worker",
        new Date("2026-08-11T10:00:10Z")
      ), null)
    }
    const refresh = await wosCommunity.claimProgressRefresh(
      firstSubmission.giftCode.id,
      "already_redeemed",
      "progress-worker",
      new Date("2026-08-11T10:00:10Z")
    )
    assert.ok(refresh)
    assert.equal(refresh.progress.successful, 4)
    assert.equal(refresh.progress.already_redeemed, 1)
    assert.equal(refresh.progress.remaining, 1)

    const stats = await wosCommunity.communityStats(guild)
    assert.equal(stats.registered_users, 1)
    assert.equal(stats.registered_accounts, 4)
    assert.equal(stats.enabled_accounts, 2)
    assert.equal(stats.verified_codes, 1)
    assert.equal(stats.latest_verified_code, "NewCode")
  } finally {
    await pool?.end().catch(() => {})
    await admin.query(`DROP SCHEMA IF EXISTS "${schema}" CASCADE`).catch(() => {})
    await admin.end()
  }
})
