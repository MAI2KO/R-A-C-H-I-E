const test = require("node:test")
const assert = require("node:assert/strict")
const fs = require("node:fs/promises")
const os = require("node:os")
const path = require("node:path")
const { Pool } = require("pg")

const { DEFAULT_MIGRATIONS_DIR, runMigrations } = require("../src/migrate")
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
    assert.equal(first.applied.length, 15)
    assert.equal(first.applied.at(-1), "015_gift_code_guild_enrolment.sql")
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
        currentlyEnabled[0].id,
        "success",
        "progress-worker",
        new Date("2026-08-11T10:00:10Z")
      ), null)
    }
    const refresh = await wosCommunity.claimProgressRefresh(
      firstSubmission.giftCode.id,
      currentlyEnabled[0].id,
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
    assert.equal(stats.auto_redeem_players, 1)
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

test("migration 014 reconciles the earlier migration 013 community schema", {
  skip: databaseUrl ? false : "TEST_DATABASE_URL is not configured"
}, async () => {
  const admin = new Pool({ connectionString: databaseUrl, max: 1 })
  const schema = `gift_community_reconcile_${process.pid}_${Date.now()}`
  const migrationsDir = await fs.mkdtemp(path.join(os.tmpdir(), "rachie-community-migrations-"))
  let pool
  try {
    await admin.query(`CREATE SCHEMA "${schema}"`)
    pool = new Pool({
      connectionString: databaseUrl,
      max: 8,
      options: `-c search_path=${schema}`
    })
    const files = (await fs.readdir(DEFAULT_MIGRATIONS_DIR))
      .filter(file => /^\d+.*\.sql$/.test(file) && file < "014_")
    await Promise.all(files.map(file => fs.copyFile(
      path.join(DEFAULT_MIGRATIONS_DIR, file),
      path.join(migrationsDir, file)
    )))
    assert.equal((await runMigrations({ pool, migrationsDir, logger })).applied.length, 13)

    // Reproduce the exact schema drift from the earlier applied 013 revision.
    await pool.query(`
      ALTER TABLE gift_code_guild_settings
        DROP CONSTRAINT gift_code_guild_settings_role_status_check,
        DROP COLUMN contributor_role_status,
        DROP COLUMN contributor_role_last_error,
        DROP COLUMN contributor_role_claimed_by,
        DROP COLUMN contributor_role_claimed_until_utc;
      DROP INDEX gift_code_engagement_pending_idx;
      ALTER TABLE gift_code_engagement_events DROP COLUMN next_attempt_at_utc;
      CREATE INDEX gift_code_engagement_pending_idx
        ON gift_code_engagement_events (game_profile, status, created_at_utc, id)
        WHERE status IN ('pending', 'claimed');
      INSERT INTO gift_code_guild_settings (
        game_profile, guild_id, contributor_role_id
      ) VALUES
        ('wos', '700000000000000001', '900000000000000001'),
        ('wos', '700000000000000002', NULL);
    `)
    await fs.copyFile(
      path.join(DEFAULT_MIGRATIONS_DIR, "014_gift_code_community_reconciliation.sql"),
      path.join(migrationsDir, "014_gift_code_community_reconciliation.sql")
    )

    const concurrent = await Promise.all([
      runMigrations({ pool, migrationsDir, logger }),
      runMigrations({ pool, migrationsDir, logger })
    ])
    assert.deepEqual(concurrent.map(result => result.applied.length).sort(), [0, 1])
    assert.deepEqual(concurrent.flatMap(result => result.applied), [
      "014_gift_code_community_reconciliation.sql"
    ])
    assert.deepEqual((await runMigrations({ pool, migrationsDir, logger })).applied, [])
    assert.equal((await pool.query(
      `SELECT COUNT(*)::integer AS count FROM schema_migrations
        WHERE version = '014_gift_code_community_reconciliation.sql'`
    )).rows[0].count, 1)

    const columns = (await pool.query(
      `SELECT table_name, column_name, is_nullable, column_default
         FROM information_schema.columns
        WHERE table_schema = $1
          AND table_name IN ('gift_code_guild_settings', 'gift_code_engagement_events')`,
      [schema]
    )).rows
    for (const name of [
      "contributor_role_status", "contributor_role_last_error",
      "contributor_role_claimed_by", "contributor_role_claimed_until_utc"
    ]) {
      assert.ok(columns.some(column =>
        column.table_name === "gift_code_guild_settings" && column.column_name === name
      ), `missing reconciled settings column ${name}`)
    }
    const roleStatus = columns.find(column =>
      column.table_name === "gift_code_guild_settings"
        && column.column_name === "contributor_role_status"
    )
    assert.equal(roleStatus.is_nullable, "NO")
    assert.match(roleStatus.column_default, /unconfigured/)
    const reconciledSettings = (await pool.query(
      `SELECT guild_id, contributor_role_status
         FROM gift_code_guild_settings
        WHERE guild_id IN ('700000000000000001', '700000000000000002')
        ORDER BY guild_id`
    )).rows
    assert.deepEqual(reconciledSettings, [
      { guild_id: "700000000000000001", contributor_role_status: "ready" },
      { guild_id: "700000000000000002", contributor_role_status: "unconfigured" }
    ])
    const roleConstraint = (await pool.query(
      `SELECT convalidated
         FROM pg_constraint
        WHERE conrelid = 'gift_code_guild_settings'::regclass
          AND conname = 'gift_code_guild_settings_role_status_check'`
    )).rows[0]
    assert.deepEqual(roleConstraint, { convalidated: true })
    assert.ok(columns.some(column =>
      column.table_name === "gift_code_engagement_events"
        && column.column_name === "next_attempt_at_utc"
    ))
    const pendingIndex = (await pool.query(
      `SELECT indexdef FROM pg_indexes
        WHERE schemaname = $1 AND indexname = 'gift_code_engagement_pending_idx'`,
      [schema]
    )).rows[0].indexdef
    assert.match(pendingIndex, /next_attempt_at_utc/)
    assert.match(pendingIndex, /failed/)

    await pool.query(`
      INSERT INTO player_accounts (
        id, game_profile, discord_user_id, player_id, state_or_kingdom_number
      ) VALUES (
        '11111111-1111-4111-8111-111111111111', 'wos',
        '100000000000000001', '200000001', '689'
      );
      INSERT INTO player_account_guilds (
        game_profile, guild_id, player_account_id
      ) VALUES (
        'wos', '700000000000000001', '11111111-1111-4111-8111-111111111111'
      );
      INSERT INTO gift_code_engagement_events (
        id, game_profile, guild_id, event_type, player_account_id, discord_user_id, status
      ) VALUES (
        '22222222-2222-4222-8222-222222222222', 'wos',
        '700000000000000001', 'auto_redeem_join',
        '11111111-1111-4111-8111-111111111111', '100000000000000001', 'completed'
      );
    `)

    await fs.copyFile(
      path.join(DEFAULT_MIGRATIONS_DIR, "015_gift_code_guild_enrolment.sql"),
      path.join(migrationsDir, "015_gift_code_guild_enrolment.sql")
    )
    assert.deepEqual((await runMigrations({ pool, migrationsDir, logger })).applied, [
      "015_gift_code_guild_enrolment.sql"
    ])
    assert.deepEqual((await runMigrations({ pool, migrationsDir, logger })).applied, [])
    const legacyEnrolment = (await pool.query(
      `SELECT gift_code_enrolled, gift_code_first_enabled_at_utc
         FROM player_account_guilds
        WHERE player_account_id = '11111111-1111-4111-8111-111111111111'`
    )).rows[0]
    assert.equal(legacyEnrolment.gift_code_enrolled, true)
    assert.ok(legacyEnrolment.gift_code_first_enabled_at_utc)

    const community = createGiftCodeCommunityRepository(pool, "wos")
    const gifts = createGiftCodeRepository(pool, "wos")
    const guildId = "777777777777777777"
    const ownerId = "111111111111111111"
    await community.setChannel(guildId, "888888888888888888")
    const submission = await gifts.recordSubmission({
      code: "ReconciledCode",
      submittedByDiscordUserId: ownerId,
      metadata: { guildId }
    })
    await pool.query(
      `UPDATE gift_codes SET status = 'active', verification_state = 'complete',
          verified_at_utc = now() WHERE id = $1`,
      [submission.giftCode.id]
    )
    const events = await community.prepareCodeEngagement(submission.giftCode.id, 2)
    assert.deepEqual(events.map(event => event.event_type).sort(), ["code_progress", "contributor_role"])

    const roleClaim = await community.claimContributorRoleProvision(
      guildId, "role-worker", new Date("2026-08-11T10:00:00Z")
    )
    assert.equal(roleClaim.contributor_role_status, "claiming")
    assert.equal((await community.completeContributorRoleProvision(
      guildId, "role-worker", "999999999999999999", new Date("2026-08-11T10:00:01Z")
    )).contributor_role_status, "ready")

    const roleEvent = events.find(event => event.event_type === "contributor_role")
    assert.ok(await community.claimEvent(
      roleEvent.id, "role-event-worker", new Date("2026-08-11T10:00:02Z")
    ))
    assert.equal((await community.completeEvent(roleEvent.id, "role-event-worker", {
      now: new Date("2026-08-11T10:00:03Z")
    })).status, "completed")

    const progressEvent = events.find(event => event.event_type === "code_progress")
    assert.ok(await community.claimEvent(
      progressEvent.id, "event-worker", new Date("2026-08-11T10:01:00Z")
    ))
    await community.failEvent(
      progressEvent.id,
      "event-worker",
      "TEST_RETRY",
      new Date("2026-08-11T10:01:01Z"),
      { retryAt: new Date("2026-08-11T10:01:02Z") }
    )
    const retry = await community.claimNextPending(
      "retry-worker", new Date("2026-08-11T10:01:03Z")
    )
    assert.equal(retry.id, progressEvent.id)
    const completed = await community.completeEvent(progressEvent.id, "retry-worker", {
      channelId: "888888888888888888",
      messageId: "555555555555555555",
      progressCount: 1,
      progress: { successful: 1, already_redeemed: 0, account_issues: 0, restricted: 0, remaining: 1 },
      now: new Date("2026-08-11T10:01:04Z")
    })
    assert.equal(completed.progress_successful, 1)
    assert.equal(completed.next_attempt_at_utc, null)
    assert.equal((await community.communityStats(guildId)).verified_codes, 1)
  } finally {
    await pool?.end().catch(() => {})
    await admin.query(`DROP SCHEMA IF EXISTS "${schema}" CASCADE`).catch(() => {})
    await admin.end()
    await fs.rm(migrationsDir, { recursive: true, force: true })
  }
})
