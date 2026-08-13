const test = require("node:test")
const assert = require("node:assert/strict")
const { Pool } = require("pg")
const { Collection, MessageReferenceType } = require("discord.js")

const { runMigrations } = require("../src/migrate")
const { createGiftCodeRepository } = require("../src/giftCodes/repository")
const { createGiftCodeService } = require("../src/giftCodes/service")
const { createGiftCodeCommunityRepository } = require("../src/giftCodes/communityRepository")
const { createGiftCodeSourceRepository } = require("../src/giftCodes/sourceRepository")
const { createGiftCodeSourceIngestionService } = require("../src/giftCodes/sourceIngestion")

const databaseUrl = process.env.TEST_DATABASE_URL

test("gift-code source migration, provenance, attribution and profile isolation are durable", {
  skip: databaseUrl ? false : "TEST_DATABASE_URL is not configured"
}, async () => {
  const admin = new Pool({ connectionString: databaseUrl, max: 1 })
  const schema = `gift_sources_${process.pid}_${Date.now()}`
  let pool
  try {
    await admin.query(`CREATE SCHEMA "${schema}"`)
    pool = new Pool({
      connectionString: databaseUrl,
      max: 8,
      options: `-c search_path=${schema}`
    })
    const first = await runMigrations({ pool, logger: { log() {}, error() {} } })
    const second = await runMigrations({ pool, logger: { log() {}, error() {} } })
    assert.equal(first.applied.length, 17)
    assert.equal(first.applied.at(-1), "017_bot_managed_discord_setup.sql")
    assert.deepEqual(second.applied, [])

    const columns = (await pool.query(
      `SELECT column_name FROM information_schema.columns
        WHERE table_schema = $1 AND table_name = 'gift_code_source_observations'`,
      [schema]
    )).rows.map(row => row.column_name)
    assert.ok(columns.includes("source_reported_expiry_at_utc"))
    assert.ok(columns.includes("no_longer_observed_at_utc"))
    assert.ok(columns.includes("provenance"))

    const wosGifts = createGiftCodeRepository(pool, "wos")
    const wosSources = createGiftCodeSourceRepository(pool, "wos")
    const wosIngestion = createGiftCodeSourceIngestionService({
      giftRepository: wosGifts,
      sourceRepository: wosSources,
      gameProfile: "wos"
    })
    const kingshotGifts = createGiftCodeRepository(pool, "kingshot")
    const kingshotSources = createGiftCodeSourceRepository(pool, "kingshot")
    const kingshotIngestion = createGiftCodeSourceIngestionService({
      giftRepository: kingshotGifts,
      sourceRepository: kingshotSources,
      gameProfile: "kingshot"
    })

    const concurrentSources = await Promise.all(Array.from({ length: 5 }, () =>
      wosSources.ensureSource({ sourceType: "public_catalogue", sourceName: "Concurrent source" })
    ))
    assert.equal(new Set(concurrentSources.map(source => source.id)).size, 1)
    assert.equal((await pool.query(
      `SELECT COUNT(*)::integer AS count FROM gift_code_sources
        WHERE game_profile = 'wos' AND source_type = 'public_catalogue'
          AND source_name = 'Concurrent source'`
    )).rows[0].count, 1)

    await wosSources.configureDiscordChannel({ guildId: "111", channelId: "222" })
    await wosSources.configureDiscordChannel({ guildId: "999", channelId: "888" })
    await kingshotSources.configureDiscordChannel({ guildId: "111", channelId: "222" })
    assert.equal((await wosSources.discordChannel("222", "111")).source_type, "discord_mirror")
    assert.equal((await kingshotSources.discordChannel("222", "111")).source_type, "discord_mirror")
    const scopedStatus = await wosSources.sourceStatus("111")
    assert.deepEqual(scopedStatus.channels.map(row => row.guild_id), ["111"])

    const sourceExpiry = new Date("2026-08-20T23:59:00Z")
    const discordMessage = {
      guildId: "111",
      channelId: "222",
      id: "333",
      webhookId: "444",
      content: "Code: WosExactCase\nValid Until: August 20, 2026 23:59 (UTC+0)",
      createdTimestamp: Date.parse("2026-08-12T10:00:00Z"),
      author: { username: "WOS mirror" }
    }
    const observed = await wosIngestion.ingestDiscordMessage(discordMessage)
    const replay = await wosIngestion.ingestDiscordMessage(discordMessage)
    assert.equal(observed.giftCode.code, "WosExactCase")
    assert.equal(observed.giftCode.status, "candidate")
    assert.equal(replay.duplicateObservation, true)
    assert.equal((await pool.query(
      "SELECT COUNT(*)::integer AS count FROM gift_code_submissions WHERE game_profile = 'wos' AND submitted_code = 'WosExactCase'"
    )).rows[0].count, 1)
    const provenance = (await pool.query(
      `SELECT observed_at_utc, source_reported_expiry_at_utc, provenance
         FROM gift_code_source_observations
        WHERE game_profile = 'wos' AND observed_code = 'WosExactCase'`
    )).rows[0]
    assert.equal(provenance.observed_at_utc.toISOString(), "2026-08-12T10:00:00.000Z")
    assert.equal(provenance.source_reported_expiry_at_utc.toISOString(), sourceExpiry.toISOString())
    assert.equal(provenance.provenance.guildId, "111")
    assert.equal(provenance.provenance.channelId, "222")
    assert.equal(provenance.provenance.messageId, "333")
    assert.equal(provenance.provenance.webhookId, "444")
    assert.equal((await wosGifts.getCode("WosExactCase")).expires_at_utc, null)

    const forwardedMessage = {
      guildId: "111",
      channelId: "222",
      id: "forwarded-333",
      webhookId: "444",
      content: "Outer forward text",
      createdTimestamp: Date.parse("2026-08-12T11:00:00Z"),
      author: { username: "WOS mirror" },
      reference: {
        type: MessageReferenceType.Forward,
        guildId: "original-guild",
        channelId: "original-channel",
        messageId: "original-message"
      },
      messageSnapshots: new Collection([["original-message", {
        content: "Gift Code: ForwardedCase7",
        embeds: []
      }]])
    }
    const forwarded = await wosIngestion.ingestDiscordMessage(forwardedMessage)
    const forwardedDuplicate = await wosIngestion.ingestDiscordMessage({
      ...forwardedMessage,
      id: "forwarded-334",
      createdTimestamp: Date.parse("2026-08-12T11:05:00Z")
    })
    assert.equal(forwarded.duplicate, false)
    assert.equal(forwardedDuplicate.duplicate, true)
    assert.equal((await pool.query(
      "SELECT COUNT(*)::integer AS count FROM gift_codes WHERE game_profile = 'wos' AND code = 'ForwardedCase7'"
    )).rows[0].count, 1)
    assert.equal((await pool.query(
      `SELECT COUNT(*)::integer AS count
         FROM gift_code_attempts a
         JOIN gift_codes c ON c.id = a.gift_code_id
        WHERE c.game_profile = 'wos' AND c.code = 'ForwardedCase7'`
    )).rows[0].count, 0)
    assert.equal((await pool.query(
      `SELECT COUNT(*)::integer AS count
         FROM gift_code_redemptions r
         JOIN gift_codes c ON c.id = r.gift_code_id
        WHERE c.game_profile = 'wos' AND c.code = 'ForwardedCase7'`
    )).rows[0].count, 0)
    const forwardedRows = (await pool.query(
      `SELECT observed_at_utc, provenance
         FROM gift_code_source_observations
        WHERE game_profile = 'wos' AND observed_code = 'ForwardedCase7'
        ORDER BY observed_at_utc`
    )).rows
    assert.equal(forwardedRows.length, 2)
    assert.equal(forwardedRows[0].provenance.messageId, "forwarded-333")
    assert.equal(forwardedRows[0].provenance.sourceMessageKind, "forwarded_snapshot")
    assert.equal(forwardedRows[0].provenance.forwardedSourceGuildId, "original-guild")
    assert.equal(forwardedRows[0].provenance.forwardedSourceChannelId, "original-channel")
    assert.equal(forwardedRows[0].provenance.forwardedSourceMessageId, "original-message")
    const mirrorHealth = (await pool.query(
      `SELECT last_observation_at_utc, last_candidate_at_utc
         FROM gift_code_sources WHERE id = $1`,
      [forwarded.source.id]
    )).rows[0]
    assert.equal(mirrorHealth.last_observation_at_utc.toISOString(), "2026-08-12T11:05:00.000Z")
    assert.equal(mirrorHealth.last_candidate_at_utc.toISOString(), "2026-08-12T11:00:00.000Z")

    const kingshotCandidate = await kingshotIngestion.ingest({
      source: { sourceType: "public_catalogue", sourceName: "KingshotRewards" },
      code: "WosExactCase",
      observationKey: "catalogue:WosExactCase",
      provenance: { transport: "http_catalogue" }
    })
    assert.equal(kingshotCandidate.duplicate, false)
    assert.notEqual(kingshotCandidate.giftCode.id, observed.giftCode.id)
    assert.equal((await wosGifts.getCode("WosExactCase")).game_profile, "wos")
    assert.equal((await kingshotGifts.getCode("WosExactCase")).game_profile, "kingshot")

    const catalogueSource = {
      sourceType: "public_catalogue",
      sourceName: "WOSRewards",
      sourceReference: "https://www.wosrewards.com/"
    }
    const catalogueFirst = await wosIngestion.ingest({
      source: catalogueSource,
      code: "CatalogueOne",
      observationKey: "catalogue:CatalogueOne",
      provenance: { transport: "http_catalogue" }
    })
    const catalogueReplay = await wosIngestion.ingest({
      source: catalogueSource,
      code: "CatalogueOne",
      observationKey: "catalogue:CatalogueOne",
      provenance: { transport: "http_catalogue" }
    })
    assert.equal(catalogueReplay.duplicateObservation, true)
    const catalogueSourceRow = catalogueReplay.source
    await wosSources.markCataloguePoll(catalogueSourceRow.id, {
      successful: true,
      observedCodes: [],
      now: new Date("2026-08-13T00:00:00Z")
    })
    assert.equal((await pool.query(
      `SELECT no_longer_observed_at_utc IS NOT NULL AS disappeared
         FROM gift_code_source_observations
        WHERE game_profile = 'wos' AND gift_code_id = $1`,
      [catalogueFirst.giftCode.id]
    )).rows[0].disappeared, true)
    assert.equal((await wosGifts.getCode("CatalogueOne")).status, "candidate")
    assert.equal((await pool.query(
      "SELECT COUNT(*)::integer AS count FROM gift_code_submissions WHERE game_profile = 'wos' AND submitted_code = 'CatalogueOne'"
    )).rows[0].count, 1)

    const manualService = createGiftCodeService({
      repository: wosGifts,
      gameProfile: "wos",
      ingestion: wosIngestion
    })
    const community = createGiftCodeCommunityRepository(pool, "wos")

    const humanFirst = await manualService.submit({
      discordUserId: "555",
      guildId: "111",
      code: "HumanFirst"
    })
    await pool.query(
      "UPDATE gift_codes SET status = 'active', verified_at_utc = now() WHERE id = $1",
      [humanFirst.giftCode.id]
    )
    const humanEvents = await community.prepareCodeEngagement(humanFirst.giftCode.id, 0)
    assert.equal(humanEvents.some(event => event.event_type === "contributor_role"), true)
    assert.equal(humanEvents.find(event => event.event_type === "contributor_role").discord_user_id, "555")

    const automatic = await wosIngestion.ingest({
      source: catalogueSource,
      code: "AutomaticOnly",
      observationKey: "catalogue:AutomaticOnly",
      provenance: { transport: "http_catalogue" }
    })
    await pool.query("UPDATE gift_codes SET status = 'active' WHERE id = $1", [automatic.giftCode.id])
    assert.deepEqual(await community.prepareCodeEngagement(automatic.giftCode.id, 0), [])

    await pool.query("UPDATE gift_codes SET status = 'active' WHERE id = $1", [forwarded.giftCode.id])
    assert.deepEqual(await community.prepareCodeEngagement(forwarded.giftCode.id, 0), [])

    const automaticFirst = await wosIngestion.ingest({
      source: catalogueSource,
      code: "AutomaticThenHuman",
      observationKey: "catalogue:AutomaticThenHuman",
      provenance: { transport: "http_catalogue" }
    })
    const lateHuman = await manualService.submit({
      discordUserId: "666",
      guildId: "111",
      code: "AutomaticThenHuman"
    })
    assert.equal(lateHuman.duplicate, true)
    await pool.query("UPDATE gift_codes SET status = 'active' WHERE id = $1", [automaticFirst.giftCode.id])
    assert.deepEqual(await community.prepareCodeEngagement(automaticFirst.giftCode.id, 0), [])
  } finally {
    await pool?.end().catch(() => {})
    await admin.query(`DROP SCHEMA IF EXISTS "${schema}" CASCADE`).catch(() => {})
    await admin.end()
  }
})
