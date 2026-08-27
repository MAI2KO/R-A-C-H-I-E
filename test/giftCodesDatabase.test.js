const test = require("node:test")
const assert = require("node:assert/strict")
const { Pool } = require("pg")

const { runMigrations } = require("../src/migrate")
const { createPlayerRepository } = require("../src/giftCodes/playerRepository")
const { createPlayerService, PlayerAccountError } = require("../src/giftCodes/playerService")
const { createGiftCodeRepository } = require("../src/giftCodes/repository")

const databaseUrl = process.env.TEST_DATABASE_URL

test("player and gift-code PostgreSQL foundation is profile scoped, transactional and durable", {
  skip: databaseUrl ? false : "TEST_DATABASE_URL is not configured"
}, async () => {
  const admin = new Pool({ connectionString: databaseUrl, max: 1 })
  const schema = `gift_codes_${process.pid}_${Date.now()}`
  let pool
  try {
    await admin.query(`CREATE SCHEMA "${schema}"`)
    pool = new Pool({
      connectionString: databaseUrl,
      max: 6,
      options: `-c search_path=${schema}`
    })
    const first = await runMigrations({ pool, logger: { log() {}, error() {} } })
    const second = await runMigrations({ pool, logger: { log() {}, error() {} } })
    assert.equal(first.applied.length, 21)
    assert.equal(first.applied.at(-1), "021_native_bot_manager_role.sql")
    assert.deepEqual(second.applied, [])

    const column = (await pool.query(
      `SELECT data_type, is_nullable
         FROM information_schema.columns
        WHERE table_schema = $1 AND table_name = 'player_accounts' AND column_name = 'id'`,
      [schema]
    )).rows[0]
    assert.deepEqual(column, { data_type: "uuid", is_nullable: "NO" })
    const indexes = (await pool.query(
      `SELECT indexname FROM pg_indexes
        WHERE schemaname = $1 AND tablename IN ('player_accounts', 'gift_code_redemptions')`,
      [schema]
    )).rows.map(row => row.indexname)
    assert.ok(indexes.includes("player_accounts_one_active_primary_idx"))
    assert.ok(indexes.includes("gift_code_redemptions_work_idx"))

    const wosRepository = createPlayerRepository(pool, "wos")
    const kingshotRepository = createPlayerRepository(pool, "kingshot")
    const wos = createPlayerService({ repository: wosRepository, gameProfile: "wos" })
    const kingshot = createPlayerService({ repository: kingshotRepository, gameProfile: "kingshot" })
    const owner = "111111111111111111"

    const wosPrimary = await wos.register({
      discordUserId: owner,
      playerId: "12345",
      locationNumber: "689", inGameName: "WOS Primary", allianceAbbreviation: "TST"
    })
    const wosSecondary = await wos.register({
      discordUserId: owner,
      playerId: "23456",
      locationNumber: "700", inGameName: "WOS Secondary", allianceAbbreviation: "TST"
    })
    const kingshotPrimary = await kingshot.register({
      discordUserId: owner,
      playerId: "12345",
      locationNumber: "521", inGameName: "King Primary", allianceAbbreviation: "TST"
    })
    assert.equal(wosPrimary.is_primary, true)
    assert.equal(wosSecondary.is_primary, false)
    assert.equal(kingshotPrimary.is_primary, true)
    assert.equal((await wos.view({ discordUserId: owner })).length, 2)
    assert.equal((await kingshot.view({ discordUserId: owner })).length, 1)
    await assert.rejects(
      wos.register({ discordUserId: "222222222222222222", playerId: "12345", locationNumber: "999",
        inGameName: "Duplicate", allianceAbbreviation: "TST" }),
      error => error instanceof PlayerAccountError && error.code === "PLAYER_ALREADY_REGISTERED"
    )

    assert.equal(
      await wosRepository.getOwnedAccount("222222222222222222", "12345"),
      null,
      "another Discord user accessed an account"
    )
    assert.equal(await wosRepository.updateLocation({
      discordUserId: "222222222222222222",
      playerId: "12345",
      newNumber: "701"
    }), null)

    const moved = await wos.changeLocation({
      discordUserId: owner,
      playerId: "12345",
      locationNumber: "700"
    })
    assert.equal(moved.previousNumber, "689")
    assert.equal(moved.account.state_or_kingdom_number, "700")
    const history = (await pool.query(
      `SELECT previous_number, new_number, change_source
         FROM player_location_history
        WHERE player_account_id = $1 AND game_profile = 'wos'
        ORDER BY changed_at_utc, id`,
      [wosPrimary.id]
    )).rows
    assert.deepEqual(history.map(row => [row.previous_number, row.new_number]), [
      [null, "689"],
      ["689", "700"]
    ])
    assert.ok(history.every(row => row.change_source === "user_command"))

    await pool.query(`
      CREATE FUNCTION reject_test_location_history() RETURNS trigger AS $$
      BEGIN
        IF NEW.new_number = '999' THEN
          RAISE EXCEPTION 'test history failure';
        END IF;
        RETURN NEW;
      END;
      $$ LANGUAGE plpgsql;
      CREATE TRIGGER reject_test_location_history_trigger
        BEFORE INSERT ON player_location_history
        FOR EACH ROW EXECUTE FUNCTION reject_test_location_history();
    `)
    await assert.rejects(wos.changeLocation({
      discordUserId: owner,
      playerId: "12345",
      locationNumber: "999"
    }), /test history failure/)
    assert.equal(
      (await wosRepository.getOwnedAccount(owner, "12345")).state_or_kingdom_number,
      "700",
      "account update was not rolled back with history failure"
    )
    await pool.query("DROP TRIGGER reject_test_location_history_trigger ON player_location_history")
    await pool.query("DROP FUNCTION reject_test_location_history()")

    const gifts = createGiftCodeRepository(pool, "wos")
    const trustedSource = await gifts.createSource({
      sourceType: "official",
      sourceName: "Official feed",
      trusted: true
    })
    const firstCode = await gifts.recordSubmission({
      code: "GiftCode",
      submittedByDiscordUserId: owner,
      sourceId: trustedSource.id
    })
    const duplicateCode = await gifts.recordSubmission({ code: "GiftCode", submittedByDiscordUserId: owner })
    const differentCase = await gifts.recordSubmission({ code: "giftcode", submittedByDiscordUserId: owner })
    assert.equal(firstCode.duplicate, false)
    assert.equal(firstCode.giftCode.status, "candidate")
    assert.equal(trustedSource.trusted, true)
    assert.equal(duplicateCode.duplicate, true)
    assert.equal(differentCase.duplicate, false)
    assert.notEqual(firstCode.giftCode.id, differentCase.giftCode.id)

    const queued = await gifts.queueRedemption({
      giftCodeId: firstCode.giftCode.id,
      playerAccountId: wosPrimary.id,
      botInstanceName: "rachie-wos"
    })
    const duplicateQueue = await gifts.queueRedemption({
      giftCodeId: firstCode.giftCode.id,
      playerAccountId: wosPrimary.id,
      botInstanceName: "rachie-wos"
    })
    assert.equal(queued.created, true)
    assert.equal(duplicateQueue.created, false)
    assert.equal(String(duplicateQueue.redemption.id), String(queued.redemption.id))
    assert.equal(queued.redemption.location_number_snapshot, "700")

    await wos.changeLocation({
      discordUserId: owner,
      playerId: "12345",
      locationNumber: "702"
    })
    const snapshot = (await pool.query(
      "SELECT player_id_snapshot, location_number_snapshot FROM gift_code_redemptions WHERE id = $1",
      [queued.redemption.id]
    )).rows[0]
    assert.deepEqual(snapshot, {
      player_id_snapshot: "12345",
      location_number_snapshot: "700"
    })

    const removal = await wos.remove({ discordUserId: owner, playerId: "12345" })
    assert.equal(removal.account.is_active, false)
    assert.equal(removal.replacement.player_id, "23456")
    assert.equal(removal.replacement.is_primary, true)
    assert.equal((await wos.view({ discordUserId: owner })).length, 1)
    assert.equal((await wosRepository.getOwnedAccount(owner, "12345", {
      includeInactive: true
    })).is_active, false)
    assert.equal((await pool.query(
      "SELECT COUNT(*)::integer AS count FROM player_location_history WHERE player_account_id = $1",
      [wosPrimary.id]
    )).rows[0].count, 3)
    assert.equal((await pool.query(
      "SELECT COUNT(*)::integer AS count FROM gift_code_redemptions WHERE player_account_id = $1",
      [wosPrimary.id]
    )).rows[0].count, 1)
  } finally {
    await pool?.end().catch(() => {})
    await admin.query(`DROP SCHEMA IF EXISTS "${schema}" CASCADE`).catch(() => {})
    await admin.end()
  }
})
