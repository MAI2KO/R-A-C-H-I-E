const assert = require("node:assert/strict")
const test = require("node:test")
const { PermissionFlagsBits } = require("discord.js")

const {
  createBotManagerAuthorizer,
  createLiveBotManagerVerifier,
  createLiveGuildOwnerVerifier
} = require("../src/botManagerAuthorization")

function interaction({ userId = "10", ownerId = "1", administrator = false, roles = [] } = {}) {
  return {
    guildId: "777",
    guild: { ownerId },
    user: { id: userId },
    memberPermissions: { has: permission => permission === PermissionFlagsBits.Administrator && administrator },
    member: { roles: { cache: { has: roleId => roles.includes(roleId) } } }
  }
}

test("native bot-manager authorization allows only owner, Administrator, or configured role", async () => {
  let lookups = 0
  const authorizer = createBotManagerAuthorizer({
    repositoryProvider: () => ({ async getManagerRole(guildId) {
      lookups++
      assert.equal(guildId, "777")
      return "42"
    } }),
    logger: { error() {} }
  })
  assert.equal(await authorizer.canManage(interaction({ userId: "1" })), true)
  assert.equal(await authorizer.canManage(interaction({ administrator: true })), true)
  assert.equal(lookups, 0, "owner and Administrator do not require a database round trip")
  assert.equal(await authorizer.canManage(interaction({ roles: ["42"] })), true)
  assert.equal(await authorizer.canManage(interaction({ roles: ["99"] })), false)
})

test("missing or failed PostgreSQL manager configuration fails closed", async () => {
  const missing = createBotManagerAuthorizer({ repositoryProvider: () => null })
  assert.equal(await missing.canManage(interaction()), false)
  const failed = createBotManagerAuthorizer({
    repositoryProvider: () => ({ async getManagerRole() { throw new Error("down") } }),
    logger: { error() {} }
  })
  assert.equal(await failed.canManage(interaction()), false)
})

function liveClient({ ownerId = "100000000000000001", administrator = false, roles = [] } = {}) {
  const member = {
    permissions: { has: permission => permission === PermissionFlagsBits.Administrator && administrator },
    roles: { cache: { has: roleId => roles.includes(roleId) } }
  }
  const guild = {
    ownerId,
    members: { fetch: async userId => {
      assert.equal(userId, "200000000000000002")
      return member
    } }
  }
  return { guilds: { cache: new Map([["300000000000000003", guild]]), fetch: async () => guild } }
}

test("live website decisions share native owner, Administrator, and PostgreSQL role authority", async () => {
  const failedRepository = () => ({ async getManagerRole() { throw new Error("down") } })
  for (const client of [
    liveClient({ ownerId: "200000000000000002" }),
    liveClient({ administrator: true })
  ]) {
    const result = await createLiveBotManagerVerifier({
      client, repositoryProvider: failedRepository, logger: { error() {} }
    })({ guildId: "300000000000000003", discordUserId: "200000000000000002" })
    assert.equal(result.status, "authorized")
    assert.equal(result.via, "administrator")
  }

  let configuredRole = "400000000000000004"
  const verify = createLiveBotManagerVerifier({
    client: liveClient({ roles: ["400000000000000004"] }),
    repositoryProvider: () => ({ async getManagerRole(guildId) {
      assert.equal(guildId, "300000000000000003")
      return configuredRole
    } }),
    logger: { error() {} }
  })
  assert.equal((await verify({ guildId: "300000000000000003",
    discordUserId: "200000000000000002" })).via, "bot_manager_role")
  configuredRole = null
  assert.equal((await verify({ guildId: "300000000000000003",
    discordUserId: "200000000000000002" })).status, "denied")
})

test("live role-only authorization fails closed when native storage is unavailable", async () => {
  const verify = createLiveBotManagerVerifier({
    client: liveClient({ roles: ["400000000000000004"] }),
    repositoryProvider: () => ({ async getManagerRole() { throw new Error("down") } }),
    logger: { error() {} }
  })
  assert.equal((await verify({ guildId: "300000000000000003",
    discordUserId: "200000000000000002" })).status, "unavailable")
})

test("guild ownership verification accepts only the exact owner and fails closed", async () => {
  const owner = createLiveGuildOwnerVerifier({
    client: liveClient({ ownerId: "200000000000000002", administrator: true,
      roles: ["400000000000000004"] }), logger: { error() {} }
  })
  assert.deepEqual(await owner({ guildId: "300000000000000003",
    discordUserId: "200000000000000002" }), { status: "owner" })
  const nonOwner = createLiveGuildOwnerVerifier({
    client: liveClient({ ownerId: "100000000000000001", administrator: true,
      roles: ["400000000000000004"] }), logger: { error() {} }
  })
  assert.deepEqual(await nonOwner({ guildId: "300000000000000003",
    discordUserId: "200000000000000002" }), { status: "not_owner" })
  const unavailable = createLiveGuildOwnerVerifier({
    client: { guilds: { cache: new Map(), fetch: async () => { throw new Error("down") } } },
    logger: { error() {} }
  })
  assert.deepEqual(await unavailable({ guildId: "300000000000000003",
    discordUserId: "200000000000000002" }), { status: "unavailable" })
})
