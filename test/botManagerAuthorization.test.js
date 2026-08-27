const assert = require("node:assert/strict")
const test = require("node:test")
const { PermissionFlagsBits } = require("discord.js")

const { createBotManagerAuthorizer } = require("../src/botManagerAuthorization")

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
