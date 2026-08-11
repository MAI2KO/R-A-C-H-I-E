const test = require("node:test")
const assert = require("node:assert/strict")
const { PermissionFlagsBits } = require("discord.js")

const {
  CONTRIBUTOR_ROLE_NAME,
  createGiftCodeCommunityService
} = require("../src/giftCodes/communityService")

function rewardRole(id, { editable = true } = {}) {
  return {
    id,
    name: CONTRIBUTOR_ROLE_NAME,
    editable,
    managed: false,
    mentionable: false,
    permissions: { bitfield: 0n }
  }
}

function discordRoles({ existingRole = null, manageRoles = true, createdEditable = true } = {}) {
  const state = { manageRoles, createdEditable }
  const roles = new Map(existingRole ? [[existingRole.id, existingRole]] : [])
  const createInputs = []
  const assigned = []
  const member = {
    roles: {
      cache: { has: roleId => assigned.includes(roleId) },
      async add(role) { assigned.push(role.id) }
    }
  }
  const guild = {
    members: {
      me: {
        permissions: {
          has(permission) {
            assert.equal(permission, PermissionFlagsBits.ManageRoles)
            return state.manageRoles
          }
        }
      },
      async fetch() { return member }
    },
    roles: {
      async fetch(roleId) { return roleId ? roles.get(roleId) || null : roles },
      async create(input) {
        createInputs.push(input)
        const role = rewardRole(`created-${createInputs.length}`, {
          editable: state.createdEditable
        })
        roles.set(role.id, role)
        return role
      }
    }
  }
  return {
    state,
    roles,
    createInputs,
    assigned,
    client: { guilds: { async fetch() { return guild } } }
  }
}

function rewardRepository({ eventCount = 1, contributorRoleId = null } = {}) {
  const events = Array.from({ length: eventCount }, (_, index) => ({
    id: `${index + 1}`.padStart(8, "0") + "-1111-4111-8111-111111111111",
    event_type: "contributor_role",
    guild_id: "777777777777777777",
    status: "pending"
  }))
  const settings = {
    guild_id: "777777777777777777",
    contributor_role_id: contributorRoleId,
    contributor_role_status: contributorRoleId ? "ready" : "unconfigured"
  }
  let provisionOwner = null
  const failures = []
  const completions = []
  return {
    events,
    settings,
    failures,
    completions,
    async claimEvent(eventId, workerId, now) {
      const event = events.find(value => value.id === eventId)
      if (!event || event.status === "completed") return null
      if (event.status === "failed" && event.nextAttemptAt > now) return null
      event.status = "claimed"
      event.claimed_by_worker = workerId
      return { ...event }
    },
    async getEventPayload(eventId) {
      const event = events.find(value => value.id === eventId)
      return {
        ...event,
        discord_user_id: `99999999999999999${events.indexOf(event)}`,
        contributor_role_id: settings.contributor_role_id
      }
    },
    async getSettings() { return { ...settings } },
    async claimContributorRoleProvision(_guildId, workerId) {
      if (provisionOwner && provisionOwner !== workerId) return null
      provisionOwner = workerId
      settings.contributor_role_status = "claiming"
      return { ...settings }
    },
    async completeContributorRoleProvision(_guildId, workerId, roleId) {
      if (provisionOwner !== workerId) return null
      settings.contributor_role_id = roleId
      settings.contributor_role_status = "ready"
      settings.contributor_role_last_error = null
      provisionOwner = null
      return { ...settings }
    },
    async failContributorRoleProvision(_guildId, workerId, errorCode) {
      if (provisionOwner === workerId) provisionOwner = null
      settings.contributor_role_status = "error"
      settings.contributor_role_last_error = errorCode
    },
    async markContributorRoleUnavailable(_guildId, errorCode) {
      settings.contributor_role_status = "error"
      settings.contributor_role_last_error = errorCode
    },
    async completeEvent(eventId) {
      const event = events.find(value => value.id === eventId)
      event.status = "completed"
      completions.push(eventId)
      return event
    },
    async failEvent(eventId, _workerId, errorCode, _now, { retryAt }) {
      const event = events.find(value => value.id === eventId)
      event.status = "failed"
      event.nextAttemptAt = retryAt
      failures.push({ eventId, errorCode, retryAt })
      return event
    }
  }
}

function rewardService(repository, discord, { workerId = "reward-worker", now } = {}) {
  return createGiftCodeCommunityService({
    repository,
    client: discord.client,
    gameProfile: "wos",
    workerId,
    now: now || (() => new Date("2026-08-11T12:00:00Z")),
    logger: { warn() {} }
  })
}

test("bot creates one persisted zero-permission non-mentionable candy role", async () => {
  const repository = rewardRepository()
  const discord = discordRoles()
  await rewardService(repository, discord).onAutoRedemptionEnabled(repository.events[0])
  assert.equal(discord.createInputs.length, 1)
  assert.deepEqual(discord.createInputs[0], {
    name: "🍭",
    permissions: [],
    mentionable: false,
    reason: "Gift-code contributor reward"
  })
  assert.equal(repository.settings.contributor_role_id, "created-1")
  assert.deepEqual(discord.assigned, ["created-1"])
  assert.equal(repository.completions.length, 1)
})

test("persisted contributor role is reused without creation", async () => {
  const existing = rewardRole("456")
  const repository = rewardRepository({ contributorRoleId: existing.id })
  const discord = discordRoles({ existingRole: existing })
  await rewardService(repository, discord).onAutoRedemptionEnabled(repository.events[0])
  assert.equal(discord.createInputs.length, 0)
  assert.deepEqual(discord.assigned, [existing.id])
})

test("concurrent rewards and a deleted configured role create at most one replacement", async () => {
  for (const contributorRoleId of [null, "deleted-role-id"]) {
    const repository = rewardRepository({ eventCount: 2, contributorRoleId })
    const discord = discordRoles()
    const first = rewardService(repository, discord, { workerId: "reward-a" })
    const second = rewardService(repository, discord, { workerId: "reward-b" })
    await Promise.all([
      first.onAutoRedemptionEnabled(repository.events[0]),
      second.onAutoRedemptionEnabled(repository.events[1])
    ])
    assert.equal(discord.createInputs.length, 1)
    assert.equal(repository.settings.contributor_role_id, "created-1")
    assert.ok(repository.completions.length + repository.failures.length === 2)
  }
})

test("missing Manage Roles is contained and a later retry awards after correction", async () => {
  let currentTime = new Date("2026-08-11T12:00:00Z")
  const repository = rewardRepository()
  const discord = discordRoles({ manageRoles: false })
  const service = rewardService(repository, discord, { now: () => currentTime })
  assert.equal(await service.onAutoRedemptionEnabled(repository.events[0]), null)
  assert.equal(discord.createInputs.length, 0)
  assert.equal(repository.failures.length, 1)
  assert.equal(repository.completions.length, 0)
  assert.equal(repository.settings.contributor_role_status, "error")

  discord.state.manageRoles = true
  currentTime = new Date("2026-08-11T12:06:00Z")
  await service.onAutoRedemptionEnabled(repository.events[0])
  assert.equal(discord.createInputs.length, 1)
  assert.equal(repository.completions.length, 1)
  assert.deepEqual(discord.assigned, ["created-1"])
})

test("hierarchy failure retains the created role and retries without another creation", async () => {
  let currentTime = new Date("2026-08-11T12:00:00Z")
  const repository = rewardRepository()
  const discord = discordRoles({ createdEditable: false })
  const service = rewardService(repository, discord, { now: () => currentTime })
  assert.equal(await service.onAutoRedemptionEnabled(repository.events[0]), null)
  assert.equal(discord.createInputs.length, 1)
  assert.equal(repository.settings.contributor_role_id, "created-1")
  assert.equal(repository.completions.length, 0)

  discord.roles.get("created-1").editable = true
  currentTime = new Date("2026-08-11T12:06:00Z")
  await service.onAutoRedemptionEnabled(repository.events[0])
  assert.equal(discord.createInputs.length, 1)
  assert.equal(repository.completions.length, 1)
  assert.deepEqual(discord.assigned, ["created-1"])
})
