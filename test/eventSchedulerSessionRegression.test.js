const test = require("node:test")
const assert = require("node:assert/strict")

const { handleEventSchedulerInteraction } = require("../src/eventSchedulerInteractions")
const { InteractionSessionStore } = require("../src/interactionSessions")

const baseContext = {
  userId: "2345678901234567",
  guildId: "1234567890123456",
  gameProfile: "wos"
}

function interaction(type, { commandName = null, customId = null, values = [], context = {} } = {}) {
  return {
    commandName,
    customId,
    values,
    user: { id: context.userId || baseContext.userId },
    guildId: context.guildId || baseContext.guildId,
    client: {},
    deferred: false,
    replied: false,
    isChatInputCommand: () => type === "command",
    isButton: () => type === "button",
    isStringSelectMenu: () => type === "select",
    isChannelSelectMenu: () => false,
    isModalSubmit: () => false,
    async deferReply() { this.deferred = true },
    async deferUpdate() { this.deferred = true },
    async editReply(payload) { this.edited = payload },
    async reply(payload) { this.replied = true; this.replyPayload = payload },
    async showModal(modal) { this.modal = modal.toJSON() }
  }
}

function repository() {
  const settings = { event_channel_id: "3456789012345678" }
  return {
    async getGuildSettings() { return settings },
    async getStateLink() { return null },
    async getStateDestination() { return null },
    async listAlliances() {
      return {
        alliances: [{ id: "42", alliance_name: "HwC", is_default: true }],
        total: 1
      }
    }
  }
}

function options(sessionStore, repo = repository(), gameProfile = "wos") {
  return {
    creationSessionStore: sessionStore,
    healthProvider: () => ({
      available: true,
      gameProfile,
      botInstanceName: gameProfile === "wos" ? "rachie-wos" : "peggie-kingshot"
    }),
    userCanManageServer: async () => true,
    repositoryProvider: () => repo,
    logger: { error() {}, warn() {} }
  }
}

async function openAllianceSelector(sessionStore) {
  const repo = repository()
  const slash = interaction("command", { commandName: "event-scheduler" })
  await handleEventSchedulerInteraction(slash, options(sessionStore, repo))
  assert.match(slash.edited.content, /Event scheduler/)

  const create = interaction("button", { customId: "ec:new" })
  await handleEventSchedulerInteraction(create, options(sessionStore, repo))
  assert.match(create.edited.content, /Select alliance or sub-alliance/)
  const menu = create.edited.components[0].components[0].toJSON()
  return { repo, customId: menu.custom_id, token: menu.options[0].value }
}

test("event scheduler alliance selection continues with a valid setup session", async () => {
  const sessions = new InteractionSessionStore()
  const opened = await openAllianceSelector(sessions)
  const selected = interaction("select", {
    customId: opened.customId,
    values: [opened.token]
  })
  await handleEventSchedulerInteraction(selected, options(sessions, opened.repo))
  assert.match(selected.modal.custom_id, /^ec:m:/)
  assert.equal(selected.deferred, false)
  assert.equal(selected.replyPayload, undefined)
})

test("missing process-local setup state is recovered from the opaque alliance token", async () => {
  const opened = await openAllianceSelector(new InteractionSessionStore())
  const replacementProcessSessions = new InteractionSessionStore()
  const selected = interaction("select", {
    customId: opened.customId,
    values: [opened.token]
  })
  await handleEventSchedulerInteraction(
    selected,
    options(replacementProcessSessions, opened.repo)
  )
  assert.match(selected.modal.custom_id, /^ec:m:/)
  assert.doesNotMatch(selected.modal.custom_id, new RegExp(opened.customId.slice("ec:as:".length)))
})

test("expired and cross-context scheduler sessions remain rejected", async () => {
  let currentTime = 1000
  const expiredSessions = new InteractionSessionStore({ ttlMs: 10, now: () => currentTime })
  const expired = await openAllianceSelector(expiredSessions)
  currentTime = 1011
  expiredSessions.cleanup()
  const expiredSelection = interaction("select", {
    customId: expired.customId,
    values: [expired.token]
  })
  await handleEventSchedulerInteraction(expiredSelection, options(expiredSessions, expired.repo))
  assert.match(expiredSelection.replyPayload.content, /setup has expired/)

  for (const [context, profile, message] of [
    [{ userId: "9999999999999999" }, "wos", /another user/],
    [{ guildId: "9999999999999999" }, "wos", /another Discord server/],
    [{}, "kingshot", /another game profile/]
  ]) {
    const sessions = new InteractionSessionStore()
    const opened = await openAllianceSelector(sessions)
    const selected = interaction("select", {
      customId: opened.customId,
      values: [opened.token],
      context
    })
    await handleEventSchedulerInteraction(selected, options(sessions, opened.repo, profile))
    assert.match(selected.replyPayload.content, message)
  }
})

test("malformed and repeated alliance selectors are handled without duplicate acknowledgement", async () => {
  const sessions = new InteractionSessionStore()
  const malformed = interaction("select", {
    customId: "ec:as:missing-session",
    values: ["not-an-opaque-alliance-token"]
  })
  await handleEventSchedulerInteraction(malformed, options(sessions))
  assert.match(malformed.replyPayload.content, /selection has expired/)
  assert.equal(malformed.deferred, false)

  const opened = await openAllianceSelector(sessions)
  for (let attempt = 0; attempt < 2; attempt += 1) {
    const selected = interaction("select", {
      customId: opened.customId,
      values: [opened.token]
    })
    await handleEventSchedulerInteraction(selected, options(sessions, opened.repo))
    assert.match(selected.modal.custom_id, /^ec:m:/)
    assert.equal(selected.deferred, false)
    assert.equal(selected.replyPayload, undefined)
  }
})
