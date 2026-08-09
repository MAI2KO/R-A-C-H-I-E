const test = require("node:test")
const assert = require("node:assert/strict")
const fs = require("node:fs")
const path = require("node:path")
const { ChannelType, MessageFlags } = require("discord.js")

const {
  handleEventSchedulerInteraction
} = require("../src/eventSchedulerInteractions")
const {
  ALREADY_ACKNOWLEDGED,
  EXPIRED_INTERACTION,
  handleInteractionFailure,
  safelyRespondToInteraction,
  schedulerInteractionWasHandled
} = require("../src/interactionResponses")

function schedulerInteraction(type, overrides = {}) {
  const interaction = {
    commandName: null,
    customId: null,
    guildId: "1234567890123456",
    user: { id: "2345678901234567" },
    values: [],
    deferred: false,
    replied: false,
    isChatInputCommand: () => type === "command",
    isButton: () => type === "button",
    isStringSelectMenu: () => type === "string-select",
    isChannelSelectMenu: () => type === "channel-select",
    isModalSubmit: () => type === "modal",
    async deferReply(options) {
      this.acknowledgedWith = "deferReply"
      this.deferOptions = options
      this.deferred = true
    },
    async deferUpdate() {
      this.acknowledgedWith = "deferUpdate"
      this.deferred = true
    },
    async editReply(payload) { this.edited = payload },
    async reply(payload) { this.replyCount = (this.replyCount || 0) + 1; this.replied = true; this.replyPayload = payload },
    async followUp(payload) { this.followUpCount = (this.followUpCount || 0) + 1; this.followUpPayload = payload },
    async showModal(modal) { this.modal = modal },
    ...overrides
  }
  return interaction
}

function options(repository, logger = { error() {}, warn() {} }) {
  return {
    healthProvider: () => ({
      available: true,
      gameProfile: "wos",
      botInstanceName: "rachie-wos"
    }),
    async userCanManageServer(interaction) {
      assert.equal(interaction.deferred, true, "authorization ran before acknowledgement")
      await new Promise(resolve => setTimeout(resolve, 5))
      return true
    },
    repositoryProvider: () => repository,
    logger
  }
}

function delayed(interaction, value) {
  return async () => {
    assert.equal(interaction.deferred, true, "database work ran before acknowledgement")
    await new Promise(resolve => setTimeout(resolve, 5))
    return value
  }
}

test("scheduler slash command acknowledges before authorization and database loading", async () => {
  const interaction = schedulerInteraction("command", {
    commandName: "event-scheduler",
    client: {}
  })
  const repository = {
    getGuildSettings: delayed(interaction, null),
    getStateLink: delayed(interaction, null),
    getStateDestination: delayed(interaction, null)
  }
  assert.equal(await handleEventSchedulerInteraction(interaction, options(repository)), true)
  assert.equal(interaction.acknowledgedWith, "deferReply")
  assert.equal(interaction.deferOptions.flags, MessageFlags.Ephemeral)
  assert.match(interaction.edited.content, /Event scheduler/)
  assert.equal(interaction.replyCount || 0, 0)
})

test("management entry button acknowledges before its alliance-list query", async () => {
  const interaction = schedulerInteraction("button", { customId: "el:0" })
  const repository = {
    listAlliances: delayed(interaction, { alliances: [], total: 0 })
  }
  await handleEventSchedulerInteraction(interaction, options(repository))
  assert.equal(interaction.acknowledgedWith, "deferUpdate")
  assert.match(interaction.edited.content, /Select an alliance/)
})

test("channel selector acknowledges before validation and persistence", async () => {
  const channel = {
    id: "3456789012345678",
    guildId: "1234567890123456",
    type: ChannelType.GuildText,
    isTextBased: () => true,
    isSendable: () => true,
    permissionsFor: () => ({ has: () => true })
  }
  const guild = {
    id: "1234567890123456",
    channels: { fetch: async () => channel },
    members: { me: {} }
  }
  const interaction = schedulerInteraction("channel-select", {
    customId: "es:reminderch",
    values: ["3456789012345678"],
    client: { guilds: { fetch: async () => guild } }
  })
  const settings = { event_channel_id: "3456789012345678" }
  const repository = {
    setEventChannel: delayed(interaction, settings),
    getGuildSettings: delayed(interaction, settings)
  }
  await handleEventSchedulerInteraction(interaction, options(repository))
  assert.equal(interaction.acknowledgedWith, "deferUpdate")
  assert.match(interaction.edited.content, /Configure channels/)
})

test("modal submission acknowledges before identity database work", async () => {
  const settings = {
    default_alliance_id: "alliance-1",
    alliance_name: "Main",
    alliance_count: 1
  }
  const interaction = schedulerInteraction("modal", {
    customId: "es:identitym",
    client: {},
    fields: { getTextInputValue: () => "Renamed" }
  })
  const repository = {
    getGuildSettings: delayed(interaction, settings),
    renameAlliance: delayed(interaction, settings),
    getStateLink: delayed(interaction, null),
    getStateDestination: delayed(interaction, null)
  }
  await handleEventSchedulerInteraction(interaction, options(repository))
  assert.equal(interaction.acknowledgedWith, "deferReply")
  assert.match(interaction.edited.content, /Event scheduler/)
})

test("roundup settings and state-link views acknowledge before database reads", async () => {
  for (const [customId, method, expected] of [
    ["es:roundset", "getGuildSettings", /Weekly roundup settings/],
    ["es:stateshare", "getStateLink", /State sharing/]
  ]) {
    const interaction = schedulerInteraction("button", { customId, client: {} })
    const repository = {
      [method]: delayed(interaction, null)
    }
    await handleEventSchedulerInteraction(interaction, options(repository))
    assert.equal(interaction.acknowledgedWith, "deferUpdate")
    assert.match(interaction.edited.content, expected)
  }
})

test("modal-opening controls respond immediately without authorization or database work", async () => {
  let authorizationChecked = false
  let repositoryCreated = false
  const interaction = schedulerInteraction("button", { customId: "es:identity" })
  await handleEventSchedulerInteraction(interaction, {
    healthProvider: () => ({ available: true, gameProfile: "wos" }),
    async userCanManageServer() { authorizationChecked = true; return true },
    repositoryProvider() { repositoryCreated = true; throw new Error("must not load repository") }
  })
  assert.ok(interaction.modal)
  assert.equal(interaction.deferred, false)
  assert.equal(authorizationChecked, false)
  assert.equal(repositoryCreated, false)
})

test("safe responses respect deferred and replied interaction state", async () => {
  const deferred = schedulerInteraction("command", { deferred: true })
  await safelyRespondToInteraction(deferred, { content: "edited" })
  assert.equal(deferred.edited.content, "edited")
  assert.equal(deferred.replyCount || 0, 0)

  const replied = schedulerInteraction("command", { replied: true })
  await safelyRespondToInteraction(replied, { content: "follow-up" })
  assert.equal(replied.followUpCount, 1)
  assert.equal(replied.replyCount || 0, 0)
})

test("handled scheduler interactions prevent generic processing", async () => {
  let genericCount = 0
  const handled = await schedulerInteractionWasHandled({}, async () => true, {})
  if (!handled) genericCount += 1
  assert.equal(handled, true)
  assert.equal(genericCount, 0)
})

test("10062 and 40060 never trigger a second response attempt", async () => {
  for (const code of [EXPIRED_INTERACTION, ALREADY_ACKNOWLEDGED]) {
    const warnings = []
    const interaction = schedulerInteraction("command", { commandName: "event-scheduler" })
    await handleInteractionFailure(interaction, { code }, {
      logger: { warn: message => warnings.push(message), error() {} }
    })
    assert.equal(interaction.replyCount || 0, 0)
    assert.equal(interaction.followUpCount || 0, 0)
    assert.equal(warnings.length, 1)
    assert.doesNotMatch(warnings[0], /token|callback|https?:/i)
  }
})

test("an expired initial acknowledgement is logged once without another reply", async () => {
  const warnings = []
  const interaction = schedulerInteraction("command", {
    commandName: "event-scheduler",
    async deferReply() { throw { code: EXPIRED_INTERACTION } }
  })
  const handled = await handleEventSchedulerInteraction(interaction, options({}, {
    warn: message => warnings.push(message),
    error() {}
  }))
  assert.equal(handled, true)
  assert.equal(interaction.replyCount || 0, 0)
  assert.deepEqual(warnings, [
    "[Discord] Interaction expired before acknowledgement: event-scheduler"
  ])
})

test("unexpected interaction errors remain logged and receive one safe response", async () => {
  const errors = []
  const interaction = schedulerInteraction("command")
  await handleInteractionFailure(interaction, new Error("unexpected"), {
    logger: { warn() {}, error: (...values) => errors.push(values) }
  })
  assert.equal(errors.length, 1)
  assert.equal(interaction.replyCount, 1)
  assert.equal(interaction.replyPayload.flags, MessageFlags.Ephemeral)
  assert.match(errors[0][1].message, /unexpected/)
})

test("scheduler interaction sources do not use the deprecated ephemeral option", () => {
  for (const filename of [
    "interactionResponses.js",
    "eventSchedulerInteractions.js",
    "eventSchedulerHelp.js",
    "eventCreationInteractions.js",
    "eventManagementInteractions.js",
    "allianceManagementInteractions.js"
  ]) {
    const source = fs.readFileSync(path.join(__dirname, "..", "src", filename), "utf8")
    assert.doesNotMatch(source, /ephemeral\s*:\s*true/, filename)
  }
})

test("Discord component validation errors are logged without request secrets", async () => {
  const errors = []
  const interaction = schedulerInteraction("button", { customId: "el:0" })
  await handleInteractionFailure(interaction, {
    code: 50035,
    message: "Invalid Form Body https://discord.com/api/webhooks/secret-token",
    rawError: {
      errors: {
        components: {
          _errors: [{ code: "COMPONENT_CUSTOM_ID_DUPLICATED" }]
        }
      }
    },
    requestBody: { token: "secret-token" }
  }, {
    logger: { warn() {}, error: message => errors.push(message) }
  })
  assert.equal(errors.length, 1)
  assert.match(errors[0], /Discord API error 50035/)
  assert.match(errors[0], /button el:0/)
  assert.match(errors[0], /duplicate custom_id/)
  assert.doesNotMatch(errors[0], /secret-token|webhooks|requestBody/)
  assert.equal(interaction.replyCount, 1)
})

test("Discord invalid-form diagnostics retain field details without secrets", async () => {
  const errors = []
  const interaction = schedulerInteraction("button", { customId: "es:statenumber" })
  await handleInteractionFailure(interaction, {
    code: 50035,
    rawError: {
      errors: {
        data: {
          components: {
            0: {
              components: {
                0: {
                  value: {
                    _errors: [{
                      code: "BASE_TYPE_BAD_LENGTH",
                      message: "Must be between 1 and 4000 in length; token=secret-token"
                    }]
                  }
                }
              }
            }
          }
        }
      }
    },
    requestBody: {
      json: { interaction_token: "interaction-secret", value: "private-input" }
    }
  }, {
    logger: { warn() {}, error: message => errors.push(message) }
  })

  assert.equal(errors.length, 1)
  assert.match(errors[0], /invalid form body/)
  assert.match(errors[0], /data\.components\.0\.components\.0\.value/)
  assert.match(errors[0], /BASE_TYPE_BAD_LENGTH/)
  assert.match(errors[0], /Must be between 1 and 4000 in length/)
  assert.doesNotMatch(errors[0], /secret-token|interaction-secret|private-input|requestBody/)
})
