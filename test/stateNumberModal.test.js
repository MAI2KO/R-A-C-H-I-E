const test = require("node:test")
const assert = require("node:assert/strict")
const { ComponentType, MessageFlags, TextInputStyle } = require("discord.js")

const {
  IDS,
  buildStateDestinationNumberModal,
  handleEventSchedulerInteraction
} = require("../src/eventSchedulerInteractions")

let guildSequence = 0

function interaction(type, overrides = {}) {
  guildSequence += 1
  return {
    commandName: null,
    customId: null,
    guildId: `state-number-guild-${guildSequence}`,
    user: { id: "state-number-admin" },
    deferred: false,
    replied: false,
    isChatInputCommand: () => false,
    isButton: () => type === "button",
    isStringSelectMenu: () => false,
    isChannelSelectMenu: () => false,
    isModalSubmit: () => type === "modal",
    async deferReply(options) {
      this.deferred = true
      this.deferOptions = options
    },
    async deferUpdate() {
      this.deferred = true
      this.deferUpdateCount = (this.deferUpdateCount || 0) + 1
    },
    async editReply(payload) { this.edited = payload },
    async reply(payload) { this.replied = true; this.replyPayload = payload },
    async showModal(modal) {
      assert.equal(this.deferred, false, "modal was deferred before showModal")
      assert.equal(this.replied, false, "modal was replied to before showModal")
      this.modal = modal
    },
    ...overrides
  }
}

function handlerOptions(repository, overrides = {}) {
  return {
    healthProvider: () => ({
      available: true,
      gameProfile: "wos",
      botInstanceName: "rachie-wos"
    }),
    async userCanManageServer() { return true },
    repositoryProvider: () => repository,
    logger: { error() {}, warn() {} },
    ...overrides
  }
}

function destination(stateNumber) {
  return {
    state_guild_id: "unused-by-view",
    state_roundup_channel_id: "2345678901234567",
    state_number: stateNumber,
    enabled: true
  }
}

function serializedInput(modal) {
  const payload = modal.toJSON()
  assert.equal(payload.components.length, 1)
  assert.equal(payload.components[0].type, ComponentType.ActionRow)
  assert.equal(payload.components[0].components.length, 1)
  return { payload, input: payload.components[0].components[0] }
}

test("state number modal serializes one valid text input without an empty value", () => {
  const { payload, input } = serializedInput(buildStateDestinationNumberModal(null))

  assert.equal(payload.custom_id, IDS.stateDestinationNumberModal)
  assert.equal(payload.title, "Set state number")
  assert.ok(payload.custom_id.length <= 100)
  assert.ok(payload.title.length <= 45)
  assert.equal(input.type, ComponentType.TextInput)
  assert.equal(input.custom_id, "n")
  assert.ok(input.custom_id.length <= 100)
  assert.equal(new Set(payload.components[0].components.map(component => component.custom_id)).size, 1)
  assert.equal(input.label, "State number")
  assert.equal(input.style, TextInputStyle.Short)
  assert.equal(input.required, true)
  assert.equal(input.min_length, 1)
  assert.equal(input.max_length, 10)
  assert.ok(input.min_length <= input.max_length)
  assert.equal(input.placeholder, "689")
  assert.equal(Object.hasOwn(input, "value"), false)
})

test("state number modal uses a populated destination as its initial value", () => {
  const { input } = serializedInput(buildStateDestinationNumberModal(destination("689")))
  assert.equal(input.value, "689")
})

for (const [label, stateNumber] of [["NULL", null], ["populated", "689"]]) {
  test(`state number button opens immediately for an existing ${label} destination`, async () => {
    const existing = destination(stateNumber)
    const viewInteraction = interaction("button", { customId: IDS.stateDestination })
    const repository = { async getStateDestination() { return existing } }
    await handleEventSchedulerInteraction(viewInteraction, handlerOptions(repository))

    let authorizationChecked = false
    let repositoryCreated = false
    const buttonInteraction = interaction("button", {
      customId: IDS.stateDestinationNumber,
      guildId: viewInteraction.guildId
    })
    await handleEventSchedulerInteraction(buttonInteraction, handlerOptions(null, {
      async userCanManageServer() { authorizationChecked = true; return true },
      repositoryProvider() {
        repositoryCreated = true
        throw new Error("repository must not be loaded before showModal")
      }
    }))

    const { input } = serializedInput(buttonInteraction.modal)
    assert.equal(buttonInteraction.deferred, false)
    assert.equal(buttonInteraction.deferUpdateCount || 0, 0)
    assert.equal(authorizationChecked, false)
    assert.equal(repositoryCreated, false)
    assert.equal(Object.hasOwn(input, "value"), stateNumber !== null)
    if (stateNumber !== null) assert.equal(input.value, stateNumber)
  })
}

for (const [label, initialStateNumber] of [["unset", null], ["populated", "123"]]) {
  test(`state number submission trims and updates an existing ${label} destination`, async () => {
    let saved
    const existing = destination(initialStateNumber)
    const updated = destination("689")
    const viewInteraction = interaction("button", { customId: IDS.stateDestination })
    const repository = {
      async getStateDestination() { return existing },
      async setStateDestinationNumber(input) { saved = input; return updated }
    }
    await handleEventSchedulerInteraction(viewInteraction, handlerOptions(repository))

    const buttonInteraction = interaction("button", {
      customId: IDS.stateDestinationNumber,
      guildId: viewInteraction.guildId
    })
    await handleEventSchedulerInteraction(buttonInteraction, handlerOptions(repository))
    const { input } = serializedInput(buttonInteraction.modal)
    assert.equal(Object.hasOwn(input, "value"), initialStateNumber !== null)

    const modalInteraction = interaction("modal", {
      customId: IDS.stateDestinationNumberModal,
      guildId: viewInteraction.guildId,
      fields: { getTextInputValue: customId => customId === "n" ? "  689  " : "" }
    })
    await handleEventSchedulerInteraction(modalInteraction, handlerOptions(repository))

    assert.equal(modalInteraction.deferred, true)
    assert.equal(modalInteraction.deferOptions.flags, MessageFlags.Ephemeral)
    assert.deepEqual(saved, {
      stateGuildId: modalInteraction.guildId,
      stateNumber: "689"
    })
    assert.match(modalInteraction.edited.content, /State number: 689/)
  })
}

test("state number submission rejects non-numeric input without database mutation", async () => {
  let mutationCount = 0
  const modalInteraction = interaction("modal", {
    customId: IDS.stateDestinationNumberModal,
    fields: { getTextInputValue: () => " 68A " }
  })
  const repository = {
    async setStateDestinationNumber() { mutationCount += 1 }
  }

  await handleEventSchedulerInteraction(modalInteraction, handlerOptions(repository))

  assert.equal(modalInteraction.deferred, true)
  assert.equal(modalInteraction.deferOptions.flags, MessageFlags.Ephemeral)
  assert.equal(mutationCount, 0)
  assert.match(modalInteraction.edited.content, /1 to 10 digits/)
})
