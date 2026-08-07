const test = require("node:test")
const assert = require("node:assert/strict")
const { MessageFlags } = require("discord.js")

const {
  HELP_IDS,
  HELP_SECTIONS,
  buildSchedulerHelpView,
  handleEventSchedulerHelpInteraction
} = require("../src/eventSchedulerHelp")
const {
  buildHomeView,
  handleEventSchedulerInteraction
} = require("../src/eventSchedulerInteractions")
const {
  buildPublishingView,
  buildTimingView
} = require("../src/eventCreationInteractions")
const { confirmationView } = require("../src/eventManagementInteractions")

test("scheduler help sections are complete and remain inside Discord limits", () => {
  assert.equal(HELP_SECTIONS.length, 9)
  const combined = HELP_SECTIONS.map(section => section.content).join("\n")
  for (const phrase of [
    "Getting started",
    "Creating an event",
    "Accepted UTC time formats",
    "Reminder behaviour",
    "Managing events",
    "Alliance weekly roundups",
    "State roundup setup",
    "State behaviour",
    "Troubleshooting"
  ]) assert.match(combined, new RegExp(phrase, "i"))

  for (const section of HELP_SECTIONS) {
    const view = buildSchedulerHelpView(section.id)
    assert.ok(view.content.length <= 2000, section.id)
    assert.ok(view.components.length <= 5, section.id)
    const componentJson = view.components.map(row => row.toJSON())
    assert.ok(componentJson.every(row => row.components.length <= 5), section.id)
    assert.ok(componentJson[0].components[0].options.length <= 25, section.id)
    assert.deepEqual(view.allowedMentions, { parse: [], repliedUser: false })
  }

  assert.match(combined, /5, 10, 15, 20 or 30 minutes/i)
  assert.match(combined, /About to start/i)
  assert.match(combined, /no exact-start post/i)
  assert.match(combined, /image is attached only to the alliance advance reminder/i)
  assert.match(combined, /never creates an individual state reminder/i)
  assert.match(combined, /select no advance reminder|cancel future reminders/i)
  assert.match(combined, /Developer Mode/i)
})

test("scheduler help is ephemeral and remains available when database health is down", async () => {
  let reply
  let healthChecked = false
  let authorizationChecked = false
  const interaction = {
    commandName: "event-scheduler-help",
    customId: null,
    isChatInputCommand: () => true,
    async reply(payload) { reply = payload }
  }
  assert.equal(await handleEventSchedulerInteraction(interaction, {
    healthProvider() { healthChecked = true; throw new Error("must not check health") },
    async userCanManageServer() { authorizationChecked = true; return false }
  }), true)
  assert.equal(healthChecked, false)
  assert.equal(authorizationChecked, false)
  assert.equal(reply.flags, MessageFlags.Ephemeral)
  assert.match(reply.content, /Getting started/i)
})

test("scheduler help section controls edit the existing ephemeral response", async () => {
  let deferred = false
  let edited
  const interaction = {
    commandName: null,
    customId: HELP_IDS.section,
    values: ["state-behaviour"],
    isChatInputCommand: () => false,
    isStringSelectMenu: () => true,
    isButton: () => false,
    async deferUpdate() { deferred = true },
    async editReply(payload) { edited = payload }
  }
  assert.equal(await handleEventSchedulerHelpInteraction(interaction), true)
  assert.equal(deferred, true)
  assert.match(edited.content, /combined weekly roundup/i)
  assert.match(edited.content, /never creates an individual state reminder/i)
})

test("main and context-sensitive scheduler views explain consequential controls", () => {
  const home = buildHomeView(null, null)
  assert.match(JSON.stringify(home.components.map(row => row.toJSON())), /eh:home/)

  const draft = {
    recurrenceDays: 7,
    advanceReminderMinutes: 15,
    advanceReminderMessage: null,
    reminderAtStart: true,
    finalReminderMessage: null,
    publishToAlliance: true,
    publishToState: true,
    includeInWeeklyRoundup: true
  }
  assert.match(buildTimingView("session", draft).content, /one alliance reminder/i)
  assert.match(buildTimingView("session", draft).content, /About to start/i)
  assert.match(buildPublishingView("session", draft).content, /never sends an individual state reminder/i)

  const event = { event_name: "Foundry" }
  assert.match(confirmationView("s", "pause", "e", event).content, /preserving the recurrence schedule/i)
  assert.match(confirmationView("s", "delete", "e", event).content, /preserving history/i)
  assert.match(confirmationView("s", "resume", "e", event).content, /original recurrence schedule/i)
})
