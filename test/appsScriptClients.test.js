const test = require("node:test")
const assert = require("node:assert/strict")
const fs = require("node:fs")
const path = require("node:path")

const {
  createAppsScriptTransport
} = require("../src/appsScript/transport")
const {
  createBookingAppsScriptClient
} = require("../src/appsScript/bookingClient")
const {
  createStateAppsScriptClient
} = require("../src/appsScript/stateClient")
const {
  createConfigAppsScriptClient
} = require("../src/appsScript/configClient")
const {
  createBanterAppsScriptClient
} = require("../src/appsScript/banterClient")

const CLIENT_FACTORIES = [
  createBookingAppsScriptClient,
  createStateAppsScriptClient,
  createConfigAppsScriptClient,
  createBanterAppsScriptClient
]

const CLIENT_URL_ENV_NAMES = [
  "BOOKING_APPS_SCRIPT_URL",
  "STATE_APPS_SCRIPT_URL",
  "CONFIG_APPS_SCRIPT_URL",
  "BANTER_APPS_SCRIPT_URL"
]

function recordingTransport() {
  const calls = []
  return {
    calls,
    transport: {
      async post(url, payload) {
        calls.push({ url, payload })
        return { ok: true }
      }
    }
  }
}

test("all Apps Script clients fall back to APPS_SCRIPT_URL", async () => {
  const recorded = recordingTransport()
  const env = { APPS_SCRIPT_URL: "https://legacy.example/exec" }

  for (const createClient of CLIENT_FACTORIES) {
    await createClient({ env, transport: recorded.transport }).post({ action: "test_action" })
  }

  assert.deepEqual(
    recorded.calls.map(call => call.url),
    Array(CLIENT_FACTORIES.length).fill(env.APPS_SCRIPT_URL)
  )
})

test("a service URL override affects only its own client", async () => {
  for (let overriddenIndex = 0; overriddenIndex < CLIENT_FACTORIES.length; overriddenIndex++) {
    const recorded = recordingTransport()
    const overrideUrl = `https://override-${overriddenIndex}.example/exec`
    const env = {
      APPS_SCRIPT_URL: "https://legacy.example/exec",
      [CLIENT_URL_ENV_NAMES[overriddenIndex]]: overrideUrl
    }
    const clients = CLIENT_FACTORIES.map(createClient => createClient({
      env,
      transport: recorded.transport
    }))

    for (const client of clients) await client.post({ action: "test_action" })

    assert.deepEqual(
      recorded.calls.map(call => call.url),
      CLIENT_FACTORIES.map((_, index) => (
        index === overriddenIndex ? overrideUrl : env.APPS_SCRIPT_URL
      ))
    )
  }
})

test("wos and kingshot use the same Apps Script client code path", async () => {
  const recorded = recordingTransport()
  const payload = {
    action: "book_for_server",
    adminKey: "test-admin-key",
    discordServerId: "123",
    discordUserId: "456",
    day: "Research",
    time: "12:00",
    fc: null,
    rfc: null,
    shards: "100",
    speedups: null
  }

  for (const gameProfile of ["wos", "kingshot"]) {
    const client = createBookingAppsScriptClient({
      env: { GAME_PROFILE: gameProfile, APPS_SCRIPT_URL: "https://legacy.example/exec" },
      transport: recorded.transport
    })
    await client.post(payload)
  }

  assert.deepEqual(recorded.calls, [
    { url: "https://legacy.example/exec", payload },
    { url: "https://legacy.example/exec", payload }
  ])
})

test("clients pass legacy action payload objects through unchanged", async () => {
  const recorded = recordingTransport()
  const env = { APPS_SCRIPT_URL: "https://legacy.example/exec" }
  const cases = [
    [createBookingAppsScriptClient, {
      action: "book_for_server",
      adminKey: "test-admin-key",
      discordServerId: "123",
      discordUserId: "456",
      day: "Construction",
      time: "09:30",
      fc: null,
      rfc: "200",
      shards: null,
      speedups: "3"
    }],
    [createStateAppsScriptClient, {
      action: "link_state",
      adminKey: "test-admin-key",
      stateCode: "1234",
      joinPassword: "test-password",
      discordServerId: "123",
      discordServerName: "Test Server",
      createdBy: "tester"
    }],
    [createConfigAppsScriptClient, {
      action: "update_setting_for_server",
      adminKey: "test-admin-key",
      discordServerId: "123",
      key: "booking_open",
      value: "false"
    }],
    [createBanterAppsScriptClient, {
      action: "set_banter_spice_for_server",
      adminKey: "test-admin-key",
      discordServerId: "123",
      spiceLevel: "spicy"
    }]
  ]

  for (const [createClient, payload] of cases) {
    await createClient({ env, transport: recorded.transport }).post(payload)
    assert.strictEqual(recorded.calls.at(-1).payload, payload)
  }
})

test("shared transport preserves the existing Axios request contract", async () => {
  const calls = []
  const httpClient = {
    async post(...args) {
      calls.push(args)
      return { data: { ok: true, value: "response-data" } }
    }
  }
  const transport = createAppsScriptTransport({ httpClient })
  const payload = { action: "get_times_for_server", adminKey: "test-admin-key" }

  const result = await transport.post("https://legacy.example/exec", payload)

  assert.deepEqual(result, { ok: true, value: "response-data" })
  assert.deepEqual(calls, [[
    "https://legacy.example/exec",
    payload,
    { headers: { "Content-Type": "application/json" } }
  ]])
})

test("index assigns every Apps Script action to its responsibility client", () => {
  const source = fs.readFileSync(path.join(__dirname, "..", "index.js"), "utf8")

  function actionsFor(clientName) {
    const pattern = new RegExp(`${clientName}\\.post\\(\\{\\s*action: "([^"]+)"`, "g")
    return [...source.matchAll(pattern)].map(match => match[1]).sort()
  }

  assert.deepEqual(actionsFor("bookingAppsScriptClient"), [
    "admin_add_booking_for_server",
    "admin_remove_booking_for_server",
    "admin_remove_reserved_slots_for_server",
    "admin_reserve_slots_for_server",
    "book_for_server",
    "clear_bookings_for_server",
    "close_bookings_for_server",
    "delete_registered_player_for_server",
    "get_booking_date_for_server",
    "get_booking_link_for_server",
    "get_my_bookings_for_server",
    "get_registered_player_for_server",
    "get_reserved_times_for_server",
    "get_times_for_server",
    "get_times_for_server",
    "get_times_for_server",
    "get_times_for_server",
    "open_bookings_for_server",
    "register_player_for_server",
    "remove_booking_for_server",
    "set_booking_date_for_server"
  ])
  assert.deepEqual(actionsFor("stateAppsScriptClient"), [
    "get_linked_servers_for_current_state",
    "get_linked_servers_for_current_state",
    "get_linked_servers_for_current_state",
    "get_sheet_link_for_server",
    "grant_sheet_access_for_server",
    "link_state",
    "reset_state_password",
    "set_announcement_channel",
    "setup_state",
    "unlink_state_server_by_id"
  ])
  assert.deepEqual(actionsFor("configAppsScriptClient"), [
    "get_booking_config_for_server",
    "get_settings_for_server",
    "update_setting_for_server",
    "update_setting_for_server",
    "update_setting_for_server"
  ])
  assert.deepEqual(actionsFor("banterAppsScriptClient"), [
    "clear_banter_channel_for_server",
    "set_banter_channel_for_server",
    "set_banter_spice_for_server"
  ])
  assert.match(source, /fetchAction: \(action, guildId\) => banterAppsScriptClient\.post\(\{\s*action,/)
  assert.doesNotMatch(source,
    /action: "(?:get|set|clear)_bot_admin_role_for_server"/)
  assert.doesNotMatch(source, /postToAppsScript/)
})
