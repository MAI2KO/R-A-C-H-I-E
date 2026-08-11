const test = require("node:test")
const assert = require("node:assert/strict")

const { createBanterConfigLookup } = require("../src/banterConfig")

test("guild with no banter channel is a cached normal configuration state", async () => {
  const actions = []
  const lookup = createBanterConfigLookup({
    fetchAction: async action => {
      actions.push(action)
      return action === "get_banter_spice_for_server"
        ? { ok: true, banter_spice_level: "standard" }
        : { ok: true, banter_channel_id: "" }
    }
  })
  assert.deepEqual(await lookup.get("777777777777777777"), {
    banterChannelId: "",
    spiceLevel: "standard"
  })
  assert.deepEqual(await lookup.get("777777777777777777"), {
    banterChannelId: "",
    spiceLevel: "standard"
  })
  assert.equal(actions.length, 2, "negative configuration was fetched repeatedly")
})

test("Apps Script banter failure is contained, redacted and negatively cached", async () => {
  let requests = 0
  const warnings = []
  const lookup = createBanterConfigLookup({
    fetchAction: async () => {
      requests += 1
      throw Object.assign(
        new Error("Request failed at https://script.google.test/x ADMIN_API_KEY=secret-value"),
        { name: "AxiosError", code: "ERR_BAD_REQUEST", response: { status: 404 } }
      )
    },
    logger: { warn(value) { warnings.push(value) } }
  })
  assert.deepEqual(await lookup.get("777777777777777777"), {
    banterChannelId: "",
    spiceLevel: "standard"
  })
  assert.deepEqual(await lookup.get("777777777777777777"), {
    banterChannelId: "",
    spiceLevel: "standard"
  })
  assert.equal(requests, 2)
  assert.equal(warnings.length, 1)
  const diagnostic = JSON.parse(warnings[0])
  assert.equal(diagnostic.event, "banter_config_lookup_failed")
  assert.equal(diagnostic.error_code, "ERR_BAD_REQUEST")
  assert.doesNotMatch(warnings[0], /secret-value|script\.google\.test/)
})

test("banter cache invalidation allows legitimate configuration changes", async () => {
  let channelId = ""
  const lookup = createBanterConfigLookup({
    fetchAction: async action => action === "get_banter_channel_for_server"
      ? { banter_channel_id: channelId }
      : { banter_spice_level: "standard" }
  })
  assert.equal((await lookup.get("777")).banterChannelId, "")
  channelId = "888"
  assert.equal((await lookup.get("777")).banterChannelId, "")
  lookup.invalidate("777")
  assert.equal((await lookup.get("777")).banterChannelId, "888")
})
