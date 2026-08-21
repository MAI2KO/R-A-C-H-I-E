const assert = require("node:assert/strict")
const fs = require("node:fs")
const path = require("node:path")
const test = require("node:test")

const { createBookingWebsiteBootstrap } = require("../src/bookingWebsiteBootstrap")

const secret = "local-placeholder-secret-value-1234567890"

function harness(env) {
  const clientHandlers = new Map()
  const processHandlers = new Map()
  const logs = []
  const errors = []
  let runtimeStarts = 0
  let runtimeStops = 0
  const api = { profile: env.GAME_PROFILE || "wos" }
  const client = {
    isReady: () => false,
    once(event, handler) { clientHandlers.set(event, handler) }
  }
  const result = createBookingWebsiteBootstrap({
    client,
    env,
    processRef: { once(event, handler) { processHandlers.set(event, handler) } },
    logger: { log(value) { logs.push(JSON.parse(value)) }, error(value) { errors.push(JSON.parse(value)) } },
    createApi({ config }) { assert.equal(config.profile, api.profile); return api },
    createRuntime({ api: suppliedApi, intervalMs }) {
      assert.equal(suppliedApi, api)
      assert.equal(intervalMs, 10000)
      return { start() { runtimeStarts++; return { started: true } }, async stop() { runtimeStops++ } }
    }
  })
  return { result, clientHandlers, processHandlers, logs, errors,
    runtimeStarts: () => runtimeStarts, runtimeStops: () => runtimeStops }
}

for (const profile of ["wos", "kingshot"]) {
  test(`production bootstrap starts the ${profile} worker on Discord clientReady`, async () => {
    const state = harness({ GAME_PROFILE: profile, BOOKING_WEBSITE_INTEGRATION_ENABLED: "true",
      BOOKING_WEBSITE_BASE_URL: `https://${profile}.example.test`, BOOKING_WEBSITE_INTEGRATION_SECRET: secret })
    assert.equal(state.result.api.profile, profile)
    assert.equal(state.runtimeStarts(), 0)
    assert.equal(typeof state.clientHandlers.get("clientReady"), "function")
    state.clientHandlers.get("clientReady")()
    assert.equal(state.runtimeStarts(), 1)
    assert.equal(state.logs.find(item => item.event === "booking_website_integration_worker_started").worker_started, true)
    state.processHandlers.get("SIGTERM")()
    await new Promise(resolve => setImmediate(resolve))
    assert.equal(state.runtimeStops(), 1)
  })
}

test("disabled and incomplete configuration never creates or starts a worker", () => {
  for (const env of [
    { GAME_PROFILE: "wos", BOOKING_WEBSITE_INTEGRATION_ENABLED: "false" },
    { GAME_PROFILE: "wos", BOOKING_WEBSITE_INTEGRATION_ENABLED: "true", BOOKING_WEBSITE_INTEGRATION_SECRET: secret },
    { GAME_PROFILE: "wos", BOOKING_WEBSITE_INTEGRATION_ENABLED: "true", BOOKING_WEBSITE_BASE_URL: "https://wos.example" }
  ]) {
    let factoryCalls = 0
    const logs = []
    const result = createBookingWebsiteBootstrap({ client: { isReady: () => false, once() {} }, env,
      processRef: { once() {} }, logger: { log(value) { logs.push(JSON.parse(value)) }, error() {} },
      createApi() { factoryCalls++; }, createRuntime() { factoryCalls++ } })
    assert.equal(result.api, null)
    assert.equal(factoryCalls, 0)
    assert.ok(logs.some(item => item.event === "booking_website_integration_disabled"))
  }
})

test("npm start production entrypoint invokes the tested booking bootstrap from main", () => {
  const packageJson = require("../package.json")
  const source = fs.readFileSync(path.join(__dirname, "../index.js"), "utf8")
  assert.equal(packageJson.scripts.start, "node index.js")
  assert.match(source, /async function main\(\)/)
  assert.match(source, /createBookingWebsiteBootstrap\(\{[\s\S]*configuration: bookingWebsiteConfiguration/)
  assert.match(source, /main\(\)\s*$/)
})
