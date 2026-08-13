const test = require("node:test")
const assert = require("node:assert/strict")

const { botOwnerIds, isBotOperator } = require("../src/botOperators")

test("BOT_OWNER_IDS is an exact, fail-closed global operator allowlist", () => {
  const operator = "707866087248756736"
  const env = {
    BOT_OWNER_IDS: ` , malformed,123, ${operator}, 000000000000000000,${operator}000 `
  }

  assert.deepEqual([...botOwnerIds(env)], [operator])
  assert.equal(isBotOperator(operator, env), true)
  assert.equal(isBotOperator("707866087248756737", env), false)
  assert.equal(isBotOperator(operator, {}), false)
  assert.equal(isBotOperator(operator, { BOT_OWNER_IDS: "" }), false)
  assert.equal(isBotOperator("malformed", env), false)
})
