const { bookingWebsiteConfig, createBookingWebsiteClient } = require("../src/bookingWebsiteClient")
const { closePool, getPool } = require("../src/db")
const { createPlayerRepository } = require("../src/giftCodes/playerRepository")
const { runPlayerMirrorReconciliation } = require("../src/giftCodes/playerMirrorReconciliation")

function parseArguments(args) {
  if (args.length !== 1 || !["--dry-run", "--execute"].includes(args[0])) {
    throw new Error("usage: --dry-run|--execute")
  }
  return { dryRun: args[0] === "--dry-run" }
}

async function run({ env = process.env, args = process.argv.slice(2), write = console.log,
  createApi = createBookingWebsiteClient, poolFactory = getPool } = {}) {
  const options = parseArguments(args)
  const config = bookingWebsiteConfig(env)
  if (!config.enabled) {
    throw new Error(`booking website integration unavailable: ${config.disabledReason}`)
  }
  const pool = poolFactory({ env })
  if (!pool) throw new Error("PostgreSQL is unavailable")
  try {
    const result = await runPlayerMirrorReconciliation({
      profile: config.profile,
      repository: createPlayerRepository(pool, config.profile),
      api: createApi({ config }),
      dryRun: options.dryRun
    })
    write(JSON.stringify(result, null, 2))
    return result
  } finally {
    await closePool()
  }
}

if (require.main === module) {
  require("dotenv").config()
  run().then(result => {
    if (result.summary.skipped > 0) process.exitCode = 2
  }).catch(error => {
    process.stderr.write(`Player mirror reconciliation failed safely: ${error.message}\n`)
    process.exitCode = 1
  })
}

module.exports = { parseArguments, run }
