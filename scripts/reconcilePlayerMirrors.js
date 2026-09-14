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

async function runCli({ operation = run, runtime = process } = {}) {
  try {
    const result = await operation()
    runtime.exitCode = result.summary.skipped > 0 ? 2 : 0
    return result
  } catch {
    runtime.stderr.write("Player mirror reconciliation failed safely.\n")
    runtime.exitCode = 1
    return null
  }
}

if (require.main === module) {
  require("dotenv").config()
  runCli()
}

module.exports = { parseArguments, run, runCli }
