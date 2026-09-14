const { bookingWebsiteConfig, createBookingWebsiteClient } = require("../src/bookingWebsiteClient")
const { closePool, getPool } = require("../src/db")
const { auditInvalidPlayerAccounts } = require("../src/giftCodes/playerAccountDiagnostics")
const { createPlayerRepository } = require("../src/giftCodes/playerRepository")

function parseArguments(args) {
  if (args.length === 1 && args[0] === "--dry-run") return { releaseRef: null, operator: null }
  const release = args.find(value => value.startsWith("--release="))?.slice(10)
  const operator = args.find(value => value.startsWith("--operator="))?.slice(11)
  if (args.length !== 2 || !release || !operator) {
    throw new Error("usage: --dry-run | --release=<account-ref> --operator=<discord-user-id>")
  }
  return { releaseRef: release, operator }
}

async function run({ env = process.env, args = process.argv.slice(2), write = console.log,
  createApi = createBookingWebsiteClient, poolFactory = getPool } = {}) {
  const options = parseArguments(args)
  const config = bookingWebsiteConfig(env)
  if (!config.enabled) throw new Error("booking website integration unavailable")
  const pool = poolFactory({ env })
  if (!pool) throw new Error("PostgreSQL is unavailable")
  try {
    const result = await auditInvalidPlayerAccounts({ profile: config.profile,
      repository: createPlayerRepository(pool, config.profile), api: createApi({ config }),
      releaseRef: options.releaseRef, operatorDiscordUserId: options.operator })
    write(JSON.stringify(result, null, 2))
    return result
  } finally {
    await closePool()
  }
}

if (require.main === module) {
  require("dotenv").config()
  run().catch(() => {
    process.stderr.write("Invalid player-account audit failed safely.\n")
    process.exitCode = 1
  })
}

module.exports = { parseArguments, run }
