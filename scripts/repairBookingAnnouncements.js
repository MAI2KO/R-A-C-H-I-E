const { Client, GatewayIntentBits } = require("discord.js")

const { runAnnouncementRepair } = require("../src/bookingAnnouncementRepair")
const { createBookingWebsiteClient, bookingWebsiteConfig } = require("../src/bookingWebsiteClient")
const { createBotSetupRepository } = require("../src/botSetupRepository")
const { closePool, getPool } = require("../src/db")

function parseArguments(args) {
  const modes = args.filter(value => value === "--dry-run" || value === "--execute")
  const cutoffIndex = args.indexOf("--sent-before")
  if (modes.length !== 1 || cutoffIndex < 0 || cutoffIndex + 1 >= args.length
      || args.length !== 3) {
    throw new Error("usage: --dry-run|--execute --sent-before <ISO timestamp>")
  }
  const sentBefore = args[cutoffIndex + 1]
  const parsed = new Date(sentBefore)
  if (!Number.isFinite(parsed.getTime()) || parsed.toISOString() !== sentBefore) {
    throw new Error("--sent-before must be a canonical ISO timestamp")
  }
  return { dryRun: modes[0] === "--dry-run", sentBefore }
}

async function run({ env = process.env, args = process.argv.slice(2), write = console.log,
  clientFactory = options => new Client(options) } = {}) {
  const options = parseArguments(args)
  const config = bookingWebsiteConfig(env)
  if (!config.enabled) throw new Error(`booking website integration unavailable: ${config.disabledReason}`)
  if (!env.BOT_TOKEN) throw new Error("BOT_TOKEN is missing")
  const pool = getPool({ env })
  if (!pool) throw new Error("PostgreSQL is unavailable")
  const client = clientFactory({ intents: [GatewayIntentBits.Guilds,
    GatewayIntentBits.GuildMembers, GatewayIntentBits.GuildMessages,
    GatewayIntentBits.MessageContent] })
  try {
    const ready = new Promise((resolve, reject) => {
      client.once("clientReady", resolve)
      client.once("error", reject)
    })
    await client.login(env.BOT_TOKEN)
    if (!client.isReady()) await ready
    const results = await runAnnouncementRepair({ client,
      api: createBookingWebsiteClient({ config }),
      setupRepository: createBotSetupRepository(pool, config.profile),
      dryRun: options.dryRun, sentBefore: options.sentBefore })
    write(JSON.stringify({ mode: options.dryRun ? "dry-run" : "execute",
      profile: config.profile, sentBefore: options.sentBefore, candidateCount: results.length,
      results }, null, 2))
    return results
  } finally {
    client.destroy()
    await closePool()
  }
}

if (require.main === module) {
  require("dotenv").config()
  run().catch(() => {
    process.stderr.write("Booking announcement repair failed safely.\n")
    process.exitCode = 1
  })
}

module.exports = { parseArguments, run }
