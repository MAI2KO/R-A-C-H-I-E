const { MessageFlags } = require("discord.js")
const { getPool } = require("../../db")
const { createGiftCodeRepository } = require("../repository")
const { GiftCodeError, createGiftCodeService } = require("../service")
const { PlayerValidationError } = require("../validation")
const { getPlayerGiftCodesHealth } = require("../runtime")
const { getGiftCodeRuntime } = require("../workflowRuntime")

function formatPlayerGiftStatus(account, terms) {
  const lines = [
    `**${terms.playerLabel} ${account.player_id}**`,
    `${terms.locationLabel}: ${account.state_or_kingdom_number}`,
    `Automatic redemption: ${account.gift_redemption_enabled ? "Enabled" : "Disabled"}`,
    `Account: ${account.is_active ? "Active" : "Inactive"}`,
    `Verification: ${account.verification_status}`,
    `Successful codes: ${account.successful_redemptions}`
  ]
  if (account.last_redemption_status) lines.push(`Last result: ${account.last_redemption_status}`)
  return lines.join("\n")
}

function formatRuntimeStatus(runtime, diagnostics) {
  const observation = runtime.recentRateLimits.at(-1)
  return [
    `Subsystem: ${runtime.started ? "Available" : "Unavailable"}`,
    `Verification worker: ${runtime.verificationEnabled ? (runtime.verificationRunning ? "Running" : "Not running") : "Disabled"}`,
    `Verifier configured: ${runtime.verifierConfigured ? "Yes" : "No"}`,
    `Redemption worker: ${runtime.redemptionEnabled ? (runtime.redemptionRunning ? "Running" : "Not running") : "Disabled"}`,
    `Pending candidates: ${diagnostics.pending_candidates}`,
    `Active codes: ${diagnostics.active_codes}`,
    `Pending redemptions: ${diagnostics.pending_redemptions}`,
    `Retry queue: ${diagnostics.retry_count}`,
    observation
      ? `Recent rate limit: ${observation.remaining ?? "?"}/${observation.limit ?? "?"} remaining`
      : "Recent rate limit: no observations"
  ].join("\n")
}

function formatCodeDiagnostics(code) {
  if (!code) return "No matching gift code was found."
  return [
    `**Gift code ${code.code}**`,
    `Status: ${code.status}`,
    `Verification: ${code.verification_state}`,
    `First seen: <t:${Math.floor(new Date(code.first_seen_at_utc).getTime() / 1000)}:f>`,
    `Source: ${code.source_name || code.source_type || "Discord submission"}`,
    `Classification: ${code.last_err_code ?? "none"} · ${code.last_api_message || "none"}`,
    `Queue counts: pending ${code.pending_count}, success ${code.success_count}, ` +
      `already redeemed ${code.already_redeemed_count}, review/failed ${code.failed_count}`
  ].join("\n")
}

async function handleGiftInteraction(interaction, {
  userCanManageServer,
  healthProvider = getPlayerGiftCodesHealth,
  runtimeProvider = getGiftCodeRuntime,
  poolProvider = () => getPool(),
  repositoryFactory = createGiftCodeRepository,
  serviceFactory = createGiftCodeService,
  logger = console
} = {}) {
  if (!interaction.isChatInputCommand?.() || !["gift", "gift-admin"].includes(interaction.commandName)) {
    return false
  }
  await interaction.deferReply({ flags: MessageFlags.Ephemeral })
  const isAdminCommand = interaction.commandName === "gift-admin"
  if (isAdminCommand && !(await userCanManageServer(interaction))) {
    await interaction.editReply("You do not have permission to use this command.")
    return true
  }
  const health = healthProvider()
  if (!health.available) {
    await interaction.editReply("Gift-code services are temporarily unavailable. Please try again later.")
    return true
  }

  try {
    const repository = repositoryFactory(poolProvider(), health.gameProfile)
    const service = serviceFactory({ repository, gameProfile: health.gameProfile })
    const subcommand = interaction.options.getSubcommand()
    const playerId = interaction.options.getString("player_id")
    const code = interaction.options.getString("code")

    if (!isAdminCommand && subcommand === "submit") {
      const result = await service.submit({ discordUserId: interaction.user.id, code })
      const runtime = runtimeProvider()?.status() || {}
      const availability = runtime.verificationEnabled && runtime.verifierConfigured
        ? "Verification is queued."
        : "Verification is currently unavailable; the candidate remains pending."
      await interaction.editReply(`Gift code ${result.giftCode.code}: ${result.outcome}. ${availability}`)
      return true
    }
    if (!isAdminCommand && ["auto-enable", "auto-disable"].includes(subcommand)) {
      const enabled = subcommand === "auto-enable"
      const account = await service.setAutomaticRedemption({
        discordUserId: interaction.user.id,
        playerId,
        enabled
      })
      await interaction.editReply(
        `Automatic gift-code redemption is ${enabled ? "enabled" : "disabled"} for ` +
        `${service.terms.playerLabel} ${account.player_id} in ` +
        `${service.terms.locationLabel} ${account.state_or_kingdom_number}.`
      )
      return true
    }
    if (!isAdminCommand && subcommand === "status") {
      const accounts = await service.status({ discordUserId: interaction.user.id, playerId })
      await interaction.editReply(accounts.length
        ? accounts.map(account => formatPlayerGiftStatus(account, service.terms)).join("\n\n")
        : "You have no matching registered player account.")
      return true
    }

    if (isAdminCommand && subcommand === "status") {
      const diagnostics = await service.adminStatus()
      const runtime = runtimeProvider()?.status() || {
        started: false,
        verificationEnabled: false,
        redemptionEnabled: false,
        verifierConfigured: false,
        recentRateLimits: []
      }
      await interaction.editReply(formatRuntimeStatus(runtime, diagnostics))
      return true
    }
    if (isAdminCommand && subcommand === "queue") {
      const diagnostics = await service.adminStatus()
      await interaction.editReply([
        `Pending: ${diagnostics.pending_redemptions}`,
        `Retrying: ${diagnostics.retry_count}`,
        `Oldest pending: ${diagnostics.oldest_pending_at_utc ? `<t:${Math.floor(new Date(diagnostics.oldest_pending_at_utc).getTime() / 1000)}:R>` : "none"}`,
        `Next retry: ${diagnostics.next_retry_at_utc ? `<t:${Math.floor(new Date(diagnostics.next_retry_at_utc).getTime() / 1000)}:R>` : "none"}`
      ].join("\n"))
      return true
    }
    if (isAdminCommand && subcommand === "code") {
      await interaction.editReply(formatCodeDiagnostics(await service.adminCode(code)))
      return true
    }
    if (isAdminCommand && subcommand === "verify") {
      const submission = await service.submit({
        discordUserId: interaction.user.id,
        code,
        isAdmin: true
      })
      const result = await runtimeProvider()?.verifyCode(submission.giftCode.code)
      await interaction.editReply(result?.processed
        ? `TEST verification completed. Result: ${result.result.giftCode.status}.`
        : `TEST verification did not run: ${result?.reason || "runtime unavailable"}.`)
      return true
    }
    throw new GiftCodeError("UNKNOWN_GIFT_ACTION", "That gift-code action is not supported.")
  } catch (error) {
    if (error instanceof PlayerValidationError || error instanceof GiftCodeError) {
      await interaction.editReply(error.message)
      return true
    }
    logger.error(`[Gift codes] Interaction failed: ${String(error?.code || error?.name || "error").slice(0, 100)}`)
    await interaction.editReply("Gift-code services are temporarily unavailable. Please try again later.")
    return true
  }
}

module.exports = {
  formatPlayerGiftStatus,
  formatRuntimeStatus,
  formatCodeDiagnostics,
  handleGiftInteraction
}
