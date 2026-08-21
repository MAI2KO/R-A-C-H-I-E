const { ActionRowBuilder, ButtonBuilder, ButtonStyle, PermissionFlagsBits } = require("discord.js")

const BUTTON_PREFIX = "booking-approval:v1:"
const permanentDiscordCodes = new Set([50001, 50007, 10013])

function locationLabel(profile) { return profile === "kingshot" ? "Kingdom" : "State" }
function dateText(value) {
  const date = new Date(`${String(value).slice(0, 10)}T12:00:00Z`)
  return Number.isNaN(date.getTime()) ? String(value) : new Intl.DateTimeFormat("en-GB", { day: "numeric", month: "long", timeZone: "UTC" }).format(date)
}
function requirementsText(requirements) {
  return (Array.isArray(requirements) ? requirements : []).map(item =>
    `${item.label}: ${item.value}${item.unit === "days" ? " days" : ""}`).join("\n")
}
function approvalComponents(requestId) {
  return [new ActionRowBuilder().addComponents(
    new ButtonBuilder().setCustomId(`${BUTTON_PREFIX}${requestId}:approve`).setLabel("Approve").setStyle(ButtonStyle.Success),
    new ButtonBuilder().setCustomId(`${BUTTON_PREFIX}${requestId}:deny`).setLabel("Deny").setStyle(ButtonStyle.Danger)
  )]
}
function renderWork(work) {
  const place = `${locationLabel(work.profile)} ${work.communityCode} — ${work.communityName}`
  if (work.type === "manager_request") return {
    content: [`New ${work.serviceLabel} appointment request`, "", place, "",
      `Player: ${work.playerName}`, `Alliance: ${work.alliance}`, `Player ID: ${work.playerId}`,
      `Time: ${work.time} UTC`, `Date: ${dateText(work.date)}`, requirementsText(work.requirements), "",
      "This request is holding the slot temporarily."].filter(line => line !== "").join("\n"),
    components: approvalComponents(work.requestId)
  }
  if (work.type === "manager_update") {
    const label = work.status === "confirmed" ? "APPROVED" : work.status === "denied" ? "DENIED" : "EXPIRED"
    const detail = work.status === "confirmed" ? `Approved by ${work.decidedByDisplayName}`
      : work.status === "denied" ? `Denied by ${work.decidedByDisplayName}` : "The temporary booking hold expired."
    return { content: `${label}\n${detail}`, components: [] }
  }
  const heading = { player_confirmed: "Appointment confirmed", player_approved: "Appointment approved",
    player_rescheduled: "Appointment rescheduled", player_cancelled: "Appointment cancelled",
    appointment_reminder: "Appointment reminder" }[work.type]
  const lines = [heading, "", work.serviceLabel]
  if (work.type === "player_rescheduled") lines.push(`Previous: ${dateText(work.previousDate)} at ${work.previousTime} UTC`, `New: ${dateText(work.date)} at ${work.time} UTC`)
  else if (work.type === "appointment_reminder") lines.push(`Your ${work.serviceLabel} appointment is in 30 minutes.`, "", place, `Today at ${work.time} UTC`, "", `Player: ${work.playerName}`, `Alliance: ${work.alliance}`)
  else lines.push(place, `${dateText(work.date)} at ${work.time} UTC`)
  if (work.type === "player_approved" && work.decidedByDisplayName) lines.push("", `Approved by ${work.decidedByDisplayName}`)
  if (work.type === "player_rescheduled" && work.attributionDisplayName) lines.push("", `Rescheduled by ${work.attributionDisplayName}`)
  if (work.type === "player_cancelled" && work.attributionDisplayName) lines.push("", `Cancelled by ${work.attributionDisplayName}`)
  return { content: lines.join("\n"), components: [] }
}

async function discoverManagers(client, work) {
  const recipients = new Map()
  for (const link of work.guilds || []) {
    const guild = await client.guilds.fetch(link.guildId)
    const members = await guild.members.fetch()
    for (const member of members.values()) {
      if (member.user?.bot) continue
      const qualified = member.id === guild.ownerId
        || member.permissions?.has(PermissionFlagsBits.Administrator)
        || (link.managerRoleId && member.roles?.cache?.has(link.managerRoleId))
      if (qualified && !recipients.has(member.id)) recipients.set(member.id, link.guildId)
    }
  }
  return [...recipients].map(([discordUserId, sourceGuildId]) => ({ discordUserId, sourceGuildId }))
}

async function sendIdempotently(client, work, message) {
  const user = await client.users.fetch(work.recipientDiscordUserId)
  const channel = await user.createDM()
  const nonce = String(work.workId).replaceAll("-", "").slice(0, 25)
  const sent = await channel.send({ ...message, nonce, enforceNonce: true })
  return { discordChannelId: channel.id, discordMessageId: sent.id }
}

async function deliverWork(client, api, work) {
  if (work.type === "manager_discovery") return api.recipients(work, await discoverManagers(client, work))
  if (work.type === "manager_request" && new Date(work.holdExpiresAt).getTime() <= Date.now()) {
    return api.outcome(work, { status: "permanent_failure", errorCode: "hold_expired_before_delivery" })
  }
  if (work.type === "manager_update") {
    try {
      const channel = await client.channels.fetch(work.discordChannelId)
      const message = await channel.messages.fetch(work.discordMessageId)
      await message.edit(renderWork(work))
      return api.outcome(work, { status: "sent", discordChannelId: channel.id, discordMessageId: message.id })
    } catch (error) {
      return api.outcome(work, { status: permanentDiscordCodes.has(error?.code) ? "permanent_failure" : "retry",
        errorCode: String(error?.code || "discord_edit_failed").slice(0, 80) })
    }
  }
  try {
    const sent = await sendIdempotently(client, work, renderWork(work))
    return api.outcome(work, { status: "sent", ...sent })
  } catch (error) {
    return api.outcome(work, { status: permanentDiscordCodes.has(error?.code) ? "permanent_failure" : "retry",
      errorCode: String(error?.code || "discord_delivery_failed").slice(0, 80) })
  }
}

function parseApprovalButton(customId) {
  const match = /^booking-approval:v1:([0-9a-f-]{36}):(approve|deny)$/i.exec(String(customId || ""))
  return match ? { requestId: match[1], action: match[2] } : null
}
async function handleBookingApprovalInteraction(interaction, api) {
  if (!interaction.isButton?.()) return false
  const parsed = parseApprovalButton(interaction.customId)
  if (!parsed) return false
  await interaction.deferReply({ flags: 64 })
  try {
    const response = await api.approval(parsed.requestId, parsed.action, {
      discordUserId: interaction.user.id,
      displayName: interaction.member?.displayName || interaction.user.globalName || interaction.user.username
    })
    const outcome = response.result?.outcome || "completed"
    const actor = response.result?.decidedByDisplayName
    const text = outcome.startsWith("already_")
      ? `This request is already ${outcome.slice(8).replace("confirmed", "approved")}${actor ? ` by ${actor}` : ""}.`
      : outcome === "expired" ? "This request has expired." : `Request ${outcome === "confirmed" ? "approved" : outcome}.`
    await interaction.editReply(text)
  } catch (error) {
    const code = error?.publicCode
    const text = code === "manager_forbidden" ? "You are no longer authorised to manage this State/Kingdom."
      : code === "manager_verification_unavailable" ? "Manager permissions could not be checked right now. Please try again."
      : "The booking website is temporarily unavailable. Please try again."
    await interaction.editReply(text)
  }
  return true
}

function createBookingWebsiteRuntime({ client, api, intervalMs = 10000, logger = console,
  setIntervalFn = setInterval, clearIntervalFn = clearInterval }) {
  let timer = null
  let active = null
  let connected = false
  let lastFailureCode = null
  async function tick() {
    if (active) return active
    active = (async () => {
      try {
        const response = await api.claim(10)
        if (!connected || lastFailureCode) {
          logger.log(JSON.stringify({
            event: lastFailureCode ? "booking_website_connection_recovered" : "booking_website_connection_established",
            game_profile: api.profile
          }))
        }
        connected = true
        lastFailureCode = null
        const claimedCount = Array.isArray(response.work) ? response.work.length : 0
        if (claimedCount > 0) {
          logger.log(JSON.stringify({
            event: "booking_website_work_claimed",
            game_profile: api.profile,
            work_count: claimedCount
          }))
        }
        for (const work of response.work || []) {
          if (work.profile !== api.profile) continue
          try { await deliverWork(client, api, work) }
          catch (error) {
            logger.error(JSON.stringify({ event: "booking_discord_work_failed", game_profile: api.profile,
              work_type: work.type, error_code: String(error?.code || "integration_failure").slice(0, 80) }))
          }
        }
        return response.work?.length || 0
      } catch (error) {
        const errorCode = String(error?.code || "integration_unavailable").slice(0, 80)
        if (lastFailureCode !== errorCode) {
          logger.error(JSON.stringify({
            event: "booking_website_poll_failed",
            game_profile: api.profile,
            operation: "claim",
            http_status: Number.isInteger(error?.status) ? error.status : null,
            error_code: errorCode
          }))
        }
        lastFailureCode = errorCode
        return 0
      }
    })().finally(() => { active = null })
    return active
  }
  function start() {
    if (timer) return { started: false, reason: "already_started" }
    timer = setIntervalFn(() => { void tick() }, intervalMs)
    timer.unref?.()
    void tick()
    return { started: true }
  }
  async function stop() { if (timer) clearIntervalFn(timer); timer = null; await active }
  return Object.freeze({ start, stop, tick })
}

module.exports = { BUTTON_PREFIX, renderWork, discoverManagers, deliverWork, parseApprovalButton,
  handleBookingApprovalInteraction, createBookingWebsiteRuntime }
