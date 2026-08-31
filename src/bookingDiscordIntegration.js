const { ActionRowBuilder, ButtonBuilder, ButtonStyle, PermissionFlagsBits } = require("discord.js")
const { createHash } = require("node:crypto")
const { discordTimestamp, utcAppointmentInstant, validInstant } = require("./discordTimeFormatting")

const BUTTON_PREFIX = "booking-approval:v1:"
const MEMBER_SIGNUP_PREFIX = "booking-member:v1:"
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
function appointmentTimeLines(date, time, appointmentAt) {
  const instant = validInstant(appointmentAt) || utcAppointmentInstant(date, time)
  const utcDate = instant ? instant.toISOString().slice(0, 10) : date
  const utcTime = instant ? instant.toISOString().slice(11, 16) : time
  const local = discordTimestamp(instant, "F")
  return [`UTC: ${dateText(utcDate)} at ${utcTime} UTC`, `Your time: ${local || "unavailable"}`]
}
function approvalComponents(requestId) {
  return [new ActionRowBuilder().addComponents(
    new ButtonBuilder().setCustomId(`${BUTTON_PREFIX}${requestId}:approve`).setLabel("Approve").setStyle(ButtonStyle.Success),
    new ButtonBuilder().setCustomId(`${BUTTON_PREFIX}${requestId}:deny`).setLabel("Deny").setStyle(ButtonStyle.Danger)
  )]
}
const pendingManagerFooter = "This request is holding the slot temporarily."
function withoutManagerActionSection(content) {
  const pendingSuffix = `\n\n${pendingManagerFooter}`
  if (content.endsWith(pendingSuffix)) return content.slice(0, -pendingSuffix.length)
  const finalMarkers = ["\n\nAPPROVED\n", "\n\nDENIED\n", "\n\nEXPIRED\n"]
  const markerIndex = Math.max(...finalMarkers.map(marker => content.lastIndexOf(marker)))
  return markerIndex === -1 ? content : content.slice(0, markerIndex)
}
function managerRequestContext(work) {
  if (work.originalContent) return withoutManagerActionSection(work.originalContent)
  const place = `${locationLabel(work.profile)} ${work.communityCode} — ${work.communityName}`
  return [`New ${work.serviceLabel} appointment request`, "", place, "",
    `Player: ${work.playerName}`, `Alliance: ${work.alliance}`, `Player ID: ${work.playerId}`,
    `Time: ${work.time} UTC`, `Date: ${dateText(work.date)}`, requirementsText(work.requirements)]
    .filter(line => line !== "").join("\n")
}
function renderManagerRequest(work) {
  const content = managerRequestContext(work)
  if (work.type === "manager_request") return {
    content: `${content}\n\n${pendingManagerFooter}`,
    components: approvalComponents(work.requestId)
  }
  const label = work.status === "confirmed" ? "APPROVED" : work.status === "denied" ? "DENIED" : "EXPIRED"
  const detail = work.status === "confirmed" ? `Approved by ${work.decidedByDisplayName}`
    : work.status === "denied" ? `Denied by ${work.decidedByDisplayName}` : "The temporary booking hold expired."
  return { content: `${content}\n\n${label}\n${detail}`, components: [] }
}
function renderWork(work) {
  const place = `${locationLabel(work.profile)} ${work.communityCode} — ${work.communityName}`
  if (work.type === "manager_request" || work.type === "manager_update") return renderManagerRequest(work)
  const heading = { player_confirmed: "Appointment confirmed", player_approved: "Appointment approved",
    player_rescheduled: "Appointment rescheduled", player_cancelled: "Appointment cancelled",
    appointment_reminder: "Appointment reminder" }[work.type]
  const lines = [heading, "", work.serviceLabel]
  if (work.type === "player_rescheduled") lines.push(
    place, "", "Previous:",
    ...appointmentTimeLines(work.previousDate, work.previousTime, work.previousAppointmentAt),
    "", "New:", ...appointmentTimeLines(work.date, work.time, work.appointmentAt)
  )
  else if (work.type === "appointment_reminder") lines.push(
    `Your ${work.serviceLabel} appointment is in 30 minutes.`, "", place, "",
    ...appointmentTimeLines(work.date, work.time, work.appointmentAt), "",
    `Player: ${work.playerName}`, `Alliance: ${work.alliance}`
  )
  else lines.push(place, "", ...appointmentTimeLines(work.date, work.time, work.appointmentAt))
  if (work.type === "player_approved" && work.decidedByDisplayName) lines.push("", `Approved by ${work.decidedByDisplayName}`)
  if (work.type === "player_rescheduled" && work.attributionDisplayName) lines.push("", `Rescheduled by ${work.attributionDisplayName}`)
  if (work.type === "player_cancelled" && work.attributionDisplayName) lines.push("", `Cancelled by ${work.attributionDisplayName}`)
  return { content: lines.join("\n"), components: [] }
}

function bookingWindowOpenMessages(work) {
  return Object.freeze({
    public: {
      content: "Minister sign-up is now open\n\nSign-up closes Sunday at 12:00 UTC.",
      components: [new ActionRowBuilder().addComponents(
        new ButtonBuilder().setLabel("Guest sign up").setStyle(ButtonStyle.Link)
          .setURL(work.guestUrl),
        new ButtonBuilder().setLabel("Member sign up").setStyle(ButtonStyle.Primary)
          .setCustomId(`${MEMBER_SIGNUP_PREFIX}${work.communityCode}`)
      )]
    },
    manager: {
      content: [
        `Booking links — ${locationLabel(work.profile)} ${work.communityCode}`,
        "", "Member sign-up:", work.memberUrl,
        "", "Guest sign-up:", work.guestUrl,
        "", "Closes:", "Sunday 12:00 UTC"
      ].join("\n"),
      components: []
    }
  })
}

function manualGuestLinkMessage(work) {
  return {
    content: [
      `New guest booking link — ${locationLabel(work.profile)} ${work.communityCode}`,
      "", "A new guest sign-up link has been created.",
      "", work.guestUrl,
      "", "Guest bookings require manager approval."
    ].join("\n"),
    components: []
  }
}

async function addQualifiedManagers(guild, managerRoleId, managers, valueFor = member => member) {
  const members = await guild.members.fetch()
  for (const member of members.values()) {
    if (member.user?.bot) continue
    const qualified = member.id === guild.ownerId
      || member.permissions?.has(PermissionFlagsBits.Administrator)
      || (managerRoleId && member.roles?.cache?.has(managerRoleId))
    if (qualified && !managers.has(member.id)) managers.set(member.id, valueFor(member))
  }
}

async function discoverManagers(client, work, setupRepository = null) {
  const recipients = new Map()
  for (const link of work.guilds || []) {
    const guildId = link.guildId
    const setup = setupRepository ? await setupRepository.get(guildId) : null
    const managerRoleId = setup?.bot_manager_role_id
    const guild = await client.guilds.fetch(guildId)
    await addQualifiedManagers(guild, managerRoleId, recipients, () => guildId)
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

async function sendChannelIdempotently(channel, nonceSeed, message) {
  const nonce = createHash("sha256").update(String(nonceSeed), "utf8")
    .digest("base64url").slice(0, 25)
  return channel.send({ ...message, nonce, enforceNonce: true })
}

async function deliverBookingWindowOpen(client, api, work, setupRepository) {
  if (new Date(work.closesAt).getTime() <= Date.now()) {
    return api.outcome(work, { status: "permanent_failure", errorCode: "booking_window_closed" })
  }
  if (!setupRepository) {
    return api.outcome(work, { status: "retry", errorCode: "bot_setup_database_unavailable" })
  }
  const messages = bookingWindowOpenMessages(work)
  let firstPublic = null
  let publicDestinations = 0
  const managers = new Map()
  try {
    for (const guildId of work.guilds || []) {
      const setup = await setupRepository.get(guildId)
      if (!setup?.minister_sign_up_channel_id) continue
      const guild = await client.guilds.fetch(guildId)
      const channel = await guild.channels.fetch(setup.minister_sign_up_channel_id)
      const sent = await sendChannelIdempotently(
        channel, `${work.workId}${guildId}public`, messages.public,
      )
      publicDestinations++
      if (!firstPublic) firstPublic = { discordChannelId: channel.id, discordMessageId: sent.id }
      await addQualifiedManagers(guild, setup.bot_manager_role_id, managers)
    }
    if (publicDestinations === 0) {
      return api.outcome(work, { status: "retry", errorCode: "minister_signup_channel_unavailable" })
    }
    for (const manager of managers.values()) {
      try {
        const dm = await manager.user.createDM()
        await sendChannelIdempotently(dm, `${work.workId}${manager.id}manager`, messages.manager)
      } catch (error) {
        if (!permanentDiscordCodes.has(error?.code)) throw error
      }
    }
    return api.outcome(work, { status: "sent", ...firstPublic })
  } catch (error) {
    return api.outcome(work, {
      status: permanentDiscordCodes.has(error?.code) ? "permanent_failure" : "retry",
      errorCode: String(error?.code || "booking_window_delivery_failed").slice(0, 80)
    })
  }
}

async function deliverManualGuestLink(client, api, work, setupRepository) {
  if (!setupRepository) {
    return api.outcome(work, { status: "retry", errorCode: "bot_setup_database_unavailable" })
  }
  const managers = new Map()
  try {
    let configuredBase
    let suppliedPath = typeof work.guestPath === "string" ? work.guestPath : null
    let legacyGuestUrl = null
    try {
      configuredBase = new URL(api.baseUrl)
      if (!suppliedPath && work.guestUrl) {
        legacyGuestUrl = new URL(work.guestUrl)
        suppliedPath = legacyGuestUrl.pathname
      }
    } catch {}
    const configuredLoopback = configuredBase
      && ["localhost", "127.0.0.1", "::1"].includes(configuredBase.hostname)
    const validConfiguredBase = configuredBase
      && ((configuredBase.protocol === "https:" && (!configuredLoopback || api.allowLoopback === true))
        || (configuredBase.protocol === "http:" && configuredLoopback && api.allowLoopback === true))
      && configuredBase.username === "" && configuredBase.password === ""
      && configuredBase.pathname.replace(/\/$/, "") === ""
      && configuredBase.search === "" && configuredBase.hash === ""
    const validGuestPath = typeof suppliedPath === "string"
      && /^\/book\/[A-Za-z0-9_-]{43}$/.test(suppliedPath)
      && (!legacyGuestUrl || (legacyGuestUrl.search === "" && legacyGuestUrl.hash === ""))
    if (!validConfiguredBase || !validGuestPath) {
      return api.outcome(work, { status: "retry", errorCode: "booking_website_public_url_invalid" })
    }
    const authoritativeWork = {
      ...work,
      guestUrl: `${configuredBase.origin}${suppliedPath}`
    }
    for (const guildId of work.guilds || []) {
      const setup = await setupRepository.get(guildId)
      if (!setup?.minister_sign_up_channel_id) continue
      const guild = await client.guilds.fetch(guildId)
      await addQualifiedManagers(guild, setup.bot_manager_role_id, managers)
    }
    if (managers.size === 0) {
      return api.outcome(work, { status: "retry", errorCode: "booking_managers_unavailable" })
    }
    let first = null
    for (const manager of managers.values()) {
      try {
        const dm = await manager.user.createDM()
        const sent = await sendChannelIdempotently(
          dm, `${work.workId}${manager.id}manual-guest`, manualGuestLinkMessage(authoritativeWork),
        )
        if (!first) first = { discordChannelId: dm.id, discordMessageId: sent.id }
      } catch (error) {
        if (!permanentDiscordCodes.has(error?.code)) throw error
      }
    }
    return api.outcome(work, { status: "sent", ...first })
  } catch (error) {
    return api.outcome(work, {
      status: permanentDiscordCodes.has(error?.code) ? "permanent_failure" : "retry",
      errorCode: String(error?.code || "manual_guest_link_delivery_failed").slice(0, 80)
    })
  }
}

async function deliverWork(client, api, work, { setupRepository = null } = {}) {
  if (work.type === "booking_window_open") {
    return deliverBookingWindowOpen(client, api, work, setupRepository)
  }
  if (work.type === "manager_guest_link") {
    return deliverManualGuestLink(client, api, work, setupRepository)
  }
  if (work.type === "manager_discovery") {
    if (!setupRepository) {
      return api.outcome(work, { status: "retry", errorCode: "bot_setup_database_unavailable" })
    }
    try {
      return api.recipients(work, await discoverManagers(client, work, setupRepository))
    } catch (error) {
      return api.outcome(work, { status: "retry",
        errorCode: String(error?.code || "manager_discovery_failed").slice(0, 80) })
    }
  }
  if (work.type === "manager_request" && new Date(work.holdExpiresAt).getTime() <= Date.now()) {
    return api.outcome(work, { status: "permanent_failure", errorCode: "hold_expired_before_delivery" })
  }
  if (work.type === "manager_update") {
    try {
      const channel = await client.channels.fetch(work.discordChannelId)
      const message = await channel.messages.fetch(work.discordMessageId)
      await message.edit(renderWork({ ...work, originalContent: message.content }))
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

async function handleBookingMemberSignupInteraction(interaction, {
  baseUrl, pool, profile
} = {}) {
  if (!interaction.isButton?.()
      || !String(interaction.customId || "").startsWith(MEMBER_SIGNUP_PREFIX)) return false
  const communityCode = String(interaction.customId).slice(MEMBER_SIGNUP_PREFIX.length)
  await interaction.deferReply({ flags: 64 })
  if (!/^\d{1,10}$/.test(communityCode) || !baseUrl || !pool) {
    await interaction.editReply("Member sign-up is temporarily unavailable. Please try again shortly.")
    return true
  }
  const registered = (await pool.query(
    `SELECT 1 FROM player_accounts
      WHERE game_profile=$1 AND discord_user_id=$2
        AND state_or_kingdom_number=$3 AND is_active=true
        AND in_game_name IS NOT NULL AND alliance_abbreviation IS NOT NULL
      LIMIT 1`,
    [profile, interaction.user.id, communityCode]
  )).rowCount === 1
  if (!registered) {
    await interaction.editReply(
      `Use \`/register\` to register your player for ${locationLabel(profile)} ${communityCode}, then press **Member sign up** again.`
    )
    return true
  }
  await interaction.editReply({
    content: `Your registered player is ready for ${locationLabel(profile)} ${communityCode}.`,
    components: [new ActionRowBuilder().addComponents(
      new ButtonBuilder().setLabel("Open native booking").setStyle(ButtonStyle.Link)
        .setURL(`${String(baseUrl).replace(/\/$/, "")}/booking`)
    )]
  })
  return true
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

function createBookingWebsiteRuntime({ client, api, setupRepository = null,
  intervalMs = 10000, logger = console,
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
          try { await deliverWork(client, api, work, { setupRepository }) }
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

module.exports = { BUTTON_PREFIX, MEMBER_SIGNUP_PREFIX, renderWork, bookingWindowOpenMessages,
  manualGuestLinkMessage,
  discoverManagers, deliverWork, parseApprovalButton, handleBookingApprovalInteraction,
  handleBookingMemberSignupInteraction, createBookingWebsiteRuntime }
