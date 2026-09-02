const { createHash } = require("node:crypto")
const { addQualifiedManagers, authoritativeGuestUrl, bookingWindowOpenMessages } = require("./bookingDiscordIntegration")

function messageText(message) {
  let components = ""
  try { components = JSON.stringify(message.components || []) } catch {}
  return `${message.content || ""}\n${components}`.toLowerCase()
}

function isRepairMessageCandidate(message, candidate, botUserId) {
  if (message.author?.id !== botUserId) return false
  const sentAt = new Date(candidate.sentAt).getTime()
  const createdAt = Number(message.createdTimestamp || new Date(message.createdAt || 0).getTime())
  if (Number.isFinite(sentAt) && Number.isFinite(createdAt)
      && Math.abs(createdAt - sentAt) > 6 * 60 * 60 * 1000) return false
  const text = messageText(message)
  const place = candidate.profile === "kingshot" ? "kingdom" : "state"
  const repairedHeading = `guest booking — ${place} ${candidate.communityCode}`.toLowerCase()
  return text.includes(repairedHeading)
    || text.includes("minister sign-up is now open")
    || text.includes("member sign-up")
    || text.includes("member sign up")
    || text.includes("guest sign-up")
    || text.includes("guest sign up")
    || text.includes("localhost")
    || text.includes("railway.internal")
}

async function findCandidateMessage(channel, candidate, botUserId) {
  if (candidate.discordChannelId === channel.id && candidate.discordMessageId) {
    try {
      return await channel.messages.fetch(candidate.discordMessageId)
    } catch (error) {
      if (error?.code !== 10008) throw error
      return null
    }
  }
  const recent = await channel.messages.fetch({ limit: 100 })
  const matches = [...recent.values()].filter((message) =>
    isRepairMessageCandidate(message, candidate, botUserId))
  if (matches.length > 1) {
    const error = new Error("announcement_message_ambiguous")
    error.code = "announcement_message_ambiguous"
    throw error
  }
  return matches[0] || null
}

async function inspectAnnouncementRepair(client, setupRepository, candidate) {
  const managers = new Map()
  const destinations = []
  const unavailable = []
  for (const guildId of candidate.guilds || []) {
    try {
      const setup = await setupRepository.get(guildId)
      if (!setup?.minister_sign_up_channel_id) continue
      const guild = await client.guilds.fetch(guildId)
      const channel = await guild.channels.fetch(setup.minister_sign_up_channel_id)
      if (!channel?.messages?.fetch || !channel?.send) throw new Error("announcement_channel_unavailable")
      await addQualifiedManagers(guild, setup.bot_manager_role_id, managers)
      const message = await findCandidateMessage(channel, candidate, client.user.id)
      destinations.push({ guildId, channel, message, action: message ? "edit" : "recreate" })
    } catch (error) {
      unavailable.push({ guildId,
        reason: String(error?.code || "announcement_channel_unavailable").slice(0, 80) })
    }
  }
  const skipReason = unavailable.length ? "configured_announcement_channel_unavailable"
    : destinations.length === 0 ? "configured_announcement_channel_unavailable"
      : managers.size === 0 ? "manager_recipients_unavailable" : null
  return { candidate, destinations, managerCount: managers.size, unavailable,
    action: skipReason ? "skip" : destinations.some((item) => item.action === "recreate")
      ? "rotate_link_and_recreate_message" : "rotate_link_and_edit_message",
    skipReason }
}

function safeResult(inspection, status, extra = {}) {
  return {
    profile: inspection.candidate.profile,
    communityCode: inspection.candidate.communityCode,
    communityId: inspection.candidate.communityId,
    windowId: inspection.candidate.windowId,
    guestLinkId: inspection.candidate.guestLinkId,
    guestLinkHint: inspection.candidate.guestLinkHint,
    announcementChannelId: inspection.candidate.discordChannelId,
    announcementMessageId: inspection.candidate.discordMessageId,
    destinations: inspection.destinations.map((item) => ({ guildId: item.guildId,
      channelId: item.channel.id, messageId: item.message?.id || null,
      botAccess: true, action: item.action })),
    managerCount: inspection.managerCount,
    managersResolved: inspection.managerCount > 0,
    action: inspection.action,
    status,
    ...extra
  }
}

async function sendReplacement(channel, candidate, message) {
  const nonce = createHash("sha256")
    .update(`${candidate.notificationId}:${channel.id}:legacy-public-url-repair-v1`, "utf8")
    .digest("base64url").slice(0, 25)
  return channel.send({ ...message, nonce, enforceNonce: true })
}

async function executeInspection(api, inspection) {
  if (inspection.skipReason) {
    return safeResult(inspection, "skipped", { reason: inspection.skipReason,
      unavailable: inspection.unavailable })
  }
  const started = await api.beginAnnouncementRepair(
    inspection.candidate.notificationId, inspection.candidate.sentBefore)
  const repair = started?.repair
  const guestUrl = repair ? authoritativeGuestUrl(api, repair) : null
  const message = guestUrl ? bookingWindowOpenMessages({ ...repair, guestUrl })?.public : null
  if (!message) return safeResult(inspection, "failed", { reason: "public_url_invalid" })

  const delivered = []
  const failures = []
  for (const destination of inspection.destinations) {
    try {
      let result = destination.message
      if (result) {
        const alreadyCurrent = result.content === message.content
          && (!result.components || result.components.length === 0)
        if (!alreadyCurrent) result = await result.edit(message)
      } else {
        result = await sendReplacement(destination.channel, inspection.candidate, message)
      }
      delivered.push({ channelId: destination.channel.id, messageId: result.id })
    } catch (error) {
      failures.push({ guildId: destination.guildId,
        reason: String(error?.code || "discord_repair_failed").slice(0, 80) })
    }
  }
  if (failures.length || delivered.length !== inspection.destinations.length) {
    return safeResult(inspection, "partial_failure", {
      reason: "discord_announcement_repair_incomplete", failures,
      managerNotification: "queued"
    })
  }
  const primary = delivered[0]
  await api.completeAnnouncementRepair(inspection.candidate.notificationId, {
    discordChannelId: primary.channelId, discordMessageId: primary.messageId
  })
  return safeResult(inspection, "completed", { managerNotification: "queued" })
}

async function runAnnouncementRepair({ client, api, setupRepository, dryRun, sentBefore }) {
  const response = await api.announcementRepairCandidates(sentBefore)
  const results = []
  for (const item of response?.candidates || []) {
    const candidate = { ...item, sentBefore }
    try {
      const inspection = await inspectAnnouncementRepair(client, setupRepository, candidate)
      results.push(dryRun
        ? safeResult(inspection, inspection.skipReason ? "skipped" : "planned",
          inspection.skipReason ? { reason: inspection.skipReason,
            unavailable: inspection.unavailable } : {})
        : await executeInspection(api, inspection))
    } catch (error) {
      results.push({ profile: candidate.profile, communityCode: candidate.communityCode,
        communityId: candidate.communityId, windowId: candidate.windowId,
        status: "failed", reason: String(error?.code || "repair_failed").slice(0, 80) })
    }
  }
  return results
}

module.exports = { isRepairMessageCandidate, findCandidateMessage, inspectAnnouncementRepair,
  executeInspection, runAnnouncementRepair }
