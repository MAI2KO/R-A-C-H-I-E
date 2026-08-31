const {
  ActionRowBuilder,
  ButtonBuilder,
  ButtonStyle,
  ChannelType,
  PermissionFlagsBits
} = require("discord.js")
const { profileTerminology } = require("./giftCodes/terminology")

const BOT_SETUP_IDS = Object.freeze({
  registerCharacter: "gcux:public-register",
  registerMinisters: "botsetup:register-ministers"
})

const PROFILE_NAMES = Object.freeze({
  wos: { botName: "R.A.C.H.I.E", categoryName: "R-A-C-H-I-E" },
  kingshot: { botName: "P.E.G.G.I.E", categoryName: "P-E-G-G-I-E" }
})

const CHANNELS = Object.freeze([
  { key: "gift_auto_redeem", name: "gift-code-auto-redeem", card: "gift" },
  { key: "gift_announcements", name: "gift-code-announcements", threads: true },
  { key: "minister_sign_up", name: "minister-sign-up", card: "minister" },
  { key: "event_scheduler", name: "event-scheduler", card: "events" },
  { key: "event_announcements", name: "event-announcements", threads: true }
])

class BotSetupError extends Error {
  constructor(code, message) {
    super(message)
    this.name = "BotSetupError"
    this.code = code
  }
}

function permissionOverwrites(guild, allowThreads, botMember = guild.members.me) {
  const botId = botMember.id
  const memberAllow = [PermissionFlagsBits.ViewChannel]
  const memberDeny = [PermissionFlagsBits.SendMessages]
  if (allowThreads) {
    memberAllow.push(
      PermissionFlagsBits.CreatePublicThreads,
      PermissionFlagsBits.SendMessagesInThreads
    )
  } else {
    memberDeny.push(
      PermissionFlagsBits.CreatePublicThreads,
      PermissionFlagsBits.SendMessagesInThreads
    )
  }
  return [
    { id: guild.roles.everyone.id, allow: memberAllow, deny: memberDeny },
    {
      id: botId,
      allow: [
        PermissionFlagsBits.ViewChannel,
        PermissionFlagsBits.SendMessages,
        PermissionFlagsBits.EmbedLinks,
        PermissionFlagsBits.AttachFiles,
        PermissionFlagsBits.ReadMessageHistory,
        PermissionFlagsBits.SendMessagesInThreads
      ]
    }
  ]
}

function giftCard(terms) {
  return {
    content: [
      "**Gift Code Auto-Redeem**",
      "",
      `Register your ${terms.gameName} character and the bot will automatically attempt eligible gift codes for you.`,
      "",
      "You need:",
      "- Player ID",
      "- In-game name",
      `- ${terms.locationLabel}`,
      "- Alliance abbreviation",
      "",
      "Auto-Redeem is enabled automatically for new registrations and can be disabled at any time.",
      "The bot never requires your game password."
    ].join("\n"),
    components: [new ActionRowBuilder().addComponents(
      new ButtonBuilder()
        .setCustomId(BOT_SETUP_IDS.registerCharacter)
        .setLabel("REGISTER CHARACTER")
        .setStyle(ButtonStyle.Primary)
    )]
  }
}

const PUBLIC_BOOKING_ORIGINS = Object.freeze({
  wos: "https://r-a-c-h-i-e.com",
  kingshot: "https://peggie.r-a-c-h-i-e.com"
})

function bookingWebsiteUrl(gameProfile, configuredBaseUrl) {
  try {
    const parsed = new URL(String(configuredBaseUrl || ""))
    if (parsed.origin !== PUBLIC_BOOKING_ORIGINS[gameProfile]
        || parsed.username || parsed.password
        || parsed.pathname.replace(/\/$/, "") !== ""
        || parsed.search || parsed.hash) return null
    return `${parsed.origin}/booking`
  } catch {
    return null
  }
}

function ministerCard(terms, configuredBaseUrl = null) {
  const websiteUrl = bookingWebsiteUrl(terms.gameProfile, configuredBaseUrl)
  const buttons = [
    new ButtonBuilder()
      .setCustomId(BOT_SETUP_IDS.registerCharacter)
      .setLabel("Register")
      .setStyle(ButtonStyle.Primary)
  ]
  if (websiteUrl) {
    buttons.push(new ButtonBuilder()
      .setLabel("Open Booking Website")
      .setStyle(ButtonStyle.Link)
      .setURL(websiteUrl))
  }
  return {
    content: [
      "**Minister Sign-Up**",
      "",
      `Register your ${terms.gameName} player account for this ${terms.locationLabel} in Discord, then book appointments on the website.`,
      "",
      "Registered players can:",
      "• book appointments",
      "• manage their own bookings",
      "• use automatic gift-code redemption",
      "• earn account points"
    ].join("\n"),
    components: [new ActionRowBuilder().addComponents(buttons)]
  }
}

function eventCard(terms, guildKind = "alliance") {
  const capabilities = guildKind === "state"
    ? [
        `- ${terms.locationLabel}-wide events, phases, recurrence, and reminders`,
        `- linked alliance event aggregation and ${terms.locationLabel}-wide roundups`,
        "- existing State/Kingdom event configuration and delivery destinations"
      ]
    : [
        "- alliance events, groups, recurrence, and reminders",
        `- ${terms.locationLabel}-wide events and linked ${terms.locationLabel} servers`,
        "- weekly alliance and community roundups",
        "- existing event configuration and delivery destinations"
      ]
  return {
    content: [
      "**Event Scheduler**",
      "",
      "Authorised event managers can use `/event-scheduler` to create and manage:",
      ...capabilities
    ].join("\n"),
    components: []
  }
}

function validateCommunitySetup(input, profile) {
  const terms = profileTerminology(profile)
  const communityNumber = String(input?.communityNumber || "").trim()
  const guildKind = String(input?.guildKind || "")
  const allianceAbbreviation = guildKind === "alliance"
    ? String(input?.allianceAbbreviation || "").trim().toUpperCase() : null
  if (!/^\d{1,10}$/.test(communityNumber)) {
    throw new BotSetupError("INVALID_COMMUNITY", `${terms.locationLabel} number must contain 1 to 10 digits.`)
  }
  if (!["state", "alliance"].includes(guildKind)) {
    throw new BotSetupError("INVALID_KIND", "Choose State/Kingdom Discord or Alliance Discord.")
  }
  if (guildKind === "alliance" && !/^[A-Z0-9]{3}$/.test(allianceAbbreviation)) {
    throw new BotSetupError("INVALID_ALLIANCE", "Alliance abbreviation must contain exactly three letters or digits.")
  }
  return Object.freeze({ communityNumber, guildKind, allianceAbbreviation })
}

function setupStatus(profile, result) {
  const names = PROFILE_NAMES[profile]
  const terms = profileTerminology(profile)
  return [
    `**${names.botName} Setup complete**`,
    "",
    `${terms.locationLabel}: ${result.community.communityNumber}`,
    `Discord server: ${result.community.guildName}`,
    `Discord type: ${result.community.guildKind === "state" ? `${terms.locationLabel} Discord` : "Alliance Discord"}`,
    ...(result.community.guildKind === "alliance"
      ? [`Alliance: ${result.community.allianceAbbreviation}`] : []),
    "",
    "**Created:**",
    ...(result.created.length ? result.created.map(name => `• ${name}`) : ["• Nothing"]),
    "",
    "**Reused:**",
    ...result.reused.map(name => `• ${name}`),
    "",
    `Native booking: ${result.bookingStatus}`,
    "Event Scheduler: available via `/event-scheduler`"
  ].join("\n")
}

function setupPreview(profile, guildName, community, inspection, bookingStatus) {
  const terms = profileTerminology(profile)
  return [
    "**Community setup**", "",
    `${terms.locationLabel}: ${community.communityNumber}`,
    `Discord server: ${guildName}`,
    `Discord type: ${community.guildKind === "state" ? `${terms.locationLabel} Discord` : "Alliance Discord"}`,
    ...(community.guildKind === "alliance" ? [`Alliance: ${community.allianceAbbreviation}`] : []), "",
    "**Existing:**",
    ...(inspection.existing.length ? inspection.existing.map(name => `✓ ${name}`) : ["None"]),
    "", "**Missing:**",
    ...(inspection.missing.length ? inspection.missing.map(name => `• ${name}`) : ["None"]),
    "", `Native booking: ${bookingStatus}`,
  ].join("\n")
}

function createBotSetupService({ repository, client, gameProfile, botInstanceName,
  bookingWebsiteBaseUrl = null, logger = console }) {
  const terms = profileTerminology(gameProfile)
  const profile = PROFILE_NAMES[gameProfile]
  if (!profile) throw new Error("Unsupported game profile")

  async function fetchChannel(guild, id, expectedType) {
    if (!id) return null
    let channel
    try {
      channel = await guild.channels.fetch(id)
    } catch (error) {
      if (String(error?.code || "") === "10003") return null
      throw error
    }
    return channel?.type === expectedType ? channel : null
  }

  async function maintainCard(channel, messageId, payload) {
    if (messageId) {
      let message
      try {
        message = await channel.messages.fetch(messageId)
      } catch (error) {
        if (String(error?.code || "") !== "10008") throw error
      }
      if (message) {
        await message.edit(payload)
        return message.id
      }
    }
    return (await channel.send(payload)).id
  }

  async function listedChannels(guild) {
    const fetched = await guild.channels.fetch()
    if (!fetched?.values) throw new Error("Discord channel inventory is unavailable")
    return [...fetched.values()].filter(Boolean)
  }

  async function inspect(guildId) {
    const guild = await client.guilds.fetch(guildId)
    const stored = await repository.get(guildId) || {}
    const listed = await listedChannels(guild)
    let category = await fetchChannel(guild, stored.category_id, ChannelType.GuildCategory)
    if (!category) category = listed.find(channel =>
      channel.type === ChannelType.GuildCategory && channel.name === profile.categoryName) || null
    const channels = {}
    for (const definition of CHANNELS) {
      const idKey = `${definition.key}_channel_id`
      let channel = await fetchChannel(guild, stored[idKey], ChannelType.GuildText)
      if (!channel) channel = listed.find(candidate =>
        candidate.type === ChannelType.GuildText && candidate.name === definition.name) || null
      channels[definition.key] = channel
    }
    const existing = []
    const missing = []
    ;(category ? existing : missing).push(`Category: ${profile.categoryName}`)
    for (const definition of CHANNELS) {
      ;(channels[definition.key] ? existing : missing).push(`#${definition.name}`)
    }
    return { guild, stored, category, channels, existing, missing }
  }

  async function validDestination(guild, channelId) {
    return Boolean(await fetchChannel(guild, channelId, ChannelType.GuildText))
  }

  return Object.freeze({
    async preview(guildId, rawCommunity, bookingStatus = "ready to link") {
      const community = validateCommunitySetup(rawCommunity, gameProfile)
      const inspection = await inspect(guildId)
      return {
        content: setupPreview(gameProfile, inspection.guild.name, community, inspection, bookingStatus),
        community,
        guildName: inspection.guild.name,
        existing: inspection.existing,
        missing: inspection.missing,
      }
    },

    async reconcile(guildId, rawCommunity, bookingStatus = "linked", nativeBooking = {}) {
      const community = validateCommunitySetup(rawCommunity, gameProfile)
      const inspection = await inspect(guildId)
      const { guild, stored } = inspection
      const botMember = guild.members.me || await guild.members.fetchMe()
      if (!botMember.permissions.has(PermissionFlagsBits.ManageChannels)) {
        throw new BotSetupError(
          "MANAGE_CHANNELS_REQUIRED",
          "I need Manage Channels before setup can begin. Grant it to the existing bot role, then run /setup again. A reinvite is not required."
        )
      }

      const values = {
        category_id: stored.category_id || null,
        ...Object.fromEntries(CHANNELS.map(definition => [
          `${definition.key}_channel_id`, stored[`${definition.key}_channel_id`] || null
        ])),
        gift_auto_redeem_message_id: stored.gift_auto_redeem_message_id || null,
        minister_sign_up_message_id: stored.minister_sign_up_message_id || null,
        event_scheduler_message_id: stored.event_scheduler_message_id || null,
        community_number: community.communityNumber,
        discord_guild_name: guild.name,
        alliance_abbreviation: community.allianceAbbreviation
      }
      const persist = async () => repository.save(guildId, values)
      const created = []
      const reused = []
      let category = inspection.category
      if (!category) {
        category = await guild.channels.create({
          name: profile.categoryName,
          type: ChannelType.GuildCategory
        })
        created.push(`Category: ${profile.categoryName}`)
      } else reused.push(`Category: ${profile.categoryName}`)
      values.category_id = category.id
      await persist()

      const channels = {}
      for (const definition of CHANNELS) {
        const idKey = `${definition.key}_channel_id`
        let channel = inspection.channels[definition.key]
        if (!channel) {
          channel = await guild.channels.create({
            name: definition.name,
            type: ChannelType.GuildText,
            parent: category.id,
            permissionOverwrites: permissionOverwrites(guild, definition.threads, botMember)
          })
          created.push(`#${definition.name}`)
        } else {
          await channel.permissionOverwrites.set(
            permissionOverwrites(guild, definition.threads, botMember)
          )
          reused.push(`#${definition.name}`)
        }
        channels[definition.key] = channel
        values[idKey] = channel.id
        await persist()
      }

      values.gift_auto_redeem_message_id = await maintainCard(
        channels.gift_auto_redeem,
        values.gift_auto_redeem_message_id,
        giftCard(terms)
      )
      await persist()
      values.minister_sign_up_message_id = await maintainCard(
        channels.minister_sign_up,
        values.minister_sign_up_message_id,
        ministerCard(terms, bookingWebsiteBaseUrl)
      )
      await persist()
      values.event_scheduler_message_id = await maintainCard(
        channels.event_scheduler,
        values.event_scheduler_message_id,
        eventCard(terms, community.guildKind)
      )
      await persist()
      const destinations = await repository.getDestinations(guildId)
      const schedulerSummary = repository.getSchedulerSummary
        ? await repository.getSchedulerSummary(guildId)
        : { configured: Boolean(destinations.event), scheduledEvents: 0, stateLinked: false }
      const giftDestination = destinations.gift?.gift_code_channel_id
        && await validDestination(guild, destinations.gift.gift_code_channel_id)
        ? destinations.gift.gift_code_channel_id : channels.gift_announcements.id
      const eventDestination = destinations.event?.event_channel_id
        && await validDestination(guild, destinations.event.event_channel_id)
        ? destinations.event.event_channel_id : channels.event_announcements.id
      const existingRoundup = community.guildKind === "state"
        ? destinations.state?.state_roundup_channel_id
        : destinations.event?.weekly_roundup_channel_id
      const roundupDestination = existingRoundup && await validDestination(guild, existingRoundup)
        ? existingRoundup : channels.event_announcements.id
      await repository.reconcileDestinations({
        guildId, giftChannelId: giftDestination, eventChannelId: eventDestination,
        roundupChannelId: roundupDestination, botInstanceName,
        allianceAbbreviation: community.allianceAbbreviation,
        guildKind: community.guildKind, communityNumber: community.communityNumber
      })
      ;(nativeBooking.created ? created : reused).push("Native booking community")
      if (schedulerSummary.configured) reused.push("Existing event scheduler configuration")
      if (schedulerSummary.scheduledEvents) {
        reused.push(`${schedulerSummary.scheduledEvents} scheduled event${schedulerSummary.scheduledEvents === 1 ? "" : "s"}`)
      }
      if (schedulerSummary.stateLinked) reused.push("State/Kingdom roundup linkage")
      logger.log(JSON.stringify({
        event: "bot_setup_reconciled",
        game_profile: gameProfile,
        guild_id: guildId
      }))
      const result = {
        community: { ...community, guildName: guild.name }, created, reused,
        bookingStatus, channels, values,
        destinations: { giftDestination, eventDestination, roundupDestination }
      }
      return { ...result, content: setupStatus(gameProfile, result) }
    }
  })
}

module.exports = {
  BOT_SETUP_IDS,
  PROFILE_NAMES,
  CHANNELS,
  BotSetupError,
  permissionOverwrites,
  giftCard,
  ministerCard,
  eventCard,
  bookingWebsiteUrl,
  validateCommunitySetup,
  setupPreview,
  setupStatus,
  createBotSetupService
}
