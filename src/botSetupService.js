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
      `- ${terms.locationLabel}`,
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

function ministerCard(terms) {
  return {
    content: [
      "**Minister Sign-Up**",
      "",
      `Register your ${terms.gameName} player details for this server's existing minister booking system.`,
      "You will need your alliance tag, in-game name, and Player ID."
    ].join("\n"),
    components: [new ActionRowBuilder().addComponents(
      new ButtonBuilder()
        .setCustomId(BOT_SETUP_IDS.registerMinisters)
        .setLabel("REGISTER FOR MINISTERS")
        .setStyle(ButtonStyle.Primary)
    )]
  }
}

function eventCard(terms) {
  return {
    content: [
      "**Event Scheduler**",
      "",
      "Authorised event managers can use `/event-scheduler` to create and manage:",
      "- alliance events, groups, recurrence, and reminders",
      `- ${terms.locationLabel}-wide events and linked ${terms.locationLabel} servers`,
      "- weekly alliance and community roundups",
      "- existing event configuration and delivery destinations"
    ].join("\n"),
    components: []
  }
}

function setupStatus(profile, channels) {
  const names = PROFILE_NAMES[profile]
  const labels = [
    ["Category", names.categoryName],
    ["Gift Auto-Redeem", channels.gift_auto_redeem.name],
    ["Gift Announcements", channels.gift_announcements.name],
    ["Minister Sign-Up", channels.minister_sign_up.name],
    ["Event Scheduler", channels.event_scheduler.name],
    ["Event Announcements", channels.event_announcements.name]
  ]
  return [
    `**${names.botName} Setup**`,
    "",
    ...labels.map(([label, value]) => `${label}: Configured (${value})`),
    "Permissions: Ready"
  ].join("\n")
}

function createBotSetupService({ repository, client, gameProfile, logger = console }) {
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

  return Object.freeze({
    async reconcile(guildId) {
      const guild = await client.guilds.fetch(guildId)
      const botMember = guild.members.me || await guild.members.fetchMe()
      if (!botMember.permissions.has(PermissionFlagsBits.ManageChannels)) {
        throw new BotSetupError(
          "MANAGE_CHANNELS_REQUIRED",
          "I need Manage Channels before setup can begin. Grant it to the existing bot role, then run /bot-setup again. A reinvite is not required."
        )
      }

      const stored = await repository.get(guildId) || {}
      const values = {
        category_id: stored.category_id || null,
        ...Object.fromEntries(CHANNELS.map(definition => [
          `${definition.key}_channel_id`, stored[`${definition.key}_channel_id`] || null
        ])),
        gift_auto_redeem_message_id: stored.gift_auto_redeem_message_id || null,
        minister_sign_up_message_id: stored.minister_sign_up_message_id || null,
        event_scheduler_message_id: stored.event_scheduler_message_id || null
      }
      const persist = async () => repository.save(guildId, values)
      let category = await fetchChannel(guild, stored.category_id, ChannelType.GuildCategory)
      if (!category) {
        category = await guild.channels.create({
          name: profile.categoryName,
          type: ChannelType.GuildCategory
        })
      }
      values.category_id = category.id
      await persist()

      const channels = {}
      for (const definition of CHANNELS) {
        const idKey = `${definition.key}_channel_id`
        let channel = await fetchChannel(guild, stored[idKey], ChannelType.GuildText)
        if (!channel || channel.parentId !== category.id) {
          channel = await guild.channels.create({
            name: definition.name,
            type: ChannelType.GuildText,
            parent: category.id,
            permissionOverwrites: permissionOverwrites(guild, definition.threads, botMember)
          })
        } else {
          await channel.permissionOverwrites.set(
            permissionOverwrites(guild, definition.threads, botMember)
          )
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
        ministerCard(terms)
      )
      await persist()
      values.event_scheduler_message_id = await maintainCard(
        channels.event_scheduler,
        values.event_scheduler_message_id,
        eventCard(terms)
      )
      await persist()
      await repository.reconcileDestinations(
        guildId,
        channels.gift_announcements.id,
        channels.event_announcements.id
      )
      logger.log(JSON.stringify({
        event: "bot_setup_reconciled",
        game_profile: gameProfile,
        guild_id: guildId
      }))
      return { content: setupStatus(gameProfile, channels), values }
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
  setupStatus,
  createBotSetupService
}
