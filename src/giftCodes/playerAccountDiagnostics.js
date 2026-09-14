const { createHash } = require("node:crypto")

const {
  normalizeDiscordUserId,
  normalizeLocationNumber,
  normalizeInGameName,
  normalizeAllianceAbbreviation
} = require("./validation")

function valid(work) {
  try { work(); return true } catch { return false }
}

function accountReference(profile, id) {
  return createHash("sha256").update(`${profile}\0${id}`, "utf8").digest("hex").slice(0, 16)
}

function diagnoseActiveAccount(account, profile) {
  const reasons = []
  if (!valid(() => normalizeLocationNumber(account.state_or_kingdom_number, "Location"))) {
    reasons.push("invalid_location")
  }
  if (!valid(() => normalizeInGameName(account.in_game_name))) reasons.push("invalid_in_game_name")
  if (!valid(() => normalizeAllianceAbbreviation(account.alliance_abbreviation))) {
    reasons.push("invalid_alliance")
  }
  if (!valid(() => normalizeDiscordUserId(account.discord_user_id))) reasons.push("invalid_owner")
  if (!reasons.length) return null
  return Object.freeze({
    accountRef: accountReference(profile, account.id),
    playerId: String(account.player_id),
    discordUserId: valid(() => normalizeDiscordUserId(account.discord_user_id))
      ? String(account.discord_user_id) : null,
    reasons: Object.freeze(reasons)
  })
}

function summarize(diagnostics, websiteResults) {
  const byReason = reason => diagnostics.filter(item => item.reasons.includes(reason)).length
  return Object.freeze({
    invalidActiveAccounts: diagnostics.length,
    invalidLocation: byReason("invalid_location"),
    invalidInGameName: byReason("invalid_in_game_name"),
    invalidAlliance: byReason("invalid_alliance"),
    ownershipProblems: byReason("invalid_owner"),
    withParticipantMirrors: websiteResults.filter(result => result.participantMirrors > 0).length,
    withBookings: websiteResults.filter(result => result.bookingReferences > 0).length,
    withHistory: websiteResults.filter(result => result.historyReferences > 0).length
  })
}

async function auditInvalidPlayerAccounts({ profile, repository, api, releaseRef = null,
  operatorDiscordUserId = null }) {
  if (repository.gameProfile !== profile || api.profile !== profile) {
    throw new Error("player account diagnostic profile mismatch")
  }
  const accounts = await repository.listActiveAccountsForDiagnostics()
  const diagnosed = accounts.map(account => ({ account,
    diagnostic: diagnoseActiveAccount(account, profile) })).filter(item => item.diagnostic)
  const publicCandidates = diagnosed.map(item => item.diagnostic)
  if (releaseRef === null) {
    const results = []
    for (let offset = 0; offset < publicCandidates.length; offset += 100) {
      const response = await api.playerAccountCleanupPreview(
        publicCandidates.slice(offset, offset + 100)
      )
      results.push(...(response.cleanup?.results || []))
    }
    return Object.freeze({ profile, mode: "dry-run", summary: summarize(publicCandidates, results),
      accounts: Object.freeze(results) })
  }
  if (!/^[a-f0-9]{16}$/.test(releaseRef) || !operatorDiscordUserId) {
    throw new Error("invalid cleanup release request")
  }
  const matches = diagnosed.filter(item => item.diagnostic.accountRef === releaseRef)
  if (matches.length > 1) throw new Error("cleanup account reference is ambiguous")
  let completed = null
  if (matches.length === 0) completed = await repository.findCompletedPlayerCleanup(releaseRef)
  if (matches.length === 0 && !completed) throw new Error("cleanup account reference not found")
  const target = matches[0] || {
    account: completed.account,
    diagnostic: {
      accountRef: releaseRef,
      playerId: String(completed.account.player_id),
      discordUserId: completed.previousOwnerDiscordUserId,
      reasons: Object.freeze(completed.cleanupMetadata?.reasons || [])
    }
  }
  normalizeDiscordUserId(operatorDiscordUserId)
  const response = await api.playerAccountCleanupExecute([target.diagnostic])
  const website = response.cleanup?.results?.[0]
  if (!website || !["deactivated", "already_completed", "safe_noop"]
    .includes(website.cleanupStatus)) {
    throw new Error("website participant cleanup was not safe")
  }
  let released = completed
  if (!released) {
    try {
      released = target.diagnostic.discordUserId === null
        ? await repository.deactivateInvalidUnownedAccount({
          accountId: target.account.id, playerId: target.account.player_id,
          operatorDiscordUserId, accountRef: releaseRef,
          reasons: target.diagnostic.reasons
        })
        : await repository.releaseAccount({
          playerId: target.account.player_id,
          performedByDiscordUserId: operatorDiscordUserId,
          actionType: "operator_release",
          expectedOwnerDiscordUserId: target.account.discord_user_id,
          expectedAccountId: target.account.id,
          sourceMetadata: { source: "invalid_player_cleanup", accountRef: releaseRef,
            reasons: target.diagnostic.reasons }
        })
    } catch (error) {
      released = await repository.findCompletedPlayerCleanup(releaseRef)
      if (!released) throw error
    }
  }
  if (!released) released = await repository.findCompletedPlayerCleanup(releaseRef)
  if (!released) throw new Error("player ownership changed before cleanup")
  if (released.replacement) {
    await api.registration({
      discordUserId: released.replacement.discord_user_id,
      playerId: released.replacement.player_id,
      isPrimary: true,
      primarySyncOnly: true
    })
  }
  return Object.freeze({ profile, mode: "release", accountRef: releaseRef,
    reasons: target.diagnostic.reasons, website, botAccountReleased: true,
    botReleaseAlreadyCompleted: released.alreadyCompleted === true,
    replacementPrimaryMirrored: Boolean(released.replacement) })
}

module.exports = { accountReference, diagnoseActiveAccount, summarize, auditInvalidPlayerAccounts }
