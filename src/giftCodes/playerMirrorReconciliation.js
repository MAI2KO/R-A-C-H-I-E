const PROFILES = new Set(["wos", "kingshot"])

function requestAccount(account, profile) {
  if (account.game_profile !== profile || account.is_active !== true
      || !account.discord_user_id) throw new Error("invalid active account reconciliation scope")
  return {
    discordUserId: account.discord_user_id,
    playerId: account.player_id,
    inGameName: account.in_game_name,
    communityCode: account.state_or_kingdom_number,
    allianceAbbreviation: account.alliance_abbreviation,
    isPrimary: account.is_primary === true
  }
}

function summarize(accounts, responses) {
  const results = responses.flatMap(response => response.reconciliation?.results || [])
  const count = status => results.filter(result => result.websiteMirrorStatus === status).length
  return {
    activeBotAccounts: accounts.length,
    matching: count("matching"),
    missingMirrors: count("missing"),
    primaryMismatches: count("primary mismatch"),
    ownershipConflicts: count("ownership mismatch"),
    ambiguousConflicted: count("ambiguous/conflict"),
    unresolvedCommunities: results.filter(result =>
      String(result.plannedAction).includes("unresolved community")).length,
    plannedCreates: responses.reduce((total, response) =>
      total + Number(response.reconciliation?.plannedCreates || 0), 0),
    plannedUpdates: responses.reduce((total, response) =>
      total + Number(response.reconciliation?.plannedUpdates || 0), 0),
    skipped: results.filter(result => String(result.plannedAction).startsWith("skip:")).length,
    mutations: responses.reduce((total, response) =>
      total + Number(response.reconciliation?.mutations || 0), 0)
  }
}

async function runPlayerMirrorReconciliation({ profile, repository, api, dryRun }) {
  if (!PROFILES.has(profile) || repository.gameProfile !== profile
      || api.profile !== profile || typeof dryRun !== "boolean") {
    throw new Error("player mirror reconciliation profile mismatch")
  }
  const accounts = await repository.listActiveAccountsForReconciliation()
  const grouped = new Map()
  for (const source of accounts) {
    const account = requestAccount(source, profile)
    const group = grouped.get(account.discordUserId) || []
    group.push(account)
    grouped.set(account.discordUserId, group)
  }
  if ([...grouped.values()].some(group => group.length > 100)) {
    throw new Error("player mirror reconciliation owner exceeds 100 active accounts")
  }
  const responses = []
  for (const ownerAccounts of grouped.values()) {
    responses.push(dryRun
      ? await api.playerMirrorPreview(ownerAccounts)
      : await api.playerMirrorExecute(ownerAccounts))
  }
  return Object.freeze({ profile, mode: dryRun ? "dry-run" : "execute",
    summary: summarize(accounts, responses), results: responses.flatMap(
      response => response.reconciliation?.results || []) })
}

module.exports = { requestAccount, summarize, runPlayerMirrorReconciliation }
