const PROFILES = Object.freeze({
  wos: Object.freeze({
    gameProfile: "wos",
    gameName: "Whiteout Survival",
    locationLabel: "State",
    locationLabelLower: "state",
    locationPlural: "States",
    playerLabel: "Player ID",
    levelLabel: "Furnace level"
  }),
  kingshot: Object.freeze({
    gameProfile: "kingshot",
    gameName: "Kingshot",
    locationLabel: "Kingdom",
    locationLabelLower: "kingdom",
    locationPlural: "Kingdoms",
    playerLabel: "Player ID",
    levelLabel: "Town level"
  })
})

function profileTerminology(gameProfile) {
  const profile = PROFILES[gameProfile]
  if (!profile) throw new Error("Unsupported game profile")
  return profile
}

module.exports = { PROFILES, profileTerminology }
