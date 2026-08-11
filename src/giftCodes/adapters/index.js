const { whiteoutAdapter } = require("./whiteout")
const { kingshotAdapter } = require("./kingshot")

function centuryAdapter(gameProfile, env = process.env) {
  if (gameProfile === "wos") return whiteoutAdapter(env)
  if (gameProfile === "kingshot") return kingshotAdapter(env)
  throw new Error("Unsupported game profile")
}

module.exports = { centuryAdapter }
