function createNoopPlayerMirror() {
  return Object.freeze({
    async mirrorRegistration() {
      return { mirrored: false, reason: "not configured" }
    }
  })
}

module.exports = { createNoopPlayerMirror }
