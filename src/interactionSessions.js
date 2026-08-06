const crypto = require("crypto")

class InteractionSessionError extends Error {
  constructor(message) {
    super(message)
    this.name = "InteractionSessionError"
  }
}

class InteractionSessionStore {
  constructor({ ttlMs = 15 * 60 * 1000, maximumSessions = 25, now = () => Date.now() } = {}) {
    this.ttlMs = ttlMs
    this.maximumSessions = maximumSessions
    this.now = now
    this.sessions = new Map()
    this.cleanupTimer = setInterval(
      () => this.cleanup(),
      Math.min(ttlMs, 60 * 1000)
    )
    this.cleanupTimer.unref?.()
  }

  cleanup() {
    const now = this.now()
    for (const [id, session] of this.sessions) {
      if (session.expiresAt <= now) this.sessions.delete(id)
    }
  }

  create({ userId, guildId, gameProfile }, data = {}) {
    this.cleanup()
    if (this.sessions.size >= this.maximumSessions) {
      throw new InteractionSessionError("Too many event setups are active. Try again shortly.")
    }
    const id = crypto.randomBytes(12).toString("base64url")
    this.sessions.set(id, {
      id,
      userId,
      guildId,
      gameProfile,
      expiresAt: this.now() + this.ttlMs,
      data
    })
    return id
  }

  get(id, { userId, guildId, gameProfile }) {
    const session = this.sessions.get(id)
    if (!session || session.expiresAt <= this.now()) {
      this.sessions.delete(id)
      throw new InteractionSessionError("This event setup has expired. Start again.")
    }
    if (session.userId !== userId) {
      throw new InteractionSessionError("This event setup belongs to another user.")
    }
    if (session.guildId !== guildId) {
      throw new InteractionSessionError("This event setup belongs to another Discord server.")
    }
    if (session.gameProfile !== gameProfile) {
      throw new InteractionSessionError("This event setup belongs to another game profile.")
    }
    return session
  }

  update(id, context, changes) {
    const session = this.get(id, context)
    session.data = { ...session.data, ...changes }
    return session
  }

  cancel(id, context) {
    const session = this.get(id, context)
    this.sessions.delete(id)
    return session.data
  }

  complete(id, context) {
    return this.cancel(id, context)
  }
}

module.exports = {
  InteractionSessionError,
  InteractionSessionStore
}
