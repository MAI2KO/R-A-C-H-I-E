const crypto = require("crypto")

class InteractionSessionError extends Error {
  constructor(message, code = "SESSION_INVALID") {
    super(message)
    this.name = "InteractionSessionError"
    this.code = code
  }
}

class InteractionSessionStore {
  constructor({ ttlMs = 15 * 60 * 1000, maximumSessions = 25, now = () => Date.now() } = {}) {
    this.ttlMs = ttlMs
    this.maximumSessions = maximumSessions
    this.now = now
    this.sessions = new Map()
    this.expiredSessions = new Map()
    this.cleanupTimer = setInterval(
      () => this.cleanup(),
      Math.min(ttlMs, 60 * 1000)
    )
    this.cleanupTimer.unref?.()
  }

  cleanup() {
    const now = this.now()
    for (const [id, session] of this.sessions) {
      if (session.expiresAt <= now) {
        this.sessions.delete(id)
        this.expiredSessions.set(id, now + this.ttlMs)
      }
    }
    for (const [id, purgeAt] of this.expiredSessions) {
      if (purgeAt <= now) this.expiredSessions.delete(id)
    }
  }

  create({ userId, guildId, gameProfile }, data = {}) {
    this.cleanup()
    if (this.sessions.size >= this.maximumSessions) {
      throw new InteractionSessionError(
        "Too many event setups are active. Try again shortly.",
        "SESSION_LIMIT"
      )
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
    if (!session) {
      if (this.expiredSessions.has(id)) {
        throw new InteractionSessionError(
          "This event setup has expired. Start again.",
          "SESSION_EXPIRED"
        )
      }
      throw new InteractionSessionError(
        "This event setup has expired. Start again.",
        "SESSION_NOT_FOUND"
      )
    }
    if (session.expiresAt <= this.now()) {
      this.sessions.delete(id)
      this.expiredSessions.set(id, this.now() + this.ttlMs)
      throw new InteractionSessionError(
        "This event setup has expired. Start again.",
        "SESSION_EXPIRED"
      )
    }
    if (session.userId !== userId) {
      throw new InteractionSessionError(
        "This event setup belongs to another user.",
        "SESSION_WRONG_USER"
      )
    }
    if (session.guildId !== guildId) {
      throw new InteractionSessionError(
        "This event setup belongs to another Discord server.",
        "SESSION_WRONG_GUILD"
      )
    }
    if (session.gameProfile !== gameProfile) {
      throw new InteractionSessionError(
        "This event setup belongs to another game profile.",
        "SESSION_WRONG_PROFILE"
      )
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
    this.expiredSessions.set(id, this.now() + this.ttlMs)
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
