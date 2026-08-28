# Booking, setup, and registration command migration

This matrix traces the registered command builders in `index.js` and the modular
command builders, their handlers, and their external calls. “Retired” means the
command is excluded from the new global registration payload and a stale command
or component receives a migration response before any legacy handler can run.

| Old/current command | Purpose | Traced backend | Replacement | Action |
| --- | --- | --- | --- | --- |
| `/setup` (legacy handler) | Create a community Sheet and registry entry | Apps Script `setup_state` | Canonical `/setup` | MERGE INTO /setup |
| `/bot-setup` | Reconcile bot category, channels, cards, and destinations | PostgreSQL `bot_managed_discord_setups`, gift/event settings; Discord channels | Canonical `/setup` using the same service | MERGE INTO /setup |
| `/setup-help` | Advertise legacy Sheet setup/write commands | Static Discord response | `/setup` status and deployment docs | MERGE INTO /setup |
| `/link-state` | Link a Discord guild to a legacy community | Apps Script `link_state` and registry Sheet | Signed native linkage in `/setup` | MERGE INTO /setup |
| `/unlink-state` | Remove a legacy guild/community link | Apps Script registry read and `unlink_state_server_by_id` | Booking Admin **Discord access** owner-only alliance unlink | REPLACED; old State/password meaning obsolete |
| `/linked-servers` | List legacy guild links | Apps Script `get_linked_servers_for_current_state` | Booking Admin **Discord access** linked-alliance list | REPLACED |
| `/reset-state-password` | Rotate legacy Sheet join password | Apps Script `reset_state_password` | Opaque website guest-link rotation | DEPRECATE |
| `/grant-access` | Grant email access to a community Sheet | Apps Script/Google Sheet permissions | Discord-authenticated Booking Admin | DEPRECATE |
| `/sheet-link` | Open the correctly guild-scoped existing historical Sheet | Apps Script read action `get_sheet_link_for_server` | Temporary emergency read-only navigation only | REMOVE AFTER ONE SUCCESSFUL NATIVE LIVE CYCLE |
| `/set-announcements` | Set legacy booking announcement destination | Apps Script `set_announcement_channel` | Native booking/scheduler destinations | REMOVE AFTER MIGRATION |
| `/settings` | Read/write booking limits and requirements | Apps Script community Sheet config | Website Booking Admin | DEPRECATE |
| `/open-bookings` | Open legacy bookings | Apps Script `open_bookings_for_server` | Booking Admin booking switch | DEPRECATE |
| `/close-bookings` | Close legacy bookings | Apps Script `close_bookings_for_server` | Booking Admin booking switch | DEPRECATE |
| `/clear-bookings` | Clear appointment-grid cells | Apps Script `clear_bookings_for_server` | No destructive native equivalent | DEPRECATE |
| `/set-booking-date` | Change a legacy service date | Apps Script `set_booking_date_for_server` | Fixed native service dates plus per-cycle open/close override | DEPRECATE |
| `/times` | Read legacy availability/date | Apps Script `get_times_for_server`, `get_booking_date_for_server` | Website appointment availability | LEGACY READ-ONLY |
| `/booking-link` | Return write-capable legacy public booking page | Apps Script booking deployment | Authenticated website or opaque guest link | DEPRECATE |
| `/book` | Create a legacy Sheet booking | Apps Script reads plus `book_for_server` | Native website booking | DEPRECATE |
| `/remove-booking` | Cancel caller's legacy Sheet booking | Apps Script `remove_booking_for_server` | Native website cancellation | DEPRECATE |
| `/my-bookings` | Read caller's current Sheet cells | Apps Script `get_my_bookings_for_server` | Native website My bookings | REPLACED; no Legacy tab/import |
| `/admin-add-booking` | Create a manager booking | Apps Script `admin_add_booking_for_server` | Native manager appointment board | DEPRECATE |
| `/admin-remove-booking` | Cancel a player's booking | Apps Script `admin_remove_booking_for_server` | Native manager board cancellation | DEPRECATE |
| `/admin-reserve-slots` | Reserve legacy Sheet cells | Apps Script read plus `admin_reserve_slots_for_server` | Native slot blocking | DEPRECATE |
| `/admin-remove-reserved` | Unreserve legacy Sheet cells | Apps Script `admin_remove_reserved_slots_for_server` | Native slot blocking | DEPRECATE |
| `/admin-help` | Advertise legacy manager writes | Static Discord response | `/setup`, Booking Admin, scheduler help | MERGE INTO /setup |
| `/help` | Advertise legacy player booking | Static Discord response | Website booking UI/docs | DEPRECATE |
| `/register` (legacy handler) | Write alliance/name/Player ID to `bot_users` | Apps Script `register_player_for_server` | Canonical `/register` | MERGE INTO /register |
| `/player-register` | Manage native player accounts in a panel | PostgreSQL `player_accounts` and gift-code services | Canonical `/register` using the same services | MERGE INTO /register |
| `/my-info` | Read legacy player registration | Apps Script `get_registered_player_for_server` | `/register` account view | MERGE INTO /register |
| `/unregister` | Delete legacy player registration | Apps Script `delete_registered_player_for_server` | `/register` account management | MERGE INTO /register |
| `/register` (canonical) | View/update one complete profile-derived identity and synchronize booking prefill | Bot PostgreSQL `player_accounts`; signed website registration route; native `booking_participants` projection | Same command | KEEP |
| `/setup` (canonical) | Preview/apply idempotent community and Discord-resource reconciliation | Bot PostgreSQL, Discord channels/messages, signed native community-link route | Same command | KEEP |
| `/event-scheduler` | Configure alliance/community events | Scheduler PostgreSQL | Separate discoverable command | KEEP |
| `/event-scheduler-help` | Read scheduler help | Static Discord response | Same command | KEEP |
| `/set-bot-admin-role` | Set custom bot-manager role | Native PostgreSQL setup row | Same command | KEEP |
| `/clear-bot-admin-role` | Clear custom bot-manager role | Native PostgreSQL setup row | Same command | KEEP |
| `/set-banter-channel` | Set banter channel | Deferred | Deferred | DORMANT |
| `/clear-banter-channel` | Clear banter channel | Deferred | Deferred | DORMANT |
| `/set-banter-spice` | Set banter tone | Deferred | Deferred | DORMANT |
| `/banter-test` | Run banter diagnostic | Deferred | Deferred | DORMANT |

Gift-code administration commands (`/gift-codes`, `/gift-codes-admin`,
`/gift-code-add`, and `/player-admin`) are native PostgreSQL specialised
operations, not duplicate booking/setup commands, and remain registered.

## Authority and remaining duplication

The command registration payload no longer includes the legacy booking, Sheet
setup, or Sheet player-registration commands. Dispatch also blocks their names
and every associated stale modal/select/button prefix before the old handler
code, so they cannot write new legacy booking data after this version is
deployed and command registration completes.

`player_accounts` is the canonical bot identity and now contains Player ID,
in-game name, State/Kingdom, and alliance. `/register` is the only registered bot
write path for that identity. Its signed profile-specific website call resolves
the Discord guild to the exact native community and upserts
`booking_participants` as the booking projection/prefill. A failed website call
does not fall back to Sheets or claim success: the bot retains the canonical
record and asks the player to rerun `/register`; the deterministic website
idempotency key makes that retry safe. Existing older bot accounts remain
readable and show `Update required` until the owner supplies the new fields.

`/setup` can link only an already bootstrapped active native community. It cannot
create a community or remap a guild already linked elsewhere. This preserves the
reviewed one-time community bootstrap boundary.

The legacy `bot_users` registration is no longer writable from registered or
stale bot interactions. Legacy public Apps Script `register`, `book`, and
`unbook` actions still exist in the undeployed reference and in the live legacy
deployment; disabling those deployment endpoints requires a separately reviewed
Apps Script deployment. Do not delete them until historical access and rollback
requirements are resolved.

## Exact Apps Script action classification

Still required solely by the registered legacy read-only navigation command:

- `/sheet-link`: `get_sheet_link_for_server`. It resolves the current Discord
  guild through the legacy registry and returns the existing Google Sheet URL;
  the command does not expose the returned write-capable booking-page URL.
  This is an emergency fallback for the next live cycle, not a supported
  mutation path, and should be deleted after that native cycle succeeds.

## Genuine remaining gaps

- Platform-admin tooling to unlink or remap the shared State/Kingdom Discord is
  intentionally deferred. Booking Admin cannot perform that operation.
- Removal of the final `/sheet-link` Apps Script read dependency waits only for
  one successful native live booking cycle.

There are no remaining player booking, manager booking, alliance-unlink,
registration, booking-window, announcement, or guest-link replacement gaps.

No other Apps Script action is required by a registered or passive live path.
Bot-manager authorization and its set/clear commands use
`bot_managed_discord_setups.bot_manager_role_id`, isolated by
`(game_profile, guild_id)`. Banter command registration and passive message
handling are dormant; their retained source is deferred and makes no Apps
Script call.

Booking/setup/registration legacy-only actions, no longer reachable from a
registered or stale bot interaction, are:

- community/Sheet setup: `setup_state`, `link_state`,
  `unlink_state_server`, `unlink_state_server_by_id`,
  `get_linked_servers_for_current_state`, `reset_state_password`,
  `grant_sheet_access_for_server`, and `set_announcement_channel`;
- booking configuration/availability: `get_booking_config_for_server`,
  `get_settings_for_server`, `update_setting_for_server`,
  `get_booking_link_for_server`, `get_times_for_server`,
  `get_booking_date_for_server`, `set_booking_date_for_server`,
  `open_bookings_for_server`, and `close_bookings_for_server`;
- player registration: `register_player_for_server`,
  `get_registered_player_for_server`, and
  `delete_registered_player_for_server`;
- booking mutations/history reads: `book_for_server`,
  `remove_booking_for_server`, `get_my_bookings_for_server`,
  `admin_add_booking_for_server`, `admin_remove_booking_for_server`,
  `clear_bookings_for_server`, `admin_reserve_slots_for_server`,
  `get_reserved_times_for_server`, and
  `admin_remove_reserved_slots_for_server`.

The public Apps Script actions `register`, `book`, and `unbook`, plus public
`doGet` booking rendering and `doGet action=times`, are also booking-legacy only.
They are outside Discord interaction dispatch and therefore must be disabled in
a later Apps Script deployment before the old public URL can be considered
fully read-only. This repository change deliberately does not alter or deploy
the immutable reference snapshots.
