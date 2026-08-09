# Event Scheduler

## Initial Setup

Run `/event-scheduler` as a user accepted by `userCanManageServer(interaction)`.

Run `/event-scheduler-help` for private setup, creation, reminder, management, roundup, state-link and troubleshooting pages. Help is also available from the scheduler management home and does not require a healthy database connection.

1. Choose **Alliance identity** and set or rename the main alliance in the text modal.
2. Choose **Configure channels** and select the shared alliance reminder channel from Discord's native channel list. The bot validates View Channel, Send Messages, Embed Links, and Attach Files.
3. Choose **Weekly roundup settings** to enable or disable the alliance roundup, independently control state publishing, select the roundup channel, and preview or change its UTC weekday/time. Roundups do not require Attach Files.
4. Open **Alliances** to add, list, rename, or delete sub-alliances.
5. For state sharing, invite the profile's bot to the state Discord. Run `/event-scheduler` there, choose **State destination**, select its weekly-roundup channel, and generate a one-time 15-minute link code. Back in the alliance Discord, choose **State sharing**, then **Link with code**. Both guilds and the channel are identified automatically.
6. In that enabled state destination, set the state number/name and use **State events** to create and manage canonical state events. Linked alliance Discords do not receive this administration control.

The scheduler state link is Postgres-backed and separate from existing Apps Script state linking. Alliance and state channels here do not reuse the existing Apps Script announcement channel. Codes are stored only as hashes, expire, work once, and cannot cross game profiles. Existing ID-backed configurations continue to work internally, but normal setup never asks users for guild or channel IDs.

## Alliances

One guild/profile may contain a main alliance and multiple sub-alliances. Names are unique without regard to case inside that scope. WOS and Kingshot can use the same names without sharing rows.

The list and event selectors use opaque session values; database IDs are not displayed. Renaming an alliance updates its events, increments their schedule versions, and reconciles unsent claims. The main alliance cannot be deleted. A sub-alliance can be deleted only when no event history references it; otherwise the user receives a blocked-delete message.

Channels remain guild/profile settings, so alliances may share the alliance reminder channel, alliance roundup channel, and linked state Discord.

## Creating Or Editing An Event

The private flow is:

1. Select an alliance or sub-alliance.
2. Enter the event name and first date.
3. Choose **Single time** and enter one UTC time, or choose **Multiple groups** and add each group through separate name/time fields. Groups may share a time; names must be unique without regard to case.
4. Choose recurrence: 2, 3, 7, 14, 21, 28, 35, or 42 days.
5. Choose one advance reminder: none, 5, 10, 15, 20, or 30 minutes.
6. Optionally enter separate advance and final custom messages, each up to 500 characters.
7. Enable or disable the one-minute final announcement.
8. Use **Manage image** to keep, replace, or remove one PNG, JPEG, GIF, or WebP image up to 8 MB.
9. Choose alliance reminders, alliance weekly roundup inclusion, and state weekly-roundup eligibility.
10. Review the clearly labelled UTC and Discord-local preview and confirm.

No event is saved before final confirmation. **Edit details** retains the current alliance and image automatically. **Change alliance** is separate and preserves all other event properties. **Manage image** is the only flow that retains, replaces, or removes the image. Invalid image replacement rolls back the complete edit and retains the prior image.

Selecting no advance reminder cancels future advance claims. Disabling the final announcement cancels future final claims. Clearing either optional message restores default wording. These edits increment `schedule_version`, invalidate unsent claims from the previous version, and preserve sent history.

Accepted time forms include `18:30`, `1830`, `1800`, `6:30pm`, and `6.30 PM`. Times are normalized to UTC. Ambiguous or invalid values are rejected rather than guessed.

## State Events

**State events** appears on `/event-scheduler` only when the current guild is the enabled state destination for the current game profile. Configure or change the displayed state number/name under **State destination** first. Administration is ephemeral and session controls are scoped to the requesting user, guild, and profile.

Choose **Create state event**, enter its name and first occurrence date, then select one-time or recurrence every 2, 3, 7, 14, 21, 28, 35, or 42 days. Every 2 days is a true anchored 48-hour interval. Add at least one structured phase before review and confirmation; nothing is stored before confirmation.

Each phase has a name and free-form UTC time parsed by the same parser as alliance events. Supported examples include `1000`, `10:00`, `10am`, `10:00am`, `1830`, `18:30`, and `6:30pm`. The editor shows the normalized UTC time and a Discord-local timestamp. Configure a pre-alert of none, 5, 10, 15, 20, or 30 minutes and independently turn the exact-time announcement on or off. Exact-time alerts may cover the game screen, so consider disabling them for critical battle or red-zone phases.

Pre-alert and exact-time messages are separate. Each delivery kind also has independent keep/upload, replace, and remove controls for PNG, JPEG, GIF, or WebP media up to 8 MB. Editing unrelated phase settings retains existing media. Discord's native GIF picker is not used.

**View state events** lists active and paused events. Selecting one provides next-occurrence preview, event-detail editing, phase management, pause/resume, soft delete, and a test announcement. Preview reads the schedule without creating claims. A test sends one clearly marked pre-alert or exact-time version to the configured state destination only; it does not create or consume production claims, change `schedule_version`, or alter sent history.

Confirmed edits update the existing event and reconcile unsent older-version claims. Stable phase rows and sent history are preserved. Pause removes the event from future alerts and roundups while retaining its anchor. Resume schedules only future work from that anchor. Soft delete stops future alerts and roundup entries while retaining sent history.

## Reminder Semantics

`advance_reminder_minutes` stores one nullable value from `5`, `10`, `15`, `20`, or `30`. One advance claim is generated for each event/group occurrence.

`reminder_at_start` is a legacy column name. When true, it creates one `final_reminder` claim with `deliver_at = occurrence_at - 1 minute`. For an 18:00 event, the final announcement is due at 17:59. The Discord handler uses the stored `deliver_at` and does not subtract another minute. Nothing is generated for 18:00 itself.

The final announcement always retains **About to start** and **Starts in approximately 1 minute**. Immediate reminders contain only alliance, event, optional group, countdown, and optional custom text. They omit dates, UTC/local timestamps, recurrence, and field labels; those details belong in previews and roundups.

The advance custom message appears only on the advance reminder. The final custom message appears only on the final announcement. Blank or whitespace-only input stores `NULL` and uses default wording. Mentions do not ping.

## Images

An image is attached only to the alliance advance reminder, at whichever valid offset is selected. It is never attached to the final announcement, a state Discord, a roundup, an exact-start legacy row, or a management preview.

If no advance reminder is selected, the image remains stored but produces no image-only message. Enabling an advance reminder later makes it eligible again. A retry follows the same database claim and has the documented at-least-once ambiguity.

## Publishing Policy

Individual advance, final, grouped, and image deliveries are alliance-only. No state event delivery claim is generated. Historical sent state and exact-start rows remain unchanged; unsent legacy rows are terminally failed and never redirected.

`publish_to_state` is retained as a legacy compatibility field. Individual alliance-event reminders remain alliance-only; linked state roundup inclusion follows weekly-roundup inclusion and enabled sharing.

## Help

`/event-scheduler-help` uses one select menu to keep every page inside Discord's message and component limits. Sections cover native channel setup, event creation, accepted UTC formats, reminders/images, event management and cancellation, alliance roundups, two-sided state linking, state-only-roundup behavior, and troubleshooting. Help responses disable mention parsing and are always ephemeral.

## Weekly Roundups

The default schedule is Monday at 09:00 UTC with a 60-minute catch-up grace. The window begins on the configured UTC weekday and is half-open through the same weekday one week later. **Weekly roundup settings** always remains available: it can enable/disable the alliance roundup, independently toggle state publishing, select the channel, and preview/save any UTC weekday/time.

Disabling retains the configured channel and schedule. Changing the configuration terminally invalidates obsolete unsent roundup work, preserves sent history, and sets a replay boundary. Re-enabling or changing the schedule begins with the next future scheduled roundup and does not replay a missed period.

Alliance roundup eligibility requires matching guild/profile, active status, `include_in_weekly_roundup=true`, and an occurrence in the window. It is independent of `publish_to_state` and includes separate alliance names for main/sub-alliance entries.

State roundup eligibility for alliance events additionally requires an enabled valid profile-scoped state link and matching state destination. Canonical state events add a single **STATE EVENTS** section to both the state roundup and each linked alliance roundup. Every phase is a milestone even when its exact alert is off; pre-alerts are not milestones. Multiple alliance identities in one Discord do not duplicate the section. WOS and Kingshot are never mixed.

Roundups are text/embed-only and disable mention parsing. Long output splits deterministically. Each sent part ID is stored immediately and recorded parts are skipped on retry; the parent claim is sent only when all parts are recorded.

## Editing, Pause, Resume, Delete

Every confirmed edit increments `schedule_version`, preserves sent history, and terminally fails all unsent older claims. Grouped claim snapshots preserve the original group identity and name when an edit replaces group rows. New generation uses the updated version and current alliance, timing, messages, and publishing settings. Previously sent Discord messages are not edited.

Pause suppresses new claims and roundups without shifting the recurrence anchor. Resume returns to the original anchor and bounded generation window. Delete sets event status to `deleted`; it does not physically remove event or sent-claim history.
