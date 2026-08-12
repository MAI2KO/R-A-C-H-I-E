# Century Gift-Code Response Protocol

## Evidence and extraction

Century responses are classified by numeric `err_code` before any other field. The tables below were recovered statically from the official public WOS and Kingshot browser bundles: only each bundle's string-array decoder and rotation initializer were evaluated, then every hexadecimal key in `errorMsgMap` was decoded to its referenced localization key. The browser applications themselves were not executed and no network request was made.

Evidence labels:

- **LIVE VERIFIED**: observed in a captured official request and response.
- **OFFICIAL FRONTEND VERIFIED**: the official numeric map and supplied official localization semantics establish the result.
- **UNRESOLVED**: the numeric key was extracted, but the supplied localization evidence does not establish safe bot behavior.

`unknown_response` rows are intentionally not adapter mappings. They remain review outcomes.

## Whiteout Survival

| `err_code` | Hex | Frontend key | Meaning | Bot classification | Retry / scope | Evidence |
|---:|---:|---|---|---|---|---|
| 20000 | `0x4e20` | `error_msg5` | Redemption accepted; reward sent to mail | `success` | Terminal; account result proving code valid | LIVE + OFFICIAL FRONTEND VERIFIED |
| 40001 | `0x9c41` | `input_roleId_tips3` | Exact semantic not established | `unknown_response` | Review; unresolved | UNRESOLVED |
| 40002 | `0x9c42` | `frequency_tip1` | Exact semantic not established | `unknown_response` | Review; unresolved | UNRESOLVED |
| 40003 | `0x9c43` | `exchange_tips4` | Exact semantic not established | `unknown_response` | Review; unresolved | UNRESOLVED |
| 40004 | `0x9c44` | `frequency_tip2` | Exact semantic not established | `unknown_response` | Review; unresolved | UNRESOLVED |
| 40005 | `0x9c45` | `error_msg3` | Claim limit exceeded | `claim_limit` | Terminal attempt; account-specific | OFFICIAL FRONTEND VERIFIED |
| 40006 | `0x9c46` | `error_msg4` | Insufficient Town Center / City level | `level_restriction` | Terminal attempt; account-specific | OFFICIAL FRONTEND VERIFIED |
| 40007 | `0x9c47` | `error_msg2` | Redemption period expired | `expired` | Terminal; code-global | LIVE + OFFICIAL FRONTEND VERIFIED |
| 40008 | `0x9c48` | `error_msg1` | Exact gift already claimed | `already_redeemed` | Terminal; account-specific, proves code valid | LIVE + OFFICIAL FRONTEND VERIFIED |
| 40009 | `0x9c49` | `exchange_tips5` | Exact semantic not established | `unknown_response` | Review; unresolved | UNRESOLVED |
| 40011 | `0x9c4b` | `error_msg7` | Another code of the same reward type was redeemed | `redemption_limit` | Terminal attempt; account-specific, proves code valid | OFFICIAL FRONTEND VERIFIED |
| 40012 | `0x9c4c` | `error_msg12` | Account age does not meet requirements | `account_age_restriction` | Terminal attempt; account-specific | OFFICIAL FRONTEND VERIFIED |
| 40014 | `0x9c4e` | `error_msg6` | Gift code not found; capitalization may be wrong | `invalid_code` | Terminal; code-global | LIVE + OFFICIAL FRONTEND VERIFIED |
| 40015 | `0x9c4f` | `error_msg8` | Generic redemption-code error | `unknown_response` | No automatic retry; review | OFFICIAL FRONTEND VERIFIED; behavior unresolved |
| 40016 | `0x9c50` | `error_msg10` | Server busy; reward delivery may be deferred | `unknown_response` | No automatic retry; duplicate-delivery risk | OFFICIAL FRONTEND VERIFIED; behavior unresolved |
| 40017 | `0x9c51` | `error_msg11` | Account does not meet redemption requirements | `account_restriction` | Terminal attempt; account-specific | OFFICIAL FRONTEND VERIFIED |
| 40018 | `0x9c52` | `error_msg18` | Localization meaning not supplied | `unknown_response` | Review; unresolved | UNRESOLVED |
| 40019 | `0x9c53` | `error_msg20` | Too many simultaneous actions | `simultaneous_action_throttle` | Existing bounded `temporary_error` retry state; transient | OFFICIAL FRONTEND VERIFIED |
| 40020 | `0x9c54` | `error_msg19` | Character/player information incorrect | `invalid_player` | Terminal; account-specific configuration issue | LIVE VERIFIED; numeric key OFFICIAL FRONTEND VERIFIED |
| 40100 | `0x9ca4` | `error_msg13` | Verification refresh requested too frequently | `verification_throttle` | Existing bounded `rate_limited` retry state; transient | OFFICIAL FRONTEND VERIFIED |
| 40101 | `0x9ca5` | `frequency_tip2` | Exact semantic not established | `unknown_response` | Review; unresolved | UNRESOLVED |
| 40102 | `0x9ca6` | `error_msg15` | Verification code expired | `verification_error` | No automatic retry; request/session-specific review | OFFICIAL FRONTEND VERIFIED |
| 40103 | `0x9ca7` | `error_msg14` | Verification code incorrect | `verification_error` | No automatic retry; request/session-specific review | OFFICIAL FRONTEND VERIFIED |

## Kingshot

| `err_code` | Hex | Frontend key | Meaning | Bot classification | Retry / scope | Evidence |
|---:|---:|---|---|---|---|---|
| 20000 | `0x4e20` | `error_msg5` | Redemption accepted; reward sent to mail | `success` | Terminal; account result proving code valid | LIVE + OFFICIAL FRONTEND VERIFIED |
| 40001 | `0x9c41` | `input_roleId_tips3` | Exact semantic not established | `unknown_response` | Review; unresolved | UNRESOLVED |
| 40002 | `0x9c42` | `frequency_tip1` | Exact semantic not established | `unknown_response` | Review; unresolved | UNRESOLVED |
| 40003 | `0x9c43` | `exchange_tips4` | Exact semantic not established | `unknown_response` | Review; unresolved | UNRESOLVED |
| 40004 | `0x9c44` | `frequency_tip2` | Exact semantic not established | `unknown_response` | Review; unresolved | UNRESOLVED |
| 40005 | `0x9c45` | `error_msg3` | Claim limit exceeded | `claim_limit` | Terminal attempt; account-specific | OFFICIAL FRONTEND VERIFIED |
| 40006 | `0x9c46` | `error_msg4` | Insufficient Town Center / City level | `level_restriction` | Terminal attempt; account-specific | OFFICIAL FRONTEND VERIFIED |
| 40007 | `0x9c47` | `error_msg2` | Redemption period expired | `expired` | Terminal; code-global | LIVE + OFFICIAL FRONTEND VERIFIED |
| 40008 | `0x9c48` | `error_msg1` | Exact gift already claimed | `already_redeemed` | Terminal; account-specific, proves code valid | LIVE + OFFICIAL FRONTEND VERIFIED |
| 40009 | `0x9c49` | `exchange_tips5` | Exact semantic not established | `unknown_response` | Review; unresolved | UNRESOLVED |
| 40011 | `0x9c4b` | `error_msg7` | Another code of the same reward type was redeemed | `redemption_limit` | Terminal attempt; account-specific, proves code valid | OFFICIAL FRONTEND VERIFIED |
| 40012 | `0x9c4c` | `error_msg12` | Account age does not meet requirements | `account_age_restriction` | Terminal attempt; account-specific | OFFICIAL FRONTEND VERIFIED |
| 40014 | `0x9c4e` | `error_msg6` | Gift code not found; capitalization may be wrong | `invalid_code` | Terminal; code-global | LIVE + OFFICIAL FRONTEND VERIFIED |
| 40015 | `0x9c4f` | `error_msg8` | Generic redemption-code error | `unknown_response` | No automatic retry; review | OFFICIAL FRONTEND VERIFIED; behavior unresolved |
| 40016 | `0x9c50` | `error_msg10` | Server busy; reward delivery may be deferred | `unknown_response` | No automatic retry; duplicate-delivery risk | OFFICIAL FRONTEND VERIFIED; behavior unresolved |
| 40017 | `0x9c51` | `error_msg11` | Account does not meet redemption requirements | `account_restriction` | Terminal attempt; account-specific | OFFICIAL FRONTEND VERIFIED |
| 40019 | `0x9c53` | `error_msg20` | Too many simultaneous actions | `simultaneous_action_throttle` | Existing bounded `temporary_error` retry state; transient | OFFICIAL FRONTEND VERIFIED |
| 40020 | `0x9c54` | `error_msg16` | Character/player information incorrect | `invalid_player` | Terminal; account-specific configuration issue | OFFICIAL FRONTEND VERIFIED |
| 40100 | `0x9ca4` | `error_msg13` | Verification refresh requested too frequently | `verification_throttle` | Existing bounded `rate_limited` retry state; transient | OFFICIAL FRONTEND VERIFIED |
| 40101 | `0x9ca5` | `frequency_tip2` | Exact semantic not established | `unknown_response` | Review; unresolved | UNRESOLVED |
| 40102 | `0x9ca6` | `error_msg15` | Verification code expired | `verification_error` | No automatic retry; request/session-specific review | OFFICIAL FRONTEND VERIFIED |
| 40103 | `0x9ca7` | `error_msg14` | Verification code incorrect | `verification_error` | No automatic retry; request/session-specific review | OFFICIAL FRONTEND VERIFIED |

## Profile differences

WOS alone contains `40018 -> error_msg18`, whose meaning remains unresolved. The clients also differ at `40020`: WOS maps it to `error_msg19` and has a live `USER INFO ERROR.` observation, while Kingshot maps it to the supplied `error_msg16` character-information text. Both safely classify it as `invalid_player`, but the evidence paths remain separate.

No numeric map entry references `error_msg9`, so the prerequisite text is not assigned to a Century number. The non-error keys `input_roleId_tips3`, `frequency_tip1`, `frequency_tip2`, `exchange_tips4`, and `exchange_tips5` remain recorded but fail closed because their exact response semantics were not established by the supplied localization catalogue.

`40016 -> error_msg10` is deliberately not retried. Its wording may indicate that Century accepted the redemption and deferred reward delivery; retrying could duplicate an accepted operation. `40015 -> error_msg8` is likewise too generic for safe automated behavior.
