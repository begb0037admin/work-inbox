# `approval_mode` deny-test — schema findings + BLOCKED on Codex usage cap (26 Aug 2026, Drew)

Authorised by Kevin via coordinator: back up `config.toml`, enumerate write tools, add `approval_mode` deny overrides for every write-capable Outlook/Calendar/Teams tool, re-verify TOML, re-run the write-gate test 3 ways, confirm reads survive.

## Status: PARTIALLY DONE — behavioural verification BLOCKED

| Step | State |
|---|---|
| 1. Back up `config.toml`, byte-verify | **DONE.** `C:\Users\admin\.codex\config.toml.bak-20260826_211513-drew-approvalmode-deny-test`, sha1 `78eeccf31930c47c82a299e79eef2ffec9c9d834`, `cmp` clean against the live file. |
| 2. Full write-tool enumeration | **BLOCKED.** Partial list only (`connector_write_tool_enumeration_partial.json`). The follow-up enumeration call failed: `ERROR: You've hit your usage limit ... try again at Aug 27th, 2026 1:20 AM.` |
| 3. Add deny overrides | **NOT DONE — deliberately.** Cannot safely apply an incomplete denylist (would give false protection) or an incomplete allowlist (would break reads I need) without the full tool inventory from step 2. `config.toml` is **UNCHANGED** this session (backup taken but not needed). |
| 4. Re-verify TOML parses | n/a yet — `codex doctor` confirmed the *current unchanged* config still parses OK and loads. |
| 5. Re-run write-gate test, verify 3 ways | **BLOCKED** by the same usage cap — the test itself needs `codex exec`. |
| 6. Confirm reads still work | **BLOCKED** — same. |
| 7. Revert any live-data side effect | n/a — no test ran, no side effect. Write-gate artifact from the 26 Aug test re-confirmed clean (`Categories == ''` via COM). |

## Schema findings (from `codex.exe` binary string inspection — offline, no API)

These narrow the approach for when Codex is available again:

- **`approval_mode` accepts NO `"deny"` / `"reject"` / `"never"` value.** The `AppToolApproval` enum's variants, read directly from the binary, are exactly: **`auto`, `prompt`, `writes`, `approve`**. (`approve` = auto-approve, the value the cc93c7b incident used for GitHub writes. `prompt` = always prompt. `writes` = prompt for writes only. `auto` = defer to session policy.) The `"never"` string in the binary belongs to a *different* enum — the session-level `approval_policy` (`untrusted` / `on-request` / `on-failure` / `never` / `granular`), where `never` means "never ask, just run" — the opposite of a denial.
- **The `[apps.<connector_id>]` config table supports these fields** (from the binary struct definition): `default_tools_approval_mode`, `enabled_tools`, `disabled_tools`, `tools` (per-tool sub-table with its own `approval_mode`), `scopes`, `oauth_resource`, `omit_tools_from`, `required`, `supports_parallel_tool_calls`, `startup_timeout_sec`, `tool_timeout_sec`, `auth`, `name`, `source`.
- **So there IS a plausible structural block, just not literally `"deny"`:**
  - `disabled_tools = ["<write tool names>"]` — remove the write tools from the app entirely (cleanest, but needs the complete write-tool name list from step 2).
  - `enabled_tools = ["<read tool names only>"]` — allowlist / default-deny (strongest, but needs the complete *read*-tool list or it breaks reads).
  - `default_tools_approval_mode = "prompt"` (+ per-write-tool `approval_mode = "prompt"`) — force a prompt on writes; in a headless `codex exec` with no TTY this *should* fail the call. **Unproven** — the 26 Aug write-gate test ran with the default and the write went through with no prompt, so the effective default for these connectors is `approve`/`auto`; explicitly setting `prompt` is the thing to test.

## Resume procedure (when Codex usage resets, ~01:20 AM 27 Aug 2026)

1. Confirm `config.toml` still matches the `.bak-20260826_211513-*` backup (nothing else changed it in the meantime).
2. Re-run the full write-tool enumeration (`brief_enum_full.txt` in the scratchpad, or re-issue: "list every codex_apps/microsoft_outlook_email|outlook_calendar|teams|planner tool, exact names, read vs write"). Save as `connector_write_tool_enumeration_full.json`.
3. Add to `config.toml`, for each of the 3 connector ids:
   ```toml
   [apps.connector_4aaab2856305417b993eca9a216aaf6e]      # Outlook Email
   default_tools_approval_mode = "prompt"
   disabled_tools = [ <every write tool from step 2> ]

   [apps.connector_e6a7394682e24467ac68c60696f275a4]      # Outlook Calendar
   default_tools_approval_mode = "prompt"
   disabled_tools = [ <every write tool> ]

   [apps.connector_246af0940da3457da0e751171dc1ce60]      # Microsoft Teams
   default_tools_approval_mode = "prompt"
   disabled_tools = [ <every write tool> ]
   ```
   (Prefer `disabled_tools` for a clean structural removal; keep `default_tools_approval_mode = "prompt"` as belt-and-braces for any write tool missed by the enumeration. Do NOT add `enabled_tools` unless the full read-tool list is confirmed — an incomplete allowlist breaks reads.)
4. `codex doctor` → confirm `config.toml parse ok`, `config loaded`, `MCP servers 3`, no new errors.
5. Re-run the write-gate test verbatim (`brief_writegate_test.txt` — categorize the Soundtrap marketing email under `codex exec -s read-only`). Verify THREE ways: (a) Codex transcript shows the call was refused/failed, (b) a second independent `codex exec -s read-only` `fetch_message` call shows no `Marketing` category, (c) Outlook COM on this machine shows `item.Categories == ''`.
6. Run a normal read pull (`brief_inbox.txt`) — confirm it still returns 40 messages, i.e. `disabled_tools`/`prompt` did not break reads.
7. If the write got through anyway (block ineffective), or reads broke: restore `config.toml` from the backup, byte-verify, and report that the local config mechanism cannot enforce denial — at which point the Oxford IT route (Task 2) becomes the only path.
8. If the block held and reads survived: keep it, and update the research doc Section 9 + HANDOVER with the confirmed result.

## Recommendation

The local `config.toml` route is **plausible** (`disabled_tools` + `default_tools_approval_mode = "prompt"` on all three connector ids) but **cannot be confirmed working this session** because Codex is capped and the test needs Codex. Do NOT treat the write path as closed until steps 5–6 above pass. The Oxford IT request (Task 2) should go out in parallel regardless — it is the durable structural fix and does not depend on whether the local block works.
