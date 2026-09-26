# Telegram demo — `/demo` (allowlist only, no public endpoint)

> Isolated WSL2 execution of the top-3 proof tools via the existing Hermes
> Telegram gateway. **Allowlist only: `6729690821`. Never public.**
> No tokens, keys or secrets in this file, in logs or in the repo —
> secrets live only in `~/.hermes/.env` (600) on the WSL2 host.
> Default model stays free (`routine`); paid (`cheap`) only on explicit order.
> Gateway method: existing `tmux hermes` session running
> `hermes gateway run` (foreground, official WSL method). Do NOT restart
> the gateway for `/demo` setup; if a restart ever becomes necessary,
> backup first (see §6) and keep diagnostics to `last4`/`len` only.

## 1. What `/demo` does

`/demo <tool> <fixture-id>` runs one of the three proof tools in a WSL2
sandbox and replies **only to the allowlisted chat** with a short summary
plus the transcript file. No HTTP endpoint is exposed; Telegram polling
via the existing gateway is the only channel.

- `<tool>`: `elearning` | `funding` | `regulation`
- `<fixture-id>`: `fictional-4righe` (4-line fictional fixture, no PII)

Examples (copy-ready, Telegram chat with the bot):

```
/demo elearning fictional-4righe
/demo funding fictional-4righe
/demo regulation fictional-4righe
```

## 2. Sandbox procedure (every run)

1. Work dir: `/dev/shm/hermes-demo-XXXX` (tmpfs, removed after reply).
2. Copy the tool `.py` from the ext4 mirror (`~/projects/Portfolio/...`,
   never from `/mnt/*`) into the sandbox dir.
3. Run with the sandbox venv python, `timeout 10` per command, no secrets
   in env, no network for the function-level demo:
   - `python3 <tool>.py --help` (captured; elearning/regulation have no
     argparse `--help` and report their normal entry banner — expected)
   - offline fixture demo (pure functions, fictional data only):
     - elearning: `parse_durata` 4/4, threshold split, morning window,
       name normalization
     - funding: Italian number parsing, confidence order, EUR 800,000
       threshold (3/4 above), FNC3 + stale filters
     - regulation: category hits on p.2, free-text hits on p.3, page cited
4. Reply to the allowlisted chat only: summary (≤10 lines) + the
   corresponding `demo-transcript.txt` as a document.
5. Remove the sandbox dir. Never log secrets; transcripts contain only
   fictional data (see `projects/*/demo-fixture.txt`).

Reference outputs (static, same-origin, CSP `connect-src 'self'` safe):

- `projects/elearning-hours-monitor/demo-transcript.txt`
- `projects/funding-call-scraper/demo-transcript.txt`
- `projects/regulation-search/demo-transcript.txt`

## 3. Gateway + skills used

- Existing persistent gateway: `tmux new-session -d -s hermes 'hermes gateway run'`
  (already running; check with `hermes gateway status` — read-only).
- Hermes tools: `code_execution` (sandbox run) + `file` (read transcript,
  send document) + `systematic-debugging` skill for failures.
- No new public endpoint, no webhook, no port forwarding.

## 4. Phone usage (allowlisted user)

1. Open the bot chat → `/start` (first time) → `/sethome` in the work chat.
2. Send one of the three `/demo ...` lines from §1.
3. Expected: summary + transcript file within ~1 minute (free-model latency).
4. If no reply in 2 minutes: check PC + WSL2 + `hermes gateway status`
   (see workspace README §7.6); do NOT paste tokens anywhere.

## 5. CLI test (free model, no paid)

From WSL2 (free `routine` alias, single ping, no retry loop):

```bash
hermes chat --model routine -q "demo elearning fictional" --oneshot
```

Expected: short fictional demo summary (no secrets, no network scrape).
Paid `cheap` is NEVER used for `/demo` unless the user orders it explicitly.

## 6. Safety rules (binding)

- NEVER expose a public HTTP endpoint for demos.
- NEVER touch `~/.hermes/.env` values for this task (read-only checks:
  `grep -c`, `wc -c`, `last4`/`len` only — never full values).
- Before ANY gateway/config change: `cp` backup with `.bak-YYYYMMDD`
  suffix first; gateway restart only if strictly necessary for `/demo`
  AND after backup.
- Diagnostics: key/token lengths and last4 only; full values never appear
  in chat, logs, docs or the repo.
- No payments, no key revocations as part of this setup.
