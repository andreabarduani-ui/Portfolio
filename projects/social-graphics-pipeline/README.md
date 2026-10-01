# Apprendistato Studio — LinkedIn Content Utility

Local-first web app for drafting LinkedIn content about apprenticeship and
chatting with a remotely hosted Hermes Agent. It turns a short brief into an
editable Italian post, a prompt for an external image-generation service, and
a clean, downloadable 1080 × 1350 graphic.

## Run the local utility

Open `index.html` in a modern browser for local copy, visual prompt, preview,
and PNG/SVG export. These features do not require a server, network access,
API key, or third-party dependency.

The optional chat requires Node.js 20.6 or newer. It uses only Node's built-in
modules and serves the same static app through a small local relay:

```powershell
Copy-Item .env.example .env
# Edit .env and set the Hermes API URL and key.
npm start
```

Then open <http://127.0.0.1:4173>. The relay binds to loopback only and does
not expose an API key to the browser. It forwards chat requests to Hermes's
OpenAI-compatible `POST /v1/chat/completions` endpoint. The URL must include
the `/v1` suffix, for example `http://<hermes-host>:8642/v1`.

## Connect the remote Hermes Agent

1. On the PC running Hermes, enable its API server and set a strong API key
   in the Hermes environment. For a Docker deployment, the relevant settings
   are typically:

   ```dotenv
   API_SERVER_ENABLED=true
   API_SERVER_HOST=0.0.0.0
   API_SERVER_PORT=8642
   API_SERVER_KEY=<long-random-secret>
   ```

   Publish the container port only on a trusted private interface and allow
   it through the host firewall only from the app PC/VPN. The container must
   be configured to accept the request on port `8642`; a port mapping alone
   does not change Hermes's bind address.
2. Keep the two PCs on a trusted private network or VPN (for example
   Tailscale). Do not expose the Hermes port or its bearer key to the public
   internet; use HTTPS when traffic crosses an untrusted network. The Hermes
   API can invoke the agent's configured tools, so treat it as a privileged
   endpoint and restrict both the network path and API key.
3. Copy `.env.example` to `.env` on the app PC and set:
   - `HERMES_API_URL` to the reachable Hermes URL ending in `/v1`.
   - `HERMES_API_KEY` to the value of Hermes's `API_SERVER_KEY`.
4. Run `npm start`, open the local page, and use **Verifica connessione**.
   The status check only requests the authenticated `/v1/models` endpoint;
   chat messages are sent only when you explicitly submit them.

The Telegram bot token and any model-provider API keys already configured on
the Hermes PC are separate credentials and are not used by this web chat.
This interface connects to Hermes's HTTP API server, not to the Telegram bot;
do not put a Telegram token or provider key in the app's `.env`. Keep the
Telegram integration working independently as it does now.

The provided example URL is a placeholder. The actual hostname, port, Docker
port mapping, firewall rules, and Hermes API key depend on the other PC's
network setup. Never commit `.env` or put the key into frontend code. For
current Hermes API server options, see the
[official API server guide](https://github.com/NousResearch/hermes-agent/blob/main/website/docs/user-guide/features/api-server.md).
A GitHub Pages/static deployment can still use the local generation features,
but cannot use this private relay unless a separately secured backend is
deployed alongside it.

## Workflow

1. Add a topic, audience, key message, and any verified supporting details.
2. Generate and review the LinkedIn copy and visual prompt.
3. Edit the copy and export a branded graphic as PNG or SVG.
4. Optionally insert the brief in the Hermes chat for a second opinion on
   front-end, UX, or digital communication; it is not sent until submitted.
5. Optionally use the visual prompt with an image generator and compose that
   image with the exported graphic in a design tool.

Claims, figures, legal details, and course-specific information are supplied
by the author; review them before publishing. Avoid entering personal or
confidential data in the app or the Hermes conversation. Chat history is held
in the current browser tab only and is not saved by this app; Hermes or its
configured provider may apply separate session and data-retention policies.

## Privacy and next steps

The copy generator and graphic exports run in the browser. The optional local
relay stores no conversations and sends no prompt to Hermes until the user
sends a chat message or explicitly checks the connection. The relay keeps the
Hermes URL and key in its server-side environment and only binds to
`127.0.0.1`; it accepts same-origin local requests, limits request sizes, and
does not expose arbitrary proxy routes. The included Node tests can be run
with `npm test`.

The visual prompt is still provided as text for an image service of your
choice; this app does not call an image-generation service or generate
raster imagery.
