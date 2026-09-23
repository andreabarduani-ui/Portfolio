# Deploy gratis - Portfolio statico

> Stato 2026-09-23: repo `andreabarduani-ui/Portfolio`, branch `main` pulito su
> `de04d99`. Sito statico: `index.html` (65 KB) + `projects/csr-energy-dashboard/`
> + `files/*.pdf` (~14 MB). Totale repo ~16,6 MB / 52 file (prima dei file frontend).
> Check pre-deploy: 2/2 HTML con `<title>`+`meta description` OK -
> 4/4 JSON validi OK - 5/5 link interni esistenti OK.
> `netlify.toml` con `publish="."` OK (nessuna build, nessun secret).
>
> NOTA CONFLITTO 14:41 CEST: il frontend ha creato in parallelo `css/site.css`,
> `js/site.js`, `dashboard/` e le sue versioni di `netlify.toml`, `_redirects`,
> `robots.txt`, `sitemap.xml` (sovrascrivendo le mie bozze delle 14:40:45).
> Per regola non ho sovrascritto: valgono i file del frontend.
> Review dei suoi file sotto in "Review file frontend (read-only)".
> Questo `DEPLOY_NETLIFY.md` e' l'unico file mio rimasto ed e' intatto.

## Opzione A - Netlify drag & drop / dashboard (CONSIGLIATA, piu' facile)

Pro: zero CLI, 5 click, preview automatiche, HTTPS + CDN globale, rollback 1-click.
Contro: deploy manuale (zip) se non colleghi Git; con Git collegato diventa automatico.

1. Vai su <https://app.netlify.com> -> **Add new site -> Deploy manually**
   (oppure **Import from Git** -> GitHub -> `andreabarduani-ui/Portfolio` -> branch `main`,
   build command vuota, publish directory `.`).
2. Trascina uno zip della cartella (solo file sito, vedi nota peso sotto).
3. **Site settings -> Change site name** -> es. `andreabarduani-portfolio`.
4. Verifica `https://TUO-SITO.netlify.app/` + `/dashboard/` + `/projects/csr-energy-dashboard/`.
5. Aggiorna `robots.txt` + `sitemap.xml` con il dominio reale e ri-deploya.

## Opzione B - Netlify CLI (senza interattivita' pesante)

Pro: deploy dal terminale, `--prod` in un comando, draft preview con URL.
Contro: richiede `npm` + login via browser una sola volta (`netlify login` apre il browser).

```powershell
npm i -g netlify-cli --prefix D:\Workspace\tools\npm-global
$env:PATH = "D:\Workspace\tools\npm-global;$env:PATH"
netlify login            # una sola volta, via browser
netlify deploy --dir=. --prod   # dal root D:\Workspace\my-projects\Portfolio
```

Draft (anteprima non produttiva): `netlify deploy --dir=.`

## Opzione C - Alternative gratis: Cloudflare Pages / Vercel

| | Netlify | Cloudflare Pages | Vercel |
|---|---|---|---|
| Deploy statico da Git | OK | OK | OK |
| Banda free indicativa | ~15 GB/mese equivalenti (300 crediti, 20 cred./GB) | Ampia su piano free, ottima per PDF pesanti | 100 GB/mese |
| Form contatti free | OK Netlify Forms illimitati (2026) | NO (serve worker esterno) | NO |
| Analytics free | OK base (lookback 1 giorno su free) | OK Web Analytics free, no cookie | OK Vercel Analytics (eventi limitati) |
| Ideale se | Vuoi form gratis integrato | Vuoi banda massima per i PDF | Usi gia' Vercel/Next.js |

## Limiti free tier Netlify 2026 (piano a crediti - verificare su netlify.com/pricing al deploy)

- **300 crediti/mese, hard limit**: a esaurimento il sito va in pausa fino al mese dopo.
- **Production deploy = 15 crediti** cad. -> ~20 deploy prod/mese se fai solo quello.
- **Banda = 20 crediti/GB** -> ~15 GB/mese se usi solo banda
  (il sito pesa ~16,6 MB ma i PDF da ~7 MB l'uno consumano banda a ogni download -
  con ~1000 visite complete/mese resti ampiamente dentro).
- **Web requests = 2 crediti/10k richieste**.
- **Form submissions = gratis e illimitati** (da aprile 2026).
- 1 build concorrente, 500 progetti/team, preview deploy illimitate.

## Review file frontend (read-only, ore 14:41 - NON modificati)

- `netlify.toml`: `[build] publish="."`, `command=""` OK (requisito verificato).
  Redirect `/dashboard -> /dashboard/` 301 force=true: normalizza trailing slash,
  non oscura file reali. Header JSON (`Content-Type application/json`, cache 1h),
  CORS `*` su `/dashboard/data/*`, CSP restrittiva
  (self + cdn.jsdelivr.net + Google Fonts), `X-Frame-Options DENY`. OK.
- `_redirects`: mirror della regola TOML, coerente. OK.
- `robots.txt` + `sitemap.xml`: dominio `https://andreabarduani-ui.netlify.app/`.
  GAP: `sitemap.xml` elenca `/` e `/dashboard/` ma NON
  `/projects/csr-energy-dashboard/` (pushato, live, linkato da index.html).
  Suggerimento al frontend: aggiungere la terza URL + i PDF pubblici.
- Diff vs mia bozza (ritirata, non ripristinare senza accordo):
  io avevo alias `/csr-energy-dashboard` e `/dashboard` -> `/projects/.../`
  piu' cache lunga su `/files/*`. L'alias corto e' perso; se serve, concordarlo
  col frontend (conflitto potenziale con la nuova `dashboard/` reale).

## Form contatti gratis

- **Netlify Forms (consigliato qui)**: aggiungi `netlify` al `<form>` + campo
  `name="contact"`, ricevi submission da dashboard, spam filter incluso, gratis.
  ```html
  <form name="contact" method="POST" data-netlify="true">
    <input type="hidden" name="form-name" value="contact" />
    ...
  </form>
  ```
- **Formspree (alternativa)**: nessun vincolo di hosting, piano free ~50 invii/mese,
  basta `action="https://formspree.io/f/TUO-ID"`. Utile se resti su GitHub Pages.

## Analytics gratis (no-cookie consigliati in EU)

- **Cloudflare Web Analytics**: gratis, privacy-friendly, script leggero - ok su qualsiasi host.
- **Plausible/Umami self-host o trial**: Plausible cloud a pagamento, ma trial 30gg;
  Umami Cloud free per 3 siti / 6k eventi/mese (verificare soglie attuali).
- **Google Analytics 4**: gratis senza limiti pratici, ma cookie banner + GDPR
  (serve consent mode + informativa). Solo se servono funnel avanzati.

## Nota peso (importante per drag & drop)

Totale ~16,6 MB + nuovi file frontend (css ~21 KB, js ~25 KB, dashboard ~25 KB + dati),
di cui ~14,6 MB nei 2 PDF. Lo zip per drag & drop va bene.
Se in futuro i file crescono >100 MB, meglio Git-connected deploy o Cloudflare Pages.

## Push proposto (NON eseguito - serve ordine esplicito)

```powershell
git checkout -b site/netlify
git add netlify.toml _redirects robots.txt sitemap.xml DEPLOY_NETLIFY.md css js dashboard
git commit -m "Add Netlify deploy config (static publish .) + dashboard and site assets"
git push -u origin site/netlify
gh pr create --fill --base main
```

Nessun push eseguito. Nessun file frontend toccato.

## Stima tempi

- **GitHub Pages** (stato attuale, live): re-deploy ~1-3 min a push su `main`.
- **Netlify da Git**: primo deploy ~1-2 min (statico, nessuna build) + DNS/SSL ~1 min;
  deploy successivi ~30-60 s. Drag & drop: online in <1 min.
