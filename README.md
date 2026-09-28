# Andrea Barduani — Digital Marketing, Communication & Automation Portfolio

[![CI](https://github.com/andreabarduani-ui/Portfolio/actions/workflows/ci.yml/badge.svg)](https://github.com/andreabarduani-ui/Portfolio/actions/workflows/ci.yml)

Personal portfolio website focused on digital marketing and communication,
supported by market research, data storytelling, web concepts and 13 Python
automation tools covering the full lifecycle of funded training programs
(launch → delivery → reporting).

**Live site:** https://andreabarduani-ui.github.io/Portfolio/

## Tech and structure

Static site — `index.html` plus the canonical external assets
`css/site.css` and `js/site.js` (vanilla JavaScript, zero npm dependencies).
The homepage intentionally keeps content in HTML for accessibility and
progressive enhancement; styles and behavior live in the external files so
there is one place to maintain each layer. Hosted for free on GitHub Pages.

The portfolio presents **19 projects across four sections, including the AI
assistant**. Three selected case studies at the top of the projects section
show the challenge, approach and deliverable before the full project list.

## Develop locally

Just open `index.html` in a browser — no build step required.

## Deploy

The site is published automatically via GitHub Pages from the `main` branch.
The same static root is also Netlify-compatible through [netlify.toml](./netlify.toml);
there is no build step and no dependency installation for the frontend.

```bash
git add index.html
git commit -m "Update portfolio"
git push
```

## License

All rights reserved — this is a personal portfolio.
