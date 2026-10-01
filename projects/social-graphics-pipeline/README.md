# Apprendistato Studio — LinkedIn Content Utility

Small, local-first web app for drafting LinkedIn content about apprenticeship.
It turns a short brief into an editable Italian post, a prompt for an external
image-generation service, and a clean, downloadable 1080 × 1350 graphic.

## Run locally

Open `index.html` in a modern browser. No installation, build step, server,
network access, API key, or third-party dependency is required.

The app processes the brief in the browser. Copy and prompt are editable;
the graphic can be exported as PNG or SVG. The AI visual prompt is provided
as text for use with an image service of your choice—the app does not call an
AI service or generate raster imagery in this first version.

## Workflow

1. Add a topic, audience, key message, and any verified supporting details.
2. Generate and review the LinkedIn copy and visual prompt.
3. Edit the copy and export a branded graphic as PNG or SVG.
4. Optionally use the visual prompt with an image generator and compose that
   image with the exported graphic in a design tool.

Claims, figures, legal details, and course-specific information are supplied
by the author; review them before publishing. Avoid entering personal or
confidential data.

## Privacy and next steps

This static app has no backend and makes no external requests. Do not put
provider API keys into browser code or commit them to GitHub. A future AI
integration should use a local backend or another private secret store, with
an explicit provider adapter and clear disclosure of any external requests.
