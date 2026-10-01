const form = document.querySelector("#content-form");
const copyOutput = document.querySelector("#copy-output");
const promptOutput = document.querySelector("#prompt-output");
const copyButton = document.querySelector("#copy-button");
const promptCopyButton = document.querySelector("#prompt-copy-button");
const downloadPngButton = document.querySelector("#download-png");
const downloadSvgButton = document.querySelector("#download-svg");
const graphicPreview = document.querySelector("#graphic-preview");
const statusMessage = document.querySelector("#status-message");

const graphic = {
  title: "La tua prossima idea prende forma.",
  message: "Il tuo messaggio chiave apparirà qui.",
};

function readBrief() {
  const data = new FormData(form);
  return {
    topic: String(data.get("topic") || "").trim(),
    audience: String(data.get("audience") || "").trim(),
    tone: String(data.get("tone") || "").trim(),
    message: String(data.get("message") || "").trim(),
    proof: String(data.get("proof") || "").trim(),
    cta: String(data.get("cta") || "").trim(),
    visualStyle: String(data.get("visualStyle") || "").trim(),
  };
}

function buildCopy(brief) {
  const hooks = {
    "chiaro e professionale": "Le competenze crescono quando la formazione incontra l’esperienza.",
    "diretto e concreto": "Apprendere facendo: da qui può partire un percorso professionale.",
    "caldo e ispirazionale": "Ogni nuova competenza comincia con un’occasione per mettersi in gioco.",
  };
  const hook = hooks[brief.tone] || hooks["chiaro e professionale"];
  const paragraphs = [
    hook,
    `Oggi parliamo di ${brief.topic}: un tema di interesse per ${brief.audience}.`,
    brief.message,
  ];

  if (brief.proof) {
    paragraphs.push(brief.proof);
  }

  if (brief.cta) {
    paragraphs.push(brief.cta);
  }

  paragraphs.push("#Apprendistato #Formazione #Competenze #Lavoro");
  return paragraphs.join("\n\n");
}

function buildVisualPrompt(brief) {
  return [
    `Crea un visual verticale 4:5 (1080 × 1350 px) per un post LinkedIn italiano sul tema "${brief.topic}".`,
    `Pubblico: ${brief.audience}. Stile: ${brief.visualStyle}, tono ${brief.tone}.`,
    `Idea da evocare: ${brief.message}`,
    "Composizione editoriale pulita, gerarchia visiva forte, spazio negativo generoso, palette petrolio e verde salvia con piccoli accenti caldi, luce naturale, aspetto professionale e umano.",
    "Non inserire parole, lettere, loghi, marchi d'acqua, uniformi con loghi o elementi che sembrino documenti ufficiali. Lascia una zona libera per aggiungere in seguito titolo e marchio con un editor grafico.",
  ].join("\n\n");
}

function escapeXml(value) {
  return value.replace(/[<>&'"]/g, (character) => {
    const entities = {
      "<": "&lt;",
      ">": "&gt;",
      "&": "&amp;",
      "'": "&apos;",
      '"': "&quot;",
    };
    return entities[character];
  });
}

function wrapText(text, maxChars, maxLines) {
  const words = text.split(/\s+/).filter(Boolean);
  const lines = [];
  let line = "";

  for (const word of words) {
    if (word.length > maxChars) {
      if (line) {
        lines.push(line);
        line = "";
      }
      for (let index = 0; index < word.length; index += maxChars) {
        const part = word.slice(index, index + maxChars);
        if (index + maxChars < word.length) {
          lines.push(part);
        } else {
          line = part;
        }
      }
      continue;
    }
    const next = line ? `${line} ${word}` : word;
    if (next.length <= maxChars) {
      line = next;
    } else {
      lines.push(line);
      line = word;
    }
  }
  if (line) {
    lines.push(line);
  }

  if (lines.length > maxLines) {
    lines.length = maxLines;
    const last = lines[maxLines - 1];
    lines[maxLines - 1] = `${last.replace(/[.,;:!?…]*$/, "")}…`;
  }
  return lines;
}

function createGraphicSvg(title, message) {
  const titleLines = wrapText(title, 17, 4);
  const messageLines = wrapText(message, 42, 4);
  const titleMarkup = titleLines
    .map((line, index) => `<tspan x="112" dy="${index === 0 ? 0 : 76}">${escapeXml(line)}</tspan>`)
    .join("");
  const messageMarkup = messageLines
    .map((line, index) => `<tspan x="116" dy="${index === 0 ? 0 : 42}">${escapeXml(line)}</tspan>`)
    .join("");
  const messageY = 805 + Math.max(0, 3 - titleLines.length) * 76;

  return `<svg xmlns="http://www.w3.org/2000/svg" width="1080" height="1350" viewBox="0 0 1080 1350">
  <defs>
    <linearGradient id="background" x1="0" y1="0" x2="1" y2="1">
      <stop offset="0" stop-color="#123c4a"/>
      <stop offset=".58" stop-color="#126e6a"/>
      <stop offset="1" stop-color="#6ba89a"/>
    </linearGradient>
    <linearGradient id="glow" x1="0" y1="0" x2="1" y2="1">
      <stop offset="0" stop-color="#e7bd91" stop-opacity=".22"/>
      <stop offset="1" stop-color="#e7bd91" stop-opacity="0"/>
    </linearGradient>
  </defs>
  <rect width="1080" height="1350" fill="url(#background)"/>
  <circle cx="1000" cy="120" r="295" fill="url(#glow)"/>
  <circle cx="1000" cy="120" r="230" fill="none" stroke="#ffffff" stroke-opacity=".22" stroke-width="2"/>
  <circle cx="1000" cy="120" r="190" fill="none" stroke="#e7bd91" stroke-opacity=".55" stroke-width="2"/>
  <path d="M0 1110 C260 1000 480 1300 1080 1080 V1350 H0Z" fill="#0c4e50" fill-opacity=".24"/>
  <text x="116" y="174" fill="#e7bd91" font-family="Arial, sans-serif" font-size="23" font-weight="700" letter-spacing="4">APPRENDISTATO · IDEE IN PRATICA</text>
  <text x="108" y="440" fill="#ffffff" font-family="Georgia, 'Times New Roman', serif" font-size="83" font-weight="400" letter-spacing="-2">${titleMarkup}</text>
  <rect x="116" y="702" width="126" height="8" rx="4" fill="#e7bd91"/>
  <text x="116" y="${messageY}" fill="#eff7f4" font-family="Arial, sans-serif" font-size="35" font-weight="400">${messageMarkup}</text>
  <text x="116" y="1244" fill="#ffffff" fill-opacity=".78" font-family="Arial, sans-serif" font-size="20" font-weight="700" letter-spacing="3">FORMAZIONE · COMPETENZE · FUTURO</text>
  <text x="964" y="1244" fill="#e7bd91" font-family="Georgia, serif" font-size="27" font-style="italic">A.</text>
</svg>`;
}

function updateGraphic(brief) {
  graphic.title = brief.topic;
  graphic.message = brief.message;
  graphicPreview.querySelector(".graphic-title").textContent = brief.topic;
  graphicPreview.querySelector(".graphic-message").textContent = brief.message;
  downloadPngButton.disabled = false;
  downloadSvgButton.disabled = false;
}

function makeFilename(topic, extension) {
  const safeTopic = topic
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, "-")
    .replace(/^-|-$/g, "")
    .slice(0, 48);
  return `linkedin-apprendistato-${safeTopic || "post"}.${extension}`;
}

function downloadBlob(blob, filename) {
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = filename;
  link.click();
  window.setTimeout(() => URL.revokeObjectURL(url), 1000);
}

async function copyText(text, label) {
  try {
    await navigator.clipboard.writeText(text);
    statusMessage.textContent = `${label} copiato negli appunti.`;
  } catch {
    statusMessage.textContent = "Il browser non consente la copia automatica: seleziona il testo e copialo manualmente.";
  }
}

form.addEventListener("submit", (event) => {
  event.preventDefault();
  if (!form.reportValidity()) {
    return;
  }
  const brief = readBrief();
  copyOutput.value = buildCopy(brief);
  promptOutput.value = buildVisualPrompt(brief);
  updateGraphic(brief);
  copyButton.disabled = false;
  promptCopyButton.disabled = false;
  statusMessage.textContent = "Bozza, prompt e anteprima aggiornati. Rivedi il copy prima di usarlo.";
});

copyButton.addEventListener("click", () => copyText(copyOutput.value, "Copy"));
promptCopyButton.addEventListener("click", () => copyText(promptOutput.value, "Prompt"));

downloadSvgButton.addEventListener("click", () => {
  const svg = createGraphicSvg(graphic.title, graphic.message);
  downloadBlob(new Blob([svg], { type: "image/svg+xml;charset=utf-8" }), makeFilename(graphic.title, "svg"));
  statusMessage.textContent = "Grafica SVG scaricata.";
});

downloadPngButton.addEventListener("click", async () => {
  const svg = createGraphicSvg(graphic.title, graphic.message);
  const svgBlob = new Blob([svg], { type: "image/svg+xml;charset=utf-8" });
  const svgUrl = URL.createObjectURL(svgBlob);
  try {
    const image = new Image();
    image.src = svgUrl;
    await image.decode();
    const canvas = document.createElement("canvas");
    canvas.width = 1080;
    canvas.height = 1350;
    const context = canvas.getContext("2d");
    if (!context) {
      throw new Error("Canvas rendering is unavailable.");
    }
    context.drawImage(image, 0, 0);
    const pngBlob = await new Promise((resolve, reject) => {
      canvas.toBlob((blob) => {
        if (blob) {
          resolve(blob);
        } else {
          reject(new Error("PNG export failed."));
        }
      }, "image/png");
    });
    downloadBlob(pngBlob, makeFilename(graphic.title, "png"));
    statusMessage.textContent = "Grafica PNG scaricata.";
  } catch (error) {
    statusMessage.textContent = `Non è stato possibile esportare il PNG: ${error.message}`;
  } finally {
    URL.revokeObjectURL(svgUrl);
  }
});
