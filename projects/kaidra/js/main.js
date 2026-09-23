/* ============================================================
   Kaidra · Concept Redesign — interactions
   Vanilla JS, no dependencies. All effects honor
   prefers-reduced-motion and pause when off-screen.
   ============================================================ */
(function () {
  "use strict";

  const $ = (sel, ctx) => (ctx || document).querySelector(sel);
  const $$ = (sel, ctx) => Array.from((ctx || document).querySelectorAll(sel));
  const prefersReduced = window.matchMedia("(prefers-reduced-motion: reduce)").matches;

  /* ---------- Footer year ---------- */
  const yearEl = $("#year");
  if (yearEl) yearEl.textContent = String(new Date().getFullYear());

  /* ---------- Header: scroll state ---------- */
  const header = $("#siteHeader");
  const onScroll = () => {
    if (header) header.classList.toggle("is-scrolled", window.scrollY > 24);
  };
  window.addEventListener("scroll", onScroll, { passive: true });
  onScroll();

  /* ---------- Mobile nav ---------- */
  const navToggle = $("#navToggle");
  const siteNav = $("#siteNav");
  if (navToggle && siteNav) {
    navToggle.addEventListener("click", () => {
      const open = siteNav.classList.toggle("is-open");
      navToggle.setAttribute("aria-expanded", String(open));
      navToggle.setAttribute("aria-label", open ? "Close menu" : "Open menu");
    });
    siteNav.addEventListener("click", (e) => {
      if (e.target.tagName === "A") {
        siteNav.classList.remove("is-open");
        navToggle.setAttribute("aria-expanded", "false");
      }
    });
  }

  /* ---------- Active nav link ---------- */
  const navLinks = $$(".site-nav ul a");
  const linkById = new Map();
  navLinks.forEach((a) => linkById.set(a.getAttribute("href").slice(1), a));
  if ("IntersectionObserver" in window && linkById.size) {
    const sectionObserver = new IntersectionObserver(
      (entries) => {
        entries.forEach((entry) => {
          const link = linkById.get(entry.target.id);
          if (!link) return;
          if (entry.isIntersecting) {
            navLinks.forEach((a) => a.removeAttribute("aria-current"));
            link.setAttribute("aria-current", "true");
          }
        });
      },
      { rootMargin: "-40% 0px -55% 0px" }
    );
    linkById.forEach((_, id) => {
      const section = document.getElementById(id);
      if (section) sectionObserver.observe(section);
    });
  }

  /* ---------- Reveal on scroll ---------- */
  const revealTargets = $$(".section-head, .process, .split > *, .carousel, .lifecycle, .trace, .connector-tools, .connector-grid, .compare, .ai-grid, .claude-banner, .guide-grid, .cta-inner, .ledger");
  if (!prefersReduced && "IntersectionObserver" in window) {
    revealTargets.forEach((el) => el.classList.add("reveal"));
    const revealObserver = new IntersectionObserver(
      (entries, obs) => {
        entries.forEach((entry) => {
          if (entry.isIntersecting) {
            entry.target.classList.add("is-visible");
            obs.unobserve(entry.target);
          }
        });
      },
      { threshold: 0.12 }
    );
    revealTargets.forEach((el) => revealObserver.observe(el));
    /* Safety net: if IntersectionObserver never fires (exotic browsers,
       embedded webviews), reveal everything after a short delay. */
    setTimeout(() => {
      revealTargets.forEach((el) => el.classList.add("is-visible"));
      armHero();
    }, 2500);
  }

  /* ---------- Animated counters ---------- */
  const easeOutCubic = (t) => 1 - Math.pow(1 - t, 3);
  function animateCount(el) {
    const target = parseInt(el.dataset.countTo, 10) || 0;
    const grouped = el.dataset.format === "grouped";
    const fmt = (v) => (grouped ? v.toLocaleString("en-US") : String(v));
    if (prefersReduced) {
      el.textContent = fmt(target);
      return;
    }
    const duration = 1400;
    let start = null;
    let done = false;
    /* Progress is computed against the rAF timestamp, with start taken from
       the first frame so both come from the same clock. Clamped so a clock
       jump can never render a value outside [0, target]. */
    function frame(now) {
      if (done) return;
      if (start === null) start = now;
      const p = Math.min(Math.max((now - start) / duration, 0), 1);
      el.textContent = fmt(Math.round(target * easeOutCubic(p)));
      if (p >= 1) { done = true; return; }
      requestAnimationFrame(frame);
    }
    /* Safety net: if frames stop firing (heavy throttling, embedded
       webviews), the final value is still written. */
    setTimeout(() => {
      if (!done) { done = true; el.textContent = fmt(target); }
    }, duration + 400);
    requestAnimationFrame(frame);
  }

  /* ---------- Fill bars + counters when visible ---------- */
  function armFills(scope) {
    $$(".src-fill, .campaign-fill", scope).forEach((bar) => {
      bar.style.width = (bar.dataset.w || 0) + "%";
    });
  }
  const heroPanel = $(".hero-panel");
  let heroArmed = false;
  function armHero() {
    if (heroArmed || !heroPanel) return;
    heroArmed = true;
    armFills(heroPanel);
    $$(".count", heroPanel).forEach(animateCount);
  }
  if (heroPanel && "IntersectionObserver" in window) {
    const heroObserver = new IntersectionObserver(
      (entries, obs) => {
        entries.forEach((entry) => {
          if (!entry.isIntersecting) return;
          armHero();
          obs.unobserve(entry.target);
        });
      },
      { threshold: 0.35 }
    );
    heroObserver.observe(heroPanel);
  } else if (heroPanel) {
    armHero();
  }

  /* Fills inside hidden panels: arm when the panel becomes visible */
  function watchPanels() {
    const panels = $$(".process-panel, .compare-panel");
    if (!("IntersectionObserver" in window)) {
      panels.forEach((p) => !p.hidden && armFills(p));
      return;
    }
    const panelObserver = new IntersectionObserver((entries, obs) => {
      entries.forEach((entry) => {
        if (entry.isIntersecting && !entry.target.hidden) {
          armFills(entry.target);
          obs.unobserve(entry.target);
        }
      });
    }, { threshold: 0.3 });
    panels.forEach((p) => panelObserver.observe(p));
  }
  watchPanels();

  /* ---------- Hero toast rotation ---------- */
  const toastText = $("#heroToastText");
  const toastMessages = [
    "Femke V. off-boarded 12 days ago. f.vosloo is still active in SAP.",
    "Toxic combination: PRD_APPROVE + PAYROLL_RW on a.janssens.",
    "Dormant account: e.vandenborne idle 214 days in SharePoint."
  ];
  if (toastText && !prefersReduced) {
    let toastIndex = 0;
    setInterval(() => {
      toastText.style.opacity = "0";
      setTimeout(() => {
        toastIndex = (toastIndex + 1) % toastMessages.length;
        toastText.textContent = toastMessages[toastIndex];
        toastText.style.opacity = "1";
      }, 200);
    }, 4200);
  }

  /* ---------- Hero canvas: identity graph ---------- */
  const canvas = $("#graphCanvas");
  if (canvas) {
    const ctx = canvas.getContext("2d");
    let nodes = [];
    let edges = [];
    let width = 0;
    let height = 0;
    let rafId = null;
    let running = false;
    const pointer = { x: -9999, y: -9999 };

    function buildGraph() {
      nodes = [];
      edges = [];
      const hub = { x: 0.5, y: 0.52 };
      const clusterCenters = [
        { x: 0.22, y: 0.3 }, { x: 0.5, y: 0.16 }, { x: 0.8, y: 0.28 },
        { x: 0.16, y: 0.72 }, { x: 0.84, y: 0.74 }, { x: 0.5, y: 0.88 }
      ];
      nodes.push({ x: hub.x, y: hub.y, r: 4.5, hub: true, ox: hub.x, oy: hub.y });
      clusterCenters.forEach((c, ci) => {
        const centerIndex = nodes.length;
        nodes.push({ x: c.x, y: c.y, r: 3, hub: false, ox: c.x, oy: c.y, cluster: ci });
        edges.push([0, centerIndex]);
        const satellites = 3 + (ci % 3);
        for (let s = 0; s < satellites; s++) {
          const angle = (s / satellites) * Math.PI * 2 + ci * 0.7;
          const dist = 0.055 + Math.random() * 0.035;
          const nx = c.x + Math.cos(angle) * dist;
          const ny = c.y + Math.sin(angle) * dist * 0.85;
          nodes.push({ x: nx, y: ny, r: 1.6 + Math.random() * 1.2, hub: false, ox: nx, oy: ny, cluster: ci });
          edges.push([centerIndex, nodes.length - 1]);
        }
      });
      // A few cross-links for organic feel
      for (let i = 0; i < 5; i++) {
        const a = 1 + Math.floor(Math.random() * (nodes.length - 1));
        const b = 1 + Math.floor(Math.random() * (nodes.length - 1));
        if (a !== b) edges.push([a, b]);
      }
    }

    function resize() {
      const rect = canvas.parentElement.getBoundingClientRect();
      width = rect.width;
      height = rect.height;
      const dpr = Math.min(window.devicePixelRatio || 1, 2);
      canvas.width = width * dpr;
      canvas.height = height * dpr;
      ctx.setTransform(dpr, 0, 0, dpr, 0, 0);
    }

    let tick = 0;
    function draw() {
      tick += 0.006;
      ctx.clearRect(0, 0, width, height);

      // Update node positions (gentle drift + pointer repulsion)
      nodes.forEach((n) => {
        const driftX = Math.sin(tick + n.ox * 12) * 0.006;
        const driftY = Math.cos(tick * 0.8 + n.oy * 10) * 0.006;
        let nx = n.ox + driftX;
        let ny = n.oy + driftY;
        const px = n.x * width - pointer.x;
        const py = n.y * height - pointer.y;
        const distSq = px * px + py * py;
        const radius = 130;
        if (distSq < radius * radius && distSq > 0.01) {
          const dist = Math.sqrt(distSq);
          const force = (1 - dist / radius) * 0.045;
          nx += (px / dist) * force;
          ny += (py / dist) * force;
        }
        n.x += (nx - n.x) * 0.08;
        n.y += (ny - n.y) * 0.08;
      });

      // Edges
      edges.forEach(([a, b]) => {
        const na = nodes[a];
        const nb = nodes[b];
        ctx.beginPath();
        ctx.moveTo(na.x * width, na.y * height);
        ctx.lineTo(nb.x * width, nb.y * height);
        ctx.strokeStyle = "rgba(105, 155, 210, 0.26)";
        ctx.lineWidth = 1;
        ctx.stroke();
      });

      // Nodes
      nodes.forEach((n) => {
        ctx.beginPath();
        ctx.arc(n.x * width, n.y * height, n.r, 0, Math.PI * 2);
        ctx.fillStyle = n.hub ? "rgba(5, 122, 198, 0.9)" : "rgba(105, 155, 210, 0.55)";
        ctx.fill();
        if (n.hub) {
          ctx.beginPath();
          ctx.arc(n.x * width, n.y * height, n.r + 3.5, 0, Math.PI * 2);
          ctx.strokeStyle = "rgba(252, 126, 79, 0.55)";
          ctx.lineWidth = 1;
          ctx.stroke();
        }
      });

      if (running) rafId = requestAnimationFrame(draw);
    }

    function start() {
      if (running) return;
      running = true;
      rafId = requestAnimationFrame(draw);
    }
    function stop() {
      running = false;
      if (rafId) cancelAnimationFrame(rafId);
    }

    buildGraph();
    resize();

    if (prefersReduced) {
      draw(); // single static frame
      stop();
    } else {
      start();
      canvas.parentElement.addEventListener("pointermove", (e) => {
        const rect = canvas.getBoundingClientRect();
        pointer.x = e.clientX - rect.left;
        pointer.y = e.clientY - rect.top;
      });
      canvas.parentElement.addEventListener("pointerleave", () => {
        pointer.x = -9999;
        pointer.y = -9999;
      });
      window.addEventListener("resize", () => {
        resize();
        if (prefersReduced) draw();
      });
      document.addEventListener("visibilitychange", () => {
        document.hidden ? stop() : start();
      });
      if ("IntersectionObserver" in window) {
        new IntersectionObserver((entries) => {
          entries.forEach((entry) => (entry.isIntersecting ? start() : stop()));
        }).observe(canvas);
      }
    }
  }

  /* ---------- Mini identity graph (SVG) ---------- */
  const mgEdges = $("#mgEdges");
  const mgNodes = $("#mgNodes");
  if (mgEdges && mgNodes) {
    const svgNS = "http://www.w3.org/2000/svg";
    const data = [
      { id: "mira", label: "Femke V.", x: 70, y: 90, type: "user" },
      { id: "jonas", label: "Jonas P.", x: 60, y: 200, type: "user" },
      { id: "admin", label: "svc-admin", x: 90, y: 280, type: "user" },
      { id: "eng", label: "Engineering", x: 190, y: 120, type: "group" },
      { id: "fin", label: "Finance", x: 180, y: 250, type: "group" },
      { id: "sap", label: "SAP", x: 320, y: 110, type: "app" },
      { id: "entra", label: "Entra ID", x: 330, y: 210, type: "app" },
      { id: "workday", label: "Workday", x: 300, y: 300, type: "app" }
    ];
    const links = [
      ["mira", "eng"], ["jonas", "eng"], ["jonas", "fin"],
      ["admin", "entra"], ["eng", "sap"], ["eng", "entra"],
      ["fin", "sap"], ["fin", "workday"], ["workday", "eng"]
    ];
    const byId = new Map(data.map((n) => [n.id, n]));

    const edgeEls = links.map(([a, b]) => {
      const line = document.createElementNS(svgNS, "line");
      line.setAttribute("x1", byId.get(a).x);
      line.setAttribute("y1", byId.get(a).y);
      line.setAttribute("x2", byId.get(b).x);
      line.setAttribute("y2", byId.get(b).y);
      line.setAttribute("class", "mg-edge");
      mgEdges.appendChild(line);
      return { line, a, b };
    });

    const nodeEls = data.map((n) => {
      const g = document.createElementNS(svgNS, "g");
      g.setAttribute("class", "mg-node");
      const circle = document.createElementNS(svgNS, "circle");
      circle.setAttribute("cx", n.x);
      circle.setAttribute("cy", n.y);
      circle.setAttribute("r", n.type === "app" ? 12 : n.type === "group" ? 9 : 7);
      const text = document.createElementNS(svgNS, "text");
      text.setAttribute("x", n.x);
      text.setAttribute("y", n.y + (n.type === "app" ? 28 : 22));
      text.setAttribute("text-anchor", "middle");
      text.textContent = n.label;
      g.appendChild(circle);
      g.appendChild(text);
      mgNodes.appendChild(g);
      return { g, node: n };
    });

    function highlight(id) {
      if (!id) {
        edgeEls.forEach((e) => e.line.classList.remove("is-lit"));
        nodeEls.forEach((n) => n.g.classList.remove("is-lit", "is-dim"));
        return;
      }
      const lit = new Set([id]);
      edgeEls.forEach((e) => {
        const on = e.a === id || e.b === id;
        e.line.classList.toggle("is-lit", on);
        if (on) { lit.add(e.a); lit.add(e.b); }
      });
      nodeEls.forEach((n) => {
        n.g.classList.toggle("is-lit", n.node.id === id);
        n.g.classList.toggle("is-dim", !lit.has(n.node.id));
      });
    }

    nodeEls.forEach(({ g, node }) => {
      g.addEventListener("pointerenter", () => highlight(node.id));
      g.addEventListener("pointerleave", () => highlight(null));
      g.addEventListener("focus", () => highlight(node.id));
      g.addEventListener("blur", () => highlight(null));
    });
  }

  /* ---------- Process tabs (how it works) ---------- */
  $$("[data-process]").forEach((process) => {
    const tabs = $$(".process-step", process);
    const panels = $$(".process-panel", process);
    tabs.forEach((tab) => {
      tab.addEventListener("click", () => {
        tabs.forEach((t) => {
          const active = t === tab;
          t.classList.toggle("is-active", active);
          t.setAttribute("aria-selected", String(active));
        });
        panels.forEach((panel) => {
          const active = panel.id === tab.getAttribute("aria-controls");
          panel.classList.toggle("is-active", active);
          panel.hidden = !active;
          if (active) armFills(panel);
        });
      });
    });
  });

  /* ---------- Carousel ---------- */
  $$("[data-carousel]").forEach((carousel) => {
    const stage = $(".carousel-stage", carousel);
    const quotes = $$(".quote", stage);
    const dotsWrap = $("#carouselDots");
    let index = 0;
    let timer = null;

    const dots = quotes.map((_, i) => {
      const dot = document.createElement("button");
      dot.className = "carousel-dot" + (i === 0 ? " is-active" : "");
      dot.setAttribute("role", "tab");
      dot.setAttribute("aria-label", "Quote " + (i + 1) + " of " + quotes.length);
      dot.setAttribute("aria-selected", String(i === 0));
      dot.addEventListener("click", () => goTo(i, true));
      dotsWrap.appendChild(dot);
      return dot;
    });

    function sizeStage() {
      let tallest = 0;
      quotes.forEach((q) => {
        q.style.position = "relative";
        const h = q.offsetHeight;
        q.style.position = "";
        tallest = Math.max(tallest, h);
      });
      stage.style.minHeight = tallest + "px";
    }

    function goTo(i, userInitiated) {
      index = (i + quotes.length) % quotes.length;
      quotes.forEach((q, qi) => {
        const active = qi === index;
        q.classList.toggle("is-active", active);
        q.setAttribute("aria-hidden", String(!active));
      });
      dots.forEach((d, di) => {
        d.classList.toggle("is-active", di === index);
        d.setAttribute("aria-selected", String(di === index));
      });
      if (userInitiated) restartAuto();
    }

    function next() { goTo(index + 1, false); }
    function restartAuto() {
      if (prefersReduced) return;
      clearInterval(timer);
      timer = setInterval(next, 6500);
    }

    $("[data-carousel-prev]", carousel).addEventListener("click", () => goTo(index - 1, true));
    $("[data-carousel-next]", carousel).addEventListener("click", () => goTo(index + 1, true));

    carousel.addEventListener("keydown", (e) => {
      if (e.key === "ArrowLeft") { e.preventDefault(); goTo(index - 1, true); }
      if (e.key === "ArrowRight") { e.preventDefault(); goTo(index + 1, true); }
    });
    carousel.addEventListener("pointerenter", () => clearInterval(timer));
    carousel.addEventListener("pointerleave", restartAuto);
    carousel.addEventListener("focusin", () => clearInterval(timer));
    carousel.addEventListener("focusout", restartAuto);
    document.addEventListener("visibilitychange", () => {
      document.hidden ? clearInterval(timer) : restartAuto();
    });

    window.addEventListener("resize", sizeStage);
    window.addEventListener("load", sizeStage);
    sizeStage();
    restartAuto();
  });

  /* ---------- Lifecycle ---------- */
  $$("[data-lifecycle]").forEach((scope) => {
    const stages = $$(".life-stage", scope);
    const dayEl = $("#lifeDay");
    const textEl = $("#lifeText");
    const content = [
      { day: "Day 0", text: "HR creates the identity. The graph links person, position and cost centre before the first account exists." },
      { day: "Month 4", text: "Projects come and go: group memberships, roles and direct grants accumulate across SAP, Entra and SharePoint. Each hop lands in the graph with its reason." },
      { day: "Month 11", text: "The Q4 campaign flags her SAP finance entitlement as toxic in combination with payment approval. Her reviewer sees the full path, not just a row." },
      { day: "Day 12 after exit", text: "Femke leaves the company. HR marks the exit, Kaidra revokes access at the source within the hour, and the audit trail records every decision." }
    ];
    stages.forEach((stage, i) => {
      stage.addEventListener("click", () => {
        stages.forEach((s, si) => {
          const active = si === i;
          s.classList.toggle("is-active", active);
          s.setAttribute("aria-selected", String(active));
        });
        dayEl.textContent = content[i].day;
        textEl.textContent = content[i].text;
      });
    });
  });

  /* ---------- Trace stepper ---------- */
  $$("[data-trace]").forEach((scope) => {
    const hops = $$(".trace-hop", scope);
    const stepEl = $(".trace-step", scope);
    const textEl = $(".trace-text", scope);
    const content = [
      "Femke V., product designer. Off-boarded 12 days ago, yet her account still resolves to live entitlements downstream.",
      "Membership in Engineering came from her team assignment in Workday. It grants broad application access, including systems she never used.",
      "SAP_FI_READ was never assigned to Femke directly. She inherited it through Engineering, which is why the spreadsheet never explained it.",
      "The entitlement itself: read access to production finance data in SAP. One revoke at the group level removes it for 214 identities at once."
    ];
    hops.forEach((hop, i) => {
      hop.addEventListener("click", () => {
        hops.forEach((h, hi) => {
          const active = hi <= i;
          h.classList.toggle("is-active", active);
          h.setAttribute("aria-pressed", String(hi === i));
        });
        stepEl.textContent = "Hop " + (i + 1) + " of " + hops.length;
        textEl.textContent = content[i];
      });
    });
  });

  /* ---------- Connector filter + search ---------- */
  const grid = $("#connectorGrid");
  if (grid) {
    const connectors = $$(".connector", grid);
    const searchInput = $("#connectorSearch");
    const emptyState = $("#connectorEmpty");
    let activeFilter = "all";
    let query = "";

    function apply() {
      let visible = 0;
      connectors.forEach((item) => {
        const cat = item.dataset.cat;
        const text = item.textContent.toLowerCase();
        const matchFilter = activeFilter === "all" || cat === activeFilter;
        const matchQuery = !query || text.includes(query);
        const show = matchFilter && matchQuery;
        item.classList.toggle("is-hidden", !show);
        if (show) visible++;
      });
      emptyState.hidden = visible > 0;
    }

    $$("#connectorFilters .chip").forEach((chip) => {
      chip.addEventListener("click", () => {
        $$("#connectorFilters .chip").forEach((c) => {
          const active = c === chip;
          c.classList.toggle("is-active", active);
          c.setAttribute("aria-pressed", String(active));
        });
        activeFilter = chip.dataset.filter;
        apply();
      });
    });

    searchInput.addEventListener("input", () => {
      query = searchInput.value.trim().toLowerCase();
      apply();
    });
  }

  /* ---------- Compare tabs (old way / Kaidra) ---------- */
  $$("[data-compare]").forEach((scope) => {
    const tabs = $$(".compare-tab", scope);
    const panels = $$(".compare-panel", scope);
    tabs.forEach((tab) => {
      tab.addEventListener("click", () => {
        tabs.forEach((t) => {
          const active = t === tab;
          t.classList.toggle("is-active", active);
          t.setAttribute("aria-selected", String(active));
        });
        panels.forEach((panel) => {
          const active = panel.id === tab.getAttribute("aria-controls");
          panel.classList.toggle("is-active", active);
          panel.hidden = !active;
          if (active) armFills(panel);
        });
      });
    });
  });

  /* ---------- Marquee: duplicate track for seamless loop ---------- */
  $$("[data-marquee]").forEach((marquee) => {
    const track = $(".marquee-track", marquee);
    if (!track) return;
    const clone = track.cloneNode(true);
    clone.setAttribute("aria-hidden", "true");
    const wrapper = document.createElement("div");
    wrapper.className = "marquee-track";
    track.parentNode.replaceChild(wrapper, track);
    wrapper.appendChild(track);
    wrapper.appendChild(clone);
  });

  /* ---------- Hero variant switcher (A/B) ---------- */
  const heroSection = $(".hero");
  const heroB = $("#heroB");
  if (heroSection && heroB) {
    let initial = "a";
    try {
      const param = new URLSearchParams(location.search).get("hero");
      if (param === "b" || param === "a") initial = param;
      else if (sessionStorage.getItem("heroVariant") === "b") initial = "b";
    } catch (e) { /* storage unavailable: default variant */ }
    const vtBtns = $$(".vt-btn");
    function setVariant(v, persist) {
      heroSection.dataset.variant = v;
      heroB.hidden = v !== "b";
      vtBtns.forEach((b) => {
        const on = b.dataset.setVariant === v;
        b.classList.toggle("is-active", on);
        b.setAttribute("aria-checked", String(on));
      });
      if (persist) {
        try { sessionStorage.setItem("heroVariant", v); } catch (e) { /* ignore */ }
        try {
          const url = new URL(location.href);
          url.searchParams.set("hero", v);
          history.replaceState(null, "", url);
        } catch (e) { /* sandboxed webview: URL param is a nice-to-have */ }
      }
    }
    vtBtns.forEach((b) => b.addEventListener("click", () => setVariant(b.dataset.setVariant, true)));
    setVariant(initial, false);
  }

  /* ---------- Live demo console ---------- */
  const rowsBody = $("#consoleRows");
  if (rowsBody) {
    const IDENTITIES = [
      { id: 1, name: "Femke Vosloo", account: "f.vosloo", dept: "Engineering", sources: ["Entra", "SAP"], ents: 214, risk: "high", reason: "Off-boarded 12 days ago, SAP access still active", path: ["f.vosloo · user", "Engineering · group", "SAP_FI_READ · role", "SAP production · entitlement"], status: "active", last: "12 d ago" },
      { id: 2, name: "Milan Ockers", account: "m.ockers", dept: "Finance", sources: ["Entra", "SAP", "Workday"], ents: 187, risk: "high", reason: "Toxic combination: PRD_APPROVE + PAYROLL_RW", path: ["m.ockers · user", "Finance controllers · group", "PRD_APPROVE · role", "SAP payments · entitlement"], status: "active", last: "2 h ago" },
      { id: 3, name: "Elke Vandenborne", account: "e.vandenborne", dept: "Operations", sources: ["Entra", "SharePoint"], ents: 96, risk: "medium", reason: "Dormant 214 days, 37 entitlements with no owner", path: ["e.vandenborne · user", "Ops legacy · group", "SP_EDIT_ALL · role", "SharePoint ops · entitlement"], status: "active", last: "214 d ago" },
      { id: 4, name: "Julie Rasson", account: "j.rasson", dept: "HR", sources: ["Workday", "Entra"], ents: 143, risk: "medium", reason: "Left HR in March, HR_ADMIN still inherited", path: ["j.rasson · user", "HR admins · group", "HR_ADMIN · role", "Workday HR · entitlement"], status: "review", last: "3 d ago" },
      { id: 5, name: "Timo Ghekiere", account: "t.ghekiere", dept: "Sales", sources: ["Entra", "Salesforce"], ents: 88, risk: "low", reason: "Direct grants without owner", path: ["t.ghekiere · user", "Sales EMEA · group", "CRM_OPP_RW · role", "Salesforce · entitlement"], status: "active", last: "1 d ago" },
      { id: 6, name: "Lotte Brams", account: "l.brams", dept: "Engineering", sources: ["Entra", "SAP", "Jira"], ents: 231, risk: "none", reason: "", path: ["l.brams · user", "Engineering · group", "DEV_JIRA_RW · role", "Jira · entitlement"], status: "active", last: "20 min ago" },
      { id: 7, name: "Yari Drost", account: "y.drost", dept: "IT", sources: ["Entra", "Okta", "CyberArk"], ents: 176, risk: "medium", reason: "Privileged vault access never reviewed", path: ["y.drost · user", "Platform ops · group", "VAULT_ADMIN · role", "CyberArk · entitlement"], status: "active", last: "4 h ago" },
      { id: 8, name: "Senne Aerts", account: "s.aerts", dept: "Finance", sources: ["Entra", "NetSuite"], ents: 64, risk: "none", reason: "", path: ["s.aerts · user", "Finance · group", "NS_READ · role", "NetSuite · entitlement"], status: "active", last: "3 h ago" },
      { id: 9, name: "Piet Vrancken", account: "p.vrancken", dept: "Operations", sources: ["Entra", "TOPDESK"], ents: 71, risk: "low", reason: "Group has no owner since 2024", path: ["p.vrancken · user", "Ops support · group", "TD_TICKET_RW · role", "TOPDESK · entitlement"], status: "active", last: "2 d ago" },
      { id: 10, name: "Mae Vermeir", account: "m.vermeir", dept: "Engineering", sources: ["Entra", "Jira"], ents: 52, risk: "none", reason: "", path: ["m.vermeir · user", "Engineering · group", "DEV_JIRA_RW · role", "Jira · entitlement"], status: "active", last: "35 min ago" },
      { id: 11, name: "Wout Segers", account: "w.segers", dept: "IT", sources: ["Entra", "ServiceNow"], ents: 118, risk: "none", reason: "", path: ["w.segers · user", "IT service · group", "SN_INC_RW · role", "ServiceNow · entitlement"], status: "active", last: "1 h ago" },
      { id: 12, name: "Roos Michiels", account: "r.michiels", dept: "Sales", sources: ["Entra", "Salesforce"], ents: 97, risk: "low", reason: "Seasonal contractor, access expires in 5 days", path: ["r.michiels · user", "Contractors · group", "CRM_LEAD_RW · role", "Salesforce · entitlement"], status: "active", last: "6 h ago" },
      { id: 13, name: "Jef Vandelanotte", account: "j.vandelanotte", dept: "Finance", sources: ["Entra", "SAP"], ents: 133, risk: "medium", reason: "Duplicate account j.vandelanotte2 pending merge", path: ["j.vandelanotte · user", "Finance · group", "SAP_FI_READ · role", "SAP production · entitlement"], status: "review", last: "1 d ago" },
      { id: 14, name: "Hanne Clijsters", account: "h.clijsters", dept: "HR", sources: ["Workday", "Entra", "SharePoint"], ents: 149, risk: "none", reason: "", path: ["h.clijsters · user", "HR · group", "SP_HR_RW · role", "SharePoint HR · entitlement"], status: "active", last: "45 min ago" }
    ];
    const state = IDENTITIES.map((p) => ({ ...p, inCampaign: false }));
    const panel = $("#idPanel");
    const emptyMsg = $("#consoleEmpty");
    const searchInput = $("#consoleSearch");
    const toast = $("#consoleToast");
    let riskFilter = "all";
    let query = "";
    let selectedId = null;
    let toastTimer = null;

    const RISK_LABEL = { high: "High", medium: "Medium", low: "Low", none: "None" };
    const STATUS_LABEL = { active: "Active", review: "Under review", revoked: "Revoked" };

    function initialsOf(name) {
      return name.split(/\s+/).map((w) => w[0]).slice(0, 2).join("").toUpperCase();
    }
    function showToast(msg) {
      toast.textContent = msg;
      toast.classList.add("is-visible");
      clearTimeout(toastTimer);
      toastTimer = setTimeout(() => toast.classList.remove("is-visible"), 2600);
    }
    function visibleRows() {
      const q = query.toLowerCase();
      return state.filter((p) => {
        const matchRisk = riskFilter === "all" || p.risk === riskFilter;
        const matchQuery = !q || (p.name + " " + p.account + " " + p.dept).toLowerCase().includes(q);
        return matchRisk && matchQuery;
      });
    }
    function renderRows() {
      const rows = visibleRows();
      rowsBody.innerHTML = "";
      rows.forEach((p) => {
        const tr = document.createElement("tr");
        tr.tabIndex = 0;
        tr.dataset.id = String(p.id);
        tr.setAttribute("aria-selected", String(p.id === selectedId));
        tr.setAttribute("aria-label", p.name + ", risk " + RISK_LABEL[p.risk] + ", press to inspect");
        tr.innerHTML =
          '<td class="id-cell-name">' + p.name + "<small>" + p.account + "</small></td>" +
          '<td class="id-sources">' + p.sources.join(", ") + "</td>" +
          '<td class="num">' + p.ents + "</td>" +
          '<td><span class="risk-pill risk-' + p.risk + '">' + RISK_LABEL[p.risk] + "</span></td>" +
          '<td><span class="status-chip status-' + p.status + '">' + STATUS_LABEL[p.status] + "</span></td>";
        tr.addEventListener("click", () => selectRow(p.id));
        tr.addEventListener("keydown", (e) => {
          if (e.key === "Enter" || e.key === " ") { e.preventDefault(); selectRow(p.id); }
        });
        rowsBody.appendChild(tr);
      });
      emptyMsg.hidden = rows.length > 0;
      if (selectedId !== null && !rows.some((p) => p.id === selectedId)) {
        selectedId = null;
        renderPanel();
      }
    }
    function renderPanel() {
      const p = state.find((x) => x.id === selectedId);
      if (!p) {
        panel.innerHTML = '<p class="id-panel-hint">Select a row to inspect why access exists, then revoke it. Everything runs locally.</p>';
        return;
      }
      const revoked = p.status === "revoked";
      panel.innerHTML =
        '<div class="idp-head">' +
        '<span class="idp-avatar">' + initialsOf(p.name) + "</span>" +
        '<div><p class="idp-name">' + p.name + "</p>" +
        '<p class="idp-dept">' + p.dept + " · " + STATUS_LABEL[p.status] + "</p></div></div>" +
        '<ul class="idp-meta">' +
        "<li><span class=\"k\">Account</span><span class=\"v\">" + p.account + "</span></li>" +
        "<li><span class=\"k\">Sources</span><span class=\"v\">" + p.sources.join(", ") + "</span></li>" +
        "<li><span class=\"k\">Entitlements</span><span class=\"v\">" + p.ents + "</span></li>" +
        "<li><span class=\"k\">Last active</span><span class=\"v\">" + p.last + "</span></li>" +
        "</ul>" +
        (p.reason ? '<p class="idp-reason">Why flagged: ' + p.reason + ".</p>" : "") +
        '<div class="idp-path"><p class="mock-caption">Access path</p><ol>' +
        p.path.map((hop) => "<li>" + hop + "</li>").join("") +
        "</ol></div>" +
        '<div class="idp-actions">' +
        '<button class="btn btn-danger btn-sm" data-action="revoke"' + (revoked ? " disabled" : "") + ">" +
        (revoked ? "Revoked" : "Revoke access") + "</button>" +
        '<button class="btn btn-quiet btn-sm" data-action="campaign" aria-pressed="' + String(p.inCampaign) + '">' +
        (p.inCampaign ? "In Q4 campaign" : "Add to campaign") + "</button>" +
        "</div>";
      panel.querySelectorAll("[data-action]").forEach((btn) => {
        btn.addEventListener("click", () => {
          if (btn.dataset.action === "revoke" && p.status !== "revoked") {
            p.status = "revoked";
            showToast("Access revoked for " + p.name + " · decision logged with reason");
            renderRows();
            renderPanel();
          } else if (btn.dataset.action === "campaign") {
            p.inCampaign = !p.inCampaign;
            showToast(p.inCampaign
              ? p.name + " added to the Q4 campaign"
              : p.name + " removed from the Q4 campaign");
            renderPanel();
          }
        });
      });
    }
    function selectRow(id) {
      selectedId = id;
      renderRows();
      renderPanel();
      const row = rowsBody.querySelector('tr[data-id="' + id + '"]');
      if (row && document.activeElement && document.activeElement.tagName === "TR") row.focus();
    }

    searchInput.addEventListener("input", () => {
      query = searchInput.value.trim();
      renderRows();
    });
    $$("#consoleFilters .chip").forEach((chip) => {
      chip.addEventListener("click", () => {
        $$("#consoleFilters .chip").forEach((c) => {
          const active = c === chip;
          c.classList.toggle("is-active", active);
          c.setAttribute("aria-pressed", String(active));
        });
        riskFilter = chip.dataset.risk;
        renderRows();
      });
    });
    document.addEventListener("keydown", (e) => {
      if (e.key === "Escape" && selectedId !== null) {
        const row = rowsBody.querySelector('tr[data-id="' + selectedId + '"]');
        selectedId = null;
        renderRows();
        renderPanel();
        if (row) row.focus();
      }
    });

    renderRows();
  }

  /* ---------- Hero insights: view switcher ---------- */
  const segBtns = $$(".seg-btn");
  if (segBtns.length) {
    const datasets = $$(".risk-dataset");
    const caption = $("#riskChartCaption");
    const legend = $("#riskChartLegend");
    const copy = {
      overview: { caption: "At-risk accounts · 12 weeks", legend: "37 flagged · 9 orphans · 12 toxic combos" },
      identities: { caption: "New identities per week", legend: "+64 joiners · +41 movers · 12 leavers" }
    };
    segBtns.forEach((btn) => {
      btn.addEventListener("click", () => {
        const v = btn.dataset.view;
        segBtns.forEach((b) => {
          const on = b === btn;
          b.classList.toggle("is-active", on);
          b.setAttribute("aria-selected", String(on));
        });
        datasets.forEach((g) => g.classList.toggle("is-active", g.dataset.dataset === v));
        if (caption && copy[v]) caption.textContent = copy[v].caption;
        if (legend && copy[v]) legend.textContent = copy[v].legend;
      });
    });
  }

  /* ---------- Depth layer from the Figma export (optional) ---------- */
  const windowBack = $("#windowBack");
  if (windowBack) {
    fetch("assets/figma-insights.png", { method: "HEAD" })
      .then((res) => {
        if (res.ok) {
          const img = document.createElement("img");
          img.src = "assets/figma-insights.png";
          img.alt = "";
          windowBack.appendChild(img);
        }
      })
      .catch(() => { /* no export yet: the layer stays empty, fine */ });
  }

  /* ---------- Asset upgrades from future Figma exports ---------- */
  function whenAssetExists(path, onReady) {
    fetch(path, { method: "HEAD" })
      .then((res) => { if (res.ok) onReady(); })
      .catch(() => { /* not exported yet */ });
  }
  whenAssetExists("assets/figma-connectors.png", () => {
    const img = document.querySelector("#integrations .section-figure img");
    if (img) {
      img.src = "assets/figma-connectors.png";
      img.alt = "Connector catalog: sixty-plus source cards by category, from HR to security.";
    }
  });
  whenAssetExists("assets/figma-excel.png", () => {
    const section = document.getElementById("reviews");
    const container = section && section.querySelector(".container");
    if (container && !container.querySelector(".excel-figure")) {
      const fig = document.createElement("figure");
      fig.className = "section-figure excel-figure";
      const img = document.createElement("img");
      img.src = "assets/figma-excel.png";
      img.alt = "The certification loop, killed: crossed-out spreadsheet on the left, completed Kaidra campaign on the right.";
      fig.appendChild(img);
      container.appendChild(fig);
    }
  });
})();
