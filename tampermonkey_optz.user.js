// ==UserScript==
// @name         OPTZ — Interceptor de API v3
// @namespace    https://viacaocatedral.com.br
// @version      3.0.0
// @description  Intercepta respostas da API do OPTZ e envia dados estruturados para API local de ocorrências. Não depende de DOM/CSS.
// @author       Lucas Inácio
// @match        https://app.optz.com.br/viacaocatedral*
// @grant        none
// @run-at       document-start
// ==/UserScript==

(function () {
  "use strict";

  const LOCAL_API = "http://localhost:3000/api";
  const OPTZ_API  = "otimizacao.optz.com.br";

  let _lastJwt = null;

  // ── Utilitários ──────────────────────────────────────────
  function post(path, payload) {
    fetch(`${LOCAL_API}${path}`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(payload),
    }).catch(() => {});
  }

  function toast(msg, err = false) {
    const el = document.createElement("div");
    Object.assign(el.style, {
      position: "fixed", bottom: "28px", left: "50%", transform: "translateX(-50%)",
      background: err ? "#820014" : "#001d66", color: "#fff",
      padding: "10px 20px", borderRadius: "8px", fontSize: "13px",
      zIndex: "2147483647", pointerEvents: "none",
    });
    el.textContent = msg;
    document.body?.appendChild(el);
    setTimeout(() => el.remove(), 3500);
  }

  // ── Lê nomes dos motoristas/frotas do DOM do Gantt ───────
  function lerGruposDom(debugLabel) {
    const painel = document.querySelector("#area-coord-gantt\\:grupo");
    if (!painel) {
      console.warn(`[OPTZ v3] lerGruposDom(${debugLabel}): painel não encontrado`);
      return {};
    }
    const rows = painel.querySelectorAll("div[data-grupo-id]");
    console.log(`[OPTZ v3] lerGruposDom(${debugLabel}): ${rows.length} linhas no DOM`);
    const map = {};
    rows.forEach((row) => {
      const id = row.dataset.grupoId;
      if (id) map[id] = row.getAttribute("title") || "";
    });
    // Log primeiros 3 para diagnóstico
    const sample = Object.entries(map).slice(0, 3);
    if (sample.length) console.log(`[OPTZ v3] DOM sample:`, sample);
    return map;
  }

  // ── Parsers de título (mesmo formato do DOM) ─────────────
  function parseTripulante(title) {
    // "VIACAO CATEDRAL LTDA - 0009 - CLEISON VIEIRA DA COSTA"
    const p = title.replace(/^VIACAO CATEDRAL LTDA\s+-\s+/i, "").split(" - ");
    return p.length >= 2
      ? { matricula: p[0].trim(), nome: p.slice(1).join(" - ").trim() }
      : { matricula: "", nome: title.trim() };
  }

  function parseFrota(title) {
    // "VIACAO CATEDRAL LTDA - DD46 LT CARVA - 24109 - SSJ6B09"
    const sem = title.replace(/^VIACAO CATEDRAL LTDA\s+-\s+/i, "").trim();
    const p = sem.split(" - ");
    if (p.length >= 3) {
      return { placa: p[p.length - 1].trim(), prefixo: p[p.length - 2].trim(), linha: p.slice(0, p.length - 2).join(" - ").trim() };
    }
    return p.length === 2
      ? { prefixo: p[1].trim(), linha: p[0].trim(), placa: "" }
      : { prefixo: "", linha: sem, placa: "" };
  }

  // ── Extrai metadados da URL de servicos ──────────────────
  function metaDeUrl(url) {
    try {
      const u = new URL(url);
      // URLSearchParams não reconhece chaves com colchetes — parseia via regex
      const empresaMatch = url.match(/[?&]empresas(?:\[0\]|%5B0%5D)=([^&]+)/i);
      const empresaUuid  = empresaMatch ? decodeURIComponent(empresaMatch[1]) : "";
      return {
        area:   u.searchParams.get("area") || "1",
        inicio: u.searchParams.get("inicio") || "",
        fim:    u.searchParams.get("fim") || "",
        empresaUuid,
      };
    } catch { return { area: "1", inicio: "", fim: "", empresaUuid: "" }; }
  }

  // ── Cache de labels vindos de /coordenacoes/labels ──────
  // Chave: uuid → título completo (independe de scroll/DOM)
  const _labelsCache = {};

  function handleLabelsData(url, responseText) {
    try {
      const data  = JSON.parse(responseText);
      const lista = Array.isArray(data) ? data : (data?.data || []);
      const payload = [];

      lista.forEach((item) => {
        const uuid = item.key || item.labelId || "";
        const help = item.help || "";
        if (!uuid) return;

        // Armazena no cache para enriquecer trilhos
        if (help) _labelsCache[uuid] = help;

        // Detecta tipo pelo formato do help
        const sem     = help.replace(/^VIACAO CATEDRAL LTDA\s+-\s+/i, "").trim();
        const partes  = sem.split(" - ");
        const isFrota = partes.length >= 3;

        if (isFrota) {
          // title = prefixo, subTitle = linha, última parte do help = placa
          const prefixo = String(item.title || "").trim();
          const linha   = String(item.subTitle || "").trim();
          const placa   = partes[partes.length - 1]?.trim() || "";
          if (prefixo) payload.push({ uuid, tipo: "frota", prefixo, linha, placa });
        } else {
          // title = matricula, resto do help = nome
          const matricula = String(item.title || "").trim();
          const nome      = partes.slice(1).join(" - ").trim();
          if (matricula) payload.push({ uuid, tipo: "tripulante", matricula, nome });
        }
      });

      console.log(`[OPTZ v3] labels: ${payload.length} grupos parseados`);
      if (payload.length) post("/optz/grupos", payload);
    } catch (e) {
      console.error("[OPTZ v3] labels parse error:", e);
    }
  }

  // Detecta se o DOM atual é de frotas ou tripulantes pelo formato do título
  function detectarViewDom(gruposTitulos) {
    const titulos = Object.values(gruposTitulos).filter(Boolean);
    if (!titulos.length) return null;
    // Título de frota: "VIACAO CATEDRAL LTDA - LINHA - PREFIXO - PLACA" (≥3 partes após remover empresa)
    const isFrota = titulos.some((t) => {
      const sem = t.replace(/^VIACAO CATEDRAL LTDA\s+-\s+/i, "").trim();
      return sem.split(" - ").length >= 3;
    });
    return isFrota ? "frotas" : "tripulantes";
  }

  // ── Handler comum para respostas de /coordenacoes/servicos ──
  function handleServicosData(url, data) {
    const trilhos = data?.data?.trilhos || [];
    const { area, inicio, fim, empresaUuid } = metaDeUrl(url);
    console.log(`[OPTZ v3] servicos area=${area}: ${trilhos.length} trilhos (${inicio} → ${fim})`);

    const delays = [600, 1400, 2500];

    function tryEnrich(attempt) {
      setTimeout(() => {
        const cacheSize   = Object.keys(_labelsCache).length;
        const gruposTitulos = cacheSize > 0 ? _labelsCache : lerGruposDom(`area${area}-t${attempt + 1}`);
        const n = Object.keys(gruposTitulos).length;
        console.log(`[OPTZ v3] fonte: ${cacheSize > 0 ? "labels-cache" : "DOM"} (${n} grupos)`);

        if (n === 0 && attempt < delays.length - 1) {
          console.log(`[OPTZ v3] grupos vazios, aguardando...`);
          tryEnrich(attempt + 1);
          return;
        }

        const viewTipo = detectarViewDom(gruposTitulos);
        const isFrota  = viewTipo === "frotas";
        console.log(`[OPTZ v3] view detectada: ${viewTipo || "desconhecida"} (area URL=${area})`);

        const trilhosEnriquecidos = trilhos.map((t) => {
          const titulo = gruposTitulos[String(t.key)] || "";
          const meta = isFrota ? parseFrota(titulo) : parseTripulante(titulo);
          return { ...t, ...meta };
        });

        const comMeta = trilhosEnriquecidos.filter(
          (t) => (isFrota ? t.prefixo : t.matricula)
        ).length;
        console.log(`[OPTZ v3] ${comMeta}/${trilhos.length} trilhos com metadata`);

        if (comMeta > 0) {
          const storeArea = isFrota ? "2" : "1";
          post("/optz/servicos", { area: storeArea, inicio, fim, empresaUuid, trilhos: trilhosEnriquecidos });
          toast(`✅ OPTZ: ${comMeta}/${trilhos.length} ${viewTipo} sincronizados`);


        } else {
          console.log(`[OPTZ v3] Nenhum match — XHR descartado (area=${area})`);
        }
      }, delays[attempt]);
    }

    tryEnrich(0);
  }

  // ── Interceptor de XHR ───────────────────────────────────
  // O OPTZ usa XHR (não fetch) para /coordenacoes/servicos
  const origOpen        = XMLHttpRequest.prototype.open;
  const origSetHeader   = XMLHttpRequest.prototype.setRequestHeader;
  const origSend        = XMLHttpRequest.prototype.send;

  XMLHttpRequest.prototype.open = function (method, url) {
    this._optzUrl     = String(url || "");
    this._optzHeaders = {};
    return origOpen.apply(this, arguments);
  };

  XMLHttpRequest.prototype.setRequestHeader = function (name, value) {
    if (this._optzHeaders) this._optzHeaders[String(name)] = String(value);
    return origSetHeader.apply(this, arguments);
  };

  XMLHttpRequest.prototype.send = function () {
    this.addEventListener("readystatechange", function () {
      if (this.readyState !== 4) return;
      const url = this._optzUrl || "";

      // Captura JWT via XHR
      const auth = this._optzHeaders?.Authorization || this._optzHeaders?.authorization || "";
      if (auth.startsWith("Bearer ") && url.includes("optz.com.br")) {
        const jwt = auth.slice(7);
        if (jwt !== _lastJwt) {
          _lastJwt = jwt;
          post("/optz/auth", { jwt });
          console.log("[OPTZ v3] JWT (XHR) capturado");
        }
      }

      // Loga e captura empresa UUID de qualquer endpoint do OPTZ
      if (url.includes(OPTZ_API)) {
        try {
          const u = new URL(url);
          console.log(`[OPTZ v3] XHR → ${u.pathname}${u.search.slice(0, 60)}`);
        } catch (_) {}
        // Captura empresa UUID de empresas[0]=... ou empresas=... em qualquer URL
        const empMatch = url.match(/[?&]empresas(?:\[0\]|%5B0%5D)?=([0-9a-f-]{36})/i);
        if (empMatch) post("/optz/empresa", { uuid: empMatch[1] });
      }

      // Captura labels (todos os grupos — resolve virtual scrolling)
      if (url.includes(OPTZ_API) && url.includes("/coordenacoes/labels")) {
        handleLabelsData(url, this.responseText);
      }

      // Captura resposta de servicos
      if (url.includes(OPTZ_API) && url.includes("/coordenacoes/servicos")) {
        console.log("[OPTZ v3] XHR servicos interceptado:", url.slice(0, 80));
        try {
          const data = JSON.parse(this.responseText);
          handleServicosData(url, data);
        } catch (e) {
          console.error("[OPTZ v3] XHR parse error:", e);
        }
      }
    });
    return origSend.apply(this, arguments);
  };

  // ── Interceptor de fetch (captura JWT e resposta via fetch) ──
  const origFetch = window.fetch;
  window.fetch = async function (...args) {
    const [resource, options = {}] = args;
    const url = typeof resource === "string" ? resource : (resource?.url || "");

    // Captura JWT via fetch
    const authHdr = options.headers?.Authorization || options.headers?.authorization || "";
    if (authHdr.startsWith("Bearer ") && url.includes("optz.com.br")) {
      const jwt = authHdr.slice(7);
      if (jwt !== _lastJwt) {
        _lastJwt = jwt;
        post("/optz/auth", { jwt });
        console.log("[OPTZ v3] JWT (fetch) capturado e enviado");
      }
    }

    const response = await origFetch.apply(this, args);

    // Captura resposta de servicos via fetch (caso o app mude para fetch)
    if (url.includes(OPTZ_API) && url.includes("/coordenacoes/servicos")) {
      console.log("[OPTZ v3] fetch servicos interceptado:", url.slice(0, 80));
      response.clone().json()
        .then((data) => handleServicosData(url, data))
        .catch((e) => console.error("[OPTZ v3] fetch parse error:", e));
    }

    return response;
  };

  // ── MutationObserver: captura linhas do Gantt ao fazer scroll ──
  // Resolve frotas/tripulantes que não estavam visíveis quando o XHR disparou
  function iniciarObserverGantt() {
    let debounce = null;
    const pendentes = new Map(); // uuid → { tipo, meta }

    const obs = new MutationObserver((mutations) => {
      for (const m of mutations) {
        for (const node of m.addedNodes) {
          if (node.nodeType !== 1) continue;
          const rows = node.matches?.("div[data-grupo-id]")
            ? [node]
            : [...(node.querySelectorAll?.("div[data-grupo-id]") ?? [])];

          for (const row of rows) {
            const id    = row.dataset.grupoId;
            const title = row.getAttribute("title") || "";
            if (!id || !title.includes("VIACAO CATEDRAL")) continue;

            const sem    = title.replace(/^VIACAO CATEDRAL LTDA\s+-\s+/i, "").trim();
            const isFrota = sem.split(" - ").length >= 3;
            const meta   = isFrota ? parseFrota(title) : parseTripulante(title);
            const chave  = isFrota ? meta.prefixo : meta.matricula;
            if (!chave) continue;

            pendentes.set(id, { tipo: isFrota ? "frota" : "tripulante", ...meta });
          }
        }
      }

      // Envia em lote após 300ms de inatividade
      clearTimeout(debounce);
      debounce = setTimeout(() => {
        if (!pendentes.size) return;
        const payload = [];
        pendentes.forEach((v, uuid) => payload.push({ uuid, ...v }));
        pendentes.clear();
        post("/optz/grupos", payload);
        console.log(`[OPTZ v3] Observer: ${payload.length} grupos enviados`);
      }, 300);
    });

    // Aguarda o body estar disponível
    const startObs = () => {
      if (document.body) {
        obs.observe(document.body, { childList: true, subtree: true });
        console.log("[OPTZ v3] MutationObserver Gantt ativo");
      } else {
        setTimeout(startObs, 200);
      }
    };
    startObs();
  }

  iniciarObserverGantt();

  console.log("[OPTZ v3] Interceptores fetch + XHR ativos");
})();
