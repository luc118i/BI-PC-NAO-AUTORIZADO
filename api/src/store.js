// Store em memória — dados capturados via interceptor do OPTZ
// area=1 → tripulantes, area=2 → frotas

// ── OPTZ: servicos por trilho (uuid → dados) ─────────────
const _tripulantesMap = new Map(); // uuid → { matricula, nome, servicos[] }
const _frotasMap      = new Map(); // uuid → { prefixo, linha, placa, servicos[] }
let   _jwtStore       = { token: null, exp: 0 };
let   _empresaUuid    = null;

function setJwt(jwt) {
  try {
    const payload = JSON.parse(Buffer.from(jwt.split(".")[1], "base64").toString("utf8"));
    _jwtStore = { token: jwt, exp: payload.exp || 0 };
  } catch (_) {
    _jwtStore = { token: jwt, exp: 0 };
  }
}

function setEmpresaUuid(uuid) {
  if (uuid && uuid !== _empresaUuid) {
    _empresaUuid = uuid;
    console.log("[store] empresaUuid:", uuid.slice(0, 8) + "...");
  }
}
function getEmpresaUuid() { return _empresaUuid; }

function getTripulanteUuids() { return [..._tripulantesMap.keys()]; }
function getFrotaUuids()      { return [..._frotasMap.keys()]; }

function getFrotaUuidByPrefixo(prefixo) {
  const pref = String(prefixo).trim();
  for (const [uuid, f] of _frotasMap) {
    if (String(f.prefixo || "").trim() === pref) return uuid;
  }
  return null;
}

function getJwt() {
  if (!_jwtStore.token) return null;
  if (_jwtStore.exp && Math.floor(Date.now() / 1000) > _jwtStore.exp - 30) return null;
  return _jwtStore.token;
}

function setServicosOptz(area, inicio, fim, trilhos) {
  const map = String(area) === "2" ? _frotasMap : _tripulantesMap;

  for (const t of trilhos) {
    const key = t.key;
    const existing = map.get(key) || {};

    // Mescla serviços, deduplicando pelo campo key
    const servicosAntigos = existing.servicos || [];
    const servicosNovos   = t.servicos || [];
    const todosSvcs       = [...servicosAntigos, ...servicosNovos];
    const dedup           = todosSvcs.filter((s, i, arr) => arr.findIndex(x => x.key === s.key) === i);

    // Meta do DOM (matricula/nome para tripulantes; prefixo/linha/placa para frotas)
    const meta = {};
    if (String(area) === "2") {
      if (t.prefixo) meta.prefixo = t.prefixo;
      if (t.linha)   meta.linha   = t.linha;
      if (t.placa)   meta.placa   = t.placa;
    } else {
      if (t.matricula) meta.matricula = t.matricula;
      if (t.nome)      meta.nome      = t.nome;
    }

    map.set(key, { ...existing, ...meta, servicos: dedup, _updatedAt: new Date().toISOString() });
  }
}

// Busca motorista pelo código de serviço e data/hora opcional
function buscarMotoristaPorServicoOptz(codigoServico, dtOcorrencia) {
  const num = String(codigoServico).trim();
  const dt  = dtOcorrencia ? new Date(dtOcorrencia) : null;

  for (const [uuid, t] of _tripulantesMap) {
    const svc = (t.servicos || []).find((s) => {
      if (s.tipo !== 1) return false; // apenas tipo Serviço (não folga/interjornada)
      // codigo pode ser "314000719" ou "103915003/103915006"
      const partes = String(s.codigo || "").split("/");
      if (!partes.some((p) => p.trim() === num)) return false;
      if (dt && s.inicio && s.fim) {
        return dt >= new Date(s.inicio) && dt <= new Date(s.fim);
      }
      return true;
    });
    if (svc) {
      return {
        uuid,
        matricula:  t.matricula || "",
        nome:       t.nome || "",
        servico: svc,
      };
    }
  }
  return null;
}

// Busca serviço ativo de um veículo em determinado momento
function buscarServicoPorFrota(prefixo, dtOcorrencia) {
  const pref = String(prefixo).trim();
  const dt   = dtOcorrencia ? new Date(dtOcorrencia) : null;

  for (const [uuid, f] of _frotasMap) {
    if (String(f.prefixo || "").trim() !== pref) continue;
    const svc = (f.servicos || []).find((s) => {
      if (s.tipo !== 1) return false;
      if (dt && s.inicio && s.fim) {
        return dt >= new Date(s.inicio) && dt <= new Date(s.fim);
      }
      return true; // sem data: retorna qualquer serviço tipo 1
    });
    if (svc) {
      return {
        uuid,
        prefixo: f.prefixo,
        linha:   f.linha || "",
        placa:   f.placa || "",
        servicoAtivo: svc,
      };
    }
  }
  return null;
}

// Retorna dados completos de uma frota pelo UUID (para debug)
function getFrotaDetalhes(uuid) {
  const f = _frotasMap.get(uuid);
  if (!f) return null;
  return { uuid, prefixo: f.prefixo, linha: f.linha, placa: f.placa, servicos: f.servicos || [] };
}

// Atualiza metadata (prefixo/nome) de grupos que chegaram via MutationObserver
function updateGruposMeta(grupos) {
  let count = 0;
  for (const g of grupos) {
    const { uuid, tipo, prefixo, linha, placa, matricula, nome } = g;
    if (!uuid) continue;
    if (tipo === "frota" && prefixo) {
      const existing = _frotasMap.get(uuid);
      if (existing && !existing.prefixo) {
        _frotasMap.set(uuid, { ...existing, prefixo, linha: linha || existing.linha, placa: placa || existing.placa });
        count++;
      } else if (!existing) {
        _frotasMap.set(uuid, { prefixo, linha: linha || "", placa: placa || "", servicos: [] });
        count++;
      }
    } else if (tipo === "tripulante" && matricula) {
      const existing = _tripulantesMap.get(uuid);
      if (existing && !existing.matricula) {
        _tripulantesMap.set(uuid, { ...existing, matricula, nome: nome || existing.nome });
        count++;
      } else if (!existing) {
        _tripulantesMap.set(uuid, { matricula, nome: nome || "", servicos: [] });
        count++;
      }
    }
  }
  return count;
}

function temServicosParaDia(dia) {
  const prefix = String(dia).slice(0, 10);
  let frotas = false, tripulantes = false;
  for (const [, f] of _frotasMap) {
    if ((f.servicos || []).some(s => s.inicio && String(s.inicio).startsWith(prefix))) { frotas = true; break; }
  }
  for (const [, t] of _tripulantesMap) {
    if ((t.servicos || []).some(s => s.inicio && String(s.inicio).startsWith(prefix))) { tripulantes = true; break; }
  }
  return { frotas, tripulantes, loaded: frotas && tripulantes };
}

function statusOptz() {
  const jwtExp  = _jwtStore.exp ? new Date(_jwtStore.exp * 1000).toISOString() : null;
  const jwtOk   = !!getJwt();
  const trips   = [];
  _tripulantesMap.forEach((v, k) => trips.push({ uuid: k, matricula: v.matricula, nome: v.nome, svcs: (v.servicos || []).length }));
  const frotas  = [];
  _frotasMap.forEach((v, k) => frotas.push({ uuid: k, prefixo: v.prefixo, svcs: (v.servicos || []).length }));
  return { jwtOk, jwtExp, tripulantes: trips.length, frotas: frotas.length, trips, frotas };
}

// ── Legado: escala/frota via Tampermonkey DOM (mantido para compatibilidade) ──
const _escalaStore = new Map();
const _frotaStore  = new Map();

function setEscala(matricula, payload) {
  _escalaStore.set(String(matricula).trim(), { ...payload, _savedAt: new Date().toISOString() });
}
function getEscala(matricula) {
  return _escalaStore.get(String(matricula).trim()) || null;
}
function listEscalas() {
  const r = [];
  _escalaStore.forEach((v, k) => r.push({ matricula: k, nome: v.nome, savedAt: v._savedAt }));
  return r;
}
function setFrota(prefixo, payload) {
  _frotaStore.set(String(prefixo).trim(), { ...payload, _savedAt: new Date().toISOString() });
}
function getFrota(prefixo) {
  return _frotaStore.get(String(prefixo).trim()) || null;
}
function listFrotas() {
  const r = [];
  _frotaStore.forEach((v, k) => r.push({ prefixo: k, linha: v.linha, savedAt: v._savedAt }));
  return r;
}
function buscarMotoristaPorServico(numeroServico) {
  const num = String(numeroServico).trim();
  for (const [matricula, payload] of _escalaStore) {
    const bloco = (payload.escala || []).find((s) => {
      const partes = String(s.numero || "").split("/");
      return partes.some((p) => p.trim() === num);
    });
    if (bloco) return { matricula, nome: payload.nome, servico: bloco };
  }
  return null;
}

module.exports = {
  // OPTZ API interceptor (v3)
  setJwt, getJwt,
  setEmpresaUuid, getEmpresaUuid,
  getTripulanteUuids, getFrotaUuids,
  getFrotaUuidByPrefixo,
  setServicosOptz,
  updateGruposMeta,
  getFrotaDetalhes,
  buscarMotoristaPorServicoOptz,
  buscarServicoPorFrota,
  temServicosParaDia,
  statusOptz,
  // Legado DOM-scraping (v2)
  setEscala, getEscala, listEscalas,
  setFrota, getFrota, listFrotas,
  buscarMotoristaPorServico,
};
