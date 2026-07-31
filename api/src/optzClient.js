const axios = require("axios");

const BASE = "https://otimizacao.optz.com.br";

const headers = (jwt) => ({ Authorization: `Bearer ${jwt}` });

async function fetchLabels({ jwt, empresaUuid, area, inicio, fim }) {
  const parts = [
    `area=${area}`,
    `inicio=${inicio}`,
    `fim=${fim}`,
    `empresas=${empresaUuid}`,
  ];
  const url = `${BASE}/coordenacoes/labels?${parts.join("&")}`;
  try {
    const { data } = await axios.get(url, { headers: headers(jwt), timeout: 20000 });
    return Array.isArray(data) ? data : (data?.data || []);
  } catch (e) {
    const body = e.response?.data;
    if (body) console.error("[optzClient] labels erro body:", JSON.stringify(body).slice(0, 300));
    throw e;
  }
}

async function fetchServicos({ jwt, empresaUuid, area, inicio, fim, labels }) {
  const parts = [
    `area=${area}`,
    `inicio=${inicio}`,
    `fim=${fim}`,
    `empresas[0]=${empresaUuid}`,
    ...labels.map((l, i) => `labels[${i}]=${l}`),
  ];
  const url = `${BASE}/coordenacoes/servicos?${parts.join("&")}`;
  try {
    const { data } = await axios.get(url, { headers: headers(jwt), timeout: 20000 });
    return data?.data?.trilhos || [];
  } catch (e) {
    const body = e.response?.data;
    if (body) console.error("[optzClient] servicos erro body:", JSON.stringify(body).slice(0, 300));
    throw e;
  }
}

module.exports = { fetchLabels, fetchServicos };
