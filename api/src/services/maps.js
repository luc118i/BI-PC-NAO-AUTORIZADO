const axios = require("axios");

const KEY = process.env.MAPS_KEY || "";

const STATIC_BASE = "https://maps.googleapis.com/maps/api/staticmap";
const SV_BASE     = "https://maps.googleapis.com/maps/api/streetview";
const SV_META     = "https://maps.googleapis.com/maps/api/streetview/metadata";

// Retorna as URLs públicas das imagens.
// Para embutir em relatório HTML basta usar como <img src="...">.
// Se precisar de base64 (ex: PDF), chame _fetchBase64 internamente.

async function buscarImagensMaps(lat, lng) {
  if (!lat || !lng || !KEY) {
    return { satelite: null, streetView: null, svDisponivel: false };
  }

  const coord = `${lat},${lng}`;

  const sateliteUrl =
    `${STATIC_BASE}?center=${coord}&zoom=17&size=600x340` +
    `&maptype=satellite&markers=color:red|${coord}&key=${KEY}`;

  const streetViewUrl =
    `${SV_BASE}?size=600x340&location=${coord}&fov=90&pitch=10&key=${KEY}`;

  // Verifica se Street View existe para estas coordenadas
  let svDisponivel = false;
  try {
    const meta = await axios.get(SV_META, {
      params: { location: coord, key: KEY },
      timeout: 5000,
    });
    svDisponivel = meta.data?.status === "OK";
  } catch {
    // ignora — retorna URL mesmo assim, a img pode falhar no navegador
  }

  return { satelite: sateliteUrl, streetView: streetViewUrl, svDisponivel };
}

module.exports = { buscarImagensMaps };
