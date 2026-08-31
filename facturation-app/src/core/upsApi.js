// ============================================================================
//  API UPS (Tracking) -- repli automatique pour recuperer le poids officiel
//  d'un colis quand aucune autre source ne l'a fourni (poids audite en texte
//  dans le CSV Billing, export WMS expeditions_brut m/m-1/m-2). Remplace la
//  verification manuelle sur billing.ups.com/ups.com/track vue dans la video
//  process "UPS_Preparation fichier import.mp4" (2026-08-27).
//
//  Necessite UPS_CLIENT_ID / UPS_CLIENT_SECRET dans facturation-app/.env
//  (app UPS Developer, flux Client Credentials -- cf. developer.ups.com,
//  produit "Tracking" souscrit). Sans ces variables, toutes les fonctions
//  ci-dessous sont des no-op silencieux (retournent null) : l'appelant doit
//  garder son comportement actuel (alerte POIDS=0) si l'API n'est pas
//  configuree, jamais planter pour autant.
// ============================================================================

const OAUTH_URL = 'https://onlinetools.ups.com/security/v1/oauth/token';
const TRACK_URL = (trackingNumber) => `https://onlinetools.ups.com/api/track/v1/details/${encodeURIComponent(trackingNumber)}`;
// Timeout par appel -- BUG TROUVE 2026-08-27 : sans lui, un appel UPS qui ne repond jamais
// (reseau instable, throttling) bloque toute la generation indefiniment (constate : test reste
// bloque >20 min, contre ~7 min sans l'API). Le carrier UPS boucle sur chaque tracking non
// resolu de facon SEQUENTIELLE -- un seul appel bloque suffit a tout geler.
const FETCH_TIMEOUT_MS = 8000;

// Cache memoire du token OAuth (evite de redemander un token a chaque tracking --
// un process de generation peut interroger des dizaines de trackings). Un token
// UPS dure typiquement ~1h ; on le renouvelle avec une marge de securite de 60s.
let cachedToken = null;
let cachedTokenExpiry = 0;

function isConfigured() {
  return !!(process.env.UPS_CLIENT_ID && process.env.UPS_CLIENT_SECRET);
}

async function getAccessToken() {
  if (!isConfigured()) return null;
  const now = Date.now();
  if (cachedToken && now < cachedTokenExpiry) return cachedToken;

  const basic = Buffer.from(`${process.env.UPS_CLIENT_ID}:${process.env.UPS_CLIENT_SECRET}`).toString('base64');
  const res = await fetch(OAUTH_URL, {
    method: 'POST',
    headers: {
      Authorization: `Basic ${basic}`,
      'Content-Type': 'application/x-www-form-urlencoded',
    },
    body: 'grant_type=client_credentials',
    signal: AbortSignal.timeout(FETCH_TIMEOUT_MS),
  });
  if (!res.ok) {
    throw new Error(`UPS OAuth: ${res.status} ${res.statusText} -- ${(await res.text()).slice(0, 300)}`);
  }
  const data = await res.json();
  cachedToken = data.access_token;
  // expires_in en secondes (reponse UPS) ; marge de 60s pour eviter d'utiliser un
  // token expire pile au moment d'un appel.
  cachedTokenExpiry = now + (Number(data.expires_in) || 3000) * 1000 - 60000;
  return cachedToken;
}

/** Poids (en KGS, nombre) d'un tracking via l'API UPS Tracking, ou null si
 * l'API n'est pas configuree, si le tracking est introuvable, ou si aucun
 * poids n'est present dans la reponse. Ne leve JAMAIS pour un cas "normal"
 * (tracking absent/pas de poids) -- seules les erreurs reseau/auth remontent. */
async function poidsParTrackingViaApi(trackingNumber) {
  if (!isConfigured() || !trackingNumber) return null;
  const token = await getAccessToken();
  if (!token) return null;

  const res = await fetch(TRACK_URL(trackingNumber), {
    headers: {
      Authorization: `Bearer ${token}`,
      transId: String(Date.now()),
      transactionSrc: 'facturation-transporteurs',
    },
    signal: AbortSignal.timeout(FETCH_TIMEOUT_MS),
  });
  if (res.status === 404) return null; // tracking inconnu de l'API -- pas une erreur
  if (!res.ok) {
    throw new Error(`UPS Tracking API: ${res.status} ${res.statusText} -- ${(await res.text()).slice(0, 300)}`);
  }
  const data = await res.json();
  // Structure confirmee (Tracking.yaml) : trackResponse.shipment[].package[].weight
  // = { unitOfMeasurement, weight (string) }.
  const shipments = (data && data.trackResponse && data.trackResponse.shipment) || [];
  for (const shipment of shipments) {
    for (const pkg of shipment.package || []) {
      const w = pkg.weight && Number(String(pkg.weight.weight).replace(',', '.'));
      if (Number.isFinite(w) && w > 0) {
        // KGS attendu (config API), mais on convertit par securite si jamais LBS.
        const unit = String((pkg.weight.unitOfMeasurement && pkg.weight.unitOfMeasurement.code) || pkg.weight.unitOfMeasurement || 'KGS').toUpperCase();
        return unit === 'LBS' ? Math.round(w * 0.45359237 * 100) / 100 : w;
      }
    }
  }
  return null;
}

module.exports = { isConfigured, poidsParTrackingViaApi };
