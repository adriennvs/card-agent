"""
Card Opportunity Agent — v1
Univers  : DBZ années 90 (Carddass Hondan, Le Grand Combat, Super Battle,
           Power Level, Visual Adventure) + Yu-Gi-Oh! anciennes séries
Sources  : Vinted (ScrapeBadger) · eBay sold listings · Cardmarket
Scoring  : ratio prix annonce / prix marché · liquidité · mots-clés de valeur
Budget   : annonces filtrées < 50€
"""

import os
import re
import time
import json
import random
import logging
import requests
import gspread
from scrapebadger import ScrapeBadger as SBClient
import smtplib
from datetime import datetime
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from google.oauth2.service_account import Credentials

logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(message)s")
log = logging.getLogger(__name__)

# ─── Configuration ────────────────────────────────────────────────────────────

BUDGET_MAX    = 50      # € — filtre dur
SCORE_ALERT   = 70      # seuil alerte email
MIN_MARGE     = 0.40    # marge minimum 40%

SHEET_NAME       = os.getenv("GOOGLE_SHEET_NAME", "Card Agent")
GMAIL_FROM       = os.getenv("GMAIL_FROM")
GMAIL_TO         = os.getenv("GMAIL_TO")
GMAIL_PASS       = os.getenv("GMAIL_APP_PASSWORD")
GCP_CREDS        = os.getenv("GCP_SERVICE_ACCOUNT_JSON")
SCRAPEBADGER_KEY = os.getenv("SCRAPEBADGER_API_KEY")
EBAY_APP_ID      = os.getenv("EBAY_APP_ID")

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/120.0.0.0 Safari/537.36"
    )
}

# ─── Mots-clés de recherche ────────────────────────────────────────────────────

# DBZ — termes de recherche Vinted
DBZ_SEARCH_TERMS = [
    "carddass hondan",
    "le grand combat",
    "super battle dragon ball",
    "power level dragon ball",
    "visual adventure dragon ball",
    "cartes dragon ball z années 90",
    "cartes dbz vintage",
    "carddass dragon ball",
    "lot cartes dragon ball z",
]

# DBZ — mots-clés qui indiquent une carte de valeur
DBZ_HIGH_VALUE_KEYWORDS = [
    "prisme", "prism", "flash", "hors série", "hors-série",
    "hondan", "grand combat", "super battle", "power level",
    "visual adventure", "carddass", "rare", "special",
    "goku", "vegeta", "freezer", "cell", "buu", "gohan",
    "gotenks", "broly", "gogeta", "vegeto",
]

# YGO — termes de recherche Vinted
YGO_SEARCH_TERMS = [
    "yu-gi-oh lot ancien",
    "yugioh vintage",
    "cartes yugioh anciennes",
    "lot yu gi oh années 2000",
    "blue eyes white dragon",
    "dark magician yugioh",
    "yugioh 1ère édition",
    "yugioh first edition",
    "lot cartes yugioh",
]

# YGO — mots-clés haute valeur
YGO_HIGH_VALUE_KEYWORDS = [
    "1ère édition", "1st edition", "first edition",
    "ultra rare", "secret rare", "ghost rare",
    "blue eyes", "dark magician", "exodia",
    "limited edition", "lob", "mrd", "mrl", "psr",
    "sdk", "sdm", "metal raiders", "legend of blue eyes",
    "magician's force", "pharaoh's servant",
    "porte-bonheur", "holo", "holographique",
]

# ─── Google Sheets ─────────────────────────────────────────────────────────────

def get_sheet():
    creds_info = json.loads(GCP_CREDS)
    scopes = [
        "https://spreadsheets.google.com/feeds",
        "https://www.googleapis.com/auth/drive",
    ]
    creds = Credentials.from_service_account_info(creds_info, scopes=scopes)
    gc    = gspread.authorize(creds)
    return gc.open(SHEET_NAME)

def ensure_sheets(sh):
    existing = [w.title for w in sh.worksheets()]
    for name, headers in [
        ("opportunites", [
            "date_scan", "univers", "titre", "prix_annonce",
            "prix_marche_ref", "marge_estimee", "ratio_x",
            "score_global", "score_ratio", "score_liquidite", "score_valeur",
            "nb_ventes_ref", "mots_cles_valeur",
            "plateforme", "etat", "vendeur",
            "rationale", "lien_annonce"
        ]),
        ("portefeuille", [
            "date_achat", "univers", "article", "prix_achat",
            "plateforme_vente", "prix_revente", "statut",
            "date_vente", "prix_vente", "pnl"
        ]),
        ("prix_reference", [
            "date", "univers", "terme", "nb_ventes",
            "prix_moyen", "prix_median", "prix_min", "prix_max"
        ]),
    ]:
        if name not in existing:
            ws = sh.add_worksheet(title=name, rows=1000, cols=len(headers))
            ws.append_row(headers)
            log.info(f"Feuille créée : {name}")

# ─── Source 1 : Vinted via ScrapeBadger ───────────────────────────────────────

def search_vinted(query: str, max_price: float = BUDGET_MAX) -> list[dict]:
    """
    Recherche sur Vinted.fr via SDK officiel ScrapeBadger.
    Retourne les annonces avec titre, prix, état, lien.
    """
    if not SCRAPEBADGER_KEY:
        log.warning("SCRAPEBADGER_API_KEY non configurée")
        return []
    try:
        client  = SBClient(SCRAPEBADGER_KEY)
        results = client.vinted.search(
            query     = query,
            market    = "fr",
            price_to  = int(max_price),
            per_page  = 50,
        )
        items = []
        for item in results.items:
            try:
                price = float(item.price.amount)
            except Exception:
                continue
            if price <= 0 or price > max_price:
                continue
            # Lien Vinted
            item_id = item.id
            lien = f"https://www.vinted.fr/items/{item_id}"
            items.append({
                "titre":      str(item.title).strip(),
                "prix":       price,
                "etat":       str(getattr(item, "status", "") or ""),
                "vendeur":    str(getattr(item.seller, "username", "") if hasattr(item, "seller") else ""),
                "lien":       lien,
                "plateforme": "vinted",
                "query":      query,
            })
        log.info(f"Vinted '{query}' → {len(items)} annonces")
        return items
    except Exception as e:
        log.warning(f"ScrapeBadger error ('{query}'): {e}")
        return []

# ─── Source 2 : eBay sold listings — prix marché DBZ ──────────────────────────

def get_ebay_sold_prices(query: str) -> dict:
    """
    Récupère les prix de ventes réelles eBay pour un terme.
    Utilise l'API Finding de eBay (gratuite).
    Retourne : nb_ventes, prix_moyen, prix_median, prix_min, prix_max
    """
    if not EBAY_APP_ID:
        log.warning("EBAY_APP_ID non configuré")
        return {}
    try:
        r = requests.get(
            "https://svcs.ebay.com/services/search/FindingService/v1",
            params={
                "OPERATION-NAME":           "findCompletedItems",
                "SERVICE-VERSION":          "1.0.0",
                "SECURITY-APPNAME":         EBAY_APP_ID,
                "RESPONSE-DATA-FORMAT":     "JSON",
                "REST-PAYLOAD":             "",
                "keywords":                 query,
                "itemFilter(0).name":       "SoldItemsOnly",
                "itemFilter(0).value":      "true",
                "itemFilter(1).name":       "Currency",
                "itemFilter(1).value":      "EUR",
                "sortOrder":                "EndTimeSoonest",
                "paginationInput.entriesPerPage": "50",
                "outputSelector":           "AspectHistogram",
            },
            headers=HEADERS,
            timeout=15,
        )
        data = r.json()
        items_data = (
            data.get("findCompletedItemsResponse", [{}])[0]
                .get("searchResult", [{}])[0]
                .get("item", [])
        )
        prices = []
        for item in items_data:
            try:
                price = float(
                    item["sellingStatus"][0]["currentPrice"][0]["__value__"]
                )
                if 0 < price < 500:
                    prices.append(price)
            except Exception:
                pass

        if not prices:
            return {}

        prices.sort()
        n = len(prices)
        median = prices[n // 2]
        return {
            "nb_ventes":   n,
            "prix_moyen":  round(sum(prices) / n, 2),
            "prix_median": round(median, 2),
            "prix_min":    round(prices[0], 2),
            "prix_max":    round(prices[-1], 2),
        }
    except Exception as e:
        log.warning(f"eBay error ('{query}'): {e}")
        return {}

# ─── Source 3 : Cardmarket — prix marché Yu-Gi-Oh! ────────────────────────────

def get_cardmarket_ygo_price(card_name: str) -> dict:
    """
    Récupère le prix de référence d'une carte YGO sur Cardmarket.
    Utilise leur API publique (sans auth pour les prix moyens).
    """
    try:
        # Cardmarket expose des données via leur API publique
        # On utilise l'endpoint de recherche de produit
        r = requests.get(
            "https://api.cardmarket.com/ws/v2.0/products/find",
            params={
                "search":      card_name,
                "idGame":      3,      # 3 = Yu-Gi-Oh!
                "idLanguage":  2,      # 2 = Français
                "maxResults":  5,
            },
            headers={
                **HEADERS,
                "Accept": "application/json",
            },
            timeout=15,
        )
        if r.status_code != 200:
            return {}
        data = r.json()
        products = data.get("product", [])
        if not products:
            return {}
        # Prend le premier résultat le plus pertinent
        p = products[0]
        return {
            "prix_moyen":  p.get("priceGuide", {}).get("avg", 0),
            "prix_trend":  p.get("priceGuide", {}).get("trend", 0),
            "nb_ventes":   p.get("countSells", 0),
            "nb_offres":   p.get("countArticles", 0),
        }
    except Exception as e:
        log.warning(f"Cardmarket error ('{card_name}'): {e}")
        return {}

# ─── Détection mots-clés de valeur ────────────────────────────────────────────

def detect_value_keywords(text: str, keywords: list[str]) -> list[str]:
    """Retourne les mots-clés de valeur présents dans le texte."""
    text_lower = text.lower()
    found = [kw for kw in keywords if kw.lower() in text_lower]
    return found

def classify_universe(query: str, title: str) -> str:
    """Détermine si l'annonce est DBZ ou YGO."""
    text = (query + " " + title).lower()
    if any(k in text for k in ["dragon ball", "dbz", "carddass", "hondan",
                                 "grand combat", "super battle", "power level",
                                 "visual adventure"]):
        return "DBZ"
    if any(k in text for k in ["yu-gi-oh", "yugioh", "yu gi oh", "ygo",
                                 "blue eyes", "dark magician", "exodia"]):
        return "YGO"
    return "AUTRE"

# ─── Scoring ──────────────────────────────────────────────────────────────────

def score_annonce(annonce: dict, prix_ref: dict, value_keywords: list[str]) -> dict:
    """
    Score une annonce sur 3 dimensions :
    - Ratio prix annonce / prix marché (50%)
    - Liquidité = nb ventes similaires (30%)
    - Mots-clés de valeur dans le titre (20%)
    """
    prix      = annonce["prix"]
    titre     = annonce["titre"]
    univers   = annonce.get("univers", "")

    prix_ref_val = prix_ref.get("prix_median") or prix_ref.get("prix_moyen") or 0
    nb_ventes    = prix_ref.get("nb_ventes", 0)

    # Marge estimée
    marge = (prix_ref_val - prix) / max(prix_ref_val, 1) if prix_ref_val > 0 else 0
    ratio = prix_ref_val / max(prix, 1) if prix_ref_val > 0 else 1

    # Score ratio (50%)
    if ratio >= 5:    sr = 100
    elif ratio >= 3:  sr = 85
    elif ratio >= 2:  sr = 70
    elif ratio >= 1.5: sr = 55
    elif ratio >= 1.2: sr = 35
    else:              sr = 10

    # Score liquidité (30%)
    if nb_ventes >= 30:  sl = 100
    elif nb_ventes >= 15: sl = 80
    elif nb_ventes >= 8:  sl = 60
    elif nb_ventes >= 3:  sl = 40
    elif nb_ventes >= 1:  sl = 20
    else:                 sl = 0

    # Score mots-clés valeur (20%)
    n_kw = len(value_keywords)
    if n_kw >= 4:    skw = 100
    elif n_kw >= 3:  skw = 80
    elif n_kw >= 2:  skw = 60
    elif n_kw >= 1:  skw = 40
    else:            skw = 10

    score_global = int(sr * 0.50 + sl * 0.30 + skw * 0.20)

    # Rationale
    parts = []
    if prix_ref_val > 0:
        parts.append(f"prix marché ~{prix_ref_val}€ pour {prix}€ (x{round(ratio,1)})")
    if marge > 0:
        parts.append(f"marge estimée {int(marge*100)}%")
    if nb_ventes > 0:
        parts.append(f"{nb_ventes} ventes similaires récentes")
    if value_keywords:
        parts.append(f"mots-clés valeur : {', '.join(value_keywords[:3])}")

    return {
        **annonce,
        "univers":         univers,
        "prix_marche_ref": prix_ref_val,
        "marge_estimee":   round(marge * 100, 1),
        "ratio_x":         round(ratio, 1),
        "score_global":    score_global,
        "score_ratio":     sr,
        "score_liquidite": sl,
        "score_valeur":    skw,
        "nb_ventes_ref":   nb_ventes,
        "mots_cles_valeur": ", ".join(value_keywords),
        "rationale":       " · ".join(parts) if parts else "Annonce détectée — données limitées",
    }

# ─── Email ─────────────────────────────────────────────────────────────────────

def send_alert(opps: list[dict]):
    if not (GMAIL_FROM and GMAIL_TO and GMAIL_PASS):
        log.warning("Gmail non configuré")
        return
    top  = sorted(opps, key=lambda x: x["score_global"], reverse=True)[:6]
    rows = "".join(f"""
        <tr style="border-bottom:1px solid #eee">
          <td style="padding:8px">
            <strong>{o['titre'][:60]}</strong><br>
            <span style="font-size:11px;background:{'#EAF3DE' if o['univers']=='DBZ' else '#E8F0FB'};
                         padding:1px 6px;border-radius:3px;color:#333">
              {o['univers']}
            </span>
          </td>
          <td style="padding:8px;text-align:center">
            <span style="background:#EAF3DE;color:#27500A;padding:2px 8px;
                         border-radius:10px;font-weight:bold">
              {o['score_global']}/100
            </span>
          </td>
          <td style="padding:8px;text-align:center;font-weight:bold;color:#185FA5">
            x{o['ratio_x']}
          </td>
          <td style="padding:8px;text-align:center">
            <strong>{o['prix']}€</strong>
            <span style="color:#888;font-size:11px"> → ~{o['prix_marche_ref']}€</span>
          </td>
          <td style="padding:8px;font-size:11px;color:#666">
            {o['rationale'][:100]}
          </td>
          <td style="padding:8px">
            <a href="{o['lien_annonce']}" style="color:#185FA5">Voir →</a>
          </td>
        </tr>""" for o in top)

    html = f"""<html><body style="font-family:Arial,sans-serif;max-width:800px;margin:auto">
      <h2 style="color:#1a1a1a">Card Agent — Opportunités détectées</h2>
      <p style="color:#666">{datetime.now().strftime('%d/%m/%Y à %Hh%M')}</p>
      <table style="width:100%;border-collapse:collapse;border:1px solid #eee">
        <thead><tr style="background:#f5f5f5">
          <th style="padding:8px;text-align:left">Article</th>
          <th style="padding:8px">Score</th>
          <th style="padding:8px">Ratio</th>
          <th style="padding:8px">Prix → Marché</th>
          <th style="padding:8px;text-align:left">Rationale</th>
          <th style="padding:8px">Lien</th>
        </tr></thead>
        <tbody>{rows}</tbody>
      </table>
      <p style="color:#aaa;font-size:11px;margin-top:16px">Card Agent v1 · Scan automatique</p>
    </body></html>"""

    msg = MIMEMultipart("alternative")
    msg["Subject"] = (
        f"[Card Agent] {len(top)} opportunité(s) DBZ/YGO"
        f" — {datetime.now().strftime('%d/%m %Hh')}"
    )
    msg["From"] = GMAIL_FROM
    msg["To"]   = GMAIL_TO
    msg.attach(MIMEText(html, "html"))
    try:
        with smtplib.SMTP_SSL("smtp.gmail.com", 465) as s:
            s.login(GMAIL_FROM, GMAIL_PASS)
            s.send_message(msg)
        log.info(f"Email envoyé → {GMAIL_TO}")
    except Exception as e:
        log.error(f"Email error: {e}")

# ─── Main ─────────────────────────────────────────────────────────────────────

def run():
    log.info("=== Card Agent v1 ===")
    sh = get_sheet()
    ensure_sheets(sh)
    ws_opp = sh.worksheet("opportunites")
    ws_ref = sh.worksheet("prix_reference")

    all_scored: list[dict] = []
    seen_urls: set[str]    = set()
    prix_ref_cache: dict   = {}
    now = datetime.now().strftime("%Y-%m-%d %H:%M")

    # ── DBZ — recherche + scoring ─────────────────────────────────────────────
    log.info("=== Scan DBZ ===")
    dbz_queries = random.sample(DBZ_SEARCH_TERMS, min(4, len(DBZ_SEARCH_TERMS)))

    for query in dbz_queries:
        annonces = search_vinted(query)
        time.sleep(1)

        for ann in annonces:
            if ann["lien"] in seen_urls:
                continue
            seen_urls.add(ann["lien"])

            ann["univers"] = "DBZ"
            value_kw = detect_value_keywords(
                ann["titre"] + " " + query, DBZ_HIGH_VALUE_KEYWORDS
            )

            # Prix de référence eBay — avec cache
            ref_query = "carte dragon ball z " + " ".join(value_kw[:2]) if value_kw else query
            if ref_query not in prix_ref_cache:
                prix_ref_cache[ref_query] = get_ebay_sold_prices(ref_query)
                # Sauvegarde dans Sheet
                pref = prix_ref_cache[ref_query]
                if pref:
                    ws_ref.append_row([
                        now, "DBZ", ref_query,
                        pref.get("nb_ventes", 0),
                        pref.get("prix_moyen", 0),
                        pref.get("prix_median", 0),
                        pref.get("prix_min", 0),
                        pref.get("prix_max", 0),
                    ])
                time.sleep(0.5)

            scored = score_annonce(ann, prix_ref_cache.get(ref_query, {}), value_kw)
            scored["lien_annonce"] = ann["lien"]
            all_scored.append(scored)

        time.sleep(1.5)

    # ── YGO — recherche + scoring ─────────────────────────────────────────────
    log.info("=== Scan YGO ===")
    ygo_queries = random.sample(YGO_SEARCH_TERMS, min(3, len(YGO_SEARCH_TERMS)))

    for query in ygo_queries:
        annonces = search_vinted(query)
        time.sleep(1)

        for ann in annonces:
            if ann["lien"] in seen_urls:
                continue
            seen_urls.add(ann["lien"])

            ann["univers"] = "YGO"
            value_kw = detect_value_keywords(
                ann["titre"] + " " + query, YGO_HIGH_VALUE_KEYWORDS
            )

            # Prix de référence — eBay pour YGO également
            # (Cardmarket nécessite OAuth, eBay suffit pour détection d'opportunités)
            ref_query = "yugioh " + " ".join(value_kw[:2]) if value_kw else query
            if ref_query not in prix_ref_cache:
                prix_ref_cache[ref_query] = get_ebay_sold_prices(ref_query)
                pref = prix_ref_cache[ref_query]
                if pref:
                    ws_ref.append_row([
                        now, "YGO", ref_query,
                        pref.get("nb_ventes", 0),
                        pref.get("prix_moyen", 0),
                        pref.get("prix_median", 0),
                        pref.get("prix_min", 0),
                        pref.get("prix_max", 0),
                    ])
                time.sleep(0.5)

            scored = score_annonce(ann, prix_ref_cache.get(ref_query, {}), value_kw)
            scored["lien_annonce"] = ann["lien"]
            all_scored.append(scored)

        time.sleep(1.5)

    # ── Écriture Sheet ────────────────────────────────────────────────────────
    all_scored.sort(key=lambda x: x["score_global"], reverse=True)

    # Filtre : garde uniquement les annonces avec une marge > 0 ou des mots-clés valeur
    filtered = [
        o for o in all_scored
        if o["marge_estimee"] > 0 or o["mots_cles_valeur"]
    ]

    log.info(f"Total scorés : {len(all_scored)} · Filtrés : {len(filtered)}")

    rows = []
    for o in filtered[:60]:
        rows.append([
            now,
            o.get("univers", ""),
            o.get("titre", "")[:120],
            o.get("prix", 0),
            o.get("prix_marche_ref", 0),
            o.get("marge_estimee", 0),
            o.get("ratio_x", 0),
            o.get("score_global", 0),
            o.get("score_ratio", 0),
            o.get("score_liquidite", 0),
            o.get("score_valeur", 0),
            o.get("nb_ventes_ref", 0),
            o.get("mots_cles_valeur", ""),
            o.get("plateforme", "vinted"),
            o.get("etat", ""),
            o.get("vendeur", ""),
            o.get("rationale", ""),
            o.get("lien_annonce", ""),
        ])

    if rows:
        ws_opp.append_rows(rows, value_input_option="RAW")
        log.info(f"{len(rows)} opportunités écrites")
    else:
        log.warning("Aucune opportunité trouvée")

    # ── Alertes ───────────────────────────────────────────────────────────────
    alerts = [o for o in filtered if o["score_global"] >= SCORE_ALERT]
    if alerts:
        send_alert(alerts)
        log.info(f"{len(alerts)} alertes envoyées")
    else:
        log.info("Aucune alerte — seuil non atteint")

    log.info("=== Scan terminé v1 ===")

if __name__ == "__main__":
    run()
