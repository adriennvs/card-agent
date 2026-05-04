"""
Card Opportunity Agent — v5
Fix : 1 appel eBay par terme de recherche (pas par annonce)
      Requêtes eBay en français + siteid=71 (eBay France)
"""

import os
import time
import json
import random
import logging
import requests
import gspread
import smtplib
from datetime import datetime
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from google.oauth2.service_account import Credentials

logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(message)s")
log = logging.getLogger(__name__)

# ─── Configuration ────────────────────────────────────────────────────────────

SCORE_ALERT = 70

SHEET_NAME       = os.getenv("GOOGLE_SHEET_NAME", "Card Agent")
GMAIL_FROM       = os.getenv("GMAIL_FROM")
GMAIL_TO         = os.getenv("GMAIL_TO")
GMAIL_PASS       = os.getenv("GMAIL_APP_PASSWORD")
GCP_CREDS        = os.getenv("GCP_SERVICE_ACCOUNT_JSON")
SCRAPEBADGER_KEY = os.getenv("SCRAPEBADGER_API_KEY", "")
EBAY_APP_ID      = os.getenv("EBAY_APP_ID", "")

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/120.0.0.0 Safari/537.36"
    )
}

# ─── Mots-clés ────────────────────────────────────────────────────────────────

# Terme Vinted → terme eBay France correspondant
DBZ_QUERIES = {
    "le grand combat dragon ball":    "dragon ball le grand combat carte",
    "hondan dragon ball":             "dragon ball hondan carddass",
    "carddass dragon ball":           "dragon ball carddass carte",
    "super battle dragon ball":       "dragon ball super battle carte",
    "visual adventure dragon ball":   "dragon ball visual adventure carte",
    "cartes dbz":                     "dragon ball z carte française",
}

YGO_QUERIES = {
    "yugioh cartes":   "yu-gi-oh carte française",
    "yu gi oh lot":    "yu-gi-oh lot cartes",
    "cartes yugioh":   "yu-gi-oh carte française",
}

DBZ_HIGH_VALUE = [
    "prisme", "prism", "flash", "hors série", "hors-série",
    "hondan", "grand combat", "super battle", "power level",
    "visual adventure", "carddass", "rare", "special",
    "goku", "vegeta", "freezer", "cell", "buu", "gohan",
    "gotenks", "broly", "gogeta", "vegeto", "sangoku",
]

YGO_HIGH_VALUE = [
    "1ère édition", "1st edition", "first edition",
    "ultra rare", "secret rare", "ghost rare",
    "blue eyes", "dark magician", "exodia",
    "limited edition", "lob", "mrd", "mrl",
    "metal raiders", "holo", "holographique", "mirror force",
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

def search_vinted(query: str) -> list[dict]:
    if not SCRAPEBADGER_KEY:
        log.warning("SCRAPEBADGER_API_KEY non configurée")
        return []
    try:
        r = requests.get(
            "https://scrapebadger.com/v1/vinted/search",
            headers={
                "x-api-key": SCRAPEBADGER_KEY,
                "Accept":    "application/json",
            },
            params={
                "query":    query,
                "market":   "fr",
                "per_page": 50,
                "order":    "newest_first",
                "time":     "1d",
            },
            timeout=30,
        )
        if r.status_code == 429:
            log.warning("Rate limit ScrapeBadger — attente 60s")
            time.sleep(60)
            return []
        if r.status_code not in (200, 201):
            log.warning(f"ScrapeBadger status {r.status_code} pour '{query}'")
            return []

        items_raw = r.json().get("items", [])
        log.info(f"Vinted '{query}' → {len(items_raw)} annonces")

        items = []
        for item in items_raw:
            try:
                price_raw = item.get("price", {})
                if isinstance(price_raw, dict):
                    price = float(price_raw.get("amount", 0) or 0)
                else:
                    price = float(price_raw or 0)
            except Exception:
                continue
            if price <= 0:
                continue
            item_id = item.get("id", "")
            user    = item.get("user", {})
            items.append({
                "titre":      str(item.get("title", "")).strip(),
                "prix":       price,
                "etat":       str(item.get("status", item.get("condition", ""))),
                "vendeur":    str(user.get("login", user.get("username", ""))),
                "lien":       item.get("url", f"https://www.vinted.fr/items/{item_id}"),
                "plateforme": "vinted",
            })
        return items

    except Exception as e:
        log.warning(f"ScrapeBadger error ('{query}'): {e}")
        return []

# ─── Source 2 : eBay sold listings — 1 appel par terme ────────────────────────

def get_ebay_sold_prices(ebay_query: str) -> dict:
    """
    1 appel eBay par terme de recherche Vinted.
    siteid=71 → eBay France uniquement → prix en euros français.
    """
    if not EBAY_APP_ID:
        return {}
    try:
        r = requests.get(
            "https://svcs.ebay.com/services/search/FindingService/v1",
            params={
                "OPERATION-NAME":                "findCompletedItems",
                "SERVICE-VERSION":               "1.0.0",
                "SECURITY-APPNAME":              EBAY_APP_ID,
                "RESPONSE-DATA-FORMAT":          "JSON",
                "keywords":                      ebay_query,
                "siteid":                        "71",
                "itemFilter(0).name":            "SoldItemsOnly",
                "itemFilter(0).value":           "true",
                "itemFilter(1).name":            "Currency",
                "itemFilter(1).value":           "EUR",
                "sortOrder":                     "EndTimeSoonest",
                "paginationInput.entriesPerPage": "50",
            },
            headers=HEADERS,
            timeout=15,
        )
        log.info(f"eBay '{ebay_query}' → status {r.status_code}")

        if r.status_code != 200:
            return {}

        items_data = (
            r.json()
             .get("findCompletedItemsResponse", [{}])[0]
             .get("searchResult", [{}])[0]
             .get("item", [])
        )
        prices = []
        for item in items_data:
            try:
                p = float(item["sellingStatus"][0]["currentPrice"][0]["__value__"])
                if 0 < p < 1000:
                    prices.append(p)
            except Exception:
                pass

        log.info(f"eBay '{ebay_query}' → {len(prices)} prix extraits")
        if not prices:
            return {}

        prices.sort()
        n = len(prices)
        return {
            "nb_ventes":   n,
            "prix_moyen":  round(sum(prices) / n, 2),
            "prix_median": round(prices[n // 2], 2),
            "prix_min":    round(prices[0], 2),
            "prix_max":    round(prices[-1], 2),
        }
    except Exception as e:
        log.warning(f"eBay error ('{ebay_query}'): {e}")
        return {}

# ─── Détection mots-clés ──────────────────────────────────────────────────────

def detect_keywords(text: str, keywords: list) -> list:
    text_lower = text.lower()
    return [kw for kw in keywords if kw.lower() in text_lower]

# ─── Scoring ──────────────────────────────────────────────────────────────────

def score_annonce(annonce: dict, prix_ref: dict, value_kw: list) -> dict:
    prix         = annonce["prix"]
    prix_ref_val = prix_ref.get("prix_median") or prix_ref.get("prix_moyen") or 0
    nb_ventes    = prix_ref.get("nb_ventes", 0)

    marge = (prix_ref_val - prix) / max(prix_ref_val, 1) if prix_ref_val > 0 else 0
    ratio = prix_ref_val / max(prix, 1) if prix_ref_val > 0 else 1

    if ratio >= 5:     sr = 100
    elif ratio >= 3:   sr = 85
    elif ratio >= 2:   sr = 70
    elif ratio >= 1.5: sr = 55
    elif ratio >= 1.2: sr = 35
    else:              sr = 10

    if nb_ventes >= 30:   sl = 100
    elif nb_ventes >= 15: sl = 80
    elif nb_ventes >= 8:  sl = 60
    elif nb_ventes >= 3:  sl = 40
    elif nb_ventes >= 1:  sl = 20
    else:                 sl = 0

    n = len(value_kw)
    if n >= 4:   skw = 100
    elif n >= 3: skw = 80
    elif n >= 2: skw = 60
    elif n >= 1: skw = 40
    else:        skw = 10

    score_global = int(sr * 0.50 + sl * 0.30 + skw * 0.20)

    parts = []
    if prix_ref_val > 0:
        parts.append(f"prix marché eBay.fr ~{prix_ref_val}€ pour {prix}€ (x{round(ratio,1)})")
    if marge > 0:
        parts.append(f"marge estimée {int(marge*100)}%")
    if nb_ventes > 0:
        parts.append(f"{nb_ventes} ventes similaires eBay.fr")
    if value_kw:
        parts.append(f"mots-clés : {', '.join(value_kw[:3])}")

    return {
        **annonce,
        "prix_marche_ref":  prix_ref_val,
        "marge_estimee":    round(marge * 100, 1),
        "ratio_x":          round(ratio, 1),
        "score_global":     score_global,
        "score_ratio":      sr,
        "score_liquidite":  sl,
        "score_valeur":     skw,
        "nb_ventes_ref":    nb_ventes,
        "mots_cles_valeur": ", ".join(value_kw),
        "rationale":        " · ".join(parts) if parts else "Annonce détectée — données limitées",
        "lien_annonce":     annonce["lien"],
    }

# ─── Email ─────────────────────────────────────────────────────────────────────

def send_alert(opps: list):
    if not (GMAIL_FROM and GMAIL_TO and GMAIL_PASS):
        return
    top  = sorted(opps, key=lambda x: x["score_global"], reverse=True)[:6]
    rows = "".join(f"""
        <tr>
          <td style="padding:8px;border-bottom:1px solid #eee">
            <strong>{o['titre'][:60]}</strong><br>
            <span style="font-size:11px;background:{'#FEF3E6' if o.get('univers')=='DBZ' else '#E8F0FB'};
                         padding:1px 6px;border-radius:3px">{o.get('univers','')}</span>
          </td>
          <td style="padding:8px;text-align:center">
            <span style="background:#EAF3DE;color:#27500A;padding:2px 8px;
                         border-radius:10px;font-weight:bold">{o['score_global']}/100</span>
          </td>
          <td style="padding:8px;text-align:center;font-weight:bold">x{o['ratio_x']}</td>
          <td style="padding:8px;text-align:center">
            <strong>{o['prix']}€</strong>
            <span style="color:#888"> → ~{o['prix_marche_ref']}€</span>
          </td>
          <td style="padding:8px;font-size:11px;color:#666">{o['rationale'][:100]}</td>
          <td style="padding:8px"><a href="{o['lien_annonce']}">Voir →</a></td>
        </tr>""" for o in top)

    html = f"""<html><body style="font-family:Arial,sans-serif;max-width:800px;margin:auto">
      <h2>Card Agent — {datetime.now().strftime('%d/%m/%Y %Hh%M')}</h2>
      <table style="width:100%;border-collapse:collapse;border:1px solid #eee">
        <thead><tr style="background:#f5f5f5">
          <th style="padding:8px;text-align:left">Article</th>
          <th style="padding:8px">Score</th><th style="padding:8px">Ratio</th>
          <th style="padding:8px">Prix → Marché</th>
          <th style="padding:8px;text-align:left">Rationale</th>
          <th style="padding:8px">Lien</th>
        </tr></thead><tbody>{rows}</tbody>
      </table>
    </body></html>"""

    msg = MIMEMultipart("alternative")
    msg["Subject"] = f"[Card Agent] {len(top)} opportunité(s) — {datetime.now().strftime('%d/%m %Hh')}"
    msg["From"]    = GMAIL_FROM
    msg["To"]      = GMAIL_TO
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
    log.info("=== Card Agent v5 ===")
    sh = get_sheet()
    ensure_sheets(sh)
    ws_opp = sh.worksheet("opportunites")
    ws_ref = sh.worksheet("prix_reference")

    # Vide les données en gardant les entêtes
    ws_opp.clear()
    ws_opp.append_row([
        "date_scan", "univers", "titre", "prix_annonce",
        "prix_marche_ref", "marge_estimee", "ratio_x",
        "score_global", "score_ratio", "score_liquidite", "score_valeur",
        "nb_ventes_ref", "mots_cles_valeur",
        "plateforme", "etat", "vendeur",
        "rationale", "lien_annonce"
    ])

    all_scored: list[dict] = []
    seen_urls: set[str]    = set()
    now = datetime.now().strftime("%Y-%m-%d %H:%M")

    # ── DBZ ───────────────────────────────────────────────────────────────────
    log.info("=== Scan DBZ ===")
    dbz_items = list(DBZ_QUERIES.items())
    random.shuffle(dbz_items)

    for vinted_query, ebay_query in dbz_items[:3]:
        # 1 appel Vinted
        annonces = search_vinted(vinted_query)
        time.sleep(12)

        # 1 appel eBay par terme (pas par annonce)
        prix_ref = get_ebay_sold_prices(ebay_query)
        if prix_ref:
            ws_ref.append_row([
                now, "DBZ", ebay_query,
                prix_ref.get("nb_ventes", 0),
                prix_ref.get("prix_moyen", 0),
                prix_ref.get("prix_median", 0),
                prix_ref.get("prix_min", 0),
                prix_ref.get("prix_max", 0),
            ])
        time.sleep(3)

        # Score toutes les annonces avec ce prix de référence
        for ann in annonces:
            if ann["lien"] in seen_urls:
                continue
            seen_urls.add(ann["lien"])
            ann["univers"] = "DBZ"
            value_kw = detect_keywords(ann["titre"] + " " + vinted_query, DBZ_HIGH_VALUE)
            scored   = score_annonce(ann, prix_ref, value_kw)
            all_scored.append(scored)

    # ── YGO ───────────────────────────────────────────────────────────────────
    log.info("=== Scan YGO ===")
    ygo_items = list(YGO_QUERIES.items())
    random.shuffle(ygo_items)

    for vinted_query, ebay_query in ygo_items[:2]:
        annonces = search_vinted(vinted_query)
        time.sleep(12)

        prix_ref = get_ebay_sold_prices(ebay_query)
        if prix_ref:
            ws_ref.append_row([
                now, "YGO", ebay_query,
                prix_ref.get("nb_ventes", 0),
                prix_ref.get("prix_moyen", 0),
                prix_ref.get("prix_median", 0),
                prix_ref.get("prix_min", 0),
                prix_ref.get("prix_max", 0),
            ])
        time.sleep(3)

        for ann in annonces:
            if ann["lien"] in seen_urls:
                continue
            seen_urls.add(ann["lien"])
            ann["univers"] = "YGO"
            value_kw = detect_keywords(ann["titre"] + " " + vinted_query, YGO_HIGH_VALUE)
            scored   = score_annonce(ann, prix_ref, value_kw)
            all_scored.append(scored)

    # ── Écriture Sheet ────────────────────────────────────────────────────────
    all_scored.sort(key=lambda x: x["score_global"], reverse=True)
    filtered = [o for o in all_scored if o["marge_estimee"] > 0 or o["mots_cles_valeur"]]
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

    alerts = [o for o in filtered if o["score_global"] >= SCORE_ALERT]
    if alerts:
        send_alert(alerts)
        log.info(f"{len(alerts)} alertes envoyées")
    else:
        log.info("Aucune alerte — seuil non atteint")

    log.info("=== Scan terminé v5 ===")

if __name__ == "__main__":
    run()
