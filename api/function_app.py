"""
Komda Demo-Portal – Azure Function App
=======================================
Bestehende Endpunkte (aus komda-onboarding):
  GET/POST /api/status   → unverändert übernehmen

Neue Endpunkte (Demo-Verwaltung):
  GET  /api/demo               → Token  → Interessenten-Daten (öffentlich)
  GET  /api/demo               → kein Token, X-Admin-Key → alle Einträge (Management)
  POST /api/demo               → Interessenten anlegen / aktualisieren (Management)
  POST /api/demo-track         → Link-Öffnung aufzeichnen + interne Benachrichtigung
  POST /api/demo-mail          → Demo-Mail an Interessenten versenden (Management)

Env-Variablen (zusätzlich zu bestehenden):
  DEMO_LIST_ID     – SharePoint-Listen-ID für „DemoInteressenten"
  DEMO_ADMIN_KEY   – Shared Secret für Management-Endpunkte
  DEMO_BASE_URL    – Basis-URL der Interessenten-Seite
                     (z. B. https://testen.komda-software.de)
  NOTIFY_FROM_EMAIL – bereits vorhanden (onboarding@komda-software.de)

SharePoint-Liste „DemoInteressenten" – Felder:
  Title (= Name), Firma, Email, Produkt, Token,
  MailGesendet (bool), LinkGeoeffnet (bool), LinkGeoeffnetAm (DateTime),
  SachbearbeiterEmail, Notizen, Aktiv (bool)
"""

import azure.functions as func
import logging
import json
import os
import base64
import requests
from datetime import datetime, timezone

app = func.FunctionApp()

# ---------------------------------------------------------------------------
# CORS
# ---------------------------------------------------------------------------

CORS_HEADERS = {
    "Access-Control-Allow-Origin":  "https://demo.komda-software.de",
    "Access-Control-Allow-Methods": "GET, POST, OPTIONS",
    "Access-Control-Allow-Headers": "Content-Type, X-Admin-Key",
}


def json_resp(data, status_code: int = 200) -> func.HttpResponse:
    """Erstellt eine JSON-Response mit CORS-Headern."""
    return func.HttpResponse(
        json.dumps(data, ensure_ascii=False),
        status_code=status_code,
        mimetype="application/json",
        headers=CORS_HEADERS,
    )


def options_resp() -> func.HttpResponse:
    """Antwortet auf CORS-Preflight-Anfragen."""
    return func.HttpResponse(status_code=204, headers=CORS_HEADERS)


# ---------------------------------------------------------------------------
# Hilfsfunktionen
# ---------------------------------------------------------------------------

def get_app_token() -> str:
    """Holt ein App-Access-Token via Client Credentials."""
    tenant_id = os.environ["TENANT_ID"]
    url = f"https://login.microsoftonline.com/{tenant_id}/oauth2/v2.0/token"
    data = {
        "grant_type":    "client_credentials",
        "client_id":     os.environ["CLIENT_ID"],
        "client_secret": os.environ["CLIENT_SECRET"],
        "scope":         "https://graph.microsoft.com/.default",
    }
    resp = requests.post(url, data=data, timeout=15)
    resp.raise_for_status()
    return resp.json()["access_token"]


def graph_headers(token: str) -> dict:
    return {"Authorization": f"Bearer {token}", "Content-Type": "application/json"}


def get_site_id() -> str:
    return os.environ["SITE_ID"]


def get_demo_list_id() -> str:
    return os.environ["DEMO_LIST_ID"]


def require_admin(req: func.HttpRequest) -> bool:
    """Prüft X-Admin-Key Header gegen Umgebungsvariable."""
    expected = os.environ.get("DEMO_ADMIN_KEY", "")
    provided = req.headers.get("X-Admin-Key", "")
    return bool(expected and provided == expected)


# ---------------------------------------------------------------------------
# Token-Hilfsfunktionen
# ---------------------------------------------------------------------------

def encode_token(item_id: int) -> str:
    """Enkodiert die SharePoint-Item-ID als URL-sicheres Base64."""
    return base64.urlsafe_b64encode(str(item_id).encode()).decode()


def decode_token(token: str) -> int | None:
    """Dekodiert den Token zurück zur Item-ID. None bei ungültigem Token."""
    try:
        decoded = base64.urlsafe_b64decode(token.encode()).decode()
        return int(decoded)
    except Exception:
        return None


# ---------------------------------------------------------------------------
# SharePoint-Hilfsfunktionen
# ---------------------------------------------------------------------------

def sp_base(token: str) -> str:
    site_id = get_site_id()
    list_id = get_demo_list_id()
    return f"https://graph.microsoft.com/v1.0/sites/{site_id}/lists/{list_id}"


def get_item_by_id(token: str, item_id: int) -> dict | None:
    url = f"{sp_base(token)}/items/{item_id}?$expand=fields"
    resp = requests.get(url, headers=graph_headers(token), timeout=15)
    if resp.status_code == 404:
        return None
    resp.raise_for_status()
    return resp.json().get("fields", {})


def get_all_items(token: str) -> list[dict]:
    url = (
        f"{sp_base(token)}/items"
        "?$expand=fields"
        "&$select=id,fields"
        "&$orderby=fields/Created desc"
        "&$top=500"
    )
    resp = requests.get(url, headers=graph_headers(token), timeout=15)
    resp.raise_for_status()
    items = resp.json().get("value", [])
    result = []
    for item in items:
        f = item.get("fields", {})
        f["_itemId"] = item.get("id")
        result.append(f)
    return result


def create_item(token: str, fields: dict) -> dict:
    url = f"{sp_base(token)}/items"
    body = {"fields": fields}
    resp = requests.post(url, headers=graph_headers(token),
                         json=body, timeout=15)
    resp.raise_for_status()
    return resp.json()


def update_item(token: str, item_id: int, fields: dict) -> None:
    url = f"{sp_base(token)}/items/{item_id}/fields"
    resp = requests.patch(url, headers=graph_headers(token),
                          json=fields, timeout=15)
    resp.raise_for_status()


# ---------------------------------------------------------------------------
# E-Mail-Versand via Graph API
# ---------------------------------------------------------------------------

MOBILE_DATA = {
    "Stationär": {
        "kundennummer": "STAT-DEMO",
        "passwort":     "komda2024",
        "pin":          "1234",
        "android_url":  "https://play.google.com/store/apps/details?id=de.komda.stationaer",
        "ios_url":      "https://apps.apple.com/de/app/komda-stationaer/id000000001",
    },
    "Teilstationär": {
        "kundennummer": "TEIL-DEMO",
        "passwort":     "komda2024",
        "pin":          "5678",
        "android_url":  "https://play.google.com/store/apps/details?id=de.komda.teilstationaer",
        "ios_url":      "https://apps.apple.com/de/app/komda-teilstationaer/id000000002",
    },
    "Ambulant": {
        "kundennummer": "AMBU-DEMO",
        "passwort":     "komda2024",
        "pin":          "9012",
        "android_url":  "https://play.google.com/store/apps/details?id=de.komda.ambulant",
        "ios_url":      "https://apps.apple.com/de/app/komda-ambulant/id000000003",
    },
    "Betreuung": {
        "kundennummer": "BETR-DEMO",
        "passwort":     "komda2024",
        "pin":          "3456",
        "android_url":  "https://play.google.com/store/apps/details?id=de.komda.betreuung",
        "ios_url":      "https://apps.apple.com/de/app/komda-betreuung/id000000004",
    },
}

RDP_USERNAME = "u37009-02"
RDP_PASSWORD = "Komda222"


def build_demo_email_html(name: str, firma: str, produkt: str,
                           token: str, sachbearbeiter: str) -> str:
    base_url = os.environ.get("DEMO_BASE_URL", "https://testen.komda-software.de")
    zugurl = f"{base_url}/zugang.html?token={token}"
    mobile = MOBILE_DATA.get(produkt, {})

    return f"""<!DOCTYPE html>
<html lang="de">
<head><meta charset="UTF-8"></head>
<body style="font-family:Arial,sans-serif;max-width:600px;margin:0 auto;color:#333;">
  <div style="background:#0055a5;padding:24px;border-radius:8px 8px 0 0;">
    <h1 style="color:#fff;margin:0;font-size:22px;">Komda® Software – Ihr persönlicher Demo-Zugang</h1>
  </div>
  <div style="background:#f8f9fa;padding:24px;border-radius:0 0 8px 8px;border:1px solid #dee2e6;">

    <p>Sehr geehrte/r {name},</p>
    <p>vielen Dank für Ihr Interesse an <strong>Komda® {produkt}</strong>
       von <strong>{firma}</strong>.</p>
    <p>Wir haben für Sie einen persönlichen Demo-Zugang eingerichtet.
       Klicken Sie auf den folgenden Link, um direkt zu starten:</p>

    <div style="text-align:center;margin:32px 0;">
      <a href="{zugurl}"
         style="background:#0055a5;color:#fff;padding:14px 32px;
                border-radius:6px;text-decoration:none;font-size:16px;
                font-weight:bold;display:inline-block;">
        Demo starten →
      </a>
    </div>

    <hr style="border:none;border-top:1px solid #dee2e6;margin:24px 0;">

    <h3 style="color:#0055a5;">Desktop-Zugang (Remote Desktop)</h3>
    <table style="border-collapse:collapse;width:100%;">
      <tr>
        <td style="padding:8px;color:#666;">Benutzername:</td>
        <td style="padding:8px;font-family:monospace;font-weight:bold;">{RDP_USERNAME}</td>
      </tr>
      <tr style="background:#fff;">
        <td style="padding:8px;color:#666;">Kennwort:</td>
        <td style="padding:8px;font-family:monospace;font-weight:bold;">{RDP_PASSWORD}</td>
      </tr>
    </table>
    <p style="font-size:13px;color:#666;">
      Die RDP-Datei und ausführliche Schritt-für-Schritt-Anleitungen
      finden Sie auf Ihrer persönlichen Demo-Seite.
    </p>

    {"" if not mobile else f'''
    <h3 style="color:#0055a5;margin-top:24px;">Komda® Mobile App</h3>
    <table style="border-collapse:collapse;width:100%;">
      <tr>
        <td style="padding:8px;color:#666;">Kundennummer:</td>
        <td style="padding:8px;font-family:monospace;font-weight:bold;">{mobile.get("kundennummer","")}</td>
      </tr>
      <tr style="background:#fff;">
        <td style="padding:8px;color:#666;">Passwort:</td>
        <td style="padding:8px;font-family:monospace;font-weight:bold;">{mobile.get("passwort","")}</td>
      </tr>
      <tr>
        <td style="padding:8px;color:#666;">PIN:</td>
        <td style="padding:8px;font-family:monospace;font-weight:bold;">{mobile.get("pin","")}</td>
      </tr>
    </table>
    <p style="margin-top:12px;">
      <a href="{mobile.get("android_url","#")}" style="color:#0055a5;">▶ Android (Play Store)</a>
      &nbsp;&nbsp;
      <a href="{mobile.get("ios_url","#")}" style="color:#0055a5;">▶ iOS (App Store)</a>
    </p>
    '''}

    <hr style="border:none;border-top:1px solid #dee2e6;margin:24px 0;">
    <p style="font-size:13px;color:#888;">
      Bei Fragen steht Ihnen Ihr Ansprechpartner gerne zur Verfügung.<br>
      Dieser Link ist persönlich und sollte nicht weitergegeben werden.
    </p>
    <p style="font-size:13px;color:#888;">
      Mit freundlichen Grüßen<br>
      <strong>Komda® Software GmbH</strong>
    </p>
  </div>
</body>
</html>"""


def send_graph_email(to_email: str, subject: str, html_body: str,
                     graph_token: str) -> None:
    sender = os.environ.get("NOTIFY_FROM_EMAIL", "onboarding@komda-software.de")
    url = (f"https://graph.microsoft.com/v1.0"
           f"/users/{sender}/sendMail")
    payload = {
        "message": {
            "subject": subject,
            "body":    {"contentType": "HTML", "content": html_body},
            "toRecipients": [{"emailAddress": {"address": to_email}}],
        }
    }
    resp = requests.post(url, headers=graph_headers(graph_token),
                         json=payload, timeout=20)
    resp.raise_for_status()


def send_internal_notification(sachbearbeiter_email: str, name: str,
                                firma: str, graph_token: str) -> None:
    """Benachrichtigt Sachbearbeiter, wenn Interessent Link öffnet."""
    subject = f"Demo-Link geöffnet: {name} ({firma})"
    html = f"""<p>Der Demo-Link wurde soeben geöffnet:</p>
<ul>
  <li><strong>Name:</strong> {name}</li>
  <li><strong>Firma:</strong> {firma}</li>
  <li><strong>Zeitpunkt:</strong> {datetime.now(timezone.utc).strftime('%d.%m.%Y %H:%M')} UTC</li>
</ul>"""
    try:
        send_graph_email(sachbearbeiter_email, subject, html, graph_token)
    except Exception as e:
        logging.warning(f"Interne Benachrichtigung fehlgeschlagen: {e}")


# ---------------------------------------------------------------------------
# Endpunkt: GET/POST /api/demo
# ---------------------------------------------------------------------------

@app.route(route="demo", methods=["GET", "POST", "OPTIONS"],
           auth_level=func.AuthLevel.ANONYMOUS)
def demo(req: func.HttpRequest) -> func.HttpResponse:
    logging.info("demo: %s", req.method)

    if req.method == "OPTIONS":
        return options_resp()

    # ------------------------------------------------------------------
    # GET – Interessenten-Daten per Token (öffentlich)
    #      oder alle Einträge per Admin-Key (Management)
    # ------------------------------------------------------------------
    if req.method == "GET":
        token = req.params.get("token")

        # Alle Einträge für Management-Portal
        if not token:
            if not require_admin(req):
                return json_resp({"error": "Unauthorized"}, 401)
            try:
                graph_token = get_app_token()
                items = get_all_items(graph_token)
                for item in items:
                    iid = item.get("_itemId")
                    if iid:
                        item["Token"] = encode_token(int(iid))
                return json_resp(items)
            except Exception as e:
                logging.exception("demo GET all error")
                return json_resp({"error": str(e)}, 500)

        # Einzelner Eintrag per Token (für Interessenten-Seite)
        item_id = decode_token(token)
        if item_id is None:
            return json_resp({"error": "Ungültiger Token"}, 400)
        try:
            graph_token = get_app_token()
            fields = get_item_by_id(graph_token, item_id)
            if fields is None:
                return json_resp({"error": "Interessent nicht gefunden"}, 404)
            if not fields.get("Aktiv", True):
                return json_resp({"error": "Demo-Zugang wurde deaktiviert"}, 410)

            public = {
                "name":    fields.get("Title", ""),
                "firma":   fields.get("Firma", ""),
                "produkt": fields.get("Produkt", ""),
                "aktiv":   fields.get("Aktiv", True),
            }
            public["mobile"] = MOBILE_DATA.get(public["produkt"], {})
            return json_resp(public)
        except Exception as e:
            logging.exception("demo GET single error")
            return json_resp({"error": str(e)}, 500)

    # ------------------------------------------------------------------
    # POST – Interessenten anlegen oder aktualisieren (Management)
    # ------------------------------------------------------------------
    if req.method == "POST":
        if not require_admin(req):
            return json_resp({"error": "Unauthorized"}, 401)
        try:
            body = req.get_json()
        except Exception:
            return json_resp({"error": "Kein gültiges JSON"}, 400)

        action = body.get("action", "create")

        try:
            graph_token = get_app_token()

            if action == "create":
                required = ["name", "firma", "email", "produkt", "sachbearbeiterEmail"]
                for f in required:
                    if not body.get(f):
                        return json_resp({"error": f"Feld '{f}' fehlt"}, 400)
                fields = {
                    "Title":               body["name"],
                    "Firma":               body["firma"],
                    "Email":               body["email"],
                    "Produkt":             body["produkt"],
                    "MailGesendet":        False,
                    "LinkGeoeffnet":       False,
                    "SachbearbeiterEmail": body["sachbearbeiterEmail"],
                    "Notizen":             body.get("notizen", ""),
                    "Aktiv":               True,
                }
                created = create_item(graph_token, fields)
                item_id = int(created["id"])
                token = encode_token(item_id)
                update_item(graph_token, item_id, {"Token": token})
                return json_resp({"success": True, "itemId": item_id, "token": token}, 201)

            if action in ("update", "deactivate", "extend"):
                item_id = body.get("itemId")
                if not item_id:
                    return json_resp({"error": "itemId fehlt"}, 400)
                item_id = int(item_id)

                if action == "deactivate":
                    update_item(graph_token, item_id, {"Aktiv": False})
                elif action == "extend":
                    update_item(graph_token, item_id, {"Aktiv": True})
                else:
                    allowed = ["Notizen", "SachbearbeiterEmail", "Aktiv"]
                    patch = {k: body[k] for k in allowed if k in body}
                    if patch:
                        update_item(graph_token, item_id, patch)
                return json_resp({"success": True})

            return json_resp({"error": f"Unbekannte action: {action}"}, 400)

        except Exception as e:
            logging.exception("demo POST error")
            return json_resp({"error": str(e)}, 500)


# ---------------------------------------------------------------------------
# Endpunkt: POST /api/demo-track
# ---------------------------------------------------------------------------

@app.route(route="demo-track", methods=["POST", "OPTIONS"],
           auth_level=func.AuthLevel.ANONYMOUS)
def demo_track(req: func.HttpRequest) -> func.HttpResponse:
    """Aufgerufen von der Interessenten-Seite beim Laden der Seite."""
    logging.info("demo-track")

    if req.method == "OPTIONS":
        return options_resp()

    try:
        body = req.get_json()
        token = body.get("token", "")
    except Exception:
        return json_resp({"error": "Kein gültiges JSON"}, 400)

    item_id = decode_token(token)
    if item_id is None:
        return json_resp({"error": "Ungültiger Token"}, 400)

    try:
        graph_token = get_app_token()
        fields = get_item_by_id(graph_token, item_id)
        if fields is None:
            return json_resp({"error": "Nicht gefunden"}, 404)

        already_tracked = fields.get("LinkGeoeffnet", False)
        now_iso = datetime.now(timezone.utc).isoformat()

        update_item(graph_token, item_id, {
            "LinkGeoeffnet":   True,
            "LinkGeoeffnetAm": now_iso,
        })

        if not already_tracked:
            sachbearbeiter = fields.get("SachbearbeiterEmail", "")
            if sachbearbeiter:
                send_internal_notification(
                    sachbearbeiter,
                    fields.get("Title", ""),
                    fields.get("Firma", ""),
                    graph_token,
                )

        return json_resp({"success": True})
    except Exception as e:
        logging.exception("demo-track error")
        return json_resp({"error": str(e)}, 500)


# ---------------------------------------------------------------------------
# Endpunkt: POST /api/demo-mail
# ---------------------------------------------------------------------------

@app.route(route="demo-mail", methods=["POST", "OPTIONS"],
           auth_level=func.AuthLevel.ANONYMOUS)
def demo_mail(req: func.HttpRequest) -> func.HttpResponse:
    """Sendet die personalisierte Demo-Mail an den Interessenten."""
    logging.info("demo-mail")

    if req.method == "OPTIONS":
        return options_resp()

    if not require_admin(req):
        return json_resp({"error": "Unauthorized"}, 401)
    try:
        body = req.get_json()
        item_id = int(body.get("itemId", 0))
        if not item_id:
            return json_resp({"error": "itemId fehlt"}, 400)
    except Exception:
        return json_resp({"error": "Kein gültiges JSON"}, 400)

    try:
        graph_token = get_app_token()
        fields = get_item_by_id(graph_token, item_id)
        if fields is None:
            return json_resp({"error": "Interessent nicht gefunden"}, 404)

        token   = encode_token(item_id)
        name    = fields.get("Title", "")
        firma   = fields.get("Firma", "")
        email   = fields.get("Email", "")
        produkt = fields.get("Produkt", "")
        sachb   = fields.get("SachbearbeiterEmail", "")

        html_body = build_demo_email_html(name, firma, produkt, token, sachb)
        subject   = f"DEMO Komda® Software {produkt} – Ihr persönlicher Zugang"

        send_graph_email(email, subject, html_body, graph_token)

        update_item(graph_token, item_id, {
            "MailGesendet": True,
            "Token":        token,
        })

        return json_resp({"success": True, "token": token})
    except Exception as e:
        logging.exception("demo-mail error")
        return json_resp({"error": str(e)}, 500)


# ---------------------------------------------------------------------------
# Bestehender Endpunkt: GET/POST /api/status  (unverändert übernehmen)
# ---------------------------------------------------------------------------

@app.route(route="status", methods=["GET", "POST"],
           auth_level=func.AuthLevel.ANONYMOUS)
def status(req: func.HttpRequest) -> func.HttpResponse:
    """Placeholder – vorhandene Implementierung hier einfügen."""
    return func.HttpResponse(
        json.dumps({"info": "status endpoint – vorhandene Implementierung einfügen"}),
        mimetype="application/json",
    )
