import os
import hmac
import hashlib
import html
import io
from datetime import datetime, timedelta, timezone
from functools import wraps
from zoneinfo import ZoneInfo

import requests
from flask import Flask, jsonify, request, send_file
from flask_sqlalchemy import SQLAlchemy
from sqlalchemy import or_
from sqlalchemy.exc import IntegrityError
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter


def normalize_database_url(url: str) -> str:
    if not url:
        return "sqlite:///henna_cloud_local.db"
    if url.startswith("postgres://"):
        return "postgresql+psycopg://" + url[len("postgres://"):]
    if url.startswith("postgresql://"):
        return "postgresql+psycopg://" + url[len("postgresql://"):]
    return url


DATABASE_URL = normalize_database_url(os.getenv("DATABASE_URL", ""))
VERIFY_TOKEN = os.getenv("VERIFY_TOKEN", "")
META_APP_SECRET = os.getenv("META_APP_SECRET", "")
META_SYSTEM_USER_TOKEN = os.getenv("META_SYSTEM_USER_TOKEN", "")
META_GRAPH_VERSION = os.getenv("META_GRAPH_VERSION", "v26.0")
ADMIN_API_KEY = os.getenv("ADMIN_API_KEY", "")
TIMEZONE_NAME = os.getenv("TIMEZONE", "Asia/Jerusalem")
NEW_CONVERSATION_AFTER_HOURS = int(os.getenv("NEW_CONVERSATION_AFTER_HOURS", "24"))
DEFAULT_OWNER = os.getenv("DEFAULT_OWNER", "מעיין")
DEFAULT_PRIORITY = os.getenv("DEFAULT_PRIORITY", "בינונית")
RESEND_API_KEY = os.getenv("RESEND_API_KEY", "")
ALERT_EMAIL = os.getenv("ALERT_EMAIL", "")
EMAIL_FROM = os.getenv("EMAIL_FROM", "Olam HaHina <onboarding@resend.dev>")
MORNING_API_KEY_ID = os.getenv("MORNING_API_KEY_ID", "")
MORNING_API_KEY_SECRET = os.getenv("MORNING_API_KEY_SECRET", "")
MORNING_ENV = os.getenv("MORNING_ENV", "production").strip().lower()
MORNING_WEBHOOK_TOKEN = os.getenv("MORNING_WEBHOOK_TOKEN", "")

TZ = ZoneInfo(TIMEZONE_NAME)

app = Flask(__name__)
app.config["SQLALCHEMY_DATABASE_URI"] = DATABASE_URL
app.config["SQLALCHEMY_TRACK_MODIFICATIONS"] = False
app.config["SQLALCHEMY_ENGINE_OPTIONS"] = {
    "pool_pre_ping": True,
    "pool_recycle": 300,
}

db = SQLAlchemy(app)


class Lead(db.Model):
    __tablename__ = "leads"

    id = db.Column(db.Integer, primary_key=True)
    lead_uid = db.Column(db.String(80), unique=True, nullable=False, index=True)
    created_at = db.Column(db.DateTime(timezone=True), nullable=False, default=lambda: datetime.now(timezone.utc))
    updated_at = db.Column(db.DateTime(timezone=True), nullable=False, default=lambda: datetime.now(timezone.utc))
    last_contact_at = db.Column(db.DateTime(timezone=True), nullable=False, default=lambda: datetime.now(timezone.utc))

    customer_name = db.Column(db.String(255), default="")
    phone = db.Column(db.String(50), nullable=False, index=True)
    email = db.Column(db.String(255), default="")

    source = db.Column(db.String(100), nullable=False, default="WhatsApp Campaign")
    platform = db.Column(db.String(50), default="WhatsApp")
    campaign_id = db.Column(db.String(100), default="")
    campaign_name = db.Column(db.String(255), default="")
    adset_id = db.Column(db.String(100), default="")
    adset_name = db.Column(db.String(255), default="")
    ad_id = db.Column(db.String(100), default="")
    ad_name = db.Column(db.String(255), default="")
    referral_source = db.Column(db.Text, default="")

    first_message = db.Column(db.Text, default="")
    last_message = db.Column(db.Text, default="")
    message_count = db.Column(db.Integer, nullable=False, default=0)

    lead_status = db.Column(db.String(100), nullable=False, default="ליד חדש", index=True)
    priority = db.Column(db.String(50), default=DEFAULT_PRIORITY)
    owner = db.Column(db.String(100), default=DEFAULT_OWNER)

    next_follow_up_at = db.Column(db.DateTime(timezone=True), nullable=True, index=True)
    lost_reason = db.Column(db.Text, default="")
    close_date = db.Column(db.DateTime(timezone=True), nullable=True)

    event_type = db.Column(db.String(100), default="")
    event_date = db.Column(db.Date, nullable=True, index=True)
    start_time = db.Column(db.Time, nullable=True)
    end_time = db.Column(db.Time, nullable=True)
    city = db.Column(db.String(150), default="")
    venue_address = db.Column(db.Text, default="")
    guests = db.Column(db.Integer, nullable=True)

    package = db.Column(db.String(150), default="")
    concept_world = db.Column(db.String(150), default="")
    colors_style = db.Column(db.String(255), default="")
    customer_choice_summary = db.Column(db.Text, default="")
    seating_furniture = db.Column(db.String(150), default="")
    backdrop_canopy = db.Column(db.String(150), default="")
    carpets = db.Column(db.String(150), default="")
    entrance_gate = db.Column(db.Boolean, default=False)
    market_table = db.Column(db.Boolean, default=False)
    vitrines_qty = db.Column(db.Integer, nullable=True)
    balloons = db.Column(db.Boolean, default=False)
    proposal_heart = db.Column(db.Boolean, default=False)
    challah_separation = db.Column(db.Boolean, default=False)
    cookies_kg = db.Column(db.Numeric(10, 2), nullable=True)
    sfenj_mufleta = db.Column(db.Boolean, default=False)

    guest_costumes_qty = db.Column(db.Integer, nullable=True)
    bride_dress_size = db.Column(db.String(50), default="")
    groom_outfit_size = db.Column(db.String(50), default="")
    bride_outfits_qty = db.Column(db.Integer, nullable=True)
    groom_outfits_qty = db.Column(db.Integer, nullable=True)

    quote_price = db.Column(db.Numeric(12, 2), nullable=True)
    discount = db.Column(db.Numeric(12, 2), nullable=True)
    final_price = db.Column(db.Numeric(12, 2), nullable=True)

    client_requirements = db.Column(db.Text, default="")
    internal_notes = db.Column(db.Text, default="")

    messages = db.relationship("Message", backref="lead", lazy=True, cascade="all, delete-orphan")
    payments = db.relationship("Payment", backref="lead", lazy=True, cascade="all, delete-orphan")
    followups = db.relationship("FollowUp", backref="lead", lazy=True, cascade="all, delete-orphan")


class Message(db.Model):
    __tablename__ = "messages"

    id = db.Column(db.Integer, primary_key=True)
    meta_message_id = db.Column(db.String(255), unique=True, nullable=False, index=True)
    lead_id = db.Column(db.Integer, db.ForeignKey("leads.id"), nullable=True, index=True)
    phone = db.Column(db.String(50), nullable=False, index=True)
    customer_name = db.Column(db.String(255), default="")
    received_at = db.Column(db.DateTime(timezone=True), nullable=False)
    message_type = db.Column(db.String(50), default="unknown")
    text = db.Column(db.Text, default="")
    campaign_entry = db.Column(db.Boolean, nullable=False, default=False)
    referral_json = db.Column(db.JSON, nullable=True)


class Payment(db.Model):
    __tablename__ = "payments"

    id = db.Column(db.Integer, primary_key=True)
    payment_uid = db.Column(db.String(80), unique=True, nullable=True, index=True)

    lead_id = db.Column(
        db.Integer,
        db.ForeignKey("leads.id"),
        nullable=False,
        index=True
    )

    paid_at = db.Column(
        db.DateTime(timezone=True),
        nullable=False,
        default=lambda: datetime.now(timezone.utc)
    )

    payment_type = db.Column(db.String(100), default="")
    amount = db.Column(db.Numeric(12, 2), nullable=False)
    method = db.Column(db.String(100), default="")
    reference = db.Column(db.String(255), default="")
    status = db.Column(db.String(50), default="שולם")
    note = db.Column(db.Text, default="")

    morning_payment_id = db.Column(db.String(255), default="")
    morning_document_id = db.Column(db.String(255), default="")
    document_type = db.Column(db.String(100), default="")
    document_number = db.Column(db.String(100), default="")
    document_url = db.Column(db.Text, default="")
    last_sync_at = db.Column(db.DateTime(timezone=True), nullable=True)


class FollowUp(db.Model):
    __tablename__ = "followups"

    id = db.Column(db.Integer, primary_key=True)
    lead_id = db.Column(db.Integer, db.ForeignKey("leads.id"), nullable=False, index=True)
    due_at = db.Column(db.DateTime(timezone=True), nullable=False, index=True)
    status = db.Column(db.String(50), nullable=False, default="פתוח")
    note = db.Column(db.Text, default="")
    created_at = db.Column(db.DateTime(timezone=True), nullable=False, default=lambda: datetime.now(timezone.utc))
    completed_at = db.Column(db.DateTime(timezone=True), nullable=True)


def utc_now():
    return datetime.now(timezone.utc)


def local_iso(dt):
    if not dt:
        return None
    if dt.tzinfo is None:
        dt = dt.replace(tzinfo=timezone.utc)
    return dt.astimezone(TZ).isoformat(timespec="seconds")


def parse_meta_timestamp(value):
    try:
        return datetime.fromtimestamp(int(value), tz=timezone.utc)
    except Exception:
        return utc_now()


def verify_signature(raw_body: bytes, header: str | None) -> bool:
    if not META_APP_SECRET or not header or not header.startswith("sha256="):
        return False
    received = header.split("=", 1)[1]
    expected = hmac.new(
        META_APP_SECRET.encode("utf-8"),
        raw_body,
        hashlib.sha256,
    ).hexdigest()
    return hmac.compare_digest(received, expected)


def extract_text(message):
    msg_type = message.get("type", "unknown")
    if msg_type == "text":
        return message.get("text", {}).get("body", "")
    if msg_type == "image":
        return ("[תמונה] " + message.get("image", {}).get("caption", "")).strip()
    if msg_type == "video":
        return ("[וידאו] " + message.get("video", {}).get("caption", "")).strip()
    if msg_type == "audio":
        return "[הודעה קולית]"
    if msg_type == "document":
        doc = message.get("document", {})
        return f"[מסמך] {doc.get('filename','')} {doc.get('caption','')}".strip()
    if msg_type == "location":
        loc = message.get("location", {})
        return f"[מיקום] {loc.get('name','')} {loc.get('address','')}".strip()
    if msg_type == "sticker":
        return "[מדבקה]"
    if msg_type == "button":
        return message.get("button", {}).get("text", "[כפתור]")
    if msg_type == "interactive":
        interactive = message.get("interactive", {})
        if "button_reply" in interactive:
            return interactive["button_reply"].get("title", "[כפתור]")
        if "list_reply" in interactive:
            return interactive["list_reply"].get("title", "[רשימה]")
        return "[הודעה אינטראקטיבית]"
    return f"[{msg_type}]"


def is_campaign_message(message):
    referral = message.get("referral")
    return isinstance(referral, dict) and bool(
        referral.get("source_id") or referral.get("source_url") or referral.get("ctwa_clid")
    )


def infer_platform(referral):
    url = str((referral or {}).get("source_url", "")).lower()
    if "instagram" in url:
        return "Instagram"
    if "facebook" in url or "fb." in url:
        return "Facebook"
    return "WhatsApp"


def fetch_ad_metadata(ad_id):
    if not ad_id or not META_SYSTEM_USER_TOKEN:
        return {}
    try:
        response = requests.get(
            f"https://graph.facebook.com/{META_GRAPH_VERSION}/{ad_id}",
            params={
                "fields": "id,name,adset{id,name},campaign{id,name}",
                "access_token": META_SYSTEM_USER_TOKEN,
            },
            timeout=6,
        )
        if not response.ok:
            return {}
        data = response.json()
        return {
            "ad_id": data.get("id", ad_id),
            "ad_name": data.get("name", ""),
            "adset_id": (data.get("adset") or {}).get("id", ""),
            "adset_name": (data.get("adset") or {}).get("name", ""),
            "campaign_id": (data.get("campaign") or {}).get("id", ""),
            "campaign_name": (data.get("campaign") or {}).get("name", ""),
        }
    except Exception:
        return {}


def referral_metadata(message):
    referral = message.get("referral") or {}
    source_type = str(referral.get("source_type", "") or "")
    source_id = str(referral.get("source_id", "") or "")
    meta = {
        "platform": infer_platform(referral),
        "campaign_id": "",
        "campaign_name": "",
        "adset_id": "",
        "adset_name": "",
        "ad_id": source_id if source_type == "ad" else "",
        "ad_name": str(referral.get("headline", "") or ""),
        "referral_source": str(
            referral.get("source_url") or referral.get("ctwa_clid") or source_type or ""
        ),
        "raw": referral,
    }
    if source_type == "ad" and source_id:
        graph = fetch_ad_metadata(source_id)
        for key, value in graph.items():
            if value:
                meta[key] = value
    return meta


def active_campaign_lead(phone, received_at):
    cutoff = received_at - timedelta(hours=NEW_CONVERSATION_AFTER_HOURS)
    return (
        Lead.query
        .filter(
            Lead.phone == phone,
            Lead.source == "WhatsApp Campaign",
            Lead.last_contact_at >= cutoff,
        )
        .order_by(Lead.last_contact_at.desc())
        .first()
    )


def new_lead_uid(received_at):
    return "LEAD-" + received_at.astimezone(TZ).strftime("%Y%m%d-%H%M%S-%f")

def new_payment_uid():
    now = datetime.now(TZ)
    return "PAY-" + now.strftime("%Y%m%d-%H%M%S-%f")
    
def send_new_lead_email(lead):
    """Send a best-effort email alert for a newly created campaign lead.

    This runs only *after* the lead and message have been committed to PostgreSQL.
    Email failures are logged and never roll back or lose the lead.
    """
    if not RESEND_API_KEY or not ALERT_EMAIL:
        return {"status": "disabled"}

    def esc(value):
        return html.escape(str(value or ""))

    local_time = lead.created_at.astimezone(TZ).strftime("%d/%m/%Y %H:%M")
    customer = lead.customer_name or "ללא שם"
    platform = lead.platform or "WhatsApp"
    campaign = lead.campaign_name or lead.campaign_id or "לא התקבל שם קמפיין"
    ad = lead.ad_name or lead.ad_id or "לא התקבל שם מודעה"
    message = lead.first_message or ""

    subject = f"ליד חדש - עולם החינה - {customer}"
    html_body = f"""
    <div dir="rtl" style="font-family:Arial,sans-serif;line-height:1.6">
      <h2>🔔 ליד חדש - עולם החינה</h2>
      <table style="border-collapse:collapse">
        <tr><td><b>שם:</b></td><td>{esc(customer)}</td></tr>
        <tr><td><b>טלפון:</b></td><td>{esc(lead.phone)}</td></tr>
        <tr><td><b>מקור:</b></td><td>{esc(platform)}</td></tr>
        <tr><td><b>קמפיין:</b></td><td>{esc(campaign)}</td></tr>
        <tr><td><b>מודעה:</b></td><td>{esc(ad)}</td></tr>
        <tr><td><b>נכנס בתאריך:</b></td><td>{esc(local_time)}</td></tr>
        <tr><td><b>Lead ID:</b></td><td>{esc(lead.lead_uid)}</td></tr>
      </table>
      <p><b>הודעה ראשונה:</b></p>
      <div style="padding:10px;border:1px solid #ddd;border-radius:8px;white-space:pre-wrap">{esc(message)}</div>
    </div>
    """

    response = requests.post(
        "https://api.resend.com/emails",
        headers={
            "Authorization": f"Bearer {RESEND_API_KEY}",
            "Content-Type": "application/json",
            "Idempotency-Key": f"new-lead/{lead.lead_uid}",
        },
        json={
            "from": EMAIL_FROM,
            "to": [ALERT_EMAIL],
            "subject": subject,
            "html": html_body,
        },
        timeout=10,
    )
    response.raise_for_status()
    data = response.json()
    return {"status": "sent", "email_id": data.get("id", "")}


def process_message(phone, name, message):
    meta_message_id = str(message.get("id", ""))
    if not meta_message_id:
        return {"status": "ignored_no_message_id"}

    existing_message = Message.query.filter_by(meta_message_id=meta_message_id).first()
    if existing_message:
        return {"status": "duplicate", "message_id": meta_message_id}

    received_at = parse_meta_timestamp(message.get("timestamp"))
    text = extract_text(message)
    campaign_entry = is_campaign_message(message)
    meta = referral_metadata(message) if campaign_entry else {}

    lead = active_campaign_lead(phone, received_at)

    # A normal direct WhatsApp message never creates a new CRM lead.
    if lead is None and not campaign_entry:
        msg = Message(
            meta_message_id=meta_message_id,
            lead_id=None,
            phone=phone,
            customer_name=name,
            received_at=received_at,
            message_type=message.get("type", "unknown"),
            text=text,
            campaign_entry=False,
            referral_json=None,
        )
        db.session.add(msg)
        db.session.commit()
        return {"status": "ignored_non_campaign"}

    if lead is None:
        lead = Lead(
            lead_uid=new_lead_uid(received_at),
            created_at=received_at,
            updated_at=received_at,
            last_contact_at=received_at,
            customer_name=name,
            phone=phone,
            source="WhatsApp Campaign",
            platform=meta.get("platform", "WhatsApp"),
            campaign_id=meta.get("campaign_id", ""),
            campaign_name=meta.get("campaign_name", ""),
            adset_id=meta.get("adset_id", ""),
            adset_name=meta.get("adset_name", ""),
            ad_id=meta.get("ad_id", ""),
            ad_name=meta.get("ad_name", ""),
            referral_source=meta.get("referral_source", ""),
            first_message=text,
            last_message=text,
            message_count=1,
            lead_status="ליד חדש",
            priority=DEFAULT_PRIORITY,
            owner=DEFAULT_OWNER,
            client_requirements=text,
            internal_notes=f"[WhatsApp {local_iso(received_at)}] {text}",
        )
        db.session.add(lead)
        db.session.flush()
        result_status = "new_campaign_lead"
    else:
        lead.updated_at = utc_now()
        lead.last_contact_at = received_at
        lead.customer_name = name or lead.customer_name
        lead.last_message = text
        lead.message_count = (lead.message_count or 0) + 1
        line = f"[WhatsApp {local_iso(received_at)}] {text}"
        lead.internal_notes = (lead.internal_notes + "\n" + line).strip() if lead.internal_notes else line

        if campaign_entry:
            lead.platform = lead.platform or meta.get("platform", "")
            lead.campaign_id = lead.campaign_id or meta.get("campaign_id", "")
            lead.campaign_name = lead.campaign_name or meta.get("campaign_name", "")
            lead.adset_id = lead.adset_id or meta.get("adset_id", "")
            lead.adset_name = lead.adset_name or meta.get("adset_name", "")
            lead.ad_id = lead.ad_id or meta.get("ad_id", "")
            lead.ad_name = lead.ad_name or meta.get("ad_name", "")
            lead.referral_source = lead.referral_source or meta.get("referral_source", "")

        result_status = "updated_campaign_lead"

    msg = Message(
        meta_message_id=meta_message_id,
        lead_id=lead.id,
        phone=phone,
        customer_name=name,
        received_at=received_at,
        message_type=message.get("type", "unknown"),
        text=text,
        campaign_entry=campaign_entry,
        referral_json=meta.get("raw") if campaign_entry else None,
    )
    db.session.add(msg)
    db.session.commit()

    email_alert = None
    if result_status == "new_campaign_lead":
        try:
            email_alert = send_new_lead_email(lead)
        except Exception as exc:
            # The lead is already safely committed to PostgreSQL.
            # Notification failure must never make Meta retry/lose the lead.
            app.logger.exception("New lead email alert failed")
            email_alert = {"status": "failed", "error": str(exc)[:500]}

    return {
        "status": result_status,
        "lead_id": lead.id,
        "lead_uid": lead.lead_uid,
        "email_alert": email_alert,
    }

def morning_access_token():
    if not MORNING_API_KEY_ID or not MORNING_API_KEY_SECRET:
        raise RuntimeError("Morning API credentials are not configured")

    if MORNING_ENV == "sandbox":
        token_url = "https://api.sandbox.morning.dev/idp/v1/oauth/token"
    else:
        token_url = "https://api.morning.co/idp/v1/oauth/token"

    response = requests.post(
        token_url,
        headers={
            "Content-Type": "application/x-www-form-urlencoded"
        },
        data={
            "grant_type": "client_credentials",
            "client_id": MORNING_API_KEY_ID,
            "client_secret": MORNING_API_KEY_SECRET,
        },
        timeout=10,
    )

    response.raise_for_status()

    data = response.json()

    token = data.get("accessToken") or data.get("access_token")

    if not token:
        raise RuntimeError(
            "Morning authentication succeeded but no access token was returned"
        )

    return token
def admin_required(fn):
    @wraps(fn)
    def wrapper(*args, **kwargs):
        if not ADMIN_API_KEY:
            return jsonify({"error": "ADMIN_API_KEY is not configured"}), 503
        key = request.headers.get("X-Admin-Key", "")
        if not hmac.compare_digest(key, ADMIN_API_KEY):
            return jsonify({"error": "Unauthorized"}), 401
        return fn(*args, **kwargs)
    return wrapper

@app.post("/api/morning/test")
@admin_required
def morning_test():
    try:
        morning_access_token()

        return jsonify({
            "ok": True,
            "authenticated": True,
            "environment": MORNING_ENV,
            "message": "Morning API authentication successful",
        })

    except requests.HTTPError as exc:
        body = (
            exc.response.text[:500]
            if exc.response is not None
            else str(exc)
        )

        return jsonify({
            "ok": False,
            "authenticated": False,
            "environment": MORNING_ENV,
            "error": body,
        }), 400

    except Exception as exc:
        return jsonify({
            "ok": False,
            "authenticated": False,
            "environment": MORNING_ENV,
            "error": str(exc)[:500],
        }), 502

@app.post("/webhook/morning")
@app.post("/webhook/morning/<path_token>")
def morning_webhook(path_token=None):
    token = path_token or request.args.get("token", "")

    if (
        not MORNING_WEBHOOK_TOKEN
        or not hmac.compare_digest(token, MORNING_WEBHOOK_TOKEN)
    ):
        return jsonify({
            "ok": False,
            "error": "Unauthorized"
        }), 401

    raw = request.get_data(cache=True)

    payload = request.get_json(silent=True) or {}

    document_id = (
        payload.get("id")
        or payload.get("documentId")
        or (payload.get("document") or {}).get("id")
    )

    # זמני לצורך הבדיקה הראשונה בלבד:
    # נראה בדיוק איזה מבנה Morning שולחת.
    app.logger.info(
        "MORNING WEBHOOK | document_id=%s | payload=%s",
        document_id,
        raw.decode("utf-8", errors="replace")[:4000],
    )

    return jsonify({
        "ok": True,
        "received": True,
        "document_id": document_id,
    }), 200
        
@app.post("/api/test-email")
@admin_required
def test_email():
    """Send a test email from Railway without creating a lead."""
    if not RESEND_API_KEY:
        return jsonify({"ok": False, "error": "RESEND_API_KEY is not configured"}), 503
    if not ALERT_EMAIL:
        return jsonify({"ok": False, "error": "ALERT_EMAIL is not configured"}), 503

    html_body = """
    <div dir="rtl" style="font-family:Arial,sans-serif;line-height:1.6">
      <h2>✅ בדיקת התראות - עולם החינה</h2>
      <p>Railway הצליח לשלוח מייל דרך Resend.</p>
      <p><b>Railway → Resend → Email</b></p>
    </div>
    """
    try:
        response = requests.post(
            "https://api.resend.com/emails",
            headers={
                "Authorization": f"Bearer {RESEND_API_KEY}",
                "Content-Type": "application/json",
            },
            json={
                "from": EMAIL_FROM,
                "to": [ALERT_EMAIL],
                "subject": "בדיקת התראות - עולם החינה",
                "html": html_body,
            },
            timeout=10,
        )
        response.raise_for_status()
        data = response.json()
        return jsonify({"ok": True, "email_id": data.get("id", ""), "to": ALERT_EMAIL})
    except Exception as exc:
        app.logger.exception("Test email failed")
        return jsonify({"ok": False, "error": str(exc)[:500]}), 502


def lead_to_dict(lead):
    total_paid = sum(float(p.amount or 0) for p in lead.payments)
    final_price = float(lead.final_price or 0) if lead.final_price is not None else None
    balance = (final_price - total_paid) if final_price is not None else None

    if final_price is None:
        payment_status = ""
    elif total_paid <= 0:
        payment_status = "לא שולם"
    elif balance is not None and balance <= 0:
        payment_status = "שולם במלואו"
    else:
        payment_status = "שולם חלקית"

    return {
        "id": lead.id,
        "lead_uid": lead.lead_uid,
        "created_at": local_iso(lead.created_at),
        "last_contact_at": local_iso(lead.last_contact_at),
        "customer_name": lead.customer_name,
        "phone": lead.phone,
        "email": lead.email,
        "source": lead.source,
        "platform": lead.platform,
        "campaign_id": lead.campaign_id,
        "campaign_name": lead.campaign_name,
        "adset_id": lead.adset_id,
        "adset_name": lead.adset_name,
        "ad_id": lead.ad_id,
        "ad_name": lead.ad_name,
        "referral_source": lead.referral_source,
        "first_message": lead.first_message,
        "last_message": lead.last_message,
        "message_count": lead.message_count,
        "lead_status": lead.lead_status,
        "priority": lead.priority,
        "owner": lead.owner,
        "next_follow_up_at": local_iso(lead.next_follow_up_at),
        "lost_reason": lead.lost_reason,
        "close_date": local_iso(lead.close_date),
        "event_type": lead.event_type,
        "event_date": lead.event_date.isoformat() if lead.event_date else None,
        "start_time": lead.start_time.isoformat(timespec="minutes") if lead.start_time else None,
        "end_time": lead.end_time.isoformat(timespec="minutes") if lead.end_time else None,
        "city": lead.city,
        "venue_address": lead.venue_address,
        "guests": lead.guests,
        "package": lead.package,
        "concept_world": lead.concept_world,
        "colors_style": lead.colors_style,
        "customer_choice_summary": lead.customer_choice_summary,
        "seating_furniture": lead.seating_furniture,
        "backdrop_canopy": lead.backdrop_canopy,
        "carpets": lead.carpets,
        "entrance_gate": lead.entrance_gate,
        "market_table": lead.market_table,
        "vitrines_qty": lead.vitrines_qty,
        "balloons": lead.balloons,
        "proposal_heart": lead.proposal_heart,
        "challah_separation": lead.challah_separation,
        "cookies_kg": float(lead.cookies_kg) if lead.cookies_kg is not None else None,
        "sfenj_mufleta": lead.sfenj_mufleta,
        "guest_costumes_qty": lead.guest_costumes_qty,
        "bride_dress_size": lead.bride_dress_size,
        "groom_outfit_size": lead.groom_outfit_size,
        "bride_outfits_qty": lead.bride_outfits_qty,
        "groom_outfits_qty": lead.groom_outfits_qty,
        "quote_price": float(lead.quote_price) if lead.quote_price is not None else None,
        "discount": float(lead.discount) if lead.discount is not None else None,
        "final_price": final_price,
        "total_paid": total_paid,
        "balance": balance,
        "payment_status": payment_status,
        "client_requirements": lead.client_requirements,
        "internal_notes": lead.internal_notes,
    }


@app.get("/")
@app.get("/health")
def health():
    try:
        db.session.execute(db.text("SELECT 1"))
        db_ok = True
    except Exception:
        db_ok = False
    return jsonify({
        "ok": db_ok,
        "service": "Henna Cloud CRM",
        "database": "connected" if db_ok else "error",
        "campaign_only": True,
    }), 200 if db_ok else 503


@app.get("/webhook")
def verify_webhook():
    mode = request.args.get("hub.mode")
    token = request.args.get("hub.verify_token")
    challenge = request.args.get("hub.challenge")
    if mode == "subscribe" and token == VERIFY_TOKEN and challenge is not None:
        return challenge, 200
    return "Forbidden", 403


@app.post("/webhook")
def webhook():
    raw = request.get_data(cache=True)
    if not verify_signature(raw, request.headers.get("X-Hub-Signature-256")):
        return jsonify({"error": "Invalid Meta signature"}), 403

    payload = request.get_json(silent=True) or {}
    results = []

    try:
        for entry in payload.get("entry", []):
            for change in entry.get("changes", []):
                if change.get("field") != "messages":
                    continue
                value = change.get("value", {})
                contacts = {
                    str(c.get("wa_id", "")): c.get("profile", {}).get("name", "")
                    for c in value.get("contacts", [])
                }
                for message in value.get("messages", []) or []:
                    phone = str(message.get("from", ""))
                    if not phone:
                        continue
                    results.append(process_message(phone, contacts.get(phone, ""), message))
        return jsonify({"received": True, "results": results}), 200
    except IntegrityError:
        db.session.rollback()
        return jsonify({"received": True, "status": "duplicate_race"}), 200
    except Exception as exc:
        db.session.rollback()
        app.logger.exception("Webhook database error")
        # Non-2xx lets Meta know this delivery was not processed successfully.
        return jsonify({"received": False, "error": str(exc)}), 503

def payment_to_dict(payment):
    lead = payment.lead

    return {
        "id": payment.id,
        "payment_uid": payment.payment_uid,
        "lead_id": payment.lead_id,
        "lead_uid": lead.lead_uid if lead else "",
        "customer_name": lead.customer_name if lead else "",

        "paid_at": local_iso(payment.paid_at),
        "payment_type": payment.payment_type,
        "amount": float(payment.amount or 0),
        "method": payment.method,
        "reference": payment.reference,
        "status": payment.status,
        "note": payment.note,

        "morning_payment_id": payment.morning_payment_id,
        "morning_document_id": payment.morning_document_id,
        "document_type": payment.document_type,
        "document_number": payment.document_number,
        "document_url": payment.document_url,
        "last_sync_at": local_iso(payment.last_sync_at),
    }
    
@app.get("/api/leads")
@admin_required
def api_leads():
    page = max(int(request.args.get("page", 1)), 1)
    per_page = min(max(int(request.args.get("per_page", 50)), 1), 200)
    status = request.args.get("status")
    q = request.args.get("q", "").strip()

    query = Lead.query.order_by(Lead.created_at.desc())
    if status:
        query = query.filter(Lead.lead_status == status)
    if q:
        like = f"%{q}%"
        query = query.filter(or_(
            Lead.customer_name.ilike(like),
            Lead.phone.ilike(like),
            Lead.city.ilike(like),
            Lead.event_type.ilike(like),
        ))

    pagination = query.paginate(page=page, per_page=per_page, error_out=False)
    return jsonify({
        "page": page,
        "pages": pagination.pages,
        "total": pagination.total,
        "items": [lead_to_dict(x) for x in pagination.items],
    })


@app.get("/api/leads/<int:lead_id>")
@admin_required
def api_lead(lead_id):
    lead = db.get_or_404(Lead, lead_id)
    data = lead_to_dict(lead)
    data["messages"] = [{
        "received_at": local_iso(m.received_at),
        "type": m.message_type,
        "text": m.text,
        "campaign_entry": m.campaign_entry,
    } for m in sorted(lead.messages, key=lambda x: x.received_at)]
    data["payments"] = [{
        "id": p.id,
        "paid_at": local_iso(p.paid_at),
        "amount": float(p.amount),
        "method": p.method,
        "note": p.note,
    } for p in lead.payments]
    return jsonify(data)


EDITABLE_FIELDS = {
    "customer_name", "email", "lead_status", "priority", "owner",
    "event_type", "city", "venue_address", "package", "concept_world",
    "colors_style", "customer_choice_summary", "seating_furniture",
    "backdrop_canopy", "carpets", "entrance_gate", "market_table",
    "vitrines_qty", "balloons", "proposal_heart", "challah_separation",
    "cookies_kg", "sfenj_mufleta", "guest_costumes_qty", "bride_dress_size",
    "groom_outfit_size", "bride_outfits_qty", "groom_outfits_qty",
    "quote_price", "discount", "final_price", "client_requirements",
    "internal_notes", "lost_reason",
}


@app.patch("/api/leads/<int:lead_id>")
@admin_required
def update_lead(lead_id):
    lead = db.get_or_404(Lead, lead_id)
    body = request.get_json(silent=True) or {}

    for key, value in body.items():
        if key in EDITABLE_FIELDS:
            setattr(lead, key, value)

    if "event_date" in body:
        lead.event_date = datetime.fromisoformat(body["event_date"]).date() if body["event_date"] else None
    if "next_follow_up_at" in body:
        lead.next_follow_up_at = datetime.fromisoformat(body["next_follow_up_at"]).astimezone(timezone.utc) if body["next_follow_up_at"] else None

    lead.updated_at = utc_now()
    db.session.commit()
    return jsonify(lead_to_dict(lead))

@app.get("/api/payments")
@admin_required
def api_payments():
    payments = (
        Payment.query
        .order_by(Payment.paid_at.desc())
        .all()
    )

    return jsonify({
        "total": len(payments),
        "items": [payment_to_dict(p) for p in payments],
    })
    
@app.post("/api/leads/<int:lead_id>/payments")
@admin_required
def add_payment(lead_id):
    lead = db.get_or_404(Lead, lead_id)
    body = request.get_json(silent=True) or {}

    if "amount" not in body:
        return jsonify({
            "ok": False,
            "error": "amount is required"
        }), 400

    paid_at = utc_now()

    if body.get("paid_at"):
        parsed = datetime.fromisoformat(body["paid_at"])

        if parsed.tzinfo is None:
            parsed = parsed.replace(tzinfo=TZ)

        paid_at = parsed.astimezone(timezone.utc)

    payment = Payment(
        payment_uid=new_payment_uid(),
        lead_id=lead.id,
        paid_at=paid_at,

        payment_type=body.get("payment_type", ""),
        amount=body["amount"],
        method=body.get("method", ""),
        reference=body.get("reference", ""),
        status=body.get("status", "שולם"),
        note=body.get("note", ""),

        morning_payment_id=body.get("morning_payment_id", ""),
        morning_document_id=body.get("morning_document_id", ""),
        document_type=body.get("document_type", ""),
        document_number=body.get("document_number", ""),
        document_url=body.get("document_url", ""),
        last_sync_at=utc_now(),
    )

    db.session.add(payment)
    db.session.commit()

    return jsonify({
        "ok": True,
        "payment": payment_to_dict(payment)
    }), 201

@app.post("/api/leads/<int:lead_id>/followups")
@admin_required
def add_followup(lead_id):
    lead = db.get_or_404(Lead, lead_id)
    body = request.get_json(silent=True) or {}
    due_at = datetime.fromisoformat(body["due_at"])
    if due_at.tzinfo is None:
        due_at = due_at.replace(tzinfo=TZ)
    followup = FollowUp(
        lead_id=lead.id,
        due_at=due_at.astimezone(timezone.utc),
        note=body.get("note", ""),
    )
    lead.next_follow_up_at = followup.due_at
    db.session.add(followup)
    db.session.commit()
    return jsonify({"ok": True, "followup_id": followup.id}), 201


@app.get("/api/export.xlsx")
@admin_required
def export_xlsx():
    leads = Lead.query.order_by(Lead.created_at.desc()).all()

    wb = Workbook()
    ws = wb.active
    ws.title = "CRM"

    headers = [
        "Lead ID", "Created At", "Last Contact", "Customer Name", "Phone",
        "Source", "Platform", "Campaign", "Ad", "Lead Status", "Priority",
        "Event Type", "Event Date", "City", "Package", "Concept / World",
        "Bride Size", "Groom Size", "Final Price", "Total Paid", "Balance",
        "Next Follow-Up", "First Message", "Last Message", "Internal Notes"
    ]
    ws.append(headers)
    header_fill = PatternFill("solid", fgColor="1F4E78")
    for cell in ws[1]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center")

    for lead in leads:
        total_paid = sum(float(p.amount or 0) for p in lead.payments)
        final_price = float(lead.final_price or 0) if lead.final_price is not None else 0
        ws.append([
            lead.lead_uid,
            local_iso(lead.created_at),
            local_iso(lead.last_contact_at),
            lead.customer_name,
            lead.phone,
            lead.source,
            lead.platform,
            lead.campaign_name,
            lead.ad_name,
            lead.lead_status,
            lead.priority,
            lead.event_type,
            lead.event_date.isoformat() if lead.event_date else "",
            lead.city,
            lead.package,
            lead.concept_world,
            lead.bride_dress_size,
            lead.groom_outfit_size,
            final_price,
            total_paid,
            final_price - total_paid if lead.final_price is not None else "",
            local_iso(lead.next_follow_up_at),
            lead.first_message,
            lead.last_message,
            lead.internal_notes,
        ])

    widths = [24, 22, 22, 22, 18, 20, 14, 22, 22, 18, 12, 18, 14, 16, 18, 22, 14, 14, 14, 14, 14, 22, 36, 36, 50]
    for idx, width in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(idx)].width = width

    bio = io.BytesIO()
    wb.save(bio)
    bio.seek(0)
    return send_file(
        bio,
        as_attachment=True,
        download_name="Henna_CRM_Export.xlsx",
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


def migrate_payments_table():
    if db.engine.dialect.name != "postgresql":
        return

    statements = [
        "ALTER TABLE payments ADD COLUMN IF NOT EXISTS payment_uid VARCHAR(80)",
        "ALTER TABLE payments ADD COLUMN IF NOT EXISTS payment_type VARCHAR(100) DEFAULT ''",
        "ALTER TABLE payments ADD COLUMN IF NOT EXISTS reference VARCHAR(255) DEFAULT ''",
        "ALTER TABLE payments ADD COLUMN IF NOT EXISTS status VARCHAR(50) DEFAULT 'שולם'",
        "ALTER TABLE payments ADD COLUMN IF NOT EXISTS morning_payment_id VARCHAR(255) DEFAULT ''",
        "ALTER TABLE payments ADD COLUMN IF NOT EXISTS morning_document_id VARCHAR(255) DEFAULT ''",
        "ALTER TABLE payments ADD COLUMN IF NOT EXISTS document_type VARCHAR(100) DEFAULT ''",
        "ALTER TABLE payments ADD COLUMN IF NOT EXISTS document_number VARCHAR(100) DEFAULT ''",
        "ALTER TABLE payments ADD COLUMN IF NOT EXISTS document_url TEXT DEFAULT ''",
        "ALTER TABLE payments ADD COLUMN IF NOT EXISTS last_sync_at TIMESTAMPTZ",
        "CREATE UNIQUE INDEX IF NOT EXISTS ix_payments_payment_uid ON payments (payment_uid)"
    ]

    # Gunicorn runs more than one worker. The advisory lock makes sure only
    # one worker performs the startup migration at a time.
    with db.engine.begin() as conn:
        conn.execute(db.text("SELECT pg_advisory_xact_lock(392026)"))
        for statement in statements:
            conn.execute(db.text(statement))


with app.app_context():
    db.create_all()
    migrate_payments_table()

if __name__ == "__main__":
    port = int(os.getenv("PORT", "5000"))
    app.run(host="0.0.0.0", port=port, debug=False)
