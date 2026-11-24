import logging
import os
import secrets
import threading
import time
from datetime import datetime, timedelta, timezone
from urllib.parse import urlencode, urlsplit

import requests
from dotenv import load_dotenv
from fastapi import FastAPI, Header, HTTPException, Request
from fastapi.middleware.cors import CORSMiddleware
from slowapi import Limiter, _rate_limit_exceeded_handler
from slowapi.errors import RateLimitExceeded
from slowapi.util import get_remote_address

load_dotenv()

logger = logging.getLogger(__name__)

DEFAULT_LIMIT = "30/minute"
GRAPH_URL = "https://graph.microsoft.com/v1.0"
HTTP_TIMEOUT = (5, 30)
BYTES_PER_GB = 1024 * 1024 * 1024

CLIENT_ID = os.getenv("CLIENT_ID")
TENANT_ID = os.getenv("TENANT_ID")
CLIENT_SECRET = os.getenv("CLIENT_SECRET")
REFRESH_TOKEN = os.getenv("REFRESH_TOKEN")
API_KEY = os.getenv("API_KEY")

required_vars = {
    "CLIENT_ID": CLIENT_ID,
    "TENANT_ID": TENANT_ID,
    "CLIENT_SECRET": CLIENT_SECRET,
    "REFRESH_TOKEN": REFRESH_TOKEN,
    "API_KEY": API_KEY,
}

for name, value in required_vars.items():
    if not value:
        raise RuntimeError(f"Missing environment variable: {name}")

DEBUG = os.getenv("DEBUG", "false").lower() == "true"

TOKEN_CACHE: str | None = None
TOKEN_EXPIRES: float | None = None
TOKEN_LOCK = threading.Lock()

app = FastAPI(
    title="Fredonet M365 API",
    version="1.0.0",
)

limiter = Limiter(key_func=get_remote_address)
app.state.limiter = limiter
app.add_exception_handler(RateLimitExceeded, _rate_limit_exceeded_handler)

app.add_middleware(
    CORSMiddleware,
    allow_origins=["https://links.fredonet.net"],
    allow_credentials=False,
    allow_methods=["GET"],
    allow_headers=["X-API-Key", "Content-Type"],
)


def graph_error() -> None:
    raise HTTPException(
        status_code=502,
        detail="Microsoft Graph unavailable",
    )


def require_api_key(x_api_key: str | None) -> None:
    if not API_KEY:
        raise HTTPException(
            status_code=500,
            detail="API_KEY is not configured",
        )
    if not x_api_key:
        raise HTTPException(
            status_code=401,
            detail="Missing API key",
        )
    if not secrets.compare_digest(x_api_key, API_KEY):
        raise HTTPException(
            status_code=401,
            detail="Invalid API key",
        )


def require_debug() -> None:
    if not DEBUG:
        raise HTTPException(
            status_code=404,
            detail="Not found",
        )


def get_access_token() -> str:
    global TOKEN_CACHE, TOKEN_EXPIRES

    now = time.time()
    if TOKEN_CACHE and TOKEN_EXPIRES and now < TOKEN_EXPIRES:
        return TOKEN_CACHE

    with TOKEN_LOCK:
        now = time.time()
        if TOKEN_CACHE and TOKEN_EXPIRES and now < TOKEN_EXPIRES:
            return TOKEN_CACHE

        token_url = f"https://login.microsoftonline.com/{TENANT_ID}/oauth2/v2.0/token"
        try:
            response = requests.post(
                token_url,
                data={
                    "client_id": CLIENT_ID,
                    "client_secret": CLIENT_SECRET,
                    "refresh_token": REFRESH_TOKEN,
                    "grant_type": "refresh_token",
                },
                timeout=HTTP_TIMEOUT,
            )
            response.raise_for_status()
            data = response.json()
        except (requests.RequestException, ValueError) as error:
            logger.warning("Failed to refresh Microsoft Graph access token: %s", type(error).__name__)
            graph_error()

        access_token = data.get("access_token") if isinstance(data, dict) else None
        if not isinstance(access_token, str) or not access_token:
            logger.warning("Microsoft token endpoint returned an invalid payload")
            graph_error()

        expires_in = data.get("expires_in", 3600)
        if not isinstance(expires_in, (int, float)):
            expires_in = 3600

        TOKEN_CACHE = access_token
        TOKEN_EXPIRES = now + max(int(expires_in) - 300, 60)
        return TOKEN_CACHE


def graph_get(endpoint: str, extra_headers: dict[str, str] | None = None) -> dict:
    token = get_access_token()
    headers = {"Authorization": f"Bearer {token}"}
    if extra_headers:
        headers.update(extra_headers)

    try:
        response = requests.get(
            f"{GRAPH_URL}{endpoint}",
            headers=headers,
            timeout=HTTP_TIMEOUT,
        )
        payload = response.json()
    except (requests.RequestException, ValueError) as error:
        logger.warning("Microsoft Graph request failed (%s): %s", endpoint, type(error).__name__)
        graph_error()

    if not isinstance(payload, dict):
        logger.warning("Microsoft Graph returned non-object JSON (%s)", endpoint)
        graph_error()

    return {
        "status": response.status_code,
        "data": payload,
    }


def get_unread_mail_count() -> int:
    result = graph_get(
        "/me/mailFolders/inbox/messages?$filter=isRead eq false&$count=true",
        {"ConsistencyLevel": "eventual"},
    )
    if result["status"] != 200:
        graph_error()

    count = result["data"].get("@odata.count", 0)
    return count if isinstance(count, int) else 0


def get_onedrive_quota() -> dict[str, float]:
    result = graph_get("/me/drive")
    if result["status"] != 200:
        graph_error()

    quota = result["data"].get("quota")
    if not isinstance(quota, dict):
        quota = {}

    return {
        "used_gb": round((quota.get("used") or 0) / BYTES_PER_GB, 2),
        "remaining_gb": round((quota.get("remaining") or 0) / BYTES_PER_GB, 2),
        "total_gb": round((quota.get("total") or 0) / BYTES_PER_GB, 2),
    }


def get_open_tasks() -> dict:
    lists = graph_get("/me/todo/lists")
    if lists["status"] != 200:
        return {"count": 0, "tasks": []}

    raw_lists = lists["data"].get("value")
    if not isinstance(raw_lists, list):
        return {"count": 0, "tasks": []}

    all_tasks = []
    for todo_list in raw_lists:
        if not isinstance(todo_list, dict):
            continue
        list_id = todo_list.get("id")
        if not isinstance(list_id, str) or not list_id:
            continue

        tasks = graph_get(f"/me/todo/lists/{list_id}/tasks")
        if tasks["status"] != 200:
            continue

        raw_tasks = tasks["data"].get("value")
        if not isinstance(raw_tasks, list):
            continue

        for task in raw_tasks:
            if isinstance(task, dict) and task.get("status") != "completed":
                all_tasks.append(task)

    return {
        "count": len(all_tasks),
        "tasks": all_tasks,
    }


def format_calendar_event(event: dict) -> dict:
    start = event.get("start") if isinstance(event.get("start"), dict) else {}
    end = event.get("end") if isinstance(event.get("end"), dict) else {}
    location = event.get("location") if isinstance(event.get("location"), dict) else {}
    return {
        "subject": event.get("subject"),
        "start": start.get("dateTime"),
        "end": end.get("dateTime"),
        "location": location.get("displayName"),
    }


def get_next_event(days_ahead: int = 30) -> dict | None:
    now = datetime.now(timezone.utc)
    end = now + timedelta(days=days_ahead)

    params = {
        "startDateTime": now.isoformat(),
        "endDateTime": end.isoformat(),
        "$orderby": "start/dateTime",
        "$top": 1,
    }

    result = graph_get(f"/me/calendarView?{urlencode(params)}")
    if result["status"] != 200:
        graph_error()

    events = result["data"].get("value")
    if not isinstance(events, list) or not events:
        return None

    first_event = events[0]
    if not isinstance(first_event, dict):
        return None

    return format_calendar_event(first_event)


def build_dashboard_data() -> dict:
    me_result = graph_get("/me")
    if me_result["status"] != 200:
        graph_error()

    unread = get_unread_mail_count()
    onedrive = get_onedrive_quota()
    next_event = get_next_event(days_ahead=30)
    todo = get_open_tasks()

    return {
        "user": {
            "name": me_result["data"].get("displayName"),
            "email": me_result["data"].get("mail"),
        },
        "mail": {
            "unread": unread,
        },
        "calendar": next_event,
        "todo": {
            "open": todo["count"],
        },
        "onedrive": onedrive,
    }


@app.middleware("http")
async def security_headers(request: Request, call_next):
    response = await call_next(request)
    response.headers["X-Content-Type-Options"] = "nosniff"
    response.headers["X-Frame-Options"] = "DENY"
    response.headers["Referrer-Policy"] = "no-referrer"
    response.headers["Cache-Control"] = "no-store"
    response.headers["Permissions-Policy"] = "camera=(), microphone=(), geolocation=()"
    response.headers["Strict-Transport-Security"] = "max-age=31536000; includeSubDomains"
    return response


@app.get("/")
def root():
    return {
        "service": "Fredonet M365 API",
        "version": app.version,
    }


@app.get("/health")
def health():
    return {"status": "ok"}


@app.get("/me")
@limiter.limit("10/minute")
def me(
    request: Request,
    x_api_key: str | None = Header(default=None),
):
    require_api_key(x_api_key)
    return graph_get("/me")


@app.get("/debug/mail", include_in_schema=False)
@limiter.limit("10/minute")
def debug_mail(
    request: Request,
    x_api_key: str | None = Header(default=None),
):
    require_debug()
    require_api_key(x_api_key)
    return graph_get("/me/mailFolders/inbox/messages?$top=1")


@app.get("/debug/calendar", include_in_schema=False)
@limiter.limit("10/minute")
def debug_calendar(
    request: Request,
    x_api_key: str | None = Header(default=None),
):
    require_debug()
    require_api_key(x_api_key)
    return graph_get("/me/events?$top=1")


@app.get("/mail/unread")
@limiter.limit(DEFAULT_LIMIT)
def unread_mail(
    request: Request,
    x_api_key: str | None = Header(default=None),
):
    require_api_key(x_api_key)
    return {"unread": get_unread_mail_count()}


@app.get("/calendar/next")
@limiter.limit(DEFAULT_LIMIT)
def calendar_next(
    request: Request,
    x_api_key: str | None = Header(default=None),
):
    require_api_key(x_api_key)
    next_event = get_next_event(days_ahead=30)
    if next_event is None:
        return {"event": None}
    return next_event


@app.get("/dashboard")
@limiter.limit("20/minute")
def dashboard(
    request: Request,
    x_api_key: str | None = Header(default=None),
):
    require_api_key(x_api_key)
    return build_dashboard_data()


@app.get("/onedrive/storage")
@limiter.limit(DEFAULT_LIMIT)
def onedrive_storage(
    request: Request,
    x_api_key: str | None = Header(default=None),
):
    require_api_key(x_api_key)
    return get_onedrive_quota()


@app.get("/todo/lists")
@limiter.limit(DEFAULT_LIMIT)
def todo_lists(
    request: Request,
    x_api_key: str | None = Header(default=None),
):
    require_api_key(x_api_key)
    return graph_get("/me/todo/lists")


@app.get("/todo/open")
@limiter.limit(DEFAULT_LIMIT)
def todo_open(
    request: Request,
    x_api_key: str | None = Header(default=None),
):
    require_api_key(x_api_key)
    return get_open_tasks()


@app.get("/calendar/upcoming")
@limiter.limit("10/minute")
def calendar_upcoming(
    request: Request,
    x_api_key: str | None = Header(default=None),
):
    require_api_key(x_api_key)

    now = datetime.now(timezone.utc)
    end = datetime(now.year + 1, now.month, 1, tzinfo=timezone.utc)

    params = {
        "startDateTime": now.isoformat(),
        "endDateTime": end.isoformat(),
        "$orderby": "start/dateTime",
        "$top": 100,
    }

    events = []
    path = f"/me/calendarView?{urlencode(params)}"
    seen_paths = set()

    while True:
        if path in seen_paths:
            raise HTTPException(status_code=502, detail="Invalid calendar pagination")
        seen_paths.add(path)

        result = graph_get(path)
        if result["status"] != 200:
            graph_error(result)

        data = result["data"]
        if not isinstance(data, dict) or not isinstance(data.get("value"), list):
            raise HTTPException(status_code=502, detail="Invalid calendar response")

        for event in data["value"]:
            if isinstance(event, dict):
                events.append(format_calendar_event(event))

        next_link = data.get("@odata.nextLink")
        if next_link is None:
            break
        if not isinstance(next_link, str):
            raise HTTPException(status_code=502, detail="Invalid calendar pagination")
        try:
            next_url = urlsplit(next_link)
        except ValueError as error:
            raise HTTPException(status_code=502, detail="Invalid calendar pagination") from error

        if (
            next_url.scheme != "https"
            or next_url.netloc != "graph.microsoft.com"
            or next_url.path != "/v1.0/me/calendarView"
            or not next_url.query
            or next_url.fragment
        ):
            raise HTTPException(status_code=502, detail="Invalid calendar pagination")
        path = f"/me/calendarView?{next_url.query}"

    return {
        "count": len(events),
        "events": events,
    }


@app.get("/summary")
@limiter.limit(DEFAULT_LIMIT)
def summary(
    request: Request,
    x_api_key: str | None = Header(default=None),
):
    require_api_key(x_api_key)
    dashboard_data = build_dashboard_data()
    return {
        "mail_unread": dashboard_data["mail"]["unread"],
        "todo_open": dashboard_data["todo"]["open"],
        "next_event": dashboard_data["calendar"],
        "onedrive_used_gb": dashboard_data["onedrive"]["used_gb"],
    }
