"""Портал сотрудников: вход через корпоративный OIDC (Keycloak), установщик OpenCode и каталог MCP.

Сотрудник входит на сайт своей учётной записью: логин и пароль он вводит на странице
Keycloak, портал их не видит. Портал выдаёт одноразовую ссылку на установщик Windows
и персональный токен устройства. Этим токеном OpenCode ходит во все MCP-серверы
каталога; сторонние серверы проверяют его через /api/introspect.
"""
from __future__ import annotations

import base64
import hashlib
import hmac
import html
import json
import logging
import os
import re
import secrets
import threading
import time
from dataclasses import dataclass
from datetime import datetime, timezone
from pathlib import Path
from urllib.parse import urlencode, urlsplit

import httpx
import jwt
from starlette.requests import Request
from starlette.responses import HTMLResponse, RedirectResponse, Response
from starlette.routing import Route

from google_connection import Caller, reports_folder_name, save_token

LOGGER = logging.getLogger("google-sheets-mcp")

SESSION_COOKIE = "mcp_portal"
STATE_COOKIE = "mcp_portal_oauth"
SESSION_TTL = 12 * 3600
STATE_TTL = 15 * 60
INSTALL_CODE_TTL = 30 * 60
LAST_USED_RESOLUTION = 3600
METADATA_TTL = 300
SIGNING_ALGORITHMS = ["RS256", "RS384", "RS512", "PS256", "PS384", "PS512", "ES256", "ES384", "ES512"]
INSTALLER_SCRIPT = Path(__file__).resolve().parent / "installer" / "install.ps1"
ID_PATTERN = re.compile(r"[a-z0-9][a-z0-9-]{0,62}")


def _b64(data: bytes) -> str:
    return base64.urlsafe_b64encode(data).rstrip(b"=").decode("ascii")


def _unb64(text: str) -> bytes:
    return base64.urlsafe_b64decode(text + "=" * (-len(text) % 4))


def _same(left, right) -> bool:
    """Сравнение секретов за постоянное время; не падает на не-ASCII, пустое не равно пустому."""
    if not left or not right:
        return False
    return hmac.compare_digest(str(left or "").encode("utf-8"), str(right or "").encode("utf-8"))


def _hash(token: str) -> str:
    return hashlib.sha256(token.encode("utf-8")).hexdigest()


def _json(data, status_code: int = 200) -> Response:
    # ASCII-JSON: Windows PowerShell 5.1 без charset декодирует ответ как Latin-1.
    return Response(json.dumps(data, ensure_ascii=True, indent=2) + "\n", status_code=status_code,
                    media_type="application/json")


def _write_json(path: Path, data) -> None:
    save_token(path, json.dumps(data, ensure_ascii=False, indent=2) + "\n")


NAME_CLAIMS = ("family_name", "given_name", "middle_name", "name")


def _claim_names() -> tuple[str, str]:
    """Claim с отчеством и (необязательный) claim с полным ФИО «Фамилия Имя Отчество».

    По стандарту OIDC отчество — middle_name. Если в Keycloak настроен другой claim
    (например, display_name из AD displayName), его задают OIDC_MIDDLE_NAME_CLAIM /
    OIDC_FULL_NAME_CLAIM.
    """
    return (os.environ.get("OIDC_MIDDLE_NAME_CLAIM", "middle_name").strip(),
            os.environ.get("OIDC_FULL_NAME_CLAIM", "").strip())


def name_claims(claims: dict) -> dict:
    """Поля ФИО ровно в том виде, как их прислал провайдер (для кабинета и диагностики)."""
    keys = [*NAME_CLAIMS, *filter(None, _claim_names())]
    return {key: str(claims[key]).strip() for key in dict.fromkeys(keys) if str(claims.get(key) or "").strip()}


def person_from_claims(claims: dict) -> dict:
    """Фамилия, имя, отчество из claims провайдера входа (Keycloak) — единственный источник.

    Полное ФИО из OIDC_FULL_NAME_CLAIM главнее отдельных полей. Если в given_name попала
    и фамилия («Сидоров Сидор»), повтор фамилии отбрасывается.
    """
    def words(value) -> list[str]:
        return str(value or "").split()

    middle_claim, full_claim = _claim_names()
    full = words(claims.get(full_claim)) if full_claim else []
    if len(full) >= 2:
        return {"family": full[0], "given": full[1], "middle": full[2] if len(full) > 2 else ""}
    family = next(iter(words(claims.get("family_name"))), "")
    given = words(claims.get("given_name"))
    if not family:
        given = words(claims.get("name"))
        family = given.pop(0) if given else ""
    given = [word for word in given if word.casefold() != family.casefold()]
    middle = next(iter(words(claims.get(middle_claim))), "")
    return {"family": family, "given": given[0] if given else "", "middle": middle}


def short_fio(person: dict) -> str:
    """«Сидоров С.С.» — формат ФИО в столбце S заявок; без отчества — «Сидоров И.»."""
    family, given, middle = person.get("family", ""), person.get("given", ""), person.get("middle", "")
    initials = "".join(f"{part[0].upper()}." for part in (given, middle) if part)
    return f"{family} {initials}".strip()


@dataclass
class PortalConfig:
    public_url: str
    state_dir: Path
    oidc_issuer: str
    oidc_client_id: str
    oidc_client_secret: str
    secret: bytes
    oidc_scopes: str = "openid profile email"
    allowed_users: frozenset[str] = frozenset()
    letai_base_url: str = "http://10.19.27.103:4000/v1"
    letai_model: str = "LetAI-ONE"
    service_keys: tuple[tuple[str, str], ...] = ()

    @property
    def secure_cookies(self) -> bool:
        return self.public_url.startswith("https://")

    @property
    def redirect_uri(self) -> str:
        return self.public_url + "/auth/callback"

    @classmethod
    def from_env(cls) -> "PortalConfig | None":
        public_url = os.environ.get("PORTAL_PUBLIC_URL", "").strip().rstrip("/")
        issuer = os.environ.get("OIDC_ISSUER", "").strip().rstrip("/")
        if not public_url or not issuer:
            return None
        for name, value in (("PORTAL_PUBLIC_URL", public_url), ("OIDC_ISSUER", issuer)):
            if urlsplit(value).scheme not in {"http", "https"}:
                raise RuntimeError(f"{name} должен начинаться с http:// или https://")
        if urlsplit(public_url).path or urlsplit(public_url).query:
            raise RuntimeError("PORTAL_PUBLIC_URL указывается без пути, например http://10.0.0.10:8082")
        client_secret = os.environ.get("OIDC_CLIENT_SECRET", "")
        secret_file = os.environ.get("OIDC_CLIENT_SECRET_FILE", "")
        if not client_secret and secret_file:
            client_secret = Path(secret_file).read_text(encoding="utf-8").strip()
        state_dir = Path(os.environ.get("PORTAL_STATE_DIR") or Path(
            os.environ.get("USERS_FILE") or os.environ.get("TOKEN_PATH", ".runtime/token.json")).parent)
        secret = os.environ.get("PORTAL_SECRET", "").encode("utf-8")
        if not secret:
            secret_path = state_dir / "portal-secret"
            if not secret_path.is_file():
                save_token(secret_path, secrets.token_hex(32))
            secret = secret_path.read_text(encoding="ascii").strip().encode("ascii")
        if len(secret) < 32:
            raise RuntimeError("PORTAL_SECRET должен содержать не менее 32 символов")
        service_keys = []
        for item in filter(None, (part.strip() for part in os.environ.get("PORTAL_SERVICE_KEYS", "").split(","))):
            name, _, key = item.partition(":")
            if not name or len(key) < 32:
                raise RuntimeError("PORTAL_SERVICE_KEYS: формат имя:ключ, ключ от 32 символов")
            service_keys.append((name, key))
        allowed = frozenset(part.strip().lower() for part in
                            os.environ.get("PORTAL_ALLOWED_USERS", "").split(",") if part.strip())
        return cls(
            public_url=public_url, state_dir=state_dir, oidc_issuer=issuer,
            oidc_client_id=os.environ.get("OIDC_CLIENT_ID", "mcp-portal"),
            oidc_client_secret=client_secret, secret=secret,
            oidc_scopes=" ".join(os.environ.get("OIDC_SCOPES", "openid profile email").replace(",", " ").split()),
            allowed_users=allowed,
            letai_base_url=os.environ.get("LETAI_BASE_URL", cls.letai_base_url),
            letai_model=os.environ.get("LETAI_MODEL", cls.letai_model),
            service_keys=tuple(service_keys),
        )


class OidcClient:
    """Authorization Code + PKCE + nonce; id_token проверяется по JWKS провайдера."""

    def __init__(self, config: PortalConfig) -> None:
        self.config = config
        self.transport: httpx.AsyncBaseTransport | None = None  # подменяется в тестах
        self._metadata: dict | None = None
        self._metadata_at = 0.0
        self._jwks: dict | None = None

    def _client(self) -> httpx.AsyncClient:
        return httpx.AsyncClient(timeout=15, transport=self.transport)

    async def metadata(self) -> dict:
        if self._metadata and time.monotonic() - self._metadata_at < METADATA_TTL:
            return self._metadata
        async with self._client() as client:
            response = await client.get(self.config.oidc_issuer + "/.well-known/openid-configuration")
            response.raise_for_status()
            metadata = response.json()
        if metadata.get("issuer", "").rstrip("/") != self.config.oidc_issuer:
            raise RuntimeError("OIDC discovery вернул другой issuer")
        self._metadata, self._metadata_at, self._jwks = metadata, time.monotonic(), None
        return metadata

    async def _signing_key(self, kid: str | None, refresh: bool = False):
        if self._jwks is None or refresh:
            async with self._client() as client:
                response = await client.get((await self.metadata())["jwks_uri"])
                response.raise_for_status()
                self._jwks = response.json()
        for data in self._jwks.get("keys", []):
            if data.get("use", "sig") == "sig" and (kid is None or data.get("kid") == kid):
                return jwt.PyJWK.from_dict(data).key
        if not refresh:  # ключи провайдера могли смениться
            return await self._signing_key(kid, refresh=True)
        raise RuntimeError("Не найден ключ подписи id_token")

    async def authorize_url(self, state: str, nonce: str, verifier: str) -> str:
        endpoint = (await self.metadata())["authorization_endpoint"]
        return endpoint + "?" + urlencode({
            "client_id": self.config.oidc_client_id, "response_type": "code",
            "redirect_uri": self.config.redirect_uri, "scope": self.config.oidc_scopes,
            "state": state, "nonce": nonce,
            "code_challenge": _b64(hashlib.sha256(verifier.encode("ascii")).digest()),
            "code_challenge_method": "S256",
        })

    async def exchange(self, code: str, verifier: str, nonce: str) -> dict:
        """Код → токены → проверенные claims пользователя (id_token + userinfo)."""
        metadata = await self.metadata()
        data = {"grant_type": "authorization_code", "code": code,
                "redirect_uri": self.config.redirect_uri, "code_verifier": verifier,
                "client_id": self.config.oidc_client_id}
        auth = None
        methods = metadata.get("token_endpoint_auth_methods_supported") or ["client_secret_post"]
        if self.config.oidc_client_secret:
            if "client_secret_post" in methods:
                data["client_secret"] = self.config.oidc_client_secret
            else:
                auth = (self.config.oidc_client_id, self.config.oidc_client_secret)
        async with self._client() as client:
            response = await client.post(metadata["token_endpoint"], data=data, auth=auth)
            if response.status_code != 200:
                LOGGER.warning("OIDC token exchange: %s %s", response.status_code, response.text[:300])
                raise RuntimeError("Провайдер входа не подтвердил код авторизации")
            tokens = response.json()
            if not tokens.get("id_token"):
                raise RuntimeError("Провайдер входа не вернул id_token (нужен scope openid)")
            header = jwt.get_unverified_header(tokens["id_token"])
            if header.get("alg") not in SIGNING_ALGORITHMS:
                raise RuntimeError("Недопустимый алгоритм подписи id_token")
            claims = jwt.decode(
                tokens["id_token"], key=await self._signing_key(header.get("kid")),
                algorithms=[header["alg"]], audience=self.config.oidc_client_id,
                issuer=metadata["issuer"], leeway=60,
                options={"require": ["exp", "iat", "iss", "aud", "sub"]})
            if not _same(claims.get("nonce"), nonce):
                raise RuntimeError("nonce id_token не совпал")
            if metadata.get("userinfo_endpoint") and tokens.get("access_token"):
                info = await client.get(metadata["userinfo_endpoint"],
                                        headers={"Authorization": "Bearer " + tokens["access_token"]})
                if info.status_code == 200 and info.json().get("sub") == claims["sub"]:
                    claims = {**info.json(), **claims}
        return claims


class PortalStore:
    """Пользователи, токены устройств и одноразовые коды установки в одном JSON.

    Токены хранятся только как SHA-256. Файл перечитывается по mtime, чтобы
    portal_admin.py мог блокировать пользователей без перезапуска службы.
    """

    def __init__(self, state_dir: Path) -> None:
        self.state_dir = state_dir
        self.path = state_dir / "portal.json"
        self._lock = threading.RLock()
        self._mtime = None
        self._data = {"users": {}, "tokens": {}, "codes": {}}

    def _load(self) -> None:
        try:
            mtime = self.path.stat().st_mtime_ns
        except OSError:
            mtime = None
        if mtime == self._mtime:
            return
        data = {"users": {}, "tokens": {}, "codes": {}}
        if mtime is not None:
            with self.path.open(encoding="utf-8") as stream:
                data.update(json.load(stream))
        self._data, self._mtime = data, mtime

    def _save(self) -> None:
        now = time.time()
        self._data["codes"] = {key: value for key, value in self._data["codes"].items()
                               if value.get("expires", 0) > now}
        _write_json(self.path, self._data)
        self._mtime = self.path.stat().st_mtime_ns

    def user(self, sub: str) -> dict | None:
        with self._lock:
            self._load()
            user = self._data["users"].get(sub)
            return dict(user) if user else None

    def users(self) -> dict[str, dict]:
        with self._lock:
            self._load()
            return {sub: dict(user) for sub, user in self._data["users"].items()}

    def save_login(self, claims: dict) -> dict:
        sub = str(claims["sub"])
        person = person_from_claims(claims)
        with self._lock:
            self._load()
            now = datetime.now(timezone.utc).isoformat(timespec="seconds")
            user = self._data["users"].setdefault(sub, {"created": now, "servers": None, "disabled": False})
            user.pop("person", None)
            # ФИО обновляется при каждом входе: исправили в Keycloak — исправится и здесь.
            user.update(login=str(claims.get("preferred_username") or sub).lower(),
                        name=" ".join(filter(None, person.values())) or str(claims.get("name") or "").strip(),
                        fio=short_fio(person), has_middle_name=bool(person["middle"]),
                        name_claims=name_claims(claims),
                        email=str(claims.get("email") or "").lower(), last_login=now)
            self._save()
            return dict(user)

    def set_servers(self, sub: str, servers: list[str]) -> None:
        with self._lock:
            self._load()
            self._data["users"][sub]["servers"] = servers
            self._save()

    def set_disabled(self, login_or_email: str, disabled: bool) -> int:
        """Блокирует по логину или почте; при блокировке отзывает все токены ПК."""
        needle = login_or_email.lower()
        changed = 0
        with self._lock:
            self._load()
            for sub, user in self._data["users"].items():
                if needle in {user.get("login"), user.get("email")}:
                    user["disabled"] = disabled
                    changed += 1
                    if disabled:
                        self._data["tokens"] = {key: value for key, value in self._data["tokens"].items()
                                                if value["sub"] != sub}
            if changed:
                self._save()
        return changed

    def create_install_code(self, sub: str) -> str:
        code = secrets.token_urlsafe(24)
        with self._lock:
            self._load()
            self._data["codes"][_hash(code)] = {"sub": sub, "expires": time.time() + INSTALL_CODE_TTL}
            self._save()
        return code

    def claim_install_code(self, code: str, device: str) -> tuple[dict, str] | None:
        """Одноразовый код установщика → новый токен устройства."""
        with self._lock:
            self._load()
            entry = self._data["codes"].pop(_hash(code), None)
            if not entry or entry["expires"] < time.time():
                if entry:
                    self._save()
                return None
            user = self._data["users"].get(entry["sub"])
            if not user or user.get("disabled"):
                self._save()
                return None
            token = secrets.token_urlsafe(32)
            now = time.time()
            self._data["tokens"][_hash(token)] = {"sub": entry["sub"], "device": device[:64] or "ПК",
                                                  "created": now, "last_used": now}
            self._save()
            return dict(user, sub=entry["sub"]), token

    def user_for_token(self, token: str) -> tuple[str, dict] | None:
        key = _hash(token)
        with self._lock:
            self._load()
            entry = self._data["tokens"].get(key)
            if not entry:
                return None
            user = self._data["users"].get(entry["sub"])
            if not user or user.get("disabled"):
                return None
            if time.time() - entry.get("last_used", 0) > LAST_USED_RESOLUTION:
                entry["last_used"] = time.time()
                self._save()
            return entry["sub"], dict(user)

    def caller_for_token(self, token: str) -> Caller | None:
        found = self.user_for_token(token)
        if found is None:
            return None
        sub, user = found
        return Caller(login=user.get("login") or sub, email=user.get("email", ""),
                      name=user.get("name", ""), fio=user.get("fio", ""))

    def devices(self, sub: str) -> list[dict]:
        with self._lock:
            self._load()
            return sorted(({"id": key[:12], **value} for key, value in self._data["tokens"].items()
                           if value["sub"] == sub), key=lambda item: item["created"])

    def revoke_device(self, sub: str, device_id: str) -> bool:
        with self._lock:
            self._load()
            for key, value in list(self._data["tokens"].items()):
                if value["sub"] == sub and key[:12] == device_id:
                    del self._data["tokens"][key]
                    self._save()
                    return True
        return False


class Catalog:
    """Каталог MCP-серверов: catalog.json в каталоге состояния, иначе встроенный по умолчанию."""

    def __init__(self, path: Path, public_url: str) -> None:
        self.path = path
        self.public_url = public_url
        self._mtime = None
        self._servers: list[dict] = []

    def default(self) -> list[dict]:
        return [{
            "id": "google-sheets",
            "name": "Google Таблицы и закупки",
            "description": "Заявки «В работе», ОКПД2, черновики ТЗ и Google-таблицы в вашей папке.",
            "url": "{PUBLIC_URL}/mcp",
            "auth": "portal",
            "required": True,
            "default": True,
            "timeout": 60000,
        }]

    def servers(self) -> list[dict]:
        try:
            mtime = self.path.stat().st_mtime_ns
        except OSError:
            mtime = None
        if mtime != self._mtime or not self._servers:
            raw = self.default()
            if mtime is not None:
                with self.path.open(encoding="utf-8-sig") as stream:
                    raw = json.load(stream).get("servers", [])
            self._servers = [self._validate(item) for item in raw]
            self._mtime = mtime
        return self._servers

    def _validate(self, item: dict) -> dict:
        server = dict(item)
        # servers/timeout — служебные ключи секции mcp в конфиге OpenCode.
        if not ID_PATTERN.fullmatch(str(server.get("id", ""))) or server.get("id") in {"servers", "timeout"}:
            raise RuntimeError(f"catalog.json: некорректный id {server.get('id')!r}")
        server["url"] = str(server.get("url", "")).replace("{PUBLIC_URL}", self.public_url)
        if urlsplit(server["url"]).scheme not in {"http", "https"}:
            raise RuntimeError(f"catalog.json: некорректный url у {server['id']}")
        server.setdefault("auth", "portal")
        if server["auth"] not in {"portal", "oauth", "none"}:
            raise RuntimeError(f"catalog.json: auth у {server['id']} — portal, oauth или none")
        server.setdefault("name", server["id"])
        server.setdefault("description", "")
        server["required"] = bool(server.get("required"))
        server["default"] = bool(server.get("default", server["required"]))
        return server

    def selected(self, choice: list[str] | None) -> list[dict]:
        servers = self.servers()
        if choice is None:
            return [server for server in servers if server["default"] or server["required"]]
        return [server for server in servers if server["required"] or server["id"] in choice]


class Portal:
    def __init__(self, config: PortalConfig) -> None:
        self.config = config
        self.store = PortalStore(config.state_dir)
        self.catalog = Catalog(config.state_dir / "catalog.json", config.public_url)
        self.oidc = OidcClient(config)

    # --- подписанные cookie -------------------------------------------------
    def _sign(self, payload: dict) -> str:
        body = _b64(json.dumps(payload, separators=(",", ":")).encode("utf-8"))
        return body + "." + _b64(hmac.new(self.config.secret, body.encode("ascii"), hashlib.sha256).digest())

    def _unsign(self, value: str | None) -> dict | None:
        if not value or "." not in value:
            return None
        body, _, signature = value.partition(".")
        expected = _b64(hmac.new(self.config.secret, body.encode("ascii"), hashlib.sha256).digest())
        if not _same(signature, expected):
            return None
        try:
            payload = json.loads(_unb64(body))
        except ValueError:
            return None
        return payload if payload.get("exp", 0) > time.time() else None

    def _set_cookie(self, response: Response, name: str, payload: dict, ttl: int) -> None:
        response.set_cookie(name, self._sign(dict(payload, exp=time.time() + ttl)), max_age=ttl,
                            httponly=True, secure=self.config.secure_cookies, samesite="lax", path="/")

    def _session_user(self, request: Request) -> tuple[str, dict] | None:
        payload = self._unsign(request.cookies.get(SESSION_COOKIE))
        if not payload:
            return None
        user = self.store.user(str(payload.get("sub")))
        if not user or user.get("disabled"):
            return None
        return str(payload["sub"]), user

    def _csrf(self, sub: str) -> str:
        return hmac.new(self.config.secret, b"csrf:" + sub.encode("utf-8"), hashlib.sha256).hexdigest()[:32]

    async def _checked_form(self, request: Request) -> tuple[str, dict, dict] | None:
        session = self._session_user(request)
        if session is None:
            return None
        form = await request.form()
        if not _same(form.get("csrf"), self._csrf(session[0])):
            return None
        return session[0], session[1], form

    # --- страницы -----------------------------------------------------------
    def _page(self, title: str, body: str, status_code: int = 200) -> HTMLResponse:
        return HTMLResponse(PAGE_TEMPLATE.format(title=html.escape(title), body=body), status_code=status_code,
                            headers={"Cache-Control": "no-store", "X-Frame-Options": "DENY",
                                     "Referrer-Policy": "no-referrer"})

    def _error(self, message: str, status_code: int = 400) -> HTMLResponse:
        return self._page("Ошибка", f'<div class="card"><h1>Не получилось</h1><p>{html.escape(message)}</p>'
                                    '<p><a class="button" href="/">На главную</a></p></div>', status_code)

    async def index(self, request: Request) -> Response:
        if self._session_user(request):
            return RedirectResponse("/me", status_code=303)
        folder = html.escape(reports_folder_name() or "Отчёты MCP")
        return self._page("Вход", f"""
<div class="card hero">
  <h1>MCP-серверы для OpenCode</h1>
  <p>Войдите своей корпоративной учётной записью (логин и пароль вводятся на странице единого входа).
  После входа вы скачаете установщик: он поставит OpenCode (если его ещё нет) и подключит MCP-серверы.
  Google-таблицы, созданные для вас, будут лежать в вашей папке «{folder}».</p>
  <p><a class="button primary" href="/auth/login">Войти</a></p>
</div>""")

    async def login(self, _request: Request) -> Response:
        state, nonce, verifier = secrets.token_urlsafe(24), secrets.token_urlsafe(24), secrets.token_urlsafe(48)
        try:
            target = await self.oidc.authorize_url(state, nonce, verifier)
        except Exception:
            LOGGER.exception("OIDC discovery")
            return self._error("Сервер единого входа недоступен. Попробуйте позже.", 502)
        response = RedirectResponse(target, status_code=303)
        self._set_cookie(response, STATE_COOKIE, {"state": state, "nonce": nonce, "verifier": verifier}, STATE_TTL)
        return response

    async def callback(self, request: Request) -> Response:
        saved = self._unsign(request.cookies.get(STATE_COOKIE))
        query = request.query_params
        if query.get("error"):
            return self._error("Вход отклонён: " + query.get("error_description", query["error"]))
        if not saved or not _same(saved.get("state"), query.get("state")):
            return self._error("Сессия входа устарела. Начните вход заново.")
        if not query.get("code"):
            return self._error("Сервер входа не вернул код авторизации.")
        try:
            claims = await self.oidc.exchange(query["code"], saved["verifier"], saved["nonce"])
        except Exception as error:
            LOGGER.warning("OIDC callback: %s", error)
            return self._error("Не удалось подтвердить вход. Попробуйте ещё раз.", 502)
        login = str(claims.get("preferred_username") or "").lower()
        if self.config.allowed_users and login not in self.config.allowed_users:
            return self._error(f"Учётной записи {login} не выдан доступ к порталу.", 403)
        existing = self.store.user(str(claims["sub"]))
        if existing and existing.get("disabled"):
            return self._error("Доступ этой учётной записи к порталу отключён администратором.", 403)
        user = self.store.save_login(claims)
        # Для диагностики: какие поля прислал провайдер (без значений) и есть ли отчество.
        LOGGER.info("Портал: вход %s, ФИО %s%s, claims: %s", user["login"], user["fio"],
                    "" if user["has_middle_name"] else " (нет отчества в Keycloak)",
                    ",".join(sorted(key for key in claims if key not in {"nonce", "at_hash", "sid"})))
        response = RedirectResponse("/me", status_code=303)
        response.delete_cookie(STATE_COOKIE, path="/")
        self._set_cookie(response, SESSION_COOKIE, {"sub": str(claims["sub"])}, SESSION_TTL)
        return response

    async def logout(self, request: Request) -> Response:
        if await self._checked_form(request) is None:
            return self._error("Сессия устарела.", 403)
        response = RedirectResponse("/", status_code=303)
        response.delete_cookie(SESSION_COOKIE, path="/")
        return response

    async def me(self, request: Request) -> Response:
        session = self._session_user(request)
        if session is None:
            return RedirectResponse("/", status_code=303)
        sub, user = session
        csrf = html.escape(self._csrf(sub))
        selected = {server["id"] for server in self.catalog.selected(user.get("servers"))}
        rows = []
        for server in self.catalog.servers():
            checked = " checked" if server["id"] in selected else ""
            disabled = " disabled" if server["required"] else ""
            note = " <span class=\"muted\">(обязательный)</span>" if server["required"] else ""
            rows.append(f'<label class="server"><input type="checkbox" name="server" value="{html.escape(server["id"])}"'
                        f'{checked}{disabled}> <span><b>{html.escape(server["name"])}</b>{note}<br>'
                        f'<span class="muted">{html.escape(server["description"])}</span></span></label>')
        devices = []
        for device in self.store.devices(sub):
            created = datetime.fromtimestamp(device["created"]).strftime("%d.%m.%Y %H:%M")
            used = datetime.fromtimestamp(device["last_used"]).strftime("%d.%m.%Y")
            devices.append(f'<li><b>{html.escape(device["device"])}</b> — подключён {created}, активен {used} '
                           f'<form method="post" action="/me/devices/revoke" class="inline">'
                           f'<input type="hidden" name="csrf" value="{csrf}">'
                           f'<input type="hidden" name="device" value="{html.escape(device["id"])}">'
                           f'<button class="link">отключить</button></form></li>')
        devices_html = "<ul>" + "".join(devices) + "</ul>" if devices else '<p class="muted">Пока нет.</p>'
        title = html.escape(user.get("name") or user.get("login", ""))
        subtitle = html.escape(" · ".join(filter(None, [user.get("login"), user.get("email")])))
        received = "".join(f"<li><code>{html.escape(key)}</code>: {html.escape(value)}</li>"
                           for key, value in (user.get("name_claims") or {}).items())
        warning = ("" if user.get("has_middle_name") else
                   '<p class="muted">Учётная запись не передаёт отчество (claim <code>middle_name</code>). '
                   'Попросите администратора Keycloak добавить его — инструкция в README портала. '
                   'После исправления выйдите и войдите заново.</p>')
        return self._page("Личный кабинет", f"""
<div class="card">
  <div class="top"><div><h1>{title}</h1><p class="muted">{subtitle}</p></div>
  <form method="post" action="/logout"><input type="hidden" name="csrf" value="{csrf}"><button class="link">Выйти</button></form></div>
  <p>ФИО для заявок: <b>{html.escape(user.get("fio") or "не определено")}</b></p>
  {warning}
  <details><summary class="muted">Что прислал сервер входа (Keycloak)</summary><ul>{received}</ul></details>
  <h2>1. Установите OpenCode и MCP</h2>
  <p>Скачайте и запустите установщик на рабочем ПК с Windows. Если OpenCode уже установлен, этот шаг
  пропустится. Установщик попросит ваш API-ключ LetAI (Enter — оставить текущий) и подключит отмеченные ниже MCP-серверы.
  Ссылка одноразовая и действует 30 минут.</p>
  <p><a class="button primary" href="/installer.bat">Скачать установщик</a></p>
  <h2>2. MCP-серверы</h2>
  <form method="post" action="/me/servers">
    <input type="hidden" name="csrf" value="{csrf}">
    {''.join(rows)}
    <p><button class="button">Сохранить выбор</button></p>
  </form>
  <p class="muted">Новые серверы подключаются без переустановки: установщик ставит задачу
  «MCP Portal sync», она обновляет список при входе в Windows. Обновить сразу можно файлом
  <code>%USERPROFILE%\\.config\\mcp-portal\\mcp-sync.cmd</code>; затем перезапустите OpenCode.
  Ссылку на вашу папку с таблицами ИИ даст по команде «кто я».</p>
  <h2>Подключённые ПК</h2>
  {devices_html}
</div>""")

    async def save_servers(self, request: Request) -> Response:
        checked = await self._checked_form(request)
        if checked is None:
            return self._error("Сессия устарела, обновите страницу.", 403)
        sub, _user, form = checked
        known = {server["id"] for server in self.catalog.servers()}
        self.store.set_servers(sub, [item for item in form.getlist("server") if item in known])
        return RedirectResponse("/me", status_code=303)

    async def revoke_device(self, request: Request) -> Response:
        checked = await self._checked_form(request)
        if checked is None:
            return self._error("Сессия устарела, обновите страницу.", 403)
        self.store.revoke_device(checked[0], str(checked[2].get("device", "")))
        return RedirectResponse("/me", status_code=303)

    async def installer_bat(self, request: Request) -> Response:
        session = self._session_user(request)
        if session is None:
            return RedirectResponse("/", status_code=303)
        code = self.store.create_install_code(session[0])
        portal = self.config.public_url
        command = (
            "$ErrorActionPreference='Stop'; "
            "[Net.ServicePointManager]::SecurityProtocol=[Net.SecurityProtocolType]::Tls12; "
            f"$r=Invoke-WebRequest -UseBasicParsing -Uri '{portal}/install.ps1'; "
            "$s=[Text.Encoding]::UTF8.GetString($r.RawContentStream.ToArray()); "
            f"& ([scriptblock]::Create($s)) -Portal '{portal}' -Code '{code}'"
        )
        body = "\r\n".join([
            "@echo off", "chcp 65001 >nul", "title OpenCode + MCP setup",
            f'powershell.exe -NoLogo -NoProfile -ExecutionPolicy Bypass -Command "{command}"',
            'set "RC=%ERRORLEVEL%"', "echo.", "pause", "exit /b %RC%", ""])
        return Response(body.encode("ascii"), media_type="application/octet-stream", headers={
            "Content-Disposition": 'attachment; filename="mcp-setup.bat"', "Cache-Control": "no-store"})

    async def install_script(self, _request: Request) -> Response:
        return Response(INSTALLER_SCRIPT.read_bytes(), media_type="text/plain; charset=utf-8",
                        headers={"Cache-Control": "no-store"})

    # --- API для установщика и других MCP-серверов -----------------------------
    async def api_claim(self, request: Request) -> Response:
        try:
            data = await request.json()
        except ValueError:
            return _json({"error": "invalid_request"}, 400)
        claimed = self.store.claim_install_code(str(data.get("code", "")), str(data.get("device", "")))
        if claimed is None:
            return _json({"error": "invalid_code",
                          "message": "Ссылка установщика устарела или уже использована. "
                                     "Скачайте установщик на портале заново."}, 400)
        user, token = claimed
        return _json({"token": token, "login": user.get("login"), "email": user.get("email"),
                      "name": user.get("name") or user.get("fio"), "portal": self.config.public_url})

    def _bearer(self, request: Request) -> str:
        scheme, _, credential = request.headers.get("authorization", "").partition(" ")
        return credential.strip() if scheme.lower() == "bearer" else ""

    def opencode_config(self, user: dict, letai: bool, schema: str = "v1") -> dict:
        """Фрагменты для основного конфига OpenCode; установщик вписывает их сам.

        v1 (OpenCode 1.x): mcp.<имя>, enabled, timeout — число; LetAI в provider.<имя>
        через npm-пакет @ai-sdk/openai-compatible и options.
        v2 (OpenCode 2.x): mcp.servers.<имя>, disabled, timeout — объект; LetAI в
        providers.<имя> через package @opencode/ai/providers/openai-compatible и settings.
        Токен ПК и ключ LetAI подставляет сам OpenCode из файлов: в конфиг они не попадают.
        provider_key говорит установщику, в какую секцию конфига вписать провайдера.
        """
        servers = {}
        for server in self.catalog.selected(user.get("servers")):
            timeout = int(server.get("timeout", 30000))
            entry = {"type": "remote", "url": server["url"]}
            if schema == "v2":
                entry["timeout"] = {"startup": 30000, "catalog": timeout, "execution": 600000}
            else:
                entry.update(enabled=True, timeout=timeout)
            if server["auth"] in {"portal", "none"}:
                entry["oauth"] = False
            if server["auth"] == "portal":
                entry["headers"] = {"Authorization": "Bearer {file:~/.config/mcp-portal/portal.token}"}
            servers[server["id"]] = entry
        result = {"schema": schema, "servers": servers}
        if letai:
            api_key = "{file:~/.config/mcp-portal/letai.key}"
            limit = {"context": 262144, "output": 65536}
            if schema == "v2":
                result["provider_key"] = "providers"
                result["provider"] = {"letai-one": {
                    "name": "LetAI-ONE", "package": "@opencode/ai/providers/openai-compatible",
                    "settings": {"baseURL": self.config.letai_base_url, "apiKey": api_key},
                    "models": {self.config.letai_model: {"name": self.config.letai_model, "limit": limit}},
                }}
            else:
                result["provider_key"] = "provider"
                result["provider"] = {"letai-one": {
                    "npm": "@ai-sdk/openai-compatible", "name": "LetAI-ONE",
                    "options": {"baseURL": self.config.letai_base_url, "apiKey": api_key, "timeout": 600000},
                    "models": {self.config.letai_model: {
                        "name": self.config.letai_model,
                        "modalities": {"input": ["text", "image"], "output": ["text"]},
                        "limit": limit}},
                }}
            result["model"] = f"letai-one/{self.config.letai_model}"
        return result

    async def api_config(self, request: Request) -> Response:
        found = self.store.user_for_token(self._bearer(request))
        if found is None:
            return _json({"error": "unauthorized"}, 401)
        schema = "v2" if request.query_params.get("schema") == "v2" else "v1"
        return _json(self.opencode_config(found[1], request.query_params.get("letai") == "1", schema))

    async def api_introspect(self, request: Request) -> Response:
        """Для других MCP-серверов: чей это токен устройства (RFC 7662-подобно)."""
        presented = self._bearer(request)
        if not presented or not any(_same(presented, key) for _name, key in self.config.service_keys):
            return _json({"error": "unauthorized"}, 401)
        data = await request.json()
        found = self.store.user_for_token(str(data.get("token", "")))
        if found is None:
            return _json({"active": False})
        sub, user = found
        return _json({"active": True, "sub": sub, "login": user.get("login"), "fio": user.get("fio"),
                      "email": user.get("email"), "name": user.get("name")})

    def routes(self) -> list[Route]:
        return [
            Route("/", self.index, methods=["GET"]),
            Route("/auth/login", self.login, methods=["GET"]),
            Route("/auth/callback", self.callback, methods=["GET"]),
            Route("/logout", self.logout, methods=["POST"]),
            Route("/me", self.me, methods=["GET"]),
            Route("/me/servers", self.save_servers, methods=["POST"]),
            Route("/me/devices/revoke", self.revoke_device, methods=["POST"]),
            Route("/installer.bat", self.installer_bat, methods=["GET"]),
            Route("/install.ps1", self.install_script, methods=["GET"]),
            Route("/api/install/claim", self.api_claim, methods=["POST"]),
            Route("/api/opencode/config", self.api_config, methods=["GET"]),
            Route("/api/introspect", self.api_introspect, methods=["POST"]),
        ]


PAGE_TEMPLATE = """<!doctype html>
<html lang="ru"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width, initial-scale=1">
<title>{title} — MCP-портал</title>
<style>
:root {{ --bg:#f4f6fa; --card:#fff; --text:#1d2433; --muted:#667085; --accent:#2a5c8f; --line:#e3e7ef; }}
@media (prefers-color-scheme: dark) {{ :root {{ --bg:#11151c; --card:#1a2029; --text:#e6e9ef; --muted:#98a2b3; --accent:#6aa5e0; --line:#2b3340; }} }}
* {{ box-sizing:border-box; }}
body {{ margin:0; background:var(--bg); color:var(--text); font:16px/1.5 system-ui,-apple-system,"Segoe UI",sans-serif; }}
main {{ max-width:760px; margin:40px auto; padding:0 16px; }}
.card {{ background:var(--card); border:1px solid var(--line); border-radius:12px; padding:28px; }}
h1 {{ margin:0 0 4px; font-size:26px; }} h2 {{ margin:28px 0 8px; font-size:19px; }}
.muted {{ color:var(--muted); }} .top {{ display:flex; justify-content:space-between; gap:16px; align-items:flex-start; }}
.button {{ display:inline-block; border:1px solid var(--accent); color:var(--accent); background:transparent; padding:10px 18px;
  border-radius:8px; text-decoration:none; font:inherit; cursor:pointer; }}
.button.primary {{ background:var(--accent); color:#fff; }}
.link {{ border:0; background:none; color:var(--accent); cursor:pointer; font:inherit; padding:0; text-decoration:underline; }}
.server {{ display:flex; gap:10px; padding:10px 0; border-top:1px solid var(--line); }}
.server input {{ margin-top:5px; }} .inline {{ display:inline; }}
code {{ font-size:14px; word-break:break-all; }}
</style></head><body><main>{body}</main></body></html>"""
