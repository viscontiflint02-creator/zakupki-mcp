"""Кто работает за ПК: логин Windows и ФИО из Active Directory (без портала и пароля).

Компьютер в домене может спросить AD о текущем пользователе через ADSI. Ответ
кешируется в identity.json: вне корпоративной сети (дома, без VPN) MCP помнит человека.
"""
from __future__ import annotations

import json
import logging
import os
import subprocess
import time
from pathlib import Path

from google_connection import Caller

LOGGER = logging.getLogger("google-sheets-mcp")
CACHE_TTL = 24 * 3600

# [adsisearcher] есть в Windows PowerShell 5.1 на любом ПК; вывод — UTF-8 JSON.
AD_QUERY = r"""
[Console]::OutputEncoding = [Text.Encoding]::UTF8
$login = $env:USERNAME -replace "'", ''
$searcher = [adsisearcher]"(&(objectCategory=person)(objectClass=user)(sAMAccountName=$login))"
'sn','givenname','middlename','displayname','mail' | ForEach-Object { [void]$searcher.PropertiesToLoad.Add($_) }
$found = $searcher.FindOne()
if (-not $found) { exit 2 }
$p = $found.Properties
[pscustomobject]@{
    family = "$($p.sn)"; given = "$($p.givenname)"; middle = "$($p.middlename)"
    display = "$($p.displayname)"; email = "$($p.mail)"
} | ConvertTo-Json -Compress
"""


def person_from_ad(record: dict) -> dict:
    """Фамилия, имя, отчество из атрибутов AD.

    Отчество бывает только в displayName («Фамилия Имя Отчество») — тогда берём его оттуда.
    Если в givenName попала фамилия («Сидоров Сидор»), повтор фамилии отбрасывается.
    """
    family = str(record.get("family") or "").strip()
    given = str(record.get("given") or "").split()
    middle = str(record.get("middle") or "").strip()
    display = str(record.get("display") or "").split()
    if not family and display:
        family = display[0]
    given = [word for word in given if word.casefold() != family.casefold()]
    rest = [word for word in display if word.casefold() != family.casefold()]
    if not given and rest:
        given = rest[:1]
    if not middle and len(rest) >= 2 and given and rest[0].casefold() == given[0].casefold():
        middle = rest[1]
    return {"family": family, "given": given[0] if given else "", "middle": middle.split()[0] if middle else ""}


def short_fio(person: dict) -> str:
    """«Сидоров С.С.» — как в столбце S заявок."""
    initials = "".join(f"{part[0].upper()}." for part in (person.get("given"), person.get("middle")) if part)
    return f"{person.get('family', '')} {initials}".strip()


def _query_ad() -> dict | None:
    if os.name != "nt":
        return None
    try:
        completed = subprocess.run(
            ["powershell.exe", "-NoLogo", "-NoProfile", "-NonInteractive", "-Command", AD_QUERY],
            capture_output=True, timeout=20, creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0))
    except (OSError, subprocess.TimeoutExpired):
        return None
    if completed.returncode != 0 or not completed.stdout.strip():
        return None
    try:
        return json.loads(completed.stdout.decode("utf-8-sig"))
    except ValueError:
        return None


def windows_caller(cache_path: Path, query=None) -> Caller:
    """Caller для локального MCP: логин Windows, ФИО и почта из AD (с кешем на сутки)."""
    query = query or _query_ad
    login = (os.environ.get("USERNAME") or os.environ.get("USER") or "").lower()
    cached = {}
    try:
        cached = json.loads(cache_path.read_text(encoding="utf-8"))
    except (OSError, ValueError):
        pass
    record = cached.get("ad") if cached.get("login") == login else None
    if record is None or time.time() - cached.get("checked", 0) > CACHE_TTL:
        fresh = query()
        if fresh:
            record = fresh
            try:
                cache_path.parent.mkdir(parents=True, exist_ok=True)
                cache_path.write_text(json.dumps({"login": login, "checked": time.time(), "ad": record},
                                                 ensure_ascii=False), encoding="utf-8")
            except OSError:
                LOGGER.warning("Не удалось сохранить %s", cache_path)
    if not record:
        LOGGER.warning("Active Directory недоступен: ФИО неизвестно, только логин %s", login)
        return Caller(login=login)
    person = person_from_ad(record)
    return Caller(login=login, email=str(record.get("email") or ""),
                  name=" ".join(filter(None, person.values())), fio=short_fio(person))
