"""MCP на ПК сотрудника: заявки из Google-таблицы, ФИО из AD, цены из интернета, stdio, установщик.

Google, AD и интернет подменены; MCP запускается настоящим stdio-клиентом в отдельном процессе.
"""
import asyncio
import json
import os
from pathlib import Path
import re
import shutil
import subprocess
import sys
import tempfile
import textwrap
import unittest
from unittest.mock import MagicMock

import httpx
from mcp import ClientSession, StdioServerParameters
from mcp.client.stdio import stdio_client

from local_identity import person_from_ad, short_fio, windows_caller
from price_search import page_prices, parse_results, search_prices
from technical_spec import PriceItem, build_price_rows
from work_report import GoogleSheetReport, positions_from_rows

ROOT = Path(__file__).resolve().parents[1]
POWERSHELL = shutil.which("powershell.exe")


def sheet_row(status, item, responsible, quantity="2", unit="шт", number="100"):
    """Строка листа «незакр 2025/26», как её отдаёт Google API: строки, без хвостовых пустых."""
    row = [""] * 19
    row[0], row[3], row[8], row[9], row[10], row[18] = status, number, item, unit, quantity, responsible
    return row


SHEET = [
    sheet_row("В работе", "Фаркоп для УАЗ Патриот", "Кузнецова А.Б./Сидоров С.С."),
    sheet_row("Завершено", "Перчатки нитриловые", "Кузнецова А.Б./Сидоров С.С."),
    sheet_row("в работе", "Картридж HP 85A", "Кузнецова А.Б./Сидоров С.С.", quantity="5"),
    sheet_row("В работе", "Бумага А4", "Иванов И.И./Петров П.П."),
]


class GoogleSheetSourceTests(unittest.TestCase):
    def service(self, rows):
        service = MagicMock()
        service.spreadsheets.return_value.values.return_value.get.return_value.execute.return_value = {"values": rows}
        return service

    def test_positions_and_version_from_live_sheet(self):
        service = self.service(SHEET)
        report = GoogleSheetReport("sheet-id", "незакр 2025/26", ROOT / "data" / "okpd.xlsx", lambda: service)
        mine = report.positions("Сидоров С.С.", "Кузнецова А.Б.")
        self.assertEqual([p["source_row"] for p in mine], [2, 4])
        self.assertEqual(mine[1]["quantity"], "5")
        call = service.spreadsheets.return_value.values.return_value.get.call_args.kwargs
        self.assertEqual(call["range"], "'незакр 2025/26'!A2:AC")
        version = report.source_version()
        # Правка комментария не меняет версию, вставка строки выше — меняет.
        report._fetched_at = None
        edited = [list(row) for row in SHEET]
        edited[0] = edited[0] + ["", "", "новый комментарий"]
        report._sheets_service = lambda: self.service(edited)
        self.assertEqual(report.source_version(), version)
        report._fetched_at = None
        report._sheets_service = lambda: self.service([sheet_row("Завершено", "x", "y")] + SHEET)
        self.assertNotEqual(report.source_version(), version)

    def test_short_rows_are_padded(self):
        self.assertEqual(positions_from_rows([["В работе", "", "", "", "", "", "", "", "Товар"]])[0]["item_name"], "Товар")


class IdentityTests(unittest.TestCase):
    def test_person_from_ad(self):
        cases = [
            ({"family": "Сидоров", "given": "Сидор", "middle": "Сидорович"}, "Сидоров С.С."),
            ({"family": "Сидоров", "given": "Сидоров Сидор",
              "display": "Сидоров Сидор Сидорович"}, "Сидоров С.С."),
            ({"family": "Кузнецова", "given": "Анна", "display": "Кузнецова Анна Борисовна"}, "Кузнецова А.Б."),
            ({"display": "Петров Пётр"}, "Петров П."),
        ]
        for record, fio in cases:
            with self.subTest(record=record):
                self.assertEqual(short_fio(person_from_ad(record)), fio)

    def test_windows_caller_caches_ad(self):
        with tempfile.TemporaryDirectory() as directory:
            cache = Path(directory) / "identity.json"
            os.environ["USERNAME"] = "SIDOROV"
            record = {"family": "Сидоров", "given": "Сидор", "middle": "Сидорович", "email": "n@example.ru"}
            caller = windows_caller(cache, query=lambda: record)
            self.assertEqual((caller.login, caller.fio, caller.email), ("sidorov", "Сидоров С.С.", "n@example.ru"))
            # Вне корпоративной сети AD не отвечает — берём из кеша.
            again = windows_caller(cache, query=lambda: None)
            self.assertEqual(again.fio, "Сидоров С.С.")
            cache.unlink()
            self.assertEqual(windows_caller(cache, query=lambda: None).fio, "")


DDG_PAGE = """
<div class="result"><a class="result__a" href="//duckduckgo.com/l/?uddg=https%3A%2F%2Fshop.example%2Ffarkop&rut=1">
Фаркоп УАЗ Патриот — 12&nbsp;990 ₽</a>
<a class="result__snippet" href="#">Тягово-сцепное устройство, цена от 11 500 руб.</a></div>
<div class="result"><a class="result__a" href="https://catalog.example/farkop">Каталог фаркопов</a>
<a class="result__snippet" href="#">Большой выбор</a></div>
"""
PRODUCT_PAGE = '<script type="application/ld+json">{"@type":"Product","offers":{"price":"14250.00"}}</script>'


class PriceTests(unittest.TestCase):
    def test_parse_and_page_prices(self):
        results = parse_results(DDG_PAGE, 5)
        self.assertEqual(results[0]["url"], "https://shop.example/farkop")
        self.assertEqual(results[0]["prices_rub"], [12990.0, 11500.0])
        self.assertEqual(results[1]["prices_rub"], [])
        self.assertEqual(page_prices(PRODUCT_PAGE), ([14250.0], "разметка товара"))
        self.assertEqual(page_prices("<p>Цена: 1 299,90 руб.</p>")[0], [1299.9])

    def test_search_prices_opens_pages_without_prices(self):
        def handler(request: httpx.Request) -> httpx.Response:
            if "duckduckgo" in request.url.host:
                return httpx.Response(200, text=DDG_PAGE)
            if request.url.host == "catalog.example":
                return httpx.Response(200, text=PRODUCT_PAGE, headers={"content-type": "text/html"})
            return httpx.Response(404)
        result = search_prices("Фаркоп УАЗ Патриот", transport=httpx.MockTransport(handler))
        self.assertEqual(result["results"][1]["prices_rub"], [14250.0])
        self.assertEqual(result["summary_rub"]["min"], 11500.0)
        self.assertIn("не коммерческое предложение", result["note"])

    def test_price_rows(self):
        positions = positions_from_rows(SHEET)
        rows = build_price_rows(positions, [PriceItem(source_row=4, unit_price=1500, source_url="https://shop.example/hp")])
        self.assertEqual(rows[1][:7], [4, "100", "Картридж HP 85A", "шт", "5", 1500.0, 7500.0])
        with self.assertRaises(ValueError):
            build_price_rows(positions, [PriceItem(source_row=99, unit_price=1, source_url="https://x.example")])


# Запуск server.py --stdio в отдельном процессе с подменёнными Google и AD.
STDIO_WRAPPER = textwrap.dedent("""
    import json, sys
    from unittest.mock import MagicMock
    sys.path.insert(0, sys.argv[1])
    rows = json.loads(open(sys.argv[2], encoding="utf-8").read())
    import google_connection, local_identity
    sheets, drive = MagicMock(), MagicMock()
    sheets.spreadsheets.return_value.values.return_value.get.return_value.execute.return_value = {"values": rows}
    sheets.spreadsheets.return_value.create.return_value.execute.return_value = {
        "spreadsheetId": "new-sheet", "sheets": [{"properties": {"sheetId": 0}}]}
    sheets.spreadsheets.return_value.get.return_value.execute.return_value = {"sheets": [{"properties": {"sheetId": 0}}]}
    drive.files.return_value.list.return_value.execute.return_value = {"files": [{"id": "folder-1"}]}
    drive.files.return_value.create.return_value.execute.return_value = {"id": "new-sheet"}
    drive.about.return_value.get.return_value.execute.return_value = {"user": {"emailAddress": "n@gmail.com"}}
    google_connection.load_credentials = lambda *a, **k: MagicMock()
    google_connection.build_services = lambda creds: (sheets, drive)
    local_identity._query_ad = lambda: {"family": "Сидоров", "given": "Сидор", "middle": "Сидорович"}
    sys.argv = ["server.py", "--stdio"]
    import server
    server.main()
""")


class StdioServerTests(unittest.TestCase):
    def test_local_mcp_over_stdio(self):
        with tempfile.TemporaryDirectory() as directory:
            wrapper = Path(directory) / "run.py"
            wrapper.write_text(STDIO_WRAPPER, encoding="utf-8")
            rows = Path(directory) / "rows.json"
            rows.write_text(json.dumps(SHEET, ensure_ascii=False), encoding="utf-8")
            env = dict(os.environ, GOOGLE_AUTH_MODE="local", REQUESTS_SPREADSHEET_ID="sheet-id",
                       WORK_SHEET_NAME="незакр 2025/26", LOCAL_STATE_DIR=directory, USERNAME="sidorov",
                       PYTHONIOENCODING="utf-8", PYTHONUTF8="1")
            params = StdioServerParameters(command=sys.executable, args=[str(wrapper), str(ROOT), str(rows)], env=env)

            async def run():
                async with stdio_client(params) as (read, write):
                    async with ClientSession(read, write) as session:
                        await session.initialize()
                        names = {tool.name for tool in (await session.list_tools()).tools}
                        self.assertTrue({"whoami", "get_work_positions", "search_okpd2", "search_prices",
                                         "create_work_report", "prepare_technical_specification"} <= names)
                        who = (await session.call_tool("whoami", {})).structuredContent
                        self.assertEqual((who["login"], who["fio"], who["google_account"]),
                                         ("sidorov", "Сидоров С.С.", "n@gmail.com"))
                        self.assertEqual(who["folder_url"], "https://drive.google.com/drive/folders/folder-1")
                        positions = (await session.call_tool("get_work_positions", {"employee": who["fio"]}))
                        self.assertFalse(positions.isError, positions.content)
                        data = positions.structuredContent
                        self.assertEqual(data["total"], 2)
                        self.assertIn("незакр 2025/26", data["source_file"])
                        okpd = (await session.call_tool("search_okpd2", {"query": "тягово-сцепное устройство"}))
                        code = okpd.structuredContent["candidates"][0]["code"]
                        report = await session.call_tool("create_work_report", {
                            "employee": who["fio"], "source_version": data["source_version"],
                            "assignments": [{"source_row": p["source_row"], "okpd2": code} for p in data["positions"]],
                            "prices": [{"source_row": 2, "unit_price": 12990, "source_url": "https://shop.example/f"}]})
                        self.assertFalse(report.isError, report.content)
                        self.assertEqual(report.structuredContent["prices_sheet"], "Цены")
                        self.assertEqual(report.structuredContent["row_count"], 2)
            asyncio.run(run())


@unittest.skipUnless(POWERSHELL, "нужен Windows PowerShell")
class InstallerConfigTests(unittest.TestCase):
    """Update-OpenCodeConfig из local-install.ps1 на настоящем PowerShell 5.1."""

    def run_update(self, home: Path, schema: str) -> dict:
        source = (ROOT / "installer" / "local-install.ps1").read_text(encoding="utf-8")
        functions = "\n".join(re.search(rf"(?ms)^function {name}\b.*?^}}", source).group(0)
                              for name in ("Remove-JsonComments", "Get-Child", "Remove-Child", "Update-OpenCodeConfig"))
        script = home / "update.ps1"
        script.write_text("$ErrorActionPreference='Stop'\n$serverName='zakupki'\n"
                          "$utf8=New-Object System.Text.UTF8Encoding($false)\n" + functions + f"""
$server = [pscustomobject]@{{ type='local'; command=@('C:\\venv\\python.exe','C:\\app\\server.py','--stdio');
    environment=[pscustomobject]@{{ GOOGLE_AUTH_MODE='local' }} }}
$provider = [pscustomobject]@{{ name='LetAI-ONE' }}
Update-OpenCodeConfig '{schema}' $server $provider 'letai-one/LetAI-ONE'
""", encoding="utf-8-sig")
        completed = subprocess.run([POWERSHELL, "-NoProfile", "-ExecutionPolicy", "Bypass", "-File", str(script)],
                                   env=dict(os.environ, USERPROFILE=str(home)), capture_output=True, timeout=120)
        self.assertEqual(completed.returncode, 0, completed.stderr.decode("cp866", "replace"))
        config_dir = home / ".config" / "opencode"
        name = "opencode.jsonc" if (config_dir / "opencode.jsonc").exists() else "opencode.json"
        return json.loads((config_dir / name).read_text(encoding="utf-8"))

    def test_v2_replaces_portal_entry_and_keeps_user_settings(self):
        with tempfile.TemporaryDirectory() as directory:
            home = Path(directory)
            (home / ".config" / "opencode").mkdir(parents=True)
            (home / ".config" / "mcp-portal").mkdir(parents=True)
            (home / ".config" / "mcp-portal" / "managed.json").write_text('["google-sheets"]', encoding="utf-8")
            (home / ".config" / "opencode" / "opencode.jsonc").write_text("""{
  // мои настройки
  "model": "delo/model",
  "mcp": { "servers": {
      "google-sheets": { "type": "remote", "url": "http://10.0.0.10:8082/mcp" },
      "delo": { "type": "remote", "url": "http://delo/mcp" }, }, },
}""", encoding="utf-8")
            config = self.run_update(home, "v2")
            servers = config["mcp"]["servers"]
            self.assertNotIn("google-sheets", servers)
            self.assertEqual(servers["zakupki"]["command"], ["C:\\venv\\python.exe", "C:\\app\\server.py", "--stdio"])
            self.assertEqual(servers["delo"]["url"], "http://delo/mcp")
            self.assertEqual(config["providers"]["letai-one"]["name"], "LetAI-ONE")
            self.assertEqual(config["model"], "delo/model")

    def test_v1_fresh_config(self):
        with tempfile.TemporaryDirectory() as directory:
            config = self.run_update(Path(directory), "v1")
            self.assertEqual(config["mcp"]["zakupki"]["type"], "local")
            self.assertEqual(config["provider"]["letai-one"]["name"], "LetAI-ONE")
            self.assertEqual(config["model"], "letai-one/LetAI-ONE")


if __name__ == "__main__":
    unittest.main()
