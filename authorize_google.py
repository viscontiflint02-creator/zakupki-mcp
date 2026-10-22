"""Explicit local OAuth setup; never called from the HTTP service."""
import argparse
from pathlib import Path

from google_auth_oauthlib.flow import InstalledAppFlow
from google_connection import LOCAL_SCOPES, SCOPES, save_token


def main():
    parser = argparse.ArgumentParser(description="Авторизация Google в браузере рабочего ПК")
    parser.add_argument("--credentials", type=Path, default=Path.home() / "mcp/google-sheets/credentials.json")
    parser.add_argument("--token", type=Path, default=Path.home() / "mcp/google-sheets/token.json")
    parser.add_argument("--local", action="store_true",
                        help="MCP на ПК сотрудника: доступ к Таблицам и только к файлам, созданным приложением")
    args = parser.parse_args()
    if not args.credentials.is_file():
        parser.error("Скачайте OAuth client JSON типа Desktop app из Google Cloud и укажите --credentials")
    flow = InstalledAppFlow.from_client_secrets_file(str(args.credentials), LOCAL_SCOPES if args.local else SCOPES)
    credentials = flow.run_local_server(port=0, access_type="offline", prompt="consent",
                                       authorization_prompt_message="Подтвердите доступ Google в открывшемся браузере.")
    if not credentials.refresh_token:
        raise RuntimeError("Google не выдал refresh_token; повторите авторизацию с согласием на offline-доступ")
    save_token(args.token, credentials.to_json())
    print(f"Google OAuth сохранен: {args.token}.")


if __name__ == "__main__":
    main()
