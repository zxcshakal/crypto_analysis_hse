## \file news_server.py
#  \brief Простой сервер новостей на стандартной библиотеке (http.server).
#
#  Запускается на «домашнем ПК-сервере». Хранит общие новости в Excel
#  (server_news.xlsx) и отдаёт их всем клиентам:
#    * GET  /news  → JSON-список всех новостей;
#    * POST /news  → добавить новость (тело — JSON одной записи).
#
#  Запуск:  python news_server.py 8765
#  Чтобы новости «расходились ко всем», на роутере нужно пробросить этот порт
#  и сообщить клиентам ваш внешний IP (в news_store.SERVER_URL).
#
#  \author Прохачев Никита
#  \date 2026

import os
import sys
import json
import datetime as dt
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer

import pandas as pd

## Файл общей базы новостей на сервере (в папке database/).
if getattr(sys, "frozen", False):
    _ROOT = os.path.dirname(sys.executable)
else:
    _ROOT = os.path.dirname(os.path.abspath(__file__))
_DB_DIR = os.path.join(_ROOT, "database")
os.makedirs(_DB_DIR, exist_ok=True)
SERVER_DB = os.path.join(_DB_DIR, "server_news.xlsx")
## Имя листа.
SHEET = "НОВОСТИ"
## Колонки.
COLUMNS = ["ID", "АВТОР", "ЗАГОЛОВОК", "ТЕКСТ", "ДАТА"]


## \brief Прочитать серверную базу новостей.
#  \return Список словарей-новостей.
def _read() -> list:
    if not os.path.exists(SERVER_DB):
        return []
    df = pd.read_excel(SERVER_DB, sheet_name=SHEET, dtype=str).fillna("")
    return df.to_dict("records")


## \brief Записать список новостей в серверную базу.
#  \param records Список словарей.
#  \return None
def _write(records: list) -> None:
    pd.DataFrame(records, columns=COLUMNS).to_excel(
        SERVER_DB, index=False, sheet_name=SHEET)


## \class NewsHandler
#  \brief Обработчик HTTP-запросов сервера новостей.
class NewsHandler(BaseHTTPRequestHandler):

    ## \brief Отправить JSON-ответ.
    #  \param obj Сериализуемый объект.
    #  \param code HTTP-код.
    #  \return None
    def _send(self, obj, code: int = 200):
        body = json.dumps(obj, ensure_ascii=False).encode("utf-8")
        self.send_response(code)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)

    ## \brief Обработка GET /news — вернуть все новости.
    #  \return None
    def do_GET(self):
        if self.path.rstrip("/") == "/news":
            self._send(_read())
        else:
            self._send({"error": "not found"}, 404)

    ## \brief Обработка POST /news — добавить новость.
    #  \return None
    def do_POST(self):
        if self.path.rstrip("/") != "/news":
            self._send({"error": "not found"}, 404)
            return
        try:
            length = int(self.headers.get("Content-Length", 0))
            record = json.loads(self.rfile.read(length).decode("utf-8"))
        except (ValueError, json.JSONDecodeError):
            self._send({"error": "bad json"}, 400)
            return

        records = _read()
        new_id = (max(int(r["ID"]) for r in records) + 1) if records else 1
        record.setdefault("ID", str(new_id))
        record.setdefault("ДАТА", dt.datetime.now().strftime("%Y-%m-%d %H:%M"))
        records.append({c: str(record.get(c, "")) for c in COLUMNS})
        _write(records)
        self._send({"status": "ok", "id": record["ID"]})

    ## \brief Заглушить стандартный шумный лог сервера.
    def log_message(self, *args):
        pass


## \brief Запустить сервер новостей.
#  \param port Порт (по умолчанию 8765).
#  \return None
def run(port: int = 8765) -> None:
    server = ThreadingHTTPServer(("0.0.0.0", port), NewsHandler)
    print(f"Сервер новостей запущен на порту {port}. Ctrl+C — остановить.")
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        print("\nОстановлен.")
        server.shutdown()


if __name__ == "__main__":
    run(int(sys.argv[1]) if len(sys.argv) > 1 else 8765)
