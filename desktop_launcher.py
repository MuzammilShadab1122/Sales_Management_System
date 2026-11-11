import os
import shutil
import subprocess
import threading
import time
import webbrowser
from http.server import ThreadingHTTPServer

from app import RajputTradersRequestHandler
from database import init_db
from excel_sync import sync_all_excel_files

HOST = "0.0.0.0"
PORT = 8765
APP_URL = f"http://127.0.0.1:{PORT}"


class ReusableThreadingHTTPServer(ThreadingHTTPServer):
    allow_reuse_address = True


def launch_desktop_window():
    time.sleep(0.8)
    edge_paths = [
        shutil.which("msedge"),
        r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe",
        r"C:\Program Files\Microsoft\Edge\Application\msedge.exe",
    ]
    for path in edge_paths:
        if path and os.path.exists(path):
            subprocess.Popen([path, f"--app={APP_URL}", "--start-maximized"])
            return

    webbrowser.open(APP_URL)


if __name__ == "__main__":
    print("==========================================================")
    print("  RAJPOOT TRADERS — ENTERPRISE SALES & STOCK ERP v2.5     ")
    print("==========================================================")
    init_db()
    sync_result = sync_all_excel_files()
    print(f"[OK] Real-Time Excel Sheets synced at: {sync_result['excel_dir']}")
    print(f"[OK] Starting Desktop App Server on {APP_URL} ...")

    threading.Thread(target=launch_desktop_window, daemon=True).start()
    server = ReusableThreadingHTTPServer((HOST, PORT), RajputTradersRequestHandler)
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        print("\nShutting down Rajpoot Traders System...")
        server.server_close()
