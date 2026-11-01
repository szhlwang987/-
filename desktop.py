"""Windows desktop launcher for the PCBA planning application."""
import multiprocessing
import threading
import time
import socket
import uvicorn
import webview
from app.main import app

APP_TITLE = '物料需求需采购计划系统'

def free_port():
    with socket.socket() as s:
        s.bind(('127.0.0.1', 0))
        return s.getsockname()[1]

def run_server(port):
    uvicorn.run(app, host='127.0.0.1', port=port, log_level='warning')

def main():
    port = free_port()
    thread = threading.Thread(target=run_server, args=(port,), daemon=True)
    thread.start()
    for _ in range(50):
        try:
            with socket.create_connection(('127.0.0.1', port), timeout=.2): break
        except OSError: time.sleep(.1)
    webview.create_window(APP_TITLE, f'http://127.0.0.1:{port}', width=1440, height=900, min_size=(1100, 700))
    webview.start()

if __name__ == '__main__':
    multiprocessing.freeze_support()
    main()
