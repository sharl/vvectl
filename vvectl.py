# -*- coding: utf-8 -*-
import ctypes
import os
import re
import socket
import subprocess
import threading
import time

from PIL import Image, ImageDraw
import darkdetect as dd
import psutil
import pystray

LISTEN_PORT = 50021
APP_INTERNAL_PORT = 50022
VRAM_LIMIT_MB = 1500
IDLE_LIMIT = 3600
BASE_DIR = os.path.join(os.environ.get('LOCALAPPDATA'), r'Programs\VOICEVOX\vv-engine')
PROC_NAME = 'run.exe'
EXE_PATH = os.path.join(BASE_DIR, PROC_NAME)
APP_CMD = [EXE_PATH, '--host=127.0.0.1', f'--port={APP_INTERNAL_PORT}', '--use_gpu']

# 共有変数
last_access_time = time.time()
current_vram = 0.0
icon = None
PreferredAppMode = {
    'Light': 0,
    'Dark': 1,
}
# https://github.com/moses-palmer/pystray/issues/130
ctypes.windll['uxtheme.dll'][135](PreferredAppMode[dd.theme()])


def create_icon_image(perc, SIZE=64):
    """タスクバー用の簡易アイコン（Vの文字など）を作成"""
    image = Image.new('RGB', (SIZE, SIZE), color=(73, 109, 137))
    d = ImageDraw.Draw(image)
    d.rectangle((0, SIZE - int(SIZE * perc / 100), SIZE, SIZE), fill=(255, 0, 0))
    d.text((10, 10), 'VVE', fill=(255, 255, 255))
    return image


def get_vv_vram_via_pwsh():
    """pwsh 7 を使用してVRAM取得"""
    total_mib = 0
    try:
        vv_pids = [p.info['pid'] for p in psutil.process_iter(['name', 'pid'])
                   if p.info['name'] and p.info['name'].lower() == PROC_NAME.lower()]
        if not vv_pids:
            return 0.0

        ps_cmd = 'Get-Counter "\\GPU Process Memory(*)\\Dedicated Usage" | Select-Object -ExpandProperty CounterSamples | ForEach-Object { "$($_.Path) : $($_.CookedValue)" }'
        result = subprocess.check_output(
            ['pwsh', '-NoProfile', '-NonInteractive', '-Command', ps_cmd],
            stderr=subprocess.DEVNULL, encoding='utf-8', errors='ignore',
            creationflags=subprocess.CREATE_NO_WINDOW
        )
        pattern = re.compile(r'pid_(\d+).*?:\s+(\d+)')
        for pid_str, usage_str in pattern.findall(result):
            if int(pid_str) in vv_pids:
                total_mib += int(usage_str)
    except Exception:
        pass
    return total_mib / 1024 / 1024


def restart_logic(reason):
    print(f'[{time.strftime('%H:%M:%S')}] {reason}')
    subprocess.run(['taskkill', '/F', '/IM', PROC_NAME, '/T'],
                   creationflags=subprocess.CREATE_NO_WINDOW, capture_output=True)
    time.sleep(5)
    subprocess.Popen(APP_CMD, cwd=BASE_DIR, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL,
                     creationflags=subprocess.CREATE_NO_WINDOW)


def monitor_loop():
    global last_access_time, current_vram, icon
    while True:
        time.sleep(15)
        idle_time = time.time() - last_access_time
        current_vram = get_vv_vram_via_pwsh()

        # update tooltip, icon
        if icon:
            perc = 100 * current_vram / VRAM_LIMIT_MB
            icon.title = f'VRAM: {current_vram:.1f} MB / {perc:.1f} % / Idle: {int(idle_time)}s'
            icon.icon = create_icon_image(perc)

        if idle_time > IDLE_LIMIT:
            restart_logic('Idle Timeout')
            last_access_time = time.time()
        elif current_vram > VRAM_LIMIT_MB and idle_time > 30:
            restart_logic(f'VRAM Leak ({current_vram:.1f} MB)')
            last_access_time = time.time()


def bridge(src, dst):
    try:
        while True:
            data = src.recv(8192)
            if not data:
                break
            dst.sendall(data)
    except Exception:
        pass
    finally:
        src.close()
        dst.close()


def proxy_handler():
    global last_access_time
    server = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
    server.setsockopt(socket.SOL_SOCKET, socket.SO_REUSEADDR, 1)
    server.bind(('', LISTEN_PORT))
    server.listen(100)
    while True:
        client_sock, _ = server.accept()
        last_access_time = time.time()
        try:
            app_sock = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
            app_sock.connect(('127.0.0.1', APP_INTERNAL_PORT))
            threading.Thread(target=bridge, args=(client_sock, app_sock), daemon=True).start()
            threading.Thread(target=bridge, args=(app_sock, client_sock), daemon=True).start()
        except Exception:
            client_sock.close()


if __name__ == '__main__':
    threading.Thread(target=proxy_handler, daemon=True).start()
    threading.Thread(target=monitor_loop, daemon=True).start()
    restart_logic('Initial Start')

    icon = pystray.Icon('VVEctl', create_icon_image(0), 'Starting...')
    icon.menu = pystray.Menu(
        pystray.MenuItem('Manual Restart', lambda: restart_logic('Manual Request')),
        pystray.MenuItem('Exit', lambda: icon.stop())
    )
    icon.run()
