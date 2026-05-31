# -*- coding: utf-8 -*-
from dataclasses import asdict, dataclass
import ctypes
import json
import os
import re
import socket
import subprocess
import threading
import time

from PIL import Image, ImageDraw
from pystray import Icon, Menu, MenuItem
import darkdetect as dd
import psutil

from config import Config

TITLE = 'VVEctl'

LISTEN_PORT = 50021
APP_INTERNAL_PORT = 50022
SUBMENU_LEN = 8         # greater than 8
IDLE_LIMIT = 3600
BASE_DIR = os.path.join(os.environ.get('LOCALAPPDATA'), r'Programs\VOICEVOX\vv-engine')
PROC_NAME = 'run.exe'
EXE_PATH = os.path.join(BASE_DIR, PROC_NAME)
APP_CMD = [EXE_PATH, '--host=127.0.0.1', f'--port={APP_INTERNAL_PORT}', '--use_gpu']

# https://www.jma.go.jp/jma/kishou/info/colorguide/HPColorGuide_202007.pdf
COLORS = {
    90: (180, 0, 104),
    80: (165, 0, 33),
    70: (255, 40, 0),
    60: (255, 133, 0),
    50: (254, 230, 0),
    40: (250, 230, 150),
    30: (0, 65, 255),
    20: (0, 170, 255),
    10: (242, 242, 255),
}
PreferredAppMode = {
    'Light': 0,
    'Dark': 1,
}
# https://github.com/moses-palmer/pystray/issues/130
ctypes.windll['uxtheme.dll'][135](PreferredAppMode[dd.theme()])


# 保存する設定の型定義
@dataclass
class Setting:
    vram_limit_mb: int


class menu_mib:
    def __init__(self, vram_size_gb):
        self._unit = ' MB'

        defaults = [512, 1024, 1536, 2048]
        step = 1024 * vram_size_gb // 2 // SUBMENU_LEN
        self.range = range(step, 1024 * vram_size_gb // 2 + 1, step)
        b = sorted(list(set([m for m in self.range] + defaults)))
        self.list = [f'{m}{self._unit}' for m in b if self.to_mib(m) <= (1024 * vram_size_gb // 2)]

    def to_mib(self, item):
        return int(str(item).removesuffix(self._unit))


class TaskTray:
    def __init__(self):
        self.stop_event = threading.Event()
        self.config = Config(TITLE)

        # 最後に proxy にアクセスした時刻
        self.last_access_time = time.time()
        self.current_vram = 0.0
        self.enable_idle = False

        image = self.create_icon_image(0)

        # GPU / VRAM 設定
        self.gpuname = ''
        self.vram_gb = 0
        vrams = self.get_vram_info_via_pwsh()
        if vrams:
            # use first GPU detected
            self.gpuname = vrams[0].get('name', '')
            # gibibyte
            self.vram_gb = int(vrams[0].get('vram', 0) // (1024 * 1024 * 1024))

            # DEBUG
            # self.gpuname = 'NVIDIA RTX PRO 6000 Blackwell'
            # self.vram_gb = 96
            # DEBUG

        # 搭載メモリ量に応じたサブメニューを設定
        vram_limit_submenu = [
            MenuItem(f'{self.gpuname} {self.vram_gb} GB', lambda: False),
            Menu.SEPARATOR,
        ]
        self.mibs = menu_mib(self.vram_gb)
        for i in self.mibs.list:
            vram_limit_submenu.append(
                MenuItem(i, self.set_vram_limit, checked=lambda x: self.vram_limit_checked(x)),
            )

        # VRAM 上限の設定
        self.vram_limit_mb = self.mibs.to_mib(self.mibs.list[1])
        # 設定から VRAM 上限を読み込み
        self.load_config()

        main_menu = Menu(
            MenuItem('VOICEVOX Engine control', lambda: False),
            Menu.SEPARATOR,
            MenuItem('Manual Restart', lambda: self.restart_logic('Manual Request')),
            MenuItem('Enable Idle Timeout', self.toggle_idle, checked=lambda _: self.enable_idle),
            MenuItem('VRAM limit', Menu(*vram_limit_submenu)),
            Menu.SEPARATOR,
            MenuItem('Exit', self.stopApp),
        )
        self.app = Icon(name=f'PYTHON.win32.{TITLE}', title='Starting...', icon=image, menu=main_menu)

    def load_config(self):
        try:
            setting = Setting(**self.config.load())
            self.vram_limit_mb = setting.vram_limit_mb
        except TypeError:
            pass

    def save_config(self):
        setting = Setting(vram_limit_mb=self.vram_limit_mb)
        self.config.save(asdict(setting))

    def create_icon_image(self, perc, SIZE=64):
        image = Image.new('RGB', (SIZE, SIZE), color=(73, 109, 137))
        d = ImageDraw.Draw(image)
        if perc > 0:
            for c in COLORS:
                if perc >= c:
                    d.rectangle((0, SIZE - int(SIZE * perc / 100), SIZE, SIZE), fill=COLORS[c])
                    break
        d.text((10, 10), TITLE, fill=(255, 255, 255))
        return image

    def set_vram_limit(self, _, item):
        self.vram_limit_mb = self.mibs.to_mib(item)
        print(f'[{time.strftime('%H:%M:%S')}] set limit to {self.vram_limit_mb}')
        self.save_config()

    def vram_limit_checked(self, item):
        return self.mibs.to_mib(item) == self.vram_limit_mb

    def get_vram_info_via_pwsh(self):
        """pwsh 7 を使用してVRAM容量取得"""
        ps_cmd = (
            "Get-ChildItem 'HKLM:\\SYSTEM\\CurrentControlSet\\Control\\Class\\{4d36e968-e325-11ce-bfc1-08002be10318}' -ErrorAction SilentlyContinue | "
            "Where-Object { $_.GetValue('HardwareInformation.qwMemorySize') } | "
            "ForEach-Object { [PSCustomObject]@{ name = $_.GetValue('HardwareInformation.AdapterString'); vram = $_.GetValue('HardwareInformation.qwMemorySize'); } } |  "
            "ConvertTo-Json; exit 0"
        )

        try:
            stdout = subprocess.check_output(
                ["pwsh", "-NoProfile", "-NonInteractive", "-Command", ps_cmd],
                stderr=subprocess.DEVNULL, encoding='utf-8', errors='ignore',
                creationflags=subprocess.CREATE_NO_WINDOW,
            )
        except (subprocess.CalledProcessError, FileNotFoundError):
            return []

        if not stdout.strip():
            return []

        # 単一GPUの場合も list で返す
        try:
            data = json.loads(stdout)
            if isinstance(data, dict):
                return [data]
            return data
        except json.JSONDecodeError:
            return []

    def get_vv_vram_via_pwsh(self):
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

    def toggle_idle(self, _, __):
        self.enable_idle = not self.enable_idle

    def restart_logic(self, reason):
        print(f'[{time.strftime('%H:%M:%S')}] {reason}')
        subprocess.run(['taskkill', '/F', '/IM', PROC_NAME, '/T'],
                       creationflags=subprocess.CREATE_NO_WINDOW, capture_output=True)
        time.sleep(5)
        subprocess.Popen(APP_CMD, cwd=BASE_DIR, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL,
                         creationflags=subprocess.CREATE_NO_WINDOW)

    def monitor_loop(self):
        while not self.stop_event.is_set():
            if self.stop_event.wait(15):
                break
            idle_time = time.time() - self.last_access_time
            self.current_vram = self.get_vv_vram_via_pwsh()

            # update tooltip
            perc = 100 * self.current_vram / self.vram_limit_mb
            if self.enable_idle:
                self.app.title = f'VRAM: {self.current_vram:.1f} MB / {perc:.1f} % / Idle: {int(idle_time)}s'
            else:
                self.app.title = f'VRAM: {self.current_vram:.1f} MB / {perc:.1f} %'
            # update icon
            self.app.icon = self.create_icon_image(perc)

            if self.enable_idle and idle_time > IDLE_LIMIT:
                self.restart_logic('Idle Timeout')
                self.last_access_time = time.time()
            elif self.current_vram > self.vram_limit_mb and idle_time > 30:
                self.restart_logic(f'VRAM Leak ({self.current_vram:.1f} MB)')
                self.last_access_time = time.time()

    def bridge(self, src, dst):
        try:
            while not self.stop_event.is_set():
                data = src.recv(8192)
                if not data:
                    break
                dst.sendall(data)
        except Exception:
            pass
        finally:
            src.close()
            dst.close()

    def proxy_handler(self):
        server = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
        server.setsockopt(socket.SOL_SOCKET, socket.SO_REUSEADDR, 1)
        server.bind(('', LISTEN_PORT))
        server.listen(100)
        while not self.stop_event.is_set():
            client_sock, _ = server.accept()
            self.last_access_time = time.time()
            try:
                app_sock = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
                app_sock.connect(('127.0.0.1', APP_INTERNAL_PORT))
                threading.Thread(target=self.bridge, args=(client_sock, app_sock), daemon=True).start()
                threading.Thread(target=self.bridge, args=(app_sock, client_sock), daemon=True).start()
            except Exception:
                client_sock.close()

    def stopApp(self):
        self.stop_event.set()
        self.app.stop()

    def runApp(self):
        self.stop_event.clear()

        threading.Thread(target=self.proxy_handler, daemon=True).start()
        threading.Thread(target=self.monitor_loop, daemon=True).start()
        self.restart_logic('Initial Start')

        self.app.run()


if __name__ == '__main__':
    TaskTray().runApp()
