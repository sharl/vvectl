# -*- coding: utf-8 -*-
from dataclasses import asdict, dataclass
import ctypes
import logging
import logging.handlers
import os
import socket
import subprocess
import sys
import threading
import time

from PIL import Image, ImageDraw
from pystray import Icon, Menu, MenuItem
import darkdetect as dd
import requests

from config import Config
from getLog import getLog
from isDirectML import isDML
from utils import resource_path
from getVRAMInfo import get_vram_info, get_vram_usages

TITLE = 'VVEctl'

BIND_ADDRESS = '127.0.0.1'
LISTEN_PORT = 50021
APP_INTERNAL_PORT = 50022
SUBMENU_LEN = 8         # greater than 8
IDLE_LIMIT = 3600
SWEEP_INTERVAL = 1800
SWEEP_THREAD_NAME = 'sweep_loop'
SWEEP_NEVER_LABEL = 'Never sweep'
BASE_DIR = os.path.join(os.environ.get('LOCALAPPDATA'), r'Programs\VOICEVOX\vv-engine')
if not os.path.exists(BASE_DIR):
    # search alternate path for 0.26.0~
    appdata = os.environ.get('APPDATA')
    path = os.path.join(appdata, r'voicevox\vvpp-engines')
    if os.path.exists(path):
        for target_dir in os.listdir(path):
            if target_dir.startswith('VOICEVOX_Engine+'):
                BASE_DIR = os.path.join(path, target_dir)
                break
PROC_NAME = 'run.exe'
EXE_PATH = os.path.join(BASE_DIR, PROC_NAME)
APP_CMD = [EXE_PATH, f'--host={BIND_ADDRESS}', f'--port={APP_INTERNAL_PORT}', '--use_gpu']

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
    opv_threshold_mb: int
    enable_idle: bool
    theme: str
    sweep_interval: int


THEMES = {
    'Default': {
        'bg': (73, 109, 137),
        'fg': (255, 255, 255),
    },
    'black / white': {
        'bg': (0, 0, 0),
        'fg': (255, 255, 255),
    },
    'VOICEVOX': {
        'bg': (165, 212, 173),
        'fg': (0, 0, 0),
    },
}

# logger settings
logname = getLog(TITLE, 'log.log')
logging.basicConfig(
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[
        logging.handlers.RotatingFileHandler(logname, encoding='utf-8', maxBytes=1000000, backupCount=0),
        logging.StreamHandler(),
    ],
    datefmt='%Y/%m/%d %X'
)
logger = logging.getLogger(TITLE)
logger.setLevel(logging.DEBUG)


# 最小限の win32con
class win32con:
    MB_OK = 0x00000000
    MB_ICONERROR = 0x00000010
    MB_TOPMOST = 0x00040000


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


def getVersion():
    v = 'test'
    try:
        with open(resource_path('Assets/version.txt')) as fd:
            v = fd.read().strip().removeprefix('v')
    except Exception:
        pass
    return v


class TaskTray:
    def __init__(self):
        gpu_name = isDML()
        if not gpu_name:
            text = 'DirectML 対応の GPU が見つかりません\n\nアプリケーションを終了します'
            style = win32con.MB_OK | win32con.MB_ICONERROR | win32con.MB_TOPMOST
            ctypes.windll.user32.MessageBoxW(None, text, TITLE, style)
            sys.exit()

        self.stop_event = threading.Event()
        self.config = Config(TITLE)
        self.theme = list(THEMES)[0]            # Default

        # 最後に proxy にアクセスした時刻
        self.last_access_time = time.time()
        self.current_vram = 0.0
        self.enable_idle = False
        self.sweep_stop_event = threading.Event()
        self.sweep_interval = SWEEP_INTERVAL

        # 他プロセスが使用しているプロセス名と VRAM 使用量
        self.top_process_name = None
        self.top_vram = 0.0

        image = self.create_icon_image(0)

        # GPU / VRAM 設定
        self.gpuname = gpu_name
        self.vram_gb = 0
        self.vram_mb = 0
        vrams = get_vram_info()
        if vrams:
            # use first GPU detected
            self.gpuname = vrams[0].get('name', '')
            # gibibyte
            self.vram_gb = int(vrams[0].get('vram', 0) // (1024 * 1024 * 1024))
            # mebibyte, insurance
            self.vram_mb = (self.vram_gb - 1) * 1024

            # DEBUG
            # self.gpuname = 'NVIDIA RTX PRO 6000 Blackwell'
            # self.vram_gb = 96
            # self.vram_mb = (self.vram_gb - 1) * 1024
            # DEBUG
        # VRAM overflow limit
        self.limit = self.vram_mb

        # 搭載メモリ量に応じたサブメニューを設定
        # (1) VVE VRAM limit
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

        # (2) Other Process VRAM threshold
        opv_threshold_submenu = []
        self.opv_mibs = menu_mib(self.vram_gb)
        for i in self.mibs.list:
            opv_threshold_submenu.append(
                MenuItem(i, self.set_opv_threshold, checked=lambda x: self.opv_threshold_checked(x)),
            )
        # Other Process VRAM threshold の設定
        self.opv_threshold_mb = self.opv_mibs.to_mib(self.opv_mibs.list[-1])

        # icon theme submenu
        theme_submenu = []
        for theme in THEMES:
            theme_submenu.append(
                MenuItem(theme, self.set_theme, checked=lambda x: str(x) == self.theme),
            )

        # Sweep interval の設定
        sweep_interval_submenu = [
            MenuItem(SWEEP_NEVER_LABEL, self.set_sweep_interval, checked=lambda x: self.sweep_interval_checked(x)),
        ]
        for t in range(SWEEP_INTERVAL, SWEEP_INTERVAL * 6 + 1, SWEEP_INTERVAL):
            sweep_interval_submenu.append(
                MenuItem(f'{t}s', self.set_sweep_interval, checked=lambda x: self.sweep_interval_checked(x)),
            )

        self.load_config()

        main_menu = Menu(
            MenuItem(f'VOICEVOX Engine control {getVersion()}', lambda: False),
            MenuItem('Theme', Menu(*theme_submenu)),
            Menu.SEPARATOR,
            MenuItem('Manual Restart', lambda: self.restart_logic('Manual Restart')),
            MenuItem('Enable Idle Timeout', self.toggle_idle, checked=lambda _: self.enable_idle),
            Menu.SEPARATOR,
            MenuItem('Sweep interval', Menu(*sweep_interval_submenu)),
            Menu.SEPARATOR,
            MenuItem('Other Process VRAM threshold', Menu(*opv_threshold_submenu)),
            MenuItem('VOCEVOX Engine VRAM limit', Menu(*vram_limit_submenu)),
            Menu.SEPARATOR,
            MenuItem('Exit', self.stopApp),
        )
        self.app = Icon(name=f'PYTHON.win32.{TITLE}', title='Starting...', icon=image, menu=main_menu)

    def load_config(self):
        try:
            setting = Setting(**self.config.load())
            self.vram_limit_mb = setting.vram_limit_mb
            self.opv_threshold_mb = setting.opv_threshold_mb
            self.enable_idle = setting.enable_idle
            self.theme = setting.theme
            self.sweep_interval = setting.sweep_interval
        except TypeError:
            pass

    def save_config(self):
        setting = Setting(
            vram_limit_mb=self.vram_limit_mb,
            opv_threshold_mb=self.opv_threshold_mb,
            enable_idle=self.enable_idle,
            theme=self.theme,
            sweep_interval=self.sweep_interval,
        )
        self.config.save(asdict(setting))

    def set_theme(self, _, item):
        self.theme = str(item)
        self.save_config()

    def create_icon_image(self, perc, SIZE=64):
        theme = THEMES.get(self.theme)
        if theme is None:
            theme = list(THEMES)[0]             # Default

        image = Image.new('RGB', (SIZE, SIZE), color=theme['bg'])
        d = ImageDraw.Draw(image)
        if perc > 0:
            for c in COLORS:
                if perc >= c:
                    d.rectangle((0, SIZE - int(SIZE * perc / 100), SIZE, SIZE), fill=COLORS[c])
                    break
        d.text((10, 10), TITLE, fill=theme['fg'])
        return image

    def set_vram_limit(self, _, item):
        self.vram_limit_mb = self.mibs.to_mib(item)
        logger.info(f'set limit to {self.vram_limit_mb}')
        self.save_config()

    def vram_limit_checked(self, item):
        return self.mibs.to_mib(item) == self.vram_limit_mb

    def set_opv_threshold(self, _, item):
        self.opv_threshold_mb = self.opv_mibs.to_mib(item)
        logger.info(f'set other process threshold to {self.opv_threshold_mb}')
        self.save_config()

    def opv_threshold_checked(self, item):
        return self.opv_mibs.to_mib(item) == self.opv_threshold_mb

    def _detect_sweep_thread(self):
        threads = [th.name for th in threading.enumerate()]
        logger.debug(f'running threads: {threads}')
        return SWEEP_THREAD_NAME in threads

    def _start_sweep_thread(self):
        self.sweep_stop_event.clear()
        threading.Thread(target=self.sweep_loop, name=SWEEP_THREAD_NAME, daemon=True).start()

    def _sweep(self):
        server_url = f'http://{BIND_ADDRESS}:{APP_INTERNAL_PORT}'
        session = requests.Session()

        try:
            speakers = session.get(f'{server_url}/speakers', timeout=5).json()
            for speaker in speakers:
                for style in speaker['styles']:
                    sid = style['id']
                    name = f"{speaker['name']}({style['name']})"

                    initialized = session.get(f'{server_url}/is_initialized_speaker?speaker={sid}').json()
                    if initialized:
                        # すでにロード済みなら初期化する
                        session.post(
                            f'{server_url}/initialize_speaker',
                            params={
                                'speaker': sid,
                                'skip_reinit': 'false'
                            },
                            timeout=5,
                        )
                        logger.info(f'initialized {name}')
        except Exception:
            pass

    def sweep_loop(self):
        logger.info(f'start {SWEEP_THREAD_NAME}')

        while not self.sweep_stop_event.is_set():
            begin = time.perf_counter()

            if self.sweep_interval:
                self._sweep()

            elapsed = time.perf_counter() - begin
            sleep_time = max(1, self.sweep_interval - elapsed)
            if self.sweep_stop_event.wait(sleep_time):
                break

        logger.info(f'stopped {SWEEP_THREAD_NAME}')

    def set_sweep_interval(self, _, _item):
        item = str(_item)
        if item == SWEEP_NEVER_LABEL:
            interval = 0
        else:
            interval = int(item.removesuffix('s'))

        if self.sweep_interval != interval:
            self.sweep_interval = interval
            self.save_config()

            logger.info(f'set sweep interval to {self.sweep_interval}')

            # stop sweep
            self.sweep_stop_event.set()
            # wait
            while self._detect_sweep_thread():
                self.stop_event.wait(0.1)
            # start sweep
            self._start_sweep_thread()

    def sweep_interval_checked(self, _item):
        item = str(_item)
        if item == SWEEP_NEVER_LABEL:
            interval = 0
        else:
            interval = int(item.removesuffix('s'))
        return self.sweep_interval == interval

    def get_vram_usage(self) -> list[int, str | None, int]:
        # VVE current usage
        current_vram = 0
        # top process
        top_process_name = None
        top_vram = 0

        try:
            usages = get_vram_usages()
            # find VVE and most usage process
            for name in sorted(usages, key=lambda n: usages[n]):
                # desktop window manager と explorer は除外
                if name.lower() in ['dwm.exe', 'explorer.exe']:
                    continue

                mem = usages[name]
                if mem != 0:
                    if name.lower() == PROC_NAME.lower():
                        current_vram = mem
                    else:
                        if mem >= self.opv_threshold_mb * 1024 * 1024:
                            ext = '.' + name.rsplit('.', 1)[-1]
                            top_process_name = name.removesuffix(ext)
                            top_vram = mem
        except Exception:
            pass
        return current_vram / 1024 / 1024, top_process_name, top_vram / 1024 / 1024

    def toggle_idle(self, _, __):
        self.enable_idle = not self.enable_idle
        self.save_config()

    def restart_logic(self, reason=None):
        if reason:
            logger.info(reason)
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
            # VVE current usage, top process name, top process usage
            self.current_vram, top_process_name, top_vram = self.get_vram_usage()

            if top_process_name:
                if self.top_process_name != top_process_name:
                    logger.info(f'found {top_process_name} {top_vram:.1f} / {self.opv_threshold_mb} MB: set limit to {self.vram_limit_mb}')
                    self.limit = self.vram_limit_mb
                    self.top_vram = top_vram
            else:
                if self.top_process_name != top_process_name:
                    logger.info('unset limit')
                    self.limit = self.vram_mb
                    self.top_vram = 0.0

            self.top_process_name = top_process_name

            # print(f'current {self.current_vram:.1f} / {self.limit} top {self.top_process_name} {self.top_vram:.1f} / {self.opv_threshold_mb}')

            # update tooltip
            perc = 100 * self.current_vram / self.limit
            topic = f'Limit: {self.limit} MB\n' if self.limit != self.vram_mb else ''
            if self.enable_idle:
                self.app.title = f'{topic}VRAM: {self.current_vram:.1f} MB / {perc:.1f} % / Idle: {int(idle_time)}s'
            else:
                self.app.title = f'{topic}VRAM: {self.current_vram:.1f} MB / {perc:.1f} %'
            # update icon
            self.app.icon = self.create_icon_image(perc)

            if self.enable_idle and idle_time > IDLE_LIMIT:
                self.restart_logic('Idle Timeout')
                self.last_access_time = time.time()
            elif self.current_vram > self.limit and idle_time > 30:
                self.restart_logic(f'VRAM overflow ({self.current_vram:.1f} / {self.limit} MB)')
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
                app_sock.connect((BIND_ADDRESS, APP_INTERNAL_PORT))
                threading.Thread(target=self.bridge, args=(client_sock, app_sock), daemon=True).start()
                threading.Thread(target=self.bridge, args=(app_sock, client_sock), daemon=True).start()
            except Exception:
                client_sock.close()

    def stopApp(self):
        self.sweep_stop_event.set()
        self.stop_event.set()
        self.app.stop()

    def runApp(self):
        self.stop_event.clear()

        threading.Thread(target=self.proxy_handler, daemon=True).start()
        threading.Thread(target=self.monitor_loop, daemon=True).start()
        self._start_sweep_thread()
        self.restart_logic()

        self.app.run()


if __name__ == '__main__':
    TaskTray().runApp()
