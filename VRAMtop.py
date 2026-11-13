# -*- coding: utf-8 -*-
from shutil import get_terminal_size
import time

from getVRAMInfo import get_vram_info, get_vram_usages

INTERVAL = 3

lines = get_terminal_size()[1]
curoff = '\033[?25l'
curon = '\033[?25h'
fill = '\n'.join([' ' * 24 for _ in range(lines)])

cls = '\033[H\033[J'
home = f'{curoff}{fill}\n\033[0;0H'

vram_info = get_vram_info()[0]
gpu_name, vram_size = vram_info.values()
top_line = f'{home}{gpu_name:24}\t{vram_size // 1024 // 1024:>8.1f} MB'

# targets = ['Dedicated', 'Local']
targets = ['Local']
print(cls, end='')

while True:
    lines = get_terminal_size()[1]
    for target in targets:
        try:
            begin = time.perf_counter()

            ps = []
            usages = get_vram_usages(target=target)

            used = 0
            for name in sorted(usages, key=lambda n: usages[n], reverse=True)[:lines - 2]:
                mem = usages[name]
                used += mem
                ps.append(f'{name:24}\t{mem / 1024 / 1024:8.1f} MB {100 * mem / vram_size:>4.1f}%')

            elapsed = time.perf_counter() - begin
            print(
                f'{top_line} {100 * used / vram_size:>4.1f}% {target} {elapsed:.3f}s\n\n{'\n'.join(ps)}{curon}',
                flush=True,
                end='',
            )
            sleep_time = max(0, INTERVAL - elapsed)
            time.sleep(sleep_time)
        except Exception:
            print()
            break
