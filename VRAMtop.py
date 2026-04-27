# -*- coding: utf-8 -*-
import re
import subprocess

import psutil


def getVRAMUsage():
    pids = {}
    for p in psutil.process_iter(['name', 'pid']):
        pid = p.info['pid']
        name = p.info['name']
        # desktop window manager は常に VRAM 上限を返すので除外
        if name and name.lower() != 'dwm.exe':
            pids[pid] = name

    ps_cmd = 'Get-Counter "\\GPU Process Memory(*)\\Dedicated Usage" | Select-Object -ExpandProperty CounterSamples | ForEach-Object { "$($_.Path) : $($_.CookedValue)" }'
    result = subprocess.check_output(
        ['pwsh', '-NoProfile', '-NonInteractive', '-Command', ps_cmd],
        stderr=subprocess.DEVNULL, encoding='utf-8', errors='ignore',
        creationflags=subprocess.CREATE_NO_WINDOW
    )
    pattern = re.compile(r'pid_(\d+).*?:\s+(\d+)')
    usages = {}
    for pid_str, usage_str in pattern.findall(result):
        pid = int(pid_str)
        if pid in pids:
            name = pids[pid]
            if name not in usages:
                usages[name] = 0
            usages[name] += int(usage_str)

    for name in sorted(usages, key=lambda n: usages[n], reverse=True):
        mem = usages[name]
        if mem != 0:
            print(f'{name:24}\t{mem / 1024 / 1024:8.1f} MB')


if __name__ == '__main__':
    getVRAMUsage()
