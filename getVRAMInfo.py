# -*- coding: utf-8 -*-
import json
import re
import subprocess

import psutil


def get_vram_info():
    _cmd = [
        'Get-ChildItem "HKLM:\\SYSTEM\\CurrentControlSet\\Control\\Class\\{4d36e968-e325-11ce-bfc1-08002be10318}" -ErrorAction SilentlyContinue',
        'Where-Object { $_.GetValue("HardwareInformation.qwMemorySize") }',
        'ForEach-Object { [PSCustomObject]@{ name = $_.GetValue("HardwareInformation.AdapterString"); vram = $_.GetValue("HardwareInformation.qwMemorySize"); } }',
        'ConvertTo-Json; exit 0',
    ]
    ps_cmd = ' | '.join(_cmd)

    try:
        stdout = subprocess.check_output(
            ['powershell.exe', '-NoProfile', '-NonInteractive', '-Command', ps_cmd],
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


def get_vram_usages(target='Local'):
    pids = {}
    for p in psutil.process_iter(['name', 'pid']):
        pid = p.info['pid']
        name = p.info['name']
        pids[pid] = name

    _cmd = [
        f'Get-Counter "\\GPU Process Memory(*)\\{target} Usage"',
        'Select-Object -ExpandProperty CounterSamples',
        'ForEach-Object { "$($_.Path) : $($_.CookedValue)" }',
    ]
    ps_cmd = ' | '.join(_cmd)
    result = subprocess.check_output(
        ['powershell', '-NoProfile', '-NonInteractive', '-Command', ps_cmd],
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

    return usages
