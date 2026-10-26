# -*- coding: utf-8 -*-
import subprocess


def isDML() -> str:
    try:
        gpus: str = subprocess.check_output(
            [
                'powershell.exe',
                '-NoProfile',
                '-NonInteractive',
                '-Command',
                'Get-CimInstance Win32_VideoController | Select-Object -ExpandProperty Name',
            ],
            encoding='utf-8',
            timeout=15,
            creationflags=subprocess.CREATE_NO_WINDOW,
        )
        dmls = ['gtx', 'rtx', 'radeon', 'arc', 'iris', 'uhd', 'qualcomm']
        for gpu in gpus.split('\n'):
            for kw in dmls:
                if kw in gpu.lower():
                    return gpu

    except Exception:
        pass

    return ''
