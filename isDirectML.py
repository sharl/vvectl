# -*- coding: utf-8 -*-
import subprocess


def isDML() -> bool:
    try:
        gpus: str = subprocess.check_output(
            [
                'pwsh',
                '-NoProfile',
                '-NonInteractive',
                '-Command',
                'Get-CimInstance Win32_VideoController | Select-Object -ExpandProperty Name',
            ],
            encoding='utf-8',
            timeout=5,
            creationflags=subprocess.CREATE_NO_WINDOW,
        )
        dmls = ['gtx', 'rtx', 'radeon', 'arc', 'iris', 'uhd', 'qualcomm']
        for gpu in gpus.split('\n'):
            if any(kw in gpu.lower() for kw in dmls):
                return True

        return False
    except Exception:
        return False
