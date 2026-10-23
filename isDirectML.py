# -*- coding: utf-8 -*-
from pathlib import Path
import subprocess
import sys
import traceback


def isDML() -> bool:
    log = Path(sys.executable).resolve().with_name('isDML.log')

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
            timeout=15,
            creationflags=subprocess.CREATE_NO_WINDOW,
            stderr=subprocess.STDOUT,
        )

        log.write_text(
            f'GPU names:\n{gpus!r}\n',
            encoding='utf-8',
        )

        dmls = ['gtx', 'rtx', 'radeon', 'arc', 'iris', 'uhd', 'qualcomm']
        return any(
            kw in gpu.lower()
            for gpu in gpus.splitlines()
            for kw in dmls
        )

    except Exception:
        log.write_text(
            traceback.format_exc(),
            encoding='utf-8',
        )
        return False
