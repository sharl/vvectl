# vvectl

Reduce VRAM consumption of VOICEVOX Engine with GPU

## Prerequisite

PowerShell 7
https://aka.ms/install-powershell

## Run

```
git clone https://github.com/sharl/vvectl.git
cd vvectl
python -m venv .venv
.\.venv\Scripts\activate
pip install -r requirements.txt
python vvectl.py
```

## Build

```powershell
pip install pyinstaller
pyinstaller vvectl.py --onefile --noconsole --icon Assets/sample.ico --add-data "Assets/version.txt;Assets"
```
