# Bid Document Manager build guide (macOS & Windows)

This project ships a PyQt5-based desktop tool in ``tools/bid_manager.py``. You
can create native executables for both macOS and Windows with ``PyInstaller``.

## Prerequisites

* Python 3.10+ with ``pip`` available.
* ``PyQt5`` and optional text-extraction dependencies:
  ```bash
  pip install pyqt5 pymupdf python-docx
  ```
* ``PyInstaller`` installed globally or in a virtual environment:
  ```bash
  pip install pyinstaller
  ```

## macOS build steps

```bash
cd tools
pyinstaller --noconfirm --windowed --name "BidManager" \
  --add-data "../data:./data" \
  bid_manager.py
```

* The ``--windowed`` flag hides the terminal.
* ``--add-data`` can be extended if you keep company templates or icons in the
  repo; remove it if unused.
* The final ``BidManager.app`` lives in ``dist``.

## Windows build steps

```powershell
cd tools
pyinstaller --noconfirm --windowed --name "BidManager" `
  --add-data "..\\data;data" `
  bid_manager.py
```

Notes:

* Replace the ``--add-data`` path separator with ``;`` on Windows.
* The executable is created under ``dist\BidManager\BidManager.exe``.

## Common tips

* To reduce file size, consider ``--onefile`` if startup time is acceptable.
* If you need custom icons, pass ``--icon path/to/icon.ico`` (use ``.icns`` on
  macOS).
* The application stores user settings via ``QSettings`` under the
  organization name ``BidManager`` and application name ``CompanyManager``. No
  extra configuration is required for packaging.
