# Bid Document Manager build guide (macOS & Windows)

This project ships a PyQt5-based desktop tool in ``tools/bid_manager.py``. You
can create native executables for both macOS and Windows with ``PyInstaller``.

## Prerequisites

* Python 3.10+ with ``pip`` available.
* Install the toolchain with the curated requirements file:
  ```bash
  pip install -r tools/requirements_bid_manager.txt
  ```

## macOS build steps

```bash
python tools/build_bid_manager.py
```

* The helper script normalizes ``--add-data`` separators for you and includes
  the repository ``data`` directory by default if it exists.
* Use ``--onefile`` to produce a single-file bundle (slower startup) or
  ``--icon`` to apply a custom ``.icns`` icon.
* The final ``BidManager.app`` lives in ``dist``.

## Windows build steps

```powershell
python tools\build_bid_manager.py
```

Notes:

* The script auto-selects ``;`` as the data separator for Windows.
* The executable is created under ``dist\BidManager\BidManager.exe``.

## Common tips

* To reduce file size, consider ``--onefile`` if startup time is acceptable.
* If you need custom icons, pass ``--icon path/to/icon.ico`` (use ``.icns`` on
  macOS).
* The application stores user settings via ``QSettings`` under the
  organization name ``BidManager`` and application name ``CompanyManager``. No
  extra configuration is required for packaging.
