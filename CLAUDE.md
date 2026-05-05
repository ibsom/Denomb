# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

**Denombrement** is a Windows desktop application for microbiology laboratory use. It calculates the number of UFC/g (colony-forming units per gram) in a sample, following the solid-medium colony counting method (NF ISO 7218). It is a single-file Python/Tkinter GUI app packaged into an `.exe` via cx_Freeze.

## Running the Application

```bash
python denombrement.py
```

Requires Python 3 with `tkinter` (included in standard Python on Windows).

## Building the Executable (Windows only)

```bash
python setup.py build
```

This uses cx_Freeze to produce a standalone Windows `.exe`. The `setup.py` has a hardcoded `includefiles` path (`E:\Lab\Denomb\icon.ico`) that must be updated to match the local environment before building.

## Architecture

The entire application lives in `denombrement.py`:

- **`box1` (Frame subclass)** — the main UI widget. It manages:
  - A 6-dilution input grid (2 petri dishes per dilution level)
  - Input validation (`validate()`), field reset (`erasefields()`), and calculation trigger (`Calculate()`)
  - The UFC/g formula: `N = ΣC / (d × V × (n1 + 0.1·n2 + 0.01·n3))` where only plates with 30–300 colonies are retained, implemented across `tauxDilution()`, `dictDilRetenues()`, and `resultat()`

- **`Config` class** — reads/writes a JSON config file at `%USERPROFILE%AppData\Local\denombrement\conf.json`. Stores inoculation method preference (`profondeur` → 1 ml, `surface` → 0.1 ml) and a single-instance flag (`instance`).

- **Menu functions** (`save`, `save_as`, `quit`, `about`, `contact`, `mailto`) — standalone functions wired to the Tkinter menu bar. `save`/`save_as` are stubs. `mailto` sends via Gmail SMTP (credentials hardcoded — not functional as-is).

- **Entry point** — `if __name__ == '__main__'`: reads config, enforces single-instance via the `instance` flag, builds the Tk root window, attaches `box1`, builds the menu bar, starts `mainloop()`, then resets the instance flag on exit.

## Key Known Issues

- `Config.__init__` has a broken path: `os.environ.get("USERPROFILE ")` (trailing space) will return `None`, causing a crash on startup. The space must be removed.
- `setup.py` has an absolute Windows path for `includefiles` that must be updated before building.
- SMTP credentials in `mailto()` are hardcoded plaintext and the feature is non-functional.
- The `versions/` folder contains pre-built `.exe` binaries committed to the repo.
