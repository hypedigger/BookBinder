# 🌸 BookBinder

**Merge two scanned comic pages into a single double-page spread — with one keyboard shortcut, straight from Windows Explorer.**

Scanned comics, manga and graphic novels often split a double-page spread into two separate image files. BookBinder stitches them back together: select the two pages in Explorer, hit a hotkey, done.

## Features

- **Global hotkeys** — select two images in Windows Explorer, press `Ctrl+Shift+D` to merge (alphabetical order = left page first) or `Ctrl+Shift+F` for reverse order. No need to open the app window.
- **Smart border trim** — automatically crops the uniform border around each page before joining, so the spread lines up cleanly.
- **Height matching** — pages are scaled to the same height before merging.
- **Output formats** — JPEG, JPEG 2000, PNG or WebP, with per-format quality settings.
- **Safe cleanup** — optionally sends the source files to the Recycle Bin after a successful merge (never hard-deletes).
- **Configurable** — output format, quality, filename suffix and both hotkeys can be changed in the Options dialog.

## Installation

Requires **Windows** and **Python 3.10+**.

```bash
git clone https://github.com/hypedigger/BookBinder.git
cd BookBinder
python -m venv .venv
.venv\Scripts\pip install -r requirements.txt
```

Run the app:

```bash
.venv\Scripts\pythonw fusion_images.py
```

Optional — add a Start Menu shortcut (also generates the app icon):

```bash
.venv\Scripts\python install_shortcut.py
```

## Usage

1. Launch BookBinder (it sits in a small window; hotkeys are global).
2. In Windows Explorer, select **exactly two images** (the two halves of a spread).
3. Press **`Ctrl+Shift+D`** — the merged image is saved next to the sources with the `_join` suffix.
4. Wrong page order? Use **`Ctrl+Shift+F`** instead.

Settings are stored in a local `settings.json` (created on first run).

## Notes

- The global hotkey hook uses the [`keyboard`](https://github.com/boppreh/keyboard) library; some antivirus software may flag keyboard hooks — it is only used to detect the two shortcuts above.
- UI is currently in French; contributions welcome.

## Support

If this tool saved you some tedious image editing, you can support development:

[![ko-fi](https://ko-fi.com/img/githubbutton_sm.svg)](https://ko-fi.com/hypedigger)

## License

[MIT](LICENSE)
