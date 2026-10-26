# Mosaic Control Center

Web-based control system for the Nikon Eclipse Ti inverted microscope with CoolSNAP EZ camera, Nikon Intensilight epi-fluorescence illuminator, and Tecan Cavro XCalibur syringe pumps. Built as a modular alternative to Micro-Manager with real-time WebSocket streaming, modular preset/config management, XYZ stage markers, software autofocus, and a visual acquisition timeline.

## Hardware

| Device | Interface | Module |
|--------|-----------|--------|
| Nikon Eclipse Ti | COM (Nikon Ti-E SDK) | `modules/nikon_ti.py` |
| CoolSNAP EZ (Photometrics) | PVCAM SDK via ctypes | `modules/coolsnap.py` + `modules/pvcam_raw.py` |
| Nikon Intensilight C-HGFIE | Serial RS-232 (COM6) | `modules/intensilight.py` |
| Tecan Cavro XCalibur pumps (×4) | RS-485 serial bus | `modules/cavro.py` |
| OBSBOT camera (incubator monitor) | VISCA + RTSP | `modules/obsbot.py` |
| Sutter Lambda SC SmartShutter | Via Ti DiaShutter port | `modules/nikon_ti.py` |

## Quick Start

```bash
cd mos-control
pip install -r requirements.txt
python control_server.py --port 8081
```

Open `http://localhost:8081` in your browser, or use the launcher `.bat` from the project root for auto-launch with Edge/Chrome in app mode.

### Windows app launcher

Double-click `start_mosaic.bat`. It starts `control_server.py` in the
background and opens `http://localhost:8081` in a dedicated Edge or Chrome app
window. Closing that window gracefully shuts down the server and disconnects
the connected hardware, just like stopping the server from the terminal.

For desktop access, right-click `start_mosaic.bat`, choose **Show more
options → Send to → Desktop (create shortcut)**, and rename the shortcut to
`Mosaic Control Center`.

## Project Structure

```
mos-control/
├── control_server.py           # Flask entry point, WebSocket live stream
├── requirements.txt
├── modules/
│   ├── _api.py                 # Flask app, @expose decorator, event push
│   ├── nikon_ti.py             # Nikon Ti: objectives, filters, shutter, lamp, Z/XY abs+rel, PFS
│   ├── coolsnap.py             # CoolSNAP EZ: connect, snap, live, video, stacks, autofocus
│   ├── pvcam_raw.py            # Low-level ctypes wrapper for pvcam64.dll
│   ├── intensilight.py         # Intensilight: epi shutter, ND filter
│   ├── cavro.py                # Tecan Cavro pump API (continuous, coordinated, dispense, withdraw)
│   ├── obsbot.py               # OBSBOT camera (VISCA control + RTSP frame grab + OCR)
│   ├── config.py               # Modular config persistence (modules/ + masters/)
│   └── experiment.py           # Experiment save/load/stop-all
├── syringe_pump/               # Pump hardware drivers
│   ├── tecan_cavro.py          # TecanCavro class (plunger, valve, commands)
│   ├── ftdi_serial.py          # Serial abstraction (FTDI + PySerial)
│   └── motion.py               # Unit conversion (mL↔counts, velocity)
├── web/
│   └── index.html              # Single-page frontend (HTML + CSS + JS)
├── tests/
│   ├── test_coolsnap.py        # Camera unit tests (mocked PVCAM)
│   ├── test_nikon_ti.py        # Microscope unit tests (mocked COM)
│   └── test_camera.py          # Standalone camera diagnostic
├── config/                     # Persistent configs
│   ├── modules/<module>/*.json # Per-module saved configs
│   ├── masters/*.json          # Master configs (reference module configs by name)
│   └── obsbot_waypoints.json   # OBSBOT pan/tilt waypoints
└── captures/                   # Saved .tif + .npy + .meta.json (gitignored)
```

## Architecture

**Backend** — Flask with real OS threads (no gevent). Each hardware module registers API endpoints via the `@expose` decorator, which creates `POST /api/<name>` routes. The Nikon Ti COM interface runs on a dedicated MTA worker thread. Camera live view streams binary JPEG frames over a WebSocket at `/cam/live`.

**Frontend** — Single-page app. Sidebar panels:

- **Config** — Manage master configs and per-module saved configs. Save the currently-active set of module configs as a named master; load any master to restore all modules at once.
- **Cavro** — Connect/configure Tecan Cavro syringe pumps. Continuous, coordinated, dual-push, dispense, and withdraw modes.
- **Setup** — Connect microscope (Nikon Ti + Intensilight). Control objectives, filters, light path, shutter, lamp, Z-drive, XY stage. Save/load *Microscope Presets* (named hardware configurations: DAPI, GFP, etc.).
- **Camera** — Connect camera, set exposure/binning, save directory. Define *Acquisition Presets* (microscope preset + exposure + binning + pseudo-color).
- **Live** — Real-time camera feed via WebSocket plus stage controls: live X/Y/Z display, 8-direction jog pad, Z jog, **autofocus** (Laplacian-variance), goto-marker dropdown, mark-current-position button, snap.
- **Timeline** — Visual timeline with lanes for imaging events (video, stack, timelapse) and pump events (Cavro). Each imaging event references an acquisition preset and (optionally) a *marker* — the stage moves to that XYZ before each capture.
- **Viewer** — Browse and view saved captures with pseudo-color overlays, max/mean projection for stacks.
- **OBSBOT** — Pan/tilt/zoom control, RTSP stream view, named waypoints, region crops, optional OCR for temperature/CO2 from incubator displays.
- **Log** — Activity log.

## Configuration system

Configs are split into **module configs** (one folder each for `camera`, `presets`, `timeline`, `cavro`, `obsbot`) and **master configs** that reference one named module config per module.

```
config/
├── modules/
│   ├── camera/<name>.json      # Exposure, binning, save dir
│   ├── presets/<name>.json     # Microscope presets, acquisition presets, XYZ markers
│   ├── timeline/<name>.json    # Events, lanes, cycle settings
│   ├── cavro/<name>.json       # Bus port, syringe size, valves, addresses
│   └── obsbot/<name>.json      # IP, RTSP URL, waypoints, crops
└── masters/<name>.json         # {"modules": {camera: "default", presets: "DAPI-rig", ...}}
```

Mix and match: load a master to set every module's active config in one click, or swap a single module via the per-panel mini-picker.

## XYZ Markers and Autofocus

**Markers** are named XYZ stage positions, independent of presets. Set them up from the Live View (`+ Mark` saves current XYZ; `Goto` moves to a marker; `List` shows all). They're saved as part of the `presets` module config.

In the **Timeline**, every imaging event has a `Position` dropdown — pick a marker and the stage moves to that XYZ before the capture begins. Leave it at `— current —` to capture wherever the stage already is.

**Autofocus** is image-based: scans Z over a range, scores each frame by Laplacian variance (or numpy gradient fallback), moves to the best Z. Run from the Live View `Auto Focus` button with configurable µm range and step count.

## Dependencies

- Python 3.10+
- Flask, flask-sock
- pyserial (pump communication)
- numpy, Pillow, tifffile, opencv-python (image processing)
- comtypes (Nikon Ti COM interface, Windows only)
- PVCAM SDK (Teledyne/Photometrics, must be installed separately)

## License

Internal use
