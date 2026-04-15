# MOS-11 Control System

Web-based control system for the Nikon Eclipse Ti inverted microscope with CoolSNAP EZ camera, Nikon Intensilight epi-fluorescence illuminator, and Tecan Cavro XCalibur syringe pumps. Built as a modular alternative to Micro-Manager with real-time WebSocket streaming, preset management, and a visual acquisition timeline.

## Hardware

| Device | Interface | Module |
|--------|-----------|--------|
| Nikon Eclipse Ti | COM (Nikon Ti-E SDK) | `modules/nikon_ti.py` |
| CoolSNAP EZ (Photometrics) | PVCAM SDK via ctypes | `modules/coolsnap.py` + `modules/pvcam_raw.py` |
| Nikon Intensilight C-HGFIE | Serial RS-232 (COM6) | `modules/intensilight.py` |
| Tecan Cavro XCalibur pumps (x4) | RS-485 serial bus | `modules/cavro.py` |
| Harvard Apparatus Model 22 pumps (x4) | Serial RS-232 (9600 8N2) | `modules/pumps.py` (legacy) |
| Sutter Lambda SC SmartShutter | Via Ti DiaShutter port | `modules/nikon_ti.py` |

## Quick Start

```bash
cd mos-control
pip install -r requirements.txt
python control_server.py --port 8081
```

Open `http://localhost:8081` in your browser, or use `MOS-11.bat` from the project root for auto-launch with Edge/Chrome in app mode.

## Project Structure

```
mos-control/
├── control_server.py           # Flask entry point, WebSocket live stream
├── requirements.txt
├── modules/
│   ├── _api.py                 # Flask app, @expose decorator, event push
│   ├── nikon_ti.py             # Nikon Ti: objectives, filters, shutter, lamp, Z/XY, PFS
│   ├── coolsnap.py             # CoolSNAP EZ: connect, snap, live view, video, stacks
│   ├── pvcam_raw.py            # Low-level ctypes wrapper for pvcam64.dll
│   ├── intensilight.py         # Intensilight: epi shutter, ND filter
│   ├── cavro.py                # Tecan Cavro pump API (continuous, coordinated modes)
│   ├── pumps.py                # Harvard Apparatus pump API (legacy)
│   └── experiment.py           # Experiment save/load/stop-all
├── syringe_pump/               # Pump hardware drivers
│   ├── tecan_cavro.py          # TecanCavro class (plunger, valve, commands)
│   ├── ftdi_serial.py          # Serial abstraction (FTDI + PySerial)
│   ├── motion.py               # Unit conversion (mL↔counts, velocity)
│   └── syringe_pump_control.py # HarvardPump class
├── web/
│   └── index.html              # Single-page frontend (HTML + CSS + JS)
├── tests/
│   ├── test_coolsnap.py        # Camera unit tests (mocked PVCAM)
│   ├── test_nikon_ti.py        # Microscope unit tests (mocked COM)
│   └── test_camera.py          # Standalone camera diagnostic
└── captures/                   # Saved .npy + .meta.json (gitignored)
```

## Architecture

**Backend** — Flask with real OS threads (no gevent). Each hardware module registers API endpoints via the `@expose` decorator, which creates `POST /api/<name>` routes. The Nikon Ti COM interface runs on a dedicated MTA worker thread. Camera live view streams binary JPEG frames over a WebSocket at `/cam/live`.

**Frontend** — Single-page app with panels for:

- **Cavro Setup** — Connect/configure Tecan Cavro syringe pumps, run continuous and coordinated pumping modes
- **Microscope Setup** — Connect to microscope, control objectives, filters, light path, shutter, lamp, Z-drive, XY stage, Intensilight. Save/load microscope presets.
- **Camera Setup** — Connect camera, set exposure/binning, save directory. Camera presets and combined acquisition presets.
- **Live View** — Real-time camera feed via WebSocket, overlay controls, snap captures.
- **Timeline** — Visual timeline with lanes for imaging events (video, image stack, timelapse) and pump events. Drag to position, resize duration, attach presets.
- **Viewer** — Browse and view saved captures with pseudo-color and projection options.
- **Log** — Activity log.

## Dependencies

- Python 3.10+
- Flask, flask-sock
- pyserial (pump communication)
- numpy, Pillow (image processing)
- comtypes (Nikon Ti COM interface, Windows only)
- PVCAM SDK (Teledyne/Photometrics, must be installed separately)

## License

Internal use
