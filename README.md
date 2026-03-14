# MOS-11 Control System

Web-based control system for the Nikon Eclipse Ti inverted microscope with CoolSNAP EZ camera and Harvard Apparatus syringe pumps. Built as a modular alternative to Micro-Manager with real-time WebSocket streaming, preset management, and a visual acquisition timeline.

## Hardware

| Device | Interface | Module |
|--------|-----------|--------|
| Nikon Eclipse Ti | COM (Nikon Ti-E SDK) | `modules/nikon_ti.py` |
| CoolSNAP EZ (Photometrics) | PVCAM SDK via ctypes | `modules/coolsnap.py` + `modules/pvcam_raw.py` |
| Harvard Apparatus Model 22 pumps (x4) | Serial RS-232 (9600 8N2) | `modules/pumps.py` |
| Sutter Lambda SC SmartShutter | Via Ti DiaShutter port | `modules/nikon_ti.py` |
| Halogen DiaLamp | Via Ti DiaLamp | `modules/nikon_ti.py` |

## Quick Start

```bash
cd mos-control
pip install -r requirements.txt
python control_server.py --port 8081
```

Open `http://localhost:8081` in your browser.

### Standalone Pump Server

To run syringe pumps without microscope/camera dependencies:

```bash
python Syringe_pump/pump_server.py --port 5000
```

## Project Structure

```
mos-control/
├── control_server.py           # Flask entry point, WebSocket live stream
├── requirements.txt
│
├── modules/
│   ├── _api.py                 # Flask app, @expose decorator, event push
│   ├── nikon_ti.py             # Nikon Ti: objectives, filters, shutter, lamp, Z/XY, PFS
│   ├── coolsnap.py             # CoolSNAP EZ: connect, snap, live view, save
│   ├── pvcam_raw.py            # Low-level ctypes wrapper for pvcam64.dll
│   ├── pumps.py                # Syringe pump API (4 pumps, protocols)
│   ├── camera.py               # Micro-Manager camera backend (optional)
│   └── experiment.py           # Experiment save/load
│
├── web/
│   └── index.html              # Single-page frontend
│
└── Syringe_pump/               # Standalone pump server
    ├── pump_server.py
    ├── pump_ui.html
    └── syringe_pump_control.py
```

## Architecture

**Backend** -- Flask with real OS threads (no gevent). Each hardware module registers API endpoints via the `@expose` decorator, which creates `POST /api/<name>` routes. The Nikon Ti COM interface runs on a dedicated MTA worker thread. Camera live view streams binary JPEG frames over a WebSocket at `/cam/live`.

**Frontend** -- Single-page app with five tabs:

- **Setup** -- Connect to microscope, control objectives, filters, light path, shutter, lamp, Z-drive, XY stage. Save/load microscope presets.
- **Camera** -- Connect camera, set exposure/binning, save directory. Camera presets and combined acquisition presets.
- **Live** -- Real-time camera feed via WebSocket, overlay controls, snap captures.
- **Timeline** -- Visual number-line timeline with lanes for imaging events (video, image stack, timelapse) and syringe pumps. Drag to position, resize duration, attach presets. Image stacks can cycle through multiple presets per Z-slice. Pump events control rate, direction, and duration. Playback engine executes events at their timeline positions with a sweeping playhead.
- **Log** -- Activity log.

## Nikon Ti COM Mapping

| SDK Property | Hardware |
|---|---|
| `DiaShutter` | Sutter Lambda SC SmartShutter (transmitted light) |
| `DiaLamp` | Halogen lamp (intensity + on/off) |
| `Nosepiece` | Objective turret (positions 1-6) |
| `FilterBlockCassette1` | Fluorescence filter cassette |
| `LightPathDrive` | Light path selector (eye / left / right / bottom) |
| `ZDrive` | Focus (units: nm) |
| `XDrive` / `YDrive` | Stage position (units: nm) |
| `PFSOffset` / `PFSStatus` | Perfect Focus System |

## Dependencies

- Python 3.10+
- Flask, flask-sock
- pyserial (pump communication)
- numpy, Pillow (image processing)
- comtypes (Nikon Ti COM interface, Windows only)
- PVCAM SDK (Teledyne/Photometrics, must be installed separately)

## License

Internal use -- Bowman Lab.
