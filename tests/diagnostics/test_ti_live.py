#!/usr/bin/env python3
"""
Interactive live test for the Nikon Ti — uses the real COM interface.
Run with:  python test_ti_live.py
"""

import sys, os
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from modules.nikon_ti import *  # noqa: F403


def header(text):
    print(f"\n{'─' * 50}\n  {text}\n{'─' * 50}")


def safe(fn, *a, **kw):
    try:
        r = fn(*a, **kw)
        print(f"  OK: {r}")
        return r
    except Exception as e:
        print(f"  ERROR: {e}")
        return None


def menu_shutter():
    header("SHUTTER (Sutter Lambda SC — top/transmitted)")
    while True:
        print(f"  State: {safe(shutter_get_state)}  (1=open, 0=closed)")
        c = input("  [o]pen / [c]lose / [b]ack: ").strip().lower()
        if c == "o": safe(shutter_open)
        elif c == "c": safe(shutter_close)
        elif c == "b": break


def menu_lamp():
    header("DIA LAMP (Halogen)")
    while True:
        print(f"  State: {safe(dia_lamp_get_state)}")
        c = input("  [on] / [off] / [i <value>] set intensity / [b]ack: ").strip().lower()
        if c == "on": safe(dia_lamp_on)
        elif c == "off": safe(dia_lamp_off)
        elif c.startswith("i"):
            parts = c.split()
            v = int(parts[1]) if len(parts) > 1 else int(input("  Intensity: "))
            safe(dia_lamp_set_intensity, v)
        elif c == "b": break


def menu_nose():
    header("NOSEPIECE (Objectives)")
    while True:
        print(f"  Position: {safe(nosepiece_get_position)}")
        c = input("  [1-6] / [b]ack: ").strip().lower()
        if c == "b": break
        elif c.isdigit(): safe(nosepiece_set_position, int(c))


def menu_filter():
    header("FILTER BLOCK")
    while True:
        print(f"  Position: {safe(filter_get_position)}")
        c = input("  [1-6] / [b]ack: ").strip().lower()
        if c == "b": break
        elif c.isdigit(): safe(filter_set_position, int(c))


def menu_lp():
    header("LIGHT PATH")
    while True:
        print(f"  Position: {safe(light_path_get_position)}")
        c = input("  [1-4] / [b]ack: ").strip().lower()
        if c == "b": break
        elif c.isdigit(): safe(light_path_set_position, int(c))


def menu_z():
    header("Z DRIVE (Focus, nm)")
    while True:
        print(f"  Position: {safe(z_get_position)} nm")
        c = input("  [+N/-N] relative nm / [a <nm>] absolute / [b]ack: ").strip().lower()
        if c == "b": break
        elif c.startswith("a"):
            parts = c.split()
            v = int(parts[1]) if len(parts) > 1 else int(input("  Absolute nm: "))
            safe(z_move_absolute, v)
        else:
            try: safe(z_move_relative, int(c))
            except ValueError: print("  Enter a number like +1000 or -500")


def menu_xy():
    header("XY STAGE (nm)")
    while True:
        print(f"  Position: {safe(xy_get_position)}")
        c = input("  [xr <nm>] / [yr <nm>] relative / [b]ack: ").strip().lower()
        if c == "b": break
        elif c.startswith("xr"):
            v = int(c.split()[1]) if len(c.split()) > 1 else int(input("  X relative nm: "))
            safe(x_move_relative, v)
        elif c.startswith("yr"):
            v = int(c.split()[1]) if len(c.split()) > 1 else int(input("  Y relative nm: "))
            safe(y_move_relative, v)


def menu_pfs():
    header("PFS (Perfect Focus)")
    while True:
        print(f"  Status: {safe(pfs_get_status)}")
        c = input("  [on] / [off] / [s]earch / [b]ack: ").strip().lower()
        if c == "on": safe(pfs_enable)
        elif c == "off": safe(pfs_disable)
        elif c == "s": safe(pfs_search)
        elif c == "b": break


def menu_status():
    header("FULL STATUS")
    s = get_full_status()
    for k, v in s.items():
        print(f"  {k:25s}: {v}")
    input("\n  Press Enter...")


def menu_probe():
    header("PROBE DEVICE")
    devices = [
        "EpiShutter", "DiaShutter", "DiaLamp", "Nosepiece",
        "FilterBlockCassette1", "LightPathDrive", "ZDrive",
        "XDrive", "YDrive", "PFS", "Analyzer", "AuxShutter",
        "MainController", "RemoteController", "TIRF",
    ]
    print("  Available:")
    for d in devices:
        print(f"    {d}")
    name = input("\n  Device name (or 'all'): ").strip()
    targets = devices if name == "all" else [name]
    for t in targets:
        r = probe_device(t)
        print(f"\n  --- {t} ---")
        if "error" in r:
            print(f"    Error: {r['error']}")
        else:
            for a in r.get("attributes", []):
                if "error" in a:
                    print(f"    {a['name']:35s}  ERROR: {a['error'][:60]}")
                else:
                    print(f"    {a['name']:35s} = {a['value'][:80]}  ({a['type']})")
    input("\n  Press Enter...")


def main():
    print("=" * 50)
    print("  NIKON Ti LIVE TEST")
    print("=" * 50)
    print("\nConnecting to microscope...")

    try:
        connect()
        print(f"Connected!  System: {get_system_type()}")
    except TiError as e:
        print(f"FAILED: {e}")
        print("\n  1. Is the microscope powered on?")
        print("  2. Is the USB cable connected?")
        print("  3. Does 'Nikon USB microscope' appear in Device Manager?")
        sys.exit(1)

    try:
        while True:
            header("MAIN MENU")
            print("  [1] Shutter        (Lambda SC)")
            print("  [2] Halogen lamp   (intensity)")
            print("  [3] Nosepiece      (objectives)")
            print("  [4] Filter block")
            print("  [5] Light path")
            print("  [6] Z drive        (focus)")
            print("  [7] XY stage")
            print("  [8] PFS")
            print("  [s] Full status")
            print("  [p] Probe device")
            print("  [q] Quit")
            c = input("\n  Choice: ").strip().lower()
            menus = {"1": menu_shutter, "2": menu_lamp,
                     "3": menu_nose, "4": menu_filter, "5": menu_lp,
                     "6": menu_z, "7": menu_xy, "8": menu_pfs,
                     "s": menu_status, "p": menu_probe}
            if c == "q":
                break
            elif c in menus:
                menus[c]()
    except KeyboardInterrupt:
        print("\nInterrupted.")
    finally:
        print("\nDisconnecting...")
        disconnect()
        print("Done.")


if __name__ == "__main__":
    main()
